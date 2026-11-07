#include "stdafx.h"

#include "DocReader.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <objbase.h>
#include <ole2.h>

#include <algorithm>
#include <array>
#include <fstream>
#include <limits>
#include <memory>
#include <sstream>

namespace docreader
{
namespace
{
    constexpr std::uint16_t kWordIdent = 0xA5EC;
    constexpr std::size_t kFibBaseSize = 32;
    constexpr std::size_t kFibRgFcLcbClxIndex = 33;

    struct ComPtr
    {
        IUnknown* p = nullptr;
        ~ComPtr() { if (p) p->Release(); }
        template<class T> T** put() { return reinterpret_cast<T**>(&p); }
    };

    struct Storage
    {
        IStorage* p = nullptr;
        ~Storage() { if (p) p->Release(); }
    };

    struct Stream
    {
        IStream* p = nullptr;
        ~Stream() { if (p) p->Release(); }
    };

    bool ReadAt(IStream* stream, std::uint64_t offset, void* dst, std::size_t bytes)
    {
        if (!stream || (!dst && bytes))
            return false;

        LARGE_INTEGER pos{};
        pos.QuadPart = static_cast<LONGLONG>(offset);
        if (FAILED(stream->Seek(pos, STREAM_SEEK_SET, nullptr)))
            return false;

        auto* out = static_cast<std::uint8_t*>(dst);
        std::size_t done = 0;
        while (done < bytes)
        {
            ULONG got = 0;
            const ULONG request = static_cast<ULONG>(
                std::min<std::size_t>(bytes - done, (std::numeric_limits<ULONG>::max)()));
            if (FAILED(stream->Read(out + done, request, &got)) || got == 0)
                return false;
            done += got;
        }
        return true;
    }

    bool GetStreamSize(IStream* stream, std::uint64_t& size)
    {
        STATSTG stat{};
        if (FAILED(stream->Stat(&stat, STATFLAG_NONAME)))
            return false;
        size = stat.cbSize.QuadPart;
        return true;
    }

    std::uint16_t U16(const std::vector<std::uint8_t>& b, std::size_t o)
    {
        return static_cast<std::uint16_t>(b[o] | (std::uint16_t(b[o + 1]) << 8));
    }

    std::uint32_t U32(const std::vector<std::uint8_t>& b, std::size_t o)
    {
        return std::uint32_t(b[o]) |
            (std::uint32_t(b[o + 1]) << 8) |
            (std::uint32_t(b[o + 2]) << 16) |
            (std::uint32_t(b[o + 3]) << 24);
    }

    bool ReadVector(IStream* s, std::uint64_t off, std::uint64_t len,
                    std::uint64_t limit, std::vector<std::uint8_t>& out)
    {
        if (len > limit || len > (std::numeric_limits<std::size_t>::max)())
            return false;
        out.resize(static_cast<std::size_t>(len));
        return len == 0 || ReadAt(s, off, out.data(), out.size());
    }

    struct FibInfo
    {
        bool table1 = false;
        bool encrypted = false;
        std::uint32_t fcMin = 0;
        std::uint32_t fcMac = 0;
        std::uint32_t fcClx = 0;
        std::uint32_t lcbClx = 0;
    };

    bool ParseFib(const std::vector<std::uint8_t>& word, FibInfo& fib, std::wstring& error)
    {
        if (word.size() < kFibBaseSize)
        {
            error = L"WordDocument stream is too small to contain a FIB.";
            return false;
        }

        if (U16(word, 0) != kWordIdent)
        {
            error = L"Invalid WordDocument FIB signature.";
            return false;
        }

        const std::uint16_t nFib = U16(word, 2);
        if (nFib < 0x00C1)
        {
            error = L"This importer currently supports Word 97-2003 documents (nFib >= 0x00C1).";
            return false;
        }

        const std::uint16_t flags = U16(word, 10);
        fib.encrypted = (flags & 0x0100) != 0 || (flags & 0x8000) != 0;
        fib.table1 = (flags & 0x0200) != 0;
        fib.fcMin = U32(word, 24);
        fib.fcMac = U32(word, 28);

        // FibRgW: csw followed by csw 16-bit words.
        std::size_t pos = 32;
        if (pos + 2 > word.size())
        {
            error = L"Truncated FIB (csw).";
            return false;
        }
        const std::uint16_t csw = U16(word, pos);
        pos += 2;
        if (csw > 0x100 || pos + std::size_t(csw) * 2 + 2 > word.size())
        {
            error = L"Invalid FIB word-count.";
            return false;
        }
        pos += std::size_t(csw) * 2;

        // FibRgLw: cslw followed by cslw 32-bit words.
        const std::uint32_t cslw = U16(word, pos);
        pos += 2;
        if (cslw > 0x100 || pos + std::size_t(cslw) * 4 + 2 > word.size())
        {
            error = L"Invalid FIB long-word count.";
            return false;
        }
        pos += std::size_t(cslw) * 4;

        const std::uint16_t cbRgFcLcb = U16(word, pos);
        pos += 2;
        if (cbRgFcLcb <= kFibRgFcLcbClxIndex ||
            pos + std::size_t(cbRgFcLcb) * 8 > word.size())
        {
            error = L"FIB does not contain the expected Word 97 CLX location entry.";
            return false;
        }

        const std::size_t clxEntry = pos + kFibRgFcLcbClxIndex * 8;
        fib.fcClx = U32(word, clxEntry);
        fib.lcbClx = U32(word, clxEntry + 4);

        if (fib.encrypted)
        {
            error = L"Encrypted or obfuscated Word documents are not supported.";
            return false;
        }
        if (fib.lcbClx == 0)
        {
            error = L"The document has no CLX piece table. This importer requires a Word 97-2003 piece table.";
            return false;
        }
        return true;
    }

    bool DecodeCompressed(const std::uint8_t* src, std::size_t count,
                          UINT codePage, std::wstring& out)
    {
        if (count == 0)
            return true;
        if (count > static_cast<std::size_t>((std::numeric_limits<int>::max)()))
            return false;

        const int needed = MultiByteToWideChar(codePage, 0,
            reinterpret_cast<const char*>(src), static_cast<int>(count), nullptr, 0);
        if (needed <= 0)
            return false;

        const std::size_t old = out.size();
        out.resize(old + static_cast<std::size_t>(needed));
        return MultiByteToWideChar(codePage, 0,
            reinterpret_cast<const char*>(src), static_cast<int>(count),
            out.data() + old, needed) == needed;
    }

    bool DecodeUtf16(const std::uint8_t* src, std::size_t count, std::wstring& out)
    {
        if ((count & 1) != 0)
            return false;
        const std::size_t chars = count / 2;
        const std::size_t old = out.size();
        out.resize(old + chars);
        for (std::size_t i = 0; i < chars; ++i)
            out[old + i] = static_cast<wchar_t>(
                std::uint16_t(src[i * 2]) | (std::uint16_t(src[i * 2 + 1]) << 8));
        return true;
    }

    struct Piece
    {
        std::uint32_t cpStart = 0;
        std::uint32_t cpEnd = 0;
        std::uint32_t fc = 0;
        bool compressed = false;
    };

    bool ParsePieceTable(const std::vector<std::uint8_t>& table,
                         std::uint32_t fcClx, std::uint32_t lcbClx,
                         std::vector<Piece>& pieces, std::wstring& error)
    {
        if (std::uint64_t(fcClx) + lcbClx > table.size())
        {
            error = L"CLX range lies outside the table stream.";
            return false;
        }

        std::size_t p = fcClx;
        const std::size_t end = std::size_t(fcClx) + lcbClx;

        // CLX may begin with zero or more grpprl records (0x01).
        while (p < end && table[p] == 0x01)
        {
            if (p + 3 > end)
            {
                error = L"Truncated CLX grpprl.";
                return false;
            }
            const std::uint16_t cb = std::uint16_t(table[p + 1]) |
                (std::uint16_t(table[p + 2]) << 8);
            p += 3;
            if (p + cb > end)
            {
                error = L"CLX grpprl exceeds its declared range.";
                return false;
            }
            p += cb;
        }

        if (p >= end || table[p] != 0x02 || p + 5 > end)
        {
            error = L"CLX does not contain a valid Pcdt record.";
            return false;
        }

        const std::uint32_t lcb = U32(table, p + 1);
        p += 5;
        if (lcb < 4 || p + lcb > end)
        {
            error = L"Invalid Pcdt length.";
            return false;
        }

        // PlcPcd = (n+1) CP values followed by n 8-byte PCD records.
        // Therefore lcb = 4*(n+1) + 8*n = 12*n + 4.
        if ((lcb - 4) % 12 != 0)
        {
            error = L"Invalid PlcPcd size.";
            return false;
        }
        const std::uint32_t n = (lcb - 4) / 12;
        if (n == 0 || n > 1000000)
        {
            error = L"Unreasonable piece count.";
            return false;
        }

        const std::size_t cpBase = p;
        const std::size_t pcdBase = p + std::size_t(n + 1) * 4;
        pieces.reserve(n);

        std::uint32_t previous = 0;
        for (std::uint32_t i = 0; i < n; ++i)
        {
            const std::uint32_t cp0 = U32(table, cpBase + std::size_t(i) * 4);
            const std::uint32_t cp1 = U32(table, cpBase + std::size_t(i + 1) * 4);
            if (cp1 < cp0 || (i && cp0 != previous))
            {
                error = L"Piece table CP positions are not monotonic.";
                return false;
            }
            previous = cp1;

            const std::size_t pcd = pcdBase + std::size_t(i) * 8;
            const std::uint32_t rawFc = U32(table, pcd + 2);
            const bool compressed = (rawFc & 0x40000000u) != 0;
            const std::uint32_t storedFc = rawFc & 0x3FFFFFFFu;

            Piece piece;
            piece.cpStart = cp0;
            piece.cpEnd = cp1;
            piece.compressed = compressed;
            // Compressed pieces store the byte offset multiplied by two.
            piece.fc = compressed ? (storedFc / 2) : storedFc;
            pieces.push_back(piece);
        }

        const std::uint32_t finalCp = U32(table, cpBase + std::size_t(n) * 4);
        if (finalCp != previous)
        {
            error = L"Invalid final CP in piece table.";
            return false;
        }
        return true;
    }

    void AppendRtfUnicode(std::wstring& out, wchar_t ch)
    {
        const unsigned int c = static_cast<unsigned int>(ch);
        if (c == L'\\' || c == L'{' || c == L'}')
        {
            out.push_back(L'\\');
            out.push_back(static_cast<wchar_t>(c));
        }
        else if (c == L'\t')
        {
            out += L"\\tab ";
        }
        else if (c >= 0x20 && c <= 0x7E)
        {
            out.push_back(static_cast<wchar_t>(c));
        }
        else
        {
            // RTF \uN uses a signed 16-bit value; UTF-16 surrogate pairs are
            // emitted as two consecutive \u controls.
            const std::int32_t signedValue = (c <= 0x7FFF)
                ? static_cast<std::int32_t>(c)
                : static_cast<std::int32_t>(c) - 0x10000;
            out += L"\\u";
            out += std::to_wstring(signedValue);
            out += L"?";
        }
    }

    std::wstring MakeRtf(const std::wstring& text, bool placeholders)
    {
        std::wstring rtf = L"{\\rtf1\\ansi\\deff0\\uc1\\viewkind4";
        rtf += L"{\\fonttbl{\\f0\\fnil Segoe UI;}}";
        rtf += L"\\f0\\fs20 ";

        for (wchar_t ch : text)
        {
            switch (ch)
            {
            case L'\r': rtf += L"\\par\n"; break;
            case L'\n': break; // Word paragraph terminators are CR (0x0D).
            case 0x0007: rtf += L"\\cell "; break;
            case 0x000B: rtf += L"\\line "; break;
            case 0x000C: rtf += L"\\page "; break;
            case 0x0001:
            case 0x0008:
                if (placeholders)
                    rtf += L"[object] ";
                break;
            case 0x0013: // field begin
            case 0x0014: // field separator
            case 0x0015: // field end
                // Field instruction text is retained as ordinary text.
                break;
            default:
                AppendRtfUnicode(rtf, ch);
                break;
            }
        }

        rtf += L"}";
        return rtf;
    }

    bool WideToUtf8Bytes(const std::wstring& text, std::vector<std::uint8_t>& bytes)
    {
        if (text.empty())
        {
            bytes.clear();
            return true;
        }
        if (text.size() > static_cast<std::size_t>((std::numeric_limits<int>::max)()))
            return false;

        const int needed = WideCharToMultiByte(CP_UTF8, 0, text.data(),
            static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
        if (needed <= 0)
            return false;
        bytes.resize(static_cast<std::size_t>(needed));
        return WideCharToMultiByte(CP_UTF8, 0, text.data(),
            static_cast<int>(text.size()), reinterpret_cast<char*>(bytes.data()),
            needed, nullptr, nullptr) == needed;
    }
}

Result ConvertToRtf(const std::wstring& fileName,
                    std::vector<std::uint8_t>& rtf,
                    const Options& options)
{
    Result result;
    rtf.clear();

    const HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool uninitialize = SUCCEEDED(initHr);
    if (FAILED(initHr) && initHr != RPC_E_CHANGED_MODE)
    {
        result.error = L"CoInitializeEx failed.";
        return result;
    }

    Storage storage;
    HRESULT hr = StgOpenStorage(fileName.c_str(), nullptr,
        STGM_READ | STGM_SHARE_DENY_WRITE, nullptr, 0, &storage.p);
    if (FAILED(hr))
    {
        if (uninitialize) CoUninitialize();
        result.error = L"Cannot open the DOC compound file (StgOpenStorage failed).";
        return result;
    }

    auto openStream = [&](const wchar_t* name, IStream** out) -> HRESULT
    {
        *out = nullptr;
        return storage.p->OpenStream(name, nullptr,
            STGM_READ | STGM_SHARE_EXCLUSIVE, 0, out);
    };

    Stream wordStream, tableStream;
    if (FAILED(openStream(L"WordDocument", &wordStream.p)))
    {
        if (uninitialize) CoUninitialize();
        result.error = L"The OLE file does not contain a WordDocument stream.";
        return result;
    }

    std::uint64_t wordSize = 0;
    if (!GetStreamSize(wordStream.p, wordSize) || wordSize > options.maxStreamBytes)
    {
        if (uninitialize) CoUninitialize();
        result.error = L"WordDocument stream is invalid or exceeds the configured size limit.";
        return result;
    }

    std::vector<std::uint8_t> word;
    if (!ReadVector(wordStream.p, 0, wordSize, options.maxStreamBytes, word))
    {
        if (uninitialize) CoUninitialize();
        result.error = L"Failed to read WordDocument stream.";
        return result;
    }

    FibInfo fib;
    if (!ParseFib(word, fib, result.error))
    {
        if (uninitialize) CoUninitialize();
        return result;
    }

    const wchar_t* tableName = fib.table1 ? L"1Table" : L"0Table";
    if (FAILED(openStream(tableName, &tableStream.p)))
    {
        if (uninitialize) CoUninitialize();
        result.error = L"Cannot open the table stream selected by the FIB.";
        return result;
    }

    std::uint64_t tableSize = 0;
    if (!GetStreamSize(tableStream.p, tableSize) || tableSize > options.maxStreamBytes)
    {
        if (uninitialize) CoUninitialize();
        result.error = L"Table stream is invalid or exceeds the configured size limit.";
        return result;
    }

    std::vector<std::uint8_t> table;
    if (!ReadVector(tableStream.p, 0, tableSize, options.maxStreamBytes, table))
    {
        if (uninitialize) CoUninitialize();
        result.error = L"Failed to read table stream.";
        return result;
    }

    std::vector<Piece> pieces;
    if (!ParsePieceTable(table, fib.fcClx, fib.lcbClx, pieces, result.error))
    {
        if (uninitialize) CoUninitialize();
        return result;
    }

    const UINT codePage = options.codePage ? options.codePage : CP_ACP;
    std::wstring text;
    for (const Piece& piece : pieces)
    {
        const std::uint64_t cpLength = std::uint64_t(piece.cpEnd) - piece.cpStart;
        const std::uint64_t byteLength = cpLength * (piece.compressed ? 1ull : 2ull);
        if (std::uint64_t(piece.fc) + byteLength > word.size())
        {
            result.error = L"A text piece points outside the WordDocument stream.";
            if (uninitialize) CoUninitialize();
            return result;
        }

        const auto* src = word.data() + piece.fc;
        const bool ok = piece.compressed
            ? DecodeCompressed(src, static_cast<std::size_t>(byteLength), codePage, text)
            : DecodeUtf16(src, static_cast<std::size_t>(byteLength), text);

        if (!ok)
        {
            result.error = L"Failed to decode a text piece.";
            if (uninitialize) CoUninitialize();
            return result;
        }
    }

    // Normalize Word's common paragraph/line separators. Do not remove the
    // final paragraph mark: Rich Edit handles a trailing \par correctly.
    std::wstring rtfText = MakeRtf(text, options.emitObjectPlaceholders);
    if (!WideToUtf8Bytes(rtfText, rtf))
    {
        result.error = L"Failed to encode generated RTF.";
        if (uninitialize) CoUninitialize();
        return result;
    }

    result.success = true;
    result.charactersRead = text.size();
    result.pieceCount = static_cast<std::uint32_t>(pieces.size());
    result.warnings.push_back(
        L"This first version extracts document text and paragraph structure. "
        L"Character/paragraph formatting and real embedded-picture extraction are not implemented yet.");

    if (uninitialize)
        CoUninitialize();
    return result;
}

Result ConvertToRtfFile(const std::wstring& inputFile,
                        const std::wstring& outputRtfFile,
                        const Options& options)
{
    std::vector<std::uint8_t> data;
    Result result = ConvertToRtf(inputFile, data, options);
    if (!result.success)
        return result;

    std::ofstream file(outputRtfFile, std::ios::binary | std::ios::trunc);
    if (!file)
    {
        result.success = false;
        result.error = L"Cannot create output RTF file.";
        return result;
    }
    file.write(reinterpret_cast<const char*>(data.data()),
               static_cast<std::streamsize>(data.size()));
    if (!file)
    {
        result.success = false;
        result.error = L"Failed while writing output RTF file.";
    }
    return result;
}
}
