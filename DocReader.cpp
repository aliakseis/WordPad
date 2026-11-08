#include "stdafx.h"

#include "DocReader.h"
#include "ComUsageScope.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <objbase.h>
#include <ole2.h>

#include <algorithm>
#include <cstring>
#include <array>
#include <fstream>
#include <limits>
#include <memory>
#include <sstream>
#include <unordered_map>

namespace docreader
{
namespace
{
    constexpr std::uint16_t kWordIdent = 0xA5EC;
    constexpr std::size_t kFibBaseSize = 32;
    constexpr std::size_t kFibRgFcLcbClxIndex = 33;
    constexpr std::size_t kFibRgFcLcbBteChpxIndex = 12;

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
                std::min<std::size_t>(bytes - done, std::numeric_limits<ULONG>::max()));
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
        if (len > limit || len > std::numeric_limits<std::size_t>::max())
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
        std::uint32_t fcPlcfBteChpx = 0;
        std::uint32_t lcbPlcfBteChpx = 0;
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
        const std::size_t chpxEntry = pos + kFibRgFcLcbBteChpxIndex * 8;
        fib.fcPlcfBteChpx = U32(word, chpxEntry);
        fib.lcbPlcfBteChpx = U32(word, chpxEntry + 4);

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
        if (count > static_cast<std::size_t>(std::numeric_limits<int>::max()))
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

    struct ChpxRun
    {
        std::uint32_t fcFirst = 0;
        std::uint32_t fcLim = 0;
        std::uint32_t picLocation = 0;
        bool hasPicture = false;
    };

    std::size_t SprmOperandLength(std::uint16_t sprm, const std::uint8_t* p,
                                  std::size_t remaining)
    {
        const unsigned spra = (sprm >> 13) & 7;
        switch (spra)
        {
        case 0: case 1: return 1;
        case 2: return 2;
        case 3: return 4;
        case 4: return 2;
        case 5: return 2;
        case 6:
            if (remaining < 1) return remaining + 1;
            if (p[0] != 0xFF) return std::size_t(p[0]) + 1;
            if (remaining < 3) return remaining + 1;
            return std::size_t(p[1] | (std::uint16_t(p[2]) << 8)) + 3;
        case 7: return 3;
        }
        return remaining + 1;
    }

    bool ParseChpxRuns(const std::vector<std::uint8_t>& word,
                       const std::vector<std::uint8_t>& table,
                       const FibInfo& fib, std::vector<ChpxRun>& runs)
    {
        const std::uint64_t off = fib.fcPlcfBteChpx;
        const std::uint64_t len = fib.lcbPlcfBteChpx;
        if (!len || len < 12 || off + len > table.size() || (len - 4) % 8 != 0)
            return false;

        const std::uint32_t count = static_cast<std::uint32_t>((len - 4) / 8);
        if (!count || count > 100000)
            return false;
        const std::size_t fcBase = static_cast<std::size_t>(off);
        const std::size_t pnBase = fcBase + std::size_t(count + 1) * 4;
        for (std::uint32_t i = 0; i < count; ++i)
        {
            const std::uint32_t pn = U32(table, pnBase + std::size_t(i) * 4);
            const std::uint64_t pageOff64 = std::uint64_t(pn) * 512;
            if (pageOff64 + 512 > word.size())
                continue;
            const std::size_t page = static_cast<std::size_t>(pageOff64);
            const std::uint8_t crun = word[page + 511];
            if (!crun || std::size_t(crun + 1) * 4 + crun > 511)
                continue;
            const std::size_t bxBase = page + std::size_t(crun + 1) * 4;
            for (unsigned j = 0; j < crun; ++j)
            {
                ChpxRun run;
                run.fcFirst = U32(word, page + std::size_t(j) * 4);
                run.fcLim = U32(word, page + std::size_t(j + 1) * 4);
                if (run.fcLim <= run.fcFirst)
                    continue;
                const std::uint8_t chpxOffset = word[bxBase + j];
                if (!chpxOffset)
                    continue;
                const std::size_t chpx = page + std::size_t(chpxOffset) * 2;
                if (chpx >= page + 511)
                    continue;
                const std::uint8_t cb = word[chpx];
                if (std::size_t(chpx) + 1 + cb > page + 511)
                    continue;
                std::size_t pos = chpx + 1;
                const std::size_t end = pos + cb;
                while (pos + 2 <= end)
                {
                    const std::uint16_t sprm = std::uint16_t(word[pos]) |
                        (std::uint16_t(word[pos + 1]) << 8);
                    pos += 2;
                    const std::size_t operand = SprmOperandLength(sprm,
                        word.data() + pos, end - pos);
                    if (operand > end - pos)
                        break;
                    if (sprm == 0x6A03 && operand == 4)
                    {
                        run.picLocation = U32(word, pos);
                        run.hasPicture = true;
                    }
                    pos += operand;
                }
                if (run.hasPicture)
                    runs.push_back(run);
            }
        }
        return !runs.empty();
    }

    struct Picture
    {
        std::string control;
        std::vector<std::uint8_t> bytes;
        unsigned width = 0;
        unsigned height = 0;
    };

    bool FindPictureBytes(const std::vector<std::uint8_t>& data, std::uint32_t offset,
                          std::uint64_t maxBytes, Picture& picture)
    {
        if (offset > data.size() || data.size() - offset < 68)
            return false;
        const std::uint32_t lcb = U32(data, offset);
        if (lcb < 68 || lcb > maxBytes || std::uint64_t(offset) + lcb > data.size())
            return false;
        const std::size_t begin = std::size_t(offset) + 68;
        const std::size_t end = std::size_t(offset) + lcb;

        // OfficeArt blips contain their native file signature after a UID.
        // Search only inside this PICF's declared extent and require a complete
        // PNG/JPEG image terminator before accepting a candidate.
        static const std::uint8_t pngSig[8] = { 137,80,78,71,13,10,26,10 };
        for (std::size_t p = begin; p + 8 <= end; ++p)
        {
            if (std::memcmp(data.data() + p, pngSig, 8) == 0)
            {
                std::size_t q = p + 8;
                bool foundEnd = false;
                while (q + 12 <= end)
                {
                    const std::uint32_t chunk = std::uint32_t(data[q]) << 24 |
                        std::uint32_t(data[q + 1]) << 16 |
                        std::uint32_t(data[q + 2]) << 8 | data[q + 3];
                    if (chunk > end - q - 12) break;
                    const bool iend = data[q + 4] == 'I' && data[q + 5] == 'E' &&
                        data[q + 6] == 'N' && data[q + 7] == 'D';
                    q += std::size_t(chunk) + 12;
                    if (iend) { foundEnd = true; break; }
                }
                if (foundEnd)
                {
                    picture.control = "pngblip";
                    picture.bytes.assign(data.begin() + p, data.begin() + q);
                    if (p + 24 <= end && data[p + 12] == 'I' && data[p + 13] == 'H' &&
                        data[p + 14] == 'D' && data[p + 15] == 'R')
                    {
                        picture.width = (unsigned(data[p + 16]) << 24) |
                            (unsigned(data[p + 17]) << 16) | (unsigned(data[p + 18]) << 8) | data[p + 19];
                        picture.height = (unsigned(data[p + 20]) << 24) |
                            (unsigned(data[p + 21]) << 16) | (unsigned(data[p + 22]) << 8) | data[p + 23];
                    }
                    return true;
                }
            }
            if (data[p] == 0xFF && data[p + 1] == 0xD8 && data[p + 2] == 0xFF)
            {
                for (std::size_t q = p + 3; q + 1 < end; ++q)
                {
                    if (data[q] == 0xFF && data[q + 1] == 0xD9)
                    {
                        picture.control = "jpegblip";
                        picture.bytes.assign(data.begin() + p, data.begin() + q + 2);
                        // Read dimensions from a Start Of Frame marker.
                        for (std::size_t k = p + 2; k + 8 < q; )
                        {
                            if (data[k] != 0xFF) { ++k; continue; }
                            const std::uint8_t marker = data[k + 1];
                            if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2 || marker == 0xC3)
                            {
                                picture.height = (unsigned(data[k + 5]) << 8) | data[k + 6];
                                picture.width = (unsigned(data[k + 7]) << 8) | data[k + 8];
                                break;
                            }
                            if (k + 4 > q) break;
                            const unsigned seg = (unsigned(data[k + 2]) << 8) | data[k + 3];
                            if (seg < 2) break;
                            k += 2 + seg;
                        }
                        return true;
                    }
                }
            }
        }
        return false;
    }

    void AppendPictureRtf(std::wstring& rtf, const Picture& picture)
    {
        rtf += L"{\\pict\\";
        for (char c : picture.control) rtf.push_back(static_cast<wchar_t>(c));
        if (picture.width) rtf += L"\\picw" + std::to_wstring(picture.width);
        if (picture.height) rtf += L"\\pich" + std::to_wstring(picture.height);
        if (picture.width) rtf += L"\\picwgoal" + std::to_wstring(picture.width * 15u);
        if (picture.height) rtf += L"\\pichgoal" + std::to_wstring(picture.height * 15u);
        rtf += L" ";
        static const wchar_t hex[] = L"0123456789abcdef";
        unsigned column = 0;
        for (std::uint8_t b : picture.bytes)
        {
            rtf.push_back(hex[b >> 4]);
            rtf.push_back(hex[b & 15]);
            column += 2;
            if (column >= 128) { rtf += L"\n"; column = 0; }
        }
        rtf += L"}";
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

    std::wstring MakeRtf(const std::wstring& text,
                         const std::unordered_map<std::uint32_t, Picture>& pictures,
                         bool placeholders)
    {
        std::wstring rtf = L"{\\rtf1\\ansi\\deff0\\uc1\\viewkind4";
        rtf += L"{\\fonttbl{\\f0\\fnil Segoe UI;}}";
        rtf += L"\\f0\\fs20 ";

        for (std::size_t cp = 0; cp < text.size(); ++cp)
        {
            const wchar_t ch = text[cp];
            if (ch == 0x0001)
            {
                const auto it = pictures.find(static_cast<std::uint32_t>(cp));
                if (it != pictures.end()) AppendPictureRtf(rtf, it->second);
                else if (placeholders) rtf += L"[picture] ";
                continue;
            }
            switch (ch)
            {
            case L'\r': rtf += L"\\par\n"; break;
            case L'\n': break;
            case 0x0007: rtf += L"\\cell "; break;
            case 0x000B: rtf += L"\\line "; break;
            case 0x000C: rtf += L"\\page "; break;
            case 0x0008:
                if (placeholders) rtf += L"[object] ";
                break;
            case 0x0013: case 0x0014: case 0x0015: break;
            default: AppendRtfUnicode(rtf, ch); break;
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
        if (text.size() > static_cast<std::size_t>(std::numeric_limits<int>::max()))
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

    //const HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    //const bool uninitialize = SUCCEEDED(initHr);
    //if (FAILED(initHr) && initHr != RPC_E_CHANGED_MODE)
    //{
    //    result.error = L"CoInitializeEx failed.";
    //    return result;
    //}

    CComUsageScope scope;

    Storage storage;
    HRESULT hr = StgOpenStorage(fileName.c_str(), nullptr,
        STGM_READ | STGM_SHARE_DENY_WRITE, nullptr, 0, &storage.p);
    if (FAILED(hr))
    {
        //if (uninitialize) CoUninitialize();
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
        //if (uninitialize) CoUninitialize();
        result.error = L"The OLE file does not contain a WordDocument stream.";
        return result;
    }

    std::uint64_t wordSize = 0;
    if (!GetStreamSize(wordStream.p, wordSize) || wordSize > options.maxStreamBytes)
    {
        //if (uninitialize) CoUninitialize();
        result.error = L"WordDocument stream is invalid or exceeds the configured size limit.";
        return result;
    }

    std::vector<std::uint8_t> word;
    if (!ReadVector(wordStream.p, 0, wordSize, options.maxStreamBytes, word))
    {
        //if (uninitialize) CoUninitialize();
        result.error = L"Failed to read WordDocument stream.";
        return result;
    }

    FibInfo fib;
    if (!ParseFib(word, fib, result.error))
    {
        //if (uninitialize) CoUninitialize();
        return result;
    }

    const wchar_t* tableName = fib.table1 ? L"1Table" : L"0Table";
    if (FAILED(openStream(tableName, &tableStream.p)))
    {
        //if (uninitialize) CoUninitialize();
        result.error = L"Cannot open the table stream selected by the FIB.";
        return result;
    }

    std::uint64_t tableSize = 0;
    if (!GetStreamSize(tableStream.p, tableSize) || tableSize > options.maxStreamBytes)
    {
        //if (uninitialize) CoUninitialize();
        result.error = L"Table stream is invalid or exceeds the configured size limit.";
        return result;
    }

    std::vector<std::uint8_t> table;
    if (!ReadVector(tableStream.p, 0, tableSize, options.maxStreamBytes, table))
    {
        //if (uninitialize) CoUninitialize();
        result.error = L"Failed to read table stream.";
        return result;
    }

    Stream dataStream;
    std::vector<std::uint8_t> data;
    if (SUCCEEDED(openStream(L"Data", &dataStream.p)))
    {
        std::uint64_t dataSize = 0;
        if (GetStreamSize(dataStream.p, dataSize) && dataSize <= options.maxStreamBytes)
            ReadVector(dataStream.p, 0, dataSize, options.maxStreamBytes, data);
    }

    std::vector<Piece> pieces;
    if (!ParsePieceTable(table, fib.fcClx, fib.lcbClx, pieces, result.error))
    {
        //if (uninitialize) CoUninitialize();
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
            //if (uninitialize) CoUninitialize();
            return result;
        }

        const auto* src = word.data() + piece.fc;
        const bool ok = piece.compressed
            ? DecodeCompressed(src, static_cast<std::size_t>(byteLength), codePage, text)
            : DecodeUtf16(src, static_cast<std::size_t>(byteLength), text);

        if (!ok)
        {
            result.error = L"Failed to decode a text piece.";
            //if (uninitialize) CoUninitialize();
            return result;
        }
    }

    // Resolve inline picture characters (U+0001) through CHPX FKP records.
    // sprmCPicLocation (0x6A03) points into the Data stream to PICFAndOfficeArtData.
    std::unordered_map<std::uint32_t, Picture> pictures;
    std::vector<ChpxRun> chpxRuns;
    if (!data.empty() && ParseChpxRuns(word, table, fib, chpxRuns))
    {
        for (std::uint32_t cp = 0; cp < text.size(); ++cp)
        {
            if (text[cp] != 0x0001) continue;
            const auto pieceIt = std::find_if(pieces.begin(), pieces.end(),
                [cp](const Piece& piece) { return cp >= piece.cpStart && cp < piece.cpEnd; });
            if (pieceIt == pieces.end()) continue;
            const std::uint64_t delta = std::uint64_t(cp - pieceIt->cpStart);
            const std::uint64_t fc64 = std::uint64_t(pieceIt->fc) + delta * (pieceIt->compressed ? 1u : 2u);
            if (fc64 > 0xFFFFFFFFu) continue;
            const std::uint32_t fc = static_cast<std::uint32_t>(fc64);
            const auto runIt = std::find_if(chpxRuns.begin(), chpxRuns.end(),
                [fc](const ChpxRun& run) { return fc >= run.fcFirst && fc < run.fcLim; });
            if (runIt == chpxRuns.end()) continue;
            Picture picture;
            if (FindPictureBytes(data, runIt->picLocation, options.maxImageBytes, picture))
                pictures.emplace(cp, std::move(picture));
        }
    }

    // Rich Edit accepts PNG/JPEG payloads directly in RTF pict groups.
    std::wstring rtfText = MakeRtf(text, pictures, options.emitObjectPlaceholders);
    if (!WideToUtf8Bytes(rtfText, rtf))
    {
        result.error = L"Failed to encode generated RTF.";
        //if (uninitialize) CoUninitialize();
        return result;
    }

    result.success = true;
    result.charactersRead = text.size();
    result.pieceCount = static_cast<std::uint32_t>(pieces.size());
    result.imagesRead = static_cast<std::uint32_t>(pictures.size());
    if (pictures.size() < std::count(text.begin(), text.end(), wchar_t(0x0001)))
        result.warnings.push_back(L"Some inline pictures could not be decoded. This implementation currently embeds PNG and JPEG picture payloads; floating pictures and other OfficeArt formats are not handled.");
    result.warnings.push_back(L"Character/paragraph formatting, floating pictures, and complex OfficeArt objects are not implemented yet.");

    //if (uninitialize)
    //    CoUninitialize();
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
