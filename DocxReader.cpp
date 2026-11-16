#include "stdafx.h"
#include "DocxReader.h"

#include <windows.h>
#include <atlbase.h>
#include <msxml6.h>
#include <algorithm>
#include <cwctype>
#include <cwchar>
#include <iterator>
#include <cstdint>
#include <fstream>
#include <map>
#include <limits>
#include <sstream>
#include <utility>
#include <vector>
#include <thread>
#include <exception>
#include <set>
#include <zlib.h>

#pragma comment(lib, "msxml6.lib")

namespace docxreader {
    namespace {

        // In-memory archive filesystem map. Maps normalized relative virtual paths to file bytes.
        using MemoryArchive = std::map<std::wstring, std::vector<std::uint8_t>>;

        class ComApartment {
        public:
            ComApartment() : hr_(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)),
                uninitialize_(SUCCEEDED(hr_)) {
            }
            ~ComApartment() { if (uninitialize_) CoUninitialize(); }
            bool usable() const { return SUCCEEDED(hr_); }
            HRESULT result() const { return hr_; }
        private:
            HRESULT hr_;
            bool uninitialize_;
        };

        static std::uint16_t ReadLe16(const std::uint8_t* p)
        {
            return static_cast<std::uint16_t>(p[0] | (static_cast<unsigned>(p[1]) << 8));
        }

        static std::uint32_t ReadLe32(const std::uint8_t* p)
        {
            return static_cast<std::uint32_t>(p[0]) |
                (static_cast<std::uint32_t>(p[1]) << 8) |
                (static_cast<std::uint32_t>(p[2]) << 16) |
                (static_cast<std::uint32_t>(p[3]) << 24);
        }

        bool ReadWholeFile(const std::wstring& path, std::vector<std::uint8_t>& bytes,
            std::uint64_t maxBytes)
        {
            std::ifstream file(path, std::ios::binary | std::ios::ate);
            if (!file) return false;
            const std::streamoff size = file.tellg();
            if (size < 0 || static_cast<std::uint64_t>(size) > maxBytes) return false;
            bytes.resize(static_cast<size_t>(size));
            file.seekg(0, std::ios::beg);
            if (!bytes.empty()) {
                file.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
                if (file.gcount() != static_cast<std::streamsize>(bytes.size())) return false;
            }
            return true;
        }

        bool DecodeZipName(const std::string& name, bool utf8, std::wstring& decoded)
        {
            if (name.empty() || name.size() > 4096) return false;
            const UINT codePage = utf8 ? CP_UTF8 : 437;
            const DWORD flags = utf8 ? MB_ERR_INVALID_CHARS : 0;
            const int count = MultiByteToWideChar(codePage, flags, name.data(),
                static_cast<int>(name.size()), nullptr, 0);
            if (count <= 0) return false;
            decoded.resize(static_cast<size_t>(count));
            return MultiByteToWideChar(codePage, flags, name.data(),
                static_cast<int>(name.size()), &decoded[0], count) == count;
        }

        // Convert a ZIP path to a normalized relative format safely matching memory paths
        bool NormalizeZipPath(std::wstring& path, bool& isDirectory)
        {
            if (path.empty() || path[0] == L'/' || path[0] == L'\\') return false;
            std::replace(path.begin(), path.end(), L'/', L'\\');
            isDirectory = !path.empty() && path.back() == L'\\';
            if (isDirectory) path.pop_back();
            if (path.empty()) return false;

            std::wstring normalized;
            size_t pos = 0;
            while (pos < path.size()) {
                const size_t next = path.find(L'\\', pos);
                const std::wstring part = path.substr(pos, next == std::wstring::npos ? next : next - pos);
                if (part.empty() || part == L"." || part == L".." ||
                    part.back() == L'.' || part.back() == L' ')
                    return false;
                std::wstring device = part.substr(0, part.find(L'.'));
                std::transform(device.begin(), device.end(), device.begin(),
                    [](wchar_t c) { return static_cast<wchar_t>(std::towupper(c)); });
                if (device == L"CON" || device == L"PRN" || device == L"AUX" || device == L"NUL" ||
                    (device.size() == 4 && (device.compare(0, 3, L"COM") == 0 ||
                        device.compare(0, 3, L"LPT") == 0) &&
                        device[3] >= L'1' && device[3] <= L'9'))
                    return false;
                for (wchar_t c : part) {
                    if (c < 32 || c == L':' || c == L'<' || c == L'>' || c == L'"' ||
                        c == L'|' || c == L'?' || c == L'*') return false;
                }
                if (!normalized.empty()) normalized += L'\\';
                normalized += part;
                if (next == std::wstring::npos) break;
                pos = next + 1;
                if (pos == path.size()) return false;
            }
            path.swap(normalized);
            return true;
        }

        bool InflateZipEntry(const std::uint8_t* compressed, std::uint32_t compressedSize,
            std::uint32_t uncompressedSize, std::uint16_t method,
            std::vector<std::uint8_t>& output)
        {
            output.clear();
            if (uncompressedSize > 512u * 1024u * 1024u) return false;
            output.resize(uncompressedSize);

            if (method == 0) {
                if (compressedSize != uncompressedSize) return false;
                if (uncompressedSize) memcpy(output.data(), compressed, uncompressedSize);
                return true;
            }
            if (method != 8) return false;

            z_stream stream{};
            stream.next_in = const_cast<Bytef*>(reinterpret_cast<const Bytef*>(compressed));
            stream.avail_in = static_cast<uInt>(compressedSize);
            std::uint8_t emptyOutput = 0;
            stream.next_out = output.empty() ? &emptyOutput : reinterpret_cast<Bytef*>(output.data());
            stream.avail_out = output.empty() ? 1u : static_cast<uInt>(output.size());

            if (inflateInit2(&stream, -MAX_WBITS) != Z_OK) return false;
            const int status = inflate(&stream, Z_FINISH);
            const bool ok = status == Z_STREAM_END &&
                stream.total_in == compressedSize && stream.total_out == uncompressedSize;
            inflateEnd(&stream);
            return ok;
        }

        // Unzips raw bytes directly to memory using a standard key-value map layout.
        bool ExtractZipPackageToMemory(const std::wstring& zipPath, MemoryArchive& virtualFiles,
            std::uint64_t maxDocumentBytes, std::wstring& error)
        {
            std::vector<std::uint8_t> archive;
            if (maxDocumentBytes < 22 ||
                !ReadWholeFile(zipPath, archive, maxDocumentBytes) || archive.size() < 22) {
                error = L"Unable to read the DOCX ZIP package or the file exceeds the configured document-size limit.";
                return false;
            }

            const size_t tailStart = archive.size() > 22u + 65535u ? archive.size() - (22u + 65535u) : 0;
            size_t eocd = std::string::npos;
            for (size_t i = archive.size() - 22 + 1; i-- > tailStart;) {
                if (ReadLe32(archive.data() + i) == 0x06054b50u) {
                    const std::uint16_t commentLength = ReadLe16(archive.data() + i + 20);
                    if (i + 22u + commentLength == archive.size()) { eocd = i; break; }
                }
            }
            if (eocd == std::string::npos) {
                error = L"The DOCX ZIP end-of-directory record is missing or malformed.";
                return false;
            }

            const std::uint8_t* end = archive.data() + eocd;
            const std::uint16_t diskNumber = ReadLe16(end + 4);
            const std::uint16_t centralDisk = ReadLe16(end + 6);
            const std::uint16_t diskEntries = ReadLe16(end + 8);
            const std::uint16_t entryCount = ReadLe16(end + 10);
            const std::uint32_t centralSize = ReadLe32(end + 12);
            const std::uint32_t centralOffset = ReadLe32(end + 16);
            if (diskNumber || centralDisk || diskEntries != entryCount || !entryCount ||
                entryCount == 0xffffu || centralSize == 0xffffffffu || centralOffset == 0xffffffffu ||
                static_cast<std::uint64_t>(centralOffset) + centralSize > eocd) {
                error = L"The DOCX package uses a multi-disk, ZIP64 or otherwise unsupported ZIP layout.";
                return false;
            }

            std::set<std::wstring> seenNames;
            std::uint64_t totalExpanded = 0;
            size_t cursor = centralOffset;
            const size_t centralEnd = static_cast<size_t>(centralOffset) + centralSize;
            for (std::uint32_t index = 0; index < entryCount; ++index) {
                if (cursor > centralEnd || centralEnd - cursor < 46 ||
                    ReadLe32(archive.data() + cursor) != 0x02014b50u) {
                    error = L"The DOCX ZIP central directory is corrupt.";
                    return false;
                }
                const std::uint8_t* header = archive.data() + cursor;
                const std::uint16_t flags = ReadLe16(header + 8);
                const std::uint16_t method = ReadLe16(header + 10);
                const std::uint32_t expectedCrc = ReadLe32(header + 16);
                const std::uint32_t compressedSize = ReadLe32(header + 20);
                const std::uint32_t uncompressedSize = ReadLe32(header + 24);
                const std::uint16_t nameLength = ReadLe16(header + 28);
                const std::uint16_t extraLength = ReadLe16(header + 30);
                const std::uint16_t commentLength = ReadLe16(header + 32);
                const std::uint16_t startDisk = ReadLe16(header + 34);
                const std::uint32_t externalAttributes = ReadLe32(header + 38);
                const std::uint32_t localOffset = ReadLe32(header + 42);
                const std::uint64_t recordLength = 46ull + nameLength + extraLength + commentLength;
                if (!nameLength || nameLength > 4096 || recordLength > centralEnd - cursor ||
                    startDisk != 0 || (flags & (1u | 0x40u | 0x2000u)) || (method != 0 && method != 8) ||
                    compressedSize == 0xffffffffu || uncompressedSize == 0xffffffffu ||
                    localOffset == 0xffffffffu || uncompressedSize > 512u * 1024u * 1024u) {
                    error = L"The DOCX package contains an encrypted, ZIP64 or unsupported ZIP entry.";
                    return false;
                }

                const std::string rawName(reinterpret_cast<const char*>(header + 46), nameLength);
                std::wstring relativePath;
                if (!DecodeZipName(rawName, (flags & (1u << 11)) != 0, relativePath)) {
                    error = L"The DOCX package contains an invalid entry name.";
                    return false;
                }
                bool isDirectory = false;
                if (!NormalizeZipPath(relativePath, isDirectory)) {
                    error = L"The DOCX package contains an unsafe entry path.";
                    return false;
                }
                std::wstring folded = relativePath;
                std::transform(folded.begin(), folded.end(), folded.begin(),
                    [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
                if (!seenNames.insert(folded).second) {
                    error = L"The DOCX package contains duplicate entry paths.";
                    return false;
                }
                const std::uint32_t unixMode = externalAttributes >> 16;
                if ((unixMode & 0170000u) == 0120000u ||
                    (externalAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
                    error = L"The DOCX package contains a symbolic link or reparse point.";
                    return false;
                }
                if ((externalAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 ||
                    (unixMode & 0170000u) == 0040000u)
                    isDirectory = true;
                if (isDirectory && (compressedSize != 0 || uncompressedSize != 0)) {
                    error = L"A DOCX ZIP directory entry unexpectedly contains file data.";
                    return false;
                }
                if (!isDirectory) {
                    totalExpanded += uncompressedSize;
                    if (totalExpanded > maxDocumentBytes) {
                        error = L"The expanded DOCX package exceeds the configured document-size limit.";
                        return false;
                    }
                }
                if (localOffset >= centralOffset || localOffset > archive.size() ||
                    archive.size() - localOffset < 30 ||
                    ReadLe32(archive.data() + localOffset) != 0x04034b50u) {
                    error = L"A DOCX ZIP local file header is invalid.";
                    return false;
                }
                const std::uint8_t* local = archive.data() + localOffset;
                const std::uint16_t localFlags = ReadLe16(local + 6);
                const std::uint16_t localMethod = ReadLe16(local + 8);
                const std::uint16_t localNameLength = ReadLe16(local + 26);
                const std::uint16_t localExtraLength = ReadLe16(local + 28);
                const std::uint64_t dataOffset = static_cast<std::uint64_t>(localOffset) + 30u +
                    localNameLength + localExtraLength;
                if (localFlags != flags || localMethod != method || localNameLength != nameLength ||
                    dataOffset > centralOffset || compressedSize > centralOffset - dataOffset ||
                    memcmp(local + 30, rawName.data(), nameLength) != 0) {
                    error = L"A DOCX ZIP entry has inconsistent local and central headers.";
                    return false;
                }
                // Commit file bytes strictly to memory buffer map layout
                if (!isDirectory) {
                    std::vector<std::uint8_t> expanded;
                    if (!InflateZipEntry(archive.data() + static_cast<size_t>(dataOffset),
                        compressedSize, uncompressedSize, method, expanded)) {
                        error = L"Unable to decompress a DOCX ZIP entry.";
                        return false;
                    }
                    const uLong actualCrc = crc32(0L,
                        expanded.empty() ? Z_NULL : reinterpret_cast<const Bytef*>(expanded.data()),
                        static_cast<uInt>(expanded.size()));
                    if (static_cast<std::uint32_t>(actualCrc) != expectedCrc) {
                        error = L"A DOCX ZIP entry failed its CRC check.";
                        return false;
                    }
                    // Use lower case keys to handle paths case-insensitively across lookups
                    virtualFiles[folded] = std::move(expanded);
                }
                cursor += static_cast<size_t>(recordLength);
            }
            if (cursor != centralEnd) {
                error = L"The DOCX ZIP central directory has unexpected trailing data.";
                return false;
            }
            return true;
        }
        // Safe lookups inside in-memory buffers instead of disk file parsing
        std::wstring ReadTextFromMemory(const MemoryArchive& archive, const std::wstring& path)
        {
            std::wstring folded = path;
            std::transform(folded.begin(), folded.end(), folded.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
            auto it = archive.find(folded);
            if (it == archive.end() || it->second.empty()) return {};
            const auto& bytes = it->second;
            if (bytes.size() > static_cast<size_t>(128ull * 1024 * 1024)) return {};
            if (bytes.size() >= 2 && (unsigned char)bytes[0] == 0xFF && (unsigned char)bytes[1] == 0xFE) {
                if ((bytes.size() - 2) % 2) return {};
                std::wstring result((bytes.size() - 2) / 2, L'\0');
                if (!result.empty()) memcpy(&result[0], bytes.data() + 2, bytes.size() - 2);
                return result;
            }
            if (bytes.size() >= 2 && (unsigned char)bytes[0] == 0xFE && (unsigned char)bytes[1] == 0xFF) {
                if ((bytes.size() - 2) % 2) return {};
                std::wstring result((bytes.size() - 2) / 2, L'\0');
                for (size_t i = 0; i < result.size(); ++i)
                    result[i] = static_cast<wchar_t>((static_cast<unsigned char>(bytes[2 + i * 2]) << 8) |
                        static_cast<unsigned char>(bytes[3 + i * 2]));
                return result;
            }
            size_t startOffset = 0;
            if (bytes.size() >= 3 && (unsigned char)bytes[0] == 0xEF &&
                (unsigned char)bytes[1] == 0xBB && (unsigned char)bytes[2] == 0xBF) startOffset = 3;
            int n = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, reinterpret_cast<const char*>(bytes.data() + startOffset),
                (int)(bytes.size() - startOffset), nullptr, 0);
            if (!n) return {};
            std::wstring result(n, L'\0');
            MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, reinterpret_cast<const char*>(bytes.data() + startOffset),
                (int)(bytes.size() - startOffset), &result[0], n);
            return result;
        }
        bool LoadXmlFromMemory(const MemoryArchive& archive, const std::wstring& path, CComPtr<IXMLDOMDocument2>& doc)
        {
            doc.Release();
            HRESULT hr = CoCreateInstance(CLSID_DOMDocument60, nullptr, CLSCTX_INPROC_SERVER,
                IID_IXMLDOMDocument2, reinterpret_cast<void**>(&doc));
            if (FAILED(hr) || !doc) return false;
            doc->put_async(VARIANT_FALSE);
            doc->put_validateOnParse(VARIANT_FALSE);
            doc->put_resolveExternals(VARIANT_FALSE);
            doc->put_preserveWhiteSpace(VARIANT_TRUE);
            CComBSTR propertyName(L"ProhibitDTD");
            CComVariant propertyValue;
            propertyValue.vt = VT_BOOL;
            propertyValue.boolVal = VARIANT_TRUE;
            if (propertyName.m_str)
                doc->setProperty(propertyName, propertyValue);
            std::wstring xml = ReadTextFromMemory(archive, path);
            if (xml.empty() || xml.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
                return false;
            CComBSTR xmlText(static_cast<int>(xml.size()), xml.data());
            if (!xmlText.m_str) return false;
            VARIANT_BOOL loaded = VARIANT_FALSE;
            const HRESULT loadResult = doc->loadXML(xmlText, &loaded);
            return SUCCEEDED(loadResult) && loaded == VARIANT_TRUE;
        }
        std::wstring NodeName(IXMLDOMNode* node)
        {
            if (!node) return {};
            CComBSTR name;
            if (FAILED(node->get_nodeName(&name)) || !name.m_str) return {};
            std::wstring result(name.m_str, name.Length());
            const auto colon = result.find_last_of(L':');
            if (colon != std::wstring::npos) result.erase(0, colon + 1);
            return result;
        }
        std::wstring NodeText(IXMLDOMNode* node)
        {
            if (!node) return {};
            CComBSTR text;
            if (FAILED(node->get_text(&text)) || !text.m_str) return {};
            return std::wstring(text.m_str, text.Length());
        }
        std::wstring Attr(IXMLDOMNode* node, const wchar_t* name)
        {
            if (!node) return {};
            CComPtr<IXMLDOMNamedNodeMap> attrs;
            if (FAILED(node->get_attributes(&attrs)) || !attrs) return {};
            long count = 0;
            if (FAILED(attrs->get_length(&count))) return {};
            for (long i = 0; i < count; ++i) {
                CComPtr<IXMLDOMNode> attr;
                if (SUCCEEDED(attrs->get_item(i, &attr)) && attr && NodeName(attr.p) == name)
                    return NodeText(attr.p);
            }
            return {};
        }
        std::vector<CComPtr<IXMLDOMNode>> Children(IXMLDOMNode* node)
        {
            std::vector<CComPtr<IXMLDOMNode>> out;
            if (!node) return out;
            CComPtr<IXMLDOMNodeList> list;
            if (FAILED(node->get_childNodes(&list)) || !list) return out;
            long count = 0;
            if (FAILED(list->get_length(&count))) return out;
            out.reserve(static_cast<size_t>((std::max)(0L, count)));
            for (long i = 0; i < count; ++i) {
                CComPtr<IXMLDOMNode> child;
                if (SUCCEEDED(list->get_item(i, &child)) && child)
                    out.push_back(child);
            }
            return out;
        }
        std::wstring NormalizeTarget(const std::wstring& target)
        {
            std::wstring t = target;
            std::replace(t.begin(), t.end(), L'/', L'\\');
                if (!t.empty() && t[0] == L'\\') t.erase(t.begin());
                else t = L"word\\" + t;
                    std::vector<std::wstring> parts;
                    std::wstringstream ss(t); std::wstring part;
                    while (std::getline(ss, part, L'\\')) {
                        if (part.empty() || part == L".") continue;
            if (part == L"..") { if (!parts.empty()) parts.pop_back(); continue; }
            parts.push_back(part);
        }
        std::wstring out;
        for (const auto& p : parts) {
            if (!out.empty()) out += L"\\"; out += p; }
                return out;
        }
        std::wstring RtfText(const std::wstring& text)
        {
            std::wstring out;
            for (size_t i = 0; i < text.size(); ++i) {
                wchar_t c = text[i];
                if (c == L'\\' || c == L'{' || c == L'}') { out += L'\\'; out += c; }
                else if (c == L'\r') {}
                else if (c == L'\n') out += L"\\line ";
                else if (c == L'\t') out += L"\\tab ";
                else if (c < 0x20) {}
                else if (c <= 0x7f) out += c;
                else {
                    int value = (short)c;

                    // With \uc1, every \uN control word is followed by exactly
                    // one fallback character. Never consume the next source character:
                    // doing so corrupts text such as U+00E9 followed by "7".
                    out += L"\\u" + std::to_wstring(value) + L"?";
                }
            }
            return out;
        }
        bool ReadBinaryFromMemory(const MemoryArchive & archive, const std::wstring & path,
            std::uint64_t maxBytes, std::vector<std::uint8_t> & data)
        {
            std::wstring folded = path;
            std::transform(folded.begin(), folded.end(), folded.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
            auto it = archive.find(folded);
            if (it == archive.end()) return false;
            if (static_cast<std::uint64_t>(it->second.size()) > maxBytes) return false;
            data = it->second;
            return true;
        }
        std::wstring Hex(const std::vector<std::uint8_t> & bytes)
        {
            static const wchar_t digits[] = L"0123456789abcdef";
            std::wstring out; out.reserve(bytes.size() * 2);
            for (auto b : bytes) { out += digits[b >> 4]; out += digits[b & 15]; }
            return out;
        }
        struct StyleDef {
            std::wstring basedOn;
            std::wstring paragraphRtf;
            std::wstring runRtf;
            long numId = -1;
            long ilvl = 0;
        };

        struct NumberLevel {
            std::wstring format;
            std::wstring text;
            long start = 1;
            long left = 720;
            long hanging = 360;
        };

        std::wstring HexColor(const std::wstring& value)
        {
            if (value.size() != 6) return {};
            for (wchar_t c : value)
                if (!iswxdigit(c)) return {};
            std::wstring result = value;
            std::transform(result.begin(), result.end(), result.begin(),
                [](wchar_t c) { return static_cast<wchar_t>(std::towupper(c)); });
            return result;
        }

        std::wstring NormalizeWordColor(const std::wstring& value)
        {
            const std::wstring hex = HexColor(value);
            if (!hex.empty()) return hex;
            static const std::map<std::wstring, std::wstring> named = {
                {L"black",L"000000"},{L"blue",L"0000FF"},{L"cyan",L"00FFFF"},
                {L"darkblue",L"000080"},{L"darkcyan",L"008080"},{L"darkgray",L"808080"},
                {L"darkgreen",L"008000"},{L"darkmagenta",L"800080"},{L"darkred",L"800000"},
                {L"darkyellow",L"808000"},{L"green",L"00FF00"},{L"lightgray",L"C0C0C0"},
                {L"magenta",L"FF00FF"},{L"red",L"FF0000"},{L"white",L"FFFFFF"},
                {L"yellow",L"FFFF00"}
            };
            std::wstring key = value;
            std::transform(key.begin(), key.end(), key.begin(),
                [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
            auto it = named.find(key);
            return it == named.end() ? std::wstring() : it->second;
        }

        void CollectFonts(IXMLDOMNode* root, std::set<std::wstring>& fonts)
        {
            if (!root) return;
            std::vector<CComPtr<IXMLDOMNode>> pending;
            CComPtr<IXMLDOMNode> first = root;
            pending.push_back(first);
            while (!pending.empty()) {
                CComPtr<IXMLDOMNode> node = pending.back();
                pending.pop_back();
                if (NodeName(node.p) == L"rFonts") {
                    std::wstring name = Attr(node.p, L"ascii");
                    if (name.empty()) name = Attr(node.p, L"hAnsi");
                    if (name.empty()) name = Attr(node.p, L"eastAsia");
                    if (!name.empty() && name.size() <= 128) fonts.insert(name);
                }
                for (auto& child : Children(node.p)) pending.push_back(child);
            }
        }

        std::wstring RtfFontTable(const std::set<std::wstring>& fonts,
            std::map<std::wstring, int>& indexes)
        {
            std::wstring table = L"{\\fonttbl{\\f0\\fnil Segoe UI;}";
            int index = 1;
            for (const auto& font : fonts) {
                std::wstring safe;
                for (wchar_t c : font)
                    if (c >= 0x20 && c != L'{' && c != L'}' && c != L'\\' && c != L';') safe += c;
                if (safe.empty()) continue;
                table += L"{\\f" + std::to_wstring(index) + L"\\fnil " + safe + L";}";
                indexes[font] = index++;
            }
            table += L"}";
            return table;
        }

        void CollectColors(IXMLDOMNode* root, std::set<std::wstring>& colors)
        {
            if (!root) return;
            std::vector<CComPtr<IXMLDOMNode>> pending;
            CComPtr<IXMLDOMNode> first = root;
            pending.push_back(first);
            while (!pending.empty()) {
                CComPtr<IXMLDOMNode> node = pending.back();
                pending.pop_back();
                const std::wstring name = NodeName(node.p);
                if (name == L"color" || name == L"shd" || name == L"highlight" ||
                    name == L"top" || name == L"bottom" || name == L"left" ||
                    name == L"right" || name == L"insideH" || name == L"insideV" ||
                    name == L"start" || name == L"end") {
                    std::wstring color = NormalizeWordColor(Attr(node.p, L"val"));
                    if (color.empty()) color = NormalizeWordColor(Attr(node.p, L"fill"));
                    if (color.empty()) color = NormalizeWordColor(Attr(node.p, L"color"));
                    if (!color.empty()) colors.insert(color);
                }
                auto children = Children(node.p);
                for (auto& child : children) pending.push_back(child);
            }
        }

        std::wstring RtfColorTable(const std::set<std::wstring>& colors,
            std::map<std::wstring, int>& indexes)
        {
            std::wstring table = L"{\\colortbl;";
            int index = 1;
            for (const auto& color : colors) {
                const int r = std::stoi(color.substr(0, 2), nullptr, 16);
                const int g = std::stoi(color.substr(2, 2), nullptr, 16);
                const int b = std::stoi(color.substr(4, 2), nullptr, 16);
                table += L"\\red" + std::to_wstring(r) + L"\\green" +
                    std::to_wstring(g) + L"\\blue" + std::to_wstring(b) + L";";
                indexes[color] = index++;
            }
            table += L"}";
            return table;
        }

        std::wstring ColorControl(const std::wstring& value,
            const std::map<std::wstring, int>& colors, const wchar_t* control)
        {
            const std::wstring color = NormalizeWordColor(value);
            auto it = colors.find(color);
            return it == colors.end() ? std::wstring() :
                std::wstring(L"\\") + control + std::to_wstring(it->second) + L" ";
        }

        std::wstring RunPropertiesRtf(IXMLDOMNode* properties,
            const std::map<std::wstring, int>& colors,
            const std::map<std::wstring, int>& fonts)
        {
            std::wstring out;
            for (const auto& property : Children(properties)) {
                const std::wstring name = NodeName(property.p);
                const std::wstring value = Attr(property.p, L"val");
                const bool enabled = value != L"0" && value != L"false" && value != L"off";
                if (name == L"b") out += enabled ? L"\\b " : L"\\b0 ";
                else if (name == L"i") out += enabled ? L"\\i " : L"\\i0 ";
                else if (name == L"strike" || name == L"dstrike") out += enabled ? L"\\strike " : L"\\strike0 ";
                else if (name == L"u") out += (value.empty() || (value != L"none" && value != L"0")) ? L"\\ul " : L"\\ul0 ";
                else if (name == L"caps") out += enabled ? L"\\caps " : L"\\caps0 ";
                else if (name == L"smallCaps") out += enabled ? L"\\scaps " : L"\\scaps0 ";
                else if (name == L"vanish") out += enabled ? L"\\v " : L"\\v0 ";
                else if (name == L"sz" || name == L"szCs") {
                    wchar_t* end = nullptr;
                    const long n = std::wcstol(value.c_str(), &end, 10);
                    if (end && *end == L'\0' && n >= 4 && n <= 400) out += L"\\fs" + std::to_wstring(n) + L" ";
                }
                else if (name == L"vertAlign") {
                    if (value == L"superscript") out += L"\\super ";
                    else if (value == L"subscript") out += L"\\sub ";
                    else out += L"\\nosupersub ";
                }
                else if (name == L"color") out += ColorControl(value, colors, L"cf");
                else if (name == L"highlight") out += ColorControl(value, colors, L"highlight");
                else if (name == L"rFonts") {
                    std::wstring font = Attr(property.p, L"ascii");
                    if (font.empty()) font = Attr(property.p, L"hAnsi");
                    if (font.empty()) font = Attr(property.p, L"eastAsia");
                    auto it = fonts.find(font);
                    if (it != fonts.end()) out += L"\\f" + std::to_wstring(it->second) + L" ";
                }
            }
            return out;
        }

        std::wstring ParagraphPropertiesRtf(IXMLDOMNode* properties)
        {
            std::wstring out;
            for (const auto& property : Children(properties)) {
                const std::wstring name = NodeName(property.p);
                const std::wstring value = Attr(property.p, L"val");
                if (name == L"jc") {
                    if (value == L"center") out += L"\\qc ";
                    else if (value == L"right" || value == L"end") out += L"\\qr ";
                    else if (value == L"both" || value == L"distribute") out += L"\\qj ";
                    else out += L"\\ql ";
                }
                else if (name == L"ind") {
                    auto emit = [&](const wchar_t* attr, const wchar_t* control) {
                        const std::wstring v = Attr(property.p, attr);
                        if (!v.empty()) {
                            wchar_t* end = nullptr; long n = std::wcstol(v.c_str(), &end, 10);
                            if (end && *end == L'\0' && n >= -31680 && n <= 31680)
                                out += std::wstring(L"\\") + control + std::to_wstring(n) + L" ";
                        }
                    };
                    emit(L"left", L"li"); emit(L"right", L"ri");
                    emit(L"firstLine", L"fi");
                    const std::wstring hanging = Attr(property.p, L"hanging");
                    if (!hanging.empty()) {
                        wchar_t* end = nullptr; long n = std::wcstol(hanging.c_str(), &end, 10);
                        if (end && *end == L'\0' && n >= 0 && n <= 31680) out += L"\\fi-" + std::to_wstring(n) + L" ";
                    }
                }
                else if (name == L"spacing") {
                    auto emit = [&](const wchar_t* attr, const wchar_t* control) {
                        const std::wstring v = Attr(property.p, attr);
                        if (!v.empty()) {
                            wchar_t* end = nullptr; long n = std::wcstol(v.c_str(), &end, 10);
                            if (end && *end == L'\0' && n >= 0 && n <= 31680)
                                out += std::wstring(L"\\") + control + std::to_wstring(n) + L" ";
                        }
                    };
                    emit(L"before", L"sb"); emit(L"after", L"sa");
                    emit(L"line", L"sl");
                }
                else if (name == L"keepNext" && value != L"0") out += L"\\keepn ";
                else if (name == L"keepLines" && value != L"0") out += L"\\keep ";
                else if (name == L"pageBreakBefore" && value != L"0") out += L"\\pagebb ";
            }
            return out;
        }

        bool ParseLong(const std::wstring& value, long& number)
        {
            if (value.empty()) return false;
            wchar_t* end = nullptr;
            const long parsed = std::wcstol(value.c_str(), &end, 10);
            if (!end || *end != L'\0') return false;
            number = parsed;
            return true;
        }

        NumberLevel ParseNumberLevel(IXMLDOMNode* level)
        {
            NumberLevel result;
            for (const auto& child : Children(level)) {
                const std::wstring name = NodeName(child.p);
                if (name == L"numFmt") result.format = Attr(child.p, L"val");
                else if (name == L"lvlText") result.text = Attr(child.p, L"val");
                else if (name == L"start") { long n; if (ParseLong(Attr(child.p, L"val"), n)) result.start = (std::max)(1L, (std::min)(32767L, n)); }
                else if (name == L"pPr") {
                    for (const auto& p : Children(child.p)) if (NodeName(p.p) == L"ind") {
                        long n; if (ParseLong(Attr(p.p, L"left"), n)) result.left = n;
                        if (ParseLong(Attr(p.p, L"hanging"), n)) result.hanging = n;
                    }
                }
            }
            if (result.format.empty()) result.format = L"decimal";
            return result;
        }

        std::wstring FormatNumber(long value, const std::wstring& format)
        {
            if (format == L"bullet") return L"\x2022";
            if (format == L"decimalZero") {
                std::wstring s = std::to_wstring(value);
                return value < 10 ? L"0" + s : s;
            }
            if (format == L"lowerLetter" || format == L"upperLetter") {
                std::wstring s;
                long n = (std::max)(1L, value);
                while (n > 0) { --n; s.insert(s.begin(), static_cast<wchar_t>((format == L"upperLetter" ? L'A' : L'a') + n % 26)); n /= 26; }
                return s;
            }
            if (format == L"lowerRoman" || format == L"upperRoman") {
                static const long values[] = {1000,900,500,400,100,90,50,40,10,9,5,4,1};
                static const wchar_t* numerals[] = {L"m",L"cm",L"d",L"cd",L"c",L"xc",L"l",L"xl",L"x",L"ix",L"v",L"iv",L"i"};
                if (value <= 0 || value > 3999) return std::to_wstring(value);
                std::wstring out;
                for (size_t i=0; i<13; ++i) while (value >= values[i]) { out += numerals[i]; value -= values[i]; }
                if (format == L"upperRoman") std::transform(out.begin(), out.end(), out.begin(),
                    [](wchar_t c) { return static_cast<wchar_t>(std::towupper(c)); });
                return out;
            }
            return std::to_wstring(value);
        }

        struct ImageDimensions { long width = 0; long height = 0; };

        ImageDimensions ReadImageDimensions(const std::vector<std::uint8_t>& bytes, bool png)
        {
            ImageDimensions d;
            if (png && bytes.size() >= 24) {
                d.width = static_cast<long>((static_cast<std::uint32_t>(bytes[16]) << 24) | (static_cast<std::uint32_t>(bytes[17]) << 16) | (static_cast<std::uint32_t>(bytes[18]) << 8) | bytes[19]);
                d.height = static_cast<long>((static_cast<std::uint32_t>(bytes[20]) << 24) | (static_cast<std::uint32_t>(bytes[21]) << 16) | (static_cast<std::uint32_t>(bytes[22]) << 8) | bytes[23]);
            }
            else if (!png && bytes.size() > 4) {
                size_t i = 2;
                while (i + 4 < bytes.size()) {
                    if (bytes[i] != 0xFF) { ++i; continue; }
                    while (i < bytes.size() && bytes[i] == 0xFF) ++i;
                    if (i >= bytes.size()) break;
                    const std::uint8_t marker = bytes[i++];
                    if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7)) continue;
                    if (i + 2 > bytes.size()) break;
                    const std::uint16_t length = static_cast<std::uint16_t>((bytes[i] << 8) | bytes[i+1]);
                    if (length < 2 || i + length > bytes.size()) break;
                    const bool sof = (marker >= 0xC0 && marker <= 0xC3) || (marker >= 0xC5 && marker <= 0xC7) || (marker >= 0xC9 && marker <= 0xCB) || (marker >= 0xCD && marker <= 0xCF);
                    if (sof && length >= 7) { d.height = (bytes[i+3] << 8) | bytes[i+4]; d.width = (bytes[i+5] << 8) | bytes[i+6]; break; }
                    i += length;
                }
            }
            if (d.width <= 0 || d.height <= 0 || d.width > 100000 || d.height > 100000 || static_cast<std::uint64_t>(d.width) * d.height > 100000000ull) return {};
            return d;
        }

        struct ImportContext {
            const MemoryArchive& archive;
            std::map<std::wstring, std::wstring> rels;
            const Options& options;
            Result& result;
            std::map<std::wstring, int> colorIndexes;
            std::map<std::wstring, int> fontIndexes;
            std::map<std::wstring, StyleDef> styles;
            std::wstring defaultParagraphRtf;
            std::wstring defaultRunRtf;
            std::wstring defaultParagraphStyle;
            std::map<int, std::map<int, NumberLevel>> abstractNumbering;
            std::map<int, int> numberToAbstract;
            std::map<int, std::map<int, NumberLevel>> numLevelOverrides;
            std::map<std::pair<int, int>, long> startOverrides;
            std::map<int, std::vector<long>> numberCounters;
            std::wstring paragraphRunStyle;
            std::wstring rtf;
            ImportContext(const MemoryArchive& memArchive,
                std::map<std::wstring, std::wstring>&& relationships,
                const Options& importOptions, Result& importResult,
                const std::set<std::wstring>& colors, const std::set<std::wstring>& fonts)
                : archive(memArchive), rels(std::move(relationships)),
                options(importOptions), result(importResult) {
                rtf = L"{\\rtf1\\ansi\\deff0\\uc1" + RtfFontTable(fonts, fontIndexes) + RtfColorTable(colors, colorIndexes) + L"\n";
            }

            void LoadStyles(IXMLDOMNode* root)
            {
                for (const auto& node : Children(root)) {
                    if (NodeName(node.p) == L"docDefaults") {
                        for (const auto& defaults : Children(node.p)) {
                            if (NodeName(defaults.p) == L"pPrDefault") {
                                for (const auto& p : Children(defaults.p))
                                    if (NodeName(p.p) == L"pPr") defaultParagraphRtf = ParagraphPropertiesRtf(p.p);
                            }
                            else if (NodeName(defaults.p) == L"rPrDefault") {
                                for (const auto& r : Children(defaults.p))
                                    if (NodeName(r.p) == L"rPr") defaultRunRtf = RunPropertiesRtf(r.p, colorIndexes, fontIndexes);
                            }
                        }
                        continue;
                    }
                    if (NodeName(node.p) != L"style") continue;
                    const std::wstring id = Attr(node.p, L"styleId");
                    if (id.empty()) continue;
                    StyleDef def;
                    if (Attr(node.p, L"default") == L"1" && Attr(node.p, L"type") == L"paragraph") defaultParagraphStyle = id;
                    for (const auto& property : Children(node.p)) {
                        const std::wstring name = NodeName(property.p);
                        if (name == L"basedOn") def.basedOn = Attr(property.p, L"val");
                        else if (name == L"pPr") {
                            def.paragraphRtf = ParagraphPropertiesRtf(property.p);
                            for (const auto& pProperty : Children(property.p)) {
                                if (NodeName(pProperty.p) == L"rPr") def.runRtf += RunPropertiesRtf(pProperty.p, colorIndexes, fontIndexes);
                                else if (NodeName(pProperty.p) == L"numPr") {
                                    for (const auto& n : Children(pProperty.p)) {
                                        if (NodeName(n.p) == L"numId") ParseLong(Attr(n.p, L"val"), def.numId);
                                        else if (NodeName(n.p) == L"ilvl") ParseLong(Attr(n.p, L"val"), def.ilvl);
                                    }
                                }
                            }
                        }
                        else if (name == L"rPr") def.runRtf += RunPropertiesRtf(property.p, colorIndexes, fontIndexes);
                    }
                    styles[id] = std::move(def);
                }
            }

            void LoadNumbering(IXMLDOMNode* root)
            {
                for (const auto& node : Children(root)) {
                    const std::wstring name = NodeName(node.p);
                    if (name == L"abstractNum") {
                        long idValue;
                        if (!ParseLong(Attr(node.p, L"abstractNumId"), idValue) || idValue < 0 || idValue > 1000000) continue;
                        const int abstractId = static_cast<int>(idValue);
                        for (const auto& level : Children(node.p)) {
                            if (NodeName(level.p) != L"lvl") continue;
                            long ilvl = 0;
                            if (!ParseLong(Attr(level.p, L"ilvl"), ilvl) || ilvl < 0 || ilvl > 8) continue;
                            abstractNumbering[abstractId][static_cast<int>(ilvl)] = ParseNumberLevel(level.p);
                        }
                    }
                    else if (name == L"num") {
                        long numId = -1;
                        if (!ParseLong(Attr(node.p, L"numId"), numId) || numId < 0 || numId > 1000000) continue;
                        for (const auto& child : Children(node.p)) {
                            if (NodeName(child.p) == L"abstractNumId") {
                                long abstractId;
                                if (ParseLong(Attr(child.p, L"val"), abstractId)) numberToAbstract[static_cast<int>(numId)] = static_cast<int>(abstractId);
                            }
                            else if (NodeName(child.p) == L"lvlOverride") {
                                long ilvl = 0;
                                if (!ParseLong(Attr(child.p, L"ilvl"), ilvl) || ilvl < 0 || ilvl > 8) continue;
                                for (const auto& over : Children(child.p)) {
                                    if (NodeName(over.p) == L"startOverride") {
                                        long start;
                                        if (ParseLong(Attr(over.p, L"val"), start)) startOverrides[{static_cast<int>(numId), static_cast<int>(ilvl)}] = (std::max)(1L, start);
                                    }
                                    else if (NodeName(over.p) == L"lvl") numLevelOverrides[static_cast<int>(numId)][static_cast<int>(ilvl)] = ParseNumberLevel(over.p);
                                }
                            }
                        }
                    }
                }
            }

            std::pair<std::wstring, std::wstring> ResolveStyle(const std::wstring& id,
                std::set<std::wstring>& visiting, int depth = 0) const
            {
                if (id.empty() || depth > 32 || !visiting.insert(id).second) return {};
                auto it = styles.find(id);
                if (it == styles.end()) { visiting.erase(id); return {}; }
                std::pair<std::wstring, std::wstring> resultPair;
                if (!it->second.basedOn.empty()) resultPair = ResolveStyle(it->second.basedOn, visiting, depth + 1);
                resultPair.first += it->second.paragraphRtf;
                resultPair.second += it->second.runRtf;
                visiting.erase(id);
                return resultPair;
            }

            std::pair<std::wstring, std::wstring> StyleRtf(const std::wstring& id) const
            {
                std::set<std::wstring> visiting;
                return ResolveStyle(id, visiting);
            }

            std::wstring NumberPrefix(IXMLDOMNode* paragraph, const std::wstring& paragraphStyle)
            {
                CComPtr<IXMLDOMNode> pPr;
                for (const auto& child : Children(paragraph)) if (NodeName(child.p) == L"pPr") { pPr = child.p; break; }
                long numId = -1, ilvl = 0;
                if (pPr) {
                    for (const auto& child : Children(pPr.p)) if (NodeName(child.p) == L"numPr") {
                        for (const auto& item : Children(child.p)) {
                            if (NodeName(item.p) == L"numId") ParseLong(Attr(item.p, L"val"), numId);
                            else if (NodeName(item.p) == L"ilvl") ParseLong(Attr(item.p, L"val"), ilvl);
                        }
                    }
                }
                if (numId < 0) {
                    std::wstring current = paragraphStyle;
                    std::set<std::wstring> visited;
                    for (int depth = 0; !current.empty() && depth < 32 && visited.insert(current).second; ++depth) {
                        auto styleIt = styles.find(current);
                        if (styleIt == styles.end()) break;
                        if (styleIt->second.numId >= 0) { numId = styleIt->second.numId; ilvl = styleIt->second.ilvl; break; }
                        current = styleIt->second.basedOn;
                    }
                }
                if (numId <= 0 || ilvl < 0 || ilvl > 8) return {};
                auto num = numberToAbstract.find(static_cast<int>(numId));
                if (num == numberToAbstract.end()) return {};
                auto abstract = abstractNumbering.find(num->second);
                if (abstract == abstractNumbering.end()) return {};
                const NumberLevel* currentLevel = nullptr;
                auto numberOverride = numLevelOverrides.find(static_cast<int>(numId));
                if (numberOverride != numLevelOverrides.end()) {
                    auto overrideLevel = numberOverride->second.find(static_cast<int>(ilvl));
                    if (overrideLevel != numberOverride->second.end()) currentLevel = &overrideLevel->second;
                }
                if (!currentLevel) {
                    auto level = abstract->second.find(static_cast<int>(ilvl));
                    if (level == abstract->second.end()) return {};
                    currentLevel = &level->second;
                }

                auto& counters = numberCounters[static_cast<int>(numId)];
                if (counters.size() < 9) counters.resize(9, 0);
                long start = currentLevel->start;
                auto overrideIt = startOverrides.find({ static_cast<int>(numId), static_cast<int>(ilvl) });
                if (overrideIt != startOverrides.end()) start = overrideIt->second;
                if (counters[ilvl] == 0) counters[ilvl] = start;
                else ++counters[ilvl];
                for (int lower = static_cast<int>(ilvl) + 1; lower < 9; ++lower) counters[lower] = 0;

                std::wstring label = currentLevel->format == L"bullet" ? L"\x2022" : currentLevel->text;
                for (int n = 0; n < 9; ++n) {
                    const std::wstring token = L"%" + std::to_wstring(n + 1);
                    size_t pos = 0;
                    while ((pos = label.find(token, pos)) != std::wstring::npos) {
                        const NumberLevel* formatLevel = nullptr;
                        if (numberOverride != numLevelOverrides.end()) {
                            auto ov = numberOverride->second.find(n);
                            if (ov != numberOverride->second.end()) formatLevel = &ov->second;
                        }
                        if (!formatLevel) {
                            auto base = abstract->second.find(n);
                            if (base != abstract->second.end()) formatLevel = &base->second;
                        }
                        const std::wstring fmt = formatLevel ? formatLevel->format : L"decimal";
                        long count = counters[n];
                        if (count == 0) count = formatLevel ? formatLevel->start : 1;
                        const std::wstring replacement = FormatNumber(count, fmt);
                        label.replace(pos, token.size(), replacement);
                        pos += replacement.size();
                    }
                }
                if (label.empty()) label = FormatNumber(counters[ilvl], currentLevel->format) + L".";
                const long left = (std::max)(0L, (std::min)(31680L, currentLevel->left));
                const long hanging = (std::max)(0L, (std::min)(31680L, currentLevel->hanging));
                return L"\\li" + std::to_wstring(left) + L"\\fi-" + std::to_wstring(hanging) + L" " + RtfText(label) + L"\\tab ";
            }
            bool Image(IXMLDOMNode* drawing)
            {
                if (!options.importImages) return false;
                std::wstring rid, ext;
                long widthTwips = 0, heightTwips = 0;
                bool widthFound = false, heightFound = false;
                auto stack = Children(drawing);
                while (!stack.empty()) {
                    auto node = stack.back();
                    stack.pop_back();
                    const auto nodeName = NodeName(node.p);
                    if (nodeName == L"blip") rid = Attr(node.p, L"embed");
                    else if (nodeName == L"extent" && (!widthFound || !heightFound)) {
                        const auto cx = Attr(node.p, L"cx"), cy = Attr(node.p, L"cy");
                        if (!widthFound && !cx.empty()) { widthTwips = (std::max)(1L, (std::min)(20000L, _wtol(cx.c_str()) / 635)); widthFound = true; }
                        if (!heightFound && !cy.empty()) { heightTwips = (std::max)(1L, (std::min)(20000L, _wtol(cy.c_str()) / 635)); heightFound = true; }
                    }
                    auto children = Children(node.p);
                    // Push in reverse so the depth-first walk follows XML order;
                    // wp:extent is encountered before the nested a:xfrm extent.
                    for (auto it = children.rbegin(); it != children.rend(); ++it) stack.push_back(*it);
                }
                if (rid.empty()) return false;
                auto it = rels.find(rid);
                if (it == rels.end()) return false;
                std::wstring path = it->second;
                std::vector<std::uint8_t> bytes;
                if (!ReadBinaryFromMemory(archive, path, options.maxImageBytes, bytes) || bytes.empty()) return false;
                const bool isPng = bytes.size() >= 8 &&
                    bytes[0] == 0x89 && bytes[1] == 'P' && bytes[2] == 'N' && bytes[3] == 'G' &&
                    bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A;
                const bool isJpeg = bytes.size() >= 3 &&
                    bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF;
                if (isPng) ext = L"\\pngblip";
                else if (isJpeg) ext = L"\\jpegblip";
                else return false;
                const ImageDimensions intrinsic = ReadImageDimensions(bytes, isPng);
                const long pixelWidth = intrinsic.width > 0 ? intrinsic.width : (std::max)(1L, widthTwips / 15);
                const long pixelHeight = intrinsic.height > 0 ? intrinsic.height : (std::max)(1L, heightTwips / 15);
                if (!widthFound && intrinsic.width > 0) {
                    widthTwips = heightFound && pixelHeight > 0
                        ? static_cast<long>((static_cast<std::uint64_t>(heightTwips) * pixelWidth) / pixelHeight)
                        : (std::min)(20000L, pixelWidth * 15L);
                }
                if (!heightFound && intrinsic.height > 0) {
                    heightTwips = widthFound && pixelWidth > 0
                        ? static_cast<long>((static_cast<std::uint64_t>(widthTwips) * pixelHeight) / pixelWidth)
                        : (std::min)(20000L, pixelHeight * 15L);
                }
                if (widthTwips <= 0) widthTwips = 3000;
                if (heightTwips <= 0) heightTwips = 2000;
                rtf += L"{\\pict" + ext + L"\\picw" + std::to_wstring(pixelWidth) +
                    L"\\pich" + std::to_wstring(pixelHeight) +
                    L"\\picwgoal" + std::to_wstring(widthTwips) +
                    L"\\pichgoal" + std::to_wstring(heightTwips) + L"\n" + Hex(bytes) + L"}";
                ++result.imagesRead;
                return true;
            }
            void Run(IXMLDOMNode* run)
            {
                rtf += L"{" + paragraphRunStyle;
                CComPtr<IXMLDOMNode> properties;
                auto children = Children(run);
                for (const auto& child : children) {
                    if (NodeName(child.p) == L"rPr") { properties = child.p; break; }
                }
                if (properties) {
                    for (const auto& property : Children(properties.p)) {
                        if (NodeName(property.p) == L"rStyle") {
                            const auto style = StyleRtf(Attr(property.p, L"val"));
                            rtf += style.second;
                        }
                    }
                    rtf += RunPropertiesRtf(properties.p, colorIndexes, fontIndexes);
                }
                for (const auto& child : children) {
                    const auto name = NodeName(child.p);
                    if (name == L"t" || name == L"delText" || name == L"instrText") rtf += RtfText(NodeText(child.p));
                    else if (name == L"tab") rtf += L"\\tab ";
                    else if (name == L"br" || name == L"cr") rtf += (Attr(child.p, L"type") == L"page" ? L"\\page " : L"\\line ");
                    else if (name == L"drawing" || name == L"pict") Image(child.p);
                }
                // Run formatting is scoped by the RTF group; do not issue \\plain,
                // which would reset inherited formatting inside this run.
                rtf += L"}";
            }
            void Paragraph(IXMLDOMNode* paragraph, bool inTable = false)
            {
                rtf += inTable ? L"\\pard\\intbl" : L"\\pard";
                auto children = Children(paragraph);
                CComPtr<IXMLDOMNode> paragraphProperties;
                std::wstring styleId;
                for (const auto& child : children) if (NodeName(child.p) == L"pPr") { paragraphProperties = child.p; break; }
                if (paragraphProperties) {
                    for (const auto& property : Children(paragraphProperties.p)) {
                        if (NodeName(property.p) == L"pStyle") styleId = Attr(property.p, L"val");
                    }
                }
                if (styleId.empty()) styleId = defaultParagraphStyle;
                const auto style = StyleRtf(styleId);
                rtf += defaultParagraphRtf + style.first;
                if (paragraphProperties) rtf += ParagraphPropertiesRtf(paragraphProperties.p);
                const std::wstring oldParagraphRunStyle = paragraphRunStyle;
                paragraphRunStyle = defaultRunRtf + style.second;
                if (paragraphProperties) {
                    for (const auto& property : Children(paragraphProperties.p))
                        if (NodeName(property.p) == L"rPr") paragraphRunStyle += RunPropertiesRtf(property.p, colorIndexes, fontIndexes);
                }
                rtf += NumberPrefix(paragraph, styleId);
                for (const auto& child : children) {
                    const auto name = NodeName(child.p);
                    if (name == L"r") Run(child.p);
                    else if (name == L"hyperlink") Hyperlink(child.p);
                    else if (name == L"sdt") {
                        for (const auto& sdtChild : Children(child.p)) {
                            if (NodeName(sdtChild.p) == L"sdtContent")
                                for (const auto& content : Children(sdtChild.p)) {
                                    if (NodeName(content.p) == L"r") Run(content.p);
                                    else if (NodeName(content.p) == L"hyperlink") Hyperlink(content.p);
                                }
                        }
                    }
                }
                paragraphRunStyle = oldParagraphRunStyle;
                rtf += L"\\par\n";
            }
            void Hyperlink(IXMLDOMNode* link)
            {
                auto appendRuns = [&]() {
                auto children = Children(link);
                for (const auto& child : children)
                if (NodeName(child.p) == L"r") Run(child.p);
                };
                if (!options.importHyperlinks) {
                    appendRuns();
                    return;
                }
                std::wstring rid = Attr(link, L"id"), target;
                auto it = rels.find(rid);
                if (it != rels.end()) target = it->second;
                const auto colon = target.find(L':');
                std::wstring scheme = colon == std::wstring::npos ? L"" : target.substr(0, colon);
                std::transform(scheme.begin(), scheme.end(), scheme.begin(),
                    [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
                const bool safeScheme = scheme == L"http" || scheme == L"https" ||
                    scheme == L"mailto" || scheme == L"ftp";
                if (!safeScheme || target.find_first_of(L"\r\n") != std::wstring::npos ||
                    target.find(L'\0') != std::wstring::npos) {
                    appendRuns();
                    return;
                }
                std::wstring fieldTarget = RtfText(target);
                size_t quote = 0;
                while ((quote = fieldTarget.find(L'"', quote)) != std::wstring::npos) {
                    fieldTarget.replace(quote, 1, L"\\'22");
                    quote += 4;
                }
                rtf += L"{\\field{\\*\\fldinst HYPERLINK \"" + fieldTarget + L"\"}{\\fldrslt ";
                appendRuns();
                rtf += L"}}";
                ++result.hyperlinksRead;
            }
            long CellWidthTwips(IXMLDOMNode* cell) const
            {
                for (const auto& child : Children(cell)) {
                    if (NodeName(child.p) != L"tcPr") continue;
                    for (const auto& property : Children(child.p)) {
                        if (NodeName(property.p) != L"tcW" || Attr(property.p, L"type") == L"pct")
                            continue;
                        const std::wstring value = Attr(property.p, L"w");
                        if (value.empty()) continue;
                        wchar_t* end = nullptr;
                        const long width = std::wcstol(value.c_str(), &end, 10);
                        if (end && *end == L'\0' && width >= 100 && width <= 20000)
                            return width;
                    }
                }
                return 0; // caller may use tblGrid widths before applying its fallback
            }
            std::wstring BorderRtf(IXMLDOMNode* border, const wchar_t* side)
            {
                const std::wstring value = Attr(border, L"val");
                if (value.empty() || value == L"nil" || value == L"none") return {};
                const wchar_t* sideCode = std::wcscmp(side, L"top") == 0 ? L"t" :
                    std::wcscmp(side, L"bottom") == 0 ? L"b" :
                    std::wcscmp(side, L"left") == 0 ? L"l" : L"r";
                std::wstring control = L"\\clbrdr" + std::wstring(sideCode);
                if (value == L"double") control += L"\\brdrdb";
                else if (value == L"dotted") control += L"\\brdrdot";
                else if (value == L"dashed" || value == L"dashSmallGap") control += L"\\brdrdash";
                else if (value == L"thick") control += L"\\brdrth";
                else if (value == L"wave") control += L"\\brdrwavy";
                else control += L"\\brdrs";
                long size = 8;
                ParseLong(Attr(border, L"sz"), size);
                size = (std::max)(1L, (std::min)(96L, size));
                control += L"\\brdrw" + std::to_wstring(size * 2);
                control += ColorControl(Attr(border, L"color"), colorIndexes, L"brdrcf");
                return control;
            }

            int CellSpan(IXMLDOMNode* cell) const
            {
                for (const auto& child : Children(cell)) if (NodeName(child.p) == L"tcPr")
                    for (const auto& property : Children(child.p)) if (NodeName(property.p) == L"gridSpan") {
                        long span = 1;
                        if (ParseLong(Attr(property.p, L"val"), span)) return static_cast<int>((std::max)(1L, (std::min)(63L, span)));
                    }
                return 1;
            }

            std::wstring CellPropertiesRtf(IXMLDOMNode* cell, bool horizontalFirst,
                bool horizontalContinue, const std::map<std::wstring, std::wstring>& tableBorders)
            {
                std::wstring out;
                std::set<std::wstring> explicitBorders;
                if (horizontalFirst) out += L"\\clmgf";
                else if (horizontalContinue) out += L"\\clmrg";
                for (const auto& child : Children(cell)) {
                    if (NodeName(child.p) != L"tcPr") continue;
                    for (const auto& property : Children(child.p)) {
                        const std::wstring name = NodeName(property.p);
                        if (name == L"vMerge") {
                            const std::wstring value = Attr(property.p, L"val");
                            out += (value == L"restart") ? L"\\clvmgf" : L"\\clvmrg";
                        }
                        else if (name == L"vAlign") {
                            const std::wstring value = Attr(property.p, L"val");
                            if (value == L"center") out += L"\\clvertalc";
                            else if (value == L"bottom") out += L"\\clvertalb";
                            else out += L"\\clvertalt";
                        }
                        else if (name == L"shd") {
                            const std::wstring fill = Attr(property.p, L"fill");
                            out += ColorControl(fill, colorIndexes, L"clcbpat");
                        }
                        else if (name == L"tcBorders") {
                            for (const auto& border : Children(property.p)) {
                                const std::wstring side = NodeName(border.p);
                                if (side == L"top" || side == L"bottom" || side == L"left" || side == L"right") {
                                    explicitBorders.insert(side);
                                    out += BorderRtf(border.p, side.c_str());
                                }
                            }
                        }
                        else if (name == L"tcMar") {
                            for (const auto& margin : Children(property.p)) {
                                const std::wstring side = NodeName(margin.p);
                                const std::wstring value = Attr(margin.p, L"w");
                                long n = 0;
                                if (!ParseLong(value, n) || n < 0 || n > 31680) continue;
                                if (side == L"top" || side == L"bottom" || side == L"left" || side == L"right") {
                                    const wchar_t* code = side == L"top" ? L"t" : side == L"bottom" ? L"b" : side == L"left" ? L"l" : L"r";
                                    out += L"\\clpad" + std::wstring(code) + std::to_wstring(n) + L"\\clpadf" + code + L"3";
                                }
                            }
                        }
                    }
                }
                for (const auto& border : tableBorders)
                    if (explicitBorders.find(border.first) == explicitBorders.end()) out += border.second;
                return out;
            }

            void Table(IXMLDOMNode* table)
            {
                std::map<std::wstring, std::wstring> tableBorders;
                for (const auto& property : Children(table)) if (NodeName(property.p) == L"tblPr")
                    for (const auto& tblProperty : Children(property.p)) if (NodeName(tblProperty.p) == L"tblBorders")
                        for (const auto& border : Children(tblProperty.p)) {
                            const std::wstring side = NodeName(border.p);
                            if (side == L"top" || side == L"bottom" || side == L"left" || side == L"right")
                                tableBorders[side] = BorderRtf(border.p, side.c_str());
                            else if (side == L"insideH") tableBorders[side] = BorderRtf(border.p, L"bottom");
                            else if (side == L"insideV") tableBorders[side] = BorderRtf(border.p, L"right");
                        }
                std::vector<long> gridWidths;
                for (const auto& child : Children(table)) if (NodeName(child.p) == L"tblGrid")
                    for (const auto& column : Children(child.p)) if (NodeName(column.p) == L"gridCol") {
                        long width = 0;
                        if (ParseLong(Attr(column.p, L"w"), width) && width >= 100 && width <= 20000) gridWidths.push_back(width);
                    }
                std::vector<CComPtr<IXMLDOMNode>> tableRows;
                for (const auto& row : Children(table)) if (NodeName(row.p) == L"tr") tableRows.push_back(row);
                for (size_t rowIndex = 0; rowIndex < tableRows.size(); ++rowIndex) {
                    const auto& row = tableRows[rowIndex];
                    struct CellInfo { CComPtr<IXMLDOMNode> node; int span = 1; long width = 1800; bool hFirst = false; bool hContinue = false; };
                    std::vector<CellInfo> rowCells;
                    size_t gridIndex = 0;
                    for (const auto& cell : Children(row.p)) {
                        if (NodeName(cell.p) != L"tc") continue;
                        CellInfo info; info.node = cell.p; info.span = CellSpan(cell.p); info.width = CellWidthTwips(cell.p);
                        if (info.width <= 0 && !gridWidths.empty()) {
                            for (int part = 0; part < info.span && gridIndex + part < gridWidths.size(); ++part) info.width += gridWidths[gridIndex + part];
                        }
                        if (info.width <= 0) info.width = 1800L * info.span;
                        gridIndex += static_cast<size_t>(info.span);
                        for (const auto& tcPr : Children(cell.p)) if (NodeName(tcPr.p) == L"tcPr")
                            for (const auto& prop : Children(tcPr.p)) if (NodeName(prop.p) == L"hMerge") {
                                const std::wstring value = Attr(prop.p, L"val");
                                info.hFirst = value == L"restart";
                                info.hContinue = value != L"restart";
                            }
                        rowCells.push_back(std::move(info));
                    }
                    rtf += L"\\trowd\\trgaph108\\trleft0";
                    for (const auto& prop : Children(row.p)) if (NodeName(prop.p) == L"trPr")
                        for (const auto& setting : Children(prop.p)) {
                            if (NodeName(setting.p) == L"cantSplit" && Attr(setting.p, L"val") != L"0" && Attr(setting.p, L"val") != L"false") rtf += L"\\trkeep";
                            else if (NodeName(setting.p) == L"tblHeader" && Attr(setting.p, L"val") != L"0" && Attr(setting.p, L"val") != L"false") rtf += L"\\trhdr";
                        }
                    long x = 0;
                    for (size_t cellIndex = 0; cellIndex < rowCells.size(); ++cellIndex) {
                        const auto& cell = rowCells[cellIndex];
                        const long partWidth = (std::max)(1L, cell.width / cell.span);
                        std::map<std::wstring, std::wstring> cellBorders;
                        auto copyBorder = [&](const wchar_t* source, const wchar_t* destination) {
                            auto it = tableBorders.find(source);
                            if (it != tableBorders.end()) cellBorders[destination] = it->second;
                        };
                        if (rowIndex == 0) copyBorder(L"top", L"top");
                        if (rowIndex + 1 == tableRows.size()) copyBorder(L"bottom", L"bottom");
                        if (rowIndex + 1 < tableRows.size()) copyBorder(L"insideH", L"bottom");
                        if (cellIndex == 0) copyBorder(L"left", L"left");
                        if (cellIndex + 1 == rowCells.size()) copyBorder(L"right", L"right");
                        if (cellIndex + 1 < rowCells.size()) copyBorder(L"insideV", L"right");
                        for (int part = 0; part < cell.span; ++part) {
                            const bool mergedFirst = cell.span > 1 ? part == 0 : cell.hFirst;
                            const bool mergedContinue = cell.span > 1 ? part > 0 : cell.hContinue;
                            rtf += CellPropertiesRtf(cell.node.p, mergedFirst, mergedContinue, cellBorders);
                            x = (std::min)(1000000L, x + partWidth);
                            rtf += L"\\cellx" + std::to_wstring(x);
                        }
                    }
                    rtf += L"\\n";
                    for (const auto& cell : rowCells) {
                        rtf += L"\\intbl ";
                        auto blocks = Children(cell.node.p);
                        for (const auto& block : blocks) {
                            const auto name = NodeName(block.p);
                            if (name == L"p") Paragraph(block.p, true);
                            else if (name == L"tbl") Table(block.p);
                        }
                        for (int part = 0; part < cell.span; ++part) rtf += L"\\cell ";
                    }
                    rtf += L"\\row\\n";
                }
            }
        };
        bool LoadRelationshipsFromMemory(const MemoryArchive& archive, const std::wstring& path, std::map<std::wstring, std::wstring>& rels)
        {
            CComPtr<IXMLDOMDocument2> doc;
            if (!LoadXmlFromMemory(archive, path, doc)) return false;
            CComPtr<IXMLDOMElement> root;
            if (FAILED(doc->get_documentElement(&root)) || !root) return false;
            auto children = Children(root.p);
            for (const auto& node : children) {
                if (NodeName(node.p) == L"Relationship") {
                    std::wstring id = Attr(node.p, L"Id");
                    std::wstring target = Attr(node.p, L"Target");
                    std::wstring mode = Attr(node.p, L"TargetMode");
                    if (!id.empty() && !target.empty() && mode != L"External")
                        rels[id] = NormalizeTarget(target);
                    else if (!id.empty() && !target.empty() && mode == L"External")
                        rels[id] = target;
                }
            }
            return true;
        }
} // namespace
static Result ConvertToRtfOnSta(const std::wstring& fileName,
    std::vector<std::uint8_t>& rtf,
    const Options& options)
{
    Result result;
    MemoryArchive archive;
    if (options.maxDocumentBytes < 22 || options.maxImageBytes == 0) {
        result.error = L"The DOCX import size limits are invalid.";
        return result;
    }

    // Keep both the compressed package and total expanded payload within the
    // caller's configured budget. No temporary extraction directory is used.
    if (!ExtractZipPackageToMemory(fileName, archive, options.maxDocumentBytes, result.error))
        return result;
    const std::wstring documentPath = L"word\\document.xml";
    const std::wstring relsPath = L"word\\_rels\\document.xml.rels";
    // Read and parse directly out of memory
    std::wstring foldedDocPath = documentPath;
    std::transform(foldedDocPath.begin(), foldedDocPath.end(), foldedDocPath.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    if (archive.find(foldedDocPath) == archive.end()) {
        result.error = L"The DOCX package does not contain a readable word/document.xml.";
        return result;
    }
    std::map<std::wstring, std::wstring> rels;
    std::wstring foldedRelsPath = relsPath;
    std::transform(foldedRelsPath.begin(), foldedRelsPath.end(), foldedRelsPath.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    if (archive.find(foldedRelsPath) != archive.end() &&
        !LoadRelationshipsFromMemory(archive, relsPath, rels)) {
        result.error = L"Unable to parse the DOCX document relationships.";
        return result;
    }
    CComPtr<IXMLDOMDocument2> doc;
    if (!LoadXmlFromMemory(archive, documentPath, doc)) { result.error = L"Unable to parse the DOCX document XML."; return result; }
    CComPtr<IXMLDOMElement> root;
    if (FAILED(doc->get_documentElement(&root)) || !root) { result.error = L"The DOCX document XML has no root element."; return result; }
    CComPtr<IXMLDOMDocument2> stylesDoc, numberingDoc;
    CComPtr<IXMLDOMElement> stylesRoot, numberingRoot;
    const std::wstring stylesPath = L"word\\styles.xml";
    const std::wstring numberingPath = L"word\\numbering.xml";
    std::set<std::wstring> colors;
    std::set<std::wstring> fonts;
    CollectColors(root.p, colors);
    CollectFonts(root.p, fonts);
    const std::wstring foldedStylesPath = L"word\\styles.xml";
    if (archive.find(foldedStylesPath) != archive.end()) {
        if (LoadXmlFromMemory(archive, stylesPath, stylesDoc) && SUCCEEDED(stylesDoc->get_documentElement(&stylesRoot)) && stylesRoot) {
            CollectColors(stylesRoot.p, colors);
            CollectFonts(stylesRoot.p, fonts);
        }
        else result.warnings.push_back(L"DOCX styles.xml could not be parsed; direct formatting will be used.");
    }
    const std::wstring foldedNumberingPath = L"word\\numbering.xml";
    if (archive.find(foldedNumberingPath) != archive.end()) {
        if (LoadXmlFromMemory(archive, numberingPath, numberingDoc) && SUCCEEDED(numberingDoc->get_documentElement(&numberingRoot)) && numberingRoot) {
            CollectColors(numberingRoot.p, colors);
            CollectFonts(numberingRoot.p, fonts);
        }
        else result.warnings.push_back(L"DOCX numbering.xml could not be parsed; automatic numbering will be omitted.");
    }
    ImportContext context{ archive, std::move(rels), options, result, colors, fonts };
    if (stylesRoot) context.LoadStyles(stylesRoot.p);
    if (numberingRoot) context.LoadNumbering(numberingRoot.p);
    auto rootChildren = Children(root.p);
    for (const auto& child : rootChildren) {
        if (NodeName(child.p) == L"body") {
            auto blocks = Children(child.p);
            for (const auto& block : blocks) {
                const auto name = NodeName(block.p);
                if (name == L"p") context.Paragraph(block.p);
                else if (name == L"tbl") context.Table(block.p);
                const std::uint64_t maxRtfChars = options.maxDocumentBytes >
                    (std::numeric_limits<std::uint64_t>::max)() / 2
                    ? (std::numeric_limits<std::uint64_t>::max)()
                    : options.maxDocumentBytes * 2;
                if (context.rtf.size() > maxRtfChars) {
                    result.error = L"The converted DOCX RTF exceeds the configured output-size limit.";
                    return result;
                }
            }
        }
    }
    context.rtf += L"}";
    if (context.rtf.size() > static_cast<size_t>(std::numeric_limits<int>::max())) {
        result.error = L"The converted RTF stream is too large."; return result;
    }
    int bytes = WideCharToMultiByte(CP_ACP, 0, context.rtf.data(), (int)context.rtf.size(), nullptr, 0, nullptr, nullptr);
    if (bytes <= 0) { result.error = L"Unable to encode the converted RTF stream."; return result; }
    rtf.resize(bytes);
    WideCharToMultiByte(CP_ACP, 0, context.rtf.data(), (int)context.rtf.size(), reinterpret_cast<char*>(rtf.data()), bytes, nullptr, nullptr);
    result.success = true;
    return result;
}
Result ConvertToRtf(const std::wstring& fileName, std::vector<std::uint8_t>& rtf,
    const Options& options)
{
    Result result;
    std::vector<std::uint8_t> converted;
    std::exception_ptr workerException;
    std::thread worker([&]() {
    ComApartment apartment;
    if (!apartment.usable()) {
    result.error = L"Unable to initialize the DOCX import STA (HRESULT " +
    std::to_wstring(static_cast<unsigned long>(apartment.result())) + L").";
    return;
    }
    try {
    result = ConvertToRtfOnSta(fileName, converted, options);
    }
    catch (...) {
    workerException = std::current_exception();
    }
        });
    worker.join();
    if (workerException) {
        try { std::rethrow_exception(workerException); }
        catch (const std::exception& e) {
            int n = MultiByteToWideChar(CP_UTF8, 0, e.what(), -1, nullptr, 0);
            if (n > 1) {
                std::wstring message(static_cast<size_t>(n), L'\0');
                MultiByteToWideChar(CP_UTF8, 0, e.what(), -1, &message[0], n);
                message.pop_back();
                result.error = L"DOCX import failed: " + message;
            }
            else result.error = L"DOCX import failed due to an unexpected exception.";
        }
        catch (...) {
            result.error = L"DOCX import failed due to an unexpected exception.";
        }
        result.success = false;
        return result;
    }
    if (result.success) rtf.swap(converted);
    return result;
}

} // namespace docxreader