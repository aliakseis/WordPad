#include "stdafx.h"
#include "DocxReader.h"

#include <windows.h>
#include <atlbase.h>
#include <msxml6.h>
#include <algorithm>
#include <cwctype>
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
        bool ExtractZipPackageToMemory(const std::wstring& zipPath, MemoryArchive& virtualFiles, std::wstring& error)
        {
            std::vector<std::uint8_t> archive;
            if (!ReadWholeFile(zipPath, archive, 256ull * 1024 * 1024) || archive.size() < 22) {
                error = L"Unable to read the DOCX ZIP package or the file exceeds the 256 MiB limit.";
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
                    if (totalExpanded > 512ull * 1024 * 1024) {
                        error = L"The expanded DOCX package exceeds the 512 MiB limit.";
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

                    wchar_t nextChar = (i + 1 < text.size()) ? text[i + 1] : L'\0';

                    if (nextChar >= L'0' && nextChar <= L'9') {
                        out += L"\\u" + std::to_wstring(value) + nextChar;

                        ++i;
                    }
                    else {
                        out += L"\\u" + std::to_wstring(value) + L"?";
                    }
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
        struct ImportContext {
            const MemoryArchive& archive;
            std::map<std::wstring, std::wstring> rels;
            const Options& options;
            Result& result;
            std::wstring rtf = L"{\\rtf1\\ansi\\deff0\\uc1\n";
            ImportContext(const MemoryArchive& memArchive,
                std::map<std::wstring, std::wstring>&& relationships,
                const Options& importOptions, Result& importResult)
                : archive(memArchive), rels(std::move(relationships)),
                options(importOptions), result(importResult) {
            }
            bool Image(IXMLDOMNode* drawing)
            {
                if (!options.importImages) return false;
                std::wstring rid, ext;
                long widthTwips = 3000, heightTwips = 2000;
                auto stack = Children(drawing);
                while (!stack.empty()) {
                    auto node = stack.back();
                    stack.pop_back();
                    const auto nodeName = NodeName(node.p);
                    if (nodeName == L"blip") rid = Attr(node.p, L"embed");
                    else if (nodeName == L"extent") {
                        const auto cx = Attr(node.p, L"cx"), cy = Attr(node.p, L"cy");
                        if (!cx.empty()) widthTwips = std::max(1L, std::min(20000L, _wtol(cx.c_str()) / 635));
                        if (!cy.empty()) heightTwips = std::max(1L, std::min(20000L, _wtol(cy.c_str()) / 635));
                    }
                    auto children = Children(node.p);
                    for (const auto& child : children) stack.push_back(child);
                }
                if (rid.empty()) return false;
                auto it = rels.find(rid);
                if (it == rels.end()) return false;
                std::wstring path = it->second;
                std::vector<std::uint8_t> bytes;
                if (!ReadBinaryFromMemory(archive, path, options.maxImageBytes, bytes) || bytes.empty()) return false;
                std::wstring lower = it->second;
                std::transform(lower.begin(), lower.end(), lower.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
                if (lower.size() >= 4 && lower.substr(lower.size() - 4) == L".png") ext = L"\\pngblip";
                else if ((lower.size() >= 4 && lower.substr(lower.size() - 4) == L".jpg") ||
                    (lower.size() >= 5 && lower.substr(lower.size() - 5) == L".jpeg")) ext = L"\\jpegblip";
                else return false;
                rtf += L"{\\pict" + ext + L"\\picw" + std::to_wstring(std::max(1L, widthTwips / 15)) +
                    L"\\pich" + std::to_wstring(std::max(1L, heightTwips / 15)) +
                    L"\\picwgoal" + std::to_wstring(widthTwips) +
                    L"\\pichgoal" + std::to_wstring(heightTwips) + L"\n" + Hex(bytes) + L"}";
                ++result.imagesRead;
                return true;
            }
            void Run(IXMLDOMNode* run)
            {
                rtf += L"{";
                CComPtr<IXMLDOMNode> properties;
                auto children = Children(run);
                for (const auto& child : children) {
                    if (NodeName(child.p) == L"rPr") {
                        properties = child.p;
                        break;
                    }
                }
                if (properties) {
                    auto propertyNodes = Children(properties.p);
                    for (const auto& property : propertyNodes) {
                        const auto name = NodeName(property.p);
                        if (name == L"b" && Attr(property.p, L"val") != L"0" && Attr(property.p, L"val") != L"false") rtf += L"\\b";
                        else if (name == L"i" && Attr(property.p, L"val") != L"0" && Attr(property.p, L"val") != L"false") rtf += L"\\i";
                        else if ((name == L"strike" || name == L"dstrike") && Attr(property.p, L"val") != L"0") rtf += L"\\strike";
                        else if (name == L"u" && Attr(property.p, L"val") != L"none" && Attr(property.p, L"val") != L"0") rtf += L"\\ul";
                        else if (name == L"sz") {
                            auto value = Attr(property.p, L"val");
                            if (!value.empty()) {
                                rtf += L"\\fs" + value + L" ";
                            }
                        }
                    }
                }
                for (const auto& child : children) {
                    const auto name = NodeName(child.p);
                    if (name == L"t") rtf += RtfText(NodeText(child.p));
                    else if (name == L"tab") rtf += L"\\tab ";
                    else if (name == L"br") rtf += (Attr(child.p, L"type") == L"page" ? L"\\page " : L"\\line ");
                    else if (name == L"drawing" || name == L"pict") Image(child.p);
                }
                rtf += L"\\plain}";
            }
            void Paragraph(IXMLDOMNode* paragraph, bool inTable = false)
            {
                rtf += inTable ? L"\\pard\\intbl" : L"\\pard";
                auto children = Children(paragraph);
                CComPtr<IXMLDOMNode> paragraphProperties;
                for (const auto& child : children) {
                    if (NodeName(child.p) == L"pPr") {
                        paragraphProperties = child.p;
                        break;
                    }
                }
                if (paragraphProperties) {
                    auto properties = Children(paragraphProperties.p);
                    for (const auto& property : properties) {
                        if (NodeName(property.p) == L"jc") {
                            const auto value = Attr(property.p, L"val");
                            if (value == L"center") rtf += L"\\qc";
                            else if (value == L"right" || value == L"end") rtf += L"\\qr";
                            else if (value == L"both" || value == L"distribute") rtf += L"\\qj";
                            else rtf += L"\\ql";
                        }
                    }
                }
                for (const auto& child : children) {
                    const auto name = NodeName(child.p);
                    if (name == L"r") Run(child.p);
                    else if (name == L"hyperlink") Hyperlink(child.p);
                }
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
                if (target.empty() || target.find(L":") == std::wstring::npos) {
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
            void Table(IXMLDOMNode * table)
            {
                auto rows = Children(table);
                for (const auto& row : rows) {
                    if (NodeName(row.p) != L"tr") continue;
                    rtf += L"\\trowd\\trgaph108\\trleft0";
                    auto cells = Children(row.p);
                    int x = 1800;
                    for (const auto& cell : cells) {
                        if (NodeName(cell.p) == L"tc") {
                            x += 1800;
                            rtf += L"\\cellx" + std::to_wstring(x);
                        }
                    }
                    rtf += L"\n";
                    for (const auto& cell : cells) {
                        if (NodeName(cell.p) != L"tc") continue;
                        rtf += L"\\intbl ";
                        auto blocks = Children(cell.p);
                        for (const auto& block : blocks) {
                            const auto name = NodeName(block.p);
                            if (name == L"p") Paragraph(block.p, true);
                            else if (name == L"tbl") Table(block.p);
                        }
                        rtf += L"\\cell ";
                    }
                    rtf += L"\\row\n";
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
    // Unpack ZIP directly to the memory map object instead of hard disk directories
    if (!ExtractZipPackageToMemory(fileName, archive, result.error)) return result;
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
    if (archive.find(foldedRelsPath) != archive.end()) {
        LoadRelationshipsFromMemory(archive, relsPath, rels);
    }
    CComPtr<IXMLDOMDocument2> doc;
    if (!LoadXmlFromMemory(archive, documentPath, doc)) { result.error = L"Unable to parse the DOCX document XML."; return result; }
    CComPtr<IXMLDOMElement> root;
    if (FAILED(doc->get_documentElement(&root)) || !root) { result.error = L"The DOCX document XML has no root element."; return result; }
    ImportContext context{ archive, std::move(rels), options, result };
    auto rootChildren = Children(root.p);
    for (const auto& child : rootChildren) {
        if (NodeName(child.p) == L"body") {
            auto blocks = Children(child.p);
            for (const auto& block : blocks) {
                const auto name = NodeName(block.p);
                if (name == L"p") context.Paragraph(block.p);
                else if (name == L"tbl") context.Table(block.p);
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