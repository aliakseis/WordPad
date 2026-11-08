#include "stdafx.h"
#include "DocxReader.h"

#include <windows.h>
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

#pragma comment(lib, "msxml6.lib")

namespace docxreader {
namespace {

struct DispatchPtr {
    IDispatch* p = nullptr;
    ~DispatchPtr() { if (p) p->Release(); }
    DispatchPtr() = default;
    DispatchPtr(const DispatchPtr&) = delete;
    DispatchPtr& operator=(const DispatchPtr&) = delete;
    IDispatch** put() { if (p) { p->Release(); p = nullptr; } return &p; }
};
struct XmlDoc {
    IXMLDOMDocument2* p = nullptr;
    ~XmlDoc() { if (p) p->Release(); }
    IXMLDOMDocument2** put() { return &p; }
};
HRESULT Invoke(IDispatch* object, const wchar_t* name, WORD flags,
               VARIANT* args, UINT count, VARIANT* result)
{
    if (!object) return E_POINTER;
    LPOLESTR member = const_cast<LPOLESTR>(name);
    DISPID id = DISPID_UNKNOWN;
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPPARAMS params{};
    params.rgvarg = args;
    params.cArgs = count;
    EXCEPINFO ex{};
    UINT argError = 0;
    return object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params,
                          result, &ex, &argError);
}

bool GetDispatch(IDispatch* object, const wchar_t* name, IDispatch** value)
{
    VARIANT v; VariantInit(&v);
    HRESULT hr = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &v);
    if (FAILED(hr) || v.vt != VT_DISPATCH || !v.pdispVal) {
        VariantClear(&v); VariantInit(&v);
        hr = Invoke(object, name, DISPATCH_METHOD, nullptr, 0, &v);
    }
    if (SUCCEEDED(hr) && v.vt == VT_DISPATCH && v.pdispVal) {
        *value = v.pdispVal; v.pdispVal = nullptr; VariantClear(&v); return true;
    }
    VariantClear(&v); return false;
}

bool GetLong(IDispatch* object, const wchar_t* name, long& value)
{
    VARIANT v; VariantInit(&v);
    HRESULT hr = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &v);
    if (SUCCEEDED(hr)) hr = VariantChangeType(&v, &v, 0, VT_I4);
    if (SUCCEEDED(hr)) value = v.lVal;
    VariantClear(&v); return SUCCEEDED(hr);
}

bool GetBool(IDispatch* object, const wchar_t* name, bool& value)
{
    VARIANT v; VariantInit(&v);
    HRESULT hr = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &v);
    if (SUCCEEDED(hr)) hr = VariantChangeType(&v, &v, 0, VT_BOOL);
    if (SUCCEEDED(hr)) value = v.boolVal != VARIANT_FALSE;
    VariantClear(&v); return SUCCEEDED(hr);
}

bool GetItem(IDispatch* collection, long index, IDispatch** item)
{
    VARIANT arg; VariantInit(&arg); arg.vt = VT_I4; arg.lVal = index;
    VARIANT v; VariantInit(&v);
    HRESULT hr = Invoke(collection, L"Item", DISPATCH_PROPERTYGET, &arg, 1, &v);
    if (FAILED(hr) || v.vt != VT_DISPATCH || !v.pdispVal) {
        VariantClear(&v); VariantInit(&v);
        hr = Invoke(collection, L"Item", DISPATCH_METHOD, &arg, 1, &v);
    }
    if (SUCCEEDED(hr) && v.vt == VT_DISPATCH && v.pdispVal) {
        *item = v.pdispVal; v.pdispVal = nullptr; VariantClear(&v); return true;
    }
    VariantClear(&v); return false;
}

bool CountFiles(IDispatch* folder, unsigned long long& count, unsigned depth = 0)
{
    if (depth > 32) return false;
    DispatchPtr items;
    if (!GetDispatch(folder, L"Items", items.put())) return false;
    long n = 0;
    if (!GetLong(items.p, L"Count", n) || n < 0 || n > 100000) return false;
    for (long i = 0; i < n; ++i) {
        DispatchPtr item;
        if (!GetItem(items.p, i, item.put())) return false;
        bool isFolder = false;
        if (!GetBool(item.p, L"IsFolder", isFolder)) return false;
        if (isFolder) {
            DispatchPtr child;
            if (!GetDispatch(item.p, L"GetFolder", child.put()) ||
                !CountFiles(child.p, count, depth + 1)) return false;
        } else ++count;
    }
    return true;
}

bool GetDiskTreeSize(const std::wstring& dir, unsigned long long& bytes, unsigned depth = 0)
{
    if (depth > 32) return false;
    std::wstring pattern = dir + L"\\*";
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return false;
    do {
        if (!wcscmp(fd.cFileName, L".") || !wcscmp(fd.cFileName, L"..")) continue;
        std::wstring path = dir + L"\\" + fd.cFileName;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            if (!GetDiskTreeSize(path, bytes, depth + 1)) { FindClose(h); return false; }
        } else {
            bytes += (static_cast<unsigned long long>(fd.nFileSizeHigh) << 32) | fd.nFileSizeLow;
            if (bytes > 512ull * 1024 * 1024) { FindClose(h); return false; }
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return true;
}

bool CountDiskFiles(const std::wstring& dir, unsigned long long& count, unsigned depth = 0)
{
    if (depth > 32) return false;
    std::wstring pattern = dir + L"\\*";
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return false;
    do {
        if (!wcscmp(fd.cFileName, L".") || !wcscmp(fd.cFileName, L"..")) continue;
        std::wstring path = dir + L"\\" + fd.cFileName;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            if (!CountDiskFiles(path, count, depth + 1)) { FindClose(h); return false; }
        } else ++count;
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return true;
}

void RemoveTree(const std::wstring& dir)
{
    std::wstring pattern = dir + L"\\*";
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (!wcscmp(fd.cFileName, L".") || !wcscmp(fd.cFileName, L"..")) continue;
            std::wstring path = dir + L"\\" + fd.cFileName;
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) RemoveTree(path);
            else { SetFileAttributesW(path.c_str(), FILE_ATTRIBUTE_NORMAL); DeleteFileW(path.c_str()); }
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    RemoveDirectoryW(dir.c_str());
}

bool ExtractZipWithShell(const std::wstring& zipPath, std::wstring& extracted,
                         std::wstring& content, std::wstring& error)
{
    wchar_t temp[MAX_PATH]{};
    if (!GetTempPathW(MAX_PATH, temp)) { error = L"Unable to locate the temporary directory."; return false; }
    wchar_t unique[MAX_PATH]{};
    if (!GetTempFileNameW(temp, L"wpx", 0, unique)) { error = L"Unable to create a temporary path."; return false; }
    DeleteFileW(unique);
    if (!CreateDirectoryW(unique, nullptr)) { error = L"Unable to create a temporary extraction directory."; return false; }
    extracted = unique;
    WIN32_FILE_ATTRIBUTE_DATA zipInfo{};
    if (!GetFileAttributesExW(zipPath.c_str(), GetFileExInfoStandard, &zipInfo)) {
        error = L"Unable to inspect the DOCX file."; RemoveTree(extracted); return false;
    }
    const unsigned long long archiveBytes =
        (static_cast<unsigned long long>(zipInfo.nFileSizeHigh) << 32) | zipInfo.nFileSizeLow;
    if (archiveBytes == 0 || archiveBytes > 256ull * 1024 * 1024) {
        error = L"The DOCX package is empty or exceeds the 256 MiB import limit.";
        RemoveTree(extracted); return false;
    }
    const std::wstring zipCopy = extracted + L"\\package.zip";
    content = extracted + L"\\content";
    if (!CopyFileW(zipPath.c_str(), zipCopy.c_str(), FALSE) ||
        !CreateDirectoryW(content.c_str(), nullptr)) {
        error = L"Unable to prepare the DOCX extraction workspace.";
        RemoveTree(extracted); return false;
    }

    // Shell.Application exposes Windows' built-in ZIP-folder handler.
    const CLSID clsidShell = {0x13709620, 0xC279, 0x11CE, {0xA4,0x9E,0x44,0x45,0x53,0x54,0x00,0x00}};
    DispatchPtr shell;
    HRESULT hr = CoCreateInstance(clsidShell, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_IDispatch, reinterpret_cast<void**>(shell.put()));
    if (FAILED(hr)) { error = L"Windows ZIP-folder support is unavailable."; RemoveTree(extracted); return false; }

    VARIANT path; VariantInit(&path); path.vt = VT_BSTR; path.bstrVal = SysAllocString(zipCopy.c_str());
    VARIANT zipFolderVar; VariantInit(&zipFolderVar);
    hr = Invoke(shell.p, L"NameSpace", DISPATCH_METHOD, &path, 1, &zipFolderVar);
    VariantClear(&path);
    if (FAILED(hr) || zipFolderVar.vt != VT_DISPATCH || !zipFolderVar.pdispVal) {
        VariantClear(&zipFolderVar); error = L"The DOCX file is not a readable ZIP package."; RemoveTree(extracted); return false;
    }
    DispatchPtr zipFolder; zipFolder.p = zipFolderVar.pdispVal; zipFolderVar.pdispVal = nullptr; VariantClear(&zipFolderVar);

    VARIANT destPath; VariantInit(&destPath); destPath.vt = VT_BSTR; destPath.bstrVal = SysAllocString(content.c_str());
    VARIANT destVar; VariantInit(&destVar);
    hr = Invoke(shell.p, L"NameSpace", DISPATCH_METHOD, &destPath, 1, &destVar);
    VariantClear(&destPath);
    if (FAILED(hr) || destVar.vt != VT_DISPATCH || !destVar.pdispVal) {
        VariantClear(&destVar); error = L"Unable to open the temporary extraction directory."; RemoveTree(extracted); return false;
    }
    DispatchPtr dest; dest.p = destVar.pdispVal; destVar.pdispVal = nullptr; VariantClear(&destVar);

    DispatchPtr items;
    if (!GetDispatch(zipFolder.p, L"Items", items.put())) {
        error = L"Unable to enumerate the DOCX package."; RemoveTree(extracted); return false;
    }
    unsigned long long expected = 0;
    if (!CountFiles(zipFolder.p, expected) || expected == 0 || expected > 100000) {
        error = L"The DOCX package contains an invalid number of entries."; RemoveTree(extracted); return false;
    }

    VARIANT args[2]; VariantInit(&args[0]); VariantInit(&args[1]);
    args[0].vt = VT_I4; args[0].lVal = 0x414; // silent, no-confirmation, no-error-UI
    args[1].vt = VT_DISPATCH; args[1].pdispVal = items.p;
    hr = Invoke(dest.p, L"CopyHere", DISPATCH_METHOD, args, 2, nullptr);
    VariantClear(&args[0]); // args[1] is borrowed
    if (FAILED(hr)) { error = L"Windows could not extract the DOCX package."; RemoveTree(extracted); return false; }

    bool complete = false;
    for (int attempt = 0; attempt < 300; ++attempt) {
        unsigned long long actual = 0, extractedBytes = 0;
        if (!GetDiskTreeSize(content, extractedBytes)) {
            error = L"The expanded DOCX package exceeds the 512 MiB extraction limit.";
            RemoveTree(extracted); return false;
        }
        if (CountDiskFiles(content, actual) && actual >= expected) { complete = true; break; }
        Sleep(100);
    }
    if (!complete) {
        error = L"Timed out while extracting the DOCX package."; RemoveTree(extracted); return false;
    }
    return true;
}

std::wstring ReadTextFile(const std::wstring& path)
{
    std::ifstream f(path, std::ios::binary | std::ios::ate);
    if (!f) return {};
    const std::streamoff fileSize = static_cast<std::streamoff>(f.tellg());
    if (fileSize <= 0 || fileSize > static_cast<std::streamoff>(128ull * 1024 * 1024)) return {};
    std::string bytes(static_cast<size_t>(fileSize), '\0');
    f.seekg(0);
    f.read(&bytes[0], static_cast<std::streamsize>(bytes.size()));
    if (!f) return {};
    if (bytes.empty()) return {};
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
    if (bytes.size() >= 3 && (unsigned char)bytes[0] == 0xEF &&
        (unsigned char)bytes[1] == 0xBB && (unsigned char)bytes[2] == 0xBF) bytes.erase(0, 3);
    int n = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), (int)bytes.size(), nullptr, 0);
    if (!n) return {};
    std::wstring result(n, L'\0');
    MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), (int)bytes.size(), &result[0], n);
    return result;
}

bool LoadXml(const std::wstring& path, XmlDoc& doc)
{
    HRESULT hr = CoCreateInstance(CLSID_DOMDocument60, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_IXMLDOMDocument2, reinterpret_cast<void**>(doc.put()));
    if (FAILED(hr) || !doc.p) return false;
    doc.p->put_async(VARIANT_FALSE);
    doc.p->put_validateOnParse(VARIANT_FALSE);
    doc.p->put_resolveExternals(VARIANT_FALSE);
    doc.p->put_preserveWhiteSpace(VARIANT_TRUE);
    BSTR propertyName = SysAllocString(L"ProhibitDTD");
    VARIANT propertyValue; VariantInit(&propertyValue);
    propertyValue.vt = VT_BOOL; propertyValue.boolVal = VARIANT_TRUE;
    if (propertyName) doc.p->setProperty(propertyName, propertyValue);
    SysFreeString(propertyName);
    VariantClear(&propertyValue);
    std::wstring xml = ReadTextFile(path);
    if (xml.empty()) return false;
    BSTR xmlText = SysAllocStringLen(xml.data(), static_cast<UINT>(xml.size()));
    if (!xmlText) return false;
    VARIANT_BOOL loaded = VARIANT_FALSE;
    HRESULT loadResult = doc.p->loadXML(xmlText, &loaded);
    SysFreeString(xmlText);
    return SUCCEEDED(loadResult) && loaded == VARIANT_TRUE;
}

std::wstring NodeName(IXMLDOMNode* node)
{
    BSTR s = nullptr;
    if (FAILED(node->get_nodeName(&s)) || !s) return {};
    std::wstring result(s, SysStringLen(s)); SysFreeString(s);
    const auto colon = result.find_last_of(L':');
    if (colon != std::wstring::npos) result.erase(0, colon + 1);
    return result;
}
std::wstring NodeText(IXMLDOMNode* node)
{
    BSTR s = nullptr;
    if (FAILED(node->get_text(&s)) || !s) return {};
    std::wstring result(s, SysStringLen(s)); SysFreeString(s); return result;
}
std::wstring Attr(IXMLDOMNode* node, const wchar_t* name)
{
    IXMLDOMNamedNodeMap* attrs = nullptr;
    if (FAILED(node->get_attributes(&attrs)) || !attrs) return {};
    long count = 0; attrs->get_length(&count);
    std::wstring value;
    for (long i = 0; i < count; ++i) {
        IXMLDOMNode* attr = nullptr;
        if (SUCCEEDED(attrs->get_item(i, &attr)) && attr) {
            if (NodeName(attr) == name) value = NodeText(attr);
            attr->Release();
            if (!value.empty()) break;
        }
    }
    attrs->Release(); return value;
}
std::vector<IXMLDOMNode*> Children(IXMLDOMNode* node)
{
    std::vector<IXMLDOMNode*> out;
    IXMLDOMNodeList* list = nullptr;
    if (FAILED(node->get_childNodes(&list)) || !list) return out;
    long count = 0; list->get_length(&count);
    for (long i = 0; i < count; ++i) {
        IXMLDOMNode* child = nullptr;
        if (SUCCEEDED(list->get_item(i, &child)) && child) out.push_back(child);
    }
    list->Release(); return out;
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
    for (const auto& p : parts) { if (!out.empty()) out += L"\\"; out += p; }
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
            out += L"\\u" + std::to_wstring(value) + L"?";
        }
    }
    return out;
}

bool ReadBinary(const std::wstring& path, std::uint64_t maxBytes, std::vector<std::uint8_t>& data)
{
    std::ifstream f(path, std::ios::binary | std::ios::ate);
    if (!f) return false;
    auto n = f.tellg();
    if (n < 0 || static_cast<std::uint64_t>(n) > maxBytes) return false;
    data.resize(static_cast<size_t>(n)); f.seekg(0);
    if (!data.empty()) f.read(reinterpret_cast<char*>(data.data()), data.size());
    return f.good() || f.eof();
}

std::wstring Hex(const std::vector<std::uint8_t>& bytes)
{
    static const wchar_t digits[] = L"0123456789abcdef";
    std::wstring out; out.reserve(bytes.size() * 2);
    for (auto b : bytes) { out += digits[b >> 4]; out += digits[b & 15]; }
    return out;
}

struct ImportContext {
    std::wstring root;
    std::map<std::wstring, std::wstring> rels;
    const Options& options;
    Result& result;
    std::wstring rtf = L"{\\rtf1\\ansi\\deff0\\uc1\n";

    ImportContext(const std::wstring& packageRoot,
                  std::map<std::wstring, std::wstring>&& relationships,
                  const Options& importOptions, Result& importResult)
        : root(packageRoot), rels(std::move(relationships)),
          options(importOptions), result(importResult) {}

    bool Image(IXMLDOMNode* drawing)
    {
        if (!options.importImages) return false;
        std::wstring rid, ext;
        long widthTwips = 3000, heightTwips = 2000;
        std::vector<IXMLDOMNode*> stack = Children(drawing);
        while (!stack.empty()) {
            IXMLDOMNode* n = stack.back(); stack.pop_back();
            const auto nodeName = NodeName(n);
            if (nodeName == L"blip") rid = Attr(n, L"embed");
            else if (nodeName == L"extent") {
                const auto cx = Attr(n, L"cx"), cy = Attr(n, L"cy");
                if (!cx.empty()) widthTwips = std::max(1L, std::min(20000L, _wtol(cx.c_str()) / 635));
                if (!cy.empty()) heightTwips = std::max(1L, std::min(20000L, _wtol(cy.c_str()) / 635));
            }
            auto children = Children(n);
            for (auto* c : children) stack.push_back(c);
            // Nodes remain owned by stack until processed; release after visiting.
            n->Release();
        }
        if (rid.empty()) return false;
        auto it = rels.find(rid);
        if (it == rels.end()) return false;
        std::wstring path = root + L"\\" + it->second;
        std::vector<std::uint8_t> bytes;
        if (!ReadBinary(path, options.maxImageBytes, bytes) || bytes.empty()) return false;
        std::wstring lower = it->second;
        std::transform(lower.begin(), lower.end(), lower.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
        if (lower.size() >= 4 && lower.substr(lower.size()-4) == L".png") ext = L"\\pngblip";
        else if ((lower.size() >= 4 && lower.substr(lower.size()-4) == L".jpg") ||
                 (lower.size() >= 5 && lower.substr(lower.size()-5) == L".jpeg")) ext = L"\\jpegblip";
        else return false;
        rtf += L"{\\pict" + ext + L"\\picw" + std::to_wstring(std::max(1L, widthTwips / 15)) +
               L"\\pich" + std::to_wstring(std::max(1L, heightTwips / 15)) +
               L"\\picwgoal" + std::to_wstring(widthTwips) +
               L"\\pichgoal" + std::to_wstring(heightTwips) + L"\n" + Hex(bytes) + L"}";
        ++result.imagesRead; return true;
    }

    void Run(IXMLDOMNode* run)
    {
        rtf += L"{";
        IXMLDOMNode* props = nullptr;
        auto children = Children(run);
        for (auto* child : children) {
            const auto name = NodeName(child);
            if (name == L"rPr") props = child;
        }
        if (props) {
            auto p = Children(props);
            for (auto* x : p) {
                auto n = NodeName(x);
                if (n == L"b" && Attr(x, L"val") != L"0" && Attr(x, L"val") != L"false") rtf += L"\\b";
                else if (n == L"i" && Attr(x, L"val") != L"0" && Attr(x, L"val") != L"false") rtf += L"\\i";
                else if ((n == L"strike" || n == L"dstrike") && Attr(x, L"val") != L"0") rtf += L"\\strike";
                else if (n == L"u" && Attr(x, L"val") != L"none" && Attr(x, L"val") != L"0") rtf += L"\\ul";
                else if (n == L"sz") {
                    auto v = Attr(x, L"val"); if (!v.empty()) rtf += L"\\fs" + v;
                }
                x->Release();
            }
        }
        for (auto* child : children) {
            auto name = NodeName(child);
            if (name == L"t") rtf += RtfText(NodeText(child));
            else if (name == L"tab") rtf += L"\\tab ";
            else if (name == L"br") rtf += (Attr(child, L"type") == L"page" ? L"\\page " : L"\\line ");
            else if (name == L"drawing" || name == L"pict") Image(child);
            child->Release();
        }
        rtf += L"\\plain}";
    }

    void Paragraph(IXMLDOMNode* paragraph, bool inTable = false)
    {
        rtf += inTable ? L"\\pard\\intbl" : L"\\pard";
        auto children = Children(paragraph);
        IXMLDOMNode* pPr = nullptr;
        for (auto* c : children) if (NodeName(c) == L"pPr") pPr = c;
        if (pPr) {
            auto props = Children(pPr);
            for (auto* p : props) {
                auto n = NodeName(p);
                if (n == L"jc") {
                    auto v = Attr(p, L"val");
                    if (v == L"center") rtf += L"\\qc";
                    else if (v == L"right" || v == L"end") rtf += L"\\qr";
                    else if (v == L"both" || v == L"distribute") rtf += L"\\qj";
                    else rtf += L"\\ql";
                }
                p->Release();
            }
        }
        for (auto* c : children) {
            auto n = NodeName(c);
            if (n == L"r") Run(c);
            else if (n == L"hyperlink") Hyperlink(c);
            else if (n == L"bookmarkStart" || n == L"bookmarkEnd" || n == L"pPr") {}
            c->Release();
        }
        rtf += L"\\par\n";
    }

    void Hyperlink(IXMLDOMNode* link)
    {
        if (!options.importHyperlinks) {
            auto ch = Children(link); for (auto* c : ch) { if (NodeName(c) == L"r") Run(c); c->Release(); }
            return;
        }
        std::wstring rid = Attr(link, L"id"), target;
        auto it = rels.find(rid); if (it != rels.end()) target = it->second;
        if (target.empty() || target.find(L":") == std::wstring::npos) {
            auto ch = Children(link); for (auto* c : ch) { if (NodeName(c) == L"r") Run(c); c->Release(); }
            return;
        }
        std::wstring fieldTarget = RtfText(target);
        size_t quote = 0;
        while ((quote = fieldTarget.find(L'"', quote)) != std::wstring::npos) {
            fieldTarget.replace(quote, 1, L"\\'22"); quote += 4;
        }
        rtf += L"{\\field{\\*\\fldinst HYPERLINK \"" + fieldTarget + L"\"}{\\fldrslt ";
        auto ch = Children(link); for (auto* c : ch) { if (NodeName(c) == L"r") Run(c); c->Release(); }
        rtf += L"}}"; ++result.hyperlinksRead;
    }

    void Table(IXMLDOMNode* table)
    {
        auto rows = Children(table);
        for (auto* row : rows) {
            if (NodeName(row) != L"tr") { row->Release(); continue; }
            rtf += L"\\trowd\\trgaph108\\trleft0";
            auto cells = Children(row); int x = 1800;
            for (auto* cell : cells) if (NodeName(cell) == L"tc") { x += 1800; rtf += L"\\cellx" + std::to_wstring(x); }
            rtf += L"\n";
            for (auto* cell : cells) {
                if (NodeName(cell) == L"tc") {
                    rtf += L"\\intbl ";
                    auto blocks = Children(cell);
                    for (auto* b : blocks) {
                        if (NodeName(b) == L"p") Paragraph(b, true);
                        else if (NodeName(b) == L"tbl") Table(b);
                        b->Release();
                    }
                    rtf += L"\\cell ";
                }
                cell->Release();
            }
            rtf += L"\\row\n";
            row->Release();
        }
    }
};

bool LoadRelationships(const std::wstring& path, std::map<std::wstring,std::wstring>& rels)
{
    XmlDoc doc; if (!LoadXml(path, doc)) return false;
    IXMLDOMElement* root = nullptr; if (FAILED(doc.p->get_documentElement(&root)) || !root) return false;
    auto children = Children(root); root->Release();
    for (auto* n : children) {
        if (NodeName(n) == L"Relationship") {
            std::wstring id = Attr(n, L"Id"), target = Attr(n, L"Target");
            std::wstring mode = Attr(n, L"TargetMode");
            if (!id.empty() && !target.empty() && mode != L"External") rels[id] = NormalizeTarget(target);
            else if (!id.empty() && !target.empty() && mode == L"External") rels[id] = target;
        }
        n->Release();
    }
    return true;
}

} // namespace

Result ConvertToRtf(const std::wstring& fileName, std::vector<std::uint8_t>& rtf,
                    const Options& options)
{
    Result result;
    std::wstring extracted, content;
    if (!ExtractZipWithShell(fileName, extracted, content, result.error)) return result;
    struct Cleanup { std::wstring path; ~Cleanup() { if (!path.empty()) RemoveTree(path); } } cleanup{extracted};

    const std::wstring documentPath = content + L"\\word\\document.xml";
    const std::wstring relsPath = content + L"\\word\\_rels\\document.xml.rels";
    std::vector<std::uint8_t> documentBytes;
    if (!ReadBinary(documentPath, options.maxDocumentBytes, documentBytes) || documentBytes.empty()) {
        result.error = L"The DOCX package does not contain a readable word/document.xml."; return result;
    }
    std::map<std::wstring,std::wstring> rels;
    if (GetFileAttributesW(relsPath.c_str()) != INVALID_FILE_ATTRIBUTES)
        LoadRelationships(relsPath, rels);

    XmlDoc doc;
    if (!LoadXml(documentPath, doc)) { result.error = L"Unable to parse the DOCX document XML."; return result; }
    IXMLDOMElement* root = nullptr;
    if (FAILED(doc.p->get_documentElement(&root)) || !root) { result.error = L"The DOCX document XML has no root element."; return result; }

    ImportContext context{content, std::move(rels), options, result};
    auto rootChildren = Children(root); root->Release();
    for (auto* child : rootChildren) {
        if (NodeName(child) == L"body") {
            auto blocks = Children(child);
            for (auto* block : blocks) {
                auto name = NodeName(block);
                if (name == L"p") context.Paragraph(block);
                else if (name == L"tbl") context.Table(block);
                block->Release();
            }
        }
        child->Release();
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

} // namespace docxreader
