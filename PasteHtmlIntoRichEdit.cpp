#include "stdafx.h"

#include "PasteHtmlIntoRichEdit.h"

#include <tom.h>

#include <cstdlib>

#include <string>
#include <vector>
#include <algorithm>
#include <cstdint>
#include <climits>



#include <windows.h>
#include <objidl.h>

#include <cstring>


#include <winhttp.h>
#include <wincodec.h>
#include <wrl/client.h>

#include <limits>
#include <map>
#include <atomic>
#include <thread>
#include <chrono>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "ole32.lib")


#ifdef _DEBUG
#undef THIS_FILE
static char BASED_CODE THIS_FILE[] = __FILE__;
#endif


namespace
{
    using Clock = std::chrono::steady_clock;
    using HtmlImageCache = std::map<std::wstring, std::vector<BYTE>>;

    constexpr size_t MaxRemoteImages = 16;
    constexpr size_t MaxCachedImageBytes = 32u * 1024u * 1024u;
    constexpr size_t MaxParallelDownloads = 4;

    bool DownloadHttpsImage(
        const std::wstring& url,
        std::vector<BYTE>& data,
        Clock::time_point deadline)
    {
        constexpr DWORD MaxImageBytes = 16u * 1024u * 1024u;

        data.clear();

        URL_COMPONENTS parts = {};
        parts.dwStructSize = sizeof(parts);
        parts.dwSchemeLength = DWORD(-1);
        parts.dwHostNameLength = DWORD(-1);
        parts.dwUrlPathLength = DWORD(-1);
        parts.dwExtraInfoLength = DWORD(-1);

        if (Clock::now() >= deadline ||
            !WinHttpCrackUrl(url.c_str(), 0, 0, &parts))
            return false;

        if (parts.nScheme != INTERNET_SCHEME_HTTPS ||
            parts.dwHostNameLength == 0)
        {
            return false;
        }

        const std::wstring host(
            parts.lpszHostName, parts.dwHostNameLength);

        std::wstring path(parts.lpszUrlPath, parts.dwUrlPathLength);

        if (parts.dwExtraInfoLength)
            path.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);

        if (path.empty())
            path = L"/";

        HINTERNET session = WinHttpOpen(
            L"RichEditor/1.0",
            WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
            WINHTTP_NO_PROXY_NAME,
            WINHTTP_NO_PROXY_BYPASS,
            0);

        if (!session)
            return false;

        WinHttpSetTimeouts(session, 2000, 2000, 2500, 2500);

        // Do not follow redirects automatically: a redirect could point
        // to an HTTP URL or an unintended internal destination.
        HINTERNET connection = WinHttpConnect(
            session, host.c_str(), parts.nPort, 0);

        HINTERNET request = connection
            ? WinHttpOpenRequest(
                connection,
                L"GET",
                path.c_str(),
                nullptr,
                WINHTTP_NO_REFERER,
                WINHTTP_DEFAULT_ACCEPT_TYPES,
                WINHTTP_FLAG_SECURE)
            : nullptr;

        // Never follow redirects; this also prevents redirecting to HTTP.
        if (request)
        {
            DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_NEVER;
            WinHttpSetOption(
                request,
                WINHTTP_OPTION_REDIRECT_POLICY,
                &redirectPolicy,
                sizeof(redirectPolicy));
        }

        bool ok = false;

        if (request && Clock::now() < deadline &&
            WinHttpSendRequest(
                request,
                WINHTTP_NO_ADDITIONAL_HEADERS,
                0,
                WINHTTP_NO_REQUEST_DATA,
                0,
                0,
                0) &&
            WinHttpReceiveResponse(request, nullptr))
        {
            DWORD status = 0;
            DWORD statusSize = sizeof(status);

            if (WinHttpQueryHeaders(
                request,
                WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                WINHTTP_HEADER_NAME_BY_INDEX,
                &status,
                &statusSize,
                WINHTTP_NO_HEADER_INDEX) &&
                status == HTTP_STATUS_OK)
            {
                DWORD contentLength = 0;
                DWORD lengthSize = sizeof(contentLength);

                if (WinHttpQueryHeaders(
                    request,
                    WINHTTP_QUERY_CONTENT_LENGTH |
                    WINHTTP_QUERY_FLAG_NUMBER,
                    WINHTTP_HEADER_NAME_BY_INDEX,
                    &contentLength,
                    &lengthSize,
                    WINHTTP_NO_HEADER_INDEX) &&
                    contentLength > MaxImageBytes)
                {
                    // Reject oversized responses before downloading them.
                }
                else
                {
                    ok = true;

                    for (;;)
                    {
                        if (Clock::now() >= deadline)
                        {
                            ok = false;
                            break;
                        }

                        DWORD available = 0;

                        if (!WinHttpQueryDataAvailable(request, &available))
                        {
                            ok = false;
                            break;
                        }

                        if (available == 0)
                            break;

                        if (available > MaxImageBytes - data.size())
                        {
                            ok = false;
                            break;
                        }

                        const size_t oldSize = data.size();
                        data.resize(oldSize + available);

                        DWORD received = 0;

                        if (!WinHttpReadData(
                            request,
                            data.data() + oldSize,
                            available,
                            &received) ||
                            received == 0)
                        {
                            data.resize(oldSize);
                            ok = false;
                            break;
                        }

                        data.resize(oldSize + received);
                    }

                    if (data.empty())
                        ok = false;
                }
            }
        }

        if (request)
            WinHttpCloseHandle(request);

        if (connection)
            WinHttpCloseHandle(connection);

        WinHttpCloseHandle(session);

        if (!ok)
            data.clear();

        return ok;
    }

    bool ReadHtmlClipboard(
        LPDATAOBJECT dataObject,
        CLIPFORMAT cfHTML,
        std::wstring& html)
    {
        html.clear();

        if (!dataObject || cfHTML == 0)
            return false;

        FORMATETC format = {};
        format.cfFormat = cfHTML;
        format.dwAspect = DVASPECT_CONTENT;
        format.lindex = -1;
        format.tymed = TYMED_HGLOBAL;

        STGMEDIUM medium = {};

        const HRESULT hr = dataObject->GetData(&format, &medium);
        if (FAILED(hr))
            return false;

        struct ReleaseMedium
        {
            STGMEDIUM* medium;

            ~ReleaseMedium()
            {
                ReleaseStgMedium(medium);
            }
        } release{ &medium };

        if (medium.tymed != TYMED_HGLOBAL || !medium.hGlobal)
            return false;

        const SIZE_T globalSize = GlobalSize(medium.hGlobal);
        if (globalSize == 0)
            return false;

        const auto* locked = static_cast<const char*>(
            GlobalLock(medium.hGlobal));

        if (!locked)
            return false;

        struct UnlockGlobal
        {
            HGLOBAL handle;

            ~UnlockGlobal()
            {
                GlobalUnlock(handle);
            }
        } unlock{ medium.hGlobal };

        // Copy the payload while its HGLOBAL is locked. Do not assume
        // the clipboard data is null-terminated.
        std::vector<char> bytes(locked, locked + globalSize);

        // CF_HTML is an ASCII header followed by UTF-8 HTML.
        // Locate the end of the header before parsing its offset fields.
        const size_t firstHtmlByte = [&bytes]() -> size_t
            {
                const auto it = std::find(bytes.begin(), bytes.end(), '<');
                return static_cast<size_t>(it - bytes.begin());
            }();

        const size_t headerSize = firstHtmlByte;

        auto parseOffset = [&bytes, headerSize](
            const char* key,
            size_t& value) -> bool
            {
                const size_t keyLength = std::strlen(key);

                if (headerSize < keyLength)
                    return false;

                for (size_t i = 0; i + keyLength <= headerSize; ++i)
                {
                    if (_strnicmp(bytes.data() + i, key, keyLength) != 0)
                        continue;

                    size_t p = i + keyLength;

                    while (p < headerSize &&
                        (bytes[p] == ' ' || bytes[p] == '\t'))
                    {
                        ++p;
                    }

                    if (p >= headerSize || bytes[p] != ':')
                        continue;

                    ++p;

                    while (p < headerSize &&
                        (bytes[p] == ' ' || bytes[p] == '\t'))
                    {
                        ++p;
                    }

                    if (p >= headerSize || bytes[p] < '0' || bytes[p] > '9')
                        return false;

                    size_t result = 0;

                    while (p < headerSize &&
                        bytes[p] >= '0' && bytes[p] <= '9')
                    {
                        const size_t digit =
                            static_cast<size_t>(bytes[p] - '0');

                        if (result > (SIZE_MAX - digit) / 10)
                            return false;

                        result = result * 10 + digit;
                        ++p;
                    }

                    value = result;
                    return true;
                }

                return false;
            };

        size_t start = 0;
        size_t end = 0;

        bool validOffsets =
            parseOffset("StartFragment", start) &&
            parseOffset("EndFragment", end) &&
            start < end &&
            end <= bytes.size();

        // Some clipboard producers supply placeholder offsets such as -1.
        // Fall back to the standard HTML fragment comments in that case.
        if (!validOffsets)
        {
            static constexpr char startMarker[] = "<!--StartFragment-->";
            static constexpr char endMarker[] = "<!--EndFragment-->";

            auto findMarker = [&bytes](
                const char* marker,
                size_t markerLength,
                size_t from) -> size_t
                {
                    if (from > bytes.size() ||
                        markerLength > bytes.size() - from)
                    {
                        return std::string::npos;
                    }

                    for (size_t i = from;
                        i + markerLength <= bytes.size();
                        ++i)
                    {
                        if (std::memcmp(
                            bytes.data() + i, marker, markerLength) == 0)
                        {
                            return i;
                        }
                    }

                    return std::string::npos;
                };

            const size_t markerStart = findMarker(
                startMarker, sizeof(startMarker) - 1, 0);

            if (markerStart == std::string::npos)
                return false;

            start = markerStart + sizeof(startMarker) - 1;

            const size_t markerEnd = findMarker(
                endMarker, sizeof(endMarker) - 1, start);

            if (markerEnd == std::string::npos || markerEnd <= start)
                return false;

            end = markerEnd;
        }

        const size_t length = end - start;

        if (length == 0 ||
            length > static_cast<size_t>(INT_MAX))
        {
            return false;
        }

        // CF_HTML offsets are byte offsets. Decode UTF-8 only after
        // selecting the fragment.
        const char* utf8 = bytes.data() + start;

        int wideLength = MultiByteToWideChar(
            CP_UTF8,
            MB_ERR_INVALID_CHARS,
            utf8,
            static_cast<int>(length),
            nullptr,
            0);

        if (wideLength == 0)
        {
            // Some applications provide malformed UTF-8. Retry with
            // replacement characters rather than rejecting all content.
            wideLength = MultiByteToWideChar(
                CP_UTF8,
                0,
                utf8,
                static_cast<int>(length),
                nullptr,
                0);
        }

        if (wideLength <= 0)
            return false;

        std::wstring result(static_cast<size_t>(wideLength), L'\0');

        if (MultiByteToWideChar(
            CP_UTF8,
            0,
            utf8,
            static_cast<int>(length),
            &result[0],
            wideLength) != wideLength)
        {
            html.clear();
            return false;
        }

        if (!result.empty() && result.front() == L'\uFEFF')
            result.erase(result.begin());

        html = std::move(result);
        return !html.empty();
    }


    void AppendRtfText(
        std::string& rtf,
        const std::wstring& text)
    {
        for (wchar_t ch : text)
        {
            switch (ch)
            {
            case L'\\': rtf += "\\\\"; break;
            case L'{':  rtf += "\\{";  break;
            case L'}':  rtf += "\\}";  break;
            case L'\r': break;
            case L'\n': rtf += "\\line "; break;
            case L'\t': rtf += "\\tab "; break;
            default:
                if (ch >= 0x20 && ch < 0x7f)
                {
                    rtf.push_back(static_cast<char>(ch));
                }
                else
                {
                    // RTF Unicode escapes use signed 16-bit values.
                    rtf += "\\u";
                    rtf += std::to_string(
                        static_cast<short>(ch));
                    rtf += "?";
                }
                break;
            }
        }
    }



    void AppendHex(
        std::string& out,
        const BYTE* bytes,
        size_t size)
    {
        static constexpr char hex[] = "0123456789abcdef";

        for (size_t i = 0; i < size; ++i)
        {
            out.push_back(hex[bytes[i] >> 4]);
            out.push_back(hex[bytes[i] & 15]);
        }
    }

    bool AppendBitmapRtf(
        std::string& rtf,
        const std::vector<BYTE>& encoded)
    {
        using Microsoft::WRL::ComPtr;

        if (encoded.empty() ||
            encoded.size() > (std::numeric_limits<UINT>::max)())
        {
            return false;
        }

        ComPtr<IWICImagingFactory> factory;

        HRESULT hr = CoCreateInstance(
            CLSID_WICImagingFactory,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&factory));

        if (FAILED(hr))
            return false;

        ComPtr<IWICStream> stream;

        hr = factory->CreateStream(&stream);
        if (FAILED(hr))
            return false;

        hr = stream->InitializeFromMemory(
            const_cast<BYTE*>(encoded.data()),
            static_cast<DWORD>(encoded.size()));

        if (FAILED(hr))
            return false;

        ComPtr<IWICBitmapDecoder> decoder;

        hr = factory->CreateDecoderFromStream(
            stream.Get(),
            nullptr,
            WICDecodeMetadataCacheOnLoad,
            &decoder);

        if (FAILED(hr))
            return false;

        ComPtr<IWICBitmapFrameDecode> frame;

        hr = decoder->GetFrame(0, &frame);
        if (FAILED(hr))
            return false;

        UINT width = 0;
        UINT height = 0;

        if (FAILED(frame->GetSize(&width, &height)) ||
            width == 0 || height == 0 ||
            width > 8192 || height > 8192)
        {
            return false;
        }

        // Bound the decoded allocation as well as the downloaded bytes.
        const uint64_t pixelCount =
            static_cast<uint64_t>(width) * height;

        if (pixelCount > 16ull * 1024 * 1024)
            return false;

        ComPtr<IWICFormatConverter> converter;

        hr = factory->CreateFormatConverter(&converter);
        if (FAILED(hr))
            return false;

        hr = converter->Initialize(
            frame.Get(),
            GUID_WICPixelFormat24bppBGR,
            WICBitmapDitherTypeNone,
            nullptr,
            0.0,
            WICBitmapPaletteTypeCustom);

        if (FAILED(hr))
            return false;

        // DIB scanlines are padded to four-byte boundaries.
        const UINT stride = (width * 3u + 3u) & ~3u;
        const uint64_t pixelBytes64 =
            static_cast<uint64_t>(stride) * height;

        if (pixelBytes64 > (std::numeric_limits<UINT>::max)())
            return false;

        const UINT pixelBytes = static_cast<UINT>(pixelBytes64);

        std::vector<BYTE> pixels(pixelBytes);

        hr = converter->CopyPixels(
            nullptr, stride, pixelBytes, pixels.data());

        if (FAILED(hr))
            return false;

        // WIC CopyPixels returns scanlines top-to-bottom, while a
        // positive-height BI_RGB DIB stores them bottom-to-top. Reverse
        // the rows before serializing the DIB into RTF; otherwise the
        // pasted image appears vertically flipped.
        for (UINT y = 0; y < height / 2; ++y)
        {
            BYTE* top = pixels.data() +
                static_cast<size_t>(y) * stride;
            BYTE* bottom = pixels.data() +
                static_cast<size_t>(height - 1 - y) * stride;

            std::swap_ranges(top, top + stride, bottom);
        }

        BITMAPINFOHEADER bih = {};
        bih.biSize = sizeof(bih);
        bih.biWidth = static_cast<LONG>(width);
        bih.biHeight = static_cast<LONG>(height);
        bih.biPlanes = 1;
        bih.biBitCount = 24;
        bih.biCompression = BI_RGB;
        bih.biSizeImage = pixelBytes;

        // RTF dimensions are expressed in twips (1/1440 inch).
        const int picwgoal = static_cast<int>(
            std::min<uint64_t>(
                static_cast<uint64_t>(width) * 15,
                1000000));

        const int pichgoal = static_cast<int>(
            std::min<uint64_t>(
                static_cast<uint64_t>(height) * 15,
                1000000));

        rtf += "{\\pict\\dibitmap0";
        rtf += "\\picw" + std::to_string(width);
        rtf += "\\pich" + std::to_string(height);
        rtf += "\\picwgoal" + std::to_string(picwgoal);
        rtf += "\\pichgoal" + std::to_string(pichgoal);
        rtf += "\n";

        AppendHex(
            rtf,
            reinterpret_cast<const BYTE*>(&bih),
            sizeof(bih));

        AppendHex(rtf, pixels.data(), pixels.size());

        rtf += "}";

        return true;
    }

    std::wstring DecodeHtmlEntities(const std::wstring& text)
    {
        std::wstring result;
        result.reserve(text.size());

        for (size_t i = 0; i < text.size(); )
        {
            if (text[i] != L'&')
            {
                result.push_back(text[i++]);
                continue;
            }

            const size_t semicolon = text.find(L';', i + 1);

            if (semicolon == std::wstring::npos ||
                semicolon - i > 12)
            {
                result.push_back(text[i++]);
                continue;
            }

            const std::wstring entity =
                text.substr(i + 1, semicolon - i - 1);

            wchar_t decoded = 0;

            if (entity == L"amp") decoded = L'&';
            else if (entity == L"lt") decoded = L'<';
            else if (entity == L"gt") decoded = L'>';
            else if (entity == L"quot") decoded = L'"';
            else if (entity == L"apos" || entity == L"#39")
                decoded = L'\'';
            else if (entity == L"nbsp") decoded = L' ';
            else if (entity.size() > 1 && entity[0] == L'#')
            {
                wchar_t* tail = nullptr;
                const bool hex = entity.size() > 2 &&
                    (entity[1] == L'x' || entity[1] == L'X');

                const unsigned long value = wcstoul(
                    entity.c_str() + (hex ? 2 : 1),
                    &tail, hex ? 16 : 10);

                if (tail && *tail == 0 &&
                    value > 0 && value <= 0xFFFF &&
                    !(value >= 0xD800 && value <= 0xDFFF))
                {
                    decoded = static_cast<wchar_t>(value);
                }
            }

            if (decoded)
            {
                result.push_back(decoded);
                i = semicolon + 1;
            }
            else
            {
                result.append(text, i, semicolon - i + 1);
                i = semicolon + 1;
            }
        }

        return result;
    }

    std::wstring LowerHtml(std::wstring value)
    {
        std::transform(
            value.begin(), value.end(), value.begin(),
            [](wchar_t c)
            {
                return static_cast<wchar_t>(towlower(c));
            });
        return value;
    }

    std::wstring GetHtmlAttribute(
        const std::wstring& tag,
        const std::wstring& attribute)
    {
        const std::wstring lower = LowerHtml(tag);
        const std::wstring key = LowerHtml(attribute);
        size_t p = 0;

        while ((p = lower.find(key, p)) != std::wstring::npos)
        {
            const bool leftBoundary =
                p == 0 || iswspace(lower[p - 1]) ||
                lower[p - 1] == L'<';

            size_t q = p + key.size();

            while (q < lower.size() && iswspace(lower[q]))
                ++q;

            if (!leftBoundary || q >= lower.size() ||
                lower[q] != L'=')
            {
                p += key.size();
                continue;
            }

            ++q;
            while (q < tag.size() && iswspace(tag[q]))
                ++q;

            if (q >= tag.size())
                return {};

            const wchar_t quote =
                (tag[q] == L'\'' || tag[q] == L'"') ? tag[q++] : 0;

            const size_t end = quote
                ? tag.find(quote, q)
                : tag.find_first_of(L" \t\r\n>", q);

            return tag.substr(
                q, (end == std::wstring::npos ? tag.size() : end) - q);
        }

        return {};
    }

    HtmlImageCache PrefetchRemoteImages(const std::wstring& html)
    {
        std::vector<std::wstring> urls;
        std::map<std::wstring, size_t> urlIndices;
        const std::wstring lowerHtml = LowerHtml(html);

        // Extract only <img ...> tags; preserve the existing attribute parser.
        for (size_t pos = 0; pos < html.size() && urls.size() < MaxRemoteImages;)
        {
            const size_t open = lowerHtml.find(L"<img", pos);
            if (open == std::wstring::npos)
                break;

            const size_t nameEnd = open + 4;
            if (nameEnd < html.size() &&
                !iswspace(html[nameEnd]) &&
                html[nameEnd] != L'>' &&
                html[nameEnd] != L'/')
            {
                pos = nameEnd;
                continue;
            }

            const size_t end = html.find(L'>', nameEnd);
            if (end == std::wstring::npos)
                break;

            const std::wstring tag = html.substr(open + 1, end - open - 1);
            const std::wstring src = GetHtmlAttribute(tag, L"src");

            if (src.size() <= 8192 &&
                src.size() >= 8 &&
                _wcsnicmp(src.c_str(), L"https://", 8) == 0 &&
                urlIndices.find(src) == urlIndices.end())
            {
                urlIndices.emplace(src, urls.size());
                urls.push_back(src);
            }

            pos = end + 1;
        }

        HtmlImageCache cache;
        if (urls.empty())
            return cache;

        CWaitCursor wait;

        const Clock::time_point deadline =
            Clock::now() + std::chrono::seconds(5);

        std::vector<std::vector<BYTE>> results(urls.size());
        std::vector<unsigned char> succeeded(urls.size(), 0);
        std::atomic<size_t> next{ 0 };

        const size_t workerCount =
            (std::min)(MaxParallelDownloads, urls.size());

        std::vector<std::thread> workers;
        workers.reserve(workerCount);

        for (size_t n = 0; n < workerCount; ++n)
        {
            workers.emplace_back([&]()
            {
                for (;;)
                {
                    const size_t i = next.fetch_add(1);
                    if (i >= urls.size() || Clock::now() >= deadline)
                        return;

                    succeeded[i] = DownloadHttpsImage(
                        urls[i], results[i], deadline) ? 1 : 0;
                }
            });
        }

        for (auto& worker : workers)
            worker.join();

        size_t totalBytes = 0;
        for (size_t i = 0; i < urls.size(); ++i)
        {
            if (!succeeded[i] || results[i].empty())
                continue;

            if (results[i].size() > MaxCachedImageBytes - totalBytes)
                continue;

            totalBytes += results[i].size();
            cache.emplace(urls[i], std::move(results[i]));
        }

        return cache;
    }

    void AppendRtfAlignment(
        std::string& rtf,
        const std::wstring& tag)
    {
        const std::wstring style =
            LowerHtml(GetHtmlAttribute(tag, L"style"));

        if (style.find(L"text-align:center") != std::wstring::npos)
            rtf += "\\qc ";
        else if (style.find(L"text-align:right") != std::wstring::npos)
            rtf += "\\qr ";
        else if (style.find(L"text-align:justify") != std::wstring::npos)
            rtf += "\\qj ";
        else
            rtf += "\\ql ";
    }


    bool AppendRemoteRtfImage(
        std::string& rtf,
        const std::wstring& tag,
        const HtmlImageCache& imageCache)
    {
        const std::wstring src = GetHtmlAttribute(tag, L"src");

        if (src.empty() || src.size() > 8192)
            return false;

        const auto it = imageCache.find(src);
        if (it == imageCache.end())
            return false;

        // Keep the original, known-working DIB-to-RTF serialization.
        return AppendBitmapRtf(rtf, it->second);
    }

    bool DecodeBase64(
        const std::wstring& input,
        std::vector<unsigned char>& output)
    {
        static const char alphabet[] =
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

        int val = 0;
        int bits = -8;

        output.clear();

        for (wchar_t ch : input)
        {
            if (ch == L'=')
                break;

            if (iswspace(ch))
                continue;

            if (ch > 127)
                return false;

            const char* p = strchr(alphabet, static_cast<char>(ch));
            if (!p)
                return false;

            val = (val << 6) | static_cast<int>(p - alphabet);
            bits += 6;

            if (bits >= 0)
            {
                output.push_back(
                    static_cast<unsigned char>((val >> bits) & 0xFF));
                bits -= 8;

                // Defend against unexpectedly large clipboard payloads.
                if (output.size() > 16 * 1024 * 1024)
                    return false;
            }
        }

        return !output.empty();
    }

    bool AppendRtfImage(
        std::string& rtf,
        const std::wstring& tag)
    {
        const std::wstring src = GetHtmlAttribute(tag, L"src");
        const std::wstring lowerSrc = LowerHtml(src);

        const bool isPng =
            lowerSrc.find(L"data:image/png;base64,") == 0;
        const bool isJpeg =
            lowerSrc.find(L"data:image/jpeg;base64,") == 0 ||
            lowerSrc.find(L"data:image/jpg;base64,") == 0;

        // Only embedded raster data is handled here. Arbitrary URLs are
        // intentionally not downloaded during a clipboard paste.
        if (!isPng && !isJpeg)
            return false;

        const size_t comma = src.find(L',');
        if (comma == std::wstring::npos)
            return false;

        std::vector<unsigned char> bytes;
        if (!DecodeBase64(src.substr(comma + 1), bytes))
            return false;

        // Validate the file signature before placing data into the RTF.
        if (isPng)
        {
            static const unsigned char signature[] =
            { 0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A };

            if (bytes.size() < sizeof(signature) ||
                memcmp(bytes.data(), signature, sizeof(signature)) != 0)
                return false;
        }
        else
        {
            if (bytes.size() < 3 ||
                bytes[0] != 0xFF ||
                bytes[1] != 0xD8 ||
                bytes[2] != 0xFF)
                return false;
        }

        // RTF picture dimensions are expressed in twips.
        // Use conservative default dimensions; Rich Edit can scale the
        // picture to fit the available line width.
        int widthTwips = 3600;
        int heightTwips = 2400;

        auto parseDimension = [](const std::wstring& value, int& twips)
            {
                if (value.empty())
                    return;

                wchar_t* end = nullptr;
                const long pixels = wcstol(value.c_str(), &end, 10);

                if (end != value.c_str() && pixels > 0 && pixels <= 4000)
                    twips = static_cast<int>(pixels * 15);
            };

        parseDimension(GetHtmlAttribute(tag, L"width"), widthTwips);
        parseDimension(GetHtmlAttribute(tag, L"height"), heightTwips);

        rtf += "{\\pict";
        rtf += isPng ? "\\pngblip" : "\\jpegblip";
        rtf += "\\picwgoal" + std::to_string(widthTwips);
        rtf += "\\pichgoal" + std::to_string(heightTwips);
        rtf += "\n";

        static const char hex[] = "0123456789abcdef";

        for (unsigned char byte : bytes)
        {
            rtf.push_back(hex[byte >> 4]);
            rtf.push_back(hex[byte & 15]);
        }

        rtf += "}\n";
        return true;
    }

    std::string HtmlToRtf(const std::wstring& html, const HtmlImageCache& imageCache)
    {
        std::string rtf =
            "{\\rtf1\\ansi\\deff0"
            "{\\fonttbl{\\f0\\fnil Segoe UI;}}"
            "\\viewkind4\\uc1 ";

        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool inPre = false;
        bool paragraphStarted = false;
        int listDepth = 0;

        auto paragraph = [&]()
            {
                if (paragraphStarted)
                    rtf += "\\par\n";

                paragraphStarted = false;
            };

        size_t i = 0;

        while (i < html.size())
        {
            if (html[i] != L'<')
            {
                size_t end = html.find(L'<', i);
                if (end == std::wstring::npos)
                    end = html.size();

                std::wstring text = DecodeHtmlEntities(
                    html.substr(i, end - i));

                if (!inPre)
                {
                    std::wstring collapsed;
                    bool previousSpace = false;

                    for (wchar_t ch : text)
                    {
                        if (iswspace(ch))
                        {
                            if (!previousSpace)
                                collapsed.push_back(L' ');
                            previousSpace = true;
                        }
                        else
                        {
                            collapsed.push_back(ch);
                            previousSpace = false;
                        }
                    }

                    text.swap(collapsed);
                }

                AppendRtfText(rtf, text);
                if (!text.empty())
                    paragraphStarted = true;
                i = end;
                continue;
            }

            const size_t end = html.find(L'>', i + 1);
            if (end == std::wstring::npos)
            {
                AppendRtfText(rtf, html.substr(i));
                break;
            }

            std::wstring tag = html.substr(i + 1, end - i - 1);
            const std::wstring lower = LowerHtml(tag);

            i = end + 1;

            if (lower.empty() || lower[0] == L'!' ||
                lower[0] == L'?')
                continue;

            const bool closing = lower[0] == L'/';
            size_t nameStart = closing ? 1 : 0;
            size_t nameEnd = nameStart;

            while (nameEnd < lower.size() &&
                (iswalnum(lower[nameEnd]) || lower[nameEnd] == L'-'))
                ++nameEnd;

            const std::wstring name =
                lower.substr(nameStart, nameEnd - nameStart);

            if (name == L"script" || name == L"style")
            {
                if (!closing)
                {
                    const std::wstring closeTag = L"</" + name;
                    const std::wstring remaining =
                        LowerHtml(html.substr(i));
                    const size_t p = remaining.find(closeTag);

                    if (p != std::wstring::npos)
                    {
                        const size_t closeEnd =
                            html.find(L'>', i + p);

                        i = closeEnd == std::wstring::npos
                            ? html.size() : closeEnd + 1;
                    }
                }
                continue;
            }

            if (name == L"img" && !closing)
            {
                const bool inserted =
                    AppendRtfImage(rtf, tag) ||
                    AppendRemoteRtfImage(rtf, tag, imageCache);

                if (!inserted)
                {
                    const std::wstring alt =
                        GetHtmlAttribute(tag, L"alt");

                    if (!alt.empty())
                        AppendRtfText(rtf, alt);
                }

                paragraphStarted = true;
            }
            else if (name == L"br" && !closing)
            {
                rtf += "\\line ";
                paragraphStarted = true;
            }
            else if (name == L"p" || name == L"div" ||
                name == L"h1" || name == L"h2" ||
                name == L"h3" || name == L"h4" ||
                name == L"li" || name == L"blockquote")
            {
                paragraph();

                if (!closing)
                {
                    if (name == L"p" || name == L"div" ||
                        name == L"h1" || name == L"h2" ||
                        name == L"h3" || name == L"h4")
                    {
                        AppendRtfAlignment(rtf, tag);
                    }

                    if (name == L"li")
                    {
                        if (listDepth > 0)
                            rtf += "\\tab ";
                        rtf += "\\bullet\\tab ";
                        paragraphStarted = true;
                    }

                    if (name == L"blockquote")
                        rtf += "\\li720 ";

                    if (name == L"h1") rtf += "\\fs36\\b ";
                    else if (name == L"h2") rtf += "\\fs30\\b ";
                    else if (name == L"h3") rtf += "\\fs26\\b ";
                    else if (name == L"h4") rtf += "\\fs24\\b ";
                }
                else
                {
                    if (name == L"h1" || name == L"h2" ||
                        name == L"h3" || name == L"h4")
                    {
                        rtf += "\\b0\\fs20 ";
                    }

                    if (name == L"blockquote")
                        rtf += "\\li0 ";

                    if (name == L"p" || name == L"div" ||
                        name == L"h1" || name == L"h2" ||
                        name == L"h3" || name == L"h4")
                    {
                        rtf += "\\ql ";
                    }
                }
            }
            else if (name == L"ul" || name == L"ol")
            {
                if (!closing)
                    ++listDepth;
                else if (listDepth > 0)
                    --listDepth;
            }
            else if (name == L"b" || name == L"strong")
            {
                bold = !closing;
                rtf += bold ? "\\b " : "\\b0 ";
            }
            else if (name == L"i" || name == L"em")
            {
                italic = !closing;
                rtf += italic ? "\\i " : "\\i0 ";
            }
            else if (name == L"u")
            {
                underline = !closing;
                rtf += underline ? "\\ul " : "\\ul0 ";
            }
            else if (name == L"s" || name == L"strike" ||
                name == L"del")
            {
                rtf += closing ? "\\strike0 " : "\\strike ";
            }
            else if (name == L"pre")
            {
                inPre = !closing;
                rtf += closing
                    ? "\\f0\\fs20 "
                    : "\\f0\\fs20 ";
            }
            else if (name == L"sub")
            {
                rtf += closing ? "\\sub0 " : "\\sub ";
            }
            else if (name == L"sup")
            {
                rtf += closing ? "\\super0 " : "\\super ";
            }
            else if (name == L"font" || name == L"span")
            {
                const std::wstring style =
                    LowerHtml(GetHtmlAttribute(tag, L"style"));

                const size_t sizePos = style.find(L"font-size:");
                if (sizePos != std::wstring::npos)
                {
                    const wchar_t* value =
                        style.c_str() + sizePos + 10;

                    while (iswspace(*value))
                        ++value;

                    wchar_t* tail = nullptr;
                    const double points = wcstod(value, &tail);

                    if (tail != value && points >= 4 && points <= 72)
                    {
                        const int halfPoints =
                            static_cast<int>(points * 2 + 0.5);

                        rtf += "\\fs" + std::to_string(halfPoints) + " ";
                    }
                }

                const size_t weightPos = style.find(L"font-weight:");
                if (weightPos != std::wstring::npos)
                {
                    const std::wstring weight =
                        style.substr(weightPos + 12, 12);

                    if (weight.find(L"bold") != std::wstring::npos ||
                        weight.find(L"600") != std::wstring::npos ||
                        weight.find(L"700") != std::wstring::npos ||
                        weight.find(L"800") != std::wstring::npos ||
                        weight.find(L"900") != std::wstring::npos)
                    {
                        bold = !closing;
                        rtf += bold ? "\\b " : "\\b0 ";
                    }
                }

                const size_t italicPos = style.find(L"font-style:");
                if (italicPos != std::wstring::npos)
                {
                    italic = !closing;
                    rtf += italic ? "\\i " : "\\i0 ";
                }

                const size_t decorationPos =
                    style.find(L"text-decoration:");

                if (decorationPos != std::wstring::npos)
                {
                    underline = !closing;
                    rtf += underline ? "\\ul " : "\\ul0 ";
                }
            }
            else if (name == L"a")
            {
                // Preserve hyperlink text and underline it.
                underline = !closing;
                rtf += underline ? "\\ul " : "\\ul0 ";
            }
        }

        // Silence compiler warnings for state retained for tag processing.
        (void)bold;
        (void)italic;
        (void)underline;

        rtf += "}";
        return rtf;
    }


    struct HtmlRtfStream
    {
        const std::string* data = nullptr;
        size_t position = 0;
    };

    DWORD CALLBACK HtmlRtfStreamCallback(
        DWORD_PTR cookie,
        LPBYTE buffer,
        LONG bytes,
        LONG* bytesWritten)
    {
        if (!cookie || !buffer || !bytesWritten || bytes < 0)
            return 1;

        HtmlRtfStream* stream =
            reinterpret_cast<HtmlRtfStream*>(cookie);

        if (!stream->data)
            return 1;

        const size_t remaining =
            stream->position < stream->data->size()
            ? stream->data->size() - stream->position
            : 0;

        const size_t count = (std::min)(
            remaining, static_cast<size_t>(bytes));

        if (count != 0)
        {
            memcpy(
                buffer,
                stream->data->data() + stream->position,
                count);
        }

        stream->position += count;
        *bytesWritten = static_cast<LONG>(count);
        return 0;
    }

} // namespace

HRESULT PasteHtmlIntoRichEdit(
    HWND hwndRichEdit,
    LPDATAOBJECT dataObject,
    CLIPFORMAT htmlFormat)
{
    if (!::IsWindow(hwndRichEdit) || !dataObject || !htmlFormat)
        return E_INVALIDARG;

    std::wstring html;

    if (!ReadHtmlClipboard(dataObject, htmlFormat, html))
        return DV_E_FORMATETC;

    // Fetch remote images concurrently before constructing the RTF.
    // The original bitmap serialization is intentionally unchanged.
    const HtmlImageCache imageCache = PrefetchRemoteImages(html);
    const std::string rtf = HtmlToRtf(html, imageCache);

    if (rtf.empty())
        return E_FAIL;

    HtmlRtfStream stream;
    stream.data = &rtf;

    EDITSTREAM editStream = {};
    editStream.dwCookie =
        reinterpret_cast<DWORD_PTR>(&stream);
    editStream.pfnCallback = HtmlRtfStreamCallback;

    // SFF_SELECTION replaces the selected text rather than the document.
    const LRESULT inserted = ::SendMessage(
        hwndRichEdit,
        EM_STREAMIN,
        SF_RTF | SFF_SELECTION,
        reinterpret_cast<LPARAM>(&editStream));

    if (editStream.dwError != 0)
        return HRESULT_FROM_WIN32(editStream.dwError);

    if (inserted == 0 && !rtf.empty())
        return E_FAIL;

    return S_OK;
}
