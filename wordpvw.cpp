//
// wordpvw.cpp : implementation of the CWordPadView class
//
// This is a part of the Microsoft Foundation Classes C++ library.
// Copyright (C) 1992-1998 Microsoft Corporation
// All rights reserved.
//
// This source code is only intended as a supplement to the
// Microsoft Foundation Classes Reference and related
// electronic documentation provided with the library.
// See these sources for detailed information regarding the
// Microsoft Foundation Classes product.

#include "stdafx.h"
#include "wordpad.h"
#include "cntritem.h"
#include "srvritem.h"

#include "wordpdoc.h"
#include "wordpvw.h"
#include "formatta.h"
#include "datedial.h"
#include "formatpa.h"
#include "ruler.h"
#include "strings.h"
#include "pageset.h"
//#include <penwin.h>

#include <tom.h>

#include <cstdlib>

#include <string>
#include <vector>
#include <algorithm>
#include <cstdint>
#include <climits>

extern CLIPFORMAT cfEmbeddedObject;
extern CLIPFORMAT cfRTO;
extern CLIPFORMAT cfHTML;

#ifdef _DEBUG
#undef THIS_FILE
static char BASED_CODE THIS_FILE[] = __FILE__;
#endif



namespace
{

#include <windows.h>
#include <objidl.h>

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

    static bool ReadHtmlClipboard(
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

    static void AppendRtfText(
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

    static std::wstring DecodeHtmlEntities(const std::wstring& text)
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

    static std::wstring LowerHtml(std::wstring value)
    {
        std::transform(
            value.begin(), value.end(), value.begin(),
            [](wchar_t c)
            {
                return static_cast<wchar_t>(towlower(c));
            });
        return value;
    }

    static std::wstring GetHtmlAttribute(
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



    static bool DecodeBase64(
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

    static bool AppendRtfImage(
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

    static std::string HtmlToRtf(const std::wstring& html)
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

                paragraphStarted = true;
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
                if (!AppendRtfImage(rtf, tag))
                {
                    // Unsupported or unavailable image: retain its alternative
                    // text rather than silently losing all indication of it.
                    const std::wstring alt = GetHtmlAttribute(tag, L"alt");
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

                if (name == L"li" && !closing)
                {
                    if (listDepth > 0)
                        rtf += "\\tab ";
                    rtf += "\\bullet\\tab ";
                }

                if (name == L"h1") rtf += "\\fs36\\b ";
                else if (name == L"h2") rtf += "\\fs30\\b ";
                else if (name == L"h3") rtf += "\\fs26\\b ";
                else if (name == L"h4") rtf += "\\fs24\\b ";
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

    static DWORD CALLBACK HtmlRtfStreamCallback(
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

    static HRESULT PasteHtmlIntoRichEdit(
        HWND hwndRichEdit,
        LPDATAOBJECT dataObject,
        CLIPFORMAT htmlFormat)
    {
        if (!::IsWindow(hwndRichEdit) || !dataObject || !htmlFormat)
            return E_INVALIDARG;

        std::wstring html;

        if (!ReadHtmlClipboard(dataObject, htmlFormat, html))
            return DV_E_FORMATETC;

        const std::string rtf = HtmlToRtf(html);

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
}


BOOL CCharFormat::operator==(CCharFormat& cf)
{
	return
	dwMask == cf.dwMask
	&& dwEffects == cf.dwEffects
	&& yHeight == cf.yHeight
	&& yOffset == cf.yOffset
	&& crTextColor == cf.crTextColor
	&& bPitchAndFamily == cf.bPitchAndFamily
	&& (lstrcmp(szFaceName, cf.szFaceName) == 0);
}

BOOL CParaFormat::operator==(WPD_PARAFORMAT& pf)
{
	if(
		dwMask != pf.dwMask
		|| wNumbering != pf.wNumbering
#if _MSC_VER < 1500
		|| wReserved != pf.wReserved
#endif
		|| dxStartIndent != pf.dxStartIndent
		|| dxRightIndent != pf.dxRightIndent
		|| dxOffset != pf.dxOffset
		|| cTabCount != pf.cTabCount
		)
	{
		return FALSE;
	}
	for (int i=0;i<pf.cTabCount;i++)
	{
		if (rgxTabs[i] != pf.rgxTabs[i])
			return FALSE;
	}
	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CWordPadView

IMPLEMENT_DYNCREATE(CWordPadView, CRichEditView)

//WM_SETTINGCHANGE -- default printer might have changed
//WM_FONTCHANGE -- pool of fonts changed
//WM_DEVMODECHANGE -- printer settings changes

BEGIN_MESSAGE_MAP(CWordPadView, CRichEditView)
	ON_COMMAND(ID_CANCEL_EDIT_CNTR, OnCancelEditCntr)
	ON_COMMAND(ID_CANCEL_EDIT_SRVR, OnCancelEditSrvr)
	//{{AFX_MSG_MAP(CWordPadView)
	ON_WM_CREATE()
	ON_COMMAND(ID_PAGE_SETUP, OnPageSetup)
	ON_COMMAND(ID_INSERT_DATE_TIME, OnInsertDateTime)
	ON_COMMAND(ID_FORMAT_PARAGRAPH, OnFormatParagraph)
	ON_COMMAND(ID_FORMAT_TABS, OnFormatTabs)
	ON_WM_TIMER()
	ON_WM_DESTROY()
	ON_WM_MEASUREITEM()
	ON_COMMAND(ID_PEN_BACKSPACE, OnPenBackspace)
	ON_COMMAND(ID_PEN_NEWLINE, OnPenNewline)
	ON_COMMAND(ID_PEN_PERIOD, OnPenPeriod)
	ON_COMMAND(ID_PEN_SPACE, OnPenSpace)
	ON_WM_KEYDOWN()
	ON_COMMAND(ID_FILE_PRINT, OnFilePrint)
#ifdef CORRECT_WRITING
	ON_COMMAND(ID_PEN_LENS, OnPenLens)
#endif
	ON_COMMAND(ID_PEN_TAB, OnPenTab)
	ON_WM_PALETTECHANGED()
	ON_WM_QUERYNEWPALETTE()
	ON_WM_SETTINGCHANGE()
	ON_WM_SIZE()
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
	ON_COMMAND(ID_CHAR_COLOR, OnCharColor)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, OnFilePrintPreview)
	ON_COMMAND(IDC_FONTNAME, OnFontname)
	ON_COMMAND(IDC_FONTSIZE, OnFontsize)
	ON_COMMAND(ID_CHAR_BOLD, OnCharBold)
	ON_UPDATE_COMMAND_UI(ID_CHAR_BOLD, OnUpdateCharBold)
	ON_COMMAND(ID_CHAR_ITALIC, OnCharItalic)
	ON_UPDATE_COMMAND_UI(ID_CHAR_ITALIC, OnUpdateCharItalic)
	ON_COMMAND(ID_CHAR_UNDERLINE, OnCharUnderline)
	ON_UPDATE_COMMAND_UI(ID_CHAR_UNDERLINE, OnUpdateCharUnderline)
	ON_COMMAND(ID_PARA_CENTER, OnParaCenter)
	ON_UPDATE_COMMAND_UI(ID_PARA_CENTER, OnUpdateParaCenter)
	ON_COMMAND(ID_PARA_LEFT, OnParaLeft)
	ON_UPDATE_COMMAND_UI(ID_PARA_LEFT, OnUpdateParaLeft)
	ON_COMMAND(ID_PARA_RIGHT, OnParaRight)
	ON_UPDATE_COMMAND_UI(ID_PARA_RIGHT, OnUpdateParaRight)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, OnFilePrint)
	ON_WM_DROPFILES()
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_INSERT_BULLET, CRichEditView::OnBullet)
	ON_UPDATE_COMMAND_UI(ID_INSERT_BULLET, CRichEditView::OnUpdateBullet)
	ON_EN_CHANGE(AFX_IDW_PANE_FIRST, OnEditChange)
	ON_WM_MOUSEACTIVATE()
	ON_NOTIFY_RANGE(NM_SETFOCUS, AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST, OnBarSetFocus)
	ON_NOTIFY_RANGE(NM_KILLFOCUS, AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST, OnBarKillFocus)
	ON_NOTIFY_RANGE(NM_RETURN, AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST, OnBarReturn)
	ON_CBN_SELENDOK(IDC_FONTNAME, OnFontname)
	ON_CBN_SELENDOK(IDC_FONTSIZE, OnFontsize)
	ON_COMMAND_RANGE(ID_BORDER_1, ID_BORDER_13, OnBorderType)
	ON_UPDATE_COMMAND_UI_RANGE(ID_BORDER_1, ID_BORDER_13, OnUpdateBorderType)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWordPadView construction/destruction

CWordPadView::CWordPadView()
{
	m_bSyncCharFormat = m_bSyncParaFormat = TRUE;
	m_uTimerID = 0;
	m_bDelayUpdateItems = FALSE;
	m_bOnBar = FALSE;
	m_bInPrint = FALSE;
	m_nPasteType = 0;
	m_rectMargin = theApp.m_rectPageMargin;
	m_nBorderType = ID_BORDER_1;
}

BOOL CWordPadView::PreCreateWindow(CREATESTRUCT& cs)
{
	BOOL bRes = CRichEditView::PreCreateWindow(cs);

    if (bRes)
        cs.lpszClass = MSFTEDIT_CLASS;

	cs.style |= ES_SELECTIONBAR;
	return bRes;
}

/////////////////////////////////////////////////////////////////////////////
// CWordPadView attributes

BOOL CWordPadView::IsFormatText()
{
	// this function checks to see if any formatting is not default text
	BOOL bRes = FALSE;
	CHARRANGE cr;
	CCharFormat cf;
	CParaFormat pf;
	GetRichEditCtrl().GetSel(cr);
	GetRichEditCtrl().HideSelection(TRUE, FALSE);
	GetRichEditCtrl().SetSel(0,-1);

	if (!(GetRichEditCtrl().GetSelectionType() & (SEL_OBJECT|SEL_MULTIOBJECT)))
	{
		GetRichEditCtrl().GetSelectionCharFormat(cf);
		if (cf == m_defTextCharFormat)
		{
			GetRichEditCtrl().GetParaFormat(pf);
			if (pf == m_defParaFormat) //compared using CParaFormat::operator==
				bRes = TRUE;
		}
	}

	GetRichEditCtrl().SetSel(cr);
	GetRichEditCtrl().HideSelection(FALSE, FALSE);
	return bRes;
}

HMENU CWordPadView::GetContextMenu(WORD, LPOLEOBJECT, CHARRANGE* )
{
	return NULL;
}

/////////////////////////////////////////////////////////////////////////////
// CWordPadView operations

void CWordPadView::WrapChanged()
{
	CWaitCursor wait;
	CFrameWnd* pFrameWnd = GetParentFrame();
	ASSERT(pFrameWnd != NULL);
	if (pFrameWnd)
	{
		pFrameWnd->SetMessageText(IDS_FORMATTING);
		CWnd* pBarWnd = pFrameWnd->GetMessageBar();
		if (pBarWnd != NULL)
			pBarWnd->UpdateWindow();

		CRichEditView::WrapChanged();

		pFrameWnd->SetMessageText(AFX_IDS_IDLEMESSAGE);
		if (pBarWnd != NULL)
			pBarWnd->UpdateWindow();
	}
}

void CWordPadView::SetUpdateTimer()
{
	if (m_uTimerID != 0) // if outstanding timer kill it
		KillTimer(m_uTimerID);
	m_uTimerID = SetTimer(1, 1000, NULL); //set a timer for 1000 milliseconds
	if (m_uTimerID == 0) // no timer available so force update now
		GetDocument()->UpdateAllItems(NULL);
	else
		m_bDelayUpdateItems = TRUE;
}

void CWordPadView::DeleteContents()
{
	ASSERT_VALID(this);
	ASSERT(m_hWnd != NULL);
	CRichEditView::DeleteContents();
	SetDefaultFont(IsTextType(GetDocument()->m_nNewDocType));
}

void CWordPadView::SetDefaultFont(BOOL bText)
{
	ASSERT_VALID(this);
	ASSERT(m_hWnd != NULL);
	m_bSyncCharFormat = m_bSyncParaFormat = TRUE;
	WPD_CHARFORMAT* pCharFormat = bText ? &m_defTextCharFormat : &m_defCharFormat;
	// set the default character format -- the FALSE makes it the default
	GetRichEditCtrl().SetSel(0,-1);
	GetRichEditCtrl().SetDefaultCharFormat(*pCharFormat);
	GetRichEditCtrl().SetSelectionCharFormat(*pCharFormat);

	GetRichEditCtrl().SetParaFormat(m_defParaFormat);

	GetRichEditCtrl().SetSel(0,0);
	GetRichEditCtrl().EmptyUndoBuffer();
	GetRichEditCtrl().SetModify(FALSE);
	ASSERT_VALID(this);
}

/////////////////////////////////////////////////////////////////////////////
// CWordPadView drawing

/////////////////////////////////////////////////////////////////////////////
// CWordPadView printing

void CWordPadView::OnPrint(CDC* pDC, CPrintInfo* pInfo)
{
	CRichEditView::OnPrint(pDC, pInfo);
	if (pInfo != NULL && pInfo->m_bPreview)
		DrawMargins(pDC);
}

void CWordPadView::DrawMargins(CDC* pDC)
{
	if (pDC->m_hAttribDC != NULL)
	{
		CRect rect;
		rect.left = m_rectMargin.left;
		rect.right = m_sizePaper.cx - m_rectMargin.right;
		rect.top = m_rectMargin.top;
		rect.bottom = m_sizePaper.cy - m_rectMargin.bottom;
		//rect in twips
		int logx = ::GetDeviceCaps(pDC->m_hDC, LOGPIXELSX);
		int logy = ::GetDeviceCaps(pDC->m_hDC, LOGPIXELSY);
		rect.left = MulDiv(rect.left, logx, 1440);
		rect.right = MulDiv(rect.right, logx, 1440);
		rect.top = MulDiv(rect.top, logy, 1440);
		rect.bottom = MulDiv(rect.bottom, logy, 1440);
		CPen pen(PS_DOT, 0, pDC->GetTextColor());
		CPen* ppen = pDC->SelectObject(&pen);
		pDC->MoveTo(0, rect.top);
		pDC->LineTo(10000, rect.top);
		pDC->MoveTo(rect.left, 0);
		pDC->LineTo(rect.left, 10000);
		pDC->MoveTo(0, rect.bottom);
		pDC->LineTo(10000, rect.bottom);
		pDC->MoveTo(rect.right, 0);
		pDC->LineTo(rect.right, 10000);
		pDC->SelectObject(ppen);
	}
}

BOOL CWordPadView::OnPreparePrinting(CPrintInfo* pInfo)
{
	return DoPreparePrinting(pInfo);
}

/////////////////////////////////////////////////////////////////////////////
// OLE Client support and commands

inline int roundleast(int n)
{
	int mod = n%10;
	n -= mod;
	if (mod >= 5)
		n += 10;
	else if (mod <= -5)
		n -= 10;
	return n;
}

static void RoundRect(LPRECT r1)
{
	r1->left = roundleast(r1->left);
	r1->right = roundleast(r1->right);
	r1->top = roundleast(r1->top);
	r1->bottom = roundleast(r1->bottom);
}

static void MulDivRect(LPRECT r1, LPRECT r2, int num, int div)
{
	r1->left = MulDiv(r2->left, num, div);
	r1->top = MulDiv(r2->top, num, div);
	r1->right = MulDiv(r2->right, num, div);
	r1->bottom = MulDiv(r2->bottom, num, div);
}

void CWordPadView::OnPageSetup()
{
	CPageSetupDialog dlg;
	PAGESETUPDLG& psd = dlg.m_psd;
	BOOL bMetric = theApp.GetUnits() == 1; //centimeters
	psd.Flags |= PSD_MARGINS | (bMetric ? PSD_INHUNDREDTHSOFMILLIMETERS :
		PSD_INTHOUSANDTHSOFINCHES);
	int nUnitsPerInch = bMetric ? 2540 : 1000;
	MulDivRect(&psd.rtMargin, m_rectMargin, nUnitsPerInch, 1440);
	RoundRect(&psd.rtMargin);
	// get the current device from the app
	PRINTDLG pd;
	pd.hDevNames = NULL;
	pd.hDevMode = NULL;
	theApp.GetPrinterDeviceDefaults(&pd);
	psd.hDevNames = pd.hDevNames;
	psd.hDevMode = pd.hDevMode;
	if (dlg.DoModal() == IDOK)
	{
		RoundRect(&psd.rtMargin);
		MulDivRect(m_rectMargin, &psd.rtMargin, 1440, nUnitsPerInch);
		theApp.m_rectPageMargin = m_rectMargin;
		theApp.SelectPrinter(psd.hDevNames, psd.hDevMode);
		theApp.NotifyPrinterChanged();
	}
	// PageSetupDlg failed
	if (CommDlgExtendedError() != 0)
	{
		CPageSetupDlg dlgRetry;
		dlgRetry.m_nBottomMargin = m_rectMargin.bottom;
		dlgRetry.m_nLeftMargin = m_rectMargin.left;
		dlgRetry.m_nRightMargin = m_rectMargin.right;
		dlgRetry.m_nTopMargin = m_rectMargin.top;
		if (dlgRetry.DoModal() == IDOK)
		{
			m_rectMargin.SetRect(dlgRetry.m_nLeftMargin, dlgRetry.m_nTopMargin,
				dlgRetry.m_nRightMargin, dlgRetry.m_nBottomMargin);
			// m_page will be changed at this point
			theApp.m_rectPageMargin = m_rectMargin;
			theApp.NotifyPrinterChanged();
		}
	}
}

/////////////////////////////////////////////////////////////////////////////
// OLE Server support

// The following command handler provides the standard keyboard
//  user interface to cancel an in-place editing session.  Here,
//  the server (not the container) causes the deactivation.
void CWordPadView::OnCancelEditSrvr()
{
	GetDocument()->OnDeactivateUI(FALSE);
}

/////////////////////////////////////////////////////////////////////////////
// CWordPadView diagnostics

#ifdef _DEBUG
void CWordPadView::AssertValid() const
{
	CRichEditView::AssertValid();
}

void CWordPadView::Dump(CDumpContext& dc) const
{
	CRichEditView::Dump(dc);
}

CWordPadDoc* CWordPadView::GetDocument() // non-debug version is inline
{
	return (CWordPadDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CWordPadView message helpers

/////////////////////////////////////////////////////////////////////////////
// CWordPadView message handlers

int CWordPadView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CRichEditView::OnCreate(lpCreateStruct) == -1)
		return -1;
	theApp.m_listPrinterNotify.AddTail(m_hWnd);

	if (theApp.m_bWordSel)
		GetRichEditCtrl().SetOptions(ECOOP_OR, ECO_AUTOWORDSELECTION);
	else
		GetRichEditCtrl().SetOptions(ECOOP_AND, ~(DWORD)ECO_AUTOWORDSELECTION);
//  GetRichEditCtrl().SetOptions(ECOOP_OR, ECO_SELECTIONBAR);

	GetDefaultFont(m_defTextCharFormat, IDS_DEFAULTTEXTFONT);
	GetDefaultFont(m_defCharFormat, IDS_DEFAULTFONT);

	GetRichEditCtrl().GetParaFormat(m_defParaFormat);
	m_defParaFormat.cTabCount = 0;

	SetTextColor ((COLORREF) -1);	// Automatic
	return 0;
}

void CWordPadView::GetDefaultFont(CCharFormat& cf, UINT nFontNameID)
{
	USES_CONVERSION;
	CString strDefFont;
	VERIFY(strDefFont.LoadString(nFontNameID));
	ASSERT(cf.cbSize == sizeof(WPD_CHARFORMAT));
	cf.dwMask = CFM_BOLD|CFM_ITALIC|CFM_UNDERLINE|CFM_STRIKEOUT|CFM_SIZE|
		CFM_COLOR|CFM_OFFSET|CFM_PROTECTED;
	cf.dwEffects = CFE_AUTOCOLOR;
	cf.yHeight = 200; //10pt
	cf.yOffset = 0;
	cf.crTextColor = RGB(0, 0, 0);
	cf.bCharSet = 0;
	cf.bPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
	ASSERT(strDefFont.GetLength() < LF_FACESIZE);
	lstrcpyn(cf.szFaceName, strDefFont, LF_FACESIZE);
	cf.dwMask |= CFM_FACE;
}

void CWordPadView::OnInsertDateTime()
{
	CDateDialog dlg;
	if (dlg.DoModal() == IDOK)
		GetRichEditCtrl().ReplaceSel(dlg.m_strSel);;
}

void CWordPadView::OnFormatParagraph()
{
	CFormatParaDlg dlg(GetParaFormatSelection());
	dlg.m_nWordWrap = m_nWordWrap;
	if (dlg.DoModal() == IDOK)
		SetParaFormat(dlg.m_pf);
}

void CWordPadView::OnFormatTabs()
{
	CFormatTabDlg dlg(GetParaFormatSelection());
	if (dlg.DoModal() == IDOK)
		SetParaFormat(dlg.m_pf);
}

void CWordPadView::OnTextNotFound(LPCTSTR lpStr)
{
	ASSERT_VALID(this);
	MessageBeep(0);
	AfxMessageBox(IDS_FINISHED_SEARCH,MB_OK|MB_ICONINFORMATION);
	CRichEditView::OnTextNotFound(lpStr);
}

void CWordPadView::OnTimer(UINT_PTR nIDEvent)
{
	if (m_uTimerID != nIDEvent) // not our timer
		CRichEditView::OnTimer(nIDEvent);
	else
	{
		KillTimer(m_uTimerID); // kill one-shot timer
		m_uTimerID = 0;
		if (m_bDelayUpdateItems)
			GetDocument()->UpdateAllItems(NULL);
		m_bDelayUpdateItems = FALSE;

		// Update document colors:
		CFrameWndEx* pFrameEx = (CFrameWndEx*) GetTopLevelFrame ();
		CMFCColorBar* pColorBar = DYNAMIC_DOWNCAST (CMFCColorBar, 
			pFrameEx->GetPane  (ID_COLOR_TEAROFF));
		
		if (pColorBar != NULL)
		{
			CList<COLORREF,COLORREF> lstDocColors;
			GetDocumentColors (lstDocColors);

			pColorBar->SetDocumentColors (_T("Document's Colors"), lstDocColors);
		}
	}
}

void CWordPadView::OnEditChange()
{
	SetUpdateTimer();
}

void CWordPadView::OnDestroy()
{
	POSITION pos = theApp.m_listPrinterNotify.Find(m_hWnd);
	ASSERT(pos != NULL);
	theApp.m_listPrinterNotify.RemoveAt(pos);

	CRichEditView::OnDestroy();

	if (m_uTimerID != 0) // if outstanding timer kill it
		OnTimer(m_uTimerID);
	ASSERT(m_uTimerID == 0);
}

void CWordPadView::CalcWindowRect(LPRECT lpClientRect, UINT nAdjustType)
{
	CRichEditView::CalcWindowRect(lpClientRect, nAdjustType);

	if (theApp.m_bWin4 && nAdjustType != 0 && (GetStyle() & WS_VSCROLL))
		lpClientRect->right--;

	// if the ruler is visible then slide the view up under the ruler to avoid
	// showing the top border of the view
	if (GetExStyle() & WS_EX_CLIENTEDGE)
	{
		CFrameWndEx* pFrame = DYNAMIC_DOWNCAST (CFrameWndEx, GetParentFrame());
		if (pFrame != NULL)
		{
			CRulerBar* pBar = (CRulerBar*)pFrame->GetPane(ID_VIEW_RULER);
			if (pBar != NULL)
			{
				BOOL bVis = pBar->IsVisible();
				if (pBar->m_bDeferInProgress)
					bVis = !bVis;
				if (bVis)
					lpClientRect->top -= 2;
			}
		}
	}
}

void CWordPadView::OnMeasureItem(int nIDCtl, LPMEASUREITEMSTRUCT lpMIS)
{
	lpMIS->itemID = (UINT)(WORD)lpMIS->itemID;
	CRichEditView::OnMeasureItem(nIDCtl, lpMIS);
}

void CWordPadView::OnPenBackspace()
{
	SendMessage(WM_KEYDOWN, VK_BACK, 0);
	SendMessage(WM_KEYUP, VK_BACK, 0);
}

void CWordPadView::OnPenNewline()
{
	SendMessage(WM_CHAR, '\n', 0);
}

void CWordPadView::OnPenPeriod()
{
	SendMessage(WM_CHAR, '.', 0);
}

void CWordPadView::OnPenSpace()
{
	SendMessage(WM_CHAR, ' ', 0);
}

void CWordPadView::OnPenTab()
{
	SendMessage(WM_CHAR, VK_TAB, 0);
}

void CWordPadView::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (nChar == VK_F10 && GetKeyState(VK_SHIFT) < 0)
	{
		long nStart, nEnd;
		GetRichEditCtrl().GetSel(nStart, nEnd);
		CPoint pt = GetRichEditCtrl().GetCharPos(nEnd);
		SendMessage(WM_CONTEXTMENU, (WPARAM)m_hWnd, MAKELPARAM(pt.x, pt.y));
	}

	CRichEditView::OnKeyDown(nChar, nRepCnt, nFlags);
}

HRESULT CWordPadView::GetClipboardData(CHARRANGE* lpchrg, DWORD /*reco*/,
	LPDATAOBJECT lpRichDataObj, LPDATAOBJECT* lplpdataobj)
{
	CHARRANGE& cr = *lpchrg;

	if ((cr.cpMax - cr.cpMin == 1) &&
		GetRichEditCtrl().GetSelectionType() == SEL_OBJECT)
	{
		return E_NOTIMPL;
	}

	BeginWaitCursor();
	//create the data source
	COleDataSource* pDataSource = new COleDataSource;

	// put the formats into the data source
	LPENUMFORMATETC lpEnumFormatEtc;
	lpRichDataObj->EnumFormatEtc(DATADIR_SET, &lpEnumFormatEtc);
	if (lpEnumFormatEtc != NULL)
	{
		FORMATETC etc;
		while (lpEnumFormatEtc->Next(1, &etc, NULL) == S_OK)
		{
			STGMEDIUM stgMedium;
			lpRichDataObj->GetData(&etc, &stgMedium);
			pDataSource->CacheData(etc.cfFormat, &stgMedium, &etc);
		}
		lpEnumFormatEtc->Release();
	}

	CEmbeddedItem item(GetDocument(), cr.cpMin, cr.cpMax);
	item.m_lpRichDataObj = lpRichDataObj;
	// get wordpad formats
	item.GetClipboardData(pDataSource);

	// get the IDataObject from the data source
	*lplpdataobj =  (LPDATAOBJECT)pDataSource->GetInterface(&IID_IDataObject);

	EndWaitCursor();
	return S_OK;
}



HRESULT CWordPadView::QueryAcceptData(
    LPDATAOBJECT lpdataobj,
    CLIPFORMAT* lpcfFormat,
    DWORD reco,
    BOOL bReally,
    HGLOBAL hMetaPict)
{
    if (!lpcfFormat)
        return E_INVALIDARG;

    if (lpdataobj &&
        bReally &&
        *lpcfFormat == 0 &&
        m_nPasteType == 0)
    {
        COleDataObject dataobj;
        dataobj.Attach(lpdataobj, FALSE);

        // Prefer HTML when available. If conversion fails, continue
        // through the existing RTF/native-object/base-class handling.
        if (cfHTML != 0 && dataobj.IsDataAvailable(cfHTML))
        {
            const HRESULT hr = PasteHtmlIntoRichEdit(
                GetSafeHwnd(),
                lpdataobj,
                cfHTML);

            if (SUCCEEDED(hr))
                return S_FALSE;
        }

        // Retain the application's existing embedded-object handling.
        if (!dataobj.IsDataAvailable(cfRTO) &&
            dataobj.IsDataAvailable(cfEmbeddedObject))
        {
            if (PasteNative(lpdataobj))
                return S_FALSE;
        }
    }

    return CRichEditView::QueryAcceptData(
        lpdataobj,
        lpcfFormat,
        reco,
        bReally,
        hMetaPict);
}

BOOL CWordPadView::PasteNative(LPDATAOBJECT lpdataobj)
{
	// check data object for wordpad object
	// if true, suck out RTF directly
	FORMATETC etc = {NULL, NULL, DVASPECT_CONTENT, -1, TYMED_ISTORAGE};
	etc.cfFormat = (CLIPFORMAT)cfEmbeddedObject;
	STGMEDIUM stgMedium = {TYMED_ISTORAGE, 0, NULL};

	// create an IStorage to transfer the data in
	LPLOCKBYTES lpLockBytes;
	if (FAILED(::CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes)))
		return FALSE;
	ASSERT(lpLockBytes != NULL);
	if (lpLockBytes == NULL)
		return FALSE;

	HRESULT hr = ::StgCreateDocfileOnILockBytes(lpLockBytes,
		STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE, 0, &stgMedium.pstg);
	lpLockBytes->Release(); //storage addref'd
	if (FAILED(hr))
		return FALSE;

	ASSERT(stgMedium.pstg != NULL);
	CLSID clsid;
	BOOL bRes = FALSE; //let richedit do what it wants
	if (SUCCEEDED(lpdataobj->GetDataHere(&etc, &stgMedium)) &&
		SUCCEEDED(ReadClassStg(stgMedium.pstg, &clsid)) &&
		clsid == GetDocument()->GetClassID())
	{
		//suck out RTF now
		// open Contents stream
		COleStreamFile file;
		CFileException fe;
		if (file.OpenStream(stgMedium.pstg, szContents,
			CFile::modeReadWrite|CFile::shareExclusive, &fe))
		{

			// load it with CArchive (loads from Contents stream)
			CArchive loadArchive(&file, CArchive::load |
				CArchive::bNoFlushOnDelete);
			Stream(loadArchive, TRUE); //stream in selection
			hr = S_FALSE; // don't let richedit do anything
		}
	}
	::ReleaseStgMedium(&stgMedium);
	return bRes;
}

// things to fix
// if format==0 we are doing a straight EM_PASTE
//  look for native formats
//      richedit specific -- allow richedit to handle (these will be first)
//      look for RTF, CF_TEXT.  If there paste special as these
//  Do standard OLE scenario

// if pasting a particular format (format != 0)
//  if richedit specific, allow through
//  if RTF, CF_TEXT. paste special
//  if OLE format, do standard OLE scenario


void CWordPadView::OnFilePrint()
{
	// don't allow winini changes to occur while printing
	m_bInPrint = TRUE;
	CRichEditView::OnFilePrint();
	// printer may have changed
	theApp.NotifyPrinterChanged(); // this will cause a GetDocument()->PrinterChanged();
	m_bInPrint = FALSE;
}

int CWordPadView::OnMouseActivate(CWnd* pWnd, UINT nHitTest, UINT message)
{
	if (m_bOnBar)
	{
		SetFocus();
		return MA_ACTIVATEANDEAT;
	}
	else
		return CRichEditView::OnMouseActivate(pWnd, nHitTest, message);
}

#ifdef CORRECT_WRITING
typedef BOOL (WINAPI *PCWPROC)(HWND, LPSTR, UINT, LPVOID, DWORD, DWORD);
void CWordPadView::OnPenLens()
{
	USES_CONVERSION;
	HINSTANCE hLib = LoadLibrary(_T("PENWIN32.DLL"));
	if (hLib == NULL)
		return;
	PCWPROC pCorrectWriting = (PCWPROC)GetProcAddress(hLib, "CorrectWriting");
	ASSERT(pCorrectWriting != NULL);
	if (pCorrectWriting != NULL)
	{
		CHARRANGE cr;
		GetRichEditCtrl().GetSel(cr);
		int nCnt = 2*(cr.cpMax-cr.cpMin);
		BOOL bSel = (nCnt != 0);
		nCnt = max(1024, nCnt);
		char* pBuf = new char[nCnt];
		pBuf[0] = NULL;
		if (bSel)
			GetRichEditCtrl().GetSelText(pBuf);
		if (pCorrectWriting(m_hWnd, pBuf, nCnt, 0, bSel ? 0 : CWR_INSERT, 0))
			GetRichEditCtrl().ReplaceSel(A2T(pBuf));
		delete [] pBuf;
	}
	FreeLibrary(hLib);
}
#endif

LONG CWordPadView::OnPrinterChangedMsg(UINT, LONG)
{
	CDC dc;
	AfxGetApp()->CreatePrinterDC(dc);
	OnPrinterChanged(dc);
	return 0;
}

static void ForwardPaletteChanged(HWND hWndParent, HWND hWndFocus)
{
	// this is a quick and dirty hack to send the WM_QUERYNEWPALETTE to a window that is interested
	HWND hWnd = NULL;
	for (hWnd = ::GetWindow(hWndParent, GW_CHILD); hWnd != NULL; hWnd = ::GetWindow(hWnd, GW_HWNDNEXT))
	{
		if (hWnd != hWndFocus)
		{
			::SendMessage(hWnd, WM_PALETTECHANGED, (WPARAM)hWndFocus, 0L);
			ForwardPaletteChanged(hWnd, hWndFocus);
		}
	}
}

void CWordPadView::OnPaletteChanged(CWnd* pFocusWnd)
{
	ForwardPaletteChanged(m_hWnd, pFocusWnd->GetSafeHwnd());
	// allow the richedit control to realize its palette
	// remove this if if richedit fixes their code so that
	// they don't realize their palette into foreground
	if (::GetWindow(m_hWnd, GW_CHILD) == NULL)
		CRichEditView::OnPaletteChanged(pFocusWnd);
}

static BOOL FindQueryPalette(HWND hWndParent)
{
	// this is a quick and dirty hack to send the WM_QUERYNEWPALETTE to a window that is interested
	HWND hWnd = NULL;
	for (hWnd = ::GetWindow(hWndParent, GW_CHILD); hWnd != NULL; hWnd = ::GetWindow(hWnd, GW_HWNDNEXT))
	{
		if (::SendMessage(hWnd, WM_QUERYNEWPALETTE, 0, 0L))
			return TRUE;
		else if (FindQueryPalette(hWnd))
			return TRUE;
	}
	return FALSE;
}

BOOL CWordPadView::OnQueryNewPalette()
{
	if(FindQueryPalette(m_hWnd))
		return TRUE;
	return CRichEditView::OnQueryNewPalette();
}

void CWordPadView::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
	CRichEditView::OnSettingChange(uFlags, lpszSection);
	//printer might have changed
	if (!m_bInPrint)
	{
		if (_tcsicmp(lpszSection, _T("windows")) == 0)
			theApp.NotifyPrinterChanged(TRUE); // force update to defaults
	}
}

void CWordPadView::OnSize(UINT nType, int cx, int cy)
{
	CRichEditView::OnSize(nType, cx, cy);
	CRect rect(HORZ_TEXTOFFSET, VERT_TEXTOFFSET, cx, cy);
	GetRichEditCtrl().SetRect(rect);
}

void CWordPadView::OnBarSetFocus(UINT, NMHDR*, LRESULT*)
{
	m_bOnBar = TRUE;
}

void CWordPadView::OnBarKillFocus(UINT, NMHDR*, LRESULT*)
{
	m_bOnBar = FALSE;
}

void CWordPadView::OnBarReturn(UINT, NMHDR*, LRESULT* )
{
	SetFocus();
}

void CWordPadView::OnContextMenu(CWnd* /*pWnd*/, CPoint point) 
{
	if (!ShowContextMenu (point))
	{
		Default ();
	}
}

void CWordPadView::OnRButtonUp(UINT nFlags, CPoint point) 
{
	long nStartChar, nEndChar;
	GetRichEditCtrl().GetSel(nStartChar, nEndChar);
	if (nEndChar - nStartChar <= 1)
	{
		SendMessage (WM_LBUTTONDOWN, nFlags, MAKELPARAM (point.x, point.y));
		ReleaseCapture ();
	}

	CPoint ptScreen = point;
	ClientToScreen (&ptScreen);

	if (!ShowContextMenu (ptScreen))
	{
		Default ();
	}
}

BOOL CWordPadView::ShowContextMenu (CPoint point)
{
	if (DYNAMIC_DOWNCAST (CFrameWndEx, GetParentFrame ()) == NULL)
	{
		// Maybe, server mode, show the regular menu!
		return FALSE;
	}

	CRichEditCntrItem* pItem = GetSelectedItem();
	if (pItem == NULL || !pItem->IsInPlaceActive())
	{
		theApp.ShowPopupMenu (IDR_TEXT_POPUP, point, this);
		return TRUE;
	}

	return FALSE;
}

void CWordPadView::SetTextColor (COLORREF color)
{
	CRichEditView::OnColorPick (color == -1 ? ::GetSysColor (COLOR_WINDOWTEXT) : color);
}

void CWordPadView::OnCharColor() 
{
	COLORREF color = CMFCColorMenuButton::GetColorByCmdID (ID_CHAR_COLOR);
	CRichEditView::OnColorPick (color == -1 ? ::GetSysColor (COLOR_WINDOWTEXT) : color);
}

void CWordPadView::OnFilePrintPreview() 
{
	CRichEditCntrItem* pItem = GetSelectedItem();
	if (pItem != NULL && pItem->IsInPlaceActive())
	{
		pItem->Deactivate ();
	}

	AFXPrintPreview (this);
}

void CWordPadView::GetDocumentColors (CList<COLORREF,COLORREF>& lstColors)
{
	CRichEditCtrl& wndRichEdit = GetRichEditCtrl();

	// Save current selection:
	CHARRANGE cr;
	wndRichEdit.GetSel(cr);

	wndRichEdit.HideSelection (TRUE, FALSE);

	int nTextLen = min (500, wndRichEdit.GetTextLength ());
		// For demonstration purposes only!

	for (long iSel = 0; iSel < nTextLen; iSel += 2)
	{
		wndRichEdit.SetSel (iSel, iSel);
		if (!(wndRichEdit.GetSelectionType() & (SEL_OBJECT|SEL_MULTIOBJECT)))
		{
			CCharFormat cf;
			wndRichEdit.GetSelectionCharFormat (cf);

			COLORREF color = cf.crTextColor;
			if (lstColors.Find (color) == NULL)
			{
				lstColors.AddTail (color);

				if (lstColors.GetCount () == 10)
				{
					break;
				}
			}
		}
	}

	// Restore selection:
	wndRichEdit.SetSel (cr);
	wndRichEdit.HideSelection (FALSE, FALSE);
}

void CWordPadView::OnFontname() 
{
	USES_CONVERSION;

	CMFCToolBarFontComboBox* pSrcCombo = 
		(CMFCToolBarFontComboBox*) CMFCToolBarComboBoxButton::GetByCmd (IDC_FONTNAME, TRUE);
	if (pSrcCombo == NULL)
	{
		CRichEditView::OnFormatFont ();
		return;
	}

	CCharFormat cf;
	cf.szFaceName[0] = NULL;
	cf.dwMask = CFM_FACE | CFM_CHARSET;

	const CMFCFontInfo* pDesc = pSrcCombo->GetFontDesc ();
	ASSERT_VALID (pDesc);
	ASSERT(pDesc->m_strName.GetLength() < LF_FACESIZE);

	lstrcpyn(cf.szFaceName, pDesc->m_strName, LF_FACESIZE);

	cf.bCharSet = pDesc->m_nCharSet;
	cf.bPitchAndFamily = pDesc->m_nPitchAndFamily;

	CMFCToolBarFontSizeComboBox* pSizeCombo =
		DYNAMIC_DOWNCAST (CMFCToolBarFontSizeComboBox, CMFCToolBarFontSizeComboBox::GetByCmd (IDC_FONTSIZE));
	if (pSizeCombo != NULL)
	{
		int nSize = pSizeCombo->GetTwipSize();
		pSizeCombo->RebuildFontSizes (pDesc->m_strName);
		pSizeCombo->SetTwipSize (nSize);
	}

	SetCharFormat (cf);
	SetFocus ();
}
//****************************************************************************************
void CWordPadView::OnFontsize() 
{
	CMFCToolBarFontSizeComboBox* pSrcCombo = 
		(CMFCToolBarFontSizeComboBox*) CMFCToolBarComboBoxButton::GetByCmd (IDC_FONTSIZE, TRUE);
	if (pSrcCombo == NULL)
	{
		CRichEditView::OnFormatFont ();
		return;
	}

	int nSize = pSrcCombo->GetTwipSize();
	if (nSize == -2)
	{
		AfxMessageBox(IDS_INVALID_NUMBER, MB_OK|MB_ICONINFORMATION);
	}
	else if ((nSize >= 0 && nSize < 20) || nSize > 32760)
	{
		AfxMessageBox(IDS_INVALID_FONTSIZE, MB_OK|MB_ICONINFORMATION);
	}
	else if (nSize > 0)
	{
		CCharFormat cf;
		cf.dwMask = CFM_SIZE;
		cf.yHeight = nSize;

		SetCharFormat (cf);
		SetFocus ();
	}
}
//*********************************************************************************
void CWordPadView::OnBorderType (UINT id)
{
	m_nBorderType = id;

	MessageBox (_T("Add your code here..."));
}
//********************************************************************************
void CWordPadView::OnUpdateBorderType (CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck (pCmdUI->m_nID == m_nBorderType);
}
