#include "stdafx.h"

#include "RichEditStream.h"
#include "DocReader.h"

#ifdef _WIN32
#include <windows.h>
#include <richedit.h>
#include <algorithm>
#include <cstring>
#include <fstream>
#include <vector>

namespace docreader
{
namespace
{
    struct Input
    {
        std::vector<char> bytes;
        std::size_t offset = 0;
    };

    DWORD CALLBACK ReadCallback(DWORD_PTR cookie, LPBYTE buffer,
                                LONG bytes, LONG* transferred)
    {
        auto* input = reinterpret_cast<Input*>(cookie);
        if (!input || !transferred || bytes < 0)
            return 1;

        const std::size_t remaining = input->bytes.size() - input->offset;
        const std::size_t count = (std::min)(remaining, static_cast<std::size_t>(bytes));
        if (count)
        {
            memcpy(buffer, input->bytes.data() + input->offset, count);
            input->offset += count;
        }
        *transferred = static_cast<LONG>(count);
        return 0;
    }

    bool ReadFileBytes(const std::wstring& path, std::vector<char>& bytes)
    {
        std::ifstream file(path, std::ios::binary);
        if (!file)
            return false;
        file.seekg(0, std::ios::end);
        const auto length = file.tellg();
        if (length < 0 || static_cast<unsigned long long>(length) > 512ull * 1024 * 1024)
            return false;
        file.seekg(0, std::ios::beg);
        bytes.resize(static_cast<std::size_t>(length));
        return bytes.empty() || bool(file.read(bytes.data(), static_cast<std::streamsize>(length)));
    }
}

bool StreamRtfFileToRichEdit(HWND richEdit, const std::wstring& rtfFile)
{
    if (!IsWindow(richEdit))
        return false;

    Input input;
    if (!ReadFileBytes(rtfFile, input.bytes))
        return false;

    EDITSTREAM stream{};
    stream.dwCookie = reinterpret_cast<DWORD_PTR>(&input);
    stream.pfnCallback = ReadCallback;

    const LRESULT chars = SendMessageW(richEdit, EM_STREAMIN, SF_RTF,
                                       reinterpret_cast<LPARAM>(&stream));
    return chars >= 0 && stream.dwError == 0;
}

bool LoadDocIntoRichEdit(HWND richEdit, const std::wstring& docFile,
                         std::wstring* error)
{
    std::vector<std::uint8_t> rtf;
    Result result = ConvertToRtf(docFile, rtf);
    if (!result.success)
    {
        if (error) *error = result.error;
        return false;
    }

    if (!IsWindow(richEdit))
    {
        if (error) *error = L"Invalid Rich Edit window handle.";
        return false;
    }

    Input input;
    input.bytes.assign(reinterpret_cast<const char*>(rtf.data()),
                       reinterpret_cast<const char*>(rtf.data() + rtf.size()));

    EDITSTREAM stream{};
    stream.dwCookie = reinterpret_cast<DWORD_PTR>(&input);
    stream.pfnCallback = ReadCallback;

    SendMessageW(richEdit, EM_STREAMIN, SF_RTF,
                 reinterpret_cast<LPARAM>(&stream));

    if (stream.dwError != 0)
    {
        if (error) *error = L"Rich Edit failed to import the generated RTF.";
        return false;
    }
    return true;
}
}
#endif
