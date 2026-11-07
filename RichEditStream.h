#pragma once

#include <string>

#ifdef _WIN32
#include <windows.h>
#include <richedit.h>

namespace docreader
{
    // Streams an existing RTF file into a Rich Edit window.
    // The control must be created with a Rich Edit class (e.g. RICHEDIT50W).
    bool StreamRtfFileToRichEdit(HWND richEdit, const std::wstring& rtfFile);

    // Converts a DOC file and streams the resulting RTF into Rich Edit.
    // Returns false on conversion or streaming failure; error receives details.
    bool LoadDocIntoRichEdit(
        HWND richEdit,
        const std::wstring& docFile,
        std::wstring* error = nullptr);
}
#endif
