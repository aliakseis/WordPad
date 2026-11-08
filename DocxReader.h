#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace docxreader
{
    struct Options
    {
        std::uint64_t maxDocumentBytes = 128ull * 1024 * 1024;
        std::uint64_t maxImageBytes = 32ull * 1024 * 1024;
        bool importImages = true;
        bool importHyperlinks = true;
    };

    struct Result
    {
        bool success = false;
        std::wstring error;
        std::vector<std::wstring> warnings;
        std::uint32_t imagesRead = 0;
        std::uint32_t hyperlinksRead = 0;
    };

    // Converts a DOCX package to RTF for loading by the Rich Edit control.
    // Extraction uses the Windows ZIP-folder shell namespace; XML is parsed by MSXML 6.
    Result ConvertToRtf(const std::wstring& fileName,
                        std::vector<std::uint8_t>& rtf,
                        const Options& options = {});
}
