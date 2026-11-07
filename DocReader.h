#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace docreader
{
    struct Options
    {
        // Used for compressed (8-bit) text pieces. CP_ACP is the default.
        // Set to e.g. 1251 for a known Windows-1251 source document.
        unsigned int codePage = 0; // 0 means CP_ACP

        // Reject streams larger than this limit.
        std::uint64_t maxStreamBytes = 256ull * 1024 * 1024;

        // Emit a visible marker for Word special characters that this
        // first implementation cannot interpret as an embedded object.
        bool emitObjectPlaceholders = true;
    };

    struct Result
    {
        bool success = false;
        std::wstring error;
        std::vector<std::wstring> warnings;
        std::uint64_t charactersRead = 0;
        std::uint32_t pieceCount = 0;
    };

    // Reads a Word 97-2003 binary .doc file and returns a valid RTF byte stream.
    //
    // Current implementation:
    //   * Native OLE Compound File access via Windows Structured Storage.
    //   * FIB parsing and CLX/PlcPcd piece-table text extraction.
    //   * UTF-16 and compressed 8-bit text pieces.
    //   * Paragraphs, tabs, line breaks, page breaks and cell separators.
    //   * Safe RTF escaping.
    //
    // Not yet implemented:
    //   * CHP/PAP formatting runs, styles and numbering.
    //   * Actual PICF/OfficeArt picture extraction (object markers are represented
    //     by placeholders).
    //   * Full table reconstruction, headers/footers, fields and floating shapes.
    //
    // This deliberately returns a useful text-preserving RTF first, rather than
    // guessing at unsupported binary structures.
    Result ConvertToRtf(
        const std::wstring& fileName,
        std::vector<std::uint8_t>& rtf,
        const Options& options = {});

    // Convenience overload for callers that want a file on disk.
    Result ConvertToRtfFile(
        const std::wstring& inputFile,
        const std::wstring& outputRtfFile,
        const Options& options = {});
}
