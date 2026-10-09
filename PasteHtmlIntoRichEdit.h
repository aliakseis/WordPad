#pragma once

HRESULT PasteHtmlIntoRichEdit(
    HWND hwndRichEdit,
    LPDATAOBJECT dataObject,
    CLIPFORMAT htmlFormat);
