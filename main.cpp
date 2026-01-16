/*
 * Font Enumerator - A Windows desktop application for exploring system fonts
 *
 * This application demonstrates three different Windows APIs for font enumeration:
 * 1. GDI (Graphics Device Interface) - Legacy API, available on all Windows versions
 * 2. DirectWrite - Modern API with better Unicode support and font metrics
 * 3. FontSet API - Windows 10+ API with access to variable font axes and file paths
 *
 * Architecture Overview
 * =====================
 * The application follows a typical Win32 GUI structure:
 * - Single main window with child controls (buttons, listview, preview panel)
 * - Global application state stored in global variables
 * - Message-driven event handling through the window procedure (WndProc)
 *
 * Code Organization
 * =================
 * 1. Includes & Pragmas (lines ~45-55)
 * 2. Constants - Control IDs (lines ~57-65)
 * 3. Global Variables - Window handles, state (lines ~67-105)
 * 4. Data Structures - FontInfo, EnumMode (lines ~80-100)
 * 5. Utility Functions - ContainsIgnoreCase (lines ~107-115)
 * 6. Forward Declarations (lines ~117-125)
 * 7. GDI Font Enumeration (lines ~127-185)
 * 8. DirectWrite Font Enumeration (lines ~187-300)
 * 9. FontSet Font Enumeration (lines ~302-490)
 * 10. Font Data Management - ClearFonts, ApplyFilter (lines ~492-520)
 * 11. UI Update Functions - UpdateStatusText, PopulateListView (lines ~522-580)
 * 12. Preview Panel - UpdatePreview, PreviewWndProc (lines ~582-655)
 * 13. UI Creation - CreateControls (lines ~657-770)
 * 14. Layout - ResizeControls (lines ~772-785)
 * 15. Window Procedure - WndProc (lines ~787-865)
 * 16. Entry Point - wWinMain (lines ~867-920)
 */

// ============================================================================
// INCLUDES & LIBRARY LINKAGE
// ============================================================================

#include <windows.h>
#include <commctrl.h>      // Common controls (ListView)
#include <dwrite_3.h>      // DirectWrite 3 for FontSet API
#include <vector>
#include <string>
#include <algorithm>

// Link required libraries
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "dwrite.lib")

// Enable visual styles for modern control appearance
#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// ============================================================================
// CONSTANTS - Control IDs for child windows
// ============================================================================
// These IDs are used to identify controls in WM_COMMAND and WM_NOTIFY messages

#define IDC_LISTVIEW        1001    // Main font list
#define IDC_GDI_BUTTON      1002    // "GDI" enumeration button
#define IDC_DWRITE_BUTTON   1003    // "DirectWrite" enumeration button
#define IDC_FONTSET_BUTTON  1004    // "FontSet API" enumeration button
#define IDC_PREVIEW_STATIC  1005    // Font preview panel
#define IDC_STATUS_LABEL    1006    // Status text showing font count
#define IDC_SEARCH_EDIT     1007    // Filter text input
#define IDC_SEARCH_LABEL    1008    // "Filter:" label

// ============================================================================
// GLOBAL VARIABLES
// ============================================================================

// Window handles
HWND g_hWnd = NULL;              // Main window
HWND g_hListView = NULL;         // ListView control
HWND g_hGdiButton = NULL;        // GDI button
HWND g_hDWriteButton = NULL;     // DirectWrite button
HWND g_hFontSetButton = NULL;    // FontSet button
HWND g_hPreviewStatic = NULL;    // Preview panel
HWND g_hStatusLabel = NULL;      // Status label
HWND g_hSearchEdit = NULL;       // Filter input
HWND g_hSearchLabel = NULL;      // "Filter:" label
HINSTANCE g_hInstance = NULL;    // Application instance
std::wstring g_filterText;       // Current filter string

// ============================================================================
// DATA STRUCTURES
// ============================================================================

// Enumeration mode - tracks which API was used to enumerate fonts
enum class EnumMode {
    None,        // No enumeration performed yet
    GDI,         // EnumFontFamiliesEx (legacy)
    DirectWrite, // IDWriteFontCollection (modern)
    FontSet      // IDWriteFontSet (Windows 10+)
};

EnumMode g_currentMode = EnumMode::None;

/*
 * FontInfo - Represents information about a single font face
 *
 * Different enumeration APIs provide different levels of detail:
 * - GDI: familyName, styleName, weight, italic, fixedPitch, charSet
 * - DirectWrite: Same as GDI plus better Unicode handling
 * - FontSet: All above plus filePath, variableAxes, isVariable
 */
struct FontInfo {
    std::wstring familyName;    // e.g., "Arial", "Segoe UI"
    std::wstring styleName;     // e.g., "Regular", "Bold Italic"
    std::wstring filePath;      // Full path to font file (FontSet API only)
    std::wstring variableAxes;  // Variable font axes, e.g., "wght 100-900" (FontSet API only)
    int weight;                 // Font weight: 400=Normal, 700=Bold, etc.
    bool italic;                // Whether this is an italic/oblique style
    bool fixedPitch;            // True for monospace fonts
    bool isVariable;            // True if font has variable axes
    int charSet;                // Character set (GDI-specific)
};

// Font data storage
std::vector<FontInfo> g_fonts;              // All enumerated fonts
std::vector<size_t> g_filteredIndices;      // Indices of fonts matching filter

// Selected font state (for preview)
std::wstring g_selectedFont;                // Selected font family name
std::wstring g_selectedStyle;               // Selected font style name
int g_selectedWeight = FW_NORMAL;           // Selected font weight
bool g_selectedItalic = false;              // Selected font italic flag

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/*
 * Case-insensitive substring search
 * Used for filtering fonts by name
 */
bool ContainsIgnoreCase(const std::wstring& str, const std::wstring& substr) {
    if (substr.empty()) return true;
    std::wstring strLower = str;
    std::wstring substrLower = substr;
    std::transform(strLower.begin(), strLower.end(), strLower.begin(), ::towlower);
    std::transform(substrLower.begin(), substrLower.end(), substrLower.begin(), ::towlower);
    return strLower.find(substrLower) != std::wstring::npos;
}

// ============================================================================
// FORWARD DECLARATIONS
// ============================================================================

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void EnumerateGDIFonts();
void EnumerateDirectWriteFonts();
void EnumerateFontSetFonts();
void PopulateListView();
void ApplyFilter();
void UpdateStatusText();
void UpdatePreview();
void ClearFonts();

// ============================================================================
// FONT ENUMERATION - GDI API
// ============================================================================

/*
 * Callback function for GDI font enumeration
 *
 * Called once for each font face found by EnumFontFamiliesExW.
 * Extracts font information and adds unique fonts to the collection.
 */
int CALLBACK EnumFontFamExProc(
    const LOGFONTW* lpelfe,
    const TEXTMETRICW* lpntme,
    DWORD FontType,
    LPARAM lParam)
{
    // Cast to extended structure for style name access
    const ENUMLOGFONTEXW* elfex = reinterpret_cast<const ENUMLOGFONTEXW*>(lpelfe);

    FontInfo info;
    info.familyName = lpelfe->lfFaceName;
    info.styleName = elfex->elfStyle;
    info.weight = lpelfe->lfWeight;
    info.italic = lpelfe->lfItalic != 0;
    info.fixedPitch = (lpelfe->lfPitchAndFamily & FIXED_PITCH) != 0;
    info.charSet = lpelfe->lfCharSet;

    // Skip duplicates (same family + style combination)
    bool exists = false;
    for (const auto& f : g_fonts) {
        if (f.familyName == info.familyName && f.styleName == info.styleName) {
            exists = true;
            break;
        }
    }

    if (!exists) {
        g_fonts.push_back(info);
    }

    return 1; // Return 1 to continue enumeration
}

/*
 * Enumerates fonts using the GDI EnumFontFamiliesEx API
 *
 * This is the oldest font enumeration API, available on all Windows versions.
 * Limitations:
 * - No access to font file paths
 * - No variable font axis information
 * - Limited style name accuracy for some fonts
 */
void EnumerateGDIFonts()
{
    ClearFonts();

    HDC hdc = GetDC(g_hWnd);

    // Set up LOGFONT to enumerate all fonts
    LOGFONTW lf = {};
    lf.lfCharSet = DEFAULT_CHARSET;  // Enumerate all character sets
    lf.lfFaceName[0] = L'\0';        // Enumerate all font families
    lf.lfPitchAndFamily = 0;

    // Enumerate all font families
    EnumFontFamiliesExW(hdc, &lf, EnumFontFamExProc, 0, 0);

    ReleaseDC(g_hWnd, hdc);

    // Sort alphabetically by family name
    std::sort(g_fonts.begin(), g_fonts.end(),
        [](const FontInfo& a, const FontInfo& b) {
            return a.familyName < b.familyName;
        });

    g_currentMode = EnumMode::GDI;
    ApplyFilter();
}

// ============================================================================
// FONT ENUMERATION - DirectWrite API
// ============================================================================

/*
 * Enumerates fonts using the DirectWrite IDWriteFontCollection API
 *
 * DirectWrite provides better support for:
 * - OpenType features
 * - Complex script shaping
 * - Font fallback
 * - Accurate style names
 *
 * Available on Windows Vista and later.
 */
void EnumerateDirectWriteFonts()
{
    ClearFonts();

    // Create DirectWrite factory
    IDWriteFactory* pDWriteFactory = nullptr;
    IDWriteFontCollection* pFontCollection = nullptr;

    HRESULT hr = DWriteCreateFactory(
        DWRITE_FACTORY_TYPE_SHARED,
        __uuidof(IDWriteFactory),
        reinterpret_cast<IUnknown**>(&pDWriteFactory));

    if (FAILED(hr)) {
        MessageBoxW(g_hWnd, L"Failed to create DirectWrite factory", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Get the system font collection
    hr = pDWriteFactory->GetSystemFontCollection(&pFontCollection, FALSE);
    if (FAILED(hr)) {
        pDWriteFactory->Release();
        MessageBoxW(g_hWnd, L"Failed to get system font collection", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    UINT32 familyCount = pFontCollection->GetFontFamilyCount();

    // Iterate through each font family
    for (UINT32 i = 0; i < familyCount; i++) {
        IDWriteFontFamily* pFontFamily = nullptr;
        hr = pFontCollection->GetFontFamily(i, &pFontFamily);
        if (FAILED(hr)) continue;

        // Get family name (prefer English)
        IDWriteLocalizedStrings* pFamilyNames = nullptr;
        hr = pFontFamily->GetFamilyNames(&pFamilyNames);
        if (SUCCEEDED(hr)) {
            UINT32 index = 0;
            BOOL exists = FALSE;

            // Try to find English name first
            pFamilyNames->FindLocaleName(L"en-us", &index, &exists);
            if (!exists) {
                index = 0;
            }

            UINT32 length = 0;
            pFamilyNames->GetStringLength(index, &length);

            std::wstring familyName(length + 1, L'\0');
            pFamilyNames->GetString(index, &familyName[0], length + 1);
            familyName.resize(length);

            // Each family can contain multiple fonts (Regular, Bold, Italic, etc.)
            UINT32 fontCount = pFontFamily->GetFontCount();
            for (UINT32 j = 0; j < fontCount; j++) {
                IDWriteFont* pFont = nullptr;
                hr = pFontFamily->GetFont(j, &pFont);
                if (FAILED(hr)) continue;

                // Get face/style name
                IDWriteLocalizedStrings* pFaceNames = nullptr;
                hr = pFont->GetFaceNames(&pFaceNames);

                std::wstring styleName;
                if (SUCCEEDED(hr)) {
                    UINT32 faceIndex = 0;
                    BOOL faceExists = FALSE;
                    pFaceNames->FindLocaleName(L"en-us", &faceIndex, &faceExists);
                    if (!faceExists) faceIndex = 0;

                    UINT32 faceLength = 0;
                    pFaceNames->GetStringLength(faceIndex, &faceLength);
                    styleName.resize(faceLength + 1);
                    pFaceNames->GetString(faceIndex, &styleName[0], faceLength + 1);
                    styleName.resize(faceLength);

                    pFaceNames->Release();
                }

                FontInfo info;
                info.familyName = familyName;
                info.styleName = styleName;
                info.weight = pFont->GetWeight();
                info.italic = (pFont->GetStyle() == DWRITE_FONT_STYLE_ITALIC ||
                              pFont->GetStyle() == DWRITE_FONT_STYLE_OBLIQUE);

                // Check if font is monospaced (requires IDWriteFont1)
                info.fixedPitch = false;
                IDWriteFont1* pFont1 = nullptr;
                if (SUCCEEDED(pFont->QueryInterface(__uuidof(IDWriteFont1), (void**)&pFont1))) {
                    info.fixedPitch = pFont1->IsMonospacedFont() == TRUE;
                    pFont1->Release();
                }
                info.charSet = DEFAULT_CHARSET;

                g_fonts.push_back(info);

                pFont->Release();
            }

            pFamilyNames->Release();
        }

        pFontFamily->Release();
    }

    pFontCollection->Release();
    pDWriteFactory->Release();

    // Sort by family name, then by style name
    std::sort(g_fonts.begin(), g_fonts.end(),
        [](const FontInfo& a, const FontInfo& b) {
            if (a.familyName != b.familyName)
                return a.familyName < b.familyName;
            return a.styleName < b.styleName;
        });

    g_currentMode = EnumMode::DirectWrite;
    ApplyFilter();
}

// ============================================================================
// FONT ENUMERATION - FontSet API (Windows 10+)
// ============================================================================

/*
 * Enumerates fonts using the DirectWrite IDWriteFontSet API
 *
 * The FontSet API (Windows 10+) provides access to:
 * - Font file paths
 * - Variable font axis information (weight ranges, width ranges, etc.)
 * - More detailed font properties
 *
 * This is the most comprehensive font enumeration API available.
 */
void EnumerateFontSetFonts()
{
    ClearFonts();

    // Create DirectWrite factory (version 3 required for FontSet API)
    IDWriteFactory3* pDWriteFactory3 = nullptr;

    HRESULT hr = DWriteCreateFactory(
        DWRITE_FACTORY_TYPE_SHARED,
        __uuidof(IDWriteFactory3),
        reinterpret_cast<IUnknown**>(&pDWriteFactory3));

    if (FAILED(hr)) {
        MessageBoxW(g_hWnd, L"Failed to create DirectWrite factory 3.\nThis feature requires Windows 10 or later.",
            L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Get the system font set
    IDWriteFontSet* pFontSet = nullptr;
    hr = pDWriteFactory3->GetSystemFontSet(&pFontSet);
    if (FAILED(hr)) {
        pDWriteFactory3->Release();
        MessageBoxW(g_hWnd, L"Failed to get system font set", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    UINT32 fontCount = pFontSet->GetFontCount();

    // Iterate through each font in the set
    for (UINT32 i = 0; i < fontCount; i++) {
        IDWriteFontFaceReference* pFontFaceRef = nullptr;
        hr = pFontSet->GetFontFaceReference(i, &pFontFaceRef);
        if (FAILED(hr)) continue;

        FontInfo info;

        // --- Extract font file path ---
        IDWriteFontFile* pFontFile = nullptr;
        if (SUCCEEDED(pFontFaceRef->GetFontFile(&pFontFile))) {
            IDWriteFontFileLoader* pLoader = nullptr;
            if (SUCCEEDED(pFontFile->GetLoader(&pLoader))) {
                // Only local fonts have file paths
                IDWriteLocalFontFileLoader* pLocalLoader = nullptr;
                if (SUCCEEDED(pLoader->QueryInterface(__uuidof(IDWriteLocalFontFileLoader), (void**)&pLocalLoader))) {
                    const void* refKey = nullptr;
                    UINT32 refKeySize = 0;
                    if (SUCCEEDED(pFontFile->GetReferenceKey(&refKey, &refKeySize))) {
                        UINT32 pathLen = 0;
                        if (SUCCEEDED(pLocalLoader->GetFilePathLengthFromKey(refKey, refKeySize, &pathLen))) {
                            info.filePath.resize(pathLen + 1);
                            if (SUCCEEDED(pLocalLoader->GetFilePathFromKey(refKey, refKeySize, &info.filePath[0], pathLen + 1))) {
                                info.filePath.resize(pathLen);
                            }
                        }
                    }
                    pLocalLoader->Release();
                }
                pLoader->Release();
            }
            pFontFile->Release();
        }

        // --- Extract font properties from the font set ---

        // Get family name
        IDWriteLocalizedStrings* pFamilyNames = nullptr;
        BOOL exists = FALSE;
        hr = pFontSet->GetPropertyValues(i, DWRITE_FONT_PROPERTY_ID_FAMILY_NAME, &exists, &pFamilyNames);
        if (SUCCEEDED(hr) && exists && pFamilyNames) {
            UINT32 index = 0;
            BOOL found = FALSE;
            pFamilyNames->FindLocaleName(L"en-us", &index, &found);
            if (!found) index = 0;

            UINT32 length = 0;
            pFamilyNames->GetStringLength(index, &length);
            info.familyName.resize(length + 1);
            pFamilyNames->GetString(index, &info.familyName[0], length + 1);
            info.familyName.resize(length);
            pFamilyNames->Release();
        }

        // Get style/face name
        IDWriteLocalizedStrings* pFaceNames = nullptr;
        hr = pFontSet->GetPropertyValues(i, DWRITE_FONT_PROPERTY_ID_FACE_NAME, &exists, &pFaceNames);
        if (SUCCEEDED(hr) && exists && pFaceNames) {
            UINT32 index = 0;
            BOOL found = FALSE;
            pFaceNames->FindLocaleName(L"en-us", &index, &found);
            if (!found) index = 0;

            UINT32 length = 0;
            pFaceNames->GetStringLength(index, &length);
            info.styleName.resize(length + 1);
            pFaceNames->GetString(index, &info.styleName[0], length + 1);
            info.styleName.resize(length);
            pFaceNames->Release();
        }

        // Get weight
        IDWriteLocalizedStrings* pWeightStr = nullptr;
        hr = pFontSet->GetPropertyValues(i, DWRITE_FONT_PROPERTY_ID_WEIGHT, &exists, &pWeightStr);
        if (SUCCEEDED(hr) && exists && pWeightStr) {
            wchar_t weightBuf[32] = {};
            pWeightStr->GetString(0, weightBuf, 32);
            info.weight = _wtoi(weightBuf);
            pWeightStr->Release();
        } else {
            info.weight = 400;  // Default to normal weight
        }

        // Get style (italic/oblique)
        IDWriteLocalizedStrings* pStyleStr = nullptr;
        hr = pFontSet->GetPropertyValues(i, DWRITE_FONT_PROPERTY_ID_STYLE, &exists, &pStyleStr);
        if (SUCCEEDED(hr) && exists && pStyleStr) {
            wchar_t styleBuf[32] = {};
            pStyleStr->GetString(0, styleBuf, 32);
            int style = _wtoi(styleBuf);
            info.italic = (style == DWRITE_FONT_STYLE_ITALIC || style == DWRITE_FONT_STYLE_OBLIQUE);
            pStyleStr->Release();
        } else {
            info.italic = false;
        }

        info.fixedPitch = false;  // FontSet doesn't directly expose this
        info.charSet = DEFAULT_CHARSET;
        info.isVariable = false;

        // --- Extract variable font axis information ---
        // Requires creating a font face and querying IDWriteFontFace5
        IDWriteFontFace3* pFontFace3 = nullptr;
        if (SUCCEEDED(pFontFaceRef->CreateFontFace(&pFontFace3))) {
            IDWriteFontFace5* pFontFace5 = nullptr;
            if (SUCCEEDED(pFontFace3->QueryInterface(__uuidof(IDWriteFontFace5), (void**)&pFontFace5))) {
                IDWriteFontResource* pFontResource = nullptr;
                if (SUCCEEDED(pFontFace5->GetFontResource(&pFontResource))) {
                    UINT32 axisCount = pFontResource->GetFontAxisCount();
                    if (axisCount > 0) {
                        std::vector<DWRITE_FONT_AXIS_RANGE> axisRanges(axisCount);
                        if (SUCCEEDED(pFontResource->GetFontAxisRanges(axisRanges.data(), axisCount))) {
                            // Check if any axis has a range (min != max means it's variable)
                            for (UINT32 a = 0; a < axisCount; a++) {
                                if (axisRanges[a].minValue != axisRanges[a].maxValue) {
                                    info.isVariable = true;

                                    // Build axis description string
                                    if (!info.variableAxes.empty()) {
                                        info.variableAxes += L", ";
                                    }

                                    // Convert 4-byte axis tag to string (e.g., "wght", "wdth")
                                    DWRITE_FONT_AXIS_TAG tag = axisRanges[a].axisTag;
                                    wchar_t tagStr[5] = {
                                        (wchar_t)(tag & 0xFF),
                                        (wchar_t)((tag >> 8) & 0xFF),
                                        (wchar_t)((tag >> 16) & 0xFF),
                                        (wchar_t)((tag >> 24) & 0xFF),
                                        0
                                    };

                                    wchar_t axisBuf[64];
                                    swprintf_s(axisBuf, L"%s %.0f-%.0f", tagStr,
                                        axisRanges[a].minValue, axisRanges[a].maxValue);
                                    info.variableAxes += axisBuf;
                                }
                            }
                        }
                    }
                    pFontResource->Release();
                }
                pFontFace5->Release();
            }
            pFontFace3->Release();
        }

        if (!info.familyName.empty()) {
            g_fonts.push_back(info);
        }

        pFontFaceRef->Release();
    }

    pFontSet->Release();
    pDWriteFactory3->Release();

    // Sort by family name, then by style name
    std::sort(g_fonts.begin(), g_fonts.end(),
        [](const FontInfo& a, const FontInfo& b) {
            if (a.familyName != b.familyName)
                return a.familyName < b.familyName;
            return a.styleName < b.styleName;
        });

    g_currentMode = EnumMode::FontSet;
    ApplyFilter();
}

// ============================================================================
// FONT DATA MANAGEMENT
// ============================================================================

/*
 * Clears all font data and resets selection state
 * Called before each new enumeration
 */
void ClearFonts()
{
    g_fonts.clear();
    g_filteredIndices.clear();
    ListView_DeleteAllItems(g_hListView);
    g_selectedFont.clear();
    g_selectedStyle.clear();
    g_selectedWeight = FW_NORMAL;
    g_selectedItalic = false;
}

/*
 * Applies the current filter text to the font list
 *
 * Creates a list of indices into g_fonts for fonts that match
 * the filter (case-insensitive search in family name or style name).
 */
void ApplyFilter()
{
    g_filteredIndices.clear();

    for (size_t i = 0; i < g_fonts.size(); i++) {
        const auto& font = g_fonts[i];
        if (ContainsIgnoreCase(font.familyName, g_filterText) ||
            ContainsIgnoreCase(font.styleName, g_filterText)) {
            g_filteredIndices.push_back(i);
        }
    }

    PopulateListView();
    UpdateStatusText();
}

// ============================================================================
// UI UPDATE FUNCTIONS
// ============================================================================

/*
 * Updates the status label with current font count
 */
void UpdateStatusText()
{
    wchar_t status[256];
    const wchar_t* modeStr;
    switch (g_currentMode) {
        case EnumMode::GDI: modeStr = L"GDI"; break;
        case EnumMode::DirectWrite: modeStr = L"DirectWrite"; break;
        case EnumMode::FontSet: modeStr = L"FontSet"; break;
        default: modeStr = L"No"; break;
    }

    if (g_filterText.empty()) {
        swprintf_s(status, L"%s Enumeration: Found %zu fonts", modeStr, g_fonts.size());
    } else {
        swprintf_s(status, L"%s Enumeration: Showing %zu of %zu fonts",
            modeStr, g_filteredIndices.size(), g_fonts.size());
    }
    SetWindowTextW(g_hStatusLabel, status);
}

/*
 * Populates the ListView with filtered font data
 */
void PopulateListView()
{
    ListView_DeleteAllItems(g_hListView);

    for (size_t i = 0; i < g_filteredIndices.size(); i++) {
        const auto& font = g_fonts[g_filteredIndices[i]];

        // Insert main item (family name)
        LVITEMW item = {};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = static_cast<int>(i);
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(font.familyName.c_str());
        item.lParam = static_cast<LPARAM>(g_filteredIndices[i]);  // Store original index
        ListView_InsertItem(g_hListView, &item);

        // Set subitem columns
        ListView_SetItemText(g_hListView, static_cast<int>(i), 1,
            const_cast<LPWSTR>(font.styleName.c_str()));

        wchar_t weightStr[32];
        swprintf_s(weightStr, L"%d", font.weight);
        ListView_SetItemText(g_hListView, static_cast<int>(i), 2, weightStr);

        ListView_SetItemText(g_hListView, static_cast<int>(i), 3,
            font.italic ? const_cast<LPWSTR>(L"Yes") : const_cast<LPWSTR>(L"No"));

        ListView_SetItemText(g_hListView, static_cast<int>(i), 4,
            font.fixedPitch ? const_cast<LPWSTR>(L"Yes") : const_cast<LPWSTR>(L"No"));

        ListView_SetItemText(g_hListView, static_cast<int>(i), 5,
            const_cast<LPWSTR>(font.filePath.c_str()));

        // Variable font info - show "Yes" with axes or empty
        std::wstring varStr = font.isVariable ? (L"Yes: " + font.variableAxes) : L"";
        ListView_SetItemText(g_hListView, static_cast<int>(i), 6,
            const_cast<LPWSTR>(varStr.c_str()));
    }
}

// ============================================================================
// PREVIEW PANEL
// ============================================================================

/*
 * Triggers a repaint of the preview panel
 */
void UpdatePreview()
{
    InvalidateRect(g_hPreviewStatic, NULL, TRUE);
}

/*
 * Subclassed window procedure for the preview panel
 *
 * Handles custom painting to display the selected font with its
 * actual weight and italic style.
 */
LRESULT CALLBACK PreviewWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (message) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);

        RECT rect;
        GetClientRect(hWnd, &rect);

        // Fill background with white
        FillRect(hdc, &rect, (HBRUSH)GetStockObject(WHITE_BRUSH));

        // Draw border
        FrameRect(hdc, &rect, (HBRUSH)GetStockObject(GRAY_BRUSH));

        if (!g_selectedFont.empty()) {
            // Create font for preview with actual weight and italic
            HFONT hFont = CreateFontW(
                32, 0, 0, 0,                    // Height, width, escapement, orientation
                g_selectedWeight,               // Use actual weight (400, 700, etc.)
                g_selectedItalic ? TRUE : FALSE, // Use actual italic flag
                FALSE, FALSE,                   // No underline/strikeout
                DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS,
                CLIP_DEFAULT_PRECIS,
                CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE,
                g_selectedFont.c_str());

            if (hFont) {
                HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);

                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, RGB(0, 0, 0));

                // Preview text shows font name, style, and sample characters
                std::wstring previewText = g_selectedFont + L" " + g_selectedStyle + L"\r\nAaBbCcDdEeFfGgHhIiJjKk\r\n0123456789 !@#$%";

                RECT textRect = rect;
                textRect.left += 10;
                textRect.top += 10;
                textRect.right -= 10;
                textRect.bottom -= 10;

                DrawTextW(hdc, previewText.c_str(), -1, &textRect,
                    DT_LEFT | DT_TOP | DT_WORDBREAK);

                SelectObject(hdc, hOldFont);
                DeleteObject(hFont);
            }
        } else {
            // Show placeholder text when no font is selected
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, RGB(128, 128, 128));
            DrawTextW(hdc, L"Select a font to preview", -1, &rect,
                DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }

        EndPaint(hWnd, &ps);
        return 0;
    }
    case WM_NCDESTROY:
        // Clean up subclass when window is destroyed
        RemoveWindowSubclass(hWnd, PreviewWndProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hWnd, message, wParam, lParam);
}

// ============================================================================
// UI CREATION
// ============================================================================

/*
 * Creates all child controls for the main window
 *
 * Layout:
 * +------------------------------------------------------------------+
 * | [GDI] [DirectWrite] [FontSet API]  Filter: [____]  Status text   |
 * +--------------------------------+--------------------------------+
 * |                                |                                 |
 * |         ListView               |        Preview Panel            |
 * |     (font list table)          |    (sample text in font)        |
 * |                                |                                 |
 * +--------------------------------+---------------------------------+
 */
void CreateControls(HWND hWnd)
{
    // --- Toolbar buttons ---
    g_hGdiButton = CreateWindowW(
        L"BUTTON", L"GDI",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        10, 10, 80, 30,
        hWnd, (HMENU)IDC_GDI_BUTTON, g_hInstance, NULL);

    g_hDWriteButton = CreateWindowW(
        L"BUTTON", L"DirectWrite",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        100, 10, 100, 30,
        hWnd, (HMENU)IDC_DWRITE_BUTTON, g_hInstance, NULL);

    g_hFontSetButton = CreateWindowW(
        L"BUTTON", L"FontSet API",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        210, 10, 100, 30,
        hWnd, (HMENU)IDC_FONTSET_BUTTON, g_hInstance, NULL);

    // --- Filter controls ---
    g_hSearchLabel = CreateWindowW(
        L"STATIC", L"Filter:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        330, 17, 40, 20,
        hWnd, (HMENU)IDC_SEARCH_LABEL, g_hInstance, NULL);

    g_hSearchEdit = CreateWindowExW(
        WS_EX_CLIENTEDGE,  // Sunken edge style
        L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
        375, 12, 180, 24,
        hWnd, (HMENU)IDC_SEARCH_EDIT, g_hInstance, NULL);

    // --- Status label ---
    g_hStatusLabel = CreateWindowW(
        L"STATIC", L"Click a button to enumerate fonts",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        570, 17, 350, 20,
        hWnd, (HMENU)IDC_STATUS_LABEL, g_hInstance, NULL);

    // --- ListView (font list) ---
    g_hListView = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEWW, L"",
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
        10, 50, 600, 400,
        hWnd, (HMENU)IDC_LISTVIEW, g_hInstance, NULL);

    // Enable modern ListView features
    ListView_SetExtendedListViewStyle(g_hListView,
        LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

    // Add columns to ListView
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

    col.pszText = const_cast<LPWSTR>(L"Font Family");
    col.cx = 200;
    col.iSubItem = 0;
    ListView_InsertColumn(g_hListView, 0, &col);

    col.pszText = const_cast<LPWSTR>(L"Style");
    col.cx = 120;
    col.iSubItem = 1;
    ListView_InsertColumn(g_hListView, 1, &col);

    col.pszText = const_cast<LPWSTR>(L"Weight");
    col.cx = 80;
    col.iSubItem = 2;
    ListView_InsertColumn(g_hListView, 2, &col);

    col.pszText = const_cast<LPWSTR>(L"Italic");
    col.cx = 60;
    col.iSubItem = 3;
    ListView_InsertColumn(g_hListView, 3, &col);

    col.pszText = const_cast<LPWSTR>(L"Fixed Pitch");
    col.cx = 80;
    col.iSubItem = 4;
    ListView_InsertColumn(g_hListView, 4, &col);

    col.pszText = const_cast<LPWSTR>(L"File Path");
    col.cx = 200;
    col.iSubItem = 5;
    ListView_InsertColumn(g_hListView, 5, &col);

    col.pszText = const_cast<LPWSTR>(L"Variable Axes");
    col.cx = 200;
    col.iSubItem = 6;
    ListView_InsertColumn(g_hListView, 6, &col);

    // --- Preview panel ---
    // Using owner-draw STATIC control for custom painting
    g_hPreviewStatic = CreateWindowW(
        L"STATIC", L"",
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
        620, 50, 350, 400,
        hWnd, (HMENU)IDC_PREVIEW_STATIC, g_hInstance, NULL);

    // Subclass the preview panel to handle custom painting
    SetWindowSubclass(g_hPreviewStatic, PreviewWndProc, 0, 0);

    // --- Set default GUI font on all controls ---
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SendMessage(g_hGdiButton, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hDWriteButton, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hFontSetButton, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hSearchLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hSearchEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hStatusLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hListView, WM_SETFONT, (WPARAM)hFont, TRUE);
}

// ============================================================================
// LAYOUT
// ============================================================================

/*
 * Resizes child controls when the window size changes
 *
 * The layout splits the content area 2/3 for list, 1/3 for preview panel.
 */
void ResizeControls(HWND hWnd)
{
    RECT rect;
    GetClientRect(hWnd, &rect);

    int width = rect.right - rect.left;
    int height = rect.bottom - rect.top;

    int listWidth = (width - 30) * 2 / 3;
    int previewWidth = width - listWidth - 30;
    int listHeight = height - 70;  // Leave space for toolbar

    MoveWindow(g_hListView, 10, 50, listWidth, listHeight, TRUE);
    MoveWindow(g_hPreviewStatic, listWidth + 20, 50, previewWidth, listHeight, TRUE);
}

// ============================================================================
// WINDOW PROCEDURE - Main message handler
// ============================================================================

/*
 * Handles all window messages for the main window
 *
 * Key messages handled:
 * - WM_CREATE: Initialize child controls
 * - WM_SIZE: Resize controls to fit window
 * - WM_COMMAND: Button clicks and edit control changes
 * - WM_NOTIFY: ListView selection changes
 * - WM_GETMINMAXINFO: Set minimum window size
 * - WM_DESTROY: Clean up and exit
 */
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message) {
    case WM_CREATE:
        CreateControls(hWnd);
        break;

    case WM_SIZE:
        ResizeControls(hWnd);
        break;

    // Handle button clicks and edit control notifications
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_GDI_BUTTON:
            EnumerateGDIFonts();
            break;
        case IDC_DWRITE_BUTTON:
            EnumerateDirectWriteFonts();
            break;
        case IDC_FONTSET_BUTTON:
            EnumerateFontSetFonts();
            break;
        case IDC_SEARCH_EDIT:
            // Filter text changed - reapply filter
            if (HIWORD(wParam) == EN_CHANGE) {
                wchar_t buffer[256] = {};
                GetWindowTextW(g_hSearchEdit, buffer, 256);
                g_filterText = buffer;
                ApplyFilter();
            }
            break;
        }
        break;

    // Handle ListView notifications (selection changes)
    case WM_NOTIFY:
    {
        LPNMHDR pnmh = (LPNMHDR)lParam;
        if (pnmh->idFrom == IDC_LISTVIEW) {
            if (pnmh->code == LVN_ITEMCHANGED) {
                LPNMLISTVIEW pnmlv = (LPNMLISTVIEW)lParam;
                // Only respond to selection (not deselection)
                if (pnmlv->uNewState & LVIS_SELECTED) {
                    // Get the original font index from lParam
                    LVITEMW item = {};
                    item.mask = LVIF_PARAM;
                    item.iItem = pnmlv->iItem;
                    if (ListView_GetItem(g_hListView, &item)) {
                        size_t fontIndex = static_cast<size_t>(item.lParam);
                        if (fontIndex < g_fonts.size()) {
                            // Update selected font state
                            const auto& font = g_fonts[fontIndex];
                            g_selectedFont = font.familyName;
                            g_selectedStyle = font.styleName;
                            g_selectedWeight = font.weight;
                            g_selectedItalic = font.italic;
                            UpdatePreview();
                        }
                    }
                }
            }
        }
        break;
    }

    // Set minimum window size
    case WM_GETMINMAXINFO:
    {
        LPMINMAXINFO lpMMI = (LPMINMAXINFO)lParam;
        lpMMI->ptMinTrackSize.x = 800;
        lpMMI->ptMinTrackSize.y = 500;
        break;
    }

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

// ============================================================================
// ENTRY POINT
// ============================================================================

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    g_hInstance = hInstance;

    // Initialize common controls (required for ListView)
    INITCOMMONCONTROLSEX icex = {};
    icex.dwSize = sizeof(icex);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    // Register the main window class
    WNDCLASSEXW wcex = {};
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW;  // Redraw on size change
    wcex.lpfnWndProc = WndProc;            // Message handler
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = L"FontEnumWindowClass";
    wcex.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcex.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassExW(&wcex)) {
        MessageBoxW(NULL, L"Failed to register window class", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    // Create the main window
    g_hWnd = CreateWindowExW(
        0,
        L"FontEnumWindowClass",
        L"Font Enumerator - GDI, DirectWrite & FontSet API",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,  // Default position
        1100, 650,                      // Initial size
        NULL, NULL, hInstance, NULL);

    if (!g_hWnd) {
        MessageBoxW(NULL, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Standard Win32 message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}
