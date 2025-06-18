#include <windows.h>
#include <commdlg.h>
#include <stdio.h>
#include <commctrl.h>
#include <wctype.h>
#include <Richedit.h>
#include <RichOle.h>
#include <vector>
#include <string>
#include <chrono>
#include <sstream>
#include <wchar.h> 
#include <algorithm>
using std::min;
#include <fstream>
#include <sstream>
#include <map>
#include <vector>
#include <string>
#include <locale>
#include <codecvt>

#pragma comment(lib, "Msftedit.lib")

#define ID_FILE_SAVE 3
#define ID_EDIT_UNDO 3001
#define ID_EDIT_FONT 2001

#ifndef MSFTEDIT_CLASS
#define MSFTEDIT_CLASS L"RICHEDIT50W"
#endif
#ifndef IMF_SPELLCHECKING
#define IMF_SPELLCHECKING 0x00000800
#endif
#ifndef IMF_AUTOCORRECT
#define IMF_AUTOCORRECT   0x00000400
#endif
#ifndef EM_BEGINUNDOACTION
#define EM_BEGINUNDOACTION 0x04C8
#endif
#ifndef EM_ENDUNDOACTION
#define EM_ENDUNDOACTION 0x04C9
#endif
#ifndef EM_EMPTYUNDOBUFFER
#define EM_EMPTYUNDOBUFFER 0x0455
#endif
#define ID_PREVIEW 6

#pragma comment(lib, "comdlg32.lib")

// Global declarations

std::map<std::wstring, std::vector<std::wstring>> rhymeDict;
HBITMAP hPreviewBitmap; 
HWND hStatusBar;
HWND hEdit;
bool isModified = false;
HFONT hFont = nullptr;
LOGFONT lf = {};
void SaveAsPDF(HWND hwndParent, HWND hEdit);
HWND hSavePDFButton;  // Global handle for the button

void LoadRhymeDictionary(const std::string& filename) {
    std::ifstream file(filename);  // Use narrow string
    file.imbue(std::locale("en_US.UTF-8")); // Optional if UTF-8 content

    if (!file.is_open()) {
        MessageBoxA(NULL, "Failed to open rhyme_dictionary.txt", "Error", MB_OK | MB_ICONERROR);
        return;
    }

    std::string line;
    while (std::getline(file, line)) {
        size_t colon = line.find(':');
        if (colon == std::string::npos) continue;

        std::string key = line.substr(0, colon);
        std::string rest = line.substr(colon + 1);

        std::stringstream ss(rest);
        std::string rhyme;
        std::vector<std::wstring> rhymes;

        while (std::getline(ss, rhyme, ',')) {
            size_t first = rhyme.find_first_not_of(" \t");
            size_t last = rhyme.find_last_not_of(" \t");
            if (first != std::string::npos && last != std::string::npos) {
                std::wstring wideRhyme(rhyme.begin() + first, rhyme.begin() + last + 1);
                rhymes.push_back(wideRhyme);
            }
        }

        std::wstring wideKey(key.begin(), key.end());
        rhymeDict[wideKey] = rhymes;
    }

    file.close();
}


std::vector<std::wstring> GetRhymes(const std::wstring& word) {
    auto it = rhymeDict.find(word);
    if (it != rhymeDict.end()) return it->second;
    return {};
}

std::wstring GetLastWord(const std::wstring& line) {
    size_t end = line.find_last_not_of(L" \t\n\r");
    if (end == std::wstring::npos) return L"";

    size_t start = line.find_last_of(L" \t\n\r", end);
    return line.substr((start == std::wstring::npos) ? 0 : start + 1, end - start);
}

void ShowRhymes(HWND hwnd, const std::vector<std::wstring>& rhymes) {
    std::wstring message = L"Rhymes:\n";
    for (const auto& word : rhymes) {
        message += L"• " + word + L"\n";
    }

    MessageBox(hwnd, message.c_str(), L"Rhyme Suggestions", MB_OK | MB_ICONINFORMATION);
}

void ShowRhymesForCurrentLine(HWND hwnd, HWND hEdit) {
    DWORD selStart = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&selStart, (LPARAM)nullptr);

    TEXTRANGEW tr = {};
    wchar_t buffer[2048] = {};
    tr.chrg.cpMin = 0;
    tr.chrg.cpMax = selStart;
    tr.lpstrText = buffer;
    SendMessage(hEdit, EM_GETTEXTRANGE, 0, (LPARAM)&tr);

    std::wstring textUpToCursor = buffer;
    size_t lastNewline = textUpToCursor.find_last_of(L"\r\n");
    std::wstring currentLine = (lastNewline != std::wstring::npos) ?
        textUpToCursor.substr(lastNewline + 1) : textUpToCursor;

    std::wstring lastWord = GetLastWord(currentLine);

    if (!lastWord.empty()) {
        auto it = rhymeDict.find(lastWord);
        if (it != rhymeDict.end() && !it->second.empty()) {
            ShowRhymes(hwnd, it->second);
        } else {
            MessageBox(hwnd, L"No rhymes found.", L"Rhyme Suggestions", MB_OK | MB_ICONINFORMATION);
        }
    } else {
        MessageBox(hwnd, L"No word found.", L"Rhyme Suggestions", MB_OK | MB_ICONWARNING);
    }
}


// Function: Counts syllables
int EstimateSyllables(const std::wstring& word) {
    int count = 0;
    bool prevVowel = false;
    for (wchar_t ch : word) {
        ch = towlower(ch);
        bool isVowel = wcschr(L"aeiouy", ch);
        if (isVowel && !prevVowel) count++;
        prevVowel = isVowel;
    }

    // Handle common silent 'e'
    if (word.length() > 2 && word.back() == L'e' && !wcschr(L"aeiouy", word[word.length() - 2]))
        count--;

    return std::max(count, 1);
}

bool ConfirmDiscardChanges(HWND hwnd) {
    if (!isModified) return true;

    int result = MessageBox(
        hwnd,
        TEXT("You have unsaved changes. Do you want to discard them?"),
        TEXT("Unsaved Changes"),
        MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2
    );

    return result == IDYES;
}

void UpdateWindowTitle(HWND hwnd) {
    const wchar_t* baseTitle = L"Poet's Desk";
    wchar_t fullTitle[100];

    if (isModified) {
        _snwprintf(fullTitle, 100, L"*%s", baseTitle);
    } else {
        _snwprintf(fullTitle, 100, L"%s", baseTitle);
    }

    SetWindowText(hwnd, fullTitle);
}

void UpdateStatusBar() {
    int length = GetWindowTextLength(hEdit);
    if (length == 0) {
        SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Words: 0   Lines: 0   Chars: 0");
        SendMessage(hStatusBar, SB_SETTEXT, 1, (LPARAM)L"Syllables: 0");
        return;
    }
    wchar_t* text = new wchar_t[length + 1];
    GetWindowText(hEdit, text, length + 1);

    int wordCount = 0, lineCount = 0, charCount = length;
    bool inWord = false;

    for (int i = 0; text[i]; ++i) {
        if (text[i] == L'\n') lineCount++;
        if (iswspace(text[i])) {
            inWord = false;
        } else if (!inWord) {
            wordCount++;
            inWord = true;
        }
    }
    lineCount++;  // account for final line

    wchar_t buffer[128];
    _snwprintf(buffer, 128, L"Words: %d   Lines: %d   Chars: %d", wordCount, lineCount, charCount);
    SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)buffer);

     // Get caret position and line
     DWORD pos = SendMessage(hEdit, EM_GETSEL, 0, 0);
     int caretIndex = LOWORD(pos);
     int line = SendMessage(hEdit, EM_LINEFROMCHAR, caretIndex, 0);
 
     wchar_t lineBuf[512] = { 0 };
     *(WORD*)lineBuf = 511;
     SendMessage(hEdit, EM_GETLINE, line, (LPARAM)lineBuf);
     lineBuf[511] = L'\0';
 
     std::wstringstream stream(lineBuf);
     std::wstring word;
     int syllableCount = 0;
 
     while (stream >> word) {
         syllableCount += EstimateSyllables(word);
     }
 
     wchar_t sylBuf[64];
     _snwprintf(sylBuf, 64, L"Syllables: %d", syllableCount);
     SendMessage(hStatusBar, SB_SETTEXT, 1, (LPARAM)sylBuf);

    delete[] text;
}

LRESULT CALLBACK PreviewWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            // Create "Save as PDF" button
            hSavePDFButton = CreateWindow(
                L"BUTTON", L"Save as PDF",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                10, 10, 100, 30,
                hwnd, (HMENU)1001, GetModuleHandle(nullptr), nullptr
            );
            return 0;
        }
        case WM_COMMAND: {
            if (LOWORD(wParam) == 1001) { // Save as PDF button clicked
                SaveAsPDF(hwnd, hEdit);
            }
            return 0;
        }
        case WM_SIZE: {
            // Reposition button when window is resized
            if (hSavePDFButton) {
                MoveWindow(hSavePDFButton, 10, 10, 100, 30, TRUE);
            }
            return 0;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            if (hPreviewBitmap) {
                HDC memDC = CreateCompatibleDC(hdc);
                HBITMAP old = (HBITMAP)SelectObject(memDC, hPreviewBitmap);
                BITMAP bmp;
                GetObject(hPreviewBitmap, sizeof(BITMAP), &bmp);
                
                // Offset the bitmap to account for the button area
                BitBlt(hdc, 0, 50, bmp.bmWidth, bmp.bmHeight, memDC, 0, 0, SRCCOPY);
                
                SelectObject(memDC, old);
                DeleteDC(memDC);
            }
            EndPaint(hwnd, &ps);
            return 0;
        }        
        case WM_DESTROY:
            // Don't call PostQuitMessage - just let the window close
            // The message loop in ShowPrintPreview will exit when IsWindow returns false
            return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void GeneratePrintPreview(HWND hEdit, HDC targetDC) {
    FORMATRANGE fr = {};
    fr.hdc = targetDC;
    fr.hdcTarget = targetDC;
    fr.rcPage = {0, 0, 12240, 15840}; // 8.5x11"
    fr.rc = {1440, 1440, 12240 - 1440, 15840 - 1440};
    fr.chrg.cpMin = 0;
    fr.chrg.cpMax = -1;

    SendMessage(hEdit, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
    SendMessage(hEdit, EM_FORMATRANGE, FALSE, 0);
}


void ShowPrintPreview(HWND hwndParent, HWND hEdit) {
    const int width = 850;
    const int height = 1100;

    // Get screen DC for compatibility
    HDC hdcScreen = GetDC(hwndParent);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);

    // Create the bitmap and select it into memory DC
    if (hPreviewBitmap) DeleteObject(hPreviewBitmap);
    hPreviewBitmap = CreateCompatibleBitmap(hdcScreen, width, height);
    HBITMAP oldBitmap = (HBITMAP)SelectObject(hdcMem, hPreviewBitmap);

    // White background
    RECT rect = { 0, 0, width, height };
    FillRect(hdcMem, &rect, (HBRUSH)GetStockObject(WHITE_BRUSH));

    // DON'T set mapping mode here - work in pixels first
    // SetMapMode(hdcMem, MM_TWIPS); // Remove this line

    // Calculate scaling from twips to pixels
    int dpiX = GetDeviceCaps(hdcScreen, LOGPIXELSX);
    int dpiY = GetDeviceCaps(hdcScreen, LOGPIXELSY);
    
    // Page dimensions in twips (8.5 x 11 inches)
    int pageWidthTwips = 8.5 * 1440;   // 12240 twips
    int pageHeightTwips = 11 * 1440;   // 15840 twips
    
    // Margin in twips (1 inch)
    int marginTwips = 1440;
    
    // Convert to pixels for our preview
    int pageWidthPixels = (pageWidthTwips * dpiX) / 1440;
    int pageHeightPixels = (pageHeightTwips * dpiY) / 1440;
    int marginPixels = (marginTwips * dpiX) / 1440;
    
    // Scale to fit in our preview window
    double scaleX = (double)width / pageWidthPixels;
    double scaleY = (double)height / pageHeightPixels;
    double scale = std::min(scaleX, scaleY);
    
    int scaledWidth = (int)(pageWidthPixels * scale);
    int scaledHeight = (int)(pageHeightPixels * scale);
    int scaledMargin = (int)(marginPixels * scale);
    
    // Center the page in the preview
    int offsetX = (width - scaledWidth) / 2;
    int offsetY = (height - scaledHeight) / 2;

    // Draw page border
    HPEN pen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
    HPEN oldPen = (HPEN)SelectObject(hdcMem, pen);
    HBRUSH oldBrush = (HBRUSH)SelectObject(hdcMem, GetStockObject(NULL_BRUSH));
    Rectangle(hdcMem, offsetX, offsetY, offsetX + scaledWidth, offsetY + scaledHeight);
    SelectObject(hdcMem, oldPen);
    SelectObject(hdcMem, oldBrush);
    DeleteObject(pen);

    // Now set up for RichEdit formatting - use a separate DC with TWIPS
    HDC hdcFormat = CreateCompatibleDC(hdcScreen);
    SetMapMode(hdcFormat, MM_TWIPS);
    
    // Set viewport to match our scaled preview area
    SetViewportOrgEx(hdcFormat, offsetX + scaledMargin, offsetY + scaledMargin, NULL);
    SetViewportExtEx(hdcFormat, scaledWidth - 2 * scaledMargin, scaledHeight - 2 * scaledMargin, NULL);
    SetWindowExtEx(hdcFormat, pageWidthTwips - 2 * marginTwips, pageHeightTwips - 2 * marginTwips, NULL);
    
    // Set up format range
    FORMATRANGE fr = {};
    fr.hdc = hdcMem;        // Render to our bitmap
    fr.hdcTarget = hdcFormat; // Use format DC for measurements
    
    // Page and content rectangles in twips
    fr.rcPage.left = 0;
    fr.rcPage.top = 0;
    fr.rcPage.right = pageWidthTwips;
    fr.rcPage.bottom = pageHeightTwips;
    
    fr.rc.left = marginTwips;
    fr.rc.top = marginTwips;
    fr.rc.right = pageWidthTwips - marginTwips;
    fr.rc.bottom = pageHeightTwips - marginTwips;
    
    fr.chrg.cpMin = 0;
    fr.chrg.cpMax = -1;

    // Check if there's actually text to render
    int textLength = GetWindowTextLength(hEdit);
    if (textLength == 0) {
        // Draw "No content" message
        SetBkMode(hdcMem, TRANSPARENT);
        SetTextColor(hdcMem, RGB(128, 128, 128));
        RECT textRect = { offsetX + scaledMargin, offsetY + scaledMargin, 
                         offsetX + scaledWidth - scaledMargin, offsetY + scaledHeight - scaledMargin };
        DrawText(hdcMem, L"No content to preview", -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    } else {
        // Render the actual content
        LONG result = SendMessage(hEdit, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
        
        // Debug: Check if anything was rendered
        if (result <= 0) {
            SetBkMode(hdcMem, TRANSPARENT);
            SetTextColor(hdcMem, RGB(255, 0, 0));
            RECT textRect = { offsetX + scaledMargin, offsetY + scaledMargin, 
                             offsetX + scaledWidth - scaledMargin, offsetY + scaledHeight - scaledMargin };
            DrawText(hdcMem, L"Rendering failed", -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
    }
    
    // Clean up
    SendMessage(hEdit, EM_FORMATRANGE, FALSE, 0);
    DeleteDC(hdcFormat);
    SelectObject(hdcMem, oldBitmap);
    DeleteDC(hdcMem);
    ReleaseDC(hwndParent, hdcScreen);

    // Register window class for preview if needed
    static bool registered = false;
    if (!registered) {
        WNDCLASS wc = {};
        wc.lpfnWndProc = PreviewWndProc;
        wc.hInstance = GetModuleHandle(nullptr);
        wc.lpszClassName = L"PreviewWindow";
        wc.hbrBackground = (HBRUSH)(COLOR_APPWORKSPACE + 1);
        RegisterClass(&wc);
        registered = true;
    }

    // Create the preview window
    HWND hPreviewWnd = CreateWindowEx(
        WS_EX_CLIENTEDGE, L"PreviewWindow", L"Print Preview",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        100, 100, width + 32, height + 130,  // Extra height for button
        hwndParent, nullptr, GetModuleHandle(nullptr), nullptr
    );

    // Force a repaint
    InvalidateRect(hPreviewWnd, nullptr, TRUE);
    UpdateWindow(hPreviewWnd);

    // Simple modal loop
    MSG msg;
    while (IsWindow(hPreviewWnd) && GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

void SaveAsPDF(HWND hwndParent, HWND hEdit) {
    // Open file dialog to get PDF save location
    OPENFILENAME ofn = {};
    TCHAR szFile[MAX_PATH] = TEXT("");
    
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwndParent;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = TEXT("PDF Files (*.pdf)\0*.pdf\0All Files\0*.*\0");
    ofn.nFilterIndex = 1;
    ofn.lpstrDefExt = TEXT("pdf");
    ofn.Flags = OFN_OVERWRITEPROMPT;
    
    if (GetSaveFileName(&ofn)) {
        // Create a device context for "Microsoft Print to PDF"
        HDC hPrinterDC = CreateDC(L"WINSPOOL", L"Microsoft Print to PDF", nullptr, nullptr);
        
        if (!hPrinterDC) {
            MessageBox(hwndParent, 
                L"Microsoft Print to PDF printer not found.\n\n"
                L"Please install Microsoft Print to PDF or use File -> Print and select a PDF printer manually.",
                L"PDF Export", MB_OK | MB_ICONWARNING);
            return;
        }
        
        // Show a message about the PDF save location
        wchar_t message[512];
        _snwprintf(message, 512, 
            L"PDF will be saved to:\n%s\n\n"
            L"Click OK to continue with PDF generation.",
            szFile);
        
        int result = MessageBox(hwndParent, message, L"Save as PDF", MB_OKCANCEL | MB_ICONINFORMATION);
        if (result != IDOK) {
            DeleteDC(hPrinterDC);
            return;
        }
        
        // Set up document info
        DOCINFO di = {};
        di.cbSize = sizeof(di);
        di.lpszDocName = L"Poem";
        di.lpszOutput = szFile;  // This tells the PDF printer where to save
        
        if (StartDoc(hPrinterDC, &di) > 0) {
            // Get printer capabilities
            int physicalWidth = GetDeviceCaps(hPrinterDC, PHYSICALWIDTH);
            int physicalHeight = GetDeviceCaps(hPrinterDC, PHYSICALHEIGHT);
            int logPixelsX = GetDeviceCaps(hPrinterDC, LOGPIXELSX);
            int logPixelsY = GetDeviceCaps(hPrinterDC, LOGPIXELSY);
            
            // Convert to twips (1 inch = 1440 twips)
            int pageWidthTwips = (physicalWidth * 1440) / logPixelsX;
            int pageHeightTwips = (physicalHeight * 1440) / logPixelsY;
            
            // 1 inch margins in twips
            int marginTwips = 1440;
            
            FORMATRANGE fr = {};
            fr.hdc = hPrinterDC;
            fr.hdcTarget = hPrinterDC;
            
            // Set up page rectangle (entire page)
            fr.rcPage.left = 0;
            fr.rcPage.top = 0;
            fr.rcPage.right = pageWidthTwips;
            fr.rcPage.bottom = pageHeightTwips;
            
            // Set up content rectangle (with margins)
            fr.rc.left = marginTwips;
            fr.rc.top = marginTwips;
            fr.rc.right = pageWidthTwips - marginTwips;
            fr.rc.bottom = pageHeightTwips - marginTwips;
            
            // Character range - render all text
            fr.chrg.cpMin = 0;
            fr.chrg.cpMax = -1;
            
            // DON'T set mapping mode - let RichEdit handle the coordinate system
            // SetMapMode(hPrinterDC, MM_TWIPS); // Remove this line
            
            // Check if there's text to render
            int textLength = GetWindowTextLength(hEdit);
            if (textLength == 0) {
                MessageBox(hwndParent, L"No text to save as PDF.", L"Save as PDF", MB_OK | MB_ICONWARNING);
                EndDoc(hPrinterDC);
                DeleteDC(hPrinterDC);
                return;
            }
            
            // Render pages
            LONG nextChar = 0;
            int pageCount = 0;
            
            do {
                StartPage(hPrinterDC);
                pageCount++;
                
                fr.chrg.cpMin = nextChar;
                fr.chrg.cpMax = -1; // To end of text
                
                nextChar = SendMessage(hEdit, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
                
                EndPage(hPrinterDC);
                
                // Break if no more characters to render
                if (nextChar == -1 || nextChar <= fr.chrg.cpMin) {
                    break;
                }
                
            } while (nextChar > 0 && nextChar < textLength);
            
            EndDoc(hPrinterDC);
            
            wchar_t successMsg[256];
            _snwprintf(successMsg, 256, L"PDF saved successfully!\n\nPages: %d\nLocation: %s", pageCount, szFile);
            MessageBox(hwndParent, successMsg, L"Save as PDF", MB_OK | MB_ICONINFORMATION);
        } else {
            MessageBox(hwndParent, L"Failed to start PDF document.", L"Error", MB_OK | MB_ICONERROR);
        }
        
        // Clean up
        SendMessage(hEdit, EM_FORMATRANGE, FALSE, 0);
        DeleteDC(hPrinterDC);
    }
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_KEYDOWN: {
            if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
                switch (wParam) {
                    case 'S': {
                        SendMessage(hwnd, WM_COMMAND, 3, 0); // simulate Save menu click
                        return 0;
                    }
                }
            }
            if (wParam == 'R' && (GetKeyState(VK_CONTROL) & 0x8000)) {
                ShowRhymesForCurrentLine(hwnd, hEdit);
            }
            break;
        }        
        case WM_CLOSE:
            if (ConfirmDiscardChanges(hwnd)) {
                DestroyWindow(hwnd); // Proceed with closing
            }
            return 0;
        
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;

        case WM_SIZE: {
            SendMessage(hStatusBar, WM_SIZE, 0, 0);
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            MoveWindow(hEdit, 10, 10, width - 20, height - 40, TRUE); // 40 for room below
            return 0;
        }
        case WM_ERASEBKGND: {
            HDC hdc = (HDC)wParam;
            RECT rc;
            GetClientRect(hwnd, &rc);
            FillRect(hdc, &rc, (HBRUSH)(COLOR_WINDOW + 1));
            return 1;
        }
        case WM_COMMAND: {
            if (HIWORD(wParam) == EN_CHANGE && (HWND)lParam == hEdit) {
                isModified = true;
                UpdateWindowTitle(hwnd);
                UpdateStatusBar();
                break;
            }            
            OPENFILENAME ofn = {};
            TCHAR szFile[MAX_PATH] = TEXT("");

            switch (LOWORD(wParam)) {
                case ID_EDIT_UNDO:
                    SendMessage(hEdit, EM_UNDO, 0, 0);
                    break;
                case 2001: { // Font customization
                    CHOOSEFONT cf = {};
                    cf.lStructSize = sizeof(cf);
                    cf.hwndOwner = hwnd;
                    cf.lpLogFont = &lf;
                    cf.Flags = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT;
                
                    if (ChooseFont(&cf)) {
                        if (hFont) DeleteObject(hFont);  // free previous font
                
                        hFont = CreateFontIndirect(&lf);
                        SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
                    }
                    break;
                }             
                case 1001: { // Sonnet
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    const wchar_t* sonnetTemplate =
                        L"[Title]\r\n\r\n[Author]\r\n\r\n"
                        L"1. ____________________________________________\r\n"
                        L"2. ____________________________________________\r\n"
                        L"3. ____________________________________________\r\n"
                        L"4. ____________________________________________\r\n\r\n"
                        L"5. ____________________________________________\r\n"
                        L"6. ____________________________________________\r\n"
                        L"7. ____________________________________________\r\n"
                        L"8. ____________________________________________\r\n\r\n"
                        L"9. ____________________________________________\r\n"
                        L"10. ___________________________________________\r\n"
                        L"11. ___________________________________________\r\n\r\n"
                        L"12. ___________________________________________\r\n\r"
                        L"13. ___________________________________________\r\n"
                        L"14. ___________________________________________\r\n";
                    
                    SendMessage(hEdit, EM_BEGINUNDOACTION, 0, 0);
                    SetWindowText(hEdit, sonnetTemplate);
                    SendMessage(hEdit, EM_ENDUNDOACTION, 0, 0);
                    // SendMessage(hEdit, EM_EMPTYUNDOBUFFER, 0, 0);
                    isModified = false;
                    UpdateWindowTitle(hwnd);
                    break;
                }

                case 1002: { // Haiku
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    const wchar_t* haikuTemplate =
                        L"[Title]\r\n\r\n[Author]\r\n\r\n"
                        L"1. ______________________ (5 syllables)\r\n"
                        L"2. _____________________________ (7 syllables)\r\n"
                        L"3. ______________________ (5 syllables)\r\n";
                    
                    SendMessage(hEdit, EM_BEGINUNDOACTION, 0, 0);
                    SetWindowText(hEdit, haikuTemplate);
                    SendMessage(hEdit, EM_ENDUNDOACTION, 0, 0);
                    // SendMessage(hEdit, EM_EMPTYUNDOBUFFER, 0, 0);
                    isModified = false;
                    UpdateWindowTitle(hwnd);
                    break;
                }

                case 1003: { // Free Verse
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    const wchar_t* freeVerseTemplate =
                        L"[Title]\r\n\r\n[Author]\r\n\r\n"
                        L"(Begin your poem below)\r\n\r\n";

                    SendMessage(hEdit, EM_BEGINUNDOACTION, 0, 0);
                    SetWindowText(hEdit, freeVerseTemplate);
                    SendMessage(hEdit, EM_ENDUNDOACTION, 0, 0);
                    // SendMessage(hEdit, EM_EMPTYUNDOBUFFER, 0, 0);
                    
                    isModified = false;
                    UpdateWindowTitle(hwnd);
                    break;
                }   
                case 2: { // Open
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = hwnd;
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.lpstrFilter = TEXT("Text Files (*.txt)\0*.txt\0All Files\0*.*\0");
                    ofn.nFilterIndex = 1;
                    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

                    if (GetOpenFileName(&ofn)) {
                        HANDLE hFile = CreateFile(
                            szFile, GENERIC_READ, 0, nullptr, OPEN_EXISTING,
                            FILE_ATTRIBUTE_NORMAL, nullptr
                        );

                        if (hFile != INVALID_HANDLE_VALUE) {
                            DWORD fileSize = GetFileSize(hFile, nullptr);
                            if (fileSize > 0) {
                                char* buffer = new char[fileSize + 1];
                                DWORD bytesRead;
                                ReadFile(hFile, buffer, fileSize, &bytesRead, nullptr);
                                buffer[bytesRead] = '\0';

                                int wlen = MultiByteToWideChar(CP_ACP, 0, buffer, -1, nullptr, 0);
                                wchar_t* wbuffer = new wchar_t[wlen];
                                MultiByteToWideChar(CP_ACP, 0, buffer, -1, wbuffer, wlen);

                                SendMessage(hEdit, EM_BEGINUNDOACTION, 0, 0);
                                SetWindowText(hEdit, wbuffer);
                                SendMessage(hEdit, EM_ENDUNDOACTION, 0, 0);
                                // SendMessage(hEdit, EM_EMPTYUNDOBUFFER, 0, 0);

                                delete[] buffer;
                                delete[] wbuffer;
                            }
                            CloseHandle(hFile);
                            isModified = false;
                            UpdateWindowTitle(hwnd);
                        }
                    }
                    break;
                }

                case 3: { // Save
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = hwnd;
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.lpstrFilter = TEXT("Text Files (*.txt)\0*.txt\0All Files\0*.*\0");
                    ofn.nFilterIndex = 1;
                    ofn.Flags = OFN_OVERWRITEPROMPT;

                    if (GetSaveFileName(&ofn)) {
                        int length = GetWindowTextLength(hEdit);
                        if (length > 0) {
                            wchar_t* wbuffer = new wchar_t[length + 1];
                            GetWindowText(hEdit, wbuffer, length + 1);

                            int len = WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, nullptr, 0, nullptr, nullptr);
                            char* buffer = new char[len];
                            WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, buffer, len, nullptr, nullptr);

                            HANDLE hFile = CreateFile(
                                szFile, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                FILE_ATTRIBUTE_NORMAL, nullptr
                            );

                            if (hFile != INVALID_HANDLE_VALUE) {
                                DWORD bytesWritten;
                                WriteFile(hFile, buffer, strlen(buffer), &bytesWritten, nullptr);
                                CloseHandle(hFile);
                                isModified = false;
                                UpdateWindowTitle(hwnd);
                            }

                            delete[] wbuffer;
                            delete[] buffer;
                        }
                    }
                    break;
                }

                case 4: // Exit
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    PostMessage(hwnd, WM_CLOSE, 0, 0);
                    break;
                case 5: { // Print
                    PRINTDLG pd = {};
                    pd.lStructSize = sizeof(pd);
                    pd.hwndOwner = hwnd;
                    pd.Flags = PD_RETURNDC | PD_USEDEVMODECOPIESANDCOLLATE | PD_NOPAGENUMS | PD_NOSELECTION;

                    if (PrintDlg(&pd)) {
                        if (!pd.hDC) {
                            MessageBox(hwnd, L"Printing was canceled or failed to initialize.", L"Print", MB_ICONERROR);
                            return 0;
                        }

                        // Inform user to complete Save As if using PDF printer
                        MessageBox(hwnd, L"If using a PDF printer, complete the Save As dialog before continuing.", L"Print", MB_OK);

                        DOCINFO di = {};
                        di.cbSize = sizeof(di);
                        di.lpszDocName = L"Poem";

                        if (StartDoc(pd.hDC, &di) > 0) {
                            FORMATRANGE fr = {};
                            fr.hdc = pd.hDC;
                            fr.hdcTarget = pd.hDC;

                            // Printable area in twips (1 inch = 1440 twips)
                            fr.rcPage.left = 1440;
                            fr.rcPage.top = 1440;
                            fr.rcPage.right = 12240;
                            fr.rcPage.bottom = 15480;

                            fr.rc = fr.rcPage;           // Content area
                            fr.chrg.cpMin = 0;
                            fr.chrg.cpMax = -1;          // To end of text

                            SetMapMode(pd.hDC, MM_TWIPS);

                            int page = 1;
                            LONG nextChar;
                            do {
                                StartPage(pd.hDC);
                                nextChar = SendMessage(hEdit, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
                                EndPage(pd.hDC);

                                fr.chrg.cpMin = nextChar;
                            } while (nextChar < fr.chrg.cpMax);

                            EndDoc(pd.hDC);
                        }

                        DeleteDC(pd.hDC);
                        SendMessage(hEdit, EM_FORMATRANGE, FALSE, 0); // Cleanup cached info
                    }

                    break;
                }
                case ID_PREVIEW: {
                    ShowPrintPreview(hwnd, hEdit);
                    break;
                }
            }
            return 0;
        }
        case WM_NOTIFY: {
            NMHDR* pNMHDR = (NMHDR*)lParam;
            if (pNMHDR->hwndFrom == hEdit && pNMHDR->code == EN_CHANGE) {
                isModified = true;
                UpdateWindowTitle(hwnd);
                UpdateStatusBar();
                return 0;
            }
            break;
        }
        
        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    LoadRhymeDictionary("rhyme_dictionary.txt");
    InitCommonControls();
    
    ACCEL accelTable[] = {
        { FCONTROL | FVIRTKEY, 'S', ID_FILE_SAVE },
        { FCONTROL | FVIRTKEY, 'Z', ID_EDIT_UNDO },
    };
    
    HACCEL hAccel = CreateAcceleratorTable(accelTable, 2);
    LPCTSTR CLASS_NAME = TEXT("PoetsDeskWindowClass");

    WNDCLASS wc = {};
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;

    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(
        0,
        CLASS_NAME,
        TEXT("Poet's Desk"),
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        nullptr, nullptr, hInstance, nullptr
    );

    if (hwnd == nullptr) return 0;

    // Menu setup
    HMENU hMenu = CreateMenu();
    HMENU hFileMenu = CreatePopupMenu();
    HMENU hNewSubMenu = CreatePopupMenu();
    HMENU hFormatMenu = CreatePopupMenu();

    AppendMenu(hNewSubMenu, MF_STRING, 1001, TEXT("Sonnet"));
    AppendMenu(hNewSubMenu, MF_STRING, 1002, TEXT("Haiku"));
    AppendMenu(hNewSubMenu, MF_STRING, 1003, TEXT("Free Verse"));
    AppendMenu(hFileMenu, MF_POPUP, (UINT_PTR)hNewSubMenu, TEXT("New From Template"));
    AppendMenu(hFileMenu, MF_STRING, 2, TEXT("Open"));
    AppendMenu(hFileMenu, MF_STRING, 3, TEXT("Save"));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenu(hFileMenu, MF_STRING, 4, TEXT("Exit"));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, TEXT("File"));
    AppendMenu(hFormatMenu, MF_STRING, 2001, TEXT("Font"));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFormatMenu, TEXT("Format"));
    AppendMenu(hFileMenu, MF_STRING, 5, TEXT("Print"));
    AppendMenu(hFileMenu, MF_STRING, ID_PREVIEW, TEXT("Print Preview"));

    SetMenu(hwnd, hMenu);

    ShowWindow(hwnd, nCmdShow);

    
    hStatusBar = CreateWindowEx(
        0, STATUSCLASSNAME, nullptr,
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0,
        hwnd, nullptr, hInstance, nullptr
    );

    int parts[] = { 680, -1 };  // Two parts: first 400px, rest fills
    SendMessage(hStatusBar, SB_SETPARTS, 2, (LPARAM)parts);

    LoadLibrary(TEXT("Msftedit.dll"));  // Required to load RichEdit 4.1+

    hEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE, MSFTEDIT_CLASS, TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
        10, 10, 760, 540,
        hwnd, nullptr, hInstance, nullptr
    );

    SendMessage(hEdit, EM_SETOPTIONS, ECOOP_OR, ECO_AUTOWORDSELECTION);
    SendMessage(hEdit, EM_SETLANGOPTIONS, 0, IMF_SPELLCHECKING | IMF_AUTOCORRECT);
    SendMessage(hEdit, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
    SendMessage(hEdit, EM_SETTARGETDEVICE, (WPARAM)nullptr, 0);
    SendMessage(hEdit, EM_SETTEXTMODE, TM_PLAINTEXT, 0); // Ensure RichEdit isn't in RTF mode



    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!TranslateAccelerator(hwnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return 0;
}
