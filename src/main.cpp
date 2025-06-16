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

#pragma comment(lib, "Msftedit.lib")
#define UNICODE
#define _UNICODE
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

#pragma comment(lib, "comdlg32.lib")

HWND hStatusBar;
HWND hEdit;
bool isModified = false;
HFONT hFont = nullptr;
LOGFONT lf = {};

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

void CountSyllablesPerLine() {
    int length = GetWindowTextLength(hEdit);
    if (length == 0) return;

    std::wstring text(length, L'\0');
    GetWindowText(hEdit, &text[0], length + 1);

    std::wstringstream stream(text);
    std::wstring line;
    int lineNum = 1;

    OutputDebugString(L"\nSyllable Counts:\n");

    while (std::getline(stream, line)) {
        std::wstringstream words(line);
        std::wstring word;
        int syllableCount = 0;

        while (words >> word) {
            syllableCount += EstimateSyllables(word);
        }

        wchar_t buffer[128];
        _snwprintf(buffer, 128, L"Line %d: %d syllables\n", lineNum++, syllableCount);
        OutputDebugString(buffer);
    }
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
                        DOCINFO di = {};
                        di.cbSize = sizeof(di);
                        di.lpszDocName = L"Poem";

                        if (StartDoc(pd.hDC, &di) > 0) {
                            if (StartPage(pd.hDC) > 0) {

                                // Force layout update
                                RedrawWindow(hEdit, nullptr, nullptr, RDW_INTERNALPAINT | RDW_UPDATENOW);

                                FORMATRANGE fr = {};
                                fr.hdc = pd.hDC;
                                fr.hdcTarget = pd.hDC;
                                fr.rcPage.left = 1440;
                                fr.rcPage.top = 1440;
                                fr.rcPage.right = 12240;
                                fr.rcPage.bottom = 15480;

                                fr.rc = fr.rcPage;

                                fr.chrg.cpMin = 0;
                                fr.chrg.cpMax = -1;

                                // Set map mode for twips
                                SetMapMode(pd.hDC, MM_TWIPS);

                                // Perform actual drawing
                                SendMessage(hEdit, EM_FORMATRANGE, TRUE, (LPARAM)&fr);

                                // End printing
                                EndPage(pd.hDC);
                            }

                            EndDoc(pd.hDC);
                        }

                        DeleteDC(pd.hDC);

                        // Clear cached formatting
                        SendMessage(hEdit, EM_FORMATRANGE, FALSE, 0);
                    }
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


    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!TranslateAccelerator(hwnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return 0;
}
