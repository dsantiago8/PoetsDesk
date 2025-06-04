#include <windows.h>
#include <commdlg.h>
#include <stdio.h>
#include <commctrl.h>
#include <wctype.h>
#define UNICODE
#define _UNICODE
#define ID_FILE_SAVE 3

HWND hStatusBar;
HWND hEdit;
bool isModified = false;


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
                    SetWindowText(hEdit, sonnetTemplate);
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
                    SetWindowText(hEdit, haikuTemplate);
                    isModified = false;
                    UpdateWindowTitle(hwnd);
                    break;
                }

                case 1003: { // Free Verse
                    if (!ConfirmDiscardChanges(hwnd)) break;
                    const wchar_t* freeVerseTemplate =
                        L"[Title]\r\n\r\n[Author]\r\n\r\n"
                        L"(Begin your poem below)\r\n\r\n";
                    SetWindowText(hEdit, freeVerseTemplate);
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

                                SetWindowText(hEdit, wbuffer);

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
            }
            return 0;
        }

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    InitCommonControls();
    
    ACCEL accelTable[] = {
        { FCONTROL | FVIRTKEY, 'S', ID_FILE_SAVE },
    };
    
    HACCEL hAccel = CreateAcceleratorTable(accelTable, 1);
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
    AppendMenu(hNewSubMenu, MF_STRING, 1001, TEXT("Sonnet"));
    AppendMenu(hNewSubMenu, MF_STRING, 1002, TEXT("Haiku"));
    AppendMenu(hNewSubMenu, MF_STRING, 1003, TEXT("Free Verse"));
    AppendMenu(hFileMenu, MF_POPUP, (UINT_PTR)hNewSubMenu, TEXT("New From Template"));
    AppendMenu(hFileMenu, MF_STRING, 2, TEXT("Open"));
    AppendMenu(hFileMenu, MF_STRING, 3, TEXT("Save"));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenu(hFileMenu, MF_STRING, 4, TEXT("Exit"));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, TEXT("File"));
    SetMenu(hwnd, hMenu);

    ShowWindow(hwnd, nCmdShow);

    
    hStatusBar = CreateWindowEx(
        0, STATUSCLASSNAME, nullptr,
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0,
        hwnd, nullptr, hInstance, nullptr
    );

    hEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE, TEXT("EDIT"), TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
        10, 10, 760, 540,
        hwnd, nullptr, hInstance, nullptr
    );

    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!TranslateAccelerator(hwnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return 0;
}
