#include <windows.h>
#include <commdlg.h>
#define UNICODE
#define _UNICODE

HWND hEdit;

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;

        case WM_SIZE: {
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            MoveWindow(hEdit, 10, 10, width - 20, height - 20, TRUE);
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
            OPENFILENAME ofn = {};
            TCHAR szFile[MAX_PATH] = TEXT("");

            switch (LOWORD(wParam)) {
                case 1: // New
                    SetWindowText(hEdit, TEXT(""));
                    break;

                case 2: { // Open
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

                                // Convert ANSI to Unicode
                                int wlen = MultiByteToWideChar(CP_ACP, 0, buffer, -1, nullptr, 0);
                                wchar_t* wbuffer = new wchar_t[wlen];
                                MultiByteToWideChar(CP_ACP, 0, buffer, -1, wbuffer, wlen);

                                SetWindowText(hEdit, wbuffer);

                                delete[] buffer;
                                delete[] wbuffer;
                            }
                            CloseHandle(hFile);
                        }
                    }
                    break;
                }

                case 3: {// Save (to be implemented)
                    OPENFILENAME ofn = {};
                    TCHAR szFile[MAX_PATH] = TEXT("");
                
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = hwnd;
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.lpstrFilter = TEXT("Text Files (*.txt)\0*.txt\0All Files\0*.*\0");
                    ofn.nFilterIndex = 1;
                    ofn.Flags = OFN_OVERWRITEPROMPT;
                
                    if (GetSaveFileName(&ofn)) {
                        // Get text from the edit control
                        int length = GetWindowTextLength(hEdit);
                        if (length > 0) {
                            wchar_t* wbuffer = new wchar_t[length + 1];
                            GetWindowText(hEdit, wbuffer, length + 1);
                
                            // Convert to ANSI
                            int len = WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, nullptr, 0, nullptr, nullptr);
                            char* buffer = new char[len];
                            WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, buffer, len, nullptr, nullptr);
                
                            // Write to file
                            HANDLE hFile = CreateFile(
                                szFile, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                FILE_ATTRIBUTE_NORMAL, nullptr
                            );
                
                            if (hFile != INVALID_HANDLE_VALUE) {
                                DWORD bytesWritten;
                                WriteFile(hFile, buffer, strlen(buffer), &bytesWritten, nullptr);
                                CloseHandle(hFile);
                            }
                
                            delete[] wbuffer;
                            delete[] buffer;
                        }
                    }
                    break;
                }
                case 4: // Exit
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

    // Create menu bar
    HMENU hMenu = CreateMenu();
    HMENU hFileMenu = CreatePopupMenu();
    AppendMenu(hFileMenu, MF_STRING, 1, TEXT("New"));
    AppendMenu(hFileMenu, MF_STRING, 2, TEXT("Open"));
    AppendMenu(hFileMenu, MF_STRING, 3, TEXT("Save"));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenu(hFileMenu, MF_STRING, 4, TEXT("Exit"));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, TEXT("File"));
    SetMenu(hwnd, hMenu);

    ShowWindow(hwnd, nCmdShow);

    hEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE, TEXT("EDIT"), TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
        10, 10, 760, 540,
        hwnd, nullptr, hInstance, nullptr
    );

    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
