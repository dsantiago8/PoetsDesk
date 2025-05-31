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
            // Resize the edit control to fill the client area with padding
            MoveWindow(hEdit, 10, 10, width - 20, height - 20, TRUE);
            return 0;
        }
        case WM_ERASEBKGND:
        // Fill the background with white to avoid artifacts
        {
            HDC hdc = (HDC)wParam;
            RECT rc;
            GetClientRect(hwnd, &rc);
            FillRect(hdc, &rc, (HBRUSH)(COLOR_WINDOW + 1));
            return 1;
        }
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case 1: // New
                    SetWindowText(hEdit, TEXT(""));
                    break;
                case 2: {// Open
                    OPENFILENAME ofn = {};
                    TCHAR szFile[MAX_PATH] = TEXT("");
                
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = hwnd;
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.lpstrFilter = TEXT("Text Files (*.txt)\0*.txt\0All Files\0*.*\0");
                    ofn.nFilterIndex = 1;
                    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
                
                    if (GetOpenFileName(&ofn)) {
                        // Open and read the file
                        HANDLE hFile = CreateFile(
                            szFile, GENERIC_READ, 0, nullptr, OPEN_EXISTING,
                            FILE_ATTRIBUTE_NORMAL, nullptr
                        );
                
                        if (hFile != INVALID_HANDLE_VALUE) {
                            DWORD fileSize = GetFileSize(hFile, nullptr);
                            if (fileSize > 0) {
                                TCHAR* buffer = new TCHAR[fileSize / sizeof(TCHAR) + 1];
                                DWORD bytesRead;
                                ReadFile(hFile, buffer, fileSize, &bytesRead, nullptr);
                                buffer[bytesRead / sizeof(TCHAR)] = '\0';
                                SetWindowText(hEdit, buffer);
                                delete[] buffer;
                            }
                            CloseHandle(hFile);
                        }
                    }
                    break;
                }
                case 3: // Save
                    MessageBox(hwnd, TEXT("Save feature not implemented yet."), TEXT("Info"), MB_OK);
                    break;
                case 4: // Exit
                    PostMessage(hwnd, WM_CLOSE, 0, 0);
                    break;
            }
            return 0;


        // You can add more message handlers here (WM_COMMAND, etc.)

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}



int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    LPCTSTR CLASS_NAME = TEXT("PoetsDeskWindowClass");


    WNDCLASS wc = { };
    wc.lpfnWndProc   = WindowProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.style = CS_HREDRAW | CS_VREDRAW;


    RegisterClass(&wc);

    // Main window
    HWND hwnd = CreateWindowEx(
        0, CLASS_NAME, TEXT("Poet's Desk"),
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        nullptr, nullptr, hInstance, nullptr
    );

    if (hwnd == nullptr) return 0;

    ShowWindow(hwnd, nCmdShow);

    // Multiline edit box
    hEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE, TEXT("EDIT"), TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
        10, 10, 760, 540,  // x, y, width, height
        hwnd, nullptr, hInstance, nullptr
    );

    //Menu bar
    HMENU hMenu = CreateMenu();

    HMENU hFileMenu = CreatePopupMenu();
    AppendMenu(hFileMenu, MF_STRING, 1, TEXT("New"));
    AppendMenu(hFileMenu, MF_STRING, 2, TEXT("Open"));
    AppendMenu(hFileMenu, MF_STRING, 3, TEXT("Save"));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenu(hFileMenu, MF_STRING, 4, TEXT("Exit"));

    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, TEXT("File"));

    // Attach the menu to the window
    SetMenu(hwnd, hMenu);


    MSG msg = { };
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
