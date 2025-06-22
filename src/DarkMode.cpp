// DarkMode.cpp
#include "DarkMode.h"



static bool isDarkMode = false;

static COLORREF darkBgColor = RGB(30, 30, 30);
static COLORREF darkTextColor = RGB(220, 220, 220);

static COLORREF lightBgColor = RGB(255, 255, 255);
static COLORREF lightTextColor = RGB(0, 0, 0);

static HBRUSH hDarkBrush = CreateSolidBrush(darkBgColor);
static HBRUSH hLightBrush = CreateSolidBrush(lightBgColor);

// Toggle handler
void ApplyDarkMode(bool enable, HWND hwnd) {
    isDarkMode = enable;
    InvalidateRect(hwnd, NULL, TRUE);  // Force repaint of all child controls

}

// Rich Edit control background + text color
void ApplyDarkModeToEdit(HWND hEdit) {
    COLORREF bg = isDarkMode ? darkBgColor : lightBgColor;
    COLORREF text = isDarkMode ? darkTextColor : lightTextColor;

    SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)bg);

    CHARFORMAT2 cf = {};
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_COLOR;
    cf.crTextColor = text;
    SendMessage(hEdit, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cf);
}

// Static controls (e.g. buttons, labels)
void ApplyDarkModeToStaticControl(HDC hdc) {
    SetTextColor(hdc, isDarkMode ? darkTextColor : lightTextColor);
    SetBkColor(hdc, isDarkMode ? darkBgColor : lightBgColor);
}

// Return background brush for static controls
HBRUSH GetDarkModeBrush() {
    return isDarkMode ? hDarkBrush : hLightBrush;
}

