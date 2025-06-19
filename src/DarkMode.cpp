// DarkMode.cpp
#include "DarkMode.h"

static COLORREF bgColor = RGB(30, 30, 30);
static COLORREF textColor = RGB(220, 220, 220);
static HBRUSH hBackgroundBrush = CreateSolidBrush(bgColor);

void ApplyDarkModeToEdit(HWND hEdit) {
    SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)bgColor);

    CHARFORMAT2 cf = {};
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_COLOR;
    cf.crTextColor = textColor;
    SendMessage(hEdit, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cf);
}

void ApplyDarkModeToStaticControl(HDC hdc) {
    SetTextColor(hdc, textColor);
    SetBkColor(hdc, bgColor);
}

HBRUSH GetDarkModeBrush() {
    return hBackgroundBrush;
}
