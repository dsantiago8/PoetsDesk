// DarkMode.h
#pragma once
#include <windows.h>
#include <richedit.h>

void ApplyDarkModeToEdit(HWND hEdit);
void ApplyDarkModeToStaticControl(HDC hdc);
HBRUSH GetDarkModeBrush();
