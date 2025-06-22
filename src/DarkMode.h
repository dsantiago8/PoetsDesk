// DarkMode.h
#pragma once
#include <windows.h>
#include <richedit.h>

void ApplyDarkMode(bool enable, HWND hwnd);
void ApplyDarkModeToEdit(HWND hEdit);
void ApplyDarkModeToStaticControl(HDC hdc);
HBRUSH GetDarkModeBrush();