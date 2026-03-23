#pragma once

#include <string>

#include "Settings.h"

namespace winxsw {

ThemeMode GetCurrentThemeMode();
bool ApplyThemeMode(ThemeMode mode, std::wstring* error = nullptr);
bool ApplyWallpaperForMode(const Settings& settings, ThemeMode mode, std::wstring* error = nullptr);

}  // namespace winxsw
