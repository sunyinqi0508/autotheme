#include "Theme.h"

#include <Windows.h>

#include <filesystem>

namespace winxsw {

namespace {

constexpr wchar_t kPersonalizeKey[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";

DWORD ReadDwordValue(const wchar_t* name, DWORD fallback) {
    HKEY key = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kPersonalizeKey, 0, KEY_QUERY_VALUE, &key) != ERROR_SUCCESS) {
        return fallback;
    }

    DWORD value = fallback;
    DWORD type = REG_DWORD;
    DWORD size = sizeof(value);
    const auto status = RegQueryValueExW(
        key,
        name,
        nullptr,
        &type,
        reinterpret_cast<BYTE*>(&value),
        &size);
    RegCloseKey(key);
    return status == ERROR_SUCCESS && type == REG_DWORD ? value : fallback;
}

bool WriteDwordValue(const wchar_t* name, DWORD value) {
    HKEY key = nullptr;
    if (RegCreateKeyExW(
            HKEY_CURRENT_USER,
            kPersonalizeKey,
            0,
            nullptr,
            0,
            KEY_SET_VALUE,
            nullptr,
            &key,
            nullptr) != ERROR_SUCCESS) {
        return false;
    }

    const auto status = RegSetValueExW(
        key,
        name,
        0,
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&value),
        sizeof(value));
    RegCloseKey(key);
    return status == ERROR_SUCCESS;
}

void BroadcastThemeChange() {
    DWORD_PTR ignored = 0;
    SendMessageTimeoutW(
        HWND_BROADCAST,
        WM_SETTINGCHANGE,
        0,
        reinterpret_cast<LPARAM>(L"ImmersiveColorSet"),
        SMTO_ABORTIFHUNG,
        2000,
        &ignored);
    SendMessageTimeoutW(
        HWND_BROADCAST,
        WM_SETTINGCHANGE,
        0,
        reinterpret_cast<LPARAM>(L"WindowsThemeElement"),
        SMTO_ABORTIFHUNG,
        2000,
        &ignored);
    SendMessageTimeoutW(
        HWND_BROADCAST,
        WM_THEMECHANGED,
        0,
        0,
        SMTO_ABORTIFHUNG,
        2000,
        &ignored);
}

std::wstring WallpaperPathForMode(const Settings& settings, ThemeMode mode) {
    return mode == ThemeMode::Light ? settings.lightWallpaperPath : settings.darkWallpaperPath;
}

}  // namespace

ThemeMode GetCurrentThemeMode() {
    return ReadDwordValue(L"AppsUseLightTheme", 1) == 0 ? ThemeMode::Dark : ThemeMode::Light;
}

bool ApplyThemeMode(ThemeMode mode, std::wstring* error) {
    const DWORD lightValue = mode == ThemeMode::Light ? 1 : 0;
    if (!WriteDwordValue(L"AppsUseLightTheme", lightValue) ||
        !WriteDwordValue(L"SystemUsesLightTheme", lightValue)) {
        if (error != nullptr) {
            *error = L"failed to update the Personalize registry values";
        }
        return false;
    }

    BroadcastThemeChange();
    return true;
}

bool ApplyWallpaperForMode(const Settings& settings, ThemeMode mode, std::wstring* error) {
    const auto wallpaperPath = WallpaperPathForMode(settings, mode);
    if (wallpaperPath.empty()) {
        return true;
    }

    std::error_code ec;
    if (!std::filesystem::exists(wallpaperPath, ec) || ec) {
        if (error != nullptr) {
            *error = L"wallpaper file does not exist: " + wallpaperPath;
        }
        return false;
    }

    std::wstring mutablePath = wallpaperPath;
    if (!SystemParametersInfoW(
            SPI_SETDESKWALLPAPER,
            0,
            mutablePath.data(),
            SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)) {
        if (error != nullptr) {
            *error = L"failed to set wallpaper for " + ThemeModeToString(mode) + L" mode";
        }
        return false;
    }

    return true;
}

}  // namespace winxsw
