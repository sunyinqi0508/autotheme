#pragma once

#include <Windows.h>

#include <filesystem>
#include <string>

namespace winxsw {

enum class ThemeMode {
    Light,
    Dark,
};

enum class LocationMode {
    Auto,
    Manual,
};

struct HotkeySettings {
    bool enabled = true;
    bool ctrl = true;
    bool alt = true;
    bool shift = false;
    bool win = false;
    UINT virtualKey = 'D';
};

struct Settings {
    int schemaVersion = 1;
    bool autoSwitchEnabled = true;
    bool autostartEnabled = true;
    std::wstring lightWallpaperPath;
    std::wstring darkWallpaperPath;
    LocationMode locationMode = LocationMode::Auto;
    bool allowGps = true;
    bool allowIpFallback = true;
    double manualLatitude = 39.7684;
    double manualLongitude = -86.1581;
    int sunriseOffsetMinutes = 0;
    int sunsetOffsetMinutes = 0;
    int pollIntervalMinutes = 1;
    int locationRefreshMinutes = 360;
    HotkeySettings hotkey{};
    HotkeySettings configHotkey{true, true, true, false, false, 'C'};
};

struct Status {
    bool agentRunning = false;
    bool autoSwitchEnabled = true;
    bool hasLocation = false;
    bool manualOverrideActive = false;
    double latitude = 0.0;
    double longitude = 0.0;
    std::wstring currentMode = L"unknown";
    std::wstring desiredMode = L"unknown";
    std::wstring manualOverrideMode;
    std::wstring manualOverrideUntil;
    std::wstring sunriseTime;
    std::wstring sunsetTime;
    std::wstring nextTransitionTime;
    std::wstring locationSource = L"none";
    std::wstring lastLocationRefresh;
    std::wstring lastAppliedTime;
    std::wstring hotkeyText;
    std::wstring configHotkeyText;
    std::wstring lastError;
};

std::wstring ThemeModeToString(ThemeMode value);
std::wstring LocationModeToString(LocationMode value);

std::filesystem::path GetSettingsPath();
std::filesystem::path GetStatusPath();

Settings DefaultSettings();
bool LoadSettings(Settings& settings, std::wstring* error = nullptr);
bool SaveSettings(const Settings& settings, std::wstring* error = nullptr);
bool LoadStatus(Status& status, std::wstring* error = nullptr);
bool SaveStatus(const Status& status, std::wstring* error = nullptr);

}  // namespace winxsw
