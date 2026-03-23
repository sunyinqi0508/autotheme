#include "Settings.h"

#include "Common.h"

#include <winrt/Windows.Data.Json.h>
#include <winrt/base.h>

namespace winxsw {

using winrt::Windows::Data::Json::JsonObject;
using winrt::Windows::Data::Json::JsonValue;

namespace {

template <typename T>
T Clamp(T value, T minValue, T maxValue) {
    if (value < minValue) {
        return minValue;
    }
    if (value > maxValue) {
        return maxValue;
    }
    return value;
}

bool TryParseObject(const std::wstring& text, JsonObject& object) {
    try {
        object = JsonObject::Parse(text);
        return true;
    } catch (...) {
        return false;
    }
}

std::wstring GetStringOr(const JsonObject& object, const wchar_t* key, const std::wstring& fallback) {
    try {
        return std::wstring(object.GetNamedString(key, fallback.c_str()).c_str());
    } catch (...) {
        return fallback;
    }
}

bool GetBoolOr(const JsonObject& object, const wchar_t* key, bool fallback) {
    try {
        return object.GetNamedBoolean(key, fallback);
    } catch (...) {
        return fallback;
    }
}

double GetDoubleOr(const JsonObject& object, const wchar_t* key, double fallback) {
    try {
        return object.GetNamedNumber(key, fallback);
    } catch (...) {
        return fallback;
    }
}

int GetIntOr(const JsonObject& object, const wchar_t* key, int fallback) {
    return static_cast<int>(GetDoubleOr(object, key, static_cast<double>(fallback)));
}

JsonObject MakeHotkeyObject(const HotkeySettings& hotkey) {
    JsonObject object;
    object.SetNamedValue(L"enabled", JsonValue::CreateBooleanValue(hotkey.enabled));
    object.SetNamedValue(L"ctrl", JsonValue::CreateBooleanValue(hotkey.ctrl));
    object.SetNamedValue(L"alt", JsonValue::CreateBooleanValue(hotkey.alt));
    object.SetNamedValue(L"shift", JsonValue::CreateBooleanValue(hotkey.shift));
    object.SetNamedValue(L"win", JsonValue::CreateBooleanValue(hotkey.win));
    object.SetNamedValue(L"virtualKey", JsonValue::CreateNumberValue(hotkey.virtualKey));
    return object;
}

void LoadHotkeyObject(const JsonObject& object, HotkeySettings& hotkey) {
    hotkey.enabled = GetBoolOr(object, L"enabled", hotkey.enabled);
    hotkey.ctrl = GetBoolOr(object, L"ctrl", hotkey.ctrl);
    hotkey.alt = GetBoolOr(object, L"alt", hotkey.alt);
    hotkey.shift = GetBoolOr(object, L"shift", hotkey.shift);
    hotkey.win = GetBoolOr(object, L"win", hotkey.win);
    hotkey.virtualKey = static_cast<UINT>(
        Clamp(GetIntOr(object, L"virtualKey", static_cast<int>(hotkey.virtualKey)), 0, 255));
}

bool WriteJsonFile(const std::filesystem::path& path, const JsonObject& object, std::wstring* error) {
    const std::wstring text = object.Stringify().c_str();
    if (!WriteUtf8File(path, text)) {
        if (error != nullptr) {
            *error = L"failed to write " + path.filename().wstring();
        }
        return false;
    }
    return true;
}

}  // namespace

std::wstring ThemeModeToString(ThemeMode value) {
    return value == ThemeMode::Light ? L"light" : L"dark";
}

std::wstring LocationModeToString(LocationMode value) {
    return value == LocationMode::Manual ? L"manual" : L"auto";
}

std::filesystem::path GetSettingsPath() {
    return GetDataDirectory() / L"settings.json";
}

std::filesystem::path GetStatusPath() {
    return GetDataDirectory() / L"status.json";
}

Settings DefaultSettings() {
    return {};
}

bool LoadSettings(Settings& settings, std::wstring* error) {
    settings = DefaultSettings();
    const auto text = ReadUtf8File(GetSettingsPath());
    if (!text) {
        SaveSettings(settings, nullptr);
        return true;
    }

    JsonObject object;
    if (!TryParseObject(*text, object)) {
        if (error != nullptr) {
            *error = L"settings.json is not valid JSON";
        }
        return false;
    }

    settings.schemaVersion = Clamp(GetIntOr(object, L"schemaVersion", settings.schemaVersion), 1, 100);
    settings.autoSwitchEnabled = GetBoolOr(object, L"autoSwitchEnabled", settings.autoSwitchEnabled);
    settings.autostartEnabled = GetBoolOr(object, L"autostartEnabled", settings.autostartEnabled);
    settings.lightWallpaperPath = GetStringOr(object, L"lightWallpaperPath", settings.lightWallpaperPath);
    settings.darkWallpaperPath = GetStringOr(object, L"darkWallpaperPath", settings.darkWallpaperPath);
    settings.allowGps = GetBoolOr(object, L"allowGps", settings.allowGps);
    settings.allowIpFallback = GetBoolOr(object, L"allowIpFallback", settings.allowIpFallback);
    settings.manualLatitude = Clamp(GetDoubleOr(object, L"manualLatitude", settings.manualLatitude), -90.0, 90.0);
    settings.manualLongitude = Clamp(GetDoubleOr(object, L"manualLongitude", settings.manualLongitude), -180.0, 180.0);
    settings.sunriseOffsetMinutes = Clamp(GetIntOr(object, L"sunriseOffsetMinutes", settings.sunriseOffsetMinutes), -180, 180);
    settings.sunsetOffsetMinutes = Clamp(GetIntOr(object, L"sunsetOffsetMinutes", settings.sunsetOffsetMinutes), -180, 180);
    settings.pollIntervalMinutes = Clamp(GetIntOr(object, L"pollIntervalMinutes", settings.pollIntervalMinutes), 1, 60);
    settings.locationRefreshMinutes = Clamp(GetIntOr(object, L"locationRefreshMinutes", settings.locationRefreshMinutes), 15, 1440);

    const auto locationMode = GetStringOr(object, L"locationMode", L"auto");
    settings.locationMode = locationMode == L"manual" ? LocationMode::Manual : LocationMode::Auto;

    try {
        LoadHotkeyObject(object.GetNamedObject(L"hotkey"), settings.hotkey);
    } catch (...) {
    }

    return true;
}

bool SaveSettings(const Settings& settings, std::wstring* error) {
    JsonObject object;
    object.SetNamedValue(L"schemaVersion", JsonValue::CreateNumberValue(settings.schemaVersion));
    object.SetNamedValue(L"autoSwitchEnabled", JsonValue::CreateBooleanValue(settings.autoSwitchEnabled));
    object.SetNamedValue(L"autostartEnabled", JsonValue::CreateBooleanValue(settings.autostartEnabled));
    object.SetNamedValue(L"lightWallpaperPath", JsonValue::CreateStringValue(settings.lightWallpaperPath));
    object.SetNamedValue(L"darkWallpaperPath", JsonValue::CreateStringValue(settings.darkWallpaperPath));
    object.SetNamedValue(L"locationMode", JsonValue::CreateStringValue(LocationModeToString(settings.locationMode)));
    object.SetNamedValue(L"allowGps", JsonValue::CreateBooleanValue(settings.allowGps));
    object.SetNamedValue(L"allowIpFallback", JsonValue::CreateBooleanValue(settings.allowIpFallback));
    object.SetNamedValue(L"manualLatitude", JsonValue::CreateNumberValue(settings.manualLatitude));
    object.SetNamedValue(L"manualLongitude", JsonValue::CreateNumberValue(settings.manualLongitude));
    object.SetNamedValue(L"sunriseOffsetMinutes", JsonValue::CreateNumberValue(settings.sunriseOffsetMinutes));
    object.SetNamedValue(L"sunsetOffsetMinutes", JsonValue::CreateNumberValue(settings.sunsetOffsetMinutes));
    object.SetNamedValue(L"pollIntervalMinutes", JsonValue::CreateNumberValue(settings.pollIntervalMinutes));
    object.SetNamedValue(L"locationRefreshMinutes", JsonValue::CreateNumberValue(settings.locationRefreshMinutes));
    object.SetNamedValue(L"hotkey", MakeHotkeyObject(settings.hotkey));
    return WriteJsonFile(GetSettingsPath(), object, error);
}

bool LoadStatus(Status& status, std::wstring* error) {
    status = {};
    const auto text = ReadUtf8File(GetStatusPath());
    if (!text) {
        return true;
    }

    JsonObject object;
    if (!TryParseObject(*text, object)) {
        if (error != nullptr) {
            *error = L"status.json is not valid JSON";
        }
        return false;
    }

    status.agentRunning = GetBoolOr(object, L"agentRunning", false);
    status.autoSwitchEnabled = GetBoolOr(object, L"autoSwitchEnabled", true);
    status.hasLocation = GetBoolOr(object, L"hasLocation", false);
    status.manualOverrideActive = GetBoolOr(object, L"manualOverrideActive", false);
    status.latitude = GetDoubleOr(object, L"latitude", 0.0);
    status.longitude = GetDoubleOr(object, L"longitude", 0.0);
    status.currentMode = GetStringOr(object, L"currentMode", L"unknown");
    status.desiredMode = GetStringOr(object, L"desiredMode", L"unknown");
    status.manualOverrideMode = GetStringOr(object, L"manualOverrideMode", L"");
    status.manualOverrideUntil = GetStringOr(object, L"manualOverrideUntil", L"");
    status.sunriseTime = GetStringOr(object, L"sunriseTime", L"");
    status.sunsetTime = GetStringOr(object, L"sunsetTime", L"");
    status.nextTransitionTime = GetStringOr(object, L"nextTransitionTime", L"");
    status.locationSource = GetStringOr(object, L"locationSource", L"none");
    status.lastLocationRefresh = GetStringOr(object, L"lastLocationRefresh", L"");
    status.lastAppliedTime = GetStringOr(object, L"lastAppliedTime", L"");
    status.hotkeyText = GetStringOr(object, L"hotkeyText", L"");
    status.lastError = GetStringOr(object, L"lastError", L"");
    return true;
}

bool SaveStatus(const Status& status, std::wstring* error) {
    JsonObject object;
    object.SetNamedValue(L"agentRunning", JsonValue::CreateBooleanValue(status.agentRunning));
    object.SetNamedValue(L"autoSwitchEnabled", JsonValue::CreateBooleanValue(status.autoSwitchEnabled));
    object.SetNamedValue(L"hasLocation", JsonValue::CreateBooleanValue(status.hasLocation));
    object.SetNamedValue(L"manualOverrideActive", JsonValue::CreateBooleanValue(status.manualOverrideActive));
    object.SetNamedValue(L"latitude", JsonValue::CreateNumberValue(status.latitude));
    object.SetNamedValue(L"longitude", JsonValue::CreateNumberValue(status.longitude));
    object.SetNamedValue(L"currentMode", JsonValue::CreateStringValue(status.currentMode));
    object.SetNamedValue(L"desiredMode", JsonValue::CreateStringValue(status.desiredMode));
    object.SetNamedValue(L"manualOverrideMode", JsonValue::CreateStringValue(status.manualOverrideMode));
    object.SetNamedValue(L"manualOverrideUntil", JsonValue::CreateStringValue(status.manualOverrideUntil));
    object.SetNamedValue(L"sunriseTime", JsonValue::CreateStringValue(status.sunriseTime));
    object.SetNamedValue(L"sunsetTime", JsonValue::CreateStringValue(status.sunsetTime));
    object.SetNamedValue(L"nextTransitionTime", JsonValue::CreateStringValue(status.nextTransitionTime));
    object.SetNamedValue(L"locationSource", JsonValue::CreateStringValue(status.locationSource));
    object.SetNamedValue(L"lastLocationRefresh", JsonValue::CreateStringValue(status.lastLocationRefresh));
    object.SetNamedValue(L"lastAppliedTime", JsonValue::CreateStringValue(status.lastAppliedTime));
    object.SetNamedValue(L"hotkeyText", JsonValue::CreateStringValue(status.hotkeyText));
    object.SetNamedValue(L"lastError", JsonValue::CreateStringValue(status.lastError));
    return WriteJsonFile(GetStatusPath(), object, error);
}

}  // namespace winxsw
