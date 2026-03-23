#include <Windows.h>

#include <chrono>
#include <memory>
#include <optional>
#include <string>

#include <winrt/base.h>

#include "AgentInterop.h"
#include "Common.h"
#include "Location.h"
#include "Settings.h"
#include "Solar.h"
#include "Theme.h"

namespace winxsw {

namespace {

constexpr UINT_PTR kTimerId = 1;
constexpr int kHotkeyId = 1;
constexpr wchar_t kMutexName[] = L"Local\\WinXSwThemeAgentMutex";

void AppendError(std::wstring& target, const std::wstring& message) {
    if (message.empty()) {
        return;
    }
    if (!target.empty()) {
        target += L" | ";
    }
    target += message;
}

class AgentApp {
public:
    explicit AgentApp(HINSTANCE instance)
        : instance_(instance) {
    }

    bool Initialize() {
        if (!LoadSettings(settings_, &lastError_)) {
            SaveSettings(settings_, nullptr);
        }

        const auto autostartOk = EnsureAutostartEnabled(settings_.autostartEnabled, &lastError_);
        if (!autostartOk) {
            AppendError(lastError_, L"autostart registration failed");
        }

        WNDCLASSEXW windowClass{};
        windowClass.cbSize = sizeof(windowClass);
        windowClass.lpfnWndProc = &AgentApp::WindowProc;
        windowClass.hInstance = instance_;
        windowClass.lpszClassName = kAgentWindowClassName;
        if (RegisterClassExW(&windowClass) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
            return false;
        }

        hwnd_ = CreateWindowExW(
            0,
            kAgentWindowClassName,
            L"Auto Theme",
            WS_OVERLAPPED,
            0,
            0,
            0,
            0,
            nullptr,
            nullptr,
            instance_,
            this);
        if (hwnd_ == nullptr) {
            return false;
        }

        RebindHotkey();
        ResetTimer();
        Evaluate(true, true);
        return true;
    }

    void Shutdown() {
        if (hotkeyRegistered_) {
            UnregisterHotKey(hwnd_, kHotkeyId);
            hotkeyRegistered_ = false;
        }
        status_.agentRunning = false;
        SaveStatus(status_, nullptr);
    }

    LRESULT HandleMessage(UINT message, WPARAM wParam, LPARAM lParam) {
        switch (message) {
        case WM_TIMER:
            Evaluate(false, false);
            return 0;
        case WM_HOTKEY:
            if (wParam == kHotkeyId) {
                Evaluate(false, false);
                ToggleNow();
            }
            return 0;
        case kAgentMessageReload:
            ReloadSettings();
            return 0;
        case kAgentMessageToggle:
            Evaluate(false, false);
            ToggleNow();
            return 0;
        case WM_QUERYENDSESSION:
        case WM_ENDSESSION:
            Shutdown();
            return 0;
        case WM_DESTROY:
            Shutdown();
            PostQuitMessage(0);
            return 0;
        default:
            return DefWindowProcW(hwnd_, message, wParam, lParam);
        }
    }

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
        if (message == WM_NCCREATE) {
            const auto create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
            const auto self = static_cast<AgentApp*>(create->lpCreateParams);
            self->hwnd_ = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            return TRUE;
        }

        const auto self = reinterpret_cast<AgentApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (self == nullptr) {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }
        return self->HandleMessage(message, wParam, lParam);
    }

    void ReloadSettings() {
        std::wstring error;
        Settings updated = settings_;
        if (!LoadSettings(updated, &error)) {
            lastError_ = error;
            PersistStatus();
            return;
        }

        settings_ = updated;
        EnsureAutostartEnabled(settings_.autostartEnabled, nullptr);
        RebindHotkey();
        ResetTimer();
        Evaluate(true, true);
    }

    void ResetTimer() {
        if (hwnd_ == nullptr) {
            return;
        }
        KillTimer(hwnd_, kTimerId);
        const auto interval = std::max(1, settings_.pollIntervalMinutes) * 60 * 1000;
        SetTimer(hwnd_, kTimerId, interval, nullptr);
    }

    void RebindHotkey() {
        if (hotkeyRegistered_) {
            UnregisterHotKey(hwnd_, kHotkeyId);
            hotkeyRegistered_ = false;
        }

        if (!settings_.hotkey.enabled || hwnd_ == nullptr) {
            return;
        }

        if (RegisterHotKey(hwnd_, kHotkeyId, HotkeyModifiers(settings_.hotkey), settings_.hotkey.virtualKey)) {
            hotkeyRegistered_ = true;
        } else {
            AppendError(lastError_, L"failed to register global hotkey " + FormatHotkey(settings_.hotkey));
        }
    }

    bool NeedsLocationRefresh(const std::chrono::system_clock::time_point& nowUtc) const {
        if (settings_.locationMode == LocationMode::Manual) {
            return true;
        }
        if (!location_.success || !lastLocationRefresh_.has_value()) {
            return true;
        }
        return nowUtc - *lastLocationRefresh_ >= std::chrono::minutes(settings_.locationRefreshMinutes);
    }

    void RefreshLocation(const std::chrono::system_clock::time_point& nowUtc) {
        location_ = ResolveLocation(settings_);
        if (location_.success) {
            lastLocationRefresh_ = nowUtc;
        } else {
            AppendError(lastError_, location_.error);
        }
    }

    ThemeMode DetermineDesiredMode(const std::chrono::system_clock::time_point& nowUtc) {
        if (!settings_.autoSwitchEnabled) {
            manualOverrideActive_ = false;
            return GetCurrentThemeMode();
        }

        if (manualOverrideActive_ && manualOverrideUntil_.has_value() && nowUtc >= *manualOverrideUntil_) {
            manualOverrideActive_ = false;
            manualOverrideUntil_.reset();
        }

        if (manualOverrideActive_) {
            return manualOverrideMode_;
        }

        if (schedule_.valid) {
            return schedule_.desiredMode;
        }

        return GetCurrentThemeMode();
    }

    void Evaluate(bool forceLocation, bool forceApply) {
        const auto nowUtc = UtcNow();
        lastError_.clear();

        if (forceLocation || NeedsLocationRefresh(nowUtc)) {
            RefreshLocation(nowUtc);
        }

        if (settings_.autoSwitchEnabled && location_.success) {
            schedule_ = ComputeSchedule(location_.latitude, location_.longitude, settings_, nowUtc);
            if (!schedule_.valid) {
                AppendError(lastError_, L"failed to compute sunrise and sunset for the current location");
            }
        } else {
            schedule_ = {};
        }

        const auto desiredMode = DetermineDesiredMode(nowUtc);
        const auto currentMode = GetCurrentThemeMode();
        if (forceApply || desiredMode != currentMode) {
            std::wstring error;
            if (ApplyThemeMode(desiredMode, &error)) {
                lastApplied_ = nowUtc;
                if (!ApplyWallpaperForMode(settings_, desiredMode, &error)) {
                    AppendError(lastError_, error);
                }
            } else {
                AppendError(lastError_, error);
            }
        }

        desiredMode_ = desiredMode;
        PersistStatus();
    }

    void ToggleNow() {
        const auto nowUtc = UtcNow();
        const auto currentMode = GetCurrentThemeMode();
        const auto targetMode = currentMode == ThemeMode::Light ? ThemeMode::Dark : ThemeMode::Light;

        if (settings_.autoSwitchEnabled && schedule_.valid && schedule_.nextTransition.has_value()) {
            manualOverrideActive_ = true;
            manualOverrideMode_ = targetMode;
            manualOverrideUntil_ = schedule_.nextTransition;
        } else {
            manualOverrideActive_ = false;
            manualOverrideUntil_.reset();
        }

        std::wstring error;
        if (ApplyThemeMode(targetMode, &error)) {
            desiredMode_ = targetMode;
            lastApplied_ = nowUtc;
            if (!ApplyWallpaperForMode(settings_, targetMode, &error)) {
                AppendError(lastError_, error);
            }
        } else {
            AppendError(lastError_, error);
        }
        PersistStatus();
    }

    void PersistStatus() {
        status_.agentRunning = true;
        status_.autoSwitchEnabled = settings_.autoSwitchEnabled;
        status_.hasLocation = location_.success;
        status_.latitude = location_.latitude;
        status_.longitude = location_.longitude;
        status_.currentMode = ThemeModeToString(GetCurrentThemeMode());
        status_.desiredMode = ThemeModeToString(desiredMode_);
        status_.manualOverrideActive = manualOverrideActive_;
        status_.manualOverrideMode = manualOverrideActive_ ? ThemeModeToString(manualOverrideMode_) : L"";
        status_.manualOverrideUntil = manualOverrideUntil_.has_value()
            ? FormatLocalTimestamp(*manualOverrideUntil_)
            : L"";
        status_.sunriseTime = schedule_.sunrise.has_value() ? FormatLocalTimestamp(*schedule_.sunrise) : L"";
        status_.sunsetTime = schedule_.sunset.has_value() ? FormatLocalTimestamp(*schedule_.sunset) : L"";
        status_.nextTransitionTime = schedule_.nextTransition.has_value()
            ? FormatLocalTimestamp(*schedule_.nextTransition)
            : L"";
        status_.locationSource = location_.source;
        status_.lastLocationRefresh = lastLocationRefresh_.has_value()
            ? FormatLocalTimestamp(*lastLocationRefresh_)
            : L"";
        status_.lastAppliedTime = lastApplied_.has_value() ? FormatLocalTimestamp(*lastApplied_) : L"";
        status_.hotkeyText = FormatHotkey(settings_.hotkey);
        status_.lastError = lastError_;
        SaveStatus(status_, nullptr);
    }

    HINSTANCE instance_ = nullptr;
    HWND hwnd_ = nullptr;
    Settings settings_{};
    Status status_{};
    LocationResult location_{};
    ScheduleResult schedule_{};
    ThemeMode desiredMode_ = ThemeMode::Light;
    ThemeMode manualOverrideMode_ = ThemeMode::Dark;
    std::wstring lastError_;
    bool hotkeyRegistered_ = false;
    bool manualOverrideActive_ = false;
    std::optional<std::chrono::system_clock::time_point> manualOverrideUntil_;
    std::optional<std::chrono::system_clock::time_point> lastLocationRefresh_;
    std::optional<std::chrono::system_clock::time_point> lastApplied_;
};

}  // namespace

}  // namespace winxsw

int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    auto mutex = CreateMutexW(nullptr, FALSE, winxsw::kMutexName);
    if (mutex == nullptr) {
        return 1;
    }
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        CloseHandle(mutex);
        return 0;
    }

    winrt::init_apartment(winrt::apartment_type::single_threaded);

    winxsw::AgentApp app(instance);
    if (!app.Initialize()) {
        CloseHandle(mutex);
        return 1;
    }

    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0)) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    CloseHandle(mutex);
    return static_cast<int>(message.wParam);
}
