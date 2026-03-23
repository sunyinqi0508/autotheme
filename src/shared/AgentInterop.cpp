#include "AgentInterop.h"

#include "Common.h"

#include <shellapi.h>

namespace winxsw {

namespace {

constexpr wchar_t kRunKey[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";

std::wstring QuotePath(const std::filesystem::path& path) {
    return L"\"" + path.wstring() + L"\"";
}

std::wstring KeyName(UINT virtualKey) {
    if (virtualKey >= 'A' && virtualKey <= 'Z') {
        return std::wstring(1, static_cast<wchar_t>(virtualKey));
    }
    if (virtualKey >= '0' && virtualKey <= '9') {
        return std::wstring(1, static_cast<wchar_t>(virtualKey));
    }
    if (virtualKey >= VK_F1 && virtualKey <= VK_F12) {
        return L"F" + std::to_wstring((virtualKey - VK_F1) + 1);
    }
    if (virtualKey == VK_SPACE) {
        return L"Space";
    }
    return L"VK " + std::to_wstring(virtualKey);
}

}  // namespace

std::filesystem::path GetAgentExecutablePath() {
    return GetExecutableDirectory() / L"winxsw-agent.exe";
}

std::filesystem::path GetConfigExecutablePath() {
    return GetExecutableDirectory() / L"winxsw-config.exe";
}

bool EnsureAutostartEnabled(bool enabled, std::wstring* error) {
    HKEY key = nullptr;
    if (RegCreateKeyExW(
            HKEY_CURRENT_USER,
            kRunKey,
            0,
            nullptr,
            0,
            KEY_SET_VALUE | KEY_QUERY_VALUE,
            nullptr,
            &key,
            nullptr) != ERROR_SUCCESS) {
        if (error != nullptr) {
            *error = L"failed to open the HKCU Run registry key";
        }
        return false;
    }

    bool success = true;
    if (enabled) {
        const auto command = QuotePath(GetAgentExecutablePath());
        success = RegSetValueExW(
            key,
            kRunValueName,
            0,
            REG_SZ,
            reinterpret_cast<const BYTE*>(command.c_str()),
            static_cast<DWORD>((command.size() + 1) * sizeof(wchar_t))) == ERROR_SUCCESS;
    } else {
        const auto status = RegDeleteValueW(key, kRunValueName);
        success = status == ERROR_SUCCESS || status == ERROR_FILE_NOT_FOUND;
    }

    RegCloseKey(key);
    if (!success && error != nullptr) {
        *error = L"failed to update the autostart registry value";
    }
    return success;
}

bool IsAutostartEnabled() {
    HKEY key = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kRunKey, 0, KEY_QUERY_VALUE, &key) != ERROR_SUCCESS) {
        return false;
    }

    wchar_t buffer[2048];
    DWORD size = sizeof(buffer);
    const auto status = RegQueryValueExW(
        key,
        kRunValueName,
        nullptr,
        nullptr,
        reinterpret_cast<BYTE*>(buffer),
        &size);
    RegCloseKey(key);
    return status == ERROR_SUCCESS;
}

std::wstring FormatHotkey(const HotkeySettings& hotkey) {
    if (!hotkey.enabled) {
        return L"Disabled";
    }
    std::wstring value;
    if (hotkey.ctrl) {
        value += L"Ctrl+";
    }
    if (hotkey.alt) {
        value += L"Alt+";
    }
    if (hotkey.shift) {
        value += L"Shift+";
    }
    if (hotkey.win) {
        value += L"Win+";
    }
    value += KeyName(hotkey.virtualKey);
    return value;
}

UINT HotkeyModifiers(const HotkeySettings& hotkey) {
    UINT modifiers = MOD_NOREPEAT;
    if (hotkey.ctrl) {
        modifiers |= MOD_CONTROL;
    }
    if (hotkey.alt) {
        modifiers |= MOD_ALT;
    }
    if (hotkey.shift) {
        modifiers |= MOD_SHIFT;
    }
    if (hotkey.win) {
        modifiers |= MOD_WIN;
    }
    return modifiers;
}

std::vector<std::pair<UINT, std::wstring>> AvailableHotkeyKeys() {
    std::vector<std::pair<UINT, std::wstring>> keys;
    for (UINT value = 'A'; value <= 'Z'; ++value) {
        keys.emplace_back(value, KeyName(value));
    }
    for (UINT value = '0'; value <= '9'; ++value) {
        keys.emplace_back(value, KeyName(value));
    }
    for (UINT value = VK_F1; value <= VK_F12; ++value) {
        keys.emplace_back(value, KeyName(value));
    }
    keys.emplace_back(VK_SPACE, L"Space");
    return keys;
}

bool NotifyAgent(UINT message) {
    const auto window = FindWindowW(kAgentWindowClassName, nullptr);
    if (window == nullptr) {
        return false;
    }
    return PostMessageW(window, message, 0, 0) != FALSE;
}

bool IsAgentRunning() {
    return FindWindowW(kAgentWindowClassName, nullptr) != nullptr;
}

bool StopAgent(std::wstring* error) {
    const auto window = FindWindowW(kAgentWindowClassName, nullptr);
    if (window == nullptr) {
        if (error != nullptr) {
            *error = L"winxsw-agent.exe is not running";
        }
        return false;
    }

    DWORD_PTR ignored = 0;
    if (!SendMessageTimeoutW(window, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 2000, &ignored)) {
        if (error != nullptr) {
            *error = L"failed to signal winxsw-agent.exe to close";
        }
        return false;
    }

    for (int attempt = 0; attempt < 20; ++attempt) {
        if (!IsAgentRunning()) {
            return true;
        }
        Sleep(100);
    }

    if (error != nullptr) {
        *error = L"winxsw-agent.exe did not exit in time";
    }
    return false;
}

bool LaunchAgent(std::wstring* error) {
    const auto executable = GetAgentExecutablePath();
    const auto result = ShellExecuteW(
        nullptr,
        L"open",
        executable.c_str(),
        nullptr,
        executable.parent_path().c_str(),
        SW_HIDE);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
        if (error != nullptr) {
            *error = L"failed to launch winxsw-agent.exe";
        }
        return false;
    }
    return true;
}

}  // namespace winxsw
