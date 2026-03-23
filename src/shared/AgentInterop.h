#pragma once

#include <Windows.h>

#include <filesystem>
#include <string>
#include <utility>
#include <vector>

#include "Settings.h"

namespace winxsw {

inline constexpr wchar_t kAgentWindowClassName[] = L"WinXSwThemeAgentWindow";
inline constexpr wchar_t kRunValueName[] = L"WinXSwThemeAgent";
inline constexpr UINT kAgentMessageReload = WM_APP + 101;
inline constexpr UINT kAgentMessageToggle = WM_APP + 102;

std::filesystem::path GetAgentExecutablePath();
std::filesystem::path GetConfigExecutablePath();

bool EnsureAutostartEnabled(bool enabled, std::wstring* error = nullptr);
bool IsAutostartEnabled();

std::wstring FormatHotkey(const HotkeySettings& hotkey);
UINT HotkeyModifiers(const HotkeySettings& hotkey);
std::vector<std::pair<UINT, std::wstring>> AvailableHotkeyKeys();

bool NotifyAgent(UINT message);
bool IsAgentRunning();
bool StopAgent(std::wstring* error = nullptr);
bool LaunchAgent(std::wstring* error = nullptr);

}  // namespace winxsw
