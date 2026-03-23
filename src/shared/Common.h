#pragma once

#include <Windows.h>

#include <chrono>
#include <filesystem>
#include <optional>
#include <string>

namespace winxsw {

std::filesystem::path GetModulePath(HMODULE module = nullptr);
std::filesystem::path GetExecutableDirectory(HMODULE module = nullptr);
std::filesystem::path GetDataDirectory();

std::optional<std::wstring> ReadUtf8File(const std::filesystem::path& path);
bool WriteUtf8File(const std::filesystem::path& path, const std::wstring& text);

std::wstring Utf8ToWide(const std::string& text);
std::string WideToUtf8(const std::wstring& text);

std::wstring Trim(const std::wstring& value);
std::wstring ToUpperAscii(std::wstring value);
std::wstring FormatDouble(double value, int precision = 6);

std::chrono::system_clock::time_point UtcNow();
std::optional<std::chrono::system_clock::time_point> LocalDateTimeToUtc(
    int year,
    int month,
    int day,
    int hour,
    int minute,
    int second);
std::optional<SYSTEMTIME> UtcTimePointToLocalSystemTime(
    const std::chrono::system_clock::time_point& value);

std::wstring FormatLocalTimestamp(const std::chrono::system_clock::time_point& value);
std::wstring FormatLocalTimeOnly(const std::chrono::system_clock::time_point& value);
int GetDayOfYear(int year, int month, int day);
int GetTimezoneOffsetMinutesForLocalDate(int year, int month, int day);

bool ParseDouble(const std::wstring& text, double& value);
bool ParseInt(const std::wstring& text, int& value);

}  // namespace winxsw
