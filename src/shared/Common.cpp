#include "Common.h"

#include <ShlObj.h>

#include <cstdint>
#include <cwchar>
#include <fstream>
#include <iomanip>
#include <sstream>

namespace winxsw {

namespace {

std::chrono::system_clock::time_point FileTimeToTimePoint(const FILETIME& fileTime) {
    ULARGE_INTEGER value{};
    value.LowPart = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    constexpr std::uint64_t kUnixEpochTicks = 116444736000000000ULL;
    const auto ticks = static_cast<std::int64_t>(value.QuadPart - kUnixEpochTicks);
    return std::chrono::system_clock::time_point{} +
        std::chrono::duration_cast<std::chrono::system_clock::duration>(
            std::chrono::nanoseconds(ticks * 100));
}

FILETIME TimePointToFileTime(const std::chrono::system_clock::time_point& timePoint) {
    constexpr std::uint64_t kUnixEpochTicks = 116444736000000000ULL;
    const auto nanoseconds = std::chrono::duration_cast<std::chrono::nanoseconds>(
        timePoint.time_since_epoch());
    const auto ticks = static_cast<std::uint64_t>(nanoseconds.count() / 100) + kUnixEpochTicks;
    ULARGE_INTEGER value{};
    value.QuadPart = ticks;
    FILETIME result{};
    result.dwLowDateTime = value.LowPart;
    result.dwHighDateTime = value.HighPart;
    return result;
}

long long CivilMinutesFromSystemTime(const SYSTEMTIME& value) {
    using namespace std::chrono;
    const auto days = sys_days(
        year{value.wYear} / month{value.wMonth} / day{value.wDay});
    return duration_cast<minutes>(days.time_since_epoch()).count() +
        (value.wHour * 60LL) + value.wMinute;
}

}  // namespace

std::filesystem::path GetModulePath(HMODULE module) {
    std::wstring buffer(MAX_PATH, L'\0');
    while (true) {
        const auto copied = GetModuleFileNameW(module, buffer.data(), static_cast<DWORD>(buffer.size()));
        if (copied == 0) {
            return {};
        }
        if (copied < buffer.size() - 1) {
            buffer.resize(copied);
            return std::filesystem::path(buffer);
        }
        buffer.resize(buffer.size() * 2);
    }
}

std::filesystem::path GetExecutableDirectory(HMODULE module) {
    auto modulePath = GetModulePath(module);
    return modulePath.empty() ? std::filesystem::path{} : modulePath.parent_path();
}

std::filesystem::path GetDataDirectory() {
    PWSTR rawPath = nullptr;
    std::filesystem::path result;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &rawPath)) &&
        rawPath != nullptr) {
        result = std::filesystem::path(rawPath) / L"WinXSwTheme";
        CoTaskMemFree(rawPath);
    }
    if (result.empty()) {
        result = GetExecutableDirectory() / L"data";
    }
    std::error_code ec;
    std::filesystem::create_directories(result, ec);
    return result;
}

std::optional<std::wstring> ReadUtf8File(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream.is_open()) {
        return std::nullopt;
    }
    const std::string buffer(
        (std::istreambuf_iterator<char>(stream)),
        std::istreambuf_iterator<char>());
    return Utf8ToWide(buffer);
}

bool WriteUtf8File(const std::filesystem::path& path, const std::wstring& text) {
    std::error_code ec;
    std::filesystem::create_directories(path.parent_path(), ec);
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    if (!stream.is_open()) {
        return false;
    }
    const auto utf8 = WideToUtf8(text);
    stream.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    return stream.good();
}

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    const auto required = MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        text.data(),
        static_cast<int>(text.size()),
        nullptr,
        0);
    if (required <= 0) {
        return {};
    }
    std::wstring result(required, L'\0');
    MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        text.data(),
        static_cast<int>(text.size()),
        result.data(),
        required);
    return result;
}

std::string WideToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }
    const auto required = WideCharToMultiByte(
        CP_UTF8,
        0,
        text.data(),
        static_cast<int>(text.size()),
        nullptr,
        0,
        nullptr,
        nullptr);
    std::string result(required, '\0');
    WideCharToMultiByte(
        CP_UTF8,
        0,
        text.data(),
        static_cast<int>(text.size()),
        result.data(),
        required,
        nullptr,
        nullptr);
    return result;
}

std::wstring Trim(const std::wstring& value) {
    const auto start = value.find_first_not_of(L" \t\r\n");
    if (start == std::wstring::npos) {
        return {};
    }
    const auto end = value.find_last_not_of(L" \t\r\n");
    return value.substr(start, end - start + 1);
}

std::wstring ToUpperAscii(std::wstring value) {
    for (auto& ch : value) {
        if (ch >= L'a' && ch <= L'z') {
            ch = static_cast<wchar_t>(ch - L'a' + L'A');
        }
    }
    return value;
}

std::wstring FormatDouble(double value, int precision) {
    std::wostringstream stream;
    stream << std::fixed << std::setprecision(precision) << value;
    return stream.str();
}

std::chrono::system_clock::time_point UtcNow() {
    FILETIME fileTime{};
    GetSystemTimeAsFileTime(&fileTime);
    return FileTimeToTimePoint(fileTime);
}

std::optional<std::chrono::system_clock::time_point> LocalDateTimeToUtc(
    int year,
    int month,
    int day,
    int hour,
    int minute,
    int second) {
    SYSTEMTIME localTime{};
    localTime.wYear = static_cast<WORD>(year);
    localTime.wMonth = static_cast<WORD>(month);
    localTime.wDay = static_cast<WORD>(day);
    localTime.wHour = static_cast<WORD>(hour);
    localTime.wMinute = static_cast<WORD>(minute);
    localTime.wSecond = static_cast<WORD>(second);

    SYSTEMTIME utcTime{};
    if (!TzSpecificLocalTimeToSystemTime(nullptr, &localTime, &utcTime)) {
        return std::nullopt;
    }
    FILETIME utcFileTime{};
    if (!SystemTimeToFileTime(&utcTime, &utcFileTime)) {
        return std::nullopt;
    }
    return FileTimeToTimePoint(utcFileTime);
}

std::optional<SYSTEMTIME> UtcTimePointToLocalSystemTime(
    const std::chrono::system_clock::time_point& value) {
    const auto utcFileTime = TimePointToFileTime(value);
    SYSTEMTIME utcTime{};
    if (!FileTimeToSystemTime(&utcFileTime, &utcTime)) {
        return std::nullopt;
    }
    SYSTEMTIME localTime{};
    if (!SystemTimeToTzSpecificLocalTime(nullptr, &utcTime, &localTime)) {
        return std::nullopt;
    }
    return localTime;
}

std::wstring FormatLocalTimestamp(const std::chrono::system_clock::time_point& value) {
    const auto local = UtcTimePointToLocalSystemTime(value);
    if (!local) {
        return L"";
    }
    wchar_t buffer[64];
    swprintf_s(
        buffer,
        L"%04u-%02u-%02u %02u:%02u:%02u",
        local->wYear,
        local->wMonth,
        local->wDay,
        local->wHour,
        local->wMinute,
        local->wSecond);
    return buffer;
}

std::wstring FormatLocalTimeOnly(const std::chrono::system_clock::time_point& value) {
    const auto local = UtcTimePointToLocalSystemTime(value);
    if (!local) {
        return L"";
    }
    wchar_t buffer[16];
    swprintf_s(buffer, L"%02u:%02u", local->wHour, local->wMinute);
    return buffer;
}

int GetDayOfYear(int yearValue, int monthValue, int dayValue) {
    using namespace std::chrono;
    const auto current = sys_days(
        year{yearValue} / month{static_cast<unsigned>(monthValue)} / day{static_cast<unsigned>(dayValue)});
    const auto first = sys_days(year{yearValue} / January / day{1});
    return static_cast<int>((current - first).count()) + 1;
}

int GetTimezoneOffsetMinutesForLocalDate(int year, int month, int day) {
    SYSTEMTIME localNoon{};
    localNoon.wYear = static_cast<WORD>(year);
    localNoon.wMonth = static_cast<WORD>(month);
    localNoon.wDay = static_cast<WORD>(day);
    localNoon.wHour = 12;

    SYSTEMTIME utcNoon{};
    if (!TzSpecificLocalTimeToSystemTime(nullptr, &localNoon, &utcNoon)) {
        return 0;
    }
    return static_cast<int>(CivilMinutesFromSystemTime(localNoon) - CivilMinutesFromSystemTime(utcNoon));
}

bool ParseDouble(const std::wstring& text, double& value) {
    const auto trimmed = Trim(text);
    if (trimmed.empty()) {
        return false;
    }
    wchar_t* end = nullptr;
    value = std::wcstod(trimmed.c_str(), &end);
    return end != nullptr && *end == L'\0';
}

bool ParseInt(const std::wstring& text, int& value) {
    const auto trimmed = Trim(text);
    if (trimmed.empty()) {
        return false;
    }
    wchar_t* end = nullptr;
    const auto parsed = std::wcstol(trimmed.c_str(), &end, 10);
    if (end == nullptr || *end != L'\0') {
        return false;
    }
    value = static_cast<int>(parsed);
    return true;
}

}  // namespace winxsw
