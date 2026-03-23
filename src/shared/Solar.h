#pragma once

#include <chrono>
#include <optional>

#include "Settings.h"

namespace winxsw {

struct ScheduleResult {
    bool valid = false;
    bool alwaysLight = false;
    bool alwaysDark = false;
    ThemeMode desiredMode = ThemeMode::Light;
    std::optional<std::chrono::system_clock::time_point> sunrise;
    std::optional<std::chrono::system_clock::time_point> sunset;
    std::optional<std::chrono::system_clock::time_point> nextTransition;
};

ScheduleResult ComputeSchedule(
    double latitude,
    double longitude,
    const Settings& settings,
    const std::chrono::system_clock::time_point& nowUtc);

}  // namespace winxsw
