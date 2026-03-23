#pragma once

#include <string>

#include "Settings.h"

namespace winxsw {

struct LocationResult {
    bool success = false;
    double latitude = 0.0;
    double longitude = 0.0;
    std::wstring source = L"none";
    std::wstring error;
};

LocationResult ResolveLocation(const Settings& settings);

}  // namespace winxsw
