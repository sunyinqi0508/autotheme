#include "Location.h"

#include "Common.h"

#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Devices.Geolocation.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/base.h>

#include <winhttp.h>

#include <cmath>
#include <future>
#include <limits>
#include <optional>
#include <vector>

namespace winxsw {

using winrt::Windows::Data::Json::JsonObject;
using winrt::Windows::Devices::Geolocation::GeolocationAccessStatus;
using winrt::Windows::Devices::Geolocation::Geolocator;
using winrt::Windows::Devices::Geolocation::PositionAccuracy;

namespace {

struct IpHit {
    std::wstring provider;
    double latitude = 0.0;
    double longitude = 0.0;
};

struct ParsedUrl {
    INTERNET_SCHEME scheme = INTERNET_SCHEME_HTTPS;
    std::wstring host;
    std::wstring path;
    INTERNET_PORT port = INTERNET_DEFAULT_HTTPS_PORT;
};

std::optional<ParsedUrl> ParseUrl(const std::wstring& url) {
    URL_COMPONENTS components{};
    components.dwStructSize = sizeof(components);
    components.dwSchemeLength = static_cast<DWORD>(-1);
    components.dwHostNameLength = static_cast<DWORD>(-1);
    components.dwUrlPathLength = static_cast<DWORD>(-1);
    components.dwExtraInfoLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &components)) {
        return std::nullopt;
    }

    ParsedUrl parsed;
    parsed.scheme = components.nScheme;
    parsed.host.assign(components.lpszHostName, components.dwHostNameLength);
    parsed.path.assign(components.lpszUrlPath, components.dwUrlPathLength);
    if (components.dwExtraInfoLength > 0) {
        parsed.path.append(components.lpszExtraInfo, components.dwExtraInfoLength);
    }
    parsed.port = components.nPort;
    return parsed;
}

std::optional<std::wstring> HttpGetText(const std::wstring& url) {
    const auto parsed = ParseUrl(url);
    if (!parsed) {
        return std::nullopt;
    }

    HINTERNET session = WinHttpOpen(
        L"WinXSwTheme/1.0",
        WINHTTP_ACCESS_TYPE_NO_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS,
        0);
    if (session == nullptr) {
        return std::nullopt;
    }

    WinHttpSetTimeouts(session, 4000, 4000, 4000, 4000);
    HINTERNET connection = WinHttpConnect(session, parsed->host.c_str(), parsed->port, 0);
    if (connection == nullptr) {
        WinHttpCloseHandle(session);
        return std::nullopt;
    }

    const auto request = WinHttpOpenRequest(
        connection,
        L"GET",
        parsed->path.c_str(),
        nullptr,
        WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES,
        parsed->scheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0);
    if (request == nullptr) {
        WinHttpCloseHandle(connection);
        WinHttpCloseHandle(session);
        return std::nullopt;
    }

    const wchar_t* headers = L"Cache-Control: no-cache\r\nPragma: no-cache\r\n";
    std::optional<std::wstring> result;
    if (WinHttpSendRequest(request, headers, static_cast<DWORD>(-1), WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
        WinHttpReceiveResponse(request, nullptr)) {
        std::string body;
        while (true) {
            DWORD available = 0;
            if (!WinHttpQueryDataAvailable(request, &available) || available == 0) {
                break;
            }
            std::string chunk(available, '\0');
            DWORD read = 0;
            if (!WinHttpReadData(request, chunk.data(), available, &read) || read == 0) {
                break;
            }
            chunk.resize(read);
            body += chunk;
        }
        result = Utf8ToWide(body);
    }

    WinHttpCloseHandle(request);
    WinHttpCloseHandle(connection);
    WinHttpCloseHandle(session);
    return result;
}

double DistanceScore(double lat1, double lon1, double lat2, double lon2) {
    const auto dLat = lat1 - lat2;
    const auto dLon = lon1 - lon2;
    return std::sqrt((dLat * dLat) + (dLon * dLon));
}

std::optional<IpHit> QueryIpWhoIs() {
    const auto text = HttpGetText(L"https://ipwho.is/");
    if (!text) {
        return std::nullopt;
    }
    try {
        const auto object = JsonObject::Parse(text->c_str());
        if (!object.GetNamedBoolean(L"success", false)) {
            return std::nullopt;
        }
        return IpHit{
            L"ipwho.is",
            object.GetNamedNumber(L"latitude"),
            object.GetNamedNumber(L"longitude"),
        };
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<IpHit> QueryIpApiCo() {
    const auto text = HttpGetText(L"https://ipapi.co/json/");
    if (!text) {
        return std::nullopt;
    }
    try {
        const auto object = JsonObject::Parse(text->c_str());
        return IpHit{
            L"ipapi.co",
            object.GetNamedNumber(L"latitude"),
            object.GetNamedNumber(L"longitude"),
        };
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<IpHit> QueryIpInfo() {
    const auto text = HttpGetText(L"https://ipinfo.io/json");
    if (!text) {
        return std::nullopt;
    }
    try {
        const auto object = JsonObject::Parse(text->c_str());
        const auto loc = std::wstring(object.GetNamedString(L"loc").c_str());
        const auto comma = loc.find(L',');
        if (comma == std::wstring::npos) {
            return std::nullopt;
        }

        double latitude = 0.0;
        double longitude = 0.0;
        if (!ParseDouble(loc.substr(0, comma), latitude) ||
            !ParseDouble(loc.substr(comma + 1), longitude)) {
            return std::nullopt;
        }

        return IpHit{
            L"ipinfo.io",
            latitude,
            longitude,
        };
    } catch (...) {
        return std::nullopt;
    }
}

LocationResult TryGpsLocationCore() {
    LocationResult result;
    result.source = L"gps";
    try {
        const auto access = Geolocator::RequestAccessAsync().get();
        if (access != GeolocationAccessStatus::Allowed) {
            result.error = L"desktop location access was denied";
            return result;
        }

        Geolocator locator;
        locator.DesiredAccuracy(PositionAccuracy::High);
        const auto position = locator.GetGeopositionAsync(
            std::chrono::minutes(5),
            std::chrono::seconds(10)).get();
        const auto point = position.Coordinate().Point().Position();
        result.success = true;
        result.latitude = point.Latitude;
        result.longitude = point.Longitude;
        return result;
    } catch (const winrt::hresult_error& error) {
        result.error = error.message().c_str();
        return result;
    } catch (...) {
        result.error = L"GPS lookup failed";
        return result;
    }
}

LocationResult TryGpsLocation() {
    try {
        auto future = std::async(std::launch::async, [] {
            winrt::init_apartment(winrt::apartment_type::multi_threaded);
            return TryGpsLocationCore();
        });
        return future.get();
    } catch (...) {
        return {
            false,
            0.0,
            0.0,
            L"gps",
            L"failed to start the MTA geolocation worker",
        };
    }
}

LocationResult TryIpLocation() {
    std::vector<IpHit> hits;
    for (const auto& hit : {QueryIpWhoIs(), QueryIpApiCo(), QueryIpInfo()}) {
        if (hit) {
            hits.push_back(*hit);
        }
    }

    LocationResult result;
    result.source = L"ip";
    if (hits.empty()) {
        result.error = L"all direct no-proxy IP providers failed";
        return result;
    }

    auto best = hits.front();
    auto bestScore = std::numeric_limits<double>::max();
    for (const auto& candidate : hits) {
        double score = 0.0;
        for (const auto& other : hits) {
            score += DistanceScore(candidate.latitude, candidate.longitude, other.latitude, other.longitude);
        }
        if (score < bestScore) {
            best = candidate;
            bestScore = score;
        }
    }

    result.success = true;
    result.latitude = best.latitude;
    result.longitude = best.longitude;
    result.source = L"ip:" + best.provider + L" (direct/no-proxy)";
    if (hits.size() > 1) {
        result.source += L", consensus";
    }
    return result;
}

}  // namespace

LocationResult ResolveLocation(const Settings& settings) {
    if (settings.locationMode == LocationMode::Manual) {
        return {
            true,
            settings.manualLatitude,
            settings.manualLongitude,
            L"manual",
            L"",
        };
    }

    if (settings.allowGps) {
        auto gps = TryGpsLocation();
        if (gps.success) {
            return gps;
        }
        if (!settings.allowIpFallback) {
            return gps;
        }
    }

    if (settings.allowIpFallback) {
        return TryIpLocation();
    }

    return {
        false,
        0.0,
        0.0,
        L"none",
        L"all location providers are disabled",
    };
}

}  // namespace winxsw
