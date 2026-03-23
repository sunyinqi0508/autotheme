#include "Solar.h"

#include "Common.h"

#include <cmath>

namespace winxsw {

namespace {

constexpr double kPi = 3.14159265358979323846;
constexpr double kZenith = 90.833;

double DegToRad(double value) {
    return value * kPi / 180.0;
}

double RadToDeg(double value) {
    return value * 180.0 / kPi;
}

double NormalizeDegrees(double value) {
    while (value < 0.0) {
        value += 360.0;
    }
    while (value >= 360.0) {
        value -= 360.0;
    }
    return value;
}

double NormalizeHours(double value) {
    while (value < 0.0) {
        value += 24.0;
    }
    while (value >= 24.0) {
        value -= 24.0;
    }
    return value;
}

struct SolarEventResult {
    bool occurs = false;
    bool alwaysLight = false;
    bool alwaysDark = false;
    std::optional<std::chrono::system_clock::time_point> value;
};

struct LocalDate {
    int year = 0;
    int month = 0;
    int day = 0;
};

LocalDate GetLocalDate(const std::chrono::system_clock::time_point& nowUtc) {
    LocalDate result;
    const auto local = UtcTimePointToLocalSystemTime(nowUtc);
    if (local) {
        result.year = local->wYear;
        result.month = local->wMonth;
        result.day = local->wDay;
    }
    return result;
}

SolarEventResult ComputeSolarEventUtc(
    int yearValue,
    int monthValue,
    int dayValue,
    double latitude,
    double longitude,
    bool sunrise,
    int offsetMinutes) {
    const auto dayOfYear = GetDayOfYear(yearValue, monthValue, dayValue);
    const auto longitudeHour = longitude / 15.0;
    const auto approximateTime = dayOfYear + (((sunrise ? 6.0 : 18.0) - longitudeHour) / 24.0);
    const auto meanAnomaly = (0.9856 * approximateTime) - 3.289;

    auto trueLongitude = meanAnomaly +
        (1.916 * std::sin(DegToRad(meanAnomaly))) +
        (0.020 * std::sin(2.0 * DegToRad(meanAnomaly))) +
        282.634;
    trueLongitude = NormalizeDegrees(trueLongitude);

    auto rightAscension = RadToDeg(std::atan(0.91764 * std::tan(DegToRad(trueLongitude))));
    rightAscension = NormalizeDegrees(rightAscension);

    const auto longitudeQuadrant = std::floor(trueLongitude / 90.0) * 90.0;
    const auto rightAscensionQuadrant = std::floor(rightAscension / 90.0) * 90.0;
    rightAscension += longitudeQuadrant - rightAscensionQuadrant;
    rightAscension /= 15.0;

    const auto sinDeclination = 0.39782 * std::sin(DegToRad(trueLongitude));
    const auto cosDeclination = std::cos(std::asin(sinDeclination));
    const auto cosHour = (std::cos(DegToRad(kZenith)) -
        (sinDeclination * std::sin(DegToRad(latitude)))) /
        (cosDeclination * std::cos(DegToRad(latitude)));

    SolarEventResult result;
    if (cosHour > 1.0) {
        result.alwaysDark = true;
        return result;
    }
    if (cosHour < -1.0) {
        result.alwaysLight = true;
        return result;
    }

    auto hourAngle = sunrise
        ? (360.0 - RadToDeg(std::acos(cosHour)))
        : RadToDeg(std::acos(cosHour));
    hourAngle /= 15.0;

    const auto localMeanTime = hourAngle + rightAscension - (0.06571 * approximateTime) - 6.622;
    const auto utcHour = NormalizeHours(localMeanTime - longitudeHour);
    const auto offsetHours = static_cast<double>(GetTimezoneOffsetMinutesForLocalDate(yearValue, monthValue, dayValue)) / 60.0;
    auto localHours = utcHour + offsetHours;
    localHours += static_cast<double>(offsetMinutes) / 60.0;

    int dayAdjustment = 0;
    while (localHours < 0.0) {
        localHours += 24.0;
        --dayAdjustment;
    }
    while (localHours >= 24.0) {
        localHours -= 24.0;
        ++dayAdjustment;
    }

    using namespace std::chrono;
    const auto baseDate = sys_days(
        year{yearValue} / month{static_cast<unsigned>(monthValue)} / day{static_cast<unsigned>(dayValue)}) +
        days(dayAdjustment);
    const year_month_day adjustedDate{baseDate};
    const auto totalMinutes = localHours * 60.0;
    const auto wholeMinutes = static_cast<int>(std::floor(totalMinutes));
    const auto seconds = static_cast<int>(std::round((totalMinutes - wholeMinutes) * 60.0));

    result.value = LocalDateTimeToUtc(
        static_cast<int>(adjustedDate.year()),
        static_cast<unsigned>(adjustedDate.month()),
        static_cast<unsigned>(adjustedDate.day()),
        wholeMinutes / 60,
        wholeMinutes % 60,
        seconds);
    result.occurs = result.value.has_value();
    return result;
}

}  // namespace

ScheduleResult ComputeSchedule(
    double latitude,
    double longitude,
    const Settings& settings,
    const std::chrono::system_clock::time_point& nowUtc) {
    ScheduleResult result;
    const auto today = GetLocalDate(nowUtc);
    if (today.year == 0) {
        return result;
    }

    const auto sunrise = ComputeSolarEventUtc(
        today.year,
        today.month,
        today.day,
        latitude,
        longitude,
        true,
        settings.sunriseOffsetMinutes);
    const auto sunset = ComputeSolarEventUtc(
        today.year,
        today.month,
        today.day,
        latitude,
        longitude,
        false,
        settings.sunsetOffsetMinutes);

    if (sunrise.alwaysDark || sunset.alwaysDark) {
        result.valid = true;
        result.alwaysDark = true;
        result.desiredMode = ThemeMode::Dark;
        return result;
    }
    if (sunrise.alwaysLight || sunset.alwaysLight) {
        result.valid = true;
        result.alwaysLight = true;
        result.desiredMode = ThemeMode::Light;
        return result;
    }
    if (!sunrise.value || !sunset.value) {
        return result;
    }

    result.valid = true;
    result.sunrise = sunrise.value;
    result.sunset = sunset.value;

    using namespace std::chrono;
    const auto tomorrowBase = sys_days(
        year{today.year} / month{static_cast<unsigned>(today.month)} / day{static_cast<unsigned>(today.day)}) +
        days(1);
    const year_month_day tomorrow{tomorrowBase};
    const auto tomorrowSunrise = ComputeSolarEventUtc(
        static_cast<int>(tomorrow.year()),
        static_cast<unsigned>(tomorrow.month()),
        static_cast<unsigned>(tomorrow.day()),
        latitude,
        longitude,
        true,
        settings.sunriseOffsetMinutes);

    if (nowUtc < *sunrise.value) {
        result.desiredMode = ThemeMode::Dark;
        result.nextTransition = sunrise.value;
    } else if (nowUtc < *sunset.value) {
        result.desiredMode = ThemeMode::Light;
        result.nextTransition = sunset.value;
    } else {
        result.desiredMode = ThemeMode::Dark;
        result.nextTransition = tomorrowSunrise.value;
    }

    return result;
}

}  // namespace winxsw
