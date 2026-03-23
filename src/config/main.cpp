#include <Windows.h>
#include <windowsx.h>

#include <CommCtrl.h>
#include <CommDlg.h>
#include <dwmapi.h>
#include <Uxtheme.h>
#include <vssym32.h>
#include <ObjIdl.h>
#include <gdiplus.h>

#include <algorithm>
#include <filesystem>
#include <string>
#include <vector>

#include <shellapi.h>
#include <winrt/base.h>

#include "AgentInterop.h"
#include "Common.h"
#include "Settings.h"
#include "Theme.h"

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "Gdiplus.lib")
#pragma comment(lib, "UxTheme.lib")

namespace winxsw {

namespace {

constexpr UINT_PTR kRefreshTimerId = 1;
constexpr int kTitleBarHeight = 56;
constexpr int kContentTopOffset = -24;
constexpr int kWindowWidth = 924;
constexpr int kWindowHeight = 856;
constexpr int kEditHeight = 26;
constexpr int kCheckboxHeight = 22;
constexpr int kTitleButtonHeight = 34;

struct UiPalette {
    bool dark = false;
    COLORREF windowBackground = RGB(244, 247, 252);
    COLORREF windowBorder = RGB(178, 188, 204);
    COLORREF cardBackground = RGB(255, 255, 255);
    COLORREF cardBorder = RGB(222, 229, 239);
    COLORREF inputBackground = RGB(251, 253, 255);
    COLORREF titleBarBackground = RGB(248, 251, 255);
    COLORREF titleBarBorder = RGB(224, 230, 240);
    COLORREF titleButtonHover = RGB(233, 239, 248);
    COLORREF titleButtonPressed = RGB(220, 228, 240);
    COLORREF closeButtonHover = RGB(232, 17, 35);
    COLORREF closeButtonPressed = RGB(196, 12, 28);
    COLORREF primaryText = RGB(23, 32, 48);
    COLORREF secondaryText = RGB(96, 107, 126);
    COLORREF accent = RGB(14, 102, 184);
    COLORREF actionButtonBackground = RGB(236, 242, 249);
    COLORREF actionButtonHover = RGB(226, 234, 244);
    COLORREF actionButtonPressed = RGB(214, 223, 236);
    COLORREF actionButtonBorder = RGB(214, 223, 236);
    COLORREF actionButtonDisabled = RGB(241, 244, 248);
    COLORREF actionButtonText = RGB(23, 32, 48);
    COLORREF actionButtonDisabledText = RGB(144, 154, 172);
    COLORREF defaultButtonBackground = RGB(14, 102, 184);
    COLORREF defaultButtonHover = RGB(30, 118, 200);
    COLORREF defaultButtonPressed = RGB(8, 86, 160);
    COLORREF defaultButtonText = RGB(255, 255, 255);
    COLORREF modifierText = RGB(96, 107, 126);
    COLORREF modifierHoverText = RGB(14, 102, 184);
    COLORREF modifierCheckedText = RGB(23, 32, 48);
    COLORREF modifierCheckedHoverText = RGB(8, 86, 160);
    COLORREF modifierDisabledText = RGB(144, 154, 172);
};

UiPalette MakeUiPalette(ThemeMode mode) {
    if (mode == ThemeMode::Dark) {
        UiPalette palette;
        palette.dark = true;
        palette.windowBackground = RGB(18, 20, 24);
        palette.windowBorder = RGB(66, 72, 84);
        palette.cardBackground = RGB(28, 31, 37);
        palette.cardBorder = RGB(52, 58, 68);
        palette.inputBackground = RGB(22, 25, 30);
        palette.titleBarBackground = RGB(24, 27, 32);
        palette.titleBarBorder = RGB(56, 62, 72);
        palette.titleButtonHover = RGB(56, 62, 74);
        palette.titleButtonPressed = RGB(70, 76, 90);
        palette.closeButtonHover = RGB(232, 17, 35);
        palette.closeButtonPressed = RGB(196, 12, 28);
        palette.primaryText = RGB(236, 240, 248);
        palette.secondaryText = RGB(160, 170, 186);
        palette.accent = RGB(76, 166, 255);
        palette.actionButtonBackground = RGB(45, 50, 58);
        palette.actionButtonHover = RGB(58, 64, 74);
        palette.actionButtonPressed = RGB(68, 74, 86);
        palette.actionButtonBorder = RGB(78, 85, 97);
        palette.actionButtonDisabled = RGB(35, 39, 45);
        palette.actionButtonText = RGB(236, 240, 248);
        palette.actionButtonDisabledText = RGB(118, 126, 140);
        palette.defaultButtonBackground = palette.accent;
        palette.defaultButtonHover = RGB(95, 180, 255);
        palette.defaultButtonPressed = RGB(54, 144, 236);
        palette.defaultButtonText = RGB(12, 18, 28);
        palette.modifierText = RGB(160, 170, 186);
        palette.modifierHoverText = RGB(120, 190, 255);
        palette.modifierCheckedText = RGB(236, 240, 248);
        palette.modifierCheckedHoverText = RGB(186, 224, 255);
        palette.modifierDisabledText = RGB(118, 126, 140);
        return palette;
    }

    return {};
}

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

enum ControlId : int {
    IdAutoSwitch = 1001,
    IdAutostart,
    IdLocationAuto,
    IdLocationManual,
    IdAllowGps,
    IdAllowIp,
    IdLatitude,
    IdLongitude,
    IdSunriseOffset,
    IdSunsetOffset,
    IdPollInterval,
    IdLocationRefresh,
    IdHotkeyEnable,
    IdHotkeyCtrl,
    IdHotkeyAlt,
    IdHotkeyShift,
    IdHotkeyWin,
    IdHotkeyKey,
    IdConfigHotkeyEnable,
    IdConfigHotkeyCtrl,
    IdConfigHotkeyAlt,
    IdConfigHotkeyShift,
    IdConfigHotkeyWin,
    IdConfigHotkeyKey,
    IdStatusBox,
    IdLightWallpaper,
    IdDarkWallpaper,
    IdSave,
    IdReload,
    IdToggle,
    IdLaunch,
    IdTerminateAgent,
    IdOpenFolder,
    IdBrowseLightWallpaper,
    IdBrowseDarkWallpaper,
};

std::wstring ReadControlText(HWND control) {
    const auto length = GetWindowTextLengthW(control);
    std::wstring buffer(length + 1, L'\0');
    GetWindowTextW(control, buffer.data(), length + 1);
    buffer.resize(length);
    return buffer;
}

RECT MakeRect(int left, int top, int width, int height) {
    return RECT{left, top, left + width, top + height};
}

std::wstring EmptyToDash(const std::wstring& value) {
    return value.empty() ? L"-" : value;
}

std::wstring YesNo(bool value) {
    return value ? L"Yes" : L"No";
}

void AppendStatusLine(std::wstring& text, const wchar_t* label, const std::wstring& value) {
    text += label;
    text += L": ";
    text += EmptyToDash(value);
    text += L"\r\n";
}

bool ContainsHandle(const std::vector<HWND>& windows, HWND target) {
    return std::find(windows.begin(), windows.end(), target) != windows.end();
}

HFONT CreateUiFont(HWND window, int points, int weight, const wchar_t* faceName) {
    const auto dpi = static_cast<int>(GetDpiForWindow(window));
    return CreateFontW(
        -MulDiv(points, dpi, 72),
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_NATURAL_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        faceName);
}

Gdiplus::Color ToGdiplusColor(COLORREF color, BYTE alpha = 255) {
    return Gdiplus::Color(alpha, GetRValue(color), GetGValue(color), GetBValue(color));
}

void ConfigureGraphics(Gdiplus::Graphics& graphics) {
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
}

void AddRoundedRect(Gdiplus::GraphicsPath& path, const RECT& rect, Gdiplus::REAL radius) {
    const auto widthPixels = std::max<LONG>(1, rect.right - rect.left - 1);
    const auto heightPixels = std::max<LONG>(1, rect.bottom - rect.top - 1);
    const auto width = static_cast<Gdiplus::REAL>(widthPixels);
    const auto height = static_cast<Gdiplus::REAL>(heightPixels);
    const auto x = static_cast<Gdiplus::REAL>(rect.left);
    const auto y = static_cast<Gdiplus::REAL>(rect.top);
    const auto diameter = std::min(radius * 2.0f, std::min(width, height));

    Gdiplus::RectF arc(x, y, diameter, diameter);
    path.AddArc(arc, 180.0f, 90.0f);
    arc.X = x + width - diameter;
    path.AddArc(arc, 270.0f, 90.0f);
    arc.Y = y + height - diameter;
    path.AddArc(arc, 0.0f, 90.0f);
    arc.X = x;
    path.AddArc(arc, 90.0f, 90.0f);
    path.CloseFigure();
}

void FillRoundedRect(
    Gdiplus::Graphics& graphics,
    const RECT& rect,
    Gdiplus::REAL radius,
    COLORREF fillColor,
    COLORREF borderColor) {
    Gdiplus::GraphicsPath path;
    AddRoundedRect(path, rect, radius);

    Gdiplus::SolidBrush fillBrush(ToGdiplusColor(fillColor));
    Gdiplus::Pen borderPen(ToGdiplusColor(borderColor), 1.0f);
    borderPen.SetAlignment(Gdiplus::PenAlignmentInset);
    graphics.FillPath(&fillBrush, &path);
    graphics.DrawPath(&borderPen, &path);
}

class ConfigWindow {
public:
    explicit ConfigWindow(HINSTANCE instance)
        : instance_(instance) {
    }

    bool Create() {
        dpi_ = GetDpiForSystem();
        uiThemeMode_ = GetCurrentThemeMode();
        palette_ = MakeUiPalette(uiThemeMode_);

        WNDCLASSEXW windowClass{};
        windowClass.cbSize = sizeof(windowClass);
        windowClass.lpfnWndProc = &ConfigWindow::WindowProc;
        windowClass.hInstance = instance_;
        windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        windowClass.hbrBackground = nullptr;
        windowClass.lpszClassName = kConfigWindowClassName;
        RegisterClassExW(&windowClass);

        hwnd_ = CreateWindowExW(
            WS_EX_APPWINDOW,
            windowClass.lpszClassName,
            L"Auto Theme",
            WS_POPUP | WS_SYSMENU | WS_MINIMIZEBOX | WS_CLIPCHILDREN,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            Scale(kWindowWidth),
            Scale(kWindowHeight),
            nullptr,
            nullptr,
            instance_,
            this);
        if (hwnd_ == nullptr) {
            return false;
        }

        CenterWindow();
        return true;
    }

    int Run() {
        ShowWindow(hwnd_, SW_SHOWDEFAULT);
        UpdateWindow(hwnd_);
        MSG message{};
        while (GetMessageW(&message, nullptr, 0, 0)) {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        return static_cast<int>(message.wParam);
    }

private:
    int Scale(int value) const {
        return MulDiv(value, static_cast<int>(dpi_), USER_DEFAULT_SCREEN_DPI);
    }

    Gdiplus::REAL ScaleF(float value) const {
        return value * static_cast<Gdiplus::REAL>(dpi_) / 96.0f;
    }

    RECT AutomationCardRect() const {
        return MakeRect(Scale(24), Scale(112 + kContentTopOffset), Scale(384), Scale(198));
    }

    RECT LocationCardRect() const {
        return MakeRect(Scale(24), Scale(316 + kContentTopOffset), Scale(384), Scale(238));
    }

    RECT HotkeyCardRect() const {
        return MakeRect(Scale(24), Scale(560 + kContentTopOffset), Scale(384), Scale(240));
    }

    RECT StatusCardRect() const {
        return MakeRect(Scale(428), Scale(112 + kContentTopOffset), Scale(472), Scale(510));
    }

    RECT WallpaperCardRect() const {
        return MakeRect(Scale(428), Scale(638 + kContentTopOffset), Scale(472), Scale(162));
    }

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
        if (message == WM_NCCREATE) {
            const auto create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
            const auto self = static_cast<ConfigWindow*>(create->lpCreateParams);
            self->hwnd_ = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            return TRUE;
        }

        const auto self = reinterpret_cast<ConfigWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (self == nullptr) {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }
        return self->HandleMessage(message, wParam, lParam);
    }

    LRESULT HandleMessage(UINT message, WPARAM wParam, LPARAM lParam) {
        switch (message) {
        case WM_CREATE:
            dpi_ = GetDpiForWindow(hwnd_);
            SetWindowPos(hwnd_, nullptr, 0, 0, Scale(kWindowWidth), Scale(kWindowHeight), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
            CreateThemeResources();
            BuildUi();
            ApplyControlTheme();
            ApplyWindowThemeMode();
            LoadFromDisk();
            SetTimer(hwnd_, kRefreshTimerId, 2000, nullptr);
            return 0;
        case WM_DPICHANGED:
            return HandleDpiChanged(HIWORD(wParam), reinterpret_cast<const RECT*>(lParam));
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE:
            RefreshVisualTheme();
            return 0;
        case WM_COMMAND:
            return HandleCommand(LOWORD(wParam), HIWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WM_DRAWITEM:
            return HandleDrawItem(reinterpret_cast<const DRAWITEMSTRUCT*>(lParam));
        case WM_NOTIFY:
            return HandleNotify(reinterpret_cast<NMHDR*>(lParam));
        case WM_NCHITTEST:
            return HitTestWindow(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
        case WM_SETCURSOR:
            if (LOWORD(lParam) == HTCLIENT && IsModifierLabelControl(reinterpret_cast<HWND>(wParam))) {
                SetCursor(LoadCursorW(nullptr, IDC_HAND));
                return TRUE;
            }
            break;
        case WM_TIMER:
            if (wParam == kRefreshTimerId) {
                RefreshStatus();
            }
            return 0;
        case WM_MOUSEMOVE:
            OnMouseMove(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            return 0;
        case WM_MOUSELEAVE:
            ClearTitleButtonHotState();
            return 0;
        case WM_LBUTTONDOWN:
            if (HandleLeftButtonDown(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)})) {
                return 0;
            }
            break;
        case WM_LBUTTONUP:
            if (HandleLeftButtonUp(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)})) {
                return 0;
            }
            break;
        case WM_CAPTURECHANGED:
            ResetTitleButtonPressedState();
            return 0;
        case WM_ERASEBKGND:
            return 1;
        case WM_PAINT: {
            PAINTSTRUCT paint{};
            const auto hdc = BeginPaint(hwnd_, &paint);
            RECT client{};
            GetClientRect(hwnd_, &client);

            HDC bufferHdc = nullptr;
            const auto buffer = BeginBufferedPaint(hdc, &client, BPBF_TOPDOWNDIB, nullptr, &bufferHdc);
            if (buffer != nullptr && bufferHdc != nullptr) {
                PaintBackground(bufferHdc);
                BufferedPaintSetAlpha(buffer, &client, 255);
                EndBufferedPaint(buffer, TRUE);
            } else {
                PaintBackground(hdc);
            }
            EndPaint(hwnd_, &paint);
            return 0;
        }
        case WM_CTLCOLORBTN:
            return reinterpret_cast<LRESULT>(HandleButtonColor(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam)));
        case WM_CTLCOLORSTATIC:
            return reinterpret_cast<LRESULT>(HandleStaticColor(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam)));
        case WM_CTLCOLOREDIT:
            return reinterpret_cast<LRESULT>(HandleEditColor(reinterpret_cast<HDC>(wParam)));
        case WM_CTLCOLORLISTBOX:
            return reinterpret_cast<LRESULT>(HandleListBoxColor(reinterpret_cast<HDC>(wParam)));
        case WM_DESTROY:
            KillTimer(hwnd_, kRefreshTimerId);
            DestroyThemeResources();
            PostQuitMessage(0);
            return 0;
        }
        return DefWindowProcW(hwnd_, message, wParam, lParam);
    }

    LRESULT HandleDpiChanged(UINT newDpi, const RECT* suggestedRect) {
        dpi_ = newDpi;
        if (suggestedRect != nullptr) {
            SetWindowPos(
                hwnd_,
                nullptr,
                suggestedRect->left,
                suggestedRect->top,
                suggestedRect->right - suggestedRect->left,
                suggestedRect->bottom - suggestedRect->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
        } else {
            SetWindowPos(hwnd_, nullptr, 0, 0, Scale(kWindowWidth), Scale(kWindowHeight), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        RebuildUi();
        return 0;
    }

    void RebuildUi() {
        ClearUi();
        DestroyThemeResources();
        palette_ = MakeUiPalette(uiThemeMode_);
        CreateThemeResources();
        BuildUi();
        ApplyControlTheme();
        ApplyWindowThemeMode();
        ApplySettingsToControls();
        RefreshStatus();
        InvalidateRect(hwnd_, nullptr, TRUE);
    }

    void ClearUi() {
        while (const auto child = GetWindow(hwnd_, GW_CHILD)) {
            DestroyWindow(child);
        }

        headerLabels_.clear();
        sectionLabels_.clear();
        hintLabels_.clear();
        bodyLabels_.clear();

        automationTitle_ = nullptr;
        automationHint_ = nullptr;
        locationTitle_ = nullptr;
        locationHint_ = nullptr;
        hotkeyTitle_ = nullptr;
        hotkeyHint_ = nullptr;
        statusTitle_ = nullptr;
        statusHint_ = nullptr;
        wallpaperTitle_ = nullptr;
        wallpaperHint_ = nullptr;

        autoSwitch_ = nullptr;
        autostart_ = nullptr;
        locationAuto_ = nullptr;
        locationManual_ = nullptr;
        allowGps_ = nullptr;
        allowIp_ = nullptr;
        latitudeLabel_ = nullptr;
        latitude_ = nullptr;
        longitudeLabel_ = nullptr;
        longitude_ = nullptr;
        sunriseOffset_ = nullptr;
        sunsetOffset_ = nullptr;
        pollInterval_ = nullptr;
        locationRefresh_ = nullptr;
        hotkeyEnable_ = nullptr;
        hotkeyCtrl_ = nullptr;
        hotkeyAlt_ = nullptr;
        hotkeyShift_ = nullptr;
        hotkeyWin_ = nullptr;
        hotkeyKey_ = nullptr;
        configHotkeyEnable_ = nullptr;
        configHotkeyCtrl_ = nullptr;
        configHotkeyAlt_ = nullptr;
        configHotkeyShift_ = nullptr;
        configHotkeyWin_ = nullptr;
        configHotkeyKey_ = nullptr;
        statusBox_ = nullptr;
        lightWallpaper_ = nullptr;
        darkWallpaper_ = nullptr;
        saveButton_ = nullptr;
        reloadButton_ = nullptr;
        toggleButton_ = nullptr;
        launchButton_ = nullptr;
        terminateAgentButton_ = nullptr;
        openFolderButton_ = nullptr;
        browseLightWallpaperButton_ = nullptr;
        browseDarkWallpaperButton_ = nullptr;
    }

    void CenterWindow() {
        RECT windowRect{};
        GetWindowRect(hwnd_, &windowRect);

        const auto width = windowRect.right - windowRect.left;
        const auto height = windowRect.bottom - windowRect.top;
        const auto monitor = MonitorFromWindow(hwnd_, MONITOR_DEFAULTTONEAREST);

        MONITORINFO info{};
        info.cbSize = sizeof(info);
        GetMonitorInfoW(monitor, &info);

        const auto x = info.rcWork.left + ((info.rcWork.right - info.rcWork.left) - width) / 2;
        const auto y = info.rcWork.top + ((info.rcWork.bottom - info.rcWork.top) - height) / 2;
        SetWindowPos(hwnd_, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    void CreateThemeResources() {
        if (sectionFont_ == nullptr) {
            sectionFont_ = CreateUiFont(hwnd_, 12, FW_NORMAL, L"Segoe UI");
            bodyFont_ = CreateUiFont(hwnd_, 10, FW_NORMAL, L"Segoe UI");
            bodyBoldFont_ = CreateUiFont(hwnd_, 10, FW_SEMIBOLD, L"Segoe UI");
            hintFont_ = CreateUiFont(hwnd_, 9, FW_NORMAL, L"Segoe UI");
            buttonFont_ = CreateUiFont(hwnd_, 10, FW_LIGHT, L"Segoe UI");
            statusFont_ = CreateUiFont(hwnd_, 10, FW_NORMAL, L"Cascadia Mono");
        }

        DestroyBrushResources();
        backgroundBrush_ = CreateSolidBrush(palette_.windowBackground);
        cardBrush_ = CreateSolidBrush(palette_.cardBackground);
        inputBrush_ = CreateSolidBrush(palette_.inputBackground);
        accentBrush_ = CreateSolidBrush(palette_.accent);
    }

    void DestroyBrushResources() {
        for (auto brush : {backgroundBrush_, cardBrush_, inputBrush_, accentBrush_}) {
            if (brush != nullptr) {
                DeleteObject(brush);
            }
        }
        backgroundBrush_ = nullptr;
        cardBrush_ = nullptr;
        inputBrush_ = nullptr;
        accentBrush_ = nullptr;
    }

    void DestroyThemeResources() {
        for (auto font : {sectionFont_, bodyFont_, bodyBoldFont_, hintFont_, buttonFont_, statusFont_}) {
            if (font != nullptr) {
                DeleteObject(font);
            }
        }
        sectionFont_ = nullptr;
        bodyFont_ = nullptr;
        bodyBoldFont_ = nullptr;
        hintFont_ = nullptr;
        buttonFont_ = nullptr;
        statusFont_ = nullptr;
        DestroyBrushResources();
    }

    void ApplyWindowThemeMode() {
        if (hwnd_ == nullptr) {
            return;
        }

        const BOOL useDarkMode = palette_.dark ? TRUE : FALSE;
        DwmSetWindowAttribute(
            hwnd_,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            &useDarkMode,
            sizeof(useDarkMode));
    }

    void RefreshVisualTheme() {
        const auto newThemeMode = GetCurrentThemeMode();
        if (newThemeMode != uiThemeMode_) {
            uiThemeMode_ = newThemeMode;
            palette_ = MakeUiPalette(uiThemeMode_);
            CreateThemeResources();
        }
        ApplyControlTheme();
        ApplyWindowThemeMode();
        RedrawWindow(hwnd_, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
    }

    RECT TitleBarRect() const {
        RECT rect{};
        GetClientRect(hwnd_, &rect);
        rect.bottom = Scale(kTitleBarHeight);
        return rect;
    }

    RECT MinimizeButtonRect() const {
        RECT client{};
        GetClientRect(hwnd_, &client);
        return MakeRect(client.right - Scale(104), Scale(10), Scale(40), Scale(kTitleButtonHeight));
    }

    RECT CloseButtonRect() const {
        RECT client{};
        GetClientRect(hwnd_, &client);
        return MakeRect(client.right - Scale(56), Scale(10), Scale(40), Scale(kTitleButtonHeight));
    }

    void InvalidateTitleBar() {
        const auto rect = TitleBarRect();
        InvalidateRect(hwnd_, &rect, FALSE);
    }

    LRESULT HitTestWindow(int screenX, int screenY) const {
        POINT point{screenX, screenY};
        ScreenToClient(hwnd_, &point);
        const auto closeRect = CloseButtonRect();
        const auto minimizeRect = MinimizeButtonRect();
        const auto titleRect = TitleBarRect();
        if (PtInRect(&closeRect, point) || PtInRect(&minimizeRect, point)) {
            return HTCLIENT;
        }
        if (PtInRect(&titleRect, point)) {
            return HTCAPTION;
        }
        return HTCLIENT;
    }

    void StartMouseTracking() {
        if (trackingMouse_) {
            return;
        }
        TRACKMOUSEEVENT event{};
        event.cbSize = sizeof(event);
        event.dwFlags = TME_LEAVE;
        event.hwndTrack = hwnd_;
        if (TrackMouseEvent(&event)) {
            trackingMouse_ = true;
        }
    }

    void OnMouseMove(POINT point) {
        StartMouseTracking();

        const auto closeRect = CloseButtonRect();
        const auto minimizeRect = MinimizeButtonRect();
        const bool closeHot = PtInRect(&closeRect, point) != FALSE;
        const bool minimizeHot = PtInRect(&minimizeRect, point) != FALSE;
        if (closeHot != closeHot_ || minimizeHot != minimizeHot_) {
            closeHot_ = closeHot;
            minimizeHot_ = minimizeHot;
            InvalidateTitleBar();
        }
    }

    void ClearTitleButtonHotState() {
        trackingMouse_ = false;
        if (closeHot_ || minimizeHot_) {
            closeHot_ = false;
            minimizeHot_ = false;
            InvalidateTitleBar();
        }
    }

    void ResetTitleButtonPressedState() {
        if (closePressed_ || minimizePressed_) {
            closePressed_ = false;
            minimizePressed_ = false;
            ReleaseCapture();
            InvalidateTitleBar();
        }
    }

    bool HandleLeftButtonDown(POINT point) {
        const auto closeRect = CloseButtonRect();
        const auto minimizeRect = MinimizeButtonRect();
        if (PtInRect(&closeRect, point)) {
            closePressed_ = true;
            SetCapture(hwnd_);
            InvalidateTitleBar();
            return true;
        }
        if (PtInRect(&minimizeRect, point)) {
            minimizePressed_ = true;
            SetCapture(hwnd_);
            InvalidateTitleBar();
            return true;
        }
        return false;
    }

    bool HandleLeftButtonUp(POINT point) {
        const auto closeRect = CloseButtonRect();
        const auto minimizeRect = MinimizeButtonRect();
        const bool closePressed = closePressed_;
        const bool minimizePressed = minimizePressed_;
        const bool activateClose = closePressed_ && PtInRect(&closeRect, point) != FALSE;
        const bool activateMinimize = minimizePressed_ && PtInRect(&minimizeRect, point) != FALSE;

        if (!closePressed && !minimizePressed) {
            return false;
        }

        closePressed_ = false;
        minimizePressed_ = false;
        ReleaseCapture();
        InvalidateTitleBar();

        if (activateClose) {
            PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        } else if (activateMinimize) {
            ShowWindow(hwnd_, SW_MINIMIZE);
        }
        return true;
    }

    void BuildUi() {
        const auto s = [this](int value) { return Scale(value); };
        const auto automationCard = AutomationCardRect();
        const auto locationCard = LocationCardRect();
        const auto hotkeyCard = HotkeyCardRect();
        const auto statusCard = StatusCardRect();
        const auto wallpaperCard = WallpaperCardRect();
        const int leftMargin = s(20);

        automationTitle_ = CreateTextLabel(
            L"Automation",
            automationCard.left + leftMargin,
            automationCard.top + s(18),
            s(180),
            s(24),
            sectionFont_,
            sectionLabels_);
        automationHint_ = CreateTextLabel(
            L"General scheduling cadence and launch behavior.",
            automationCard.left + leftMargin,
            automationCard.top + s(46),
            s(280),
            s(18),
            hintFont_,
            hintLabels_);
        autoSwitch_ = CreateCheckBox(
            IdAutoSwitch,
            L"Switch automatically at local sunrise and sunset",
            automationCard.left + leftMargin,
            automationCard.top + s(74),
            s(320),
            s(kCheckboxHeight));
        autostart_ = CreateCheckBox(
            IdAutostart,
            L"Start the agent automatically when I sign in",
            automationCard.left + leftMargin,
            automationCard.top + s(102),
            s(320),
            s(kCheckboxHeight));
        CreateFieldLabel(L"Poll interval", automationCard.left + leftMargin, automationCard.top + s(132));
        pollInterval_ = CreateEdit(IdPollInterval, automationCard.left + s(182), automationCard.top + s(128), s(72), s(kEditHeight));
        CreateHintText(L"minutes", automationCard.left + s(264), automationCard.top + s(134), s(80));
        CreateFieldLabel(L"Location refresh", automationCard.left + leftMargin, automationCard.top + s(162));
        locationRefresh_ = CreateEdit(IdLocationRefresh, automationCard.left + s(182), automationCard.top + s(158), s(72), s(kEditHeight));
        CreateHintText(L"minutes", automationCard.left + s(264), automationCard.top + s(164), s(80));

        locationTitle_ = CreateTextLabel(
            L"Location",
            locationCard.left + leftMargin,
            locationCard.top + s(18),
            s(180),
            s(24),
            sectionFont_,
            sectionLabels_);
        locationHint_ = CreateTextLabel(
            L"Use GPS first, then fall back to direct IP geolocation when needed.",
            locationCard.left + leftMargin,
            locationCard.top + s(46),
            s(300),
            s(18),
            hintFont_,
            hintLabels_);
        locationAuto_ = CreateRadio(
            IdLocationAuto,
            L"Automatic lookup (GPS first, then IP)",
            locationCard.left + leftMargin,
            locationCard.top + s(74),
            s(320),
            s(kCheckboxHeight),
            true);
        locationManual_ = CreateRadio(
            IdLocationManual,
            L"Manual coordinates",
            locationCard.left + leftMargin,
            locationCard.top + s(102),
            s(220),
            s(kCheckboxHeight),
            false);
        allowGps_ = CreateCheckBox(
            IdAllowGps,
            L"Allow GPS / Windows location access",
            locationCard.left + leftMargin,
            locationCard.top + s(130),
            s(320),
            s(kCheckboxHeight));
        allowIp_ = CreateCheckBox(
            IdAllowIp,
            L"Allow direct IP fallback",
            locationCard.left + leftMargin,
            locationCard.top + s(158),
            s(320),
            s(kCheckboxHeight));
        latitudeLabel_ = CreateTextLabel(
            L"Latitude",
            locationCard.left + s(36),
            locationCard.top + s(194),
            s(54),
            s(20),
            bodyFont_,
            bodyLabels_);
        latitude_ = CreateEdit(
            IdLatitude,
            locationCard.left + s(96),
            locationCard.top + s(188),
            s(92),
            s(kEditHeight));
        longitudeLabel_ = CreateTextLabel(
            L"Longitude",
            locationCard.left + s(200),
            locationCard.top + s(194),
            s(62),
            s(20),
            bodyFont_,
            bodyLabels_);
        longitude_ = CreateEdit(
            IdLongitude,
            locationCard.left + s(270),
            locationCard.top + s(188),
            s(92),
            s(kEditHeight));

        hotkeyTitle_ = CreateTextLabel(
            L"Offsets and shortcuts",
            hotkeyCard.left + leftMargin,
            hotkeyCard.top + s(18),
            s(260),
            s(24),
            sectionFont_,
            sectionLabels_);
        hotkeyHint_ = CreateTextLabel(
            L"Set offsets plus global toggle and config shortcuts.",
            hotkeyCard.left + leftMargin,
            hotkeyCard.top + s(46),
            s(320),
            s(18),
            hintFont_,
            hintLabels_);
        CreateFieldLabel(L"Sunrise offset", hotkeyCard.left + leftMargin, hotkeyCard.top + s(78));
        sunriseOffset_ = CreateEdit(IdSunriseOffset, hotkeyCard.left + s(182), hotkeyCard.top + s(74), s(72), s(kEditHeight));
        CreateHintText(L"minutes", hotkeyCard.left + s(264), hotkeyCard.top + s(80), s(80));
        CreateFieldLabel(L"Sunset offset", hotkeyCard.left + leftMargin, hotkeyCard.top + s(110));
        sunsetOffset_ = CreateEdit(IdSunsetOffset, hotkeyCard.left + s(182), hotkeyCard.top + s(106), s(72), s(kEditHeight));
        CreateHintText(L"minutes", hotkeyCard.left + s(264), hotkeyCard.top + s(112), s(80));
        const int shortcutLeft = hotkeyCard.left + s(20);
        const int shortcutRight = hotkeyCard.left + s(198);
        const int shortcutLabelTop = hotkeyCard.top + s(146);
        const int shortcutKeyTop = hotkeyCard.top + s(180);
        const int shortcutComboTop = hotkeyCard.top + s(176);
        const int shortcutModifierTop = hotkeyCard.top + s(204);
        configHotkeyEnable_ = CreateCheckBox(
            IdConfigHotkeyEnable,
            L"Launch config hotkey",
            shortcutLeft,
            shortcutLabelTop,
            s(164),
            s(kCheckboxHeight));
        hotkeyEnable_ = CreateCheckBox(
            IdHotkeyEnable,
            L"Toggle Dark/Light",
            shortcutRight,
            shortcutLabelTop,
            s(164),
            s(kCheckboxHeight));
        CreateFieldLabel(L"Key", shortcutLeft, shortcutKeyTop);
        CreateFieldLabel(L"Key", shortcutRight, shortcutKeyTop);
        configHotkeyKey_ = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"COMBOBOX",
            nullptr,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
            shortcutLeft + s(36),
            shortcutComboTop,
            s(110),
            s(240),
            hwnd_,
            reinterpret_cast<HMENU>(IdConfigHotkeyKey),
            instance_,
            nullptr);
        PopulateHotkeyKeyCombo(configHotkeyKey_);
        hotkeyKey_ = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"COMBOBOX",
            nullptr,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
            shortcutRight + s(36),
            shortcutComboTop,
            s(110),
            s(240),
            hwnd_,
            reinterpret_cast<HMENU>(IdHotkeyKey),
            instance_,
            nullptr);
        PopulateHotkeyKeyCombo(hotkeyKey_);
        configHotkeyCtrl_ = CreateModifierToggle(IdConfigHotkeyCtrl, L"Ctrl", shortcutLeft, shortcutModifierTop, s(24), s(kCheckboxHeight));
        configHotkeyAlt_ = CreateModifierToggle(IdConfigHotkeyAlt, L"Alt", shortcutLeft + s(40), shortcutModifierTop, s(20), s(kCheckboxHeight));
        configHotkeyShift_ = CreateModifierToggle(IdConfigHotkeyShift, L"Shift", shortcutLeft + s(76), shortcutModifierTop, s(30), s(kCheckboxHeight));
        configHotkeyWin_ = CreateModifierToggle(IdConfigHotkeyWin, L"Win", shortcutLeft + s(122), shortcutModifierTop, s(24), s(kCheckboxHeight));
        hotkeyCtrl_ = CreateModifierToggle(IdHotkeyCtrl, L"Ctrl", shortcutRight, shortcutModifierTop, s(24), s(kCheckboxHeight));
        hotkeyAlt_ = CreateModifierToggle(IdHotkeyAlt, L"Alt", shortcutRight + s(40), shortcutModifierTop, s(20), s(kCheckboxHeight));
        hotkeyShift_ = CreateModifierToggle(IdHotkeyShift, L"Shift", shortcutRight + s(76), shortcutModifierTop, s(30), s(kCheckboxHeight));
        hotkeyWin_ = CreateModifierToggle(IdHotkeyWin, L"Win", shortcutRight + s(122), shortcutModifierTop, s(24), s(kCheckboxHeight));

        statusTitle_ = CreateTextLabel(
            L"Current status",
            statusCard.left + leftMargin,
            statusCard.top + s(18),
            s(220),
            s(24),
            sectionFont_,
            sectionLabels_);
        statusHint_ = CreateTextLabel(
            L"Live runtime state from the background agent.",
            statusCard.left + leftMargin,
            statusCard.top + s(46),
            s(260),
            s(18),
            hintFont_,
            hintLabels_);
        statusBox_ = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"EDIT",
            nullptr,
            WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
            statusCard.left + leftMargin,
            statusCard.top + s(78),
            s(432),
            s(422),
            hwnd_,
            reinterpret_cast<HMENU>(IdStatusBox),
            instance_,
            nullptr);
        SendMessageW(statusBox_, WM_SETFONT, reinterpret_cast<WPARAM>(statusFont_), TRUE);
        SendMessageW(statusBox_, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(s(10), s(10)));

        wallpaperTitle_ = CreateTextLabel(
            L"Wallpapers",
            wallpaperCard.left + leftMargin,
            wallpaperCard.top + s(18),
            s(220),
            s(24),
            sectionFont_,
            sectionLabels_);
        wallpaperHint_ = CreateTextLabel(
            L"Optional image paths applied when the theme switches. Leave empty to keep the current wallpaper.",
            wallpaperCard.left + leftMargin,
            wallpaperCard.top + s(46),
            s(392),
            s(18),
            hintFont_,
            hintLabels_);
        CreateFieldLabel(L"Light", wallpaperCard.left + leftMargin, wallpaperCard.top + s(80));
        lightWallpaper_ = CreateEdit(
            IdLightWallpaper,
            wallpaperCard.left + s(86),
            wallpaperCard.top + s(76),
            s(270),
            s(kEditHeight));
        browseLightWallpaperButton_ = CreateActionButton(
            IdBrowseLightWallpaper,
            L"Browse",
            wallpaperCard.left + s(366),
            wallpaperCard.top + s(76),
            s(76),
            s(30),
            false);
        CreateFieldLabel(L"Dark", wallpaperCard.left + leftMargin, wallpaperCard.top + s(112));
        darkWallpaper_ = CreateEdit(
            IdDarkWallpaper,
            wallpaperCard.left + s(86),
            wallpaperCard.top + s(108),
            s(270),
            s(kEditHeight));
        browseDarkWallpaperButton_ = CreateActionButton(
            IdBrowseDarkWallpaper,
            L"Browse",
            wallpaperCard.left + s(366),
            wallpaperCard.top + s(108),
            s(76),
            s(30),
            false);

        saveButton_ = CreateActionButton(IdSave, L"Save Settings", s(24), s(814 + kContentTopOffset), s(126), s(34), true);
        reloadButton_ = CreateActionButton(IdReload, L"Reload", s(160), s(814 + kContentTopOffset), s(92), s(34), false);
        toggleButton_ = CreateActionButton(IdToggle, L"Toggle Now", s(262), s(814 + kContentTopOffset), s(108), s(34), false);
        launchButton_ = CreateActionButton(IdLaunch, L"Launch Agent", s(428), s(814 + kContentTopOffset), s(126), s(34), false);
        terminateAgentButton_ = CreateActionButton(IdTerminateAgent, L"Terminate Agent", s(564), s(814 + kContentTopOffset), s(126), s(34), false);
        openFolderButton_ = CreateActionButton(IdOpenFolder, L"Open Data Folder", s(700), s(814 + kContentTopOffset), s(140), s(34), false);
    }

    LRESULT HandleCommand(int id, int notifyCode, HWND control) {
        if (notifyCode == BN_CLICKED && IsModifierLabelControl(control)) {
            ToggleCheckbox(control);
            return 0;
        }

        switch (id) {
        case IdLocationAuto:
        case IdLocationManual:
        case IdHotkeyEnable:
        case IdConfigHotkeyEnable:
            UpdateEnabledState();
            return 0;
        case IdSave:
            SaveCurrentSettings();
            return 0;
        case IdReload:
            LoadFromDisk();
            return 0;
        case IdToggle:
            ToggleNow();
            return 0;
        case IdLaunch:
            LaunchAgentAndRefresh();
            return 0;
        case IdTerminateAgent:
            StopAgentAndRefresh();
            return 0;
        case IdOpenFolder:
            ShellExecuteW(nullptr, L"open", GetDataDirectory().c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            return 0;
        case IdBrowseLightWallpaper:
            BrowseWallpaper(lightWallpaper_);
            return 0;
        case IdBrowseDarkWallpaper:
            BrowseWallpaper(darkWallpaper_);
            return 0;
        default:
            break;
        }
        return 0;
    }

    void LoadFromDisk() {
        std::wstring error;
        if (!LoadSettings(settings_, &error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
        }
        LoadStatus(status_, nullptr);
        ApplySettingsToControls();
        RefreshStatus();
    }

    void ApplySettingsToControls() {
        SetCheck(autoSwitch_, settings_.autoSwitchEnabled);
        SetCheck(autostart_, settings_.autostartEnabled);
        SetCheck(locationAuto_, settings_.locationMode == LocationMode::Auto);
        SetCheck(locationManual_, settings_.locationMode == LocationMode::Manual);
        SetCheck(allowGps_, settings_.allowGps);
        SetCheck(allowIp_, settings_.allowIpFallback);
        SetText(lightWallpaper_, settings_.lightWallpaperPath);
        SetText(darkWallpaper_, settings_.darkWallpaperPath);
        manualLatitudeText_ = FormatDouble(settings_.manualLatitude, 6);
        manualLongitudeText_ = FormatDouble(settings_.manualLongitude, 6);
        coordinatesShowingManual_ = false;
        SetText(sunriseOffset_, std::to_wstring(settings_.sunriseOffsetMinutes));
        SetText(sunsetOffset_, std::to_wstring(settings_.sunsetOffsetMinutes));
        SetText(pollInterval_, std::to_wstring(settings_.pollIntervalMinutes));
        SetText(locationRefresh_, std::to_wstring(settings_.locationRefreshMinutes));
        ApplyHotkeyToControls(settings_.hotkey, hotkeyEnable_, hotkeyCtrl_, hotkeyAlt_, hotkeyShift_, hotkeyWin_, hotkeyKey_);
        ApplyHotkeyToControls(
            settings_.configHotkey,
            configHotkeyEnable_,
            configHotkeyCtrl_,
            configHotkeyAlt_,
            configHotkeyShift_,
            configHotkeyWin_,
            configHotkeyKey_);
        UpdateEnabledState();
    }

    bool CollectSettings(Settings& updated, std::wstring& error) {
        updated = settings_;
        updated.autoSwitchEnabled = GetCheck(autoSwitch_);
        updated.autostartEnabled = GetCheck(autostart_);
        updated.locationMode = GetCheck(locationManual_) ? LocationMode::Manual : LocationMode::Auto;
        updated.allowGps = GetCheck(allowGps_);
        updated.allowIpFallback = GetCheck(allowIp_);
        updated.lightWallpaperPath = ReadControlText(lightWallpaper_);
        updated.darkWallpaperPath = ReadControlText(darkWallpaper_);

        if (!ParseInt(ReadControlText(pollInterval_), updated.pollIntervalMinutes)) {
            error = L"Poll interval must be an integer.";
            return false;
        }
        if (!ParseInt(ReadControlText(locationRefresh_), updated.locationRefreshMinutes)) {
            error = L"Location refresh must be an integer.";
            return false;
        }
        if (!ParseInt(ReadControlText(sunriseOffset_), updated.sunriseOffsetMinutes) ||
            !ParseInt(ReadControlText(sunsetOffset_), updated.sunsetOffsetMinutes)) {
            error = L"Offsets must be integers.";
            return false;
        }
        if (updated.locationMode == LocationMode::Manual) {
            if (!ParseDouble(ReadControlText(latitude_), updated.manualLatitude) ||
                !ParseDouble(ReadControlText(longitude_), updated.manualLongitude)) {
                error = L"Manual latitude and longitude must be valid numbers.";
                return false;
            }
        }

        if (!CollectHotkeyFromControls(
                updated.hotkey,
                hotkeyEnable_,
                hotkeyCtrl_,
                hotkeyAlt_,
                hotkeyShift_,
                hotkeyWin_,
                hotkeyKey_,
                L"manual toggle",
                error)) {
            return false;
        }
        if (!CollectHotkeyFromControls(
                updated.configHotkey,
                configHotkeyEnable_,
                configHotkeyCtrl_,
                configHotkeyAlt_,
                configHotkeyShift_,
                configHotkeyWin_,
                configHotkeyKey_,
                L"launch config",
                error)) {
            return false;
        }
        if (HotkeysConflict(updated.hotkey, updated.configHotkey)) {
            error = L"Manual toggle and launch config hotkeys must be different.";
            return false;
        }

        updated.pollIntervalMinutes = std::clamp(updated.pollIntervalMinutes, 1, 60);
        updated.locationRefreshMinutes = std::clamp(updated.locationRefreshMinutes, 15, 1440);
        updated.sunriseOffsetMinutes = std::clamp(updated.sunriseOffsetMinutes, -180, 180);
        updated.sunsetOffsetMinutes = std::clamp(updated.sunsetOffsetMinutes, -180, 180);
        return true;
    }

    void SaveCurrentSettings() {
        Settings updated;
        std::wstring error;
        if (!CollectSettings(updated, error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
            return;
        }

        if (!SaveSettings(updated, &error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
            return;
        }

        settings_ = updated;
        manualLatitudeText_ = FormatDouble(settings_.manualLatitude, 6);
        manualLongitudeText_ = FormatDouble(settings_.manualLongitude, 6);
        EnsureAutostartEnabled(settings_.autostartEnabled, nullptr);
        if (!NotifyAgent(kAgentMessageReload)) {
            LaunchAgent(nullptr);
        }
        RefreshStatus();
    }

    void ToggleNow() {
        if (!NotifyAgent(kAgentMessageToggle)) {
            LaunchAgent(nullptr);
            Sleep(400);
            NotifyAgent(kAgentMessageToggle);
        }
    }

    void BrowseWallpaper(HWND target) {
        wchar_t buffer[32768]{};
        const auto current = ReadControlText(target);
        if (!current.empty() && std::filesystem::exists(current)) {
            wcsncpy_s(buffer, current.c_str(), _TRUNCATE);
        }

        OPENFILENAMEW dialog{};
        dialog.lStructSize = sizeof(dialog);
        dialog.hwndOwner = hwnd_;
        dialog.lpstrFile = buffer;
        dialog.nMaxFile = static_cast<DWORD>(_countof(buffer));
        dialog.lpstrFilter =
            L"Image Files\0*.bmp;*.dib;*.jpg;*.jpeg;*.jpe;*.png;*.gif;*.tif;*.tiff;*.webp\0All Files\0*.*\0\0";
        dialog.nFilterIndex = 1;
        dialog.Flags = OFN_EXPLORER | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST;

        if (GetOpenFileNameW(&dialog) && buffer[0] != L'\0' && std::filesystem::exists(buffer)) {
            SetText(target, buffer);
        }
    }

    void LaunchAgentAndRefresh() {
        std::wstring error;
        if (IsAgentRunning() && !StopAgent(&error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
            return;
        }
        if (!LaunchAgent(&error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
            return;
        }
        Sleep(400);
        RefreshStatus();
    }

    void StopAgentAndRefresh() {
        std::wstring error;
        if (!StopAgent(&error)) {
            MessageBoxW(hwnd_, error.c_str(), L"Auto Theme", MB_ICONERROR | MB_OK);
            return;
        }
        RefreshStatus();
    }

    void RefreshStatus() {
        LoadStatus(status_, nullptr);
        status_.agentRunning = IsAgentRunning();
        if (status_.hotkeyText.empty()) {
            status_.hotkeyText = FormatHotkey(settings_.hotkey);
        }
        if (status_.configHotkeyText.empty()) {
            status_.configHotkeyText = FormatHotkey(settings_.configHotkey);
        }
        std::wstring text;
        text += L"Theme status\r\n";
        text += L"------------\r\n";
        AppendStatusLine(text, L"Agent running", YesNo(status_.agentRunning));
        AppendStatusLine(text, L"Current mode", status_.currentMode);
        AppendStatusLine(text, L"Desired mode", status_.desiredMode);
        AppendStatusLine(text, L"Auto switch", YesNo(status_.autoSwitchEnabled));
        AppendStatusLine(text, L"Toggle hotkey", status_.hotkeyText);
        AppendStatusLine(text, L"Config hotkey", status_.configHotkeyText);

        text += L"\r\nLocation\r\n";
        text += L"--------\r\n";
        AppendStatusLine(text, L"Source", status_.locationSource);
        if (status_.hasLocation) {
            AppendStatusLine(
                text,
                L"Coordinates",
                FormatDouble(status_.latitude, 6) + L", " + FormatDouble(status_.longitude, 6));
        } else {
            AppendStatusLine(text, L"Coordinates", L"-");
        }
        AppendStatusLine(text, L"Last refresh", status_.lastLocationRefresh);

        text += L"\r\nSchedule\r\n";
        text += L"--------\r\n";
        AppendStatusLine(text, L"Sunrise", status_.sunriseTime);
        AppendStatusLine(text, L"Sunset", status_.sunsetTime);
        AppendStatusLine(text, L"Next transition", status_.nextTransitionTime);
        if (status_.manualOverrideActive) {
            AppendStatusLine(
                text,
                L"Manual override",
                status_.manualOverrideMode + L" until " + EmptyToDash(status_.manualOverrideUntil));
        } else {
            AppendStatusLine(text, L"Manual override", L"None");
        }
        AppendStatusLine(text, L"Last applied", status_.lastAppliedTime);
        AppendStatusLine(text, L"Autostart registered", YesNo(IsAutostartEnabled()));

        if (!status_.lastError.empty()) {
            text += L"\r\nDiagnostics\r\n";
            text += L"-----------\r\n";
            text += status_.lastError;
            text += L"\r\n";
        }
        SetWindowTextW(statusBox_, text.c_str());
        UpdateLocationCoordinateDisplay();
        UpdateAgentButtons();
    }

    void UpdateEnabledState() {
        const auto manual = GetCheck(locationManual_);
        EnableWindow(allowGps_, !manual);
        EnableWindow(allowIp_, !manual);
        UpdateLocationCoordinateDisplay();

        UpdateHotkeyEnabledState(
            hotkeyEnable_,
            hotkeyCtrl_,
            hotkeyAlt_,
            hotkeyShift_,
            hotkeyWin_,
            hotkeyKey_);
        UpdateHotkeyEnabledState(
            configHotkeyEnable_,
            configHotkeyCtrl_,
            configHotkeyAlt_,
            configHotkeyShift_,
            configHotkeyWin_,
            configHotkeyKey_);

        RedrawWindow(hwnd_, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
    }

    void UpdateAgentButtons() {
        SetWindowTextW(launchButton_, status_.agentRunning ? L"Relaunch Agent" : L"Launch Agent");
        EnableWindow(terminateAgentButton_, status_.agentRunning);
    }

    void PaintBackground(HDC hdc) {
        RECT client{};
        GetClientRect(hwnd_, &client);
        FillRect(hdc, &client, backgroundBrush_);

        DrawWindowShell(hdc, client);
        DrawTitleBar(hdc);

        DrawCard(hdc, AutomationCardRect());
        DrawCard(hdc, LocationCardRect());
        DrawCard(hdc, HotkeyCardRect());
        DrawCard(hdc, StatusCardRect());
        DrawCard(hdc, WallpaperCardRect());
        DrawWindowBorder(hdc, client);
    }

    void DrawWindowShell(HDC hdc, const RECT& client) {
        Gdiplus::Graphics graphics(hdc);
        ConfigureGraphics(graphics);
        Gdiplus::SolidBrush backgroundBrush(ToGdiplusColor(palette_.windowBackground));
        graphics.FillRectangle(
            &backgroundBrush,
            static_cast<Gdiplus::REAL>(client.left),
            static_cast<Gdiplus::REAL>(client.top),
            static_cast<Gdiplus::REAL>(client.right - client.left),
            static_cast<Gdiplus::REAL>(client.bottom - client.top));
    }

    void DrawWindowBorder(HDC hdc, const RECT& client) {
        Gdiplus::Graphics graphics(hdc);
        ConfigureGraphics(graphics);

        Gdiplus::Pen borderPen(ToGdiplusColor(palette_.windowBorder), 1.0f);
        borderPen.SetAlignment(Gdiplus::PenAlignmentInset);
        graphics.DrawRectangle(
            &borderPen,
            static_cast<Gdiplus::REAL>(client.left),
            static_cast<Gdiplus::REAL>(client.top),
            static_cast<Gdiplus::REAL>(std::max(1L, client.right - client.left - 1)),
            static_cast<Gdiplus::REAL>(std::max(1L, client.bottom - client.top - 1)));
    }

    void DrawTitleBar(HDC hdc) {
        Gdiplus::Graphics graphics(hdc);
        ConfigureGraphics(graphics);

        const auto titleRect = TitleBarRect();
        Gdiplus::SolidBrush backgroundBrush(ToGdiplusColor(palette_.titleBarBackground));
        graphics.FillRectangle(
            &backgroundBrush,
            static_cast<Gdiplus::REAL>(titleRect.left),
            static_cast<Gdiplus::REAL>(titleRect.top),
            static_cast<Gdiplus::REAL>(titleRect.right - titleRect.left),
            static_cast<Gdiplus::REAL>(titleRect.bottom - titleRect.top));

        Gdiplus::SolidBrush underlineBrush(ToGdiplusColor(palette_.titleBarBorder));
        graphics.FillRectangle(&underlineBrush, 1.0f, static_cast<Gdiplus::REAL>(titleRect.bottom - 1),
            static_cast<Gdiplus::REAL>(titleRect.right - 2), 1.0f);

        Gdiplus::SolidBrush accentBrush(ToGdiplusColor(palette_.accent));
        graphics.FillEllipse(&accentBrush, ScaleF(24.0f), ScaleF(18.0f), ScaleF(12.0f), ScaleF(12.0f));

        Gdiplus::FontFamily uiFamily(L"Segoe UI");
        Gdiplus::Font titleFont(&uiFamily, ScaleF(16.5f), Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::Font subtitleFont(&uiFamily, ScaleF(11.0f), Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::StringFormat textFormat;
        textFormat.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);
        textFormat.SetAlignment(Gdiplus::StringAlignmentNear);
        textFormat.SetLineAlignment(Gdiplus::StringAlignmentNear);

        Gdiplus::SolidBrush titleBrush(ToGdiplusColor(palette_.primaryText));
        graphics.DrawString(
            L"Auto Theme",
            -1,
            &titleFont,
            Gdiplus::RectF(ScaleF(48.0f), ScaleF(8.0f), ScaleF(340.0f), ScaleF(24.0f)),
            &textFormat,
            &titleBrush);

        Gdiplus::SolidBrush subtitleBrush(ToGdiplusColor(palette_.secondaryText));
        graphics.DrawString(
            L"configuration",
            -1,
            &subtitleFont,
            Gdiplus::RectF(ScaleF(48.0f), ScaleF(30.0f), ScaleF(340.0f), ScaleF(18.0f)),
            &textFormat,
            &subtitleBrush);

        DrawTitleBarButton(hdc, MinimizeButtonRect(), L"\uE921", minimizeHot_, minimizePressed_, false);
        DrawTitleBarButton(hdc, CloseButtonRect(), L"\uE8BB", closeHot_, closePressed_, true);
    }

    void DrawTitleBarButton(
        HDC hdc,
        const RECT& rect,
        const wchar_t* text,
        bool hot,
        bool pressed,
        bool destructive) {
        COLORREF background = palette_.titleBarBackground;
        COLORREF foreground = palette_.primaryText;

        if (destructive && pressed) {
            background = palette_.closeButtonPressed;
            foreground = RGB(255, 255, 255);
        } else if (destructive && hot) {
            background = palette_.closeButtonHover;
            foreground = RGB(255, 255, 255);
        } else if (pressed) {
            background = palette_.titleButtonPressed;
        } else if (hot) {
            background = palette_.titleButtonHover;
        }

        Gdiplus::Graphics graphics(hdc);
        ConfigureGraphics(graphics);
        FillRoundedRect(graphics, rect, ScaleF(9.0f), background, background);

        Gdiplus::FontFamily iconFamily(L"Segoe MDL2 Assets");
        Gdiplus::Font iconFont(&iconFamily, ScaleF(11.5f), Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::StringFormat iconFormat;
        iconFormat.SetAlignment(Gdiplus::StringAlignmentCenter);
        iconFormat.SetLineAlignment(Gdiplus::StringAlignmentCenter);
        iconFormat.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);

        Gdiplus::SolidBrush textBrush(ToGdiplusColor(foreground));
        graphics.DrawString(
            text,
            -1,
            &iconFont,
            Gdiplus::RectF(
                static_cast<Gdiplus::REAL>(rect.left),
                static_cast<Gdiplus::REAL>(rect.top - (text[0] == L'\uE921' ? Scale(1) : 0)),
                static_cast<Gdiplus::REAL>(rect.right - rect.left),
                static_cast<Gdiplus::REAL>(rect.bottom - rect.top)),
            &iconFormat,
            &textBrush);
    }

    void DrawCard(HDC hdc, const RECT& rect) {
        Gdiplus::Graphics graphics(hdc);
        ConfigureGraphics(graphics);
        FillRoundedRect(graphics, rect, ScaleF(18.0f), palette_.cardBackground, palette_.cardBorder);

        Gdiplus::SolidBrush accentBrush(ToGdiplusColor(palette_.accent));
        graphics.FillRectangle(
            &accentBrush,
            static_cast<Gdiplus::REAL>(rect.left + Scale(18)),
            static_cast<Gdiplus::REAL>(rect.top + Scale(14)),
            ScaleF(72.0f),
            ScaleF(4.0f));
    }

    void ApplyControlTheme() {
        const wchar_t* surfaceTheme = palette_.dark ? L"DarkMode_Explorer" : L"Explorer";
        for (auto control : {
                 latitude_, longitude_, sunriseOffset_, sunsetOffset_, pollInterval_, locationRefresh_,
                 hotkeyKey_, configHotkeyKey_, statusBox_, lightWallpaper_, darkWallpaper_}) {
            if (control != nullptr) {
                SetWindowTheme(control, surfaceTheme, nullptr);
            }
        }

        for (auto control : {
                 autoSwitch_, autostart_, locationAuto_, locationManual_, allowGps_, allowIp_,
                 hotkeyEnable_, configHotkeyEnable_}) {
            if (control != nullptr) {
                SetWindowTheme(control, surfaceTheme, nullptr);
            }
        }
    }

    bool IsChoiceControl(HWND control) const {
        return control == autoSwitch_ ||
            control == autostart_ ||
            control == locationAuto_ ||
            control == locationManual_ ||
            control == allowGps_ ||
            control == allowIp_ ||
            control == hotkeyEnable_ ||
            control == configHotkeyEnable_;
    }

    bool IsModifierLabelControl(HWND control) const {
        return control == hotkeyCtrl_ ||
            control == hotkeyAlt_ ||
            control == hotkeyShift_ ||
            control == hotkeyWin_ ||
            control == configHotkeyCtrl_ ||
            control == configHotkeyAlt_ ||
            control == configHotkeyShift_ ||
            control == configHotkeyWin_;
    }

    bool IsActionButton(HWND control) const {
        return control == saveButton_ ||
            control == reloadButton_ ||
            control == toggleButton_ ||
            control == launchButton_ ||
            control == terminateAgentButton_ ||
            control == openFolderButton_ ||
            control == browseLightWallpaperButton_ ||
            control == browseDarkWallpaperButton_;
    }

    bool IsDefaultActionButton(HWND control) const {
        return control == saveButton_;
    }

    bool IsRadioControl(HWND control) const {
        return control == locationAuto_ ||
            control == locationManual_;
    }

    void SetControlVisible(HWND control, bool visible) {
        if (control == nullptr) {
            return;
        }
        ShowWindow(control, visible ? SW_SHOWNA : SW_HIDE);
    }

    void UpdateLocationCoordinateDisplay() {
        const auto manual = GetCheck(locationManual_);
        if (manual) {
            if (!coordinatesShowingManual_) {
                SetText(latitude_, manualLatitudeText_);
                SetText(longitude_, manualLongitudeText_);
                coordinatesShowingManual_ = true;
            }
            SendMessageW(latitude_, EM_SETREADONLY, FALSE, 0);
            SendMessageW(longitude_, EM_SETREADONLY, FALSE, 0);
            return;
        }

        if (coordinatesShowingManual_) {
            manualLatitudeText_ = ReadControlText(latitude_);
            manualLongitudeText_ = ReadControlText(longitude_);
            coordinatesShowingManual_ = false;
        }

        SendMessageW(latitude_, EM_SETREADONLY, TRUE, 0);
        SendMessageW(longitude_, EM_SETREADONLY, TRUE, 0);

        if (status_.hasLocation) {
            SetText(latitude_, FormatDouble(status_.latitude, 6));
            SetText(longitude_, FormatDouble(status_.longitude, 6));
        } else {
            SetText(latitude_, L"-");
            SetText(longitude_, L"-");
        }
    }

    LRESULT HandleNotify(NMHDR* header) {
        if (header == nullptr) {
            return 0;
        }
        if (header->code == NM_CUSTOMDRAW && IsChoiceControl(header->hwndFrom)) {
            return HandleChoiceCustomDraw(reinterpret_cast<NMCUSTOMDRAW*>(header));
        }
        return 0;
    }

    LRESULT HandleChoiceCustomDraw(NMCUSTOMDRAW* draw) {
        if (draw == nullptr || draw->dwDrawStage != CDDS_PREPAINT) {
            return CDRF_DODEFAULT;
        }

        DrawChoiceControl(*draw);
        return CDRF_SKIPDEFAULT;
    }

    void DrawChoiceControl(const NMCUSTOMDRAW& draw) {
        RECT rect = draw.rc;
        FillRect(draw.hdc, &rect, cardBrush_);

        const auto control = draw.hdr.hwndFrom;
        const auto checked = SendMessageW(control, BM_GETCHECK, 0, 0) == BST_CHECKED;
        const auto itemState = draw.uItemState;
        const bool disabled = (itemState & CDIS_DISABLED) != 0;
        const bool hot = (itemState & CDIS_HOT) != 0;
        const bool pressed = (itemState & CDIS_SELECTED) != 0;
        const bool focused = (itemState & CDIS_FOCUS) != 0;
        const bool radio = IsRadioControl(control);

        int stateId = 0;
        if (radio) {
            if (checked) {
                stateId = disabled ? RBS_CHECKEDDISABLED :
                    pressed ? RBS_CHECKEDPRESSED :
                    hot ? RBS_CHECKEDHOT :
                    RBS_CHECKEDNORMAL;
            } else {
                stateId = disabled ? RBS_UNCHECKEDDISABLED :
                    pressed ? RBS_UNCHECKEDPRESSED :
                    hot ? RBS_UNCHECKEDHOT :
                    RBS_UNCHECKEDNORMAL;
            }
        } else {
            if (checked) {
                stateId = disabled ? CBS_CHECKEDDISABLED :
                    pressed ? CBS_CHECKEDPRESSED :
                    hot ? CBS_CHECKEDHOT :
                    CBS_CHECKEDNORMAL;
            } else {
                stateId = disabled ? CBS_UNCHECKEDDISABLED :
                    pressed ? CBS_UNCHECKEDPRESSED :
                    hot ? CBS_UNCHECKEDHOT :
                    CBS_UNCHECKEDNORMAL;
            }
        }

        SIZE glyphSize{Scale(16), Scale(16)};
        RECT glyphRect{};
        glyphRect.left = rect.left + Scale(2);
        glyphRect.top = rect.top + ((rect.bottom - rect.top) - glyphSize.cy) / 2;
        glyphRect.right = glyphRect.left + glyphSize.cx;
        glyphRect.bottom = glyphRect.top + glyphSize.cy;

        if (const auto theme = OpenThemeData(control, L"Button")) {
            SIZE themedSize{};
            if (SUCCEEDED(GetThemePartSize(
                    theme,
                    draw.hdc,
                    radio ? BP_RADIOBUTTON : BP_CHECKBOX,
                    stateId,
                    nullptr,
                    TS_TRUE,
                    &themedSize))) {
                glyphSize = themedSize;
                glyphRect.top = rect.top + ((rect.bottom - rect.top) - glyphSize.cy) / 2;
                glyphRect.right = glyphRect.left + glyphSize.cx;
                glyphRect.bottom = glyphRect.top + glyphSize.cy;
            }

            DrawThemeBackground(theme, draw.hdc, radio ? BP_RADIOBUTTON : BP_CHECKBOX, stateId, &glyphRect, nullptr);
            CloseThemeData(theme);
        } else {
            UINT fallbackState = radio ? DFCS_BUTTONRADIO : DFCS_BUTTONCHECK;
            if (checked) {
                fallbackState |= DFCS_CHECKED;
            }
            if (disabled) {
                fallbackState |= DFCS_INACTIVE;
            }
            if (pressed) {
                fallbackState |= DFCS_PUSHED;
            }
            DrawFrameControl(draw.hdc, &glyphRect, DFC_BUTTON, fallbackState);
        }

        RECT textRect = rect;
        textRect.left = glyphRect.right + Scale(8);
        textRect.right -= Scale(2);
        SetBkMode(draw.hdc, TRANSPARENT);
        SetTextColor(draw.hdc, disabled ? palette_.secondaryText : palette_.primaryText);
        const auto previousFont = SelectObject(draw.hdc, bodyFont_);
        const auto text = ReadControlText(control);
        DrawTextW(
            draw.hdc,
            text.c_str(),
            -1,
            &textRect,
            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
        SelectObject(draw.hdc, previousFont);

        if (focused) {
            RECT focusRect = textRect;
            focusRect.left = std::max(rect.left, focusRect.left - Scale(2));
            DrawFocusRect(draw.hdc, &focusRect);
        }
    }

    LRESULT HandleDrawItem(const DRAWITEMSTRUCT* draw) {
        if (draw == nullptr || draw->CtlType != ODT_BUTTON) {
            return FALSE;
        }

        if (IsModifierLabelControl(draw->hwndItem)) {
            DrawModifierLabel(*draw);
            return TRUE;
        }
        if (!IsActionButton(draw->hwndItem)) {
            return FALSE;
        }

        DrawActionButton(*draw);
        return TRUE;
    }

    void DrawModifierLabel(const DRAWITEMSTRUCT& draw) {
        RECT rect = draw.rcItem;
        FillRect(draw.hDC, &rect, cardBrush_);

        const auto control = draw.hwndItem;
        const auto itemState = draw.itemState;
        const bool checked = GetCheck(control);
        const bool disabled = (itemState & ODS_DISABLED) != 0;
        const bool hot = (itemState & ODS_HOTLIGHT) != 0;
        const bool focused = (itemState & ODS_FOCUS) != 0;

        COLORREF foreground = palette_.modifierText;
        if (disabled) {
            foreground = palette_.modifierDisabledText;
        } else if (checked && hot) {
            foreground = palette_.modifierCheckedHoverText;
        } else if (checked) {
            foreground = palette_.modifierCheckedText;
        } else if (hot) {
            foreground = palette_.modifierHoverText;
        }

        RECT textRect = rect;
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, foreground);
        const auto previousFont = SelectObject(draw.hDC, checked ? bodyBoldFont_ : bodyFont_);
        const auto text = ReadControlText(control);
        DrawTextW(
            draw.hDC,
            text.c_str(),
            -1,
            &textRect,
            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
        SelectObject(draw.hDC, previousFont);

        if (focused) {
            RECT focusRect = rect;
            focusRect.right = std::min(rect.right, rect.left + Scale(4) + static_cast<int>(text.size()) * Scale(8));
            DrawFocusRect(draw.hDC, &focusRect);
        }
    }

    HBRUSH BackgroundBrushForActionButton(HWND control) const {
        return control == browseLightWallpaperButton_ || control == browseDarkWallpaperButton_
            ? cardBrush_
            : backgroundBrush_;
    }

    void DrawActionButton(const DRAWITEMSTRUCT& draw) {
        const auto control = draw.hwndItem;
        const auto itemState = draw.itemState;
        const bool disabled = (itemState & ODS_DISABLED) != 0;
        const bool hot = (itemState & ODS_HOTLIGHT) != 0;
        const bool pressed = (itemState & ODS_SELECTED) != 0;
        const bool focused = (itemState & ODS_FOCUS) != 0;
        const bool isDefault = IsDefaultActionButton(control);

        RECT rect = draw.rcItem;
        FillRect(draw.hDC, &rect, BackgroundBrushForActionButton(control));

        COLORREF background = palette_.actionButtonBackground;
        COLORREF foreground = palette_.actionButtonText;

        if (disabled) {
            background = palette_.actionButtonDisabled;
            foreground = palette_.actionButtonDisabledText;
        } else if (isDefault && pressed) {
            background = palette_.defaultButtonPressed;
            foreground = palette_.defaultButtonText;
        } else if (isDefault && hot) {
            background = palette_.defaultButtonHover;
            foreground = palette_.defaultButtonText;
        } else if (isDefault) {
            background = palette_.defaultButtonBackground;
            foreground = palette_.defaultButtonText;
        } else if (pressed) {
            background = palette_.actionButtonPressed;
        } else if (hot) {
            background = palette_.actionButtonHover;
        }

        Gdiplus::Graphics graphics(draw.hDC);
        ConfigureGraphics(graphics);
        FillRoundedRect(graphics, rect, ScaleF(6.0f), background, background);

        RECT textRect = rect;
        textRect.left += Scale(12);
        textRect.right -= Scale(12);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, foreground);
        const auto previousFont = SelectObject(draw.hDC, buttonFont_);
        const auto text = ReadControlText(control);
        DrawTextW(
            draw.hDC,
            text.c_str(),
            -1,
            &textRect,
            DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
        SelectObject(draw.hDC, previousFont);

        if (focused) {
            RECT focusRect = rect;
            InflateRect(&focusRect, -Scale(4), -Scale(4));
            DrawFocusRect(draw.hDC, &focusRect);
        }
    }

    HBRUSH HandleButtonColor(HDC hdc, HWND control) {
        if (IsChoiceControl(control)) {
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, IsWindowEnabled(control) ? palette_.primaryText : palette_.secondaryText);
            return reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
        }

        SetBkColor(hdc, palette_.windowBackground);
        return backgroundBrush_;
    }

    HBRUSH HandleStaticColor(HDC hdc, HWND control) {
        if (control == statusBox_) {
            SetTextColor(hdc, palette_.primaryText);
            SetBkColor(hdc, palette_.inputBackground);
            return inputBrush_;
        }

        if (control == latitude_ || control == longitude_) {
            SetTextColor(hdc, status_.hasLocation ? palette_.primaryText : palette_.secondaryText);
            SetBkColor(hdc, palette_.inputBackground);
            return inputBrush_;
        }

        SetBkMode(hdc, TRANSPARENT);
        if (ContainsHandle(hintLabels_, control)) {
            SetTextColor(hdc, palette_.secondaryText);
        } else {
            SetTextColor(hdc, palette_.primaryText);
        }
        return reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
    }

    HBRUSH HandleEditColor(HDC hdc) {
        SetTextColor(hdc, palette_.primaryText);
        SetBkColor(hdc, palette_.inputBackground);
        return inputBrush_;
    }

    HBRUSH HandleListBoxColor(HDC hdc) {
        SetTextColor(hdc, palette_.primaryText);
        SetBkColor(hdc, palette_.inputBackground);
        return inputBrush_;
    }

    HWND CreateActionButton(int id, const wchar_t* text, int x, int y, int width, int height, bool isDefault) {
        const auto button = CreateWindowExW(
            0,
            L"BUTTON",
            text,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
            x,
            y,
            width,
            height,
            hwnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
            instance_,
            nullptr);
        SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(buttonFont_), TRUE);
        if (isDefault) {
            SendMessageW(hwnd_, DM_SETDEFID, id, 0);
        }
        return button;
    }

    HWND CreateCheckBox(int id, const wchar_t* text, int x, int y, int width, int height) {
        const auto box = CreateWindowExW(
            0,
            L"BUTTON",
            text,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_NOTIFY | BS_AUTOCHECKBOX,
            x,
            y,
            width,
            height,
            hwnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
            instance_,
            nullptr);
        SendMessageW(box, WM_SETFONT, reinterpret_cast<WPARAM>(bodyFont_), TRUE);
        return box;
    }

    HWND CreateModifierToggle(int id, const wchar_t* text, int x, int y, int width, int height) {
        const auto toggle = CreateWindowExW(
            0,
            L"BUTTON",
            text,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW | BS_NOTIFY,
            x,
            y,
            width,
            height,
            hwnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
            instance_,
            nullptr);
        SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(bodyFont_), TRUE);
        SetPropW(toggle, L"WinXSwModifierChecked", reinterpret_cast<HANDLE>(static_cast<INT_PTR>(FALSE)));
        return toggle;
    }

    HWND CreateRadio(int id, const wchar_t* text, int x, int y, int width, int height, bool startsGroup) {
        DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_NOTIFY | BS_AUTORADIOBUTTON;
        if (startsGroup) {
            style |= WS_GROUP;
        }

        const auto radio = CreateWindowExW(
            0,
            L"BUTTON",
            text,
            style,
            x,
            y,
            width,
            height,
            hwnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
            instance_,
            nullptr);
        SendMessageW(radio, WM_SETFONT, reinterpret_cast<WPARAM>(bodyFont_), TRUE);
        return radio;
    }

    HWND CreateEdit(int id, int x, int y, int width, int height) {
        const auto edit = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"EDIT",
            nullptr,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
            x,
            y,
            width,
            height,
            hwnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
            instance_,
            nullptr);
        SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(bodyFont_), TRUE);
        SendMessageW(edit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(Scale(8), Scale(8)));
        return edit;
    }

    HWND CreateTextLabel(
        const wchar_t* text,
        int x,
        int y,
        int width,
        int height,
        HFONT font,
        std::vector<HWND>& bucket) {
        const auto label = CreateWindowExW(
            WS_EX_TRANSPARENT,
            L"STATIC",
            text,
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            x,
            y,
            width,
            height,
            hwnd_,
            nullptr,
            instance_,
            nullptr);
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        bucket.push_back(label);
        return label;
    }

    HWND CreateFieldLabel(const wchar_t* text, int x, int y) {
        return CreateTextLabel(text, x, y, Scale(140), Scale(20), bodyFont_, bodyLabels_);
    }

    HWND CreateHintText(const wchar_t* text, int x, int y, int width) {
        return CreateTextLabel(text, x, y, width, Scale(20), hintFont_, hintLabels_);
    }

    void SetText(HWND window, const std::wstring& text) {
        SetWindowTextW(window, text.c_str());
    }

    void SetCheck(HWND window, bool checked) {
        if (IsModifierLabelControl(window)) {
            SetPropW(window, L"WinXSwModifierChecked", reinterpret_cast<HANDLE>(static_cast<INT_PTR>(checked ? TRUE : FALSE)));
            InvalidateRect(window, nullptr, TRUE);
            return;
        }
        SendMessageW(window, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
    }

    bool GetCheck(HWND window) const {
        if (IsModifierLabelControl(window)) {
            return reinterpret_cast<INT_PTR>(GetPropW(window, L"WinXSwModifierChecked")) != 0;
        }
        return SendMessageW(window, BM_GETCHECK, 0, 0) == BST_CHECKED;
    }

    void ToggleCheckbox(HWND window) {
        SetCheck(window, !GetCheck(window));
        SetFocus(window);
    }

    void PopulateHotkeyKeyCombo(HWND combo) {
        SendMessageW(combo, WM_SETFONT, reinterpret_cast<WPARAM>(bodyFont_), TRUE);
        for (const auto& entry : AvailableHotkeyKeys()) {
            const auto index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.second.c_str())));
            SendMessageW(combo, CB_SETITEMDATA, index, entry.first);
        }
    }

    void ApplyHotkeyToControls(
        const HotkeySettings& hotkey,
        HWND enable,
        HWND ctrl,
        HWND alt,
        HWND shift,
        HWND win,
        HWND combo) {
        SetCheck(enable, hotkey.enabled);
        SetCheck(ctrl, hotkey.ctrl);
        SetCheck(alt, hotkey.alt);
        SetCheck(shift, hotkey.shift);
        SetCheck(win, hotkey.win);
        SelectHotkeyKey(combo, hotkey.virtualKey);
    }

    bool CollectHotkeyFromControls(
        HotkeySettings& hotkey,
        HWND enable,
        HWND ctrl,
        HWND alt,
        HWND shift,
        HWND win,
        HWND combo,
        const wchar_t* displayName,
        std::wstring& error) const {
        hotkey.enabled = GetCheck(enable);
        hotkey.ctrl = GetCheck(ctrl);
        hotkey.alt = GetCheck(alt);
        hotkey.shift = GetCheck(shift);
        hotkey.win = GetCheck(win);

        if (hotkey.enabled &&
            !hotkey.ctrl &&
            !hotkey.alt &&
            !hotkey.shift &&
            !hotkey.win) {
            error = std::wstring(L"Enable at least one modifier for the ") + displayName + L" hotkey.";
            return false;
        }

        const auto selected = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
        if (selected == CB_ERR) {
            error = std::wstring(L"Choose a key for the ") + displayName + L" hotkey.";
            return false;
        }
        hotkey.virtualKey = static_cast<UINT>(SendMessageW(combo, CB_GETITEMDATA, selected, 0));
        return true;
    }

    static bool HotkeysConflict(const HotkeySettings& left, const HotkeySettings& right) {
        return left.enabled &&
            right.enabled &&
            left.ctrl == right.ctrl &&
            left.alt == right.alt &&
            left.shift == right.shift &&
            left.win == right.win &&
            left.virtualKey == right.virtualKey;
    }

    void UpdateHotkeyEnabledState(
        HWND enable,
        HWND ctrl,
        HWND alt,
        HWND shift,
        HWND win,
        HWND combo) {
        const auto enabled = GetCheck(enable);
        EnableWindow(ctrl, enabled);
        EnableWindow(alt, enabled);
        EnableWindow(shift, enabled);
        EnableWindow(win, enabled);
        EnableWindow(combo, enabled);
        InvalidateRect(ctrl, nullptr, TRUE);
        InvalidateRect(alt, nullptr, TRUE);
        InvalidateRect(shift, nullptr, TRUE);
        InvalidateRect(win, nullptr, TRUE);
    }

    void SelectHotkeyKey(HWND combo, UINT virtualKey) {
        const auto count = static_cast<int>(SendMessageW(combo, CB_GETCOUNT, 0, 0));
        for (int i = 0; i < count; ++i) {
            if (static_cast<UINT>(SendMessageW(combo, CB_GETITEMDATA, i, 0)) == virtualKey) {
                SendMessageW(combo, CB_SETCURSEL, i, 0);
                return;
            }
        }
        SendMessageW(combo, CB_SETCURSEL, 0, 0);
    }

    HINSTANCE instance_ = nullptr;
    HWND hwnd_ = nullptr;
    UINT dpi_ = USER_DEFAULT_SCREEN_DPI;
    ThemeMode uiThemeMode_ = ThemeMode::Light;
    UiPalette palette_{};
    Settings settings_{};
    Status status_{};

    HFONT sectionFont_ = nullptr;
    HFONT bodyFont_ = nullptr;
    HFONT bodyBoldFont_ = nullptr;
    HFONT hintFont_ = nullptr;
    HFONT buttonFont_ = nullptr;
    HFONT statusFont_ = nullptr;

    HBRUSH backgroundBrush_ = nullptr;
    HBRUSH cardBrush_ = nullptr;
    HBRUSH inputBrush_ = nullptr;
    HBRUSH accentBrush_ = nullptr;

    std::vector<HWND> headerLabels_;
    std::vector<HWND> sectionLabels_;
    std::vector<HWND> hintLabels_;
    std::vector<HWND> bodyLabels_;

    bool trackingMouse_ = false;
    bool closeHot_ = false;
    bool minimizeHot_ = false;
    bool closePressed_ = false;
    bool minimizePressed_ = false;
    bool coordinatesShowingManual_ = false;

    std::wstring manualLatitudeText_;
    std::wstring manualLongitudeText_;

    HWND automationTitle_ = nullptr;
    HWND automationHint_ = nullptr;
    HWND locationTitle_ = nullptr;
    HWND locationHint_ = nullptr;
    HWND hotkeyTitle_ = nullptr;
    HWND hotkeyHint_ = nullptr;
    HWND statusTitle_ = nullptr;
    HWND statusHint_ = nullptr;
    HWND wallpaperTitle_ = nullptr;
    HWND wallpaperHint_ = nullptr;

    HWND autoSwitch_ = nullptr;
    HWND autostart_ = nullptr;
    HWND locationAuto_ = nullptr;
    HWND locationManual_ = nullptr;
    HWND allowGps_ = nullptr;
    HWND allowIp_ = nullptr;
    HWND latitudeLabel_ = nullptr;
    HWND latitude_ = nullptr;
    HWND longitudeLabel_ = nullptr;
    HWND longitude_ = nullptr;
    HWND sunriseOffset_ = nullptr;
    HWND sunsetOffset_ = nullptr;
    HWND pollInterval_ = nullptr;
    HWND locationRefresh_ = nullptr;
    HWND hotkeyEnable_ = nullptr;
    HWND hotkeyCtrl_ = nullptr;
    HWND hotkeyAlt_ = nullptr;
    HWND hotkeyShift_ = nullptr;
    HWND hotkeyWin_ = nullptr;
    HWND hotkeyKey_ = nullptr;
    HWND configHotkeyEnable_ = nullptr;
    HWND configHotkeyCtrl_ = nullptr;
    HWND configHotkeyAlt_ = nullptr;
    HWND configHotkeyShift_ = nullptr;
    HWND configHotkeyWin_ = nullptr;
    HWND configHotkeyKey_ = nullptr;
    HWND statusBox_ = nullptr;
    HWND lightWallpaper_ = nullptr;
    HWND darkWallpaper_ = nullptr;
    HWND saveButton_ = nullptr;
    HWND reloadButton_ = nullptr;
    HWND toggleButton_ = nullptr;
    HWND launchButton_ = nullptr;
    HWND terminateAgentButton_ = nullptr;
    HWND openFolderButton_ = nullptr;
    HWND browseLightWallpaperButton_ = nullptr;
    HWND browseDarkWallpaperButton_ = nullptr;
};

}  // namespace

}  // namespace winxsw

int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    INITCOMMONCONTROLSEX controls{};
    controls.dwSize = sizeof(controls);
    controls.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
    InitCommonControlsEx(&controls);
    BufferedPaintInit();

    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken = 0;
    const auto gdiplusStatus = Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);
    if (gdiplusStatus != Gdiplus::Ok) {
        BufferedPaintUnInit();
        return 1;
    }

    winrt::init_apartment(winrt::apartment_type::single_threaded);

    winxsw::ConfigWindow window(instance);
    if (!window.Create()) {
        Gdiplus::GdiplusShutdown(gdiplusToken);
        BufferedPaintUnInit();
        return 1;
    }
    const auto exitCode = window.Run();
    Gdiplus::GdiplusShutdown(gdiplusToken);
    BufferedPaintUnInit();
    return exitCode;
}
