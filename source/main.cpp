#include "pch.h"
#include "resource.h"
using namespace winrt;

constexpr int MAX_LOADSTR = 260;
constexpr UINT UM_TRAY = WM_USER + 1;
constexpr UINT UM_THEMECHANGED = WM_USER + 2;
constexpr UINT UM_SETMENUITEMICON = WM_USER + 3;
constexpr UINT UM_MENU_CLOSED = WM_USER + 4;
HINSTANCE hInst;
HWND hWnd;
WCHAR app_title[MAX_LOADSTR];
WCHAR wnd_class[MAX_LOADSTR];

Windows::UI::Xaml::Controls::TextBlock anchor{nullptr};
Windows::UI::Xaml::Controls::MenuFlyout menu_flyout{nullptr};
bool apps_use_dark_theme, system_uses_dark_theme;

void load_resource();
void init_theme();
void register_wndclass();
void make_window();
void init_tray(bool = false);
void destroy_tray();
void init_hotkey();
void init_island();
void init_icon_thread();
void show_menu();
void toggle_top(HWND wnd);
std::vector<HWND> get_app_windows();
HICON get_window_icon(HWND);
std::optional<std::wstring> get_uwp_icon_path(HWND);
void write_icon(HICON, IStream *);
int main_loop();
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow) {
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);
  UNREFERENCED_PARAMETER(nCmdShow);

  hInst = hInstance;

  init_apartment();
  load_resource();
  init_theme();
  register_wndclass();
  make_window();
  init_tray();
  init_hotkey();
  init_island();
  init_icon_thread();
  return main_loop();
}

int get_iconsm_metric() {
  auto dpi{GetDpiForSystem()};
  return GetSystemMetricsForDpi(SM_CXSMICON, dpi);
}

void load_resource() {
  THROW_LAST_ERROR_IF(
      LoadStringW(hInst, IDS_APP_TITLE, app_title, MAX_LOADSTR) == 0);
  THROW_LAST_ERROR_IF(
      LoadStringW(hInst, IDC_WNDCLASS, wnd_class, MAX_LOADSTR) == 0);
}

void register_wndclass() {
  WNDCLASSEXW wc{};
  wc.cbSize = sizeof(WNDCLASSEXW);
  wc.lpfnWndProc = WndProc;
  wc.hInstance = hInst;
  wc.hIcon =
      THROW_LAST_ERROR_IF_NULL(LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP)));
  wc.hCursor = THROW_LAST_ERROR_IF_NULL(LoadCursor(nullptr, IDC_ARROW));
  wc.hbrBackground = HBRUSH(COLOR_WINDOW + 1);
  wc.lpszClassName = wnd_class;
  const int cx = get_iconsm_metric();
  wc.hIconSm = HICON(THROW_LAST_ERROR_IF_NULL(
      LoadImageW(hInst, MAKEINTRESOURCEW(IDI_APP), IMAGE_ICON, cx, cx, 0)));
  THROW_LAST_ERROR_IF(!RegisterClassExW(&wc));
}

void make_window() {
  THROW_LAST_ERROR_IF_NULL(
      hWnd = CreateWindowExW(
          WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT | WS_EX_LAYERED, wnd_class,
          app_title, 0, 0, 0, 0, 0, nullptr, nullptr, hInst, nullptr));
  THROW_LAST_ERROR_IF(SetWindowLongW(hWnd, GWL_STYLE, 0) == 0);
  THROW_IF_WIN32_BOOL_FALSE(
      SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW));
}

void init_tray(bool reinit) {
  NOTIFYICONDATAW ncd{};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  ncd.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE | NIF_SHOWTIP;
  const int cx = get_iconsm_metric();
  ncd.hIcon = HICON(THROW_LAST_ERROR_IF_NULL(
      LoadImageW(hInst,
                 system_uses_dark_theme ? MAKEINTRESOURCEW(IDI_APP_THEME_DARK)
                                        : MAKEINTRESOURCEW(IDI_APP),
                 IMAGE_ICON, cx, cx, 0)));
  ncd.uVersion = NOTIFYICON_VERSION_4;
  ncd.uCallbackMessage = UM_TRAY;
  THROW_IF_FAILED(
      StringCchCopyW(ncd.szTip, sizeof(ncd.szTip) / sizeof(WCHAR), app_title));
  THROW_IF_WIN32_BOOL_FALSE(
      Shell_NotifyIconW(reinit ? NIM_MODIFY : NIM_ADD, &ncd));
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_SETVERSION, &ncd));
}

void destroy_tray() {
  NOTIFYICONDATAW ncd{};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_DELETE, &ncd));
}

void tray_return_focus() {
  NOTIFYICONDATAW ncd{};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_SETFOCUS, &ncd));
}

void notify_error(UINT title_id, UINT info_id) {
  NOTIFYICONDATAW ncd{};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  ncd.uFlags = NIF_INFO | NIF_REALTIME;
  THROW_LAST_ERROR_IF(LoadStringW(hInst, info_id, ncd.szInfo,
                                  sizeof(ncd.szInfo) / sizeof(WCHAR)) == 0);
  THROW_LAST_ERROR_IF(LoadStringW(hInst, title_id, ncd.szInfoTitle,
                                  sizeof(ncd.szInfoTitle) / sizeof(WCHAR)) ==
                      0);
  ncd.dwInfoFlags = NIIF_ERROR | NIIF_LARGE_ICON;
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_MODIFY, &ncd));
}

void init_hotkey() { RegisterHotKey(hWnd, 0, MOD_CONTROL | MOD_ALT, 0x54); }

void create_menu_flyout() {
  menu_flyout = Windows::UI::Xaml::Controls::MenuFlyout{};
  menu_flyout.Placement(Windows::UI::Xaml::Controls::Primitives::
                            FlyoutPlacementMode::TopEdgeAlignedLeft);
  Windows::UI::Xaml::Controls::Primitives::FlyoutBase::SetAttachedFlyout(
      anchor, menu_flyout);
  menu_flyout.Closed([](const auto &, const auto &) {
    SendNotifyMessageW(hWnd, UM_MENU_CLOSED, 0, 0);
  });
}

void init_island() {
  static Windows::UI::Xaml::Hosting::WindowsXamlManager xaml_manager{nullptr};
  static Windows::UI::Xaml::Hosting::DesktopWindowXamlSource desktop_source{
      nullptr};
  xaml_manager = Windows::UI::Xaml::Hosting::WindowsXamlManager::
      InitializeForCurrentThread();
  desktop_source = Windows::UI::Xaml::Hosting::DesktopWindowXamlSource{};
  auto interop = desktop_source.as<IDesktopWindowXamlSourceNative>();
  THROW_IF_FAILED(interop->AttachToWindow(hWnd));
  HWND hIsland;
  THROW_IF_FAILED(interop->get_WindowHandle(&hIsland));
  SetWindowPos(hIsland, 0, 0, 0, 0, 0, SWP_SHOWWINDOW);
  anchor = Windows::UI::Xaml::Controls::TextBlock{};
  desktop_source.Content(anchor);
  create_menu_flyout();
}

constexpr WCHAR reg_theme_path[] =
    L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";

void get_theme() {
  wil::unique_hkey key;
  THROW_IF_WIN32_ERROR(RegOpenKeyExW(HKEY_CURRENT_USER, reg_theme_path, 0,
                                     KEY_QUERY_VALUE | KEY_WOW64_64KEY, &key));
  DWORD use_light_theme;
  DWORD buf_len;
  buf_len = sizeof(DWORD);
  THROW_IF_WIN32_ERROR(RegQueryValueExW(key.get(), L"AppsUseLightTheme",
                                        nullptr, nullptr,
                                        (BYTE *)&use_light_theme, &buf_len));
  apps_use_dark_theme = !use_light_theme;
  buf_len = sizeof(DWORD);
  THROW_IF_WIN32_ERROR(RegQueryValueExW(key.get(), L"SystemUsesLightTheme",
                                        nullptr, nullptr,
                                        (BYTE *)&use_light_theme, &buf_len));
  system_uses_dark_theme = !use_light_theme;
}

void init_theme() {
  static wil::unique_registry_watcher watcher;
  get_theme();
  watcher = wil::make_registry_watcher(
      HKEY_CURRENT_USER, reg_theme_path, false, [](auto) {
        auto previous_apps_use_dark_theme = apps_use_dark_theme;
        auto previous_system_uses_dark_theme = system_uses_dark_theme;
        get_theme();
        apps_use_dark_theme = previous_apps_use_dark_theme;
        WPARAM wParam =
            MAKELONG(apps_use_dark_theme != previous_apps_use_dark_theme,
                     system_uses_dark_theme != previous_system_uses_dark_theme);
        if (wParam) {
          THROW_IF_WIN32_BOOL_FALSE(
              SendNotifyMessageW(hWnd, UM_THEMECHANGED, wParam, 0));
        }
      });
}

int main_loop() {
  MSG msg;
  while (GetMessageW(&msg, nullptr, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessageW(&msg);
  }
  return int(msg.wParam);
}

bool is_window_topmost(HWND wnd) {
  return GetWindowLongW(wnd, GWL_EXSTYLE) & WS_EX_TOPMOST;
}

std::queue<std::pair<HWND, Windows::UI::Xaml::Controls::MenuFlyoutItem>> itq;
std::queue<std::pair<std::wstring, Windows::UI::Xaml::Controls::MenuFlyoutItem>>
    itbq;
std::condition_variable itcv;
std::mutex itmutex;

void show_menu(int x, int y) {
  THROW_IF_WIN32_BOOL_FALSE(
      SetWindowPos(hWnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE));
  if (!SetForegroundWindow(hWnd)) {
    return;
  }
  WCHAR exit_str[MAX_LOADSTR];
  THROW_LAST_ERROR_IF(LoadStringW(hInst, IDS_EXIT, exit_str, MAX_LOADSTR) == 0);
  auto menu_items{menu_flyout.Items()};
  menu_items.Clear();
  auto wnds{get_app_windows()};
  {
    std::scoped_lock lck{itmutex};
    for (auto wnd : wnds) {
      WCHAR wnd_text[MAX_LOADSTR];
      wnd_text[GetWindowTextW(wnd, wnd_text, MAX_LOADSTR)] = 0;
      Windows::UI::Xaml::Controls::ToggleMenuFlyoutItem item;
      Windows::UI::Xaml::Controls::BitmapIcon icon;
      icon.ShowAsMonochrome(false);
      item.Text(wnd_text);
      item.Icon(icon);
      item.IsChecked(is_window_topmost(wnd));
      item.Click([wnd](const auto &, const auto &) { toggle_top(wnd); });
      menu_items.Append(item);

      itq.push({wnd, std::move(item)});
    }
    itcv.notify_one();
  }
  if (!wnds.empty()) {
    menu_items.Append(Windows::UI::Xaml::Controls::MenuFlyoutSeparator{});
  }
  Windows::UI::Xaml::Controls::MenuFlyoutItem exit_item;
  exit_item.Text(exit_str);
  exit_item.Click([](const auto &, const auto &) { DestroyWindow(hWnd); });
  menu_items.Append(exit_item);
  Windows::UI::Xaml::Controls::Primitives::FlyoutBase::ShowAttachedFlyout(
      anchor);
}

std::optional<std::wstring> get_window_icon_uri(HWND wnd) {
  if (!IsWindow(wnd)) {
    return std::nullopt;
  }
  auto uwp_icon_path{get_uwp_icon_path(wnd)};
  if (uwp_icon_path) {
    return *uwp_icon_path;
  } else {
    auto hicon{get_window_icon(wnd)};
    if (hicon) {
      WCHAR temp_path[MAX_LOADSTR], temp_file[MAX_LOADSTR];
      THROW_LAST_ERROR_IF(GetTempPathW(MAX_LOADSTR, temp_path) == 0);
      THROW_IF_FAILED(StringCchPrintfW(temp_file, MAX_LOADSTR, L"%ws%p.png",
                                       temp_path, hicon));
      {
        wil::com_ptr<IStream> stream;
        auto hr = SHCreateStreamOnFileEx(
            temp_file, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_FAILIFTHERE,
            FILE_ATTRIBUTE_NORMAL, TRUE, nullptr, &stream);
        if (SUCCEEDED(hr)) {
          write_icon(hicon, stream.get());
        } else if ((hr & 0xffff) != ERROR_FILE_EXISTS) {
          THROW_HR(hr);
        }
      }
      return temp_file;
    }
  }
  return std::nullopt;
}

void init_icon_thread() {
  std::thread{[] {
    while (true) {
      std::unique_lock lck{itmutex};
      itcv.wait(lck);
      while (!itq.empty()) {
        auto wnd{itq.front().first};
        auto item{std::move(itq.front().second)};
        itq.pop();
        auto uri{get_window_icon_uri(wnd)};
        if (uri) {
          itbq.push({std::move(*uri), std::move(item)});
        }
      }
      THROW_IF_WIN32_BOOL_FALSE(
          SendNotifyMessage(hWnd, UM_SETMENUITEMICON, 0, 0));
    }
  }}.detach();
}

void toggle_top(HWND wnd) {
  if (!SetWindowPos(wnd, is_window_topmost(wnd) ? HWND_NOTOPMOST : HWND_TOPMOST,
                    0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)) {
    DWORD err = GetLastError();
    switch (err) {
    case 5:
      notify_error(IDS_FAIL_INFOTITLE, IDS_WND_ACCESS_DENIED_INFO);
      break;
    case 1400:
      notify_error(IDS_FAIL_INFOTITLE, IDS_WND_CLOSED_INFO);
      break;
    default:
      THROW_WIN32(err);
      break;
    }
  }
}

bool is_app_window(HWND wnd) {
  static IVirtualDesktopManager *vdm;
  if (!vdm) {
    THROW_IF_FAILED(CoCreateInstance(
        __uuidof(VirtualDesktopManager), nullptr, CLSCTX_INPROC_SERVER,
        __uuidof(IVirtualDesktopManager), (void **)&vdm));
  }
  LONG ex_sty = GetWindowLongW(wnd, GWL_EXSTYLE);
  WCHAR wnd_class[MAX_LOADSTR];
  wnd_class[RealGetWindowClassW(wnd, wnd_class, MAX_LOADSTR)] = 0;
  bool is_app = IsWindowVisible(wnd) &&
                ((ex_sty & WS_EX_APPWINDOW) ||
                 (!GetWindow(wnd, GW_OWNER) && !(ex_sty & WS_EX_NOACTIVATE) &&
                  !(ex_sty & WS_EX_TOOLWINDOW) && GetWindowTextLengthW(wnd) &&
                  wcscmp(wnd_class, L"Windows.UI.Core.CoreWindow")));
  BOOL is_on_current_vd;
  THROW_IF_FAILED(vdm->IsWindowOnCurrentVirtualDesktop(wnd, &is_on_current_vd));
  return is_app && is_on_current_vd;
}

std::vector<HWND> get_app_windows() {
  std::vector<HWND> wnds;
  EnumWindows(
      [](HWND wnd, LPARAM param) -> BOOL {
        std::vector<HWND> &dsks = *(std::vector<HWND> *)(param);
        if (is_app_window(wnd)) {
          dsks.push_back(wnd);
        }
        return TRUE;
      },
      LPARAM(&wnds));
  return wnds;
}

HICON get_window_icon(HWND wnd) {
  auto icon = HICON(GetClassLongPtrW(wnd, -34));
  if (icon) {
    return icon;
  }
  icon = HICON(SendMessageW(wnd, WM_GETICON, ICON_SMALL2, 0));
  if (icon) {
    return icon;
  }
  icon = HICON(GetClassLongPtrW(wnd, -14));
  return icon;
}

bool parse_modifiers(std::wstring_view s,
                     std::unordered_map<std::wstring, std::wstring> &mod) {
  bool finished = false;
  while (!finished) {
    auto delimiter = s.find(L"_");
    std::wstring_view sub{s};
    if (delimiter != std::wstring::npos) {
      sub = sub.substr(0, delimiter);
      s = s.substr(delimiter + 1);
    } else {
      finished = true;
    }
    auto pos = sub.find(L"-");
    if (pos == std::wstring_view::npos) {
      return false;
    }
    std::wstring name{sub.substr(0, pos)};
    std::transform(name.begin(), name.end(), name.begin(),
                   [](wchar_t ch) { return towlower(ch); });
    std::wstring value{sub.substr(pos + 1)};
    std::transform(value.begin(), value.end(), value.begin(),
                   [](wchar_t ch) { return towlower(ch); });
    mod[name] = value;
  }
  return true;
}

std::optional<std::wstring> get_uwp_icon_path(HWND wnd) {
  DWORD pid;
  GetWindowThreadProcessId(wnd, &pid);
  wil::unique_handle p{
      OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid)};
  THROW_LAST_ERROR_IF_NULL(p);
  WCHAR exename[MAX_LOADSTR];
  DWORD exelen = MAX_LOADSTR;
  THROW_IF_WIN32_BOOL_FALSE(
      QueryFullProcessImageNameW(p.get(), 0, exename, &exelen));
  if (_wcsicmp(exename, L"C:\\Windows\\System32\\ApplicationFrameHost.exe") !=
      0) {
    return std::nullopt;
  }
  DWORD real_pid = pid;
  EnumChildWindows(
      wnd,
      [](HWND wnd, LPARAM param) -> BOOL {
        auto p_real_pid = (DWORD *)param;
        DWORD pid;
        GetWindowThreadProcessId(wnd, &pid);
        if (pid != *p_real_pid) {
          *p_real_pid = pid;
          return FALSE;
        } else {
          return TRUE;
        }
      },
      LPARAM(&real_pid));
  if (real_pid == pid) {
    return L"";
  }
  p.reset(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, real_pid));
  THROW_LAST_ERROR_IF_NULL(p);
  UINT32 buflen = 0;
  LONG r = GetPackageId(p.get(), &buflen, nullptr);
  std::vector<BYTE> pkg_id_buf;
  if (r == ERROR_INSUFFICIENT_BUFFER) {
    pkg_id_buf.resize(buflen);
    THROW_IF_WIN32_ERROR(GetPackageId(p.get(), &buflen, &pkg_id_buf[0]));
  } else
    THROW_IF_WIN32_ERROR(r);
  auto pkg_id = (PACKAGE_ID *)&pkg_id_buf[0];
  WCHAR path[MAX_LOADSTR];
  WCHAR manifest_path[MAX_LOADSTR];
  buflen = MAX_LOADSTR;
  THROW_IF_WIN32_ERROR(GetPackagePath(pkg_id, 0, &buflen, path));
  THROW_IF_FAILED(StringCchCopyW(manifest_path, MAX_LOADSTR, path));
  THROW_IF_FAILED(
      StringCchCatW(manifest_path, MAX_LOADSTR, L"\\AppxManifest.xml"));
  wil::com_ptr<IStream> is;
  THROW_IF_FAILED(
      SHCreateStreamOnFileEx(manifest_path, STGM_READ, 0, 0, 0, &is));
  static IAppxFactory *factory;
  if (!factory) {
    THROW_IF_FAILED(
        CoCreateInstance(CLSID_AppxFactory, nullptr, CLSCTX_INPROC_SERVER,
                         __uuidof(IAppxFactory), (void **)&factory));
  }
  wil::com_ptr<IAppxManifestReader> reader;
  THROW_IF_FAILED(factory->CreateManifestReader(is.get(), &reader));
  wil::com_ptr<IAppxManifestApplicationsEnumerator> iter;
  THROW_IF_FAILED(reader->GetApplications(&iter));
  wil::com_ptr<IAppxManifestApplication> app;
  THROW_IF_FAILED(iter->GetCurrent(&app));
  WCHAR *logo;
  THROW_IF_FAILED(app->GetStringValue(L"Square44x44Logo", &logo));
  std::filesystem::path img_path{std::wstring(path) + L"\\" + logo};
  auto logo_stw{img_path.stem().native() + L"."};
  auto folder_path{img_path.parent_path()};
  std::vector<std::pair<std::filesystem::path,
                        std::unordered_map<std::wstring, std::wstring>>>
      candidates;
  for (const auto &file :
       std::filesystem::recursive_directory_iterator{folder_path}) {
    if (file.is_regular_file() &&
        std::wstring_view{file.path().filename().native()}.substr(
            0, logo_stw.size()) == logo_stw) {
      std::filesystem::path loc{
          file.path().native().substr(folder_path.native().size() + 1)};
      std::unordered_map<std::wstring, std::wstring> modifiers;
      bool valid = true;
      for (auto it{loc.begin()}; it != loc.end(); ++it) {
        auto it2{it};
        if (++it2 == loc.end()) {
          break;
        }
        if (!parse_modifiers(it->native(), modifiers)) {
          valid = false;
          break;
        }
      }
      auto stem{loc.stem()};
      if (valid && stem.has_extension()) {
        if (!parse_modifiers(
                std::wstring_view{stem.extension().native()}.substr(1),
                modifiers)) {
          valid = false;
        }
      }
      if (valid) {
        candidates.push_back({file.path().native(), std::move(modifiers)});
      }
    }
  }
  if (candidates.empty()) {
    return std::nullopt;
  }

  HIGHCONTRASTW hc;
  hc.cbSize = sizeof(HIGHCONTRASTW);
  SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(HIGHCONTRASTW), (void *)&hc,
                        0);
  int hc_mode;
  if (hc.dwFlags & HCF_HIGHCONTRASTON) {
    if (wcsstr(hc.lpszDefaultScheme, L"White") != nullptr) {
      hc_mode = 2;
    } else {
      hc_mode = 1;
    }
  } else {
    hc_mode = 0;
  }

  int cx = get_iconsm_metric();

  auto best{std::min_element(
      candidates.begin(), candidates.end(),
      [hc_mode, cx](const auto &left, const auto &right) {
        const auto &lm{left.second}, &rm{right.second};
        auto lc{lm.find(L"contrast")}, rc{rm.find(L"contrast")};
        if (hc_mode == 1 && lc != lm.end() && lc->second == L"black" &&
            (rc == rm.end() || rc->second != L"black")) {
          return true;
        }
        if (hc_mode == 1 && rc != rm.end() && rc->second == L"black" &&
            (lc == lm.end() || lc->second != L"black")) {
          return false;
        }
        if (hc_mode == 2 && lc != lm.end() && lc->second == L"white" &&
            (rc == rm.end() || rc->second != L"white")) {
          return true;
        }
        if (hc_mode == 2 && rc != rm.end() && rc->second == L"white" &&
            (lc == lm.end() || lc->second != L"white")) {
          return false;
        }
        if (hc_mode == 0 && (lc == lm.end() || lc->second == L"standard") &&
            rc != rm.end() && rc->second != L"standard") {
          return true;
        }
        if (hc_mode == 0 && (rc == rm.end() || rc->second == L"standard") &&
            lc != lm.end() && lc->second != L"standard") {
          return false;
        }
        lc = lm.find(L"altform");
        rc = rm.find(L"altform");
        if (lc != lm.end() && rc == rm.end()) {
          return true;
        }
        if (lc == lm.end() && rc != rm.end()) {
          return false;
        }
        if (lc != lm.end() && rc != rm.end()) {
          if (!apps_use_dark_theme) {
            if (lc->second == L"lightunplated" &&
                rc->second != L"lightunplated") {
              return true;
            }
            if (lc->second != L"lightunplated" &&
                rc->second == L"lightunplated") {
              return false;
            }
          } else {
            if (lc->second == L"unplated" && rc->second != L"unplated") {
              return true;
            }
            if (lc->second != L"unplated" && rc->second == L"unplated") {
              return false;
            }
          }
        }
        lc = lm.find(L"theme");
        rc = rm.find(L"theme");
        if (lc != lm.end() && rc == rm.end()) {
          return true;
        }
        if (lc == lm.end() && rc != rm.end()) {
          return false;
        }
        if (lc != lm.end() && rc != rm.end()) {
          if (!apps_use_dark_theme) {
            if (lc->second == L"light" && rc->second != L"light") {
              return true;
            }
            if (lc->second != L"light" && rc->second == L"light") {
              return false;
            }
          } else {
            if (lc->second == L"dark" && rc->second != L"dark") {
              return true;
            }
            if (lc->second != L"dark" && rc->second == L"dark") {
              return false;
            }
          }
        }
        lc = lm.find(L"targetsize");
        rc = rm.find(L"targetsize");
        if (lc != lm.end() && rc == rm.end()) {
          return true;
        }
        if (lc == lm.end() && rc != rm.end()) {
          return false;
        }
        if (lc != lm.end() && rc != rm.end()) {
          int lx = std::stoi(lc->second);
          int rx = std::stoi(rc->second);
          if (lx >= cx && rx < cx) {
            return true;
          }
          if (lx < cx && rx >= cx) {
            return false;
          }
          int ld = lx - cx;
          int rd = rx - cx;
          ld = ld < 0 ? -ld : ld;
          rd = rd < 0 ? -rd : rd;
          return ld < rd;
        }
        lc = lm.find(L"scale");
        rc = rm.find(L"scale");
        if (lc != lm.end() && rc == rm.end()) {
          return true;
        }
        if (lc == lm.end() && rc != rm.end()) {
          return false;
        }
        if (lc != lm.end() && rc != rm.end()) {
          int lx = std::stoi(lc->second);
          int rx = std::stoi(rc->second);
          return lx < rx;
        }
        return lm.size() < rm.size();
      })};

  return best->first;
}

void write_icon(HICON icon, IStream *stream) {
  static IWICImagingFactory *factory;
  if (!factory) {
    THROW_IF_FAILED(
        CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                         __uuidof(IWICImagingFactory), (void **)&factory));
  }
  wil::com_ptr<IWICBitmap> source;
  THROW_IF_FAILED(factory->CreateBitmapFromHICON(icon, &source));
  wil::com_ptr<IWICBitmapEncoder> encoder;
  THROW_IF_FAILED(
      factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, &encoder));
  THROW_IF_FAILED(encoder->Initialize(stream, WICBitmapEncoderNoCache));
  wil::com_ptr<IWICBitmapFrameEncode> frame;
  THROW_IF_FAILED(encoder->CreateNewFrame(&frame, nullptr));
  THROW_IF_FAILED(frame->Initialize(nullptr));
  THROW_IF_FAILED(frame->WriteSource(source.get(), nullptr));
  THROW_IF_FAILED(frame->Commit());
  THROW_IF_FAILED(encoder->Commit());
}

LRESULT CALLBACK WndProc(HWND thisHWnd, UINT message, WPARAM wParam,
                         LPARAM lParam) {
  try {
    if (thisHWnd == hWnd) {
      switch (message) {
      case UM_TRAY: {
        static bool menu_showing = false;
        if ((lParam == WM_LBUTTONUP || lParam == WM_RBUTTONUP) &&
            !menu_showing) {
          menu_showing = true;
          show_menu(GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam));
          menu_showing = false;
        }
        break;
      }
      case WM_HOTKEY: {
        HWND foreground = GetForegroundWindow();
        if (foreground) {
          for (HWND owner; (owner = GetWindow(foreground, GW_OWNER));
               foreground = owner)
            ;
          if (is_app_window(foreground)) {
            toggle_top(foreground);
          }
        }
        break;
      }
      case UM_THEMECHANGED:
        if (LOWORD(wParam)) {
          create_menu_flyout();
        }
        if (HIWORD(wParam)) {
          init_tray(true);
        }
        break;
      case UM_SETMENUITEMICON: {
        std::scoped_lock lck{itmutex};
        while (!itbq.empty()) {
          auto uri{std::move(itbq.front().first)};
          auto item{std::move(itbq.front().second)};
          itbq.pop();
          if (!uri.empty()) {
            auto icon{
                item.Icon().as<Windows::UI::Xaml::Controls::BitmapIcon>()};
            icon.UriSource(Windows::Foundation::Uri(uri));
          } else {
            Windows::UI::Xaml::Controls::FontIcon icon;
            icon.FontFamily(
                Windows::UI::Xaml::Media::FontFamily(L"Segoe MDL2 Assets"));
            icon.Glyph(L"\uE8BE");
            Windows::UI::Color green;
            green.R = green.B = 0;
            green.G = 128;
            green.A = 255;
            icon.Foreground(Windows::UI::Xaml::Media::SolidColorBrush(green));
            item.Icon(icon);
          }
        }
        break;
      }
      case UM_MENU_CLOSED:
        menu_flyout.Items().Clear();
        try {
          tray_return_focus();
        } catch (...) {
        }
        break;
      case WM_DPICHANGED:
        init_tray(true);
        break;
      case WM_DESTROY:
        PostQuitMessage(0);
        destroy_tray();
        break;
      default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
      }
    } else {
      return DefWindowProcW(thisHWnd, message, wParam, lParam);
    }
  } catch (const std::exception &exc) {
    CHAR fatal_title[MAX_LOADSTR];
    LoadStringA(hInst, IDS_FATAL_MSGBOX_TITLE, fatal_title, MAX_LOADSTR);
    MessageBoxA(hWnd, exc.what(), fatal_title, MB_ICONERROR | MB_SYSTEMMODAL);
    throw;
  }
  return 0;
}
