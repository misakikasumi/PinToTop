#include <sdkddkver.h>

#include <windows.h>

#include <appmodel.h>
#include <appxpackaging.h>
#include <comdef.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl_core.h>
#include <string.h>
#include <strsafe.h>
#include <wincodec.h>

#include <algorithm>
#include <filesystem>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "wil/com.h"
#include "wil/resource.h"
#include "wil/result.h"

#include "resource.h"

constexpr int MAX_LOADSTR = 260;
constexpr UINT UM_TRAY = WM_USER + 1;
HINSTANCE hInst;
HWND hWnd;
WCHAR app_title[MAX_LOADSTR];
WCHAR wnd_class[MAX_LOADSTR];
constexpr CLSID CLSID_ImmersiveShell = {
    0xC2F03A33, 0x21F5, 0x47FA, 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39};
IVirtualDesktopManager *vdm;

void load_resource();
void register_wndclass();
void make_window();
void init_tray();
void destroy_tray();
void init_hotkey();
void show_menu();
void toggle_top(HWND wnd);
std::vector<HWND> get_app_windows();
HICON get_window_icon(HWND);
HBITMAP get_uwp_icon(HWND);
HBITMAP icon_to_bitmap(HICON);
int main_loop();
HBITMAP load_img(const WCHAR *path, int w, int h);
void init_vdm();
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow) {
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);
  UNREFERENCED_PARAMETER(nCmdShow);

  hInst = hInstance;

  THROW_IF_FAILED(CoInitializeEx(nullptr, COINIT_MULTITHREADED));

  load_resource();
  register_wndclass();
  make_window();
  init_tray();
  init_vdm();
  init_hotkey();
  return main_loop();
}

int get_iconsm_metric() { return GetSystemMetrics(SM_CXSMICON); }

void load_resource() {
  THROW_LAST_ERROR_IF(
      LoadStringW(hInst, IDS_APP_TITLE, app_title, MAX_LOADSTR) == 0);
  THROW_LAST_ERROR_IF(
      LoadStringW(hInst, IDC_WNDCLASS, wnd_class, MAX_LOADSTR) == 0);
}

void register_wndclass() {
  WNDCLASSEXW wc{0};
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
  THROW_LAST_ERROR_IF_NULL(hWnd = CreateWindowW(wnd_class, app_title, 0, 0, 0,
                                                0, 0, HWND_MESSAGE, nullptr,
                                                hInst, nullptr));
}

void init_tray() {
  NOTIFYICONDATAW ncd{0};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  ncd.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE | NIF_SHOWTIP;
  const int cx = get_iconsm_metric();
  ncd.hIcon = HICON(THROW_LAST_ERROR_IF_NULL(
      LoadImageW(hInst, MAKEINTRESOURCEW(IDI_APP), IMAGE_ICON, cx, cx, 0)));
  ncd.uVersion = NOTIFYICON_VERSION_4;
  ncd.uCallbackMessage = UM_TRAY;
  THROW_IF_FAILED(
      StringCchCopyW(ncd.szTip, sizeof(ncd.szTip) / sizeof(WCHAR), app_title));
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_ADD, &ncd));
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_SETVERSION, &ncd));
}

void destroy_tray() {
  NOTIFYICONDATAW ncd{0};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  THROW_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_DELETE, &ncd));
}

void notify_error(UINT title_id, UINT info_id) {
  NOTIFYICONDATAW ncd{0};
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

void init_hotkey() {
  RegisterHotKey(hWnd, 0, MOD_CONTROL|MOD_ALT, 0x54);
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

void show_menu() {
  WCHAR exit_str[MAX_LOADSTR];
  THROW_LAST_ERROR_IF(LoadStringW(hInst, IDS_EXIT, exit_str, MAX_LOADSTR) == 0);
  wil::unique_hmenu menu{CreatePopupMenu()};
  THROW_LAST_ERROR_IF_NULL(menu);
  auto wnds{get_app_windows()};
  UINT id = 1;
  for (auto wnd : wnds) {
    WCHAR wnd_text[MAX_LOADSTR];
    wnd_text[GetWindowTextW(wnd, wnd_text, MAX_LOADSTR)] = 0;
    THROW_IF_WIN32_BOOL_FALSE(AppendMenuW(
        menu.get(), is_window_topmost(wnd) ? MF_CHECKED : 0, id, wnd_text));
    HBITMAP bitmap = get_uwp_icon(wnd);
    if (!bitmap) {
      HICON icon = get_window_icon(wnd);
      if (icon) {
        bitmap = icon_to_bitmap(icon);
      }
    }
    if (bitmap) {
      MENUITEMINFOW mii{0};
      mii.cbSize = sizeof(mii);
      mii.fMask = MIIM_CHECKMARKS;
      mii.hbmpUnchecked = bitmap;
      THROW_IF_WIN32_BOOL_FALSE(SetMenuItemInfoW(menu.get(), id, FALSE, &mii));
    }
    ++id;
  }
  THROW_IF_WIN32_BOOL_FALSE(
      AppendMenuW(menu.get(), MF_SEPARATOR, id + 1, nullptr));
  THROW_IF_WIN32_BOOL_FALSE(AppendMenuW(menu.get(), 0, id, exit_str));
  POINT pt;
  THROW_IF_WIN32_BOOL_FALSE(GetCursorPos(&pt));
  THROW_IF_WIN32_BOOL_FALSE(SetForegroundWindow(hWnd));
  int cmd =
      TrackPopupMenu(menu.get(), TPM_BOTTOMALIGN | TPM_NONOTIFY | TPM_RETURNCMD,
                     pt.x, pt.y, 0, hWnd, nullptr);
  if (cmd == id) {
    DestroyWindow(hWnd);
  } else if (cmd != 0) {
    toggle_top(wnds[cmd - 1]);
  }
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

void init_vdm() {
  wil::com_ptr<IServiceProvider> sp;
  THROW_IF_FAILED(CoCreateInstance(CLSID_ImmersiveShell, nullptr,
                                   CLSCTX_LOCAL_SERVER,
                                   __uuidof(IServiceProvider), (void **)&sp));
  THROW_IF_FAILED(sp->QueryService(__uuidof(IVirtualDesktopManager), &vdm));
}

bool is_app_window(HWND wnd) {
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

HBITMAP get_uwp_icon(HWND wnd) {
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
    return nullptr;
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
    return nullptr;
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
  wil::com_ptr<IAppxFactory> factory;
  THROW_IF_FAILED(CoCreateInstance(CLSID_AppxFactory, nullptr,
                                   CLSCTX_INPROC_SERVER, __uuidof(IAppxFactory),
                                   (void **)&factory));
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
  for (auto file : std::filesystem::recursive_directory_iterator{folder_path}) {
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
        if(!parse_modifiers(std::wstring_view{stem.extension().native()}.substr(1),
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
    return nullptr;
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
          if (lc->second == L"lightunplated" &&
              rc->second != L"lightunplated") {
            return true;
          }
          if (lc->second != L"lightunplated" &&
              rc->second == L"lightunplated") {
            return false;
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
          if (lc->second == L"light" &&
              rc->second != L"light") {
            return true;
          }
          if (lc->second != L"light" &&
              rc->second == L"light") {
            return false;
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

  return load_img(best->first.c_str(), cx, cx);
}

HBITMAP create_dib(LONG width, LONG height, void **pp_bits = nullptr) {
  BITMAPV5HEADER header{0};
  header.bV5Size = sizeof(BITMAPV5HEADER);
  header.bV5Width = width;
  header.bV5Height = height;
  header.bV5Planes = 1;
  header.bV5BitCount = 32;
  header.bV5Compression = BI_RGB;
  header.bV5CSType = LCS_sRGB;
  header.bV5Intent = LCS_GM_GRAPHICS;
  return CreateDIBSection(nullptr, (BITMAPINFO *)&header, DIB_RGB_COLORS,
                          pp_bits, nullptr, 0);
}

HBITMAP icon_to_bitmap(HICON icon) {
  wil::unique_hdc_window hdc{GetDC(nullptr)};
  wil::unique_hdc hMemDC{CreateCompatibleDC(hdc.get())};
  const int cx = get_iconsm_metric();
  HBITMAP hBitmap = create_dib(cx, cx);
  wil::unique_hgdiobj hOrgBmp{SelectObject(hMemDC.get(), hBitmap)};
  THROW_IF_WIN32_BOOL_FALSE(
      DrawIconEx(hMemDC.get(), 0, 0, icon, cx, cx, 0, nullptr, DI_NORMAL));
  SelectObject(hMemDC.get(), hOrgBmp.release());
  return hBitmap;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam,
                         LPARAM lParam) {
  static bool menu_showing = false;
  switch (message) {
  case UM_TRAY:
    if ((lParam == WM_LBUTTONUP || lParam == WM_RBUTTONUP) && !menu_showing) {
      try {
        menu_showing = true;
        show_menu();
        menu_showing = false;
      } catch (const std::exception &exc) {
        CHAR fatal_title[MAX_LOADSTR];
        LoadStringA(hInst, IDS_FATAL_MSGBOX_TITLE, fatal_title, MAX_LOADSTR);
        MessageBoxA(hWnd, exc.what(), fatal_title, MB_ICONERROR|MB_SYSTEMMODAL);
        throw;
      }
    }
    break;
  case WM_HOTKEY: {
    HWND foreground = GetForegroundWindow();
    if (foreground) {
      for (HWND owner; (owner = GetWindow(foreground, GW_OWNER)); foreground = owner);
      if (is_app_window(foreground)) {
        toggle_top(foreground);
      }
    }
    break;
  }
  case WM_DESTROY:
    PostQuitMessage(0);
    destroy_tray();
    break;
  default:
    return DefWindowProcW(hWnd, message, wParam, lParam);
  }
  return 0;
}

HBITMAP load_img(const WCHAR *path, int w, int h) {
  wil::com_ptr<IWICImagingFactory> factory;
  THROW_IF_FAILED(
      CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                       __uuidof(IWICImagingFactory), (void **)&factory));
  wil::com_ptr<IWICBitmapDecoder> decoder;
  THROW_IF_FAILED(factory->CreateDecoderFromFilename(
      path, nullptr, GENERIC_READ, WICDecodeMetadataCacheOnDemand, &decoder));
  wil::com_ptr<IWICBitmapFrameDecode> frame;
  THROW_IF_FAILED(decoder->GetFrame(0, &frame));
  wil::com_ptr<IWICFormatConverter> converter;
  THROW_IF_FAILED(factory->CreateFormatConverter(&converter));
  THROW_IF_FAILED(converter->Initialize(
      frame.get(), GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone,
      nullptr, 0, WICBitmapPaletteTypeCustom));
  wil::com_ptr<IWICColorContext> src_context;
  THROW_IF_FAILED(factory->CreateColorContext(&src_context));
  UINT icc_count;
  THROW_IF_FAILED(
      frame->GetColorContexts(1, src_context.addressof(), &icc_count));
  if (icc_count == 0) {
    THROW_IF_FAILED(src_context->InitializeFromExifColorSpace(1));
  }
  wil::com_ptr<IWICColorContext> dst_context;
  THROW_IF_FAILED(factory->CreateColorContext(&dst_context));
  THROW_IF_FAILED(dst_context->InitializeFromExifColorSpace(1));
  wil::com_ptr<IWICColorTransform> transformer;
  THROW_IF_FAILED(factory->CreateColorTransformer(&transformer));
  THROW_IF_FAILED(transformer->Initialize(converter.get(), src_context.get(),
                                          dst_context.get(),
                                          GUID_WICPixelFormat32bppBGRA));
  wil::com_ptr<IWICBitmapScaler> scaled;
  THROW_IF_FAILED(factory->CreateBitmapScaler(&scaled));
  THROW_IF_FAILED(scaled->Initialize(
      transformer.get(), w, h, WICBitmapInterpolationModeHighQualityCubic));
  void *p_bits;
  wil::unique_hbitmap bitmap{create_dib(w, -h, &p_bits)};
  THROW_IF_FAILED(
      scaled->CopyPixels(nullptr, w * 4, w * h * 4, (BYTE *)p_bits));
  return bitmap.release();
}
