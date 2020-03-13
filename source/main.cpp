#include <sdkddkver.h>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <appmodel.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <string.h>
#include <strsafe.h>
#include <wincodec.h>
#include <xmllite.h>

#include <string>
#include <vector>

#include "wil/com.h"
#include "wil/resource.h"
#include "wil/result.h"

#include "resource.h"

constexpr int MAX_LOADSTR = 128;
constexpr UINT UM_TRAY = WM_USER + 1;
HINSTANCE hInst;
HWND hWnd;
WCHAR app_title[MAX_LOADSTR];
WCHAR wnd_class[MAX_LOADSTR];

void load_resource();
void register_wndclass();
void make_window();
void init_tray();
void destroy_tray();
void show_menu();
void toggle_top(HWND wnd);
std::vector<HWND> get_app_windows();
HICON get_window_icon(HWND);
HBITMAP get_uwp_icon(HWND);
HBITMAP icon_to_bitmap(HICON);
int main_loop();
HBITMAP load_img(const WCHAR *path, int w, int h);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow) {
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);
  UNREFERENCED_PARAMETER(nCmdShow);

  hInst = hInstance;

  CoInitialize(nullptr);

  load_resource();
  register_wndclass();
  make_window();
  init_tray();
  return main_loop();
}

void load_resource() {
  LoadStringW(hInst, IDS_APP_TITLE, app_title, MAX_LOADSTR);
  LoadStringW(hInst, IDC_WNDCLASS, wnd_class, MAX_LOADSTR);
}

void register_wndclass() {
  WNDCLASSEXW wc{0};
  wc.cbSize = sizeof(WNDCLASSEXW);
  wc.lpfnWndProc = WndProc;
  wc.hInstance = hInst;
  wc.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_APP));
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wc.hbrBackground = HBRUSH(COLOR_WINDOW + 1);
  wc.lpszClassName = wnd_class;
  int cx = GetSystemMetrics(SM_CXSMICON);
  wc.hIconSm =
      HICON(LoadImage(hInst, MAKEINTRESOURCE(IDI_APP), IMAGE_ICON, cx, cx, 0));
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
  int cx = GetSystemMetrics(SM_CXSMICON);
  ncd.hIcon =
      HICON(LoadImage(hInst, MAKEINTRESOURCE(IDI_APP), IMAGE_ICON, cx, cx, 0));
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

void notify(UINT title_id, UINT info_id) {
  NOTIFYICONDATAW ncd{0};
  ncd.cbSize = sizeof(ncd);
  ncd.hWnd = hWnd;
  ncd.uFlags = NIF_INFO | NIF_REALTIME;
  LoadStringW(hInst, info_id, ncd.szInfo, sizeof(ncd.szInfo) / sizeof(WCHAR));
  LoadStringW(hInst, title_id, ncd.szInfoTitle,
              sizeof(ncd.szInfoTitle) / sizeof(WCHAR));
  ncd.dwInfoFlags = NIIF_ERROR | NIIF_LARGE_ICON;
  Shell_NotifyIconW(NIM_MODIFY, &ncd);
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
  LoadStringW(hInst, IDS_EXIT, exit_str, MAX_LOADSTR);

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
      notify(IDS_FAIL_INFOTITLE, IDS_WND_ACCESS_DENIED_INFO);
      break;
    case 1400:
      notify(IDS_FAIL_INFOTITLE, IDS_WND_CLOSED_INFO);
      break;
    default:
      THROW_WIN32(err);
      break;
    }
  }
}

bool is_app_window(HWND wnd) {
  LONG ex_sty = GetWindowLongW(wnd, GWL_EXSTYLE);
  WCHAR wnd_class[MAX_LOADSTR];
  wnd_class[RealGetWindowClassW(wnd, wnd_class, MAX_LOADSTR)] = 0;
  return IsWindowVisible(wnd) &&
         ((ex_sty & WS_EX_APPWINDOW) ||
          (!GetWindow(wnd, GW_OWNER) && !(ex_sty & WS_EX_NOACTIVATE) &&
           !(ex_sty & WS_EX_TOOLWINDOW) && GetWindowTextLengthW(wnd) &&
           wcscmp(wnd_class, L"Windows.UI.Core.CoreWindow")));
}

std::vector<HWND> get_app_windows() {
  std::vector<HWND> wnds;
  WNDENUMPROC cb{[](HWND wnd, LPARAM param) -> BOOL {
    std::vector<HWND> &dsks = *(std::vector<HWND> *)(param);
    if (is_app_window(wnd)) {
      dsks.push_back(wnd);
    }
    return 1;
  }};
  EnumWindows(cb, LPARAM(&wnds));
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
  EnumChildWindows(wnd, WNDENUMPROC([](HWND wnd, LPARAM param) -> BOOL {
                     auto p_real_pid = (DWORD *)param;
                     DWORD pid;
                     GetWindowThreadProcessId(wnd, &pid);
                     if (pid != *p_real_pid) {
                       *p_real_pid = pid;
                       return FALSE;
                     } else {
                       return TRUE;
                     }
                   }),
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
  WCHAR img_path[MAX_LOADSTR];
  buflen = MAX_LOADSTR;
  THROW_IF_WIN32_ERROR(GetPackagePath(pkg_id, 0, &buflen, path));
  THROW_IF_FAILED(StringCchCopyW(img_path, MAX_LOADSTR, path));
  THROW_IF_FAILED(StringCchCatW(path, MAX_LOADSTR, L"\\AppxManifest.xml"));
  THROW_IF_FAILED(StringCchCatW(img_path, MAX_LOADSTR, L"\\"));
  wil::com_ptr<IStream> is;
  THROW_IF_FAILED(SHCreateStreamOnFileEx(path, STGM_READ, 0, 0, 0, &is));
  wil::com_ptr<IXmlReader> rd;
  THROW_IF_FAILED(CreateXmlReader(__uuidof(IXmlReader), (void **)&rd, nullptr));
  THROW_IF_FAILED(rd->SetInput(is.get()));
  bool in_logo = false, suc = false;
  for (XmlNodeType nt; rd->Read(&nt) == S_OK;) {
    if (nt == XmlNodeType_Element) {
      WCHAR *local_name;
      THROW_IF_FAILED(rd->GetLocalName((LPCWSTR *)&local_name, nullptr));
      if (wcscmp(local_name, L"Logo") == 0) {
        in_logo = true;
      }
    } else if (in_logo && nt == XmlNodeType_Text) {
      WCHAR *value;
      THROW_IF_FAILED(rd->GetValue((LPCWSTR *)&value, nullptr));
      StringCchCatW(img_path, MAX_LOADSTR, value);
      suc = true;
      break;
    }
  }
  if (!suc) {
    return nullptr;
  }
  int cx = GetSystemMetrics(SM_CXSMICON);
  return load_img(img_path, cx, cx);
}

HBITMAP create_dib(LONG width, LONG height, void **pp_bits = nullptr) {
  BITMAPV5HEADER header{0};
  header.bV5Size = sizeof(BITMAPV5HEADER);
  header.bV5Width = width;
  header.bV5Height = height;
  header.bV5Planes = 1;
  header.bV5BitCount = 32;
  header.bV5Compression = BI_RGB;
  header.bV5CSType = LCS_WINDOWS_COLOR_SPACE;
  header.bV5Intent = LCS_GM_GRAPHICS;
  return CreateDIBSection(nullptr, (BITMAPINFO *)&header, DIB_RGB_COLORS,
                          pp_bits, nullptr, 0);
}

HBITMAP icon_to_bitmap(HICON icon) {
  wil::unique_hdc_window hdc{GetDC(nullptr)};
  wil::unique_hdc hMemDC{CreateCompatibleDC(hdc.get())};
  int cx = GetSystemMetrics(SM_CXSMICON);
  HBITMAP hBitmap = create_dib(cx, cx);
  wil::unique_hgdiobj hOrgBmp{SelectObject(hMemDC.get(), hBitmap)};
  THROW_IF_WIN32_BOOL_FALSE(
      DrawIconEx(hMemDC.get(), 0, 0, icon, cx, cx, 0, nullptr, DI_NORMAL));
  SelectObject(hMemDC.get(), hOrgBmp.release());
  return hBitmap;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam,
                         LPARAM lParam) {
  switch (message) {
  case UM_TRAY:
    if (lParam == WM_LBUTTONUP || lParam == WM_RBUTTONUP) {
      show_menu();
    }
    break;
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
  std::wstring match(path);
  size_t dot_pos = match.rfind(L'.');
  THROW_HR_IF(E_INVALIDARG, dot_pos == std::wstring::npos);
  size_t slash_pos = match.rfind(L'\\');
  THROW_HR_IF(E_INVALIDARG, slash_pos == std::wstring::npos);
  match.insert(dot_pos, L"*");
  WIN32_FIND_DATAW fd;
  wil::unique_handle sh{FindFirstFileW(match.c_str(), &fd)};
  THROW_LAST_ERROR_IF_NULL(sh);
  wil::unique_hfind sq{sh.release()};
  std::vector<std::wstring> imgs;
  do {
    imgs.push_back(fd.cFileName);
  } while (FindNextFileW(sq.get(), &fd));
  std::wstring found{match.substr(0, slash_pos + 1) + imgs.at(0)};
  const WCHAR *found_path = found.c_str();
  wil::com_ptr<IWICImagingFactory> factory;
  THROW_IF_FAILED(
      CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                       __uuidof(IWICImagingFactory), (void **)&factory));
  wil::com_ptr<IWICBitmapDecoder> decoder;
  THROW_IF_FAILED(factory->CreateDecoderFromFilename(
      found_path, nullptr, GENERIC_READ, WICDecodeMetadataCacheOnLoad,
      &decoder));
  wil::com_ptr<IWICBitmapFrameDecode> frame;
  THROW_IF_FAILED(decoder->GetFrame(0, &frame));
  wil::com_ptr<IWICFormatConverter> converter;
  THROW_IF_FAILED(factory->CreateFormatConverter(&converter));
  THROW_IF_FAILED(converter->Initialize(
      frame.get(), GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeSpiral4x4,
      nullptr, 0, WICBitmapPaletteTypeMedianCut));
  wil::com_ptr<IWICBitmapScaler> scaled;
  THROW_IF_FAILED(factory->CreateBitmapScaler(&scaled));
  THROW_IF_FAILED(scaled->Initialize(
      converter.get(), w, h, WICBitmapInterpolationModeHighQualityCubic));
  void *p_bits;
  wil::unique_hbitmap bitmap{create_dib(w, -h, &p_bits)};
  THROW_IF_FAILED(
      scaled->CopyPixels(nullptr, w * 4, w * h * 4, (BYTE *)p_bits));
  return bitmap.release();
}