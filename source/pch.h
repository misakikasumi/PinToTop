#include <sdkddkver.h>

#include <windows.h>

#include <appmodel.h>
#include <appxpackaging.h>
#include <comdef.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <string.h>
#include <strsafe.h>
#include <wincodec.h>

#ifdef GetCurrentTime
#undef GetCurrentTime
#endif
#include <winrt/base.h>

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.UI.Xaml.Hosting.h>
#include <Windows.UI.Xaml.Hosting.DesktopWindowXamlSource.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Controls.Primitives.h>

#include <algorithm>
#include <filesystem>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "wil/com.h"
#include "wil/resource.h"
#include "wil/result.h"
