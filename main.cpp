#include <exdisp.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <windows.h>

#include <algorithm>
#include <set>
#include <string>
#include <vector>

#ifndef WH_MOD
#ifndef Wh_log
#define Wh_log printf
#endif
#endif

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shell32.lib")

#define WM_COMMAND 0x0111
#define CMD_NEW_TAB 0xA21B
#define CMD_SELECT_TAB_BASE 0xA221

std::wstring GetComparisonPath(std::wstring path) {
    if (path.empty()) return L"";
    if (path.back() == L'\\' || path.back() == L'/') path.pop_back();
    if (path.find(L"file:///") == 0) path = path.substr(8);
    std::replace(path.begin(), path.end(), L'/', L'\\');
    std::transform(path.begin(), path.end(), path.begin(), ::towlower);
    return path;
}

std::wstring GetRawBrowserUrl(IWebBrowser2* pWebBrowser) {
    if (!pWebBrowser) return L"";
    BSTR bstrUrl = NULL;
    if (SUCCEEDED(pWebBrowser->get_LocationURL(&bstrUrl)) && bstrUrl) {
        std::wstring url = bstrUrl;
        SysFreeString(bstrUrl);
        return url;
    }
    return L"";
}

HWND GetTabHandle(IWebBrowser2* pWebBrowser) {
    if (!pWebBrowser) return NULL;
    IServiceProvider* pServiceProvider = NULL;
    if (SUCCEEDED(pWebBrowser->QueryInterface(IID_IServiceProvider, (void**)&pServiceProvider))) {
        IShellBrowser* pShellBrowser = NULL;
        if (SUCCEEDED(pServiceProvider->QueryService(SID_SShellBrowser, IID_IShellBrowser, (void**)&pShellBrowser))) {
            HWND hwndTab = NULL;
            if (SUCCEEDED(pShellBrowser->GetWindow(&hwndTab))) {
                pShellBrowser->Release();
                pServiceProvider->Release();
                return hwndTab;
            }
            pShellBrowser->Release();
        }
        pServiceProvider->Release();
    }
    return NULL;
}

std::set<HWND> GetWindowTabHandles(HWND mainWindow) {
    std::set<HWND> tabs;
    HWND hTab = FindWindowExW(mainWindow, NULL, L"ShellTabWindowClass", NULL);
    while (hTab) {
        tabs.insert(hTab);
        hTab = FindWindowExW(mainWindow, hTab, L"ShellTabWindowClass", NULL);
    }
    return tabs;
}

// Advanced Navigation using IShellBrowser::BrowseObject
HRESULT NavigateToPath(IWebBrowser2* pWB, const std::wstring& path) {
    HRESULT hrFinal = E_FAIL;
    IServiceProvider* pServiceProvider = NULL;

    // 1. Try IShellBrowser::BrowseObject (Lower level, more reliable)
    if (SUCCEEDED(pWB->QueryInterface(IID_IServiceProvider, (void**)&pServiceProvider))) {
        IShellBrowser* pShellBrowser = NULL;
        if (SUCCEEDED(pServiceProvider->QueryService(SID_SShellBrowser, IID_IShellBrowser, (void**)&pShellBrowser))) {
            LPITEMIDLIST pidl = NULL;
            if (SUCCEEDED(SHParseDisplayName(path.c_str(), NULL, &pidl, 0, NULL))) {
                hrFinal = pShellBrowser->BrowseObject(pidl, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
                ILFree(pidl);
            }
            pShellBrowser->Release();
        }
        pServiceProvider->Release();
    }

    // 2. Fallback to Navigate2 if BrowseObject didn't work or wasn't available
    if (FAILED(hrFinal)) {
        Wh_log("[Log] BrowseObject failed, falling back to Navigate2...\n");
        VARIANT vUrl;
        vUrl.vt = VT_BSTR;
        vUrl.bstrVal = SysAllocString(path.c_str());
        hrFinal = pWB->Navigate2(&vUrl, NULL, NULL, NULL, NULL);
        VariantClear(&vUrl);
    }

    return hrFinal;
}

void CALLBACK WinEventProc(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd, LONG idObject, LONG idChild,
                           DWORD dwEventThread, DWORD dwmsEventTime) {
    if (event == EVENT_OBJECT_SHOW && idObject == OBJID_WINDOW && idChild == CHILDID_SELF) {
        wchar_t className[256];
        if (GetClassNameW(hwnd, className, 256) && wcscmp(className, L"CabinetWClass") == 0) {
            // CRITICAL: Hide window IMMEDIATELY to prevent flicker
            // Move it off-screen AND hide it
            SetWindowPos(hwnd, NULL, -32000, -32000, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);

            IShellWindows* pShellWindows = NULL;
            if (FAILED(
                    CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&pShellWindows)))
                return;

            IWebBrowser2* pNewWebBrowser = NULL;
            // Wait for COM registration without blocking the whole system, but very short
            for (int r = 0; r < 10; ++r) {
                long count = 0;
                pShellWindows->get_Count(&count);
                for (long i = 0; i < count; ++i) {
                    VARIANT vIdx;
                    vIdx.vt = VT_I4;
                    vIdx.lVal = i;
                    IDispatch* pDisp = NULL;
                    if (SUCCEEDED(pShellWindows->Item(vIdx, &pDisp)) && pDisp) {
                        IWebBrowser2* pWB = NULL;
                        if (SUCCEEDED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB))) {
                            SHANDLE_PTR hPtr = 0;
                            if (SUCCEEDED(pWB->get_HWND(&hPtr)) && (HWND)hPtr == hwnd) {
                                pNewWebBrowser = pWB;
                                pDisp->Release();
                                break;
                            }
                            pWB->Release();
                        }
                        pDisp->Release();
                    }
                }
                if (pNewWebBrowser) break;
                Sleep(50);
            }

            if (!pNewWebBrowser) {
                pShellWindows->Release();
                return;
            }

            std::wstring rawUrl = GetRawBrowserUrl(pNewWebBrowser);
            std::wstring normalizedUrl = GetComparisonPath(rawUrl);
            if (normalizedUrl.empty()) {
                pNewWebBrowser->Release();
                pShellWindows->Release();
                return;
            }

            // Check if we should merge it
            HWND mainHwnd = NULL;
            HWND tempHwnd = FindWindowExW(NULL, NULL, L"CabinetWClass", NULL);
            while (tempHwnd) {
                if (tempHwnd != hwnd && IsWindowVisible(tempHwnd)) {
                    mainHwnd = tempHwnd;
                    break;
                }
                tempHwnd = FindWindowExW(NULL, tempHwnd, L"CabinetWClass", NULL);
            }

            if (mainHwnd) {
                // Check if path already exists in any window
                HWND existingTabHwnd = NULL;
                HWND existingMainHwnd = NULL;
                long count = 0;
                pShellWindows->get_Count(&count);
                for (long i = 0; i < count; ++i) {
                    VARIANT vIdx;
                    vIdx.vt = VT_I4;
                    vIdx.lVal = i;
                    IDispatch* pDisp = NULL;
                    if (SUCCEEDED(pShellWindows->Item(vIdx, &pDisp)) && pDisp) {
                        IWebBrowser2* pWB = NULL;
                        if (SUCCEEDED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB))) {
                            SHANDLE_PTR hPtr = 0;
                            if (SUCCEEDED(pWB->get_HWND(&hPtr)) && (HWND)hPtr != hwnd) {
                                if (GetComparisonPath(GetRawBrowserUrl(pWB)) == normalizedUrl) {
                                    existingMainHwnd = (HWND)hPtr;
                                    existingTabHwnd = GetTabHandle(pWB);
                                    pWB->Release();
                                    pDisp->Release();
                                    break;
                                }
                            }
                            pWB->Release();
                        }
                        pDisp->Release();
                    }
                }

                if (existingMainHwnd && existingTabHwnd) {
                    // Just switch
                    std::vector<HWND> tabList;
                    HWND hTab = FindWindowExW(existingMainHwnd, NULL, L"ShellTabWindowClass", NULL);
                    while (hTab) {
                        tabList.push_back(hTab);
                        hTab = FindWindowExW(existingMainHwnd, hTab, L"ShellTabWindowClass", NULL);
                    }
                    auto it = std::find(tabList.begin(), tabList.end(), existingTabHwnd);
                    if (it != tabList.end()) {
                        int idx = (int)std::distance(tabList.begin(), it);
                        SendMessageW(existingMainHwnd, WM_COMMAND, CMD_SELECT_TAB_BASE + idx + 1, 0);
                    }
                    SetForegroundWindow(existingMainHwnd);
                    pNewWebBrowser->Quit();
                } else {
                    // Convert to tab
                    std::set<HWND> oldTabs = GetWindowTabHandles(mainHwnd);

                    PostMessageW(mainHwnd, WM_COMMAND, CMD_NEW_TAB, 0);
                    pNewWebBrowser->Quit();  // Standalone window is hidden anyway, so Quitting is seamless

                    HWND newTabHwnd = NULL;
                    for (int r = 0; r < 30; ++r) {
                        Sleep(50);
                        std::set<HWND> currentTabs = GetWindowTabHandles(mainHwnd);
                        for (HWND h : currentTabs) {
                            if (oldTabs.find(h) == oldTabs.end()) {
                                newTabHwnd = h;
                                break;
                            }
                        }
                        if (newTabHwnd) break;
                    }

                    if (newTabHwnd) {
                        bool navigated = false;
                        for (int retry = 0; retry < 10; ++retry) {
                            pShellWindows->get_Count(&count);
                            for (long i = 0; i < count; ++i) {
                                VARIANT vIdx;
                                vIdx.vt = VT_I4;
                                vIdx.lVal = i;
                                IDispatch* pDisp = NULL;
                                if (SUCCEEDED(pShellWindows->Item(vIdx, &pDisp)) && pDisp) {
                                    IWebBrowser2* pWB = NULL;
                                    if (SUCCEEDED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB))) {
                                        if (GetTabHandle(pWB) == newTabHwnd) {
                                            NavigateToPath(pWB, normalizedUrl);
                                            navigated = true;
                                            pWB->Release();
                                            pDisp->Release();
                                            break;
                                        }
                                        pWB->Release();
                                    }
                                    pDisp->Release();
                                }
                            }
                            if (navigated) break;
                            Sleep(100);
                        }
                    }
                    SetForegroundWindow(mainHwnd);
                }
            } else {
                // No existing window, so we must show THIS window
                SetWindowPos(hwnd, NULL, 100, 100, 1000, 700, SWP_SHOWWINDOW);
            }
            pNewWebBrowser->Release();
            pShellWindows->Release();
        }
    }
}

int main() {
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    Wh_log("[Log] Explorer Merger start...\n");
    HWINEVENTHOOK hHook =
        SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW, NULL, WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    UnhookWinEvent(hHook);
    CoUninitialize();
    return 0;
}
