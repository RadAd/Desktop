#include <windows.h>
#include <atlbase.h>
#include <shlobj.h>
#include <winstring.h>
#include <assert.h>

#include <tchar.h>
#include "Win10Desktops.h"
#include "ComUtils.h"
#include <combaseapi.h>
#include <inttypes.h>
#include <cwctype>

#include <cstdlib>
#include <cstdio>

// "1841C6D7-4F9D-42C0-AF41-8747538F10E5"
const IID IID_IApplicationViewCollection = {
    0x1841C6D7, 0x4F9D, 0x42C0, { 0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5 }
};

void PrintDesktop(IVirtualDesktop2* pDesktop, int dn)
{
    GUID guid;
    CHECK(pDesktop->GetID(&guid), L"GetID", void());

    OLECHAR* guidString;
    CHECK(StringFromCLSID(guid, &guidString), L"StringFromCLSID", void());

    HSTRING s = NULL;
    CHECK(pDesktop->GetName(&s), L"GetName", void());

    if (s != nullptr)
        _tprintf(L"%s %s\n", guidString, WindowsGetStringRawBuffer(s, nullptr));
    else
        _tprintf(L"%s Desktop %d\n", guidString, dn);

    ::CoTaskMemFree(guidString);
    CHECK(WindowsDeleteString(s), L"WindowsDeleteString", void());
}

void ShowCurrentDesktop(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), L"obtaining CLSID_VirtualDesktopManagerInternal", void());

    CComPtr<IVirtualDesktop> pCurrentDesktop;
    CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pCurrentDesktop), L"GetCurrentDesktop", void());

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), L"GetDesktops", void());

    int dn = 0;
    for (CComPtr<IVirtualDesktop2> pDesktop : ObjectArrayRange<IVirtualDesktop2>(pDesktopArray))
    {
        if (pDesktop.IsEqualObject(pCurrentDesktop))
        {
            ++dn;
            PrintDesktop(pDesktop, dn);
        }
    }
}

void ListViews(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), L"obtaining IID_IVirtualDesktopPinnedApps", void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), L"obtaining IID_IApplicationViewCollection", void());

    CComPtr<IObjectArray> pViewArray;
    CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), L"GetViewsByZOrder", void());

    for (CComPtr<IApplicationView> pView : ObjectArrayRange<IApplicationView>(pViewArray))
    {
        BOOL bShowInSwitchers = FALSE;
        CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), L"GetShowInSwitchers", void());

        if (!bShowInSwitchers)
            continue;

        HWND hWnd;
        CHECK(pView->GetThumbnailWindow(&hWnd), L"GetThumbnailWindow", void());

        GUID guid;
        CHECK(pView->GetVirtualDesktopId(&guid), L"GetVirtualDesktopId", void());

        OLECHAR* guidString;
        CHECK(StringFromCLSID(guid, &guidString), L"StringFromCLSID", void());

        BOOL bPinned = FALSE;
        CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), L"IsViewPinned", void());

        TCHAR title[256];
        GetWindowText(hWnd, title, ARRAYSIZE(title));

        _tprintf(_T("0x%08") _T(PRIXPTR) _T(" %c %s %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), guidString, title);

        ::CoTaskMemFree(guidString);
    }
}

void ListDesktops(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), L"obtaining CLSID_VirtualDesktopManagerInternal", void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), L"obtaining IID_IApplicationViewCollection", void());

    CComPtr<IObjectArray> pViewArray;
    CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), L"GetViewsByZOrder", void());

    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), L"obtaining IID_IVirtualDesktopPinnedApps", void());

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), L"GetDesktops", void());

    int dn = 0;
    for (CComPtr<IVirtualDesktop2> pDesktop : ObjectArrayRange<IVirtualDesktop2>(pDesktopArray))
    {
        ++dn;

        PrintDesktop(pDesktop, dn);

        if (pViewArray)
        {
            for (CComPtr<IApplicationView> pView : ObjectArrayRangeRev<IApplicationView>(pViewArray))
            {
                BOOL bShowInSwitchers = FALSE;
                CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), L"GetShowInSwitchers", void());

                if (!bShowInSwitchers)
                    continue;

                BOOL bVisible = FALSE;
                CHECK(pDesktop->IsViewVisible(pView, &bVisible), L"IsViewVisible", void());

                if (!bVisible)
                    continue;

                HWND hWnd;
                CHECK(pView->GetThumbnailWindow(&hWnd), L"GetThumbnailWindow", void());

                BOOL bPinned = FALSE;
                CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), L"IsViewPinned", void());

                TCHAR title[256];
                GetWindowText(hWnd, title, ARRAYSIZE(title));

                _tprintf(_T("  0x%08") _T(PRIXPTR) _T(" %c %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), title);
            }
        }
    }
}

void PinView(IServiceProvider* pServiceProvider, HWND hWnd, BOOL bPin)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), L"obtaining IID_IVirtualDesktopPinnedApps", void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), L"obtaining IID_IApplicationViewCollection", void());

    CComPtr<IApplicationView> pView;
    CHECK(pApplicationViewCollection->GetViewForHwnd(hWnd, &pView), L"obtaining GetViewForHwnd", void());

    if (bPin)
    {
        CHECK(pVirtualDesktopPinnedApps->PinView(pView), L"PinView", void());
    }
    else
    {
        CHECK(pVirtualDesktopPinnedApps->UnpinView(pView), L"UnpinView", void());
    }
}

HWND GetWindow(const TCHAR* s)
{
    if (s == nullptr)
        return NULL;
    else if (_tcsicmp(s, _T("{cursor}")) == 0)
    {
        POINT pt;
        GetCursorPos(&pt);
        return WindowFromPoint(pt);
    }
    else if (_tcsicmp(s, _T("{console}")) == 0)
        return GetConsoleWindow();
    else if (_tcsicmp(s, _T("{foreground}")) == 0)
        return GetForegroundWindow();
    else
        return static_cast<HWND>(UlongToHandle(_tcstoul(s, nullptr, 0)));
}

CComPtr<IVirtualDesktop> GetAdjacentDesktop(IVirtualDesktopManagerInternal* pDesktopManagerInternal, AdjacentDesktop uDirection)
{
    CComPtr<IVirtualDesktop> pDesktop;
    CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop), L"GetCurrentDesktop", {});

    CComPtr<IVirtualDesktop> pAdjacentDesktop;
    CHECK(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, uDirection, &pAdjacentDesktop), L"GetAdjacentDesktop", {});

    return pAdjacentDesktop;
}

CComPtr<IVirtualDesktop> GetDesktop(IVirtualDesktopManagerInternal* pDesktopManagerInternal, const TCHAR* sDesktop)
{
    // TODO Support
    // by name
    // by guid
    // by position
    if (sDesktop == nullptr)
        return {};
    else if (_tcsicmp(sDesktop, _T("{current}")) == 0)
    {
        CComPtr<IVirtualDesktop> pDesktop;
        CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop), L"GetCurrentDesktop", {});
        return pDesktop;
    }
    else if (_tcsicmp(sDesktop, _T("{prev}")) == 0)
        return GetAdjacentDesktop(pDesktopManagerInternal, LeftDirection);
    else if (_tcsicmp(sDesktop, _T("{next}")) == 0)
        return GetAdjacentDesktop(pDesktopManagerInternal, RightDirection);
    else if (std::_istdigit(sDesktop[0]))
    {
        int i = _tstoi(sDesktop);
        CComPtr<IVirtualDesktop> pDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), L"GetDesktops", {});

        CHECK(pDesktopArray->GetAt(i, IID_PPV_ARGS(&pDesktop)), L"GetAt", {});

        return pDesktop;
    }
    else
    {
        CComPtr<IVirtualDesktop> pRetDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), L"GetDesktops", {});

        for (CComPtr<IVirtualDesktop2> pDesktop : ObjectArrayRange<IVirtualDesktop2>(pDesktopArray))
        {
            HSTRING s = NULL;
            CHECK(pDesktop->GetName(&s), L"GetName", {});

            if (_tcsicmp(WindowsGetStringRawBuffer(s, nullptr), sDesktop) == 0)
            {
                pRetDesktop = pDesktop;
                break;
            }
        }

        return pRetDesktop;
    }
}

void SwitchDesktop(IVirtualDesktopManagerInternal* pDesktopManagerInternal, CComPtr<IVirtualDesktop> pDesktop)
{
    if (pDesktop)
    {
        CHECK(pDesktopManagerInternal->SwitchDesktop(pDesktop), L"SwitchDesktop", void());
    }
    else
    {
        _ftprintf(stderr, L"Unknown desktop\n");
    }
}

int _tmain(int argc, const TCHAR* const argv[])
{
    CHECK(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), L"CoInitialize", EXIT_FAILURE);

    {
        CComPtr<IServiceProvider> pServiceProvider;
        CHECK(pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER), L"creating CLSID_ImmersiveShell", EXIT_FAILURE);

        const TCHAR* cmd = argc > 0 ? argv[1] : nullptr;
        const TCHAR* wnd = argc > 1 ? argv[2] : nullptr;
        const TCHAR* desktop = argc > 1 ? argv[2] : nullptr;

        if (cmd != nullptr && _tcsicmp(cmd, _T("current")) == 0)
            ShowCurrentDesktop(pServiceProvider);
        else if (cmd != nullptr && _tcsicmp(cmd, _T("views")) == 0)
            ListViews(pServiceProvider);
        else if (cmd != nullptr && _tcsicmp(cmd, _T("list")) == 0)
            ListDesktops(pServiceProvider);
        else if (cmd != nullptr && wnd != nullptr && _tcsicmp(cmd, _T("pin")) == 0)
            PinView(pServiceProvider, GetWindow(wnd), TRUE);
        else if (cmd != nullptr && wnd != nullptr && _tcsicmp(cmd, _T("unpin")) == 0)
            PinView(pServiceProvider, GetWindow(wnd), FALSE);
        else if (cmd != nullptr && desktop != nullptr && _tcsicmp(cmd, _T("switch")) == 0)
        {
            CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
            CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), L"obtaining CLSID_VirtualDesktopManagerInternal", EXIT_FAILURE);

            SwitchDesktop(pDesktopManagerInternal, GetDesktop(pDesktopManagerInternal, desktop));
        }
        else
        {
            _tprintf(L"Desktop [cmd] <options>\n\n");
            _tprintf(L"where [cmd] is one of:\n");
            _tprintf(L"  current\n");
            _tprintf(L"  list\n");
            _tprintf(L"  views\n");
            _tprintf(L"  pin <HWND>\n");
            _tprintf(L"  unpin <HWND>\n");
            _tprintf(L"  switch <desktop>\n");
            //_tprintf(L"  create <name>\n");
            //_tprintf(L"  remove <desktop>\n");
            //_tprintf(L"  rename <desktop> <name>\n");
        }
    }

    CoUninitialize();

    return EXIT_SUCCESS;
}
