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
    pDesktop->GetID(&guid);

    OLECHAR* guidString;
    if (FAILED(StringFromCLSID(guid, &guidString)))
        _ftprintf(stderr, L"FAILED StringFromCLSID %d\n", __LINE__);

    HSTRING s = NULL;
    if (FAILED(pDesktop->GetName(&s)))
        _ftprintf(stderr, L"FAILED GetName %d\n", __LINE__);

    if (s != nullptr)
        _tprintf(L"%s %s\n", guidString, WindowsGetStringRawBuffer(s, nullptr));
    else
        _tprintf(L"%s Desktop %d\n", guidString, dn);

    ::CoTaskMemFree(guidString);
    WindowsDeleteString(s);
}

void ShowCurrentDesktop(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal)))
        _ftprintf(stderr, L"FAILED obtaining CLSID_VirtualDesktopManagerInternal %d\n", __LINE__);

    CComPtr<IVirtualDesktop> pCurrentDesktop;
    if (FAILED(pDesktopManagerInternal->GetCurrentDesktop(&pCurrentDesktop)))
        _ftprintf(stderr, L"FAILED GetCurrentDesktop %d\n", __LINE__);

    CComPtr<IObjectArray> pDesktopArray;
    if (pDesktopManagerInternal && SUCCEEDED(pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
    {
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
}

void ListViews(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps)))
        _ftprintf(stderr, L"FAILED obtaining IID_IVirtualDesktopPinnedApps %d\n", __LINE__);

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    if (FAILED(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection)))
        _ftprintf(stderr, L"FAILED obtaining IID_IApplicationViewCollection %d\n", __LINE__);

    CComPtr<IObjectArray> pViewArray;
    if (pApplicationViewCollection && SUCCEEDED(pApplicationViewCollection->GetViewsByZOrder(&pViewArray)))
    {
        for (CComPtr<IApplicationView> pView : ObjectArrayRange<IApplicationView>(pViewArray))
        {
            BOOL bShowInSwitchers = FALSE;
            if (FAILED(pView->GetShowInSwitchers(&bShowInSwitchers)))
                _ftprintf(stderr, L"FAILED GetShowInSwitchers %d\n", __LINE__);

            if (!bShowInSwitchers)
                continue;

            HWND hWnd;
            if (FAILED(pView->GetThumbnailWindow(&hWnd)))
                _ftprintf(stderr, L"FAILED GetThumbnailWindow %d\n", __LINE__);

            GUID guid;
            if (FAILED(pView->GetVirtualDesktopId(&guid)))
                _ftprintf(stderr, L"FAILED GetVirtualDesktopId %d\n", __LINE__);

            OLECHAR* guidString;
            if (FAILED(StringFromCLSID(guid, &guidString)))
                _ftprintf(stderr, L"FAILED StringFromCLSID %d\n", __LINE__);

            BOOL bPinned = FALSE;
            if (FAILED(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned)))
                _ftprintf(stderr, L"FAILED IsViewPinned %d\n", __LINE__);

            TCHAR title[256];
            GetWindowText(hWnd, title, ARRAYSIZE(title));

            _tprintf(_T("0x%08") _T(PRIXPTR) _T(" %c %s %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), guidString, title);

            ::CoTaskMemFree(guidString);
        }
    }
}

void ListDesktops(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal)))
        _ftprintf(stderr, L"FAILED obtaining CLSID_VirtualDesktopManagerInternal %d\n", __LINE__);

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    if (FAILED(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection)))
        _ftprintf(stderr, L"FAILED obtaining IID_IApplicationViewCollection %d\n", __LINE__);

    CComPtr<IObjectArray> pViewArray;
    if (pApplicationViewCollection && FAILED(pApplicationViewCollection->GetViewsByZOrder(&pViewArray)))
        _ftprintf(stderr, L"FAILED GetViewsByZOrder %d\n", __LINE__);

    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps)))
        _ftprintf(stderr, L"FAILED obtaining IID_IVirtualDesktopPinnedApps %d\n", __LINE__);

    CComPtr<IObjectArray> pDesktopArray;
    if (pDesktopManagerInternal && SUCCEEDED(pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
    {
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
                    if (FAILED(pView->GetShowInSwitchers(&bShowInSwitchers)))
                        _ftprintf(stderr, L"FAILED GetShowInSwitchers %d\n", __LINE__);

                    if (!bShowInSwitchers)
                        continue;

                    BOOL bVisible = FALSE;
                    if (FAILED(pDesktop->IsViewVisible(pView, &bVisible)))
                        _ftprintf(stderr, L"FAILED IsViewVisible %d\n", __LINE__);

                    if (!bVisible)
                        continue;

                    HWND hWnd;
                    if (FAILED(pView->GetThumbnailWindow(&hWnd)))
                        _ftprintf(stderr, L"FAILED GetThumbnailWindow %d\n", __LINE__);

                    BOOL bPinned = FALSE;
                    if (FAILED(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned)))
                        _ftprintf(stderr, L"FAILED IsViewPinned %d\n", __LINE__);

                    TCHAR title[256];
                    GetWindowText(hWnd, title, ARRAYSIZE(title));

                    _tprintf(_T("  0x%08") _T(PRIXPTR) _T(" %c %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), title);
                }
            }
        }
    }
}

void PinView(IServiceProvider* pServiceProvider, HWND hWnd, BOOL bPin)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps)))
        _ftprintf(stderr, L"FAILED obtaining IID_IVirtualDesktopPinnedApps %d\n", __LINE__);

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    if (FAILED(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection)))
        _ftprintf(stderr, L"FAILED obtaining IID_IApplicationViewCollection %d\n", __LINE__);

    CComPtr<IApplicationView> pView;
    if (FAILED(pApplicationViewCollection->GetViewForHwnd(hWnd, &pView)))
        _ftprintf(stderr, L"FAILED obtaining GetViewForHwnd %d\n", __LINE__);

    if (bPin)
    {
        if (FAILED(pVirtualDesktopPinnedApps->PinView(pView)))
            _ftprintf(stderr, L"FAILED PinView %d\n", __LINE__);
    }
    else
    {
        if (FAILED(pVirtualDesktopPinnedApps->UnpinView(pView)))
            _ftprintf(stderr, L"FAILED UnpinView %d\n", __LINE__);
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
    if (FAILED(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop)))
        _ftprintf(stderr, L"FAILED GetCurrentDesktop %d\n", __LINE__);

    CComPtr<IVirtualDesktop> pAdjacentDesktop;
    if (FAILED(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, uDirection, &pAdjacentDesktop)))
        _ftprintf(stderr, L"FAILED GetAdjacentDesktop %d\n", __LINE__);

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
        if (FAILED(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop)))
            _ftprintf(stderr, L"FAILED GetCurrentDesktop %d\n", __LINE__);
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
        if (FAILED(pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
            _ftprintf(stderr, L"FAILED GetDesktops %d\n", __LINE__);

        if (FAILED(pDesktopArray->GetAt(i, IID_PPV_ARGS(&pDesktop))))
            _ftprintf(stderr, L"FAILED GetAt %d\n", __LINE__);

        return pDesktop;
    }
    else
    {
        CComPtr<IVirtualDesktop> pRetDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        if (FAILED(pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
            _ftprintf(stderr, L"FAILED GetDesktops %d\n", __LINE__);

        if (pDesktopManagerInternal && SUCCEEDED(pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
        {
            for (CComPtr<IVirtualDesktop2> pDesktop : ObjectArrayRange<IVirtualDesktop2>(pDesktopArray))
            {
                HSTRING s = NULL;
                if (FAILED(pDesktop->GetName(&s)))
                    _ftprintf(stderr, L"FAILED GetName %d\n", __LINE__);

                if (_tcsicmp(WindowsGetStringRawBuffer(s, nullptr), sDesktop) == 0)
                {
                    pRetDesktop = pDesktop;
                    break;
                }
            }
        }

        return pRetDesktop;
    }
}

void SwitchDesktop(IVirtualDesktopManagerInternal* pDesktopManagerInternal, CComPtr<IVirtualDesktop> pDesktop)
{
    if (pDesktop)
    {
        if (FAILED(pDesktopManagerInternal->SwitchDesktop(pDesktop)))
            _ftprintf(stderr, L"FAILED SwitchDesktop %d\n", __LINE__);
    }
    else
    {
        _ftprintf(stderr, L"Unknown desktop\n");
    }
}

int _tmain(int argc, const TCHAR* const argv[])
{
    if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)))
        _ftprintf(stderr, L"FAILED CoInitialize %d\n", __LINE__);

    {
        CComPtr<IServiceProvider> pServiceProvider;
        if (FAILED(pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER)))
            _ftprintf(stderr, L"FAILED creating CLSID_ImmersiveShell %d\n", __LINE__);

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
            if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal)))
                _ftprintf(stderr, L"FAILED obtaining CLSID_VirtualDesktopManagerInternal %d\n", __LINE__);

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
