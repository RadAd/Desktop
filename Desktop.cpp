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
#include <cctype>
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
    CHECK(pDesktop->GetID(&guid), _T("GetID"), void());

    OLECHAR* guidString;
    CHECK(StringFromCLSID(guid, &guidString), _T("StringFromCLSID"), void());

    HSTRING s = NULL;
    CHECK(pDesktop->GetName(&s), _T("GetName"), void());

    if (s != nullptr)
        _tprintf(_T("%ls %ls\n"), guidString, WindowsGetStringRawBuffer(s, nullptr));
    else
        _tprintf(_T("%ls Desktop %d\n"), guidString, dn);

    ::CoTaskMemFree(guidString);
    CHECK(WindowsDeleteString(s), _T("WindowsDeleteString"), void());
}

void ShowCurrentDesktop(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), void());

    CComPtr<IVirtualDesktop> pCurrentDesktop;
    CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pCurrentDesktop), _T("GetCurrentDesktop"), void());

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), _T("GetDesktops"), void());

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
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), void());

    CComPtr<IObjectArray> pViewArray;
    CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), _T("GetViewsByZOrder"), void());

    for (CComPtr<IApplicationView> pView : ObjectArrayRange<IApplicationView>(pViewArray))
    {
        BOOL bShowInSwitchers = FALSE;
        CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), _T("GetShowInSwitchers"), void());

        if (!bShowInSwitchers)
            continue;

        HWND hWnd;
        CHECK(pView->GetThumbnailWindow(&hWnd), _T("GetThumbnailWindow"), void());

        GUID guid;
        CHECK(pView->GetVirtualDesktopId(&guid), _T("GetVirtualDesktopId"), void());

        OLECHAR* guidString;
        CHECK(StringFromCLSID(guid, &guidString), _T("StringFromCLSID"), void());

        BOOL bPinned = FALSE;
        CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), _T("IsViewPinned"), void());

        TCHAR title[256];
        GetWindowText(hWnd, title, ARRAYSIZE(title));

        _tprintf(_T("0x%08") _T(PRIXPTR) _T(" %c %ls %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), guidString, title);

        ::CoTaskMemFree(guidString);
    }
}

void ListDesktops(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopManagerInternal> pDesktopManagerInternal;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), void());

    CComPtr<IObjectArray> pViewArray;
    CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), _T("GetViewsByZOrder"), void());

    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), void());

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), _T("GetDesktops"), void());

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
                CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), _T("GetShowInSwitchers"), void());

                if (!bShowInSwitchers)
                    continue;

                BOOL bVisible = FALSE;
                CHECK(pDesktop->IsViewVisible(pView, &bVisible), _T("IsViewVisible"), void());

                if (!bVisible)
                    continue;

                HWND hWnd;
                CHECK(pView->GetThumbnailWindow(&hWnd), _T("GetThumbnailWindow"), void());

                BOOL bPinned = FALSE;
                CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), _T("IsViewPinned"), void());

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
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), void());

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), void());

    CComPtr<IApplicationView> pView;
    CHECK(pApplicationViewCollection->GetViewForHwnd(hWnd, &pView), _T("obtaining GetViewForHwnd"), void());

    if (bPin)
    {
        CHECK(pVirtualDesktopPinnedApps->PinView(pView), _T("PinView"), void());
    }
    else
    {
        CHECK(pVirtualDesktopPinnedApps->UnpinView(pView), _T("UnpinView"), void());
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
    CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop), _T("GetCurrentDesktop"), {});

    CComPtr<IVirtualDesktop> pAdjacentDesktop;
    CHECK(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, uDirection, &pAdjacentDesktop), _T("GetAdjacentDesktop"), {});

    return pAdjacentDesktop;
}

CComPtr<IVirtualDesktop> GetDesktop(IVirtualDesktopManagerInternal* pDesktopManagerInternal, const LPCTSTR sDesktop)
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
        CHECK(pDesktopManagerInternal->GetCurrentDesktop(&pDesktop), _T("GetCurrentDesktop"), {});
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
        CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), _T("GetDesktops"), {});

        CHECK(pDesktopArray->GetAt(i, IID_PPV_ARGS(&pDesktop)), _T("GetAt"), {});

        return pDesktop;
    }
    else
    {
#ifdef UNICODE
        LPCWSTR wDesktop = sDesktop;
#else
        WCHAR wDesktop[100];
        MultiByteToWideChar(CP_UTF8, 0, sDesktop, -1, wDesktop, ARRAYSIZE(wDesktop));
#endif
        CComPtr<IVirtualDesktop> pRetDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(pDesktopManagerInternal->GetDesktops(&pDesktopArray), _T("GetDesktops"), {});

        for (CComPtr<IVirtualDesktop2> pDesktop : ObjectArrayRange<IVirtualDesktop2>(pDesktopArray))
        {
            HSTRING s = NULL;
            CHECK(pDesktop->GetName(&s), _T("GetName"), {});

            if (_wcsicmp(WindowsGetStringRawBuffer(s, nullptr), wDesktop) == 0)
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
        CHECK(pDesktopManagerInternal->SwitchDesktop(pDesktop), _T("SwitchDesktop"), void());
    }
    else
    {
        _ftprintf(stderr, _T("Unknown desktop\n"));
    }
}

int _tmain(int argc, const TCHAR* const argv[])
{
    CHECK(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), _T("CoInitialize"), EXIT_FAILURE);

    {
        CComPtr<IServiceProvider> pServiceProvider;
        CHECK(pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER), _T("creating CLSID_ImmersiveShell"), EXIT_FAILURE);

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
            CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);

            SwitchDesktop(pDesktopManagerInternal, GetDesktop(pDesktopManagerInternal, desktop));
        }
        else
        {
            _tprintf(_T("Desktop [cmd] <options>\n\n"));
            _tprintf(_T("where [cmd] is one of:\n"));
            _tprintf(_T("  current\n"));
            _tprintf(_T("  list\n"));
            _tprintf(_T("  views\n"));
            _tprintf(_T("  pin <HWND>\n"));
            _tprintf(_T("  unpin <HWND>\n"));
            _tprintf(_T("  switch <desktop>\n"));
            //_tprintf(_T("  create <name>\n"));
            //_tprintf(_T("  remove <desktop>\n"));
            //_tprintf(_T("  rename <desktop> <name>\n"));
        }
    }

    CoUninitialize();

    return EXIT_SUCCESS;
}
