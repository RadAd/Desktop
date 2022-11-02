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

template <class T>
typename ATL::CComPtr<T>::_PtrClass* operator-(const ATL::CComPtr<T>& p)
{
    return static_cast<typename ATL::CComPtr<T>::_PtrClass*>(p);
}

template <class VD>
void PrintDesktop(VD* pDesktop, int dn)
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

template <class VDTypes>
int ShowCurrentDesktop(CComPtr<typename VDTypes::IVirtualDesktopManagerInternal> pDesktopManagerInternal)
{
    CComPtr<typename VDTypes::IVirtualDesktop> pCurrentDesktop;
    CHECK(GetCurrentDesktop(-pDesktopManagerInternal, &pCurrentDesktop), _T("GetCurrentDesktop"), EXIT_FAILURE);

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(GetDesktops(-pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), EXIT_FAILURE);

    int dn = 0;
    for (CComPtr<typename VDTypes::IVirtualDesktop2> pDesktop : ObjectArrayRange<typename VDTypes::IVirtualDesktop2>(pDesktopArray))
    {
        if (pDesktop.IsEqualObject(pCurrentDesktop))
        {
            ++dn;
            PrintDesktop(-pDesktop, dn);
        }
    }
    return EXIT_SUCCESS;
}

int ShowCurrentDesktop(IServiceProvider* pServiceProvider)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return ShowCurrentDesktop<Win10VDTypes>(pDesktopManagerInternal10);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return ShowCurrentDesktop<Win11VDTypes>(pDesktopManagerInternal11);
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

int ListViews(IServiceProvider* pServiceProvider)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), EXIT_FAILURE);

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(IID_IApplicationViewCollection, &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), EXIT_FAILURE);

    CComPtr<IObjectArray> pViewArray;
    CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), _T("GetViewsByZOrder"), EXIT_FAILURE);

    for (CComPtr<IApplicationView> pView : ObjectArrayRange<IApplicationView>(pViewArray))
    {
        BOOL bShowInSwitchers = FALSE;
        CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), _T("GetShowInSwitchers"), EXIT_FAILURE);

        if (!bShowInSwitchers)
            continue;

        HWND hWnd;
        CHECK(pView->GetThumbnailWindow(&hWnd), _T("GetThumbnailWindow"), EXIT_FAILURE);

        GUID guid;
        CHECK(pView->GetVirtualDesktopId(&guid), _T("GetVirtualDesktopId"), EXIT_FAILURE);

        OLECHAR* guidString;
        CHECK(StringFromCLSID(guid, &guidString), _T("StringFromCLSID"), EXIT_FAILURE);

        BOOL bPinned = FALSE;
        CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), _T("IsViewPinned"), EXIT_FAILURE);

        TCHAR title[256];
        GetWindowText(hWnd, title, ARRAYSIZE(title));

        _tprintf(_T("0x%08") _T(PRIXPTR) _T(" %c %ls %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), guidString, title);

        ::CoTaskMemFree(guidString);
    }
    return EXIT_SUCCESS;
}

template <class VDTypes>
int ListDesktops(IServiceProvider* pServiceProvider, CComPtr<typename VDTypes::IVirtualDesktopManagerInternal> pDesktopManagerInternal, bool bShowViews)
{
    CComPtr<IObjectArray> pDesktopArray;
    CHECK(GetDesktops(-pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), EXIT_FAILURE);

    if (bShowViews)
    {
        CComPtr<IApplicationViewCollection> pApplicationViewCollection;
        CHECK(pServiceProvider->QueryService(IID_IApplicationViewCollection, &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), EXIT_FAILURE);

        CComPtr<IObjectArray> pViewArray;
        CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), _T("GetViewsByZOrder"), EXIT_FAILURE);

        CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
        CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), EXIT_FAILURE);

        int dn = 0;
        for (CComPtr<typename VDTypes::IVirtualDesktop2> pDesktop : ObjectArrayRange<typename VDTypes::IVirtualDesktop2>(pDesktopArray))
        {
            ++dn;

            PrintDesktop(-pDesktop, dn);

            if (pViewArray)
            {
                for (CComPtr<IApplicationView> pView : ObjectArrayRangeRev<IApplicationView>(pViewArray))
                {
                    BOOL bShowInSwitchers = FALSE;
                    CHECK(pView->GetShowInSwitchers(&bShowInSwitchers), _T("GetShowInSwitchers"), EXIT_FAILURE);

                    if (!bShowInSwitchers)
                        continue;

                    BOOL bVisible = FALSE;
                    CHECK(pDesktop->IsViewVisible(pView, &bVisible), _T("IsViewVisible"), EXIT_FAILURE);

                    if (!bVisible)
                        continue;

                    HWND hWnd;
                    CHECK(pView->GetThumbnailWindow(&hWnd), _T("GetThumbnailWindow"), EXIT_FAILURE);

                    BOOL bPinned = FALSE;
                    CHECK(pVirtualDesktopPinnedApps->IsViewPinned(pView, &bPinned), _T("IsViewPinned"), EXIT_FAILURE);

                    TCHAR title[256];
                    GetWindowText(hWnd, title, ARRAYSIZE(title));

                    _tprintf(_T("  0x%08") _T(PRIXPTR) _T(" %c %s\n"), reinterpret_cast<uintptr_t>(hWnd), bPinned ? _T('*') : _T(' '), title);
                }
            }
        }
    }
    else
    {
        int dn = 0;
        for (CComPtr<typename VDTypes::IVirtualDesktop2> pDesktop : ObjectArrayRange<typename VDTypes::IVirtualDesktop2>(pDesktopArray))
        {
            ++dn;

            PrintDesktop(-pDesktop, dn);
        }
    }

    return EXIT_SUCCESS;
}

int ListDesktops(IServiceProvider* pServiceProvider, bool bShowViews)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return ListDesktops<Win10VDTypes>(pServiceProvider, pDesktopManagerInternal10, bShowViews);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return ListDesktops<Win11VDTypes>(pServiceProvider, pDesktopManagerInternal11, bShowViews);
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

int PinView(IServiceProvider* pServiceProvider, HWND hWnd, BOOL bPin)
{
    CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
    CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), EXIT_FAILURE);

    CComPtr<IApplicationViewCollection> pApplicationViewCollection;
    CHECK(pServiceProvider->QueryService(IID_IApplicationViewCollection, &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), EXIT_FAILURE);

    CComPtr<IApplicationView> pView;
    CHECK(pApplicationViewCollection->GetViewForHwnd(hWnd, &pView), _T("obtaining GetViewForHwnd"), EXIT_FAILURE);

    if (bPin)
    {
        CHECK(pVirtualDesktopPinnedApps->PinView(pView), _T("PinView"), EXIT_FAILURE);
    }
    else
    {
        CHECK(pVirtualDesktopPinnedApps->UnpinView(pView), _T("UnpinView"), EXIT_FAILURE);
    }
    return EXIT_SUCCESS;
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

template <class VDTypes>
CComPtr<typename VDTypes::IVirtualDesktop> GetAdjacentDesktop(typename VDTypes::IVirtualDesktopManagerInternal* pDesktopManagerInternal, AdjacentDesktop uDirection)
{
    CComPtr<typename VDTypes::IVirtualDesktop> pDesktop;
    CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pDesktop), _T("GetCurrentDesktop"), {});

    CComPtr<typename VDTypes::IVirtualDesktop> pAdjacentDesktop;
    CHECK(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, uDirection, &pAdjacentDesktop), _T("GetAdjacentDesktop"), {});

    return pAdjacentDesktop;
}

template <class VDTypes>
CComPtr<typename VDTypes::IVirtualDesktop> GetDesktop(typename VDTypes::IVirtualDesktopManagerInternal* pDesktopManagerInternal, const LPCTSTR sDesktop)
{
    // TODO Support
    // by name
    // by guid
    // by position
    if (sDesktop == nullptr)
        return {};
    else if (_tcsicmp(sDesktop, _T("{current}")) == 0)
    {
        CComPtr<typename VDTypes::IVirtualDesktop> pDesktop;
        CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pDesktop), _T("GetCurrentDesktop"), {});
        return pDesktop;
    }
    else if (_tcsicmp(sDesktop, _T("{prev}")) == 0)
        return GetAdjacentDesktop<VDTypes>(pDesktopManagerInternal, LeftDirection);
    else if (_tcsicmp(sDesktop, _T("{next}")) == 0)
        return GetAdjacentDesktop<VDTypes>(pDesktopManagerInternal, RightDirection);
    else if (std::_istdigit(sDesktop[0]))
    {
        int i = _tstoi(sDesktop);
        CComPtr<typename VDTypes::IVirtualDesktop> pDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), {});

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
        CComPtr<typename VDTypes::IVirtualDesktop> pRetDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), {});

        for (CComPtr<typename VDTypes::IVirtualDesktop2> pDesktop : ObjectArrayRange<typename VDTypes::IVirtualDesktop2>(pDesktopArray))
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

template <class VDTypes>
int SwitchDesktop(typename VDTypes::IVirtualDesktopManagerInternal* pDesktopManagerInternal, CComPtr<typename VDTypes::IVirtualDesktop> pDesktop)
{
    if (pDesktop)
    {
        CHECK(SwitchDesktop(pDesktopManagerInternal, -pDesktop), _T("SwitchDesktop"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
    else
    {
        _ftprintf(stderr, _T("Unknown desktop\n"));
        return EXIT_FAILURE;
    }
}

int SwitchDesktop(IServiceProvider* pServiceProvider, LPCTSTR strDesktop)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    //CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return SwitchDesktop<Win10VDTypes>(pDesktopManagerInternal10, GetDesktop<Win10VDTypes>(pDesktopManagerInternal10, strDesktop));
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return SwitchDesktop<Win11VDTypes>(pDesktopManagerInternal11, GetDesktop<Win11VDTypes>(pDesktopManagerInternal11, strDesktop));
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

template <class VDTypes>
int RenameDesktop(typename VDTypes::IVirtualDesktopManagerInternal2* pDesktopManagerInternal, CComPtr<typename VDTypes::IVirtualDesktop> pDesktop, LPCTSTR strName)
{
    if (pDesktop)
    {
        HSTRING hName = NULL;
        CHECK(WindowsCreateString(strName, UINT32(_tcslen(strName)), &hName), _T("WindowsCreateString"), EXIT_FAILURE);
        CHECK(pDesktopManagerInternal->SetName(pDesktop, hName), _T("SwitchDesktop"), EXIT_FAILURE);
        CHECK(WindowsDeleteString(hName), _T("WindowsDeleteString"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
    else
    {
        _ftprintf(stderr, _T("Unknown desktop\n"));
        return EXIT_FAILURE;
    }
}

int RenameDesktop(IServiceProvider* pServiceProvider, LPCTSTR strDesktop, LPCTSTR strName)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal2> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    //CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return RenameDesktop<Win10VDTypes>(pDesktopManagerInternal10, GetDesktop<Win10VDTypes>(pDesktopManagerInternal10, strDesktop), strName);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return RenameDesktop<Win11VDTypes>(pDesktopManagerInternal11, GetDesktop<Win11VDTypes>(pDesktopManagerInternal11, strDesktop), strName);
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
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
        const TCHAR* showviews = argc > 1 ? argv[2] : nullptr;
        const TCHAR* name = argc > 1 ? argv[3] : nullptr;

        if (cmd != nullptr && _tcsicmp(cmd, _T("current")) == 0)
            return ShowCurrentDesktop(pServiceProvider);
        else if (cmd != nullptr && _tcsicmp(cmd, _T("views")) == 0)
            return ListViews(pServiceProvider);
        else if (cmd != nullptr && _tcsicmp(cmd, _T("list")) == 0)
            return ListDesktops(pServiceProvider, showviews != nullptr && _tcsicmp(showviews, _T("/views")) == 0);
        else if (cmd != nullptr && wnd != nullptr && _tcsicmp(cmd, _T("pin")) == 0)
            return PinView(pServiceProvider, GetWindow(wnd), TRUE);
        else if (cmd != nullptr && wnd != nullptr && _tcsicmp(cmd, _T("unpin")) == 0)
            return PinView(pServiceProvider, GetWindow(wnd), FALSE);
        else if (cmd != nullptr && desktop != nullptr && _tcsicmp(cmd, _T("switch")) == 0)
            return SwitchDesktop(pServiceProvider, desktop);
        else if (cmd != nullptr && desktop != nullptr && _tcsicmp(cmd, _T("rename")) == 0 && name != nullptr)
            return RenameDesktop(pServiceProvider, desktop, name);
        else
        {
            _tprintf(_T("Desktop [cmd] <options>\n\n"));
            _tprintf(_T("where [cmd] is one of:\n"));
            _tprintf(_T("  current\n"));
            _tprintf(_T("  list [/views]\n"));
            _tprintf(_T("  views\n"));
            _tprintf(_T("  pin <HWND>\n"));
            _tprintf(_T("  unpin <HWND>\n"));
            _tprintf(_T("  switch <desktop>\n"));
            _tprintf(_T("  rename <desktop> <name>\n"));
            //_tprintf(_T("  create <name>\n"));
            //_tprintf(_T("  remove <desktop>\n"));
            _tprintf(_T("\n"));
            _tprintf(_T("where <desktop> is one of:>\n"));
            _tprintf(_T("  {current}\n"));
            _tprintf(_T("  {prev}\n"));
            _tprintf(_T("  {next}\n"));
            _tprintf(_T("  [desktop number]\n"));
            _tprintf(_T("  [desktop name]\n"));
        }
    }

    CoUninitialize();

    return EXIT_SUCCESS;
}
