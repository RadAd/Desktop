#include <windows.h>
#include <atlbase.h>
#include <shlobj.h>
#include <winstring.h>
#include <assert.h>

#include <tchar.h>
#include "Win10Desktops.h"
#include "ComUtils.h"
#include "VDNotification.h"
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

template <class VD, class VD2 = VD, class VDMI>
int ShowCurrentDesktop(VDMI* pDesktopManagerInternal)
{
    CComPtr<VD> pCurrentDesktop;
    CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pCurrentDesktop), _T("GetCurrentDesktop"), EXIT_FAILURE);

    CComPtr<IObjectArray> pDesktopArray;
    CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), EXIT_FAILURE);

    int dn = 0;
    for (CComPtr<VD2> pDesktop : ObjectArrayRange<VD2>(pDesktopArray))
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
        return ShowCurrentDesktop<Win10::IVirtualDesktop, Win10::IVirtualDesktop2>(-pDesktopManagerInternal10);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return ShowCurrentDesktop<Win11::IVirtualDesktop>(-pDesktopManagerInternal11);
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

template <class VD, class VDMI>
int ListDesktops(IServiceProvider* pServiceProvider, VDMI* pDesktopManagerInternal, bool bShowViews)
{
    CComPtr<IObjectArray> pDesktopArray;
    CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), EXIT_FAILURE);

    if (bShowViews)
    {
        CComPtr<IApplicationViewCollection> pApplicationViewCollection;
        CHECK(pServiceProvider->QueryService(IID_IApplicationViewCollection, &pApplicationViewCollection), _T("obtaining IID_IApplicationViewCollection"), EXIT_FAILURE);

        CComPtr<IObjectArray> pViewArray;
        CHECK(pApplicationViewCollection->GetViewsByZOrder(&pViewArray), _T("GetViewsByZOrder"), EXIT_FAILURE);

        CComPtr<IVirtualDesktopPinnedApps> pVirtualDesktopPinnedApps;
        CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopPinnedApps, &pVirtualDesktopPinnedApps), _T("obtaining IID_IVirtualDesktopPinnedApps"), EXIT_FAILURE);

        int dn = 0;
        for (CComPtr<VD> pDesktop : ObjectArrayRange<VD>(pDesktopArray))
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
        for (CComPtr<VD> pDesktop : ObjectArrayRange<VD>(pDesktopArray))
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
        return ListDesktops<Win10::IVirtualDesktop2>(pServiceProvider, -pDesktopManagerInternal10, bShowViews);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return ListDesktops<Win11::IVirtualDesktop>(pServiceProvider, -pDesktopManagerInternal11, bShowViews);
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
    // TODO Support
    // by title
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

template <class VD, class VDMI>
CComPtr<VD> GetAdjacentDesktop(VDMI* pDesktopManagerInternal, AdjacentDesktop uDirection)
{
    CComPtr<VD> pDesktop;
    CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pDesktop), _T("GetCurrentDesktop"), {});

    CComPtr<VD> pAdjacentDesktop;
    CHECK(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, uDirection, &pAdjacentDesktop), _T("GetAdjacentDesktop"), {});

    return pAdjacentDesktop;
}

template <class VD, class VD2 = VD, class VDMI>
CComPtr<VD> GetDesktop(VDMI* pDesktopManagerInternal, const LPCTSTR sDesktop)
{
    // TODO Support
    // by guid
    if (sDesktop == nullptr)
        return {};
    else if (_tcsicmp(sDesktop, _T("{current}")) == 0)
    {
        CComPtr<VD> pDesktop;
        CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pDesktop), _T("GetCurrentDesktop"), {});
        return pDesktop;
    }
    else if (_tcsicmp(sDesktop, _T("{prev}")) == 0)
        return GetAdjacentDesktop<VD>(pDesktopManagerInternal, LeftDirection);
    else if (_tcsicmp(sDesktop, _T("{next}")) == 0)
        return GetAdjacentDesktop<VD>(pDesktopManagerInternal, RightDirection);
    else if (std::_istdigit(sDesktop[0]))
    {
        int i = _tstoi(sDesktop) - 1;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), {});

        CComPtr<VD> pDesktop;
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
        CComPtr<VD> pRetDesktop;

        CComPtr<IObjectArray> pDesktopArray;
        CHECK(GetDesktops(pDesktopManagerInternal, &pDesktopArray), _T("GetDesktops"), {});

        for (CComPtr<VD2> pDesktop : ObjectArrayRange<VD2>(pDesktopArray))
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

template <class VD, class VDMI>
int DoSwitchDesktop(VDMI* pDesktopManagerInternal, VD* pDesktop)
{
    if (pDesktop)
    {
        CHECK(SwitchDesktop(pDesktopManagerInternal, pDesktop), _T("SwitchDesktop"), EXIT_FAILURE);
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
        return DoSwitchDesktop(-pDesktopManagerInternal10, -GetDesktop<Win10::IVirtualDesktop, Win10::IVirtualDesktop2>(-pDesktopManagerInternal10, strDesktop));
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return DoSwitchDesktop(-pDesktopManagerInternal11, -GetDesktop<Win11::IVirtualDesktop>(-pDesktopManagerInternal11, strDesktop));
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

template <class VD, class VDMI>
int RenameDesktop(VDMI* pDesktopManagerInternal, VD* pDesktop, LPCTSTR strName)
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
        return RenameDesktop(-pDesktopManagerInternal10, -GetDesktop<Win10::IVirtualDesktop, Win10::IVirtualDesktop2>(-pDesktopManagerInternal10, strDesktop), strName);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return RenameDesktop(-pDesktopManagerInternal11, -GetDesktop<Win11::IVirtualDesktop>(-pDesktopManagerInternal11, strDesktop), strName);
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

template <class VD, class VD2 = VD, class VDMI>
int CreateDesktop(VDMI* pDesktopManagerInternal, LPCTSTR strName)
{
    CComPtr<VD> pNewDesktop;
    CHECK(CreateDesktop(pDesktopManagerInternal, &pNewDesktop), _T("CreateDesktop"), EXIT_FAILURE);

    if (strName != nullptr)
    {
        CComQIPtr<VD2> pNewDesktop2(pNewDesktop);
        return RenameDesktop(pDesktopManagerInternal, -pNewDesktop2, strName);
    }
    else
        return EXIT_SUCCESS;
}

int CreateDesktop(IServiceProvider* pServiceProvider, LPCTSTR strName)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal2> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    //CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return CreateDesktop<Win10::IVirtualDesktop, Win10::IVirtualDesktop2>(-pDesktopManagerInternal10, strName);
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return CreateDesktop<Win11::IVirtualDesktop>(-pDesktopManagerInternal11, strName);
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

template <class VD, class VDMI>
int RemoveDesktop(VDMI* pDesktopManagerInternal, VD* pDesktop)
{
    if (pDesktop)
    {
        CComPtr<VD> pFallbackDesktop;
        CHECK(GetCurrentDesktop(pDesktopManagerInternal, &pFallbackDesktop), _T("GetCurrentDesktop"), EXIT_FAILURE);
        if (pFallbackDesktop.IsEqualObject(pDesktop))
        {
            if (FAILED(pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, LeftDirection, &pFallbackDesktop)))
                pDesktopManagerInternal->GetAdjacentDesktop(pDesktop, RightDirection, &pFallbackDesktop);
        }

        CHECK(pDesktopManagerInternal->RemoveDesktop(pDesktop, pFallbackDesktop), _T("RemoveDesktop"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
    else
    {
        _ftprintf(stderr, _T("Unknown desktop\n"));
        return EXIT_FAILURE;
    }
}


int RemoveDesktop(IServiceProvider* pServiceProvider, LPCTSTR strDesktop)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    //CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
        return RemoveDesktop(-pDesktopManagerInternal10, -GetDesktop<Win10::IVirtualDesktop, Win10::IVirtualDesktop2>(-pDesktopManagerInternal10, strDesktop));
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
        return RemoveDesktop(-pDesktopManagerInternal11, -GetDesktop<Win11::IVirtualDesktop>(-pDesktopManagerInternal11, strDesktop));
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

class VDNotification: public IVDNotification
{
public:
    // Win10::IVirtualDesktopNotification
    virtual void VirtualDesktopCreated(Win10::IVirtualDesktop* pDesktop) override
    {
        CComQIPtr<Win10::IVirtualDesktop2> pDesktop2(pDesktop);

        _tprintf(_T("VirtualDesktopCreated\n "));
        PrintDesktop(-pDesktop2, 0);
    }

    virtual void VirtualDesktopDestroyed(Win10::IVirtualDesktop* pDesktopDestroyed, Win10::IVirtualDesktop* pDesktopFallback) override
    {
        CComQIPtr<Win10::IVirtualDesktop2> pDesktopDestroyed2(pDesktopDestroyed);
        CComQIPtr<Win10::IVirtualDesktop2> pDesktopFallback2(pDesktopFallback);

        _tprintf(_T("VirtualDesktopDestroyed\n Destroyed: "));
        PrintDesktop(-pDesktopDestroyed2, 0);
        _tprintf(_T(" Fallback: "));
        PrintDesktop(-pDesktopFallback2, 0);
    }

    virtual void CurrentVirtualDesktopChanged(Win10::IVirtualDesktop* pDesktopOld, Win10::IVirtualDesktop* pDesktopNew) override
    {
        CComQIPtr<Win10::IVirtualDesktop2> pDesktopOld2(pDesktopOld);
        CComQIPtr<Win10::IVirtualDesktop2> pDesktopNew2(pDesktopNew);

        _tprintf(_T("CurrentVirtualDesktopChanged\n Old: "));
        PrintDesktop(-pDesktopOld2, 0);
        _tprintf(_T(" New: "));
        PrintDesktop(-pDesktopNew2, 0);
    }

    // Win11::IVirtualDesktopNotification
    virtual void VirtualDesktopCreated(Win11::IVirtualDesktop* pDesktop) override
    {
        _tprintf(_T("VirtualDesktopCreated\n "));
        PrintDesktop(pDesktop, 0);
    }

    virtual void VirtualDesktopDestroyed(Win11::IVirtualDesktop* pDesktopDestroyed, Win11::IVirtualDesktop* pDesktopFallback) override
    {
        _tprintf(_T("VirtualDesktopDestroyed\n Destroyed: "));
        PrintDesktop(pDesktopDestroyed, 0);
        _tprintf(_T(" Fallback: "));
        PrintDesktop(pDesktopFallback, 0);
    }

    virtual void VirtualDesktopMoved(Win11::IVirtualDesktop* pDesktop, int64_t oldIndex, int64_t newIndex) override
    {
        _tprintf(_T("VirtualDesktopMoved\n "));
        PrintDesktop(pDesktop, 0);
        _tprintf(_T(" From: %d To: %d\n"), static_cast<int>(oldIndex), static_cast<int>(newIndex));
    }

    virtual void VirtualDesktopNameChanged(Win11::IVirtualDesktop* pDesktop, HSTRING name) override
    {
        _tprintf(_T("VirtualDesktopNameChanged\n "));
        PrintDesktop(pDesktop, 0);
        _tprintf(_T(" Name:s %ls\n"), WindowsGetStringRawBuffer(name, nullptr));
    }

    virtual void CurrentVirtualDesktopChanged(Win11::IVirtualDesktop* pDesktopOld, Win11::IVirtualDesktop* pDesktopNew) override
    {
        _tprintf(_T("CurrentVirtualDesktopChanged\nOld: "));
        PrintDesktop(pDesktopOld, 0);
        _tprintf(_T(" New: "));
        PrintDesktop(pDesktopNew, 0);
    }
};

UINT DoMessageLoop()
{
    MSG msg = { };
    while (GetMessage(&msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return msg.message;
}

int WatchDesktops(IServiceProvider* pServiceProvider)
{
    CComPtr<Win10::IVirtualDesktopManagerInternal> pDesktopManagerInternal10;
    CComPtr<Win11::IVirtualDesktopManagerInternal> pDesktopManagerInternal11;
    //CHECK(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal), _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);

    VDNotification TheVDNotification;
    DWORD idVirtualDesktopNotification = 0;
    CComPtr<VirtualDesktopNotification> pNotify = new VirtualDesktopNotification(&TheVDNotification);

    if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal10)))
    {
        CComPtr<Win10::IVirtualDesktopNotificationService> pVirtualNotificationService;
        CHECK(pServiceProvider->QueryService(CLSID_VirtualNotificationService, &pVirtualNotificationService), _T("creating CLSID_IVirtualNotificationService"), EXIT_FAILURE);
        if (pVirtualNotificationService)
            CHECK(pVirtualNotificationService->Register(pNotify, &idVirtualDesktopNotification), _T("Register DesktopNotificationService"), EXIT_FAILURE);
        DoMessageLoop();
        CHECK(pVirtualNotificationService->Unregister(idVirtualDesktopNotification), _T("Unregister DesktopNotificationService"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
    else if (SUCCEEDED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &pDesktopManagerInternal11)))
    {
        CComPtr<Win11::IVirtualDesktopNotificationService> pVirtualNotificationService;
        CHECK(pServiceProvider->QueryService(CLSID_VirtualNotificationService, &pVirtualNotificationService), _T("creating CLSID_IVirtualNotificationService"), EXIT_FAILURE);
        if (pVirtualNotificationService)
            CHECK(pVirtualNotificationService->Register(pNotify, &idVirtualDesktopNotification), _T("Register DesktopNotificationService"), EXIT_FAILURE);
        DoMessageLoop();
        CHECK(pVirtualNotificationService->Unregister(idVirtualDesktopNotification), _T("Unregister DesktopNotificationService"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
    else
    {
        CHECK(false, _T("obtaining CLSID_VirtualDesktopManagerInternal"), EXIT_FAILURE);
        return EXIT_SUCCESS;
    }
}

const TCHAR* GetNextArg(int argc, const TCHAR* const argv[], int* arg)
{
    if (argc > *arg)
        return argv[++(*arg)];
    else
        return nullptr;
}

template <class T>
class AutoCall
{
public:
    AutoCall(T* f)
        : f(f)
    {
    }

    ~AutoCall()
    {
        f();
    }
private:
    T* f;
};

template <class T>
AutoCall<T> MakeAutoCall(T* f)
{
    return AutoCall<T>(f);
};

int _tmain(const int argc, const TCHAR* const argv[])
{
    CHECK(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), _T("CoInitialize"), EXIT_FAILURE);
    auto AutoCoUninitialize(MakeAutoCall(CoUninitialize));

    CComPtr<IServiceProvider> pServiceProvider;
    CHECK(pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER), _T("creating CLSID_ImmersiveShell"), EXIT_FAILURE);

    int arg = 0;
    const TCHAR* cmd = GetNextArg(argc, argv, &arg);

    bool showusage = true;

    if (cmd != nullptr)
    {
        if (_tcsicmp(cmd, _T("current")) == 0)
            return ShowCurrentDesktop(pServiceProvider);
        else if (_tcsicmp(cmd, _T("views")) == 0)
            return ListViews(pServiceProvider);
        else if (_tcsicmp(cmd, _T("list")) == 0)
        {
            const TCHAR* showviews = GetNextArg(argc, argv, &arg);
            bool bshowviews = showviews != nullptr && _tcsicmp(showviews, _T("/views")) == 0;
            return ListDesktops(pServiceProvider, bshowviews);
        }
        else if (_tcsicmp(cmd, _T("pin")) == 0)
        {
            const TCHAR* wnd = GetNextArg(argc, argv, &arg);
            if (wnd != nullptr)
                return PinView(pServiceProvider, GetWindow(wnd), TRUE);
        }
        else if (_tcsicmp(cmd, _T("unpin")) == 0)
        {
            const TCHAR* wnd = GetNextArg(argc, argv, &arg);
            if (wnd != nullptr)
                return PinView(pServiceProvider, GetWindow(wnd), FALSE);
        }
        else if (_tcsicmp(cmd, _T("switch")) == 0)
        {
            const TCHAR* desktop = GetNextArg(argc, argv, &arg);
            if (desktop != nullptr)
                return SwitchDesktop(pServiceProvider, desktop);
        }
        else if (_tcsicmp(cmd, _T("create")) == 0)
        {
            const TCHAR* name = GetNextArg(argc, argv, &arg);
            return CreateDesktop(pServiceProvider, name);
        }
        else if (_tcsicmp(cmd, _T("remove")) == 0)
        {
            const TCHAR* desktop = GetNextArg(argc, argv, &arg);
            if (desktop != nullptr)
                return RemoveDesktop(pServiceProvider, desktop);
        }
        else if (_tcsicmp(cmd, _T("rename")) == 0)
        {
            const TCHAR* desktop = GetNextArg(argc, argv, &arg);
            const TCHAR* name = GetNextArg(argc, argv, &arg);
            if (desktop != nullptr && name != nullptr)
                return RenameDesktop(pServiceProvider, desktop, name);
        }
        else if (_tcsicmp(cmd, _T("watch")) == 0)
        {
            return WatchDesktops(pServiceProvider);
        }
        else
        {
            _ftprintf(stderr, _T("Unknown cmd: %s\n"), cmd);
            return EXIT_FAILURE;
        }

    }

    if (showusage)
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
        _tprintf(_T("  create <name>\n"));
        _tprintf(_T("  remove <desktop>\n"));
        _tprintf(_T("  watch\n"));
        _tprintf(_T("\n"));
        _tprintf(_T("where <desktop> is one of:\n"));
        _tprintf(_T("  {current}\n"));
        _tprintf(_T("  {prev}\n"));
        _tprintf(_T("  {next}\n"));
        _tprintf(_T("  [desktop number]\n"));
        _tprintf(_T("  [desktop name]\n"));
    }

    return EXIT_SUCCESS;
}
