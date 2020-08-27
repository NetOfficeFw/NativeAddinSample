#include "pch.h"
#include "NativeAddin.h"

CNativeAddin::CNativeAddin()
{
}

HRESULT __stdcall CNativeAddin::OnConnection(IDispatch* Application, ext_ConnectMode ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom)
{
    OutputDebugStringW(L"OnConnection()");

    PROCESS_DPI_AWARENESS pProcessDpiAwareness = PROCESS_DPI_UNAWARE;
    HANDLE hProcess = GetCurrentProcess();
    HRESULT hr = GetProcessDpiAwareness(hProcess, &pProcessDpiAwareness);

    DPI_AWARENESS_CONTEXT ctxThreadDpiAwareness = GetThreadDpiAwarenessContext();
    DPI_AWARENESS threadDpiAwareness = GetAwarenessFromDpiAwarenessContext(ctxThreadDpiAwareness);

    return S_OK;
}

HRESULT __stdcall CNativeAddin::OnAddInsUpdate(SAFEARRAY** custom)
{
    OutputDebugStringW(L"OnAddInsUpdate()");
    return S_OK;
}

HRESULT __stdcall CNativeAddin::OnStartupComplete(SAFEARRAY** custom)
{
    OutputDebugStringW(L"OnStartupComplete()");
    return S_OK;
}

HRESULT __stdcall CNativeAddin::OnBeginShutdown(SAFEARRAY** custom)
{
    OutputDebugStringW(L"OnBeginShutdown()");
    return S_OK;
}

HRESULT __stdcall CNativeAddin::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY** custom)
{
    OutputDebugStringW(L"OnDisconnection()");
    return S_OK;
}