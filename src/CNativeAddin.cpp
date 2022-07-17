#include "pch.h"
#include "NativeAddin.h"

IDispatch* CNativeAddin::Application = nullptr;

CNativeAddin::CNativeAddin()
{
}

HRESULT __stdcall CNativeAddin::OnConnection(IDispatch* Application, ext_ConnectMode ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom)
{
    /*OutputDebugStringW(L"OnConnection()");

    PROCESS_DPI_AWARENESS pProcessDpiAwareness = PROCESS_DPI_UNAWARE;
    HANDLE hProcess = GetCurrentProcess();
    HRESULT hr = GetProcessDpiAwareness(hProcess, &pProcessDpiAwareness);

    DPI_AWARENESS_CONTEXT ctxThreadDpiAwareness = GetThreadDpiAwarenessContext();
    DPI_AWARENESS threadDpiAwareness = GetAwarenessFromDpiAwarenessContext(ctxThreadDpiAwareness);*/

    int argc = 0;
    wchar_t* argv = nullptr;

    testing::InitGoogleTest(&argc, &argv);

    ULONG c = Application->AddRef();
    CNativeAddin::Application = Application;

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

    int result = RUN_ALL_TESTS();

    ULONG c = CNativeAddin::Application->Release();
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

class PowerPointFixture : public ::testing::Test {
protected:
    void SetUp() override {
        this->Application = CNativeAddin::Application;
        this->Application->AddRef();
    }

    void TearDown() override {
        this->Application->Release();
    }

    IDispatch* Application;
};

TEST_F(PowerPointFixture, ApplicationIsValidObject) {
    EXPECT_NE(nullptr, Application);
}

TEST(PowerPointApp, AlwaysPass) {
    EXPECT_TRUE(true);
}