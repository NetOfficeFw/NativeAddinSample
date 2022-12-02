#pragma once

using namespace AddInDesignerObjects;

class ATL_NO_VTABLE CNativeAddin :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CNativeAddin, &CLSID_NativeAddin>,
    public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects, 1, 0>
{
public:
    CNativeAddin();

    DECLARE_REGISTRY_RESOURCEID(IDR_REGNATIVEADDIN)
    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CNativeAddin)
        COM_INTERFACE_ENTRY(_IDTExtensibility2)
    END_COM_MAP()

public:
    //IDTExtensibility2.
    STDMETHOD(OnConnection)(IDispatch* Application, ext_ConnectMode ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom);

    STDMETHOD(OnAddInsUpdate)(SAFEARRAY** custom);

    STDMETHOD(OnStartupComplete)(SAFEARRAY** custom);

    STDMETHOD(OnBeginShutdown)(SAFEARRAY** custom);

    STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY** custom);

    static IDispatch* Application;
};


OBJECT_ENTRY_AUTO(__uuidof(NativeAddin), CNativeAddin)