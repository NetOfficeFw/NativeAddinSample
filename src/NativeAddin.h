#pragma once

class ATL_NO_VTABLE CNativeAddin :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CNativeAddin, &CLSID_NativeAddin>
{
public:
    CNativeAddin();

    DECLARE_REGISTRY_RESOURCEID(IDR_REGNATIVEADDIN)
    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CNativeAddin)
    END_COM_MAP()
};


OBJECT_ENTRY_AUTO(__uuidof(NativeAddin), CNativeAddin)