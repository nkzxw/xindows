
#ifndef __XINDOWS_SITE_UTIL_PROPERTYBAG_H__
#define __XINDOWS_SITE_UTIL_PROPERTYBAG_H__

class CStreamWriteBuff;

struct PROPNAMEVALUE
{
    PROPNAMEVALUE()
    {
        memset(this, 0, sizeof(*this));
    }

    HRESULT Set(TCHAR* pchName, VARIANT* pVar);
    void Free();

    VARIANT _varValue;
    CString _cstrName;
};

class CPropertyBag : public IPropertyBag, public IPropertyBag2
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();
    CPropertyBag();
    ~CPropertyBag();

    void Clear();

    // IUnknown methods
    DECLARE_STANDARD_IUNKNOWN(CPropertyBag);

    // IPropertyBag methods
    STDMETHOD(Read)(LPCOLESTR pszPropName, VARIANT* pVar, IErrorLog* pErrorLog);
    STDMETHOD(Write)(LPCOLESTR pszPropName, VARIANT* pVar);

    // IPropertyBag2 methods
    STDMETHOD(Read)(
        ULONG       cProperties, 
        PROPBAG2*   pPB2, 
        IErrorLog*  pErrorLog,
        VARIANT*    pvarValue,
        HRESULT*    phrError);
    STDMETHOD(Write)(ULONG cProperties, PROPBAG2* pPB2, VARIANT* pVar);
    STDMETHOD(CountProperties)(ULONG* pcProperties);
    STDMETHOD(GetPropertyInfo)(
        ULONG       iProperty,
        ULONG       cProperties,
        PROPBAG2*   pPropBag,
        ULONG*      pcProperties);
    STDMETHOD(LoadObject)( 
        LPCOLESTR   pstrName,
        DWORD       dwHint,
        IUnknown*   pUnkObject,
        IErrorLog*  pErrorLog);

    // BUGBUG: istvanc from OHARE -> HACKHACK davidna 5/7/96 Beta1 Hack
    HRESULT GetPropertyBagContents(VARIANT* pVar, IErrorLog* pErrorLog);
    // BUGBUG: istvanc from OHARE -> HACKHACK davidna 5/7/96 Beta1 Hack
    void FreePropertyBagContent();

    // Misc. helpers to add properties to the bag
    HRESULT AddProp(TCHAR* pchName, int cchName, TCHAR* pchValue, int cchValue);
    HRESULT AddProp(TCHAR* pchName, TCHAR* pchValue);
    HRESULT FindAndSetProp(TCHAR* pchName, TCHAR* pchValue);

    PROPNAMEVALUE* Find(TCHAR* pchName, long iLikelyIndex=-1);
    HRESULT Save(CStreamWriteBuff* pStreamWrBuff);
    HRESULT Read(LPCOLESTR pszPropName, VARIANT* pVar, IErrorLog* pErrorLog, long iLikelyIndex=-1);

    // Data members
    DECLARE_CDataAry(CAryProps, PROPNAMEVALUE)
    CAryProps   _aryProps;          // Array of property structs
    CElement*   _pElementExpandos;  // when set, each Read method will remove the corresponding expando on the element
};

#endif //__XINDOWS_SITE_UTIL_PROPERTYBAG_H__