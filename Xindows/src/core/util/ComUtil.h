
#ifndef __XINDOWS_CORE_UTIL_COMUTIL_H__
#define __XINDOWS_CORE_UTIL_COMUTIL_H__

#define DECLARE_SUBOBJECT_IUNKNOWN(parent_cls, parent_mth)      \
    DECLARE_IUNKNOWN_METHODS                                    \
    parent_cls* parent_mth();                                   \
    BOOL IsMyParentAlive();

#define DECLARE_SUBOBJECT_IUNKNOWN_NOQI(parent_cls, parent_mth) \
    STDMETHOD_(ULONG, AddRef)();                                \
    STDMETHOD_(ULONG, Release)();                               \
    parent_cls* parent_mth();                                   \
    BOOL IsMyParentAlive();

#define IMPLEMENT_SUBOBJECT_IUNKNOWN(cls, parent_cls, parent_mth, member) \
    inline parent_cls* cls::parent_mth()                        \
    {                                                           \
        return CONTAINING_RECORD(this, parent_cls, member);     \
    }                                                           \
    inline BOOL cls::IsMyParentAlive(void)                      \
    { return parent_mth()->GetObjectRefs()!=0; }                \
    STDMETHODIMP_(ULONG) cls::AddRef( )                         \
    { return parent_mth()->SubAddRef(); }                       \
    STDMETHODIMP_(ULONG) cls::Release( )                        \
    { return parent_mth()->SubRelease(); }

#define DECLARE_FORMS_SUBOBJECT_IUNKNOWN(parent_cls)            \
    DECLARE_SUBOBJECT_IUNKNOWN(parent_cls, My##parent_cls)
#define DECLARE_FORMS_SUBOBJECT_IUNKNOWN_NOQI(parent_cls)       \
    DECLARE_SUBOBJECT_IUNKNOWN_NOQI(parent_cls, My##parent_cls)
#define IMPLEMENT_FORMS_SUBOBJECT_IUNKNOWN(cls, parent_cls, member) \
    IMPLEMENT_SUBOBJECT_IUNKNOWN(cls, parent_cls, My##parent_cls, member)


inline BOOL DISPID_NOT_FOUND(HRESULT hr)
{
    return hr==DISP_E_MEMBERNOTFOUND||hr==TYPE_E_LIBNOTREGISTERED;
}

inline HRESULT ClampITFResult(HRESULT hr)
{
    return (HRESULT_FACILITY(hr)==FACILITY_ITF)?E_FAIL:hr;
}

HRESULT ValidateInvoke(DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);

HRESULT ReadyStateInvoke(DISPID dispid, WORD wFlags, long lReadyState, VARIANT* pvarResult);

BOOL IsSameObject(IUnknown* pUnkLeft, IUnknown* pUnkRight);

HRESULT FormSetClipboard(IDataObject* pdo);

#define EVENTPARAMS_MAX     16

#endif //__XINDOWS_CORE_UTIL_COMUTIL_H__