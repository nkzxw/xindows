
#ifndef __XINDOWSUTIL_COM_UNKNOWN_H__
#define __XINDOWSUTIL_COM_UNKNOWN_H__

#define ULREF_IN_DESTRUCTOR     256

#define DECLARE_IUNKNOWN_METHODS                        \
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv); \
    STDMETHOD_(ULONG, AddRef)(void);                    \
    STDMETHOD_(ULONG, Release)(void);

#define DECLARE_STANDARD_IUNKNOWN(cls)                  \
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv); \
    ULONG _ulRefs;                                      \
    STDMETHOD_(ULONG, AddRef)(void)                     \
    {                                                   \
        return ++_ulRefs;                               \
    }                                                   \
    STDMETHOD_(ULONG, Release)(void)                    \
    {                                                   \
        if(--_ulRefs == 0)                              \
        {                                               \
            _ulRefs = ULREF_IN_DESTRUCTOR;              \
            delete this;                                \
            return 0;                                   \
        }                                               \
        return _ulRefs;                                 \
    }                                                   \
    ULONG GetRefs(void)                                 \
    {                                                   \
        return _ulRefs;                                 \
    }

#define DECLARE_PLAIN_IUNKNOWN(cls)                     \
    STDMETHOD(QueryInterface)(REFIID iid, void** ppv)   \
    {                                                   \
        return PrivateQueryInterface(iid, ppv);         \
    }                                                   \
    STDMETHOD_(ULONG, AddRef)(void)                     \
    {                                                   \
        return PrivateAddRef();                         \
    }                                                   \
    STDMETHOD_(ULONG, Release)(void)                    \
    {                                                   \
        return PrivateRelease();                        \
    }

XINDOWS_PUBLIC void ClearInterfaceFn(IUnknown** ppUnk);
XINDOWS_PUBLIC void ClearClassFn(void** ppv, IUnknown* pUnk);
XINDOWS_PUBLIC void ReplaceInterfaceFn(IUnknown** ppUnk, IUnknown* pUnk);
XINDOWS_PUBLIC void ReleaseInterface(IUnknown* pUnk);

//+------------------------------------------------------------------------
//
//  Function:   ClearInterface
//
//  Synopsis:   Sets an interface pointer to NULL, after first calling
//              Release if the pointer was not NULL initially
//
//  Arguments:  [ppI]   *ppI is cleared
//
//-------------------------------------------------------------------------
template<class PI> inline void ClearInterface(PI* ppI)
{
#ifdef _DEBUG
    IUnknown* pUnk = *ppI;
    Assert((void*)pUnk == (void*)*ppI);
#endif
    ClearInterfaceFn((IUnknown**)ppI);
}

//+------------------------------------------------------------------------
//
//  Function:   ClearClass
//
//  Synopsis:   Nulls a pointer to a class, releasing the class via the
//              provided IUnknown implementation if the original pointer
//              is non-NULL.
//
//  Arguments:  [ppT]       Pointer to a class
//              [pUnk]      Pointer to the class's public IUnknown impl
//
//-------------------------------------------------------------------------
template<class PT> inline void ClearClass(PT* ppT, IUnknown* pUnk)
{
    ClearClassFn((void**)ppT, pUnk);
}

//+------------------------------------------------------------------------
//
//  Function:   ReplaceInterface
//
//  Synopsis:   Replaces an interface pointer with a new interface,
//              following proper ref counting rules:
//
//              = *ppI is set to pI
//              = if pI is not NULL, it is AddRef'd
//              = if *ppI was not NULL initially, it is Release'd
//
//              Effectively, this allows pointer assignment for ref-counted
//              pointers.
//
//  Arguments:  [ppI]     Destination pointer in *ppI
//              [pI]      Source pointer in pI
//
//-------------------------------------------------------------------------
template<class PI> inline void ReplaceInterface(PI* ppI, PI pI)
{
#ifdef _DEBUG
    IUnknown* pUnk = *ppI;
    Assert((void*)pUnk == (void*)*ppI);
#endif
    ReplaceInterfaceFn((IUnknown**)ppI, pI);
}

#endif //__XINDOWSUTIL_COM_UNKNOWN_H__