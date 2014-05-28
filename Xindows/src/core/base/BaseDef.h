
#ifndef __XINDOWS_CORE_BASE_BASEDEF_H__
#define __XINDOWS_CORE_BASE_BASEDEF_H__

#define FORMS_BUFLEN    255             // constant size for stack-based buffers used for LoadString()

#define VB_TRUE         VARIANT_TRUE    // TRUE for VARIANT_BOOL
#define VB_FALSE        VARIANT_FALSE   // FALSE for VARIANT_BOOL

#define LCID_SCRIPTING  0x0409          // Mandatory VariantChangeTypeEx localeID

#define VTS_NONE        ""              // used for members with 0 params

#define pdlUrlLen       4096
#define pdlToken        128             // strings that really are some form of a token
#define pdlLength       128             // strings that really are numeric lengths
#define pdlColor        128             // strings that really are color values
#define pdlNoLimit      0xFFFF          // strings that have no limit on their max lengths
#define pdlEvent        pdlNoLimit      // strings that could be assigned to onfoo event properties

// flag for determine when gion is called from ie40. was in dispex.idl
#define fdexFromGetIdsOfNames       0x80000000L


#define NOVTABLE                    __declspec(novtable)

#define CVOID_CAST(pDerObject)      ((CVoid*)(void*)pDerObject)

#define OK(r)                       (SUCCEEDED(r))

#define CTL_E_METHODNOTAPPLICABLE   STD_CTL_SCODE(444)
#define CTL_E_CANTMOVEFOCUSTOCTRL   STD_CTL_SCODE(2110)
#define CTL_E_CONTROLNEEDSFOCUS     STD_CTL_SCODE(2185)
#define CTL_E_INVALIDPICTURETYPE    STD_CTL_SCODE(485)
#define CTL_E_INVALIDPASTETARGET    CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+0)
#define CTL_E_INVALIDPASTESOURCE    CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+1)
#define CTL_E_MISMATCHEDTAG         CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+2)
#define CTL_E_INCOMPATIBLEPOINTERS  CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+3)
#define CTL_E_UNPOSITIONEDPOINTER   CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+4)
#define CTL_E_UNPOSITIONEDELEMENT   CUSTOM_CTL_SCODE(CTL_E_CUSTOM_FIRST+5)

#define ENSURE_BOOL(_x_)            (!!(BOOL)(_x_))
#define EQUAL_BOOL(_x_,_y_)         (ENSURE_BOOL(_x_) == ENSURE_BOOL(_y_))
#define ISBOOL(_x_)                 (ENSURE_BOOL(_x_) == (BOOL)(_x_))

#define VARIANT_BOOL_FROM_BOOL(_x_) ((_x_) ? VB_TRUE : VB_FALSE)
#define BOOL_FROM_VARIANT_BOOL(_x_) ((VB_TRUE==_x_) ? TRUE : FALSE)

typedef enum _FRAMEOPTIONS_FLAGS
{
    FRAMEOPTIONS_SCROLL_YES	= 0x1,
    FRAMEOPTIONS_SCROLL_NO	= 0x2,
    FRAMEOPTIONS_SCROLL_AUTO= 0x4,
    FRAMEOPTIONS_NO3DBORDER	= 0x10,
} FRAMEOPTIONS_FLAGS;

typedef enum _htmlCaptionAlign
{
    htmlCaptionAlignNotSet	= 0,
    htmlCaptionAlignLeft    = 1,
    htmlCaptionAlignCenter	= 2,
    htmlCaptionAlignRight	= 3,
    htmlCaptionAlignJustify	= 4,
    htmlCaptionAlignTop	    = 5,
    htmlCaptionAlignBottom	= 6,
    htmlCaptionAlign_Max	= 2147483647L
} htmlCaptionAlign;

typedef enum _htmlCaptionVAlign
{
    htmlCaptionVAlignNotSet	= 0,
    htmlCaptionVAlignTop    = 1,
    htmlCaptionVAlignBottom	= 2,
    htmlCaptionVAlign_Max   = 2147483647L
} htmlCaptionVAlign;

typedef enum _htmlCellVAlign
{
    htmlCellVAlignNotSet    = 0,
    htmlCellVAlignTop       = 1,
    htmlCellVAlignMiddle    = 2,
    htmlCellVAlignBottom    = 3,
    htmlCellVAlignBaseline	= 4,
    htmlCellVAlignCenter    = htmlCellVAlignMiddle,
    htmlCellVAlign_Max      = 2147483647L
} htmlCellVAlign;


struct THREADSTATE;

typedef void (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONCALL)(DWORD_PTR);

#ifdef _DEBUG
#define GWPostMethodCall(pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg) \
    GWPostMethodCallEx(GetThreadState(), pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg)
#define GWPostMethodCallEx(pts, pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg) \
    _GWPostMethodCallEx(pts, pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg)
HRESULT _GWPostMethodCallEx(THREADSTATE* pts, void* pvObject, PFN_VOID_ONCALL pfnCall, DWORD_PTR dwContext, BOOL fIgnoreContext, char* pszCallDbg);
#else
#define GWPostMethodCall(pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg) \
    _GWPostMethodCallEx(GetThreadState(), pvObject, pfnCall, dwContext, fIgnoreContext)
#define GWPostMethodCallEx(pts, pvObject, pfnCall, dwContext, fIgnoreContext, pszCallDbg) \
    _GWPostMethodCallEx(pts, pvObject, pfnCall, dwContext, fIgnoreContext)
HRESULT _GWPostMethodCallEx(THREADSTATE* pts, void* pvObject, PFN_VOID_ONCALL pfnCall, DWORD_PTR dwContext, BOOL fIgnoreContext);
#endif

void GWKillMethodCall(void* pvObject, PFN_VOID_ONCALL pfnOnCall, DWORD_PTR dwContext);
void GWKillMethodCallEx(THREADSTATE* pts, void* pvObject, PFN_VOID_ONCALL pfnOnCall, DWORD_PTR dwContext);

#define CALL_METHOD(pObj, pfnMethod, args)  (pObj->*pfnMethod) args

#define ONCALL_METHOD(klass, fn, FN)        (PFN_VOID_ONCALL)&klass::fn
#define NV_DECLARE_ONCALL_METHOD(fn, FN, args) \
    static void STDMETHODCALLTYPE FN args; \
    void STDMETHODCALLTYPE fn args

//+-------------------------------------------------------------------
//
//  Macros for declaring Non-Virtual methods in support of tearoff
//              implementation of standard OLE interfaces.  (Tearoffs require
//              STDMETHODCALLTYPE.)
//
//--------------------------------------------------------------------
#define NV_STDMETHOD(method)                HRESULT STDMETHODCALLTYPE method
#define NV_STDMETHOD_(type, method)         type STDMETHODCALLTYPE method

#define NV_DECLARE_PROPERTY_METHOD(fn, FN, args) \
    HRESULT STDMETHODCALLTYPE fn args
#define PROPERTY_METHOD(kind, getset, klass, fn, FN) \
    (PFN_##kind##PROP##getset)(PFNB_##kind##PROP##getset)&klass::fn

#define NV_DECLARE_ONTICK_METHOD(fn, FN, args) \
    HRESULT STDMETHODCALLTYPE fn args
#define DECLARE_ONTICK_METHOD(fn, FN, args) \
    virtual HRESULT STDMETHODCALLTYPE fn args
#define ONTICK_METHOD(klass, fn, FN)        (PFN_VOID_ONTICK)&klass::fn

#define NV_DECLARE_ONMESSAGE_METHOD(fn, FN, args) \
    LRESULT STDMETHODCALLTYPE fn args
#define ONMESSAGE_METHOD(klass, fn, FN)     (PFN_VOID_ONMESSAGE)&klass::fn

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONTICK)(UINT idTimer);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONCOMMAND)(int id, HWND hwndCtrl, UINT codeNotify);
typedef LRESULT (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONMESSAGE)(UINT msg, WPARAM wParam, LPARAM lParam);
typedef void    (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONCALL)(DWORD_PTR);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_VOID_ONSEND)(void*);

#define DECLARE_PRIVATE_QI_FUNCS(cls)       STDMETHOD(PrivateQueryInterface)(REFIID, void**);

#define DISPID_XOBJ_MIN                     0x80010000
#define DISPID_XOBJ_MAX                     0x8001FFFF
#define DISPID_XOBJ_BASE                    DISPID_XOBJ_MIN


//+------------------------------------------------------------------------
//
// Useful combinations of flags for IMsoCommandTarget
//
//-------------------------------------------------------------------------
#define MSOCMDSTATE_DISABLED                MSOCMDF_SUPPORTED
#define MSOCMDSTATE_UP                      (MSOCMDF_SUPPORTED|MSOCMDF_ENABLED)
#define MSOCMDSTATE_DOWN                    (MSOCMDF_SUPPORTED|MSOCMDF_ENABLED|MSOCMDF_LATCHED)
#define MSOCMDSTATE_NINCHED                 (MSOCMDF_SUPPORTED|MSOCMDF_ENABLED|MSOCMDF_NINCHED)


HRESULT FormsTrackPopupMenu(HMENU hMenu, UINT fuFlags, int x, int y,
                            HWND hwndMessage, int* piSelection);

typedef int (__stdcall *STRINGCOMPAREFN)(const TCHAR*, const TCHAR*);

#endif //__XINDOWS_CORE_BASE_BASEDEF_H__