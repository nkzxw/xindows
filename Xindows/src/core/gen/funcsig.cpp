
#include "stdafx.h"

static const CustomHandler  _HandlerTable[] = {
   Method_void_BSTR_VARIANT_oDoLONG,
   Method_VARIANTp_BSTR_oDoLONG,
   Method_VARIANTBOOLp_BSTR_oDoLONG,
   Method_void_BSTR_BSTR_oDoBSTR,
   Method_VARIANTp_BSTR,
   Method_VARIANTBOOLp_BSTR,
   Method_VARIANTBOOLp_BSTR_IDispatchp,
   Method_void_BSTR_IDispatchp,
   GS_BSTR,
   GS_PropEnum,
   GS_VARIANT,
   GS_VARIANTBOOL,
   GS_long,
   GS_float,
   Method_BSTRp_void,
   G_PropEnum,
   G_VARIANT,
   G_BSTR,
   G_long,
   G_IUnknownp,
   Method_VARIANTp_VARIANTp,
   G_IDispatchp,
   Method_void_o0oVARIANT,
   Method_VARIANTBOOLp_IDispatchp,
   Method_void_BSTR_BSTR,
   G_VARIANTBOOL,
   Method_void_void,
   Method_void_oDoVARIANTBOOL,
   Method_BSTRp_long_long,
   Method_IDispatchpp_void,
   GS_short,
   Method_void_IUnknownp,
   Method_VARIANTBOOLp_void,
   Method_IDispatchpp_IDispatchp_o0oVARIANT,
   Method_IDispatchpp_IDispatchp,
   Method_IDispatchpp_IDispatchp_IDispatchp,
   Method_IDispatchpp_VARIANTBOOL,
   Method_void_IDispatchp,
   Method_IDispatchpp_IDispatchp_BSTR,
   Method_IDispatchpp_oDoVARIANTBOOL,
   Method_IDispatchpp_BSTR_IDispatchp,
   Method_BSTRp_BSTR,
   Method_BSTRp_BSTR_BSTR,
   Method_longp_BSTR_o0oVARIANTp,
   Method_VARIANTBOOLp_long,
   Method_IDispatchpp_BSTR,
   Method_IDispatchpp_long,
   Method_IDispatchpp_o0oVARIANTp,
   Method_IDispatchpp_VARIANT,
   Method_IDispatchpp_o0oVARIANT_o0oVARIANT,
   Method_VARIANTBOOLp_BSTR_oDolong_oDolong,
   Method_void_long_long,
   Method_longp_BSTR_oDolong,
   Method_longp_BSTR_IDispatchp,
   Method_void_BSTR,
   Method_VARIANTBOOLp_BSTR_oDoVARIANTBOOL_o0oVARIANT,
   Method_longp_BSTR_BSTR_oDolong,
   Method_void_long,
   Method_VARIANTp_long,
   G_short,
   Method_void_o0oVARIANTp,
   Method_VARIANTBOOLp_BSTR_o0oVARIANT,
   Method_void_VARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT,
   Method_VARIANTBOOLp_BSTR_BSTR_o0oVARIANT,
   Method_void_IDispatchp_o0oVARIANT,
   Method_void_oDolong,
   Method_void_SAFEARRAYPVARIANTP,
   Method_IDispatchpp_oDoBSTR_o0oVARIANT_o0oVARIANT_o0oVARIANT,
   Method_IDispatchpp_long_long,
   Method_IDispatchpp_oDoBSTR_oDolong,
   Method_longp_VARIANTp_long_o0oVARIANTp,
   Method_void_oDoBSTR,
   Method_VARIANTBOOLp_oDoBSTR,
   Method_VARIANTp_oDoBSTR_oDoBSTR,
   Method_IDispatchpp_oDoBSTR_oDoBSTR_oDoBSTR_oDoVARIANTBOOL,
   Method_VARIANTp_BSTR_o0oVARIANTp_o0oVARIANTp,
   Method_void_BSTR_o0oVARIANT_oDoBSTR,
   Method_VARIANTp_BSTR_oDoBSTR,
   Method_IDispatchpp_oDoBSTR_o0oVARIANTp_o0oVARIANTp,
   GS_LONG,
   Method_IDispatchpp_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT,
   GS_IDispatchp,
   Method_IDispatchpp_BSTR_o0oVARIANTp,
   Method_VARIANTBOOLp_BSTR_VARIANTp,
   NULL
};

HRESULT Method_void_BSTR_VARIANT_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT, LONG);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT,
                                VT_I4,
                          };
    BSTR    param1;
    VARIANT    param2;
    LONG    param3 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    return hr;
}
HRESULT Method_VARIANTp_BSTR_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, LONG, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_I4,
                          };
    BSTR    param1;
    LONG    param2 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, pvarResult);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, LONG, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_I4,
                          };
    BSTR    param1;
    LONG    param2 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_void_BSTR_BSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, BSTR);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                                VT_BSTR,
                          };
    BSTR    param1;
    BSTR    param2;
    BSTR    paramDef3 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param3 = paramDef3;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3);

Cleanup:
    SysFreeString(paramDef3);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    if (ulAlloc & 4)
        SysFreeString(param3);
    return hr;
}
HRESULT Method_VARIANTp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, pvarResult);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, IDispatch *, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_DISPATCH,
                          };
    BSTR    param1;
    IDispatch *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if ((ulAlloc & 2) && param2)
        param2->Release();
    return hr;
}
HRESULT Method_void_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, IDispatch *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_DISPATCH,
                          };
    BSTR    param1;
    IDispatch *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if ((ulAlloc & 2) && param2)
        param2->Release();
    return hr;
}
HRESULT GS_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, BSTR *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, BSTR);

    HRESULT         hr;
    ULONG   ulAlloc = 0;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BSTR };
    BSTR    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (BSTR *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (BSTR)param1);
    }
    if (ulAlloc && param1)
        SysFreeString(param1);

    return hr;
}
HRESULT GS_PropEnum (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, BSTR *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, BSTR);

    HRESULT         hr;
    ULONG   ulAlloc = 0;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BSTR };
    BSTR    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (BSTR *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (BSTR)param1);
    }
    if (ulAlloc && param1)
        SysFreeString(param1);

    return hr;
}
HRESULT GS_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, VARIANT *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, VARIANT);

    HRESULT         hr;
    ULONG   ulAlloc = 0;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_VARIANT };
    VARIANT    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl, pInstance), pvarResult);
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (VARIANT)param1);
    }
    if (ulAlloc)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }

    return hr;
}
HRESULT GS_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, VARIANT_BOOL *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, VARIANT_BOOL);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BOOL };
    VARIANT_BOOL    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (VARIANT_BOOL *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, 0L, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (VARIANT_BOOL)param1);
    }

    return hr;
}
HRESULT GS_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, LONG *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, LONG);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_I4 };
    LONG    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (LONG *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, 0L, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (LONG)param1);
    }

    return hr;
}
HRESULT GS_float (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, float *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, float);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_R4 };
    float    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (float *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, 0L, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (float)param1);
    }

    return hr;
}
HRESULT Method_BSTRp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (BSTR *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BSTR | VT_BYREF) & ~VT_BYREF;
    return hr;
}
HRESULT G_PropEnum (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, BSTR *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BSTR };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (BSTR *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT G_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, VARIANT *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_VARIANT };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl, pInstance), pvarResult);

    return hr;
}
HRESULT G_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, BSTR *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BSTR };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (BSTR *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT G_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, LONG *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_I4 };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (LONG *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT G_IUnknownp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, IUnknown * *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_UNKNOWN };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (IUnknown * *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT Method_VARIANTp_VARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT *, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT | VT_BYREF,
                          };
    VARIANT *    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, pvarResult);

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(param1) == VT_BSTR);
        SysFreeString(V_BSTR(param1));
    }
    return hr;
}
HRESULT G_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, IDispatch * *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_DISPATCH };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (IDispatch * *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT Method_void_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT,
                          };
    VARIANT    param1;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param1) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                          };
    IDispatch *    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    return hr;
}
HRESULT Method_void_BSTR_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                          };
    BSTR    param1;
    BSTR    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT G_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, VARIANT_BOOL *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_BOOL };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (VARIANT_BOOL *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT Method_void_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance));
    return hr;
}
HRESULT Method_void_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT_BOOL);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BOOL,
                          };
    VARIANT_BOOL    param1 = (VARIANT_BOOL)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    return hr;
}
HRESULT Method_BSTRp_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, LONG, BSTR *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                                VT_I4,
                          };
    LONG    param1;
    LONG    param2;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (BSTR *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BSTR | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_IDispatchpp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;
    return hr;
}
HRESULT GS_short (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, SHORT *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, SHORT);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_I2 };
    SHORT    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (SHORT *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, 0L, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (SHORT)param1);
    }

    return hr;
}
HRESULT Method_void_IUnknownp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IUnknown *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_UNKNOWN,
                          };
    IUnknown *    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    return hr;
}
HRESULT Method_VARIANTBOOLp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;
    return hr;
}
HRESULT Method_IDispatchpp_IDispatchp_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, VARIANT, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                                VT_VARIANT,
                          };
    IDispatch *    param1;
    VARIANT    param2;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    return hr;
}
HRESULT Method_IDispatchpp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                          };
    IDispatch *    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    return hr;
}
HRESULT Method_IDispatchpp_IDispatchp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, IDispatch *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                                VT_DISPATCH,
                          };
    IDispatch *    param1;
    IDispatch *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    if ((ulAlloc & 2) && param2)
        param2->Release();
    return hr;
}
HRESULT Method_IDispatchpp_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT_BOOL, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BOOL,
                          };
    VARIANT_BOOL    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_void_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                          };
    IDispatch *    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    return hr;
}
HRESULT Method_IDispatchpp_IDispatchp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, BSTR, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                                VT_BSTR,
                          };
    IDispatch *    param1;
    BSTR    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT Method_IDispatchpp_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT_BOOL, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BOOL,
                          };
    VARIANT_BOOL    param1 = (VARIANT_BOOL)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_IDispatchpp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, IDispatch *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_DISPATCH,
                          };
    BSTR    param1;
    IDispatch *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if ((ulAlloc & 2) && param2)
        param2->Release();
    return hr;
}
HRESULT Method_BSTRp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (BSTR *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BSTR | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_BSTRp_BSTR_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, BSTR *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                          };
    BSTR    param1;
    BSTR    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (BSTR *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BSTR | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT Method_longp_BSTR_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *, LONG *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT | VT_BYREF,
                          };
    BSTR    param1;
    CVariant   param2opt(VT_ERROR);      // optional variant
    VARIANT   *param2 = &param2opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (LONG *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_I4 | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(param2) == VT_BSTR);
        SysFreeString(V_BSTR(param2));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                          };
    LONG    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_IDispatchpp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_IDispatchpp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                          };
    LONG    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_IDispatchpp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT | VT_BYREF,
                          };
    CVariant   param1opt(VT_ERROR);      // optional variant
    VARIANT   *param1 = &param1opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(param1) == VT_BSTR);
        SysFreeString(V_BSTR(param1));
    }
    return hr;
}
HRESULT Method_IDispatchpp_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT,
                          };
    VARIANT    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }
    return hr;
}
HRESULT Method_IDispatchpp_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT, VARIANT, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT,
                                VT_VARIANT,
                          };
    VARIANT    param1;    // optional arg.
    VARIANT    param2;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param1) = VT_ERROR;
    V_VT(&param2) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_oDolong_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, LONG, LONG, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_I4,
                                VT_I4,
                          };
    BSTR    param1;
    LONG    param2 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    LONG    param3 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[1]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_void_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, LONG);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                                VT_I4,
                          };
    LONG    param1;
    LONG    param2;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2);

Cleanup:
    return hr;
}
HRESULT Method_longp_BSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, LONG, LONG *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_I4,
                          };
    BSTR    param1;
    LONG    param2 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (LONG *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_I4 | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_longp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, IDispatch *, LONG *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_DISPATCH,
                          };
    BSTR    param1;
    IDispatch *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (LONG *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_I4 | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if ((ulAlloc & 2) && param2)
        param2->Release();
    return hr;
}
HRESULT Method_void_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    param1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_oDoVARIANTBOOL_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT_BOOL, VARIANT, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BOOL,
                                VT_VARIANT,
                          };
    BSTR    param1;
    VARIANT_BOOL    param2 = (VARIANT_BOOL)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    VARIANT    param3;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param3) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 4)
    {
        Assert(V_VT(&param3) == VT_BSTR);
        SysFreeString(V_BSTR(&param3));
    }
    return hr;
}
HRESULT Method_longp_BSTR_BSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, LONG, LONG *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                                VT_I4,
                          };
    BSTR    param1;
    BSTR    param2;
    LONG    param3 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (LONG *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_I4 | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT Method_void_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                          };
    LONG    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    return hr;
}
HRESULT Method_VARIANTp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                          };
    LONG    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, pvarResult);

Cleanup:
    return hr;
}
HRESULT G_short (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblPropFunc)(IDispatch *, SHORT *);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_I2 };

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    hr = (*(OLEVTblPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (SHORT *)&pvarResult->iVal);
    if (!hr)
        V_VT(pvarResult) = argTypes[0];

    return hr;
}
HRESULT Method_void_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT | VT_BYREF,
                          };
    CVariant   param1opt(VT_ERROR);      // optional variant
    VARIANT   *param1 = &param1opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(param1) == VT_BSTR);
        SysFreeString(V_BSTR(param1));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT,
                          };
    BSTR    param1;
    VARIANT    param2;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    return hr;
}
HRESULT Method_void_VARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT, VARIANT, VARIANT, VARIANT, VARIANT, VARIANT);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                          };
    VARIANT    param1;
    VARIANT    param2;    // optional arg.
    VARIANT    param3;    // optional arg.
    VARIANT    param4;    // optional arg.
    VARIANT    param5;    // optional arg.
    VARIANT    param6;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;
    V_VT(&param3) = VT_ERROR;
    V_VT(&param4) = VT_ERROR;
    V_VT(&param5) = VT_ERROR;
    V_VT(&param6) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, &param4, &param5, &param6, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, param4, param5, param6);

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(&param3) == VT_BSTR);
        SysFreeString(V_BSTR(&param3));
    }
    if (ulAlloc & 8)
    {
        Assert(V_VT(&param4) == VT_BSTR);
        SysFreeString(V_BSTR(&param4));
    }
    if (ulAlloc & 16)
    {
        Assert(V_VT(&param5) == VT_BSTR);
        SysFreeString(V_BSTR(&param5));
    }
    if (ulAlloc & 32)
    {
        Assert(V_VT(&param6) == VT_BSTR);
        SysFreeString(V_BSTR(&param6));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_BSTR_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, VARIANT, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                                VT_VARIANT,
                          };
    BSTR    param1;
    BSTR    param2;
    VARIANT    param3;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param3) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    if (ulAlloc & 4)
    {
        Assert(V_VT(&param3) == VT_BSTR);
        SysFreeString(V_BSTR(&param3));
    }
    return hr;
}
HRESULT Method_void_IDispatchp_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, IDispatch *, VARIANT);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_DISPATCH,
                                VT_VARIANT,
                          };
    IDispatch *    param1;
    VARIANT    param2;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2);

Cleanup:
    if ((ulAlloc & 1) && param1)
        param1->Release();
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    return hr;
}
HRESULT Method_void_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                          };
    LONG    param1 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]);


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    return hr;
}
HRESULT Method_void_SAFEARRAYPVARIANTP (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, SAFEARRAY *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_ARRAY,
                          };
    SAFEARRAY *    param1;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to safearray.
    param1 = DispParamsToSAFEARRAY(pdispparams);
    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);
    if (param1)
    {
        HRESULT hr1 = SafeArrayDestroy(param1);
        if (hr1)
            hr = hr1;
    }
    return hr;
}
HRESULT Method_IDispatchpp_oDoBSTR_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT, VARIANT, VARIANT, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    VARIANT    param2;    // optional arg.
    VARIANT    param3;    // optional arg.
    VARIANT    param4;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;
    V_VT(&param3) = VT_ERROR;
    V_VT(&param4) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, &param4, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, param4, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    SysFreeString(paramDef1);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(&param3) == VT_BSTR);
        SysFreeString(V_BSTR(&param3));
    }
    if (ulAlloc & 8)
    {
        Assert(V_VT(&param4) == VT_BSTR);
        SysFreeString(V_BSTR(&param4));
    }
    return hr;
}
HRESULT Method_IDispatchpp_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, LONG, LONG, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_I4,
                                VT_I4,
                          };
    LONG    param1;
    LONG    param2;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, NULL, 0, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    return hr;
}
HRESULT Method_IDispatchpp_oDoBSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, LONG, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_I4,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    LONG    param2 = (LONG)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[1]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    SysFreeString(paramDef1);
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_longp_VARIANTp_long_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT *, LONG, VARIANT *, LONG *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT | VT_BYREF,
                                VT_I4,
                                VT_VARIANT | VT_BYREF,
                          };
    VARIANT *    param1;
    LONG    param2;
    CVariant   param3opt(VT_ERROR);      // optional variant
    VARIANT   *param3 = &param3opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (LONG *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_I4 | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(param1) == VT_BSTR);
        SysFreeString(V_BSTR(param1));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(param3) == VT_BSTR);
        SysFreeString(V_BSTR(param3));
    }
    return hr;
}
HRESULT Method_void_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1);

Cleanup:
    SysFreeString(paramDef1);
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTBOOLp_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    SysFreeString(paramDef1);
    if (ulAlloc & 1)
        SysFreeString(param1);
    return hr;
}
HRESULT Method_VARIANTp_oDoBSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    BSTR    paramDef2 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[1]));
    BSTR    param2 = paramDef2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, pvarResult);

Cleanup:
    SysFreeString(paramDef1);
    SysFreeString(paramDef2);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT Method_IDispatchpp_oDoBSTR_oDoBSTR_oDoBSTR_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, BSTR, VARIANT_BOOL, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                                VT_BSTR,
                                VT_BOOL,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    BSTR    paramDef2 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[1]));
    BSTR    param2 = paramDef2;
    BSTR    paramDef3 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[2]));
    BSTR    param3 = paramDef3;
    VARIANT_BOOL    param4 = (VARIANT_BOOL)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[3]);
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, &param4, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, param4, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    SysFreeString(paramDef1);
    SysFreeString(paramDef2);
    SysFreeString(paramDef3);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    if (ulAlloc & 4)
        SysFreeString(param3);
    return hr;
}
HRESULT Method_VARIANTp_BSTR_o0oVARIANTp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *, VARIANT *, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT | VT_BYREF,
                                VT_VARIANT | VT_BYREF,
                          };
    BSTR    param1;
    CVariant   param2opt(VT_ERROR);      // optional variant
    VARIANT   *param2 = &param2opt;      // optional arg.
    CVariant   param3opt(VT_ERROR);      // optional variant
    VARIANT   *param3 = &param3opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, pvarResult);

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(param2) == VT_BSTR);
        SysFreeString(V_BSTR(param2));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(param3) == VT_BSTR);
        SysFreeString(V_BSTR(param3));
    }
    return hr;
}
HRESULT Method_void_BSTR_o0oVARIANT_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT, BSTR);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT,
                                VT_BSTR,
                          };
    BSTR    param1;
    VARIANT    param2;    // optional arg.
    BSTR    paramDef3 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param3 = paramDef3;
    ULONG   ulAlloc = 0;

    V_VT(&param2) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3);

Cleanup:
    SysFreeString(paramDef3);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    if (ulAlloc & 4)
        SysFreeString(param3);
    return hr;
}
HRESULT Method_VARIANTp_BSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, BSTR, VARIANT *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_BSTR,
                          };
    BSTR    param1;
    BSTR    paramDef2 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param2 = paramDef2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, pvarResult);

Cleanup:
    SysFreeString(paramDef2);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
        SysFreeString(param2);
    return hr;
}
HRESULT Method_IDispatchpp_oDoBSTR_o0oVARIANTp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *, VARIANT *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT | VT_BYREF,
                                VT_VARIANT | VT_BYREF,
                          };
    BSTR    paramDef1 = SysAllocString((TCHAR *)(((PROPERTYDESC_METHOD *)pDesc)->c->defargs[0]));
    BSTR    param1 = paramDef1;
    CVariant   param2opt(VT_ERROR);      // optional variant
    VARIANT   *param2 = &param2opt;      // optional arg.
    CVariant   param3opt(VT_ERROR);      // optional variant
    VARIANT   *param3 = &param3opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    SysFreeString(paramDef1);
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(param2) == VT_BSTR);
        SysFreeString(V_BSTR(param2));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(param3) == VT_BSTR);
        SysFreeString(V_BSTR(param3));
    }
    return hr;
}
HRESULT GS_LONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, LONG *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, LONG);

    HRESULT         hr;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_I4 };
    LONG    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (LONG *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, 0L, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (LONG)param1);
    }

    return hr;
}
HRESULT Method_IDispatchpp_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, VARIANT, VARIANT, VARIANT, VARIANT, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                                VT_VARIANT,
                          };
    VARIANT    param1;    // optional arg.
    VARIANT    param2;    // optional arg.
    VARIANT    param3;    // optional arg.
    VARIANT    param4;    // optional arg.
    ULONG   ulAlloc = 0;

    V_VT(&param1) = VT_ERROR;
    V_VT(&param2) = VT_ERROR;
    V_VT(&param3) = VT_ERROR;
    V_VT(&param4) = VT_ERROR;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, &param3, &param4, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, param3, param4, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
    {
        Assert(V_VT(&param1) == VT_BSTR);
        SysFreeString(V_BSTR(&param1));
    }
    if (ulAlloc & 2)
    {
        Assert(V_VT(&param2) == VT_BSTR);
        SysFreeString(V_BSTR(&param2));
    }
    if (ulAlloc & 4)
    {
        Assert(V_VT(&param3) == VT_BSTR);
        SysFreeString(V_BSTR(&param3));
    }
    if (ulAlloc & 8)
    {
        Assert(V_VT(&param4) == VT_BSTR);
        SysFreeString(V_BSTR(&param4));
    }
    return hr;
}
HRESULT GS_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblGetPropFunc)(IDispatch *, IDispatch * *);
    typedef HRESULT (STDMETHODCALLTYPE *OLEVTblSetPropFunc)(IDispatch *, IDispatch *);

    HRESULT         hr;
    ULONG   ulAlloc = 0;
    VTABLE_ENTRY *pVTbl;
    VARTYPE         argTypes[] = { VT_DISPATCH };
    IDispatch *    param1;

    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY *)(((BYTE *)(*(DWORD_PTR *)pInstance)) + (wVTblOffset*sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET));

    if (wFlags & INVOKE_PROPERTYGET)
    {
        hr = (*(OLEVTblGetPropFunc)VTBL_PFN(pVTbl))((IDispatch*) VTBL_THIS(pVTbl,pInstance), (IDispatch * *)&pvarResult->iVal);
        if (!hr)
            V_VT(pvarResult) = argTypes[0];
    }
    else
    {
         // Convert dispatch params to C params.
         hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, 0, argTypes, &param1, -1);
         if (hr)
             pBase->SetErrorInfo(hr);
         else
             hr = (*(OLEVTblSetPropFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), (IDispatch *)param1);
    }
    if (ulAlloc && param1)
        param1->Release();

    return hr;
}
HRESULT Method_IDispatchpp_BSTR_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *, IDispatch **);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT | VT_BYREF,
                          };
    BSTR    param1;
    CVariant   param2opt(VT_ERROR);      // optional variant
    VARIANT   *param2 = &param2opt;      // optional arg.
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (IDispatch **)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_DISPATCH | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(param2) == VT_BSTR);
        SysFreeString(V_BSTR(param2));
    }
    return hr;
}
HRESULT Method_VARIANTBOOLp_BSTR_VARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult)
{
    typedef HRESULT (STDMETHODCALLTYPE *MethodOLEVTblFunc)(IDispatch *, BSTR, VARIANT *, VARIANT_BOOL *);

    HRESULT               hr;
    VTABLE_ENTRY     *pVTbl;
    VARTYPE               argTypes[] = {
                                VT_BSTR,
                                VT_VARIANT | VT_BYREF,
                          };
    BSTR    param1;
    VARIANT *    param2;
    ULONG   ulAlloc = 0;


    Assert(pInstance || pDesc || pdispparams);

    pVTbl = (VTABLE_ENTRY*)(((BYTE *)(*(DWORD_PTR *)pInstance)) + wVTblOffset * sizeof(VTABLE_ENTRY)/sizeof(DWORD_PTR) + FIRST_VTABLE_OFFSET);

    // Convert dispatch params to C params.
    hr = DispParamsToCParams(pSrvProvider, pdispparams, &ulAlloc, ((PROPERTYDESC_BASIC_ABSTRACT *)pDesc)->b.wMaxstrlen, argTypes, &param1, &param2, -1L);
    if (hr)
    {
        pBase->SetErrorInfo(hr);
        goto Cleanup;
    }

    hr = (*(MethodOLEVTblFunc)VTBL_PFN(pVTbl))((IDispatch*)VTBL_THIS(pVTbl,pInstance), param1, param2, (VARIANT_BOOL *)(&(pvarResult->pdispVal)));
    if (!hr)
        V_VT(pvarResult) = (VT_BOOL | VT_BYREF) & ~VT_BYREF;

Cleanup:
    if (ulAlloc & 1)
        SysFreeString(param1);
    if (ulAlloc & 2)
    {
        Assert(V_VT(param2) == VT_BSTR);
        SysFreeString(V_BSTR(param2));
    }
    return hr;
}

EXTERN_C const IID IID_IHTMLStyle2;
EXTERN_C const IID IID_IHTMLElement;
EXTERN_C const IID IID_IHTMLElement2;
EXTERN_C const IID IID_IHTMLFormElement2;
EXTERN_C const IID IID_IHTMLElementCollection2;
EXTERN_C const IID IID_IHTMLTextRangeMetrics;
EXTERN_C const IID IID_IHTMLTextRangeMetrics2;
EXTERN_C const IID IID_IHTMLControlElement;
EXTERN_C const IID IID_IHTMLInputTextElement;
EXTERN_C const IID IID_IHTMLInputHiddenElement;
EXTERN_C const IID IID_IHTMLInputButtonElement;
EXTERN_C const IID IID_IHTMLInputImage;
EXTERN_C const IID IID_IHTMLTextContainer;
EXTERN_C const IID IID_IHTMLBlockElement;
EXTERN_C const IID IID_IHTMLAreasCollection2;
EXTERN_C const IID IID_IHTMLListElement;
EXTERN_C const IID IID_IHTMLDocument3;
EXTERN_C const IID IID_IHTMLWindow3;
EXTERN_C const IID IID_IHTMLPhraseElement;
EXTERN_C const IID IID_IHTMLSelectElement2;
EXTERN_C const IID IID_IHTMLObjectElement2;


#define MAX_IIDS 31
static const IID * _IIDTable[MAX_IIDS] = {
	NULL,
	&IID_IHTMLStyle2,
	&IID_IHTMLElement,
	&IID_IHTMLElement2,
	&IID_IHTMLFormElement2,
	&IID_IHTMLElementCollection2,
	&IID_IHTMLTextRangeMetrics,
	&IID_IHTMLTextRangeMetrics2,
	&IID_IHTMLControlElement,
	&IID_IHTMLInputTextElement,
	&IID_IHTMLInputHiddenElement,
	&IID_IHTMLInputButtonElement,
	&IID_IHTMLInputImage,
	&IID_IHTMLTextContainer,
	&IID_IHTMLBlockElement,
	&IID_IHTMLAreasCollection2,
	&IID_IHTMLListElement,
	&IID_IHTMLDocument3,
	&IID_IHTMLWindow3,
	&IID_IHTMLPhraseElement,
	&IID_IHTMLSelectElement2,
	&IID_IHTMLObjectElement2,
};

#define DIID_DispBase   0x3050f500
#define DIID_DispMax    0x3050f5a0

const GUID DIID_Low12Bytes = { 0x00000000, 0x98b5, 0x11cf, { 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };


BOOL DispNonDualDIID(IID iid)
{
	BOOL    fRetVal = FALSE;

	if (iid.Data1 >= DIID_DispBase && iid.Data1 <= DIID_DispMax)
	{
		fRetVal = memcmp(&iid.Data2,
				&DIID_Low12Bytes.Data2,
				sizeof(IID) - sizeof(DWORD)) == 0;
	}

	return fRetVal;
}

