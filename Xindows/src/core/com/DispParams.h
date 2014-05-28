
#ifndef __XINDOWS_CORE_COM_DISPPARAMS_H__
#define __XINDOWS_CORE_COM_DISPPARAMS_H__

//+---------------------------------------------------------------
//
//  Class:      CDispParams
//
//  Synopsis:   Helper class to manipulate the DISPPARAMS structure
//              used by Invoke/InvokeEx.
//
//----------------------------------------------------------------
class CDispParams : public DISPPARAMS
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CDispParams(UINT cTotalArgs=0, UINT cTotalNamedArgs=0)
    {
        rgvarg = NULL;
        rgdispidNamedArgs = NULL;
        cArgs = cTotalArgs;
        cNamedArgs = cTotalNamedArgs;
    }
    ~CDispParams()
    {
        delete[] rgvarg;
        delete[] rgdispidNamedArgs;
    }

    void Initialize(UINT cTotalArgs, UINT cTotalNamedArgs)
    {
        Assert(!rgvarg && !rgdispidNamedArgs);
        cArgs = cTotalArgs;
        cNamedArgs = cTotalNamedArgs;
    }

    HRESULT Create(DISPPARAMS* pOrgDispParams);

    HRESULT MoveArgsToDispParams(DISPPARAMS* pOutDispParams, UINT cNumArgs, BOOL fFromEnd=TRUE);

    void ReleaseVariants();
};

void CParamsToDispParams(DISPPARAMS* pDispParams, BYTE* pb, va_list va);
HRESULT CDECL DispParamsToCParams(IServiceProvider* pSrvProvider, DISPPARAMS* pDP, ULONG* pAlloc, WORD wMaxstrlen, VARTYPE* pVT, ...);

SAFEARRAY* DispParamsToSAFEARRAY(DISPPARAMS* pdispparams);

#endif //__XINDOWS_CORE_COM_DISPPARAMS_H__