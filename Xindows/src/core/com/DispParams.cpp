
#include "stdafx.h"
#include "DispParams.h"

//+---------------------------------------------------------------
//
//  Member:     CDispParams::Create, public
//
//  Synopsis:   Allocated argument and named argument arrays.
//              Initial values are VT_NULL for argument array
//              and DISPID_UNKNOWN for named argument array.
//
//----------------------------------------------------------------
HRESULT CDispParams::Create(DISPPARAMS* pOrgDispParams)
{
    HRESULT hr = S_OK;
    UINT    i;

    // Nothing should exist yet.
    if(rgvarg || rgdispidNamedArgs)
    {
        hr = E_FAIL;
        goto Error;
    }

    if(cArgs+cNamedArgs)
    {
        rgvarg = new VARIANTARG[cArgs+cNamedArgs];
        if(!rgvarg)
        {
            hr = E_OUTOFMEMORY;
            goto Error;
        }

        // cArgs is now total count of args including named args.
        cArgs = cArgs + cNamedArgs;

        // Initialize all parameters to VT_NULL.
        for(i=0; i<cArgs; i++)
        {
            rgvarg[i].vt = VT_NULL;
        }

        // Any arguments to copy over?
        if(pOrgDispParams->cArgs)
        {
            if(cArgs >= pOrgDispParams->cArgs)
            {
                UINT iStartIndex;

                iStartIndex = cArgs - pOrgDispParams->cArgs;

                if(cArgs >= iStartIndex+pOrgDispParams->cArgs)
                {
                    memcpy(&rgvarg[iStartIndex], pOrgDispParams->rgvarg,
                        sizeof(VARIANTARG)*pOrgDispParams->cArgs);
                }
            }
            else
            {
                hr = E_UNEXPECTED;
                goto Cleanup;
            }
        }

        if(cNamedArgs)
        {
            rgdispidNamedArgs = new DISPID[cNamedArgs];
            if(!rgdispidNamedArgs)
            {
                hr = E_OUTOFMEMORY;
                goto Error;
            }

            // Initialize all named args to the unknown dispid.
            for(i=0; i<cNamedArgs; i++)
            {
                rgdispidNamedArgs[i] = DISPID_UNKNOWN;
            }

            if(pOrgDispParams->cNamedArgs)
            {
                if(cNamedArgs >= pOrgDispParams->cNamedArgs)
                {
                    UINT iStartIndex;

                    iStartIndex = cNamedArgs - pOrgDispParams->cNamedArgs;
                    if(cNamedArgs >= (iStartIndex+pOrgDispParams->cNamedArgs))
                    {
                        memcpy(&rgdispidNamedArgs[iStartIndex], pOrgDispParams->rgdispidNamedArgs,
                            sizeof(VARIANTARG)*pOrgDispParams->cNamedArgs);
                    }
                }
                else
                {
                    hr = E_UNEXPECTED;
                    goto Cleanup;
                }
            }
        }
    }

Cleanup:
    RRETURN(hr);

Error:
    delete[] rgvarg;
    delete[] rgdispidNamedArgs;
    goto Cleanup;
}

//+---------------------------------------------------------------
//
//  Member:     CDispParams::MoveArgsToDispParams, public
//
//  Synopsis:   Move arguments from arguments array to pOutDispParams.
//              Notice, I said move not copy so both this
//              object and the pOutDispParams hold the identical
//              VARIANTS.  So be careful to only release these
//              variants ONCE.  The fFromEnd parameter specifies
//              how the arguments are moved from our rgvar array.
//
//----------------------------------------------------------------
HRESULT CDispParams::MoveArgsToDispParams(DISPPARAMS* pOutDispParams, UINT cNumArgs, BOOL fFromEnd/*=TRUE*/)
{
    HRESULT hr = S_OK;
    UINT    iStartIndex;

    if(rgvarg && cNumArgs)
    {
        if(cArgs>=cNumArgs && pOutDispParams->cArgs>=cNumArgs)
        {
            iStartIndex = fFromEnd ? cArgs-cNumArgs : 0;

            memcpy(pOutDispParams->rgvarg, &rgvarg[iStartIndex],
                sizeof(VARIANTARG)*cNumArgs);
            goto Cleanup;
        }
        hr = E_FAIL;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDispParams::ReleaseVariants, public
//
//  Synopsis:   Release any variants in out arguments array.
//              Again, if the values were moved (exist in 2
//              places) be careful you may have just screwed
//              yourself.
//
//----------------------------------------------------------------
void CDispParams::ReleaseVariants()
{
    for(UINT i=0; i<cArgs; i++)
    {
        VariantClear(rgvarg+i);
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   VARTYPEFromBYTE
//
//  Synopsis:   Converts a byte specification of the type of a variant to
//              a VARTYPE.
//
//----------------------------------------------------------------------------
VARTYPE VARTYPEFromBYTE(BYTE b)
{
    VARTYPE vt;

    Assert(!(b & 0xB0));
    if(b & 0x40)
    {
        vt = (VARTYPE)((UINT)b ^ (0x40|VT_BYREF));
    }
    else
    {
        vt = b;
    }

    return vt;
}

//+---------------------------------------------------------------------------
//
//  Function:   CParamsToDispParams
//
//  Synopsis:   Converts a C parameter list to a dispatch parameter list.
//
//  Arguments:  [pDispParams] -- Resulting dispatch parameter list.
//                               Note that the rgvarg member of pDispParams
//                               must be initialized with an array of
//                               EVENTPARAMS_MAX VARIANTs.
//
//              [pb]         -- List of C parameter types.  May be NULL.
//                              Construct using EVENT_PARAM macro.
//
//              [va]          -- List of C arguments.
//
//  Modifies:   [pDispParams]
//
//  History:    05-Jan-94   adams   Created
//              23-Feb-94   adams   Reversed order of disp arguments, added
//                                  support for VT_R4, VT_R8, and pointer
//                                  types.
//
//----------------------------------------------------------------------------
void CParamsToDispParams(DISPPARAMS* pDispParams, BYTE* pb, va_list va)
{
    Assert(pDispParams);
    Assert(pDispParams->rgvarg);

    VARIANTARG* pvargCur;   // current variant
    BYTE*       pbCur;      // current vartype

    // Assign vals to dispatch param list.
    pDispParams->cNamedArgs         = 0;
    pDispParams->rgdispidNamedArgs  = NULL;

    pDispParams->cArgs = strlen((char*)pb);
    Assert(pDispParams->cArgs < EVENTPARAMS_MAX);

    // Convert each C-param to a dispparam.  Note that the order of dispatch
    // parameters is the reverse of the order of c-params.
    Assert(pDispParams->rgvarg);
    pvargCur = pDispParams->rgvarg + pDispParams->cArgs;
    for(pbCur=pb; *pbCur; pbCur++)
    {
        pvargCur--;
        Assert(pvargCur >= pDispParams->rgvarg);

        V_VT(pvargCur) = VARTYPEFromBYTE(*pbCur);
        if(V_VT(pvargCur) & VT_BYREF)
        {
            V_BYREF(pvargCur) = va_arg(va, long*);
        }
        else
        {
            switch(V_VT(pvargCur))
            {
            case VT_BOOL:
                // convert TRUE to VT_TRUE
                V_BOOL(pvargCur) = VARIANT_BOOL(-va_arg(va, BOOL));
                Assert(V_BOOL(pvargCur)==VB_FALSE || V_BOOL(pvargCur)==VB_TRUE);
                break;

            case VT_I2:
                V_I2(pvargCur) = va_arg(va, short);
                break;

            case VT_ERROR:
            case VT_I4:
                V_I4(pvargCur) = va_arg(va, long);
                break;

            case VT_R4:
                V_R4(pvargCur) = (float) va_arg(va, double);
                // casting & change to double inserted to fix BUG 5005
                break;

            case VT_R8:
                V_R8(pvargCur) = va_arg(va, double);
                break;

                // All Pointer types.
            case VT_PTR:
            case VT_BSTR:
            case VT_LPSTR:
            case VT_LPWSTR:
            case VT_DISPATCH:
            case VT_UNKNOWN:
                V_BYREF(pvargCur) = va_arg(va, void**);
                break;

            case VT_VARIANT:
                *pvargCur = va_arg(va, VARIANT);
                break;

            default:
                Assert(FALSE && "Unknown type.\n");
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   DispParamsToCParams
//
//  Synopsis:   Converts Dispatch::Invoke method params to C-language params.
//
//  Arguments:  [pDP] -- Dispatch params to be converted.
//              [pb]  -- Array of types of C-params.  May be NULL.
//                       Construct using EVENT_PARAM macro.
//              [...] -- List of pointers to c-params to be converted to.
//              -1    -- Last parameter to signal end of parameter list.
//
//  Returns:    HRESULT.
//
//  History:    2-23-94   adams   Created
//
//  Notes:      Supports types listed in VARIANTToCParam.
//
//----------------------------------------------------------------------------
HRESULT DispParamsToCParams(IServiceProvider* pSrvProvider, DISPPARAMS* pDP, ULONG* pAlloc, WORD wMaxstrlen, VARTYPE* pVT, ...)
{
    HRESULT     hr;
    va_list     va;         // list of pointers to c-params.
    VARIANTARG* pvargCur;   // current VARIANT being converted.
    void*       pv;         // current c-param being converted.
    UINT        cArgs;      // count of arguments.

    Assert(pDP);

    hr = S_OK;
    va_start(va, pVT);
    if(!pVT)
    {
        if(pDP->cArgs > 0)
        {
            goto BadParamCountError;
        }

        goto Cleanup;
    }

    pvargCur = pDP->rgvarg + pDP->cArgs - 1;
    for(cArgs=0; cArgs<pDP->cArgs; cArgs++)
    {
        BOOL fAlloc;

        // If the DISPID_THIS named argument is passed in skip it.
        if(pDP->cNamedArgs && (pDP->cArgs-cArgs<=pDP->cNamedArgs))
        {
            if(pDP->rgdispidNamedArgs[(pDP->cArgs-cArgs)-1] == DISPID_THIS)
            {
                pvargCur--;
                continue;
            }
        }

        pv = va_arg(va, void*);

        // Done processing arguments?
        if(pv == (void*)-1)
        {
            goto Cleanup;
        }

        // Skip all byvalue variants custom invoke doesn't pass them.
        if(!((*pVT==VT_VARIANT) && (pv==NULL)))
        {
            hr = VARIANTARGToCVar(pvargCur, &fAlloc, *pVT, pv, pSrvProvider, ((wMaxstrlen==pdlNoLimit)?0:wMaxstrlen));
            if(hr)
            {
                goto Cleanup;
            }

            // Any BSTRs or objects (IUnknow, IDispatch) allocated during
            // conversion to CVar then remember which param this occurred to so
            // we can de-allocate it when we're finished.
            if(pAlloc && fAlloc)
            {
                *pAlloc |= (1 << cArgs);
            }
        }

        pvargCur--;
        pVT++;
    }

Cleanup:
    va_end(va);
    RRETURN(hr);

BadParamCountError:
    hr = DISP_E_BADPARAMCOUNT;
    goto Cleanup;
}

//+------------------------------------------------------------------------
//  Function:   DispParamsToSAFEARRAY, public API
//
//  Synopsis:   Converts all arguments in dispparams to a SAFEARRAY
//              If the DISPPARAMS contains no arguments we should create
//              an empty SAFEARRAY.
//
//  Arguments:  [pdispparams] -- VARIANTARGs to add to safearray.
//
//  Returns:    If the DISPPARAMS contains no arguments we should create an
//              empty SAFEARRAY.  It is the responsibility of the caller to
//              call SafeArrayDestroy.
//-------------------------------------------------------------------------
SAFEARRAY* DispParamsToSAFEARRAY(DISPPARAMS* pdispparams)
{
    SAFEARRAY*  psa = NULL;
    HRESULT     hr = S_OK;

    LONG saElemIdx;
    SAFEARRAYBOUND sabounds;
    const LONG cArgsToArray = pdispparams->cArgs;
    const LONG cArgsNamed = pdispparams->cNamedArgs;

    sabounds.cElements = cArgsToArray;

    // If first named arg is DISPID_THIS then this parameter won't be part of
    // the safearray.
    if(cArgsNamed)
    {
        if(pdispparams->rgdispidNamedArgs[0] == DISPID_THIS)
        {
            sabounds.cElements--;
        }
    }

    sabounds.lLbound = 0;
    psa = SafeArrayCreate(VT_VARIANT, 1, &sabounds);
    if(psa == NULL)
    {
        goto Cleanup;
    }

    // dispparams are in right to left order.
    for(saElemIdx=0; saElemIdx<cArgsToArray; saElemIdx++)
    {
        // Don't process any DISPID_THIS named arguments.
        if(cArgsNamed && (cArgsToArray-saElemIdx<=cArgsNamed))
        {
            if(pdispparams->rgdispidNamedArgs[(cArgsToArray-saElemIdx)-1] == DISPID_THIS)
            {
                continue;
            }
        }

        hr = SafeArrayPutElement(psa, &saElemIdx, pdispparams->rgvarg+(cArgsToArray-1-saElemIdx));
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    if(hr && psa)
    {
        hr = SafeArrayDestroy(psa);
        psa = NULL;
    }

    return psa;
}