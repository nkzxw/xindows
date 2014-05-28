
#include "stdafx.h"
#include "Variant.h"

#include <DispEx.h>

#define LCID_SCRIPTING  0x0409  // Mandatory VariantChangeTypeEx localeID

// declare in DispEx.h
// {1F101481-BCCD-11d0-9336-00A0C90DCAA9}
const GUID SID_VariantConversion = { 0x1f101481, 0xbccd, 0x11d0, { 0x93, 0x36, 0x0, 0xa0, 0xc9, 0xd, 0xca, 0xa9 } };

// Coerce pArgFrom into this instance from anyvariant to a given type
HRESULT CVariant::CoerceVariantArg(VARIANT* pArgFrom, WORD wCoerceToType)
{
    HRESULT hr = S_OK;
    VARIANT* pvar;

    if(V_VT(pArgFrom) == (VT_BYREF|VT_VARIANT))
    {
        pvar = V_VARIANTREF(pArgFrom);
    }
    else
    {
        pvar = pArgFrom;
    }

    if(!(pvar->vt==VT_EMPTY || pvar->vt==VT_ERROR))
    {
        hr = VariantChangeTypeSpecial((VARIANT*)this, pvar,  wCoerceToType);
    }
    else
    {
        return S_FALSE;
    }
    RRETURN(hr);
}

// Coerce current variant into itself
HRESULT CVariant::CoerceVariantArg(WORD wCoerceToType)
{
    HRESULT hr = S_OK;

    if(!(vt==VT_EMPTY || vt==VT_ERROR))
    {
        hr = VariantChangeTypeSpecial((VARIANT*)this, (VARIANT*)this, wCoerceToType);
    }
    else
    {
        return S_FALSE;
    }
    RRETURN(hr);
}

// Coerce any numeric (VT_I* or  VT_UI*) into a VT_I4 in this instance
BOOL CVariant::CoerceNumericToI4()
{
    switch(vt)
    {
    case VT_I1:
    case VT_UI1:
        lVal = 0x000000FF & (DWORD)bVal;
        break;

    case VT_UI2:
    case VT_I2:
        lVal = 0x0000FFFF & (DWORD)iVal;
        break;

    case VT_UI4:
    case VT_I4:
    case VT_INT: 
    case VT_UINT:
        break;

    case VT_R8:
        lVal = (LONG)dblVal;
        break;

    case VT_R4:
        lVal = (LONG)fltVal;
        break;

    default:
        return FALSE;
    }

    vt = VT_I4;
    return TRUE;
}

// returns true for non-infinite nans.
int isNAN(double dbl)
{
    union
    {
        USHORT rgw[4];
        ULONG  rglu[2];
        double dbl;
    } v;

    v.dbl = dbl;

    return (0==(~v.rgw[3]&0x7FF0) && ((v.rgw[3]&0x000F) || v.rgw[2] || v.rglu[0]));
}

// returns false for infinities and nans.
int isFinite(double dbl)
{
    union
    {
        USHORT rgw[4];
        ULONG rglu[2];
        double dbl;
    } v;

    v.dbl = dbl;

    return (~v.rgw[3]&0x7FE0) || 0==(v.rgw[3]&0x0010) && (~v.rglu[0]||~v.rgw[2]||(~v.rgw[3]&0x000F));
}

HRESULT VariantChangeTypeSpecial(VARIANT* pvargDest, VARIANT* pVArg, VARTYPE vt, IServiceProvider* pSrvProvider/*=NULL*/, DWORD dwFlags/*=0*/)
{
    HRESULT hr;
    IVariantChangeType* pVarChangeType = NULL;

    if(pSrvProvider)
    {
        hr = pSrvProvider->QueryService(SID_VariantConversion,
            IID_IVariantChangeType, (void**)&pVarChangeType);
        if(hr)
        {
            goto OldWay;
        }

        // Use script engine conversion routine.
        hr = pVarChangeType->ChangeType(pvargDest, pVArg, 0, vt);

        if(!hr)
        {
            goto Cleanup; // ChangeType suceeded we're done...
        }
    }

    // Fall back to our tried & trusted type coercions
OldWay:
    hr = S_OK;

    if(vt==VT_BSTR && V_VT(pVArg)==VT_NULL)
    {
        // Converting a NULL to BSTR
        V_VT(pvargDest) = VT_BSTR;
        hr = FormsAllocString(_T("null"), &V_BSTR(pvargDest));
        goto Cleanup;
    }
    else if(vt==VT_BSTR && V_VT(pVArg)==VT_EMPTY)
    {
        // Converting "undefined" to BSTR
        V_VT(pvargDest) = VT_BSTR;
        hr = FormsAllocString(_T("undefined"), &V_BSTR(pvargDest));
        goto Cleanup;
    }
    else if(vt==VT_BOOL && V_VT(pVArg)==VT_BSTR)
    {
        // Converting from BSTR to BOOL
        // To match Navigator compatibility empty strings implies false when
        // assigned to a boolean type any other string implies true.
        V_VT(pvargDest) = VT_BOOL;
        V_BOOL(pvargDest) = FormsStringLen(V_BSTR(pVArg))==0 ? VARIANT_FALSE : VARIANT_TRUE;
        goto Cleanup;
    }
    else if(V_VT(pVArg)==VT_BOOL && vt==VT_BSTR)
    {
        // Converting from BOOL to BSTR
        // To match Nav we either get "true" or "false"
        V_VT(pvargDest) = VT_BSTR;
        hr = FormsAllocString(
            V_BOOL(pVArg)==VARIANT_TRUE?_T("true"):_T("false"),
            &V_BSTR(pvargDest));
        goto Cleanup;
    }
    // If we're converting R4 or R8 to a string then we need special handling to
    // map Nan and +/-Inf.
    else if(vt==VT_BSTR && (V_VT(pVArg)==VT_R8||V_VT(pVArg)==VT_R4))
    {
        double dblValue = V_VT(pVArg)==VT_R8 ? V_R8(pVArg) : (double)(V_R4(pVArg));

        // Infinity or NAN?
        if(!isFinite(dblValue))
        {
            if(isNAN(dblValue))
            {
                // NAN
                hr = FormsAllocStringW(_T("NaN"), &(V_BSTR(pvargDest)));
            }
            else
            {
                // Infinity
                hr = FormsAllocStringW((dblValue<0)?_T("-Infinity"):_T("Infinity"), &(V_BSTR(pvargDest)));
            }
        }
        else
        {
            goto DefaultConvert;
        }

        // Any error from allocating string?
        if(hr)
        {
            goto Cleanup;
        }

        V_VT(pvargDest) = vt;
        goto Cleanup;
    }

DefaultConvert:
    // Default VariantChangeTypeEx.

    // VARIANT_NOUSEROVERRIDE flag is undocumented flag that tells OLEAUT to convert to the lcid
    // given. Without it the conversion is done to user localeid
    hr = VariantChangeTypeEx(pvargDest, pVArg, LCID_SCRIPTING, dwFlags|VARIANT_NOUSEROVERRIDE, vt);

    if(hr == DISP_E_TYPEMISMATCH)
    {
        if(V_VT(pVArg) == VT_NULL)
        {
            hr = S_OK;
            switch(vt)
            {
            case VT_BOOL:
                V_BOOL(pvargDest) = VARIANT_FALSE;
                V_VT(pvargDest) = VT_BOOL;
                break;

                // For NS compatability - NS treats NULL args as 0
            default:
                V_I4(pvargDest)=0;
                break;
            }
        }
        else if(V_VT(pVArg) == VT_DISPATCH)
        {
            // Nav compatability - return the string [object] or null 
            V_VT(pvargDest) = VT_BSTR;
            hr = FormsAllocString((V_DISPATCH(pVArg))?_T("[object]"):_T("null"), &V_BSTR(pvargDest));
        }
        else if(V_VT(pVArg)==VT_BSTR &&
            (V_BSTR(pVArg) && ((V_BSTR(pVArg))[0]==_T('\0')) ||!V_BSTR(pVArg)) &&
            (vt==VT_I4 || vt==VT_I2 || vt==VT_UI2||vt==VT_UI4 
            || vt==VT_I8||vt==VT_UI8 || vt==VT_INT || vt==VT_UINT))
        {
            // Converting empty string to integer => Zero
            hr = S_OK;
            V_VT(pvargDest) = vt;
            V_I4(pvargDest) = 0;
            goto Cleanup;
        }
    }
    else if(hr==DISP_E_OVERFLOW && vt==VT_I4 && (V_VT(pVArg)==VT_R8 || V_VT(pVArg)==VT_R4))
    {
        // Nav compatability - return MAXLONG on overflow
        V_VT(pvargDest) = VT_I4;
        V_I4(pvargDest) = MAXLONG;
        hr = S_OK;
        goto Cleanup;
    }

    // To match Navigator change any scientific notation E to e.
    if(!hr && (vt==VT_BSTR && (V_VT(pVArg)==VT_R8 || V_VT(pVArg)==VT_R4)))
    {
        TCHAR* pENotation;

        pENotation = _tcschr(V_BSTR(pvargDest), _T('E'));
        if(pENotation)
        {
            *pENotation = _T('e');
        }
    }

Cleanup:
    ReleaseInterface(pVarChangeType);

    RRETURN(hr);
}

HRESULT ClipVarString(VARIANT* pvarSrc, VARIANT* pvarDest, BOOL* pfAlloc, WORD wMaxstrlen)
{
    HRESULT hr = S_OK;
    if(wMaxstrlen && (V_VT(pvarSrc)==VT_BSTR) && FormsStringLen(V_BSTR(pvarSrc))>wMaxstrlen)
    {
        hr = FormsAllocStringLen(V_BSTR(pvarSrc), wMaxstrlen, &V_BSTR(pvarDest));
        if(hr)
        {
            goto Cleanup;
        }

        *pfAlloc = TRUE;
        V_VT(pvarDest) = VT_BSTR;
    }
    else
    {
        hr = S_FALSE;
    }

Cleanup:
    return hr;
}

//+---------------------------------------------------------------------------
//
//  Function:   VARIANTARGToCVar
//
//  Synopsis:   Converts a VARIANT to a C-language variable.
//
//  Arguments:  [pvarg]         -- Variant to convert.
//              [pfAlloc]       -- BSTR allocated during conversion caller is
//                                 now owner of this BSTR or IUnknown or IDispatch
//                                 object allocated needs to be released.
//              [vt]            -- Type to convert to.
//              [pv]            -- Location to place C-language variable.
//
//  Modifies:   [pv].
//
//  Returns:    HRESULT.
//
//  History:    2-23-94   adams   Created
//
//  Notes:      Supports all variant pointer types, VT_I2, VT_I4, VT_R4,
//              VT_R8, VT_ERROR.
//----------------------------------------------------------------------------
HRESULT VARIANTARGToCVar(VARIANT* pvarg, BOOL* pfAlloc, VARTYPE vt, void* pv, IServiceProvider* pSrvProvider, WORD wMaxstrlen)
{
    HRESULT     hr = S_OK;
    VARIANTARG* pVArgCopy = pvarg;
    VARIANTARG  vargNew; // variant of new type
    BOOL        fAlloc;

    Assert(pvarg);
    Assert(pv);

    if(!pfAlloc)
    {
        pfAlloc = &fAlloc;
    }

    Assert((vt&~VT_TYPEMASK) == 0 || (vt&~VT_TYPEMASK)==VT_BYREF);

    // Assume no allocations yet.
    *pfAlloc = FALSE;

    if(vt & VT_BYREF)
    {
        // If the parameter is a variant pointer then everything is acceptable.
        if((vt&VT_TYPEMASK) == VT_VARIANT)
        {
            switch(V_VT(pvarg))
            {
            case VT_VARIANT|VT_BYREF:
                hr = ClipVarString(pvarg->pvarVal, *(VARIANT**)pv, pfAlloc, wMaxstrlen);
                break;
            default:
                hr = ClipVarString(pvarg, *(VARIANT**)pv, pfAlloc, wMaxstrlen);
                break;
            }
            if(hr == S_FALSE)
            {
                hr = S_OK;
                *(PVOID*)pv = (PVOID)pvarg;
            }

            goto Cleanup;
        }

        if((V_VT(pvarg)&VT_TYPEMASK) != (vt&VT_TYPEMASK))
        {
            hr = DISP_E_TYPEMISMATCH;
            goto Cleanup;
        }

        // Type of both original and destination or same type (however, original
        // may not be a byref only the original.
        if(V_ISBYREF(pvarg))
        {
            // Destination and original are byref and same type just copy pointer.
            *(PVOID*)pv = V_BYREF(pvarg);
        }
        else
        {
            // Convert original to byref.
            switch(vt & VT_TYPEMASK)
            {
            case VT_BOOL:
                *(PVOID*)pv = (PVOID)&V_BOOL(pvarg);
                break;

            case VT_I2:
                *(PVOID*)pv = (PVOID)&V_I2(pvarg);
                break;

            case VT_ERROR:
            case VT_I4:
                *(PVOID*)pv = (PVOID)&V_I4(pvarg);
                break;

            case VT_R4:
                *(PVOID*)pv = (PVOID)&V_R4(pvarg);
                break;

            case VT_R8:
                *(PVOID*)pv = (PVOID)&V_R8(pvarg);
                break;

            case VT_CY:
                *(PVOID*)pv = (PVOID)&V_CY(pvarg);
                break;

                // All pointer types.
            case VT_PTR:
            case VT_BSTR:
            case VT_LPSTR:
            case VT_LPWSTR:
            case VT_DISPATCH:
            case VT_UNKNOWN:
                *(PVOID*)pv = (PVOID)&V_UNKNOWN(pvarg);
                break;

            case VT_VARIANT:
                Assert("Dead code: shudn't have gotten here!");
                *(PVOID*)pv = (PVOID)pvarg;
                break;

            default:
                Assert(!"Unknown type in BYREF VARIANTARGToCVar().\n");
                hr = DISP_E_TYPEMISMATCH;
                goto Cleanup;
            }
        }

        goto Cleanup;
    }
    // If the c style parameter is the same type as the VARIANT then we'll just
    // move the data.  Also if the c style type is a VARIANT then there's
    // nothing to convert just copy the variant to the C parameter.
    else if((V_VT(pvarg)&(VT_TYPEMASK|VT_BYREF))!=vt && (vt!=VT_VARIANT))
    {
        // If the request type isn't the same as the variant passed in then we
        // need to convert.
        VariantInit(&vargNew);
        pVArgCopy = &vargNew;

        hr = VariantChangeTypeSpecial(pVArgCopy, pvarg, vt,pSrvProvider);

        if(hr)
        {
            goto Cleanup;
        }

        *pfAlloc = (vt==VT_BSTR) || (vt==VT_UNKNOWN) || (vt==VT_DISPATCH);
    }

    // Move the variant data to C style data.
    switch(vt)
    {
    case VT_BOOL:
        // convert VT_TRUE and any other non-zero values to TRUE
        *(VARIANT_BOOL*)pv = V_BOOL(pVArgCopy);
        break;

    case VT_I2:
        *(short*)pv = V_I2(pVArgCopy);
        break;

    case VT_ERROR:
    case VT_I4:
        *(long*)pv = V_I4(pVArgCopy);
        break;

    case VT_R4:
        *(float*)pv = V_R4(pVArgCopy);
        break;

    case VT_R8:
        *(double*)pv = V_R8(pVArgCopy);
        break;

    case VT_CY:
        *(CY*)pv = V_CY(pVArgCopy);
        break;

        // All Pointer types.
    case VT_BSTR:
        if(wMaxstrlen && FormsStringLen(V_BSTR(pVArgCopy))>wMaxstrlen)
        {
            hr = FormsAllocStringLen(V_BSTR(pVArgCopy), wMaxstrlen, (BSTR*)pv);
            if(hr)
            {
                goto Cleanup;
            }

            if(*pfAlloc)
            {
                VariantClear(&vargNew);
            }
            else
            {
                *pfAlloc = TRUE;
            }

            goto Cleanup;
        }
    case VT_PTR:
    case VT_LPSTR:
    case VT_LPWSTR:
    case VT_DISPATCH:
    case VT_UNKNOWN:
        *(void**)pv = V_BYREF(pVArgCopy);
        break;

    case VT_VARIANT:
        hr = ClipVarString(pVArgCopy, (VARIANT*)pv, pfAlloc, wMaxstrlen);
        if(hr == S_FALSE)
        {
            hr = S_OK;
            // Copy entire variant to output parameter.
            *(VARIANT*)pv = *pVArgCopy;
        }

        break;

    default:
        Assert(FALSE && "Unknown type in VARIANTARGToCVar().\n");
        hr = DISP_E_TYPEMISMATCH;
        break;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   VARIANTARGToIndex
//
//  Synopsis:   Converts a VARIANT to an index of type long. Sets the index to
//              -1 if the VARIANT type is bad or empty.
//
//  Arguments:  [pvarg]         -- Variant to convert.
//              [plIndex        -- Location to place index.
//
//  Notes:      Useful special case of VARIANTARGToCVar for reading array
//              indices.
//----------------------------------------------------------------------------
HRESULT VARIANTARGToIndex(VARIANT* pvarg, long* plIndex)
{
    HRESULT hr = S_OK;

    Assert(pvarg);
    *plIndex = -1;

    // Quick return for the common case
    if(V_VT(pvarg)==VT_I4 || V_VT(pvarg)==(VT_I4|VT_BYREF))
    {
        *plIndex = (V_VT(pvarg)==VT_I4) ? V_I4(pvarg) : *V_I4REF(pvarg);
        return S_OK;
    }

    if (V_VT(pvarg)==VT_ERROR || V_VT(pvarg)==VT_EMPTY)
    {
        return S_OK;
    }

    // Must perform type corecion
    CVariant varNum;
    hr = VariantChangeTypeEx(&varNum, pvarg, LCID_SCRIPTING, 0, VT_I4);
    if(hr)
    {
        goto Cleanup;
    }

    Assert(V_VT(&varNum)==VT_I4 || V_VT(&varNum)==(VT_I4|VT_BYREF));
    *plIndex = (V_VT(&varNum)==VT_I4) ? V_I4(&varNum) : *V_I4REF(&varNum);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   CVarToVARIANTARG
//
//  Synopsis:   Converts a C-language variable to a VARIANT.
//
//  Arguments:  [pv]    -- Pointer to C-language variable.
//              [vt]    -- Type of C-language variable.
//              [pvarg] -- Resulting VARIANT.  Must be initialized by caller.
//                         Any contents will be freed.
//
//  Modifies:   [pvarg]
//
//  History:    2-23-94   adams   Created
//
//  Notes:      Supports all variant pointer types, VT_UI2, VT_I2, VT_UI4,
//              VT_I4, VT_R4, VT_R8, VT_ERROR.
//
//----------------------------------------------------------------------------
void CVarToVARIANTARG(void* pv, VARTYPE vt, VARIANTARG* pvarg)
{
    Assert(pv);
    Assert(pvarg);

    VariantClear(pvarg);

    V_VT(pvarg) = vt;
    if(V_ISBYREF(pvarg))
    {
        // Use a supported pointer type for derefencing.
        vt = VT_UNKNOWN;
    }

    switch(vt)
    {
    case VT_BOOL:
        // convert TRUE to VT_TRUE
        Assert(*(BOOL*)pv==1 || *(BOOL*)pv==0);
        V_BOOL(pvarg) = VARIANT_BOOL(-*(BOOL*)pv);
        break;

    case VT_I2:
        V_I2(pvarg) = *(short*)pv;
        break;

    case VT_ERROR:
    case VT_I4:
        V_I4(pvarg) = *(long*)pv;
        break;

    case VT_R4:
        V_R4(pvarg) = *(float*)pv;
        break;

    case VT_R8:
        V_R8(pvarg) = *(double*)pv;
        break;

    case VT_CY:
        V_CY(pvarg) = *(CY*)pv;
        break;

        // All Pointer types.
    case VT_PTR:
    case VT_BSTR:
    case VT_LPSTR:
    case VT_LPWSTR:
    case VT_DISPATCH:
    case VT_UNKNOWN:
        V_BYREF(pvarg) = *(void**)pv;
        break;

    default:
        Assert(FALSE && "Unknown type.");
        break;
    }
}