
#include "stdafx.h"

#define VT_BOOL4    0xfe // 254 is not used for another common VT_*

//+-------------------------------------------------------------------------
//
//  Function:   GetNumberOfSize
//              SetNumberOfSize
//
//  Synopsis:   Helpers to get/set an integer value of given byte size
//              by dereferencing a pointer
//
//              pv - pointer to dereference
//              cb - size (1, 2 or 4)
//
//--------------------------------------------------------------------------
long GetNumberOfSize(void* pv, int cb)
{
    switch(cb)
    {
    case 1:
        return *(BYTE*)pv;

    case 2:
        return *(SHORT*)pv;

    case 4:
        return *(LONG*)pv;

    default:
        Assert(FALSE);
        return 0;
    }
}

void SetNumberOfSize(void* pv, int cb, long i)
{
    switch(cb)
    {
    case 1:
        Assert((char)i>=SCHAR_MIN && (char)i<=SCHAR_MAX);
        *(BYTE*)pv = BYTE(i);
        break;

    case 2:
        Assert(i>=SHRT_MIN && i<=SHRT_MAX);
        *(SHORT*)pv = SHORT(i);
        break;

    case 4:
        *(LONG*)pv = i;
        break;

    default:
        Assert(FALSE);
    }
}

//+-------------------------------------------------------------------------
//
//  Function:   GetNumberOfType
//              SetNumberOfType
//
//  Synopsis:   Helpers to get/set an integer value of given variant type
//              by dereferencing a pointer
//
//              pv - pointer to dereference
//              vt - variant type
//
//--------------------------------------------------------------------------
long GetNumberOfType(void* pv, VARENUM vt)
{
    switch(vt)
    {
    case VT_I2:
    case VT_BOOL:
        return *(SHORT*)pv;

    case VT_I4:
    case VT_BOOL4:
        return *(LONG*)pv;

    default:
        Assert(FALSE);
        return 0;
    }
}

void SetNumberOfType(void* pv, VARENUM vt, long l)
{
    switch(vt)
    {
    case VT_BOOL:
        l = l ? VB_TRUE : VB_FALSE;
        //  vvvvvvvvvvv  FALL THROUGH vvvvvvvvvvvvv

    case VT_I2:
        Assert(l>=SHRT_MIN && l<=SHRT_MAX);
        *(SHORT*)pv = SHORT(l);
        break;

    case VT_BOOL4:
        l = l ? VB_TRUE : VB_FALSE;
        //  vvvvvvvvvvv  FALL THROUGH vvvvvvvvvvvvv

    case VT_I4:
        *(LONG UNALIGNED*)pv = l;
        break;

    default:
        Assert(FALSE);
    }
}

//+-----------------------------------------------------
//
//  Member : StripCRLF
//
//  Synopsis : strips CR and LF from a provided string
//      the returned string has been allocated and needs 
//      to be freed by the caller.
//
//+-----------------------------------------------------
static HRESULT StripCRLF(TCHAR* pSrc, TCHAR** ppDest)
{
    HRESULT hr = S_OK;
    long    lLength;
    TCHAR*  pTarget = NULL;


    if(!ppDest)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(!pSrc)
    {
        goto Cleanup;
    }

    lLength = _tcslen(pSrc);
    *ppDest= (TCHAR*)MemAlloc(sizeof(TCHAR)*(lLength+1));
    if(!*ppDest)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pTarget = *ppDest;

    for(; lLength>0; lLength--)
    {
        if((*pSrc!=_T('\r')) && (*pSrc!=_T('\n'))) 
        {
            // we want it.
            *pTarget = *pSrc;
            pTarget++;
        }
        pSrc++;
    }

    *pTarget = _T('\0');

Cleanup:
    RRETURN(hr);
}

HRESULT StringToLong(LPTSTR pch, long* pl)
{
    *pl = _tcstol(pch, 0, 0);
    return S_OK;
}

HRESULT WriteText(IStream* pStm, TCHAR* pch)
{
    if(pch)
    {
        return pStm->Write(pch, _tcslen(pch)*sizeof(TCHAR), 0);
    }
    else
    {
        return S_OK;
    }
}

HRESULT WriteTextLen(IStream* pStm, TCHAR* pch, int nLen)
{
    return pStm->Write(pch, nLen*sizeof(TCHAR), 0);
}

HRESULT WriteTextCStr(CStreamWriteBuff* pStmWrBuff, CString* pstr, BOOL fAlwaysQuote, BOOL fNeverQuote)
{
    UINT u;

    if((u=pstr->Length()) != 0)
    {
        if(fNeverQuote)
        {
            RRETURN(pStmWrBuff->Write(*pstr));
        }
        else
        {
            RRETURN(pStmWrBuff->WriteQuotedText(*pstr, fAlwaysQuote));
        }
    }
    else
    {
        if(!fNeverQuote)
        {
            RRETURN(WriteText(pStmWrBuff, _T("\"\"")));
        }
    }
    return S_OK;
}

HRESULT WriteTextLong(IStream* pStm, long l)
{
    TCHAR ach[20];

    HRESULT hr = Format(_afxGlobalData._hInstResource, 0, ach, ARRAYSIZE(ach), _T("<0d>"), l);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pStm->Write(ach, _tcslen(ach)*sizeof(TCHAR), 0);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

const PROPERTYDESC* HDLDESC::FindPropDescForName(LPCTSTR szName, BOOL fCaseSensitive, long* pidx) const
{
    const PROPERTYDESC* found = NULL;
    STRINGCOMPAREFN pfnCompareString = fCaseSensitive ? StrCmpC : StrCmpIC;
    if(pidx)
    {
        *pidx = -1;
    }

    if(uNumPropDescs)
    {
        int r;
        const PROPERTYDESC* const* low = ppPropDescs;
        const PROPERTYDESC* const* high;
        const PROPERTYDESC* const* mid;

        high = low + uNumPropDescs - 1;

        // Binary search for property name
        while(high >= low)
        {
            if(high != low)
            {
                mid = low + (((high-low)+1)>>1);
                r = pfnCompareString(szName, (*mid)->pstrName);
                if(r < 0)
                {
                    high = mid - 1;
                }
                else if(r > 0)
                {
                    low = mid + 1;
                }
                else
                {
                    found = *mid;
                    if(pidx)
                    {
                        *pidx = mid - ppPropDescs;
                    }
                    break;
                }
            }
            else 
            {
                found = pfnCompareString(szName, (*low)->pstrName)?NULL:*low;
                if(pidx && found)
                {
                    *pidx = low - ppPropDescs;
                }
                break;
            }
        }
    }
    return found;
}

HRESULT ENUMDESC::EnumFromString(LPCTSTR pStr, long* plValue, BOOL fCaseSensitive) const
{   
    int i;
    HRESULT hr = S_OK;              
    STRINGCOMPAREFN pfnCompareString = fCaseSensitive ? StrCmpC : StrCmpIC;

    if(!pStr)
    {
        pStr = _T("");
    }

    for(i=cEnums-1; i>=0; i--)
    {
        if(!(pfnCompareString(pStr, aenumpairs[i].pszName)))
        {
            *plValue = aenumpairs[i].iVal;
            break;
        }
    }

    if(i < 0)
    {
        hr = E_INVALIDARG;
    }

    RRETURN1(hr, E_INVALIDARG);
}

HRESULT ENUMDESC::StringFromEnum(long lValue, BSTR* pbstr) const
{
    int i;
    HRESULT hr = E_INVALIDARG;

    for(i=0; i<cEnums; i++)
    {
        if(aenumpairs[i].iVal == lValue)
        {
            hr = FormsAllocString(aenumpairs[i].pszName, pbstr);
            break;
        }
    }
    RRETURN1(hr, E_INVALIDARG);
}

LPCTSTR ENUMDESC::StringPtrFromEnum(long lValue) const
{
    int     i;

    for(i=0; i<cEnums; i++)
    {
        if(aenumpairs[i].iVal == lValue)
        {
            return aenumpairs[i].pszName;
        }
    }
    return NULL;
}

HRESULT BASICPROPPARAMS::GetAvString(void* pObject, const void* pvParams, CString* pstr, BOOL* pfValuePresent) const
{
    HRESULT hr = S_OK;
    LPCTSTR lpStr;
    BOOL fDummy;

    if(!pfValuePresent)
    {
        pfValuePresent = &fDummy;
    }

    Assert(pstr);

    if(dwPPFlags & PROPPARAM_ATTRARRAY)
    {
        *pfValuePresent = CAttrArray::FindString(*(CAttrArray**)pObject, ((PROPERTYDESC*)this)-1, &lpStr);
        // String pointer will be set to a default if not present
        pstr->Set(lpStr);
    }
    else
    {
        // Stored as offset from a struct
        CString* pstoredstr = (CString*)((BYTE*)pObject+*(DWORD*)pvParams);
        pstr->Set((LPTSTR)*pstoredstr); 
        *pfValuePresent = TRUE;
    }
    RRETURN(hr);
}

DWORD BASICPROPPARAMS::GetAvNumber(void* pObject, const void* pvParams, UINT uNumBytes, BOOL* pfValuePresent) const
{
    DWORD dwValue;
    BOOL fDummy;

    if(!pfValuePresent)
    {
        pfValuePresent = &fDummy;
    }

    if(dwPPFlags & PROPPARAM_ATTRARRAY)
    {
        *pfValuePresent = CAttrArray::FindSimple(*(CAttrArray**)pObject, ((PROPERTYDESC*)this)-1, &dwValue);
    }
    else
    {
        //Stored as offset from a struct
        BYTE* pbValue = (BYTE*)pObject + *(DWORD*)pvParams;
        dwValue = (DWORD)GetNumberOfSize((void*)pbValue, uNumBytes);
        *pfValuePresent = TRUE;
    }
    return dwValue;
}

HRESULT BASICPROPPARAMS::SetAvNumber(void* pObject, DWORD dwNumber, const void* pvParams, UINT uNumberBytes, WORD wFlags/*=0*/) const
{
    HRESULT hr = S_OK;

    if(dwPPFlags & PROPPARAM_ATTRARRAY)
    {
        hr = CAttrArray::SetSimple((CAttrArray**)pObject, ((PROPERTYDESC*)this)-1, dwNumber, wFlags);
    }
    else
    {
        BYTE* pbData = (BYTE*)pObject + *(DWORD*)pvParams;
        SetNumberOfSize((void*)pbData, uNumberBytes, dwNumber);
    }
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Function:   GetGetMethodPtr
//
//  Synopsis:   Helper for getting get/set pointer to member functions
//
//----------------------------------------------------------------
void BASICPROPPARAMS::GetGetMethodP(const void* pvParams, void* pfn) const
{
    Assert(!(dwPPFlags&PROPPARAM_MEMBER));
    Assert(dwPPFlags&PROPPARAM_GETMFHandler);
    Assert(dwPPFlags&PROPPARAM_SETMFHandler);

    memcpy(pfn, pvParams, sizeof(PFN_ARBITRARYPROPGET));
}

//+---------------------------------------------------------------
//
//  Function:   GetSetMethodptr
//
//  Synopsis:   Helper for getting get/set pointer to member functions
//
//----------------------------------------------------------------
void BASICPROPPARAMS::GetSetMethodP(const void* pvParams, void* pfn) const
{
    Assert(!(dwPPFlags&PROPPARAM_MEMBER));
    Assert(dwPPFlags&PROPPARAM_GETMFHandler);
    Assert(dwPPFlags&PROPPARAM_SETMFHandler);

    memcpy(pfn, (BYTE*)pvParams+sizeof(PFN_ARBITRARYPROPGET), sizeof(PFN_ARBITRARYPROPSET));
}

//+---------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::GetColorProperty, public
//
//  Synopsis:   Helper for setting color valued properties
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetColorProperty(VARIANT* pVar, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;

    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    if(!pVar)
    {
        hr = E_INVALIDARG;
    }
    else
    {
        VariantInit(pVar);
        V_VT(pVar) = VT_BSTR;
        hr = GetColor(pSubObject, &(pVar->bstrVal));
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::SetColorProperty, public
//
//  Synopsis:   Helper for setting color valued properties
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetColorProperty(VARIANT var, CBase* pObject, CVoid* pSubObject, WORD wFlags) const
{
    CBase::CLock    Lock(pObject);
    HRESULT         hr=S_OK;
    DWORD           dwOldValue;
    CColorValue     cvValue;
    CVariant*       v = (CVariant*)&var;

    // check for shorts to keep vbscript happy and other VT_U* and VT_UI* types to keep C programmer's happy.
    // vbscript interprets values between 0x8000 and 0xFFFF in hex as a short!, jscript doesn't
    if(v->CoerceNumericToI4())
    {
        DWORD dwRGB = V_I4(v);

        // if -ve value or highbyte!=0x00, ignore (NS compat)
        if(dwRGB & CColorValue::MASK_FLAG)
        {
            goto Cleanup;
        }

        // flip RRGGBB to BBGGRR to be in CColorValue format
        cvValue.SetFromRGB(dwRGB);
    }
    else if(V_VT(&var) == VT_BSTR)
    {
        // Removed 4/24/97 because "" clearing a color is a useful feature.  -CWilso
        // if NULL or empty string, ignore (NS compat)
        //        if (!(V_BSTR(&var)) || !*(V_BSTR(&var)))
        //            goto Cleanup;
        hr = cvValue.FromString((LPTSTR)(V_BSTR(&var)), (dwPPFlags&PROPPARAM_STYLESHEET_PROPERTY));
    }
    else
    {
        goto Cleanup; // if invalid type, ignore
    }

    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    hr = GetColor(pSubObject, &dwOldValue);
    if(hr)
    {
        goto Cleanup;
    }

    if(dwOldValue == (DWORD)cvValue.GetRawValue())
    {
        // No change - ignore it
        hr = S_OK;
        goto Cleanup;
    }

    hr = SetColor(pSubObject, cvValue.GetRawValue(), wFlags);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pObject->OnPropertyChange(dispid, dwFlags);

    if(hr)
    {
        SetColor(pSubObject, dwOldValue, wFlags);
        pObject->OnPropertyChange(dispid, dwFlags);
    }

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------
//
//  Function:   SetString
//
//  Synopsis:   Helper for setting string value
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetString(CVoid* pObject, TCHAR* pch, WORD wFlags) const
{
    HRESULT hr;

    if(dwPPFlags & PROPPARAM_SETMFHandler)
    {
        PFN_CSTRPROPSET pmfSet;
        CString szString;
        szString.Set(pch);

        GetSetMethodP(this+1, &pmfSet);

        hr = CALL_METHOD(pObject,pmfSet,(&szString));
    }
    else
    {

        if(dwPPFlags & PROPPARAM_ATTRARRAY)
        {            
            hr = CAttrArray::SetString((CAttrArray**)(void*)pObject,
                (PROPERTYDESC*)this-1, pch, FALSE, wFlags);
            if(hr)
            {
                goto Cleanup;
            }
        }
        else
        {
            // Stored as offset from a struct
            CString* pcstr;
            pcstr = (CString*)((BYTE*)pObject + *(DWORD*)(this+1));
            hr = pcstr->Set(pch);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Function:   GetString
//
//  Synopsis:   Helper for getting string value
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetString(CVoid* pObject, CString* pcstr, BOOL* pfValuePresent) const
{
    HRESULT hr = S_OK;

    if(dwPPFlags & PROPPARAM_GETMFHandler)
    {
        PFN_CSTRPROPGET pmfGet;

        GetGetMethodP(this+1, &pmfGet);

        // Get Method fn prototype takes a BSTR ptr
        hr = CALL_METHOD(pObject,pmfGet,(pcstr));
        if(pfValuePresent)
        {
            *pfValuePresent = TRUE;
        }
    }
    else
    {
        hr = GetAvString(pObject, this+1, pcstr, pfValuePresent);
    }

    RRETURN(hr);
}

int __cdecl _wcsicmp2(const wchar_t* string1, const wchar_t* string2)
{
    int cc;

    cc = CompareString(_afxGlobalData._lcidUserDefault,
        NORM_IGNORECASE|NORM_IGNOREWIDTH|NORM_IGNOREKANATYPE,
        string1, -1, string2, -1);

    if(cc > 0)
    {
        return (cc-2);
    }
    else
    {
        return _NLSCMPERROR;
    }
}

//+---------------------------------------------------------------
//
//  Member:     SetStringProperty, public
//
//  Synopsis:   Helper for setting string values properties
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetStringProperty(BSTR bstrNew, CBase* pObject, CVoid* pSubObject, WORD wFlags) const
{
    CBase::CLock    Lock(pObject);
    HRESULT         hr;
    CString         cstrOld;
    BOOL            fOldPresent;

    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    hr = GetString(pSubObject, &cstrOld, &fOldPresent);
    if(hr)
    {
        goto Cleanup;
    }

    // HACK! HACK! HACK! (MohanB) In order to fix #64710 at this very late
    // stage in IE5, I am putting in this hack to specifically check for
    // DISPID_CElement_id. For IE6, we should not fire onpropertychange for any
    // property (INCLUDING non-string type properties!) if the value of the
    // property has not been modified.

    // Quit if the value has not been modified
    if(fOldPresent && DISPID_IHTMLELEMENT_ID==dispid && 0==_wcsicmp2(cstrOld, bstrNew))
    {
        goto Cleanup;
    }

    hr = SetString(pSubObject, (TCHAR*)bstrNew, wFlags);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pObject->OnPropertyChange(dispid, dwFlags);

    if(hr)
    {
        SetString(pSubObject, (TCHAR*)cstrOld, wFlags);
        pObject->OnPropertyChange(dispid, dwFlags);
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::GetStringProperty, public
//
//  Synopsis:   Helper for setting string valued properties
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetStringProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;

    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    if(!pbstr)
    {
        hr = E_INVALIDARG;
    }
    else
    {
        CString cstr;

        // BUGBUG: (anandra) Possibly have a helper here that returns
        // a bstr to avoid two allocations.
        hr = GetString(pSubObject, &cstr);
        if(hr)
        {
            goto Cleanup;
        }
        hr = cstr.AllocBSTR(pbstr);
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Function:   SetUrl
//
//  Synopsis:   Helper for setting url string value
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetUrl(CVoid* pObject, TCHAR* pch, WORD wFlags) const
{
    HRESULT hr = S_OK;
    TCHAR* pstrNoCRLF=NULL;

    hr = StripCRLF(pch, &pstrNoCRLF);
    if(hr)
    {
        goto Cleanup;
    }

    hr =SetString(pObject, pstrNoCRLF, wFlags);

Cleanup:
    if(pstrNoCRLF)
    {
        MemFree(pstrNoCRLF);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::GetUrlProperty, public
//
//  Synopsis:   Helper for setting url string valued properties
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetUrlProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const
{
    // SetString does the SetErrorInfoPGet
    RRETURN(GetStringProperty(pbstr, pObject, pSubObject));
}

//+---------------------------------------------------------------
//
//  Member:     SetUrlProperty, public
//
//  Synopsis:   Helper for setting url-string values properties
//     strip off CR/LF and call setString
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetUrlProperty(BSTR bstrNew, CBase* pObject, CVoid* pSubObject, WORD wFlags) const
{
    HRESULT hr=S_OK;
    TCHAR* pstrNoCRLF=NULL;

    hr = StripCRLF((TCHAR*)bstrNew, &pstrNoCRLF);
    if(hr)
    {
        goto Cleanup;
    }

    // SetStringProperty calls Set ErrorInfoPSet
    hr = SetStringProperty(pstrNoCRLF, pObject, pSubObject, wFlags);

Cleanup:
    if(pstrNoCRLF)
    {
        MemFree(pstrNoCRLF);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Function:   GetUrl
//
//  Synopsis:   Helper for getting url-string value
//
//----------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetUrl(CVoid* pObject, CString* pcstr) const
{
    RRETURN(GetString(pObject, pcstr));
}

HRESULT BASICPROPPARAMS::GetColor(CVoid* pObject, BSTR* pbstrColor, BOOL fReturnAsHex/*==FALSE*/) const
{
    HRESULT hr;
    TCHAR szBuffer[64];
    CColorValue cvColor;

    Assert(!(dwPPFlags & PROPPARAM_GETMFHandler));

    cvColor = (CColorValue)GetAvNumber(pObject, this+1, sizeof(CColorValue));
    hr = cvColor.FormatBuffer(szBuffer, ARRAYSIZE(szBuffer), fReturnAsHex);
    if(hr)
    {
        goto Cleanup;
    }

    hr = FormsAllocString(szBuffer, pbstrColor);
Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::SetCodeProperty
//
//----------------------------------------------------------------------------
HRESULT BASICPROPPARAMS::SetCodeProperty(VARIANT* pvarCode, CBase* pObject, CVoid*) const
{
    HRESULT hr = S_OK;

    Assert((dwPPFlags&PROPPARAM_ATTRARRAY) && "only attr array support implemented for 'type:code'");

    if(V_VT(pvarCode) == VT_NULL)
    {
        hr = pObject->SetCodeProperty(dispid, NULL);
    }
    else if(V_VT(pvarCode) == VT_DISPATCH)
    {
        hr = pObject->SetCodeProperty(dispid, V_DISPATCH(pvarCode));
    }
    else if(V_VT(pvarCode) == VT_BSTR)
    {
        pObject->FindAAIndexAndDelete(dispid, CAttrValue::AA_Internal);
        hr = pObject->AddString(dispid, V_BSTR(pvarCode), CAttrValue::AA_Attribute);
    }
    else
    {
        hr = E_NOTIMPL;
    }

    if(!hr)
    {
        hr = pObject->OnPropertyChange(dispid, dwFlags);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     BASICPROPPARAMS::GetCodeProperty
//
//----------------------------------------------------------------------------
HRESULT BASICPROPPARAMS::GetCodeProperty(VARIANT* pvarCode, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
    AAINDEX aaidx;

    V_VT(pvarCode) = VT_NULL;

    Assert((dwPPFlags&PROPPARAM_ATTRARRAY) && "only attr array support implemented for 'type:code'");

    aaidx = pObject->FindAAIndex(dispid, CAttrValue::AA_Internal);
    if((AA_IDX_UNKNOWN!=aaidx) && (pObject->GetVariantTypeAt(aaidx)==VT_DISPATCH))
    {
        hr = pObject->GetDispatchObjectAt(aaidx, &V_DISPATCH(pvarCode));
        if(hr)
        {
            goto Cleanup;
        }
        V_VT(pvarCode) = VT_DISPATCH;
    }
    else
    {
        LPCTSTR szCodeText;
        aaidx = pObject->FindAAIndex(dispid, CAttrValue::AA_Attribute);
        if(AA_IDX_UNKNOWN != aaidx)
        {
            hr = pObject->GetStringAt(aaidx, &szCodeText);
            if(hr)
            {
                goto Cleanup;
            }
            hr = FormsAllocString(szCodeText, &V_BSTR(pvarCode));
            if(hr)
            {
                goto Cleanup;            
            }
            V_VT(pvarCode) = VT_BSTR;
        }
    }
Cleanup:
    RRETURN(hr);
}

HRESULT BASICPROPPARAMS::GetColor(CVoid* pObject, CString* pcstr, BOOL fReturnAsHex/*=FALSE*/, BOOL* pfValuePresent/*=NULL*/) const
{
    HRESULT     hr;
    TCHAR       szBuffer[64];
    CColorValue cvColor;

    Assert(!(dwPPFlags & PROPPARAM_GETMFHandler));

    cvColor = (CColorValue)GetAvNumber(pObject, this+1, sizeof(CColorValue), pfValuePresent);
    hr = cvColor.FormatBuffer(szBuffer, ARRAYSIZE(szBuffer), fReturnAsHex);
    if(hr)
    {
        goto Cleanup;
    }
    hr = pcstr->Set(szBuffer);
Cleanup:
    RRETURN(hr);
}

HRESULT BASICPROPPARAMS::GetColor(CVoid* pObject, DWORD* pdwValue) const
{
    CColorValue cvColor;

    Assert(!(dwPPFlags & PROPPARAM_GETMFHandler));

    cvColor = (CColorValue)GetAvNumber(pObject, this+1, sizeof(CColorValue));
    *pdwValue = cvColor.GetRawValue();

    return S_OK;
}

HRESULT BASICPROPPARAMS::SetColor(CVoid* pObject, TCHAR* pch, WORD wFlags) const
{
    HRESULT hr;
    CColorValue cvColor;

    hr = cvColor.FromString(pch);
    if(hr)
    {
        goto Cleanup;
    }

    Assert(!(dwPPFlags & PROPPARAM_SETMFHandler));

    hr = SetAvNumber(pObject, (DWORD)cvColor.GetRawValue(), this+1, sizeof(CColorValue), wFlags);

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

HRESULT BASICPROPPARAMS::SetColor(CVoid* pObject, DWORD dwValue, WORD wFlags) const
{
    HRESULT hr;
    CColorValue cvColor;

    Assert(!(dwPPFlags & PROPPARAM_SETMFHandler));

    cvColor.SetRawValue(dwValue);

    hr = SetAvNumber(pObject, (DWORD)cvColor.GetRawValue(), this+1, sizeof(CColorValue), wFlags);

    RRETURN1(hr,E_INVALIDARG);
}



HRESULT LookupEnumString(const NUMPROPPARAMS* ppp, LPCTSTR pStr, long* plNewValue)
{
    HRESULT         hr = S_OK;
    PROPERTYDESC*   ppdPropDesc = ((PROPERTYDESC*)ppp) - 1;
    ENUMDESC*       pEnumDesc = ppdPropDesc->GetEnumDescriptor();

    // ### v-gsrir - Modified the 3rd parameter to result in a 
    // Bool instead of a DWORD - which was resulting in the parameter
    // being taken as 0 always.
    hr = pEnumDesc->EnumFromString(pStr, plNewValue,
        (ppp->bpp.dwPPFlags&PROPPARAM_CASESENSITIVE)?TRUE:FALSE);

    if(hr == E_INVALIDARG)
    {
        if(ppp->bpp.dwPPFlags & PROPPARAM_ANUMBER)
        {
            // Not one of the enums, is it a number??
            hr = ttol_with_error(pStr, plNewValue);
        }
    }
    RRETURN(hr);
}

// Don't call this function outside of automation - loads oleaut!
HRESULT LookupStringFromEnum(const NUMPROPPARAMS* ppp, BSTR* pbstr, long lValue)
{
    ENUMDESC* pEnumDesc;

    Assert((ppp->bpp.dwPPFlags&PROPPARAM_ENUM) && "Can't convert a non-enum to an enum string!");

    if(ppp->bpp.dwPPFlags & PROPPARAM_ANUMBER)
    {
        pEnumDesc = *(ENUMDESC**)((BYTE*)(ppp+1) + sizeof(DWORD_PTR));
    }
    else
    {
        pEnumDesc = (ENUMDESC*)ppp->lMax;
    }

    RRETURN1(pEnumDesc->StringFromEnum(lValue, pbstr), S_FALSE);
}

BOOL PROPERTYDESC::IsBOOLProperty(void) const
{
    return ((pfnHandleProperty==&PROPERTYDESC::HandleNumProperty &&
        ((NUMPROPPARAMS*)(this+1))->vt==VT_BOOL)?TRUE:FALSE);
}

//+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleNumProperty, public
//
//  Synopsis:   Helper for getting/setting number value properties
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (long in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//----------------------------------------------------------------
HRESULT PROPERTYDESC::HandleNumProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK, hr2;
    const NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(this + 1);
    const BASICPROPPARAMS* bpp = (BASICPROPPARAMS*)ppp;
    VARIANT varDest;
    LONG lNewValue, lTemp;
    CStreamWriteBuff* pStmWrBuff;
    VARIANT* pVariant;
    VARIANT varTemp;

    varDest.vt = VT_EMPTY;

    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        switch(PROPTYPE(dwOpCode))
        {
        case PROPTYPE_LPWSTR:
            // If the parameter is a BOOL, the presence of the value sets the value
            if(ppp->vt == VT_BOOL)
            {
                // For an OLE BOOL -1 == True
                lNewValue = -1;
            }
            else
            {
SetFromString:
                LPTSTR pStr = (TCHAR*)pv;
                hr = S_OK;
                if(pStr==NULL || !pStr[0])
                {
                    // Just have Tag=, set to not assigned value
                    lNewValue = GetNotPresentDefault();
                }
                else
                {
                    if(!(ppp->bpp.dwPPFlags & PROPPARAM_ENUM))
                    {
                        hr = ttol_with_error(pStr, &lNewValue);
                    }
                    else
                    {
                        // enum string, look it up
                        hr = LookupEnumString(ppp, pStr, &lNewValue);
                    }
                    if(hr)
                    {
                        lNewValue = GetInvalidDefault();
                    }
                }
            }
            // If we're just sniffing for a good parse - don't set up a default
            if(hr && ISSAMPLING(dwOpCode))
            {
                goto Cleanup;
            }
            pv = &lNewValue;
            break;

        case PROPTYPE_VARIANT:
            if(pv == NULL)
            {
                // Just have Tag=, ignore it
                return S_OK;
            }
            VariantInit(&varDest);
            if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
            {
                hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pv, VT_BSTR);
                if(hr)
                {
                    goto Cleanup;
                }
                pv = V_BSTR(&varDest);
                // Send it through the String handler
                goto SetFromString;
            }
            else
            {
                hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pv, VT_I4);
                if(hr)
                {
                    goto Cleanup;
                }
                pv = &V_I4(&varDest);
            }
            break;
        default:
            Assert(PROPTYPE(dwOpCode) == 0); // assumed native long
        }
        WORD wFlags = 0;
        if(dwOpCode & HANDLEPROP_IMPORTANT)
        {
            wFlags |= CAttrValue::AA_Extra_Important;
        }
        if(dwOpCode & HANDLEPROP_IMPLIED)
        {
            wFlags |= CAttrValue::AA_Extra_Implied;
        }
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_VALUE:
            // We need to preserve an existing hr
            hr2 = ppp->ValidateNumberProperty((long*)pv, pObject);
            if(hr2)
            {
                hr = hr2;
                goto Cleanup;
            }
            hr2 = ppp->SetNumber(pObject, pSubObject, *(long*)pv, wFlags);
            if(hr2)
            {
                hr = hr2;
                goto Cleanup;
            }

            if(dwOpCode & HANDLEPROP_MERGE)
            {
                hr2 = pObject->OnPropertyChange(bpp->dispid, bpp->dwFlags);
                if(hr2)
                {
                    hr = hr2;
                    goto Cleanup;
                }
            }
            break;

        case HANDLEPROP_AUTOMATION:
            if(hr)
            {
                goto Cleanup;
            }
            hr = ppp->SetNumberProperty(*(long*)pv, pObject, pSubObject,
                dwOpCode&HANDLEPROP_DONTVALIDATE?FALSE:TRUE, wFlags);
            break;

        case HANDLEPROP_DEFAULT:
            Assert(pv == NULL);
            hr = ppp->SetNumber(pObject, pSubObject,
                (long)ulTagNotPresentDefault, wFlags);
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }
    else
    {
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_AUTOMATION:
            hr = ppp->GetNumberProperty(&lTemp, pObject, pSubObject);
            if(hr)
            {
                goto Cleanup;
            }
            switch(PROPTYPE(dwOpCode))
            {
            case PROPTYPE_VARIANT:
                pVariant = (VARIANT*)pv;
                VariantInit(pVariant);
                if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
                {
                    hr = LookupStringFromEnum(ppp, &V_BSTR(pVariant), lTemp);
                    if(!hr)
                    {
                        V_VT(pVariant) = VT_BSTR;
                    }
                    else
                    {
                        if(hr != E_INVALIDARG)
                        {
                            goto Cleanup;
                        }
                        hr = S_OK;
                    }
                }
                if(V_VT(pVariant) == VT_EMPTY)
                {
                    V_VT(pVariant) = VT_I4;
                    V_I4(pVariant) = lTemp;
                }
                break;

            case PROPTYPE_BSTR:
                *((BSTR*)pv) = NULL;
                if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
                {
                    hr = LookupStringFromEnum(ppp, (BSTR*)pv, lTemp);
                }
                else
                {
                    TCHAR szNumber[33];
                    hr = Format(_afxGlobalData._hInstResource, 0, szNumber, 32, _T("<0d>"), lTemp);
                    if(hr == S_OK)
                    {
                        hr = FormsAllocString(szNumber, (BSTR*)pv);
                    }
                }
                break;

            default:
                *(long*)pv = lTemp;
                break;
            }
            break;

        case HANDLEPROP_STREAM:
            // until the binary persistance we assume text save
            Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);
            hr = ppp->GetNumber(pObject, pSubObject, &lNewValue);
            if(hr)
            {
                goto Cleanup;
            }
            pStmWrBuff = (CStreamWriteBuff*)pv;
            // If it's one of the enums, save out the enum string
            if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
            {
                INT i;
                ENUMDESC* pEnumDesc = GetEnumDescriptor();
                // until the binary persistance we assume text save
                for(i=pEnumDesc->cEnums-1; i>=0; i--)
                {
                    if(lNewValue == pEnumDesc->aenumpairs[i].iVal)
                    {
                        hr = pStmWrBuff->WriteQuotedText(pEnumDesc->aenumpairs[i].pszName, FALSE);
                        goto Cleanup;
                    }
                }
            }
            // Either don't have an enum array or wasn't one of the enum values
            hr = WriteTextLong(pStmWrBuff, lNewValue);
            break;

        case HANDLEPROP_VALUE:
            hr = ppp->GetNumber(pObject, pSubObject, &lNewValue);
            if(hr)
            {
                goto Cleanup;
            }
            switch(PROPTYPE(dwOpCode))
            {
            case VT_VARIANT:
                {
                    ENUMDESC* pEnumDesc = GetEnumDescriptor();
                    if(pEnumDesc)
                    {
                        hr = pEnumDesc->StringFromEnum(lNewValue, &V_BSTR((VARIANT*)pv));
                        if(!hr)
                        {
                            ((VARIANT*)pv)->vt = VT_BSTR;
                            goto Cleanup;
                        }
                    }

                    // if the Numeric prop is boolean...
                    if(GetNumPropParams()->vt == VT_BOOL)
                    {
                        ((VARIANT*)pv)->boolVal = (VARIANT_BOOL)lNewValue;
                        ((VARIANT*)pv)->vt = VT_BOOL;
                        break;
                    }

                    // Either mixed enum/integer or plain integer return I4
                    ((VARIANT*)pv)->lVal = lNewValue;
                    ((VARIANT*)pv)->vt = VT_I4;
                }
                break;

            case VT_BSTR:
                {
                    ENUMDESC* pEnumDesc = GetEnumDescriptor();
                    if(pEnumDesc)
                    {
                        hr = pEnumDesc->StringFromEnum(lNewValue, (BSTR*)pv);
                        if(!hr)
                        {
                            goto Cleanup;
                        }
                    }
                    // Either mixed enum/integer or plain integer return BSTR
                    varTemp.lVal = lNewValue;
                    varTemp.vt = VT_I4;   
                    hr = VariantChangeTypeSpecial(&varTemp, &varTemp, VT_BSTR);
                    *(BSTR*)pv = V_BSTR(&varTemp);
                }
                break;

            default:
                Assert(PROPTYPE(dwOpCode) == 0);
                *(long*)pv = lNewValue;
                break;
            }
            break;

        case HANDLEPROP_COMPARE:
            hr = ppp->GetNumber(pObject, pSubObject, &lNewValue);
            if(hr)
            {
                goto Cleanup;
            }
            hr = (lNewValue == *(long*)pv) ? S_OK : S_FALSE;
            RRETURN1(hr, S_FALSE);
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }

Cleanup:
    if(varDest.vt != VT_EMPTY)
    {
        VariantClear(&varDest);
    }

    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleStringProperty, public
//
//  Synopsis:   Helper for getting/setting string value properties
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (CStr in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//----------------------------------------------------------------
HRESULT PROPERTYDESC::HandleStringProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
    BASICPROPPARAMS* ppp = (BASICPROPPARAMS*)(this + 1);
    VARIANT varDest;

    varDest.vt = VT_EMPTY;
    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        switch(PROPTYPE(dwOpCode))
        {
        case PROPTYPE_LPWSTR:
            break;
        case PROPTYPE_VARIANT:
            if(V_VT((VARIANT*)pv) == VT_BSTR)
            {
                pv = (void*)V_BSTR((VARIANT*)pv);
            }
            else
            {
                hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pv, VT_BSTR);
                if(hr)
                {
                    goto Cleanup;
                }
                pv = V_BSTR(&varDest);
            }
        default:
            Assert(PROPTYPE(dwOpCode) == 0); // assumed native long
        }
        WORD wFlags = 0;
        if(dwOpCode & HANDLEPROP_IMPORTANT)
        {
            wFlags |= CAttrValue::AA_Extra_Important;
        }
        if(dwOpCode & HANDLEPROP_IMPLIED)
        {
            wFlags |= CAttrValue::AA_Extra_Implied;
        }
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_DEFAULT:
            Assert(pv == NULL);
            pv = (void*)ulTagNotPresentDefault;

            if(!pv)
            {
                goto Cleanup; // zero string
            }

            // fall thru
        case HANDLEPROP_VALUE:
            hr = ppp->SetString(pSubObject, (TCHAR*)pv, wFlags);
            if(dwOpCode & HANDLEPROP_MERGE)
            {
                hr = pObject->OnPropertyChange(ppp->dispid, ppp->dwFlags);
                if(hr)
                {
                    goto Cleanup;
                }
            }
            break;
        case HANDLEPROP_AUTOMATION:
            hr = ppp->SetStringProperty((TCHAR*)pv, pObject, pSubObject, wFlags);
            break;
        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }
    else
    {
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_AUTOMATION:
            {
                BSTR bstr;
                hr = ppp->GetStringProperty(&bstr, pObject, pSubObject);
                if(hr)
                {
                    goto Cleanup;
                }
                switch(PROPTYPE(dwOpCode))
                {
                case PROPTYPE_VARIANT:
                    V_VT((VARIANTARG*)pv) = VT_BSTR;
                    V_BSTR((VARIANTARG*)pv) = bstr;
                    break;
                case PROPTYPE_BSTR:
                    *(BSTR*)pv = bstr;
                    break;
                default:
                    Assert("Wrong type for property!");
                    hr = E_INVALIDARG;
                    goto Cleanup;
                }

            }
            break;

        case HANDLEPROP_STREAM:
            {
                CString cstr;

                // until the binary persistance we assume text save
                Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);
                hr = ppp->GetString(pSubObject, &cstr);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = WriteTextCStr((CStreamWriteBuff*)pv, &cstr, FALSE, (ppp->dwPPFlags&PROPPARAM_STYLESHEET_PROPERTY));
            }
            break;
        case HANDLEPROP_VALUE:
            if(PROPTYPE(dwOpCode) == PROPTYPE_VARIANT)
            {
                CString cstr;

                // BUGBUG istvanc this is bad, we allocate the string twice
                // we need a new API which allocates the BSTR directly
                hr = ppp->GetString(pSubObject, &cstr);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = cstr.AllocBSTR(&((VARIANT*)pv)->bstrVal);
                if(hr)
                {
                    goto Cleanup;
                }

                ((VARIANT*)pv)->vt = VT_BSTR;
            }
            else if(PROPTYPE(dwOpCode) == PROPTYPE_BSTR)
            {
                CString cstr;

                // BUGBUG istvanc this is bad, we allocate the string twice
                // we need a new API which allocates the BSTR directly
                hr = ppp->GetString(pSubObject, &cstr);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = cstr.AllocBSTR((BSTR*)pv);
            }
            else
            {
                Assert(PROPTYPE(dwOpCode) == 0);
                hr = ppp->GetString(pSubObject, (CString*)pv);
            }
            break;
        case HANDLEPROP_COMPARE:
            {
                CString cstr;
                hr = ppp->GetString(pSubObject, &cstr);
                if(hr)
                {
                    goto Cleanup;
                }
                LPTSTR lpThisString = (LPTSTR)cstr;
                if(lpThisString==NULL || *(TCHAR**)pv==NULL)
                {
                    hr = (lpThisString==NULL && *(TCHAR**)pv==NULL) ? S_OK : S_FALSE;
                }
                else
                {
                    hr = _tcsicmp(lpThisString, *(TCHAR**)pv) ? S_FALSE : S_OK;
                }
            }
            RRETURN1(hr, S_FALSE);
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }

Cleanup:
    if(varDest.vt != VT_EMPTY)
    {
        VariantClear(&varDest);
    }
    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleEnumProperty, public
//
//  Synopsis:   Helper for getting/setting enum value properties
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (long in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//----------------------------------------------------------------
HRESULT PROPERTYDESC::HandleEnumProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;
    NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(this + 1);
    long lNewValue;

    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        Assert(ppp->bpp.dwPPFlags & PROPPARAM_ENUM);
    }
    else if(OPCODE(dwOpCode) == HANDLEPROP_STREAM)
    {
        RRETURN1(HandleNumProperty(dwOpCode, pv, pObject, pSubObject), E_INVALIDARG);
    }
    else if(OPCODE(dwOpCode) == HANDLEPROP_COMPARE)
    {
        hr = ppp->GetNumber(pObject, pSubObject, &lNewValue);
        if(hr)
        {
            goto Cleanup;
        }
        // if inherited and not set, return S_OK
        hr = (lNewValue==*(long*)pv) ? S_OK : S_FALSE;
        RRETURN1(hr, S_FALSE);
    }

    hr = HandleNumProperty(dwOpCode, pv, pObject, pSubObject);

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleColorProperty, public
//
//  Synopsis:   Helper for getting/setting enum value properties
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (OLE_COLOR in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//----------------------------------------------------------------
HRESULT PROPERTYDESC::HandleColorProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK, hr2;
    BASICPROPPARAMS* ppp = (BASICPROPPARAMS*)(this + 1);
    VARIANT varDest;
    DWORD dwNewValue;
    CColorValue cvValue;
    LPTSTR pStr;

    // just to be a little paranoid
    Assert(sizeof(OLE_COLOR) == 4);
    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        VariantInit(&varDest);
        switch(PROPTYPE(dwOpCode))
        {
        case PROPTYPE_LPWSTR:
            V_VT(&varDest) = VT_BSTR;
            V_BSTR(&varDest) = (BSTR)pv;
            break;
        case PROPTYPE_VARIANT:
            if(pv == NULL)
            {
                // Just have Tag=, ignore it
                return S_OK;
            }
            break;
        default:
            Assert(PROPTYPE(dwOpCode) == 0);    // assumed native long
        }
        WORD wFlags = 0;
        if(dwOpCode & HANDLEPROP_IMPORTANT)
        {
            wFlags |= CAttrValue::AA_Extra_Important;
        }
        if(dwOpCode & HANDLEPROP_IMPLIED)
        {
            wFlags |= CAttrValue::AA_Extra_Implied;
        }
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_VALUE:
            pStr = (LPTSTR)pv;
            if(pStr==NULL || !pStr[0])
            {
                if(IsNotPresentAsDefaultSet() || !UseNotAssignedValue())
                {
                    cvValue = *(CColorValue*)&ulTagNotPresentDefault;
                }
                else
                {
                    cvValue = *(CColorValue*)&ulTagNotAssignedDefault;
                }
            }
            else
            {
                hr = cvValue.FromString(pStr, (ppp->dwPPFlags&PROPPARAM_STYLESHEET_PROPERTY));
                if(hr || S_OK!=cvValue.IsValid())
                {
                    if(IsInvalidAsNoAssignSet() && UseNotAssignedValue())
                    {
                        cvValue = *(CColorValue*)&ulTagNotAssignedDefault;
                    }
                    else
                    {
                        cvValue = *(CColorValue*)&ulTagNotPresentDefault;
                    }
                }
            }
            // If we're just sniffing for a good parse - don't set up a default
            if(hr && ISSAMPLING(dwOpCode))
            {
                goto Cleanup;
            }

            // if asp string, we need to store the original string itself as an unknown attr
            // skip leading white space.
            while(pStr && *pStr && _istspace(*pStr))
            {
                pStr++;
            }

            if(pStr && (*pStr==_T('<')) && (*(pStr+1)==_T('%')))
            {
                hr = E_INVALIDARG;
            }

            // We need to preserve the hr if there is one
            hr2 = ppp->SetColor(pSubObject, cvValue.GetRawValue(), wFlags);
            if(hr2)
            {
                hr = hr2;
                goto Cleanup;
            }

            if(dwOpCode & HANDLEPROP_MERGE)
            {
                hr2 = pObject->OnPropertyChange(ppp->dispid, ppp->dwFlags);
                if(hr2)
                {
                    hr = hr2;
                    goto Cleanup;
                }
            }
            break;

        case HANDLEPROP_AUTOMATION:
            if(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR)
            {
                hr = ppp->SetColorProperty(varDest, pObject, pSubObject, wFlags);
            }
            else
            {
                hr = ppp->SetColorProperty(*(VARIANT*)pv, pObject, pSubObject, wFlags);
            }
            break;

        case HANDLEPROP_DEFAULT:
            Assert(pv == NULL);
            // CColorValue is initialized to VALUE_UNDEF.
            hr = ppp->SetColor(pSubObject, *(DWORD*)&ulTagNotPresentDefault, wFlags);
            break;
        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }
    else
    {
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_AUTOMATION:
            {
                // this code path results from OM access. OM color's are 
                // always returned in Pound6 format.
                // make sure there is a subObject
                if(!pSubObject)
                {
                    pSubObject = pObject;
                }

                // get the dw Color value
                V_VT(&varDest) = VT_BSTR;
                V_BSTR(&varDest) = NULL;
                ppp->GetColor(pSubObject, &(V_BSTR(&varDest)), !(ppp->dwPPFlags&PROPPARAM_STYLESHEET_PROPERTY));

                // Coerce to type to return
                switch(PROPTYPE(dwOpCode))
                {
                case PROPTYPE_VARIANT:
                    VariantInit((VARIANT*)pv);
                    V_VT ((VARIANT*)pv) = VT_BSTR;
                    V_BSTR((VARIANT*)pv) = V_BSTR(&varDest);
                    break;
                default:
                    *(BSTR*)pv = V_BSTR(&varDest);
                }
                hr = S_OK;
            }
            break;
        case HANDLEPROP_STREAM:
            // until the binary persistance we assume text save
            Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);
            hr = ppp->GetColor(pSubObject, &dwNewValue);
            if(hr)
            {
                goto Cleanup;
            }
            cvValue = dwNewValue;

            hr = cvValue.Persist((IStream*)pv);
            break;
        case HANDLEPROP_VALUE:
            if(PROPTYPE(dwOpCode) == PROPTYPE_VARIANT)
            {
                ((VARIANT*)pv)->vt = VT_UI4;
                hr = ppp->GetColor(pSubObject, &((VARIANT*)pv)->ulVal);
            }
            else if(PROPTYPE(dwOpCode) == PROPTYPE_BSTR)
            {
                hr = ppp->GetColorProperty(&varDest, pObject, pSubObject);
                if(hr)
                {
                    goto Cleanup;
                }
                *(BSTR*)pv = V_BSTR(&varDest);
            }
            else
            {
                Assert(PROPTYPE(dwOpCode) == 0);
                hr = ppp->GetColor(pSubObject, (DWORD*)pv);
            }
            break;
        case HANDLEPROP_COMPARE:
            hr = ppp->GetColor(pSubObject, &dwNewValue);
            if(hr)
            {
                goto Cleanup;
            }
            // if inherited and not set, return S_OK
            hr = dwNewValue == *(DWORD*)pv ? S_OK : S_FALSE;
            RRETURN1(hr, S_FALSE);
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
            break;
        }
    }

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------------------
//  Member: HandleUrlProperty
//
//  Synopsis : Handler for Getting/setting Url typed properties
//              pretty much, the only thing this needs to do is strip out CR/LF;s and 
//                  call the HandleStringProperty
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (OLE_COLOR in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//+---------------------------------------------------------------------------
HRESULT PROPERTYDESC::HandleUrlProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
    TCHAR*  pstrNoCRLF = NULL;
    VARIANT varDest;

    varDest.vt = VT_EMPTY;
    if(ISSET(dwOpCode))
    {
        // First handle the storage type
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        switch(PROPTYPE(dwOpCode))
        {
        case PROPTYPE_LPWSTR:
            break;
        case PROPTYPE_VARIANT:
            if(V_VT((VARIANT*)pv) == VT_BSTR)
            {
                pv = (void*)V_BSTR((VARIANT*)pv);
            }
            else
            {
                hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pv,  VT_BSTR);
                if(hr)
                {
                    goto Cleanup;
                }
                pv = V_BSTR(&varDest);
            }
        default:
            Assert(PROPTYPE(dwOpCode) == 0); // assumed native long
        }

        //then handle the operation
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_VALUE:
        case HANDLEPROP_DEFAULT:
            hr = StripCRLF((TCHAR*)pv, &pstrNoCRLF);
            if(hr)
            {
                goto Cleanup;
            }

            hr = HandleStringProperty(dwOpCode, (void*)pstrNoCRLF, pObject, pSubObject);

            if(pstrNoCRLF)
            {
                MemFree(pstrNoCRLF);
            }
            goto Cleanup;
            break;
        }
    }
    else
    {   // BUGBUG: This code should not be necessary - it forces us to always quote URLs when
        // persisting.  It is here to circumvent an Athena bug (IEv4.1 RAID #20953) - CWilso
        if(ISSTREAM(dwOpCode))
        {
            // Have to make sure we quote URLs when persisting.
            CString cstr;
            BASICPROPPARAMS* ppp = (BASICPROPPARAMS*)(this + 1);

            // until the binary persistance we assume text save
            Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);
            hr = ppp->GetString(pSubObject, &cstr);
            if(hr)
            {
                goto Cleanup;
            }
            hr = WriteTextCStr((CStreamWriteBuff*)pv, &cstr, TRUE, FALSE);
            goto Cleanup;
        }
    }

    hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);

Cleanup:
    if(varDest.vt != VT_EMPTY)
    {
        VariantClear(&varDest);
    }
    if(OPCODE(dwOpCode)== HANDLEPROP_COMPARE)
    {
        RRETURN1( hr, S_FALSE);
    }
    else
    {
        RRETURN1(hr, E_INVALIDARG);
    }
}

//+---------------------------------------------------------------------------
//
//  Member: PROPERTYDESC::HandleCodeProperty
//
//+---------------------------------------------------------------------------
HRESULT PROPERTYDESC::HandleCodeProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
    BASICPROPPARAMS* ppp = (BASICPROPPARAMS*)(this + 1);

    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...

        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_DEFAULT:
        case HANDLEPROP_VALUE:
            hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
            break;

        case HANDLEPROP_AUTOMATION:
            Assert (PROPTYPE_VARIANT==PROPTYPE(dwOpCode) && "Only this type supported for this case");

            hr = ppp->SetCodeProperty((VARIANT*)pv, pObject, pSubObject);
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
            break;
        }
    }
    else
    {
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_VALUE:
        case HANDLEPROP_STREAM:
            hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
            break;

        case HANDLEPROP_COMPARE:
            hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
            RRETURN1(hr,S_FALSE);
            break;

        case HANDLEPROP_AUTOMATION:
            if(!pv)
            {
                hr = E_POINTER;
                goto Cleanup;
            }

            Assert(PROPTYPE_VARIANT==PROPTYPE(dwOpCode) && "Only V_DISPATCH supported for this case");

            hr = ppp->GetCodeProperty((VARIANT*)pv, pObject, pSubObject);
            if(hr)
            {
                goto Cleanup;
            }
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
        }
    }

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

HRESULT STDMETHODCALLTYPE PROPERTYDESC::HandleUnitValueProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK, hr2;
    const NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(this + 1);
    const BASICPROPPARAMS* bpp = (BASICPROPPARAMS*)ppp;
    CVariant varDest;
    LONG lNewValue;
    long lTemp;
    BOOL fValidated = FALSE;

    // The code in this fn assumes that the address of the object
    // is the address of the long value
    if(ISSET(dwOpCode))
    {
        CUnitValue uvValue;
        uvValue.SetDefaultValue();
        LPCTSTR pStr = (TCHAR*)pv;

        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...

        switch(PROPTYPE(dwOpCode))
        {
        case PROPTYPE_LPWSTR:
FromString:
            if(pStr==NULL || !pStr[0])
            {
                lNewValue = GetNotPresentDefault();
            }
            else
            {
                hr = uvValue.FromString(pStr, this);
                if(!hr)
                {
                    lNewValue = uvValue.GetRawValue();
                    if(!(dwOpCode&HANDLEPROP_DONTVALIDATE))
                    {
                        hr = uvValue.IsValid(this);
                    }
                }

                if(hr == E_INVALIDARG)
                {
                    // Ignore invalid values when parsing a stylesheet
                    if(IsStyleSheetProperty())
                    {
                        goto Cleanup;
                    }

                    lNewValue = GetInvalidDefault();
                }
                else if(hr)
                {
                    goto Cleanup;
                }
            }
            // If we're just sniffing for a good parse - don't set up a default
            if(hr && ISSAMPLING(dwOpCode))
            {
                goto Cleanup;
            }
            fValidated = TRUE;
            pv = &lNewValue;
            break;

        case PROPTYPE_VARIANT:
            {
                VARIANT* pVar = (VARIANT*)pv;

                switch(pVar -> vt)
                {
                default:
                    hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pVar, VT_BSTR);
                    if(hr)
                    {
                        goto Cleanup;
                    }
                    pVar = &varDest;

                    // Fall thru to VT_BSTR code below

                case VT_BSTR:
                    pStr = pVar -> bstrVal;
                    goto FromString;

                case VT_NULL:
                    lNewValue = 0;
                    break;

                }
                pv = &lNewValue;
            }
            break;

        default:
            Assert(PROPTYPE(dwOpCode) == 0); // assumed native long
        }

        WORD wFlags = 0;
        if(dwOpCode & HANDLEPROP_IMPORTANT)
        {
            wFlags |= CAttrValue::AA_Extra_Important;
        }
        if(dwOpCode & HANDLEPROP_IMPLIED)
        {
            wFlags |= CAttrValue::AA_Extra_Implied;
        }
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_DEFAULT:
            // Set default Value from propdesc
            hr = ppp->SetNumber(pObject, pSubObject, (long)ulTagNotPresentDefault, wFlags);
            goto Cleanup;

        case HANDLEPROP_AUTOMATION:
            if(hr)
            {
                goto Cleanup;
            }

            // The unit value handler can be called directly through automation
            // This is pretty unique - we usualy have a helper fn
            hr = HandleNumProperty(SETPROPTYPE(dwOpCode, PROPTYPE_EMPTY),
                pv, pObject, pSubObject);
            if(hr)
            {
                goto Cleanup;
            }
            break;

        case HANDLEPROP_VALUE:
            if(!fValidated)
            {
                // We need to preserve an existing error code if there is one
                hr2 = ppp->ValidateNumberProperty((long*)pv, pObject);
                if(hr2)
                {
                    hr = hr2;
                    goto Cleanup;
                }
            }
            hr2 = ppp->SetNumber(pObject, pSubObject, *(long*)pv, wFlags);
            if(hr2)
            {
                hr = hr2;
                goto Cleanup;
            }

            if(dwOpCode & HANDLEPROP_MERGE)
            {
                hr2 = pObject->OnPropertyChange(bpp->dispid, bpp->dwFlags);
                if(hr2)
                {
                    hr = hr2;
                    goto Cleanup;
                }
            }

            break;
        }
        RRETURN1(hr, E_INVALIDARG);
    }
    else
    {
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_AUTOMATION:
            hr = ppp->GetNumberProperty(&lTemp, pObject, pSubObject);
            if(hr)
            {
                goto Cleanup;
            }
            // Coerce to type
            {
                BSTR* pbstr = (BSTR*)pv;

                switch(PROPTYPE(dwOpCode))
                {
                case PROPTYPE_VARIANT:
                    // pv points to a VARIANT
                    VariantInit((VARIANT*)pv);
                    V_VT((VARIANT*)pv) = VT_BSTR;
                    pbstr = &V_BSTR((VARIANT*)pv);
                    // intentional fall-through
                case PROPTYPE_BSTR:
                    {
                        TCHAR cValue[30];

                        CUnitValue uvThis;
                        uvThis.SetRawValue(lTemp);

                        if(uvThis.IsNull())
                        {
                            // Return an empty BSTR
                            hr = FormsAllocString(_afxGlobalData._Zero.ach, pbstr);
                        }
                        else
                        {
                            hr = uvThis.FormatBuffer((LPTSTR)cValue, (UINT)ARRAYSIZE(cValue), this);
                            if(hr)
                            {
                                goto Cleanup;
                            }

                            hr = FormsAllocString(cValue, pbstr);
                        }
                    }
                    break;

                default:
                    *(long*)pv = lTemp;
                    break;
                }
            }
            goto Cleanup;
            break;

        case HANDLEPROP_VALUE:
            {
                CUnitValue uvValue;

                // get raw value
                hr = ppp->GetNumber(pObject, pSubObject, (long*)&uvValue);

                if(PROPTYPE(dwOpCode) == VT_VARIANT)
                {
                    ((VARIANT*)pv)->vt = VT_I4;
                    ((VARIANT*)pv)->lVal = uvValue.GetUnitValue();
                }
                else if(PROPTYPE(dwOpCode) == PROPTYPE_BSTR)
                {
                    TCHAR cValue[30];

                    hr = uvValue.FormatBuffer((LPTSTR)cValue, (UINT)ARRAYSIZE(cValue), this);
                    if(hr)
                    {
                        goto Cleanup;
                    }

                    hr = FormsAllocString(cValue, (BSTR*)pv);
                }
                else
                {
                    Assert(PROPTYPE(dwOpCode) == 0);
                    *(long*)pv = uvValue.GetUnitValue();
                }
            }
            goto Cleanup;

        case HANDLEPROP_COMPARE:
            hr = ppp->GetNumber(pObject, pSubObject, &lTemp);
            if(hr)
            {
                goto Cleanup;
            }
            // if inherited and not set, return S_OK
            hr = (lTemp==*(long*)pv) ? S_OK : S_FALSE;
            RRETURN1(hr, S_FALSE);
            break;

        case HANDLEPROP_STREAM:
            {
                // Get value into an IStream
                // until the binary persistance we assume text save
                Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);

                hr = ppp->GetNumber(pObject, pSubObject, &lNewValue);
                if(hr)
                {
                    goto Cleanup;
                }

                CUnitValue uvValue;
                uvValue.SetRawValue(lNewValue);

                hr = uvValue.Persist((IStream*)pv, this);

                goto Cleanup;
            }
            break;
        }
    }
    // Let the basic numproperty handler handle all other cases
    hr = HandleNumProperty(dwOpCode, pv, pObject, pSubObject);

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}



//+---------------------------------------------------------------
//
//  Function:   GetNumber
//
//  Synopsis:   Helper for getting number values properties
//
//----------------------------------------------------------------
HRESULT NUMPROPPARAMS::GetNumber(CBase* pObject, CVoid* pSubObject, long* plNum, BOOL *pfValuePresent) const
{
    HRESULT hr = S_OK;

    if(bpp.dwPPFlags & PROPPARAM_GETMFHandler)
    {
        PFN_NUMPROPGET pmfGet;
        bpp.GetGetMethodP(this+1, &pmfGet);
        hr = CALL_METHOD((CVoid*)(void*)pObject, pmfGet, (plNum));
        if(pfValuePresent)
        {
            *pfValuePresent = TRUE;
        }
    }
    else
    {
        *plNum = bpp.GetAvNumber(pSubObject, this+1, cbMember, pfValuePresent);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Function:   SetNumber
//
//  Synopsis:   Helper for setting number values properties
//
//----------------------------------------------------------------
HRESULT NUMPROPPARAMS::SetNumber(CBase* pObject, CVoid* pSubObject, long lNum, WORD wFlags) const
{
    HRESULT hr = S_OK;

    if(bpp.dwPPFlags & PROPPARAM_SETMFHandler)
    {
        PFN_NUMPROPSET pmfSet;
        bpp.GetSetMethodP(this+1, &pmfSet);
        hr = CALL_METHOD((CVoid*)(void*)pObject, pmfSet, (lNum));
    }
    else
    {
        hr = bpp.SetAvNumber(pSubObject, (DWORD)lNum, this+1, cbMember, wFlags);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     ValidateNumberProperty, public
//
//  Synopsis:   Helper for testing the validity of numeric properties
//
//----------------------------------------------------------------
HRESULT NUMPROPPARAMS::ValidateNumberProperty(long* lArg, CBase* pObject) const
{
    long lMinValue=0, lMaxValue=LONG_MAX;
    HRESULT hr = S_OK;
    int ids = 0;
    const PROPERTYDESC* pPropDesc = (const PROPERTYDESC*)this - 1;

    if(vt==VT_BOOL || vt==VT_BOOL4)
    {
        return S_OK;
    }

    // Check validity of input argument
    if(bpp.dwPPFlags & PROPPARAM_UNITVALUE)
    {
        CUnitValue uv;
        uv.SetRawValue(*lArg);
        hr = uv.IsValid(pPropDesc);
        if(hr)
        {
            if(bpp.dwPPFlags & PROPPARAM_MINOUT)
            {
                // set value to min;
                *lArg = lMin;
                hr = S_OK;
            }
            else
            {
                //otherwise return error, i.e. set to default
                ids = IDS_ES_ENTER_PROPER_VALUE;
            }
        }
        goto Cleanup;
    }

    if(((bpp.dwPPFlags&PROPPARAM_ENUM) && (bpp.dwPPFlags&PROPPARAM_ANUMBER))
        || !(bpp.dwPPFlags&PROPPARAM_ENUM))
    {
        lMinValue = lMin; lMaxValue = lMax;
        if((*lArg<lMinValue || *lArg>lMaxValue) && 
            *lArg!=(long)pPropDesc->GetNotPresentDefault() &&
            *lArg!=(long)pPropDesc->GetInvalidDefault())
        {
            if(lMaxValue != LONG_MAX)
            {
                ids = IDS_ES_ENTER_VALUE_IN_RANGE;
            }
            else if(lMinValue == 0)
            {
                ids = IDS_ES_ENTER_VALUE_GE_ZERO;
            }
            else if(lMinValue == 1)
            {
                ids = IDS_ES_ENTER_VALUE_GT_ZERO;
            }
            else
            {
                ids = IDS_ES_ENTER_VALUE_IN_RANGE;
            }
            hr = E_INVALIDARG;
        }
        else
        {
            // inside of range
            goto Cleanup;
        }
    }

    // We have 3 scenarios to check for :-
    // 1) Just a number validate w min & max
    // 2) Just an emum validated w pEnumDesc
    // 3) A number w. min & max & one or more enum values

    // If we've got a number OR enum type, first check the
    if(bpp.dwPPFlags & PROPPARAM_ENUM)
    {
        ENUMDESC* pEnumDesc = pPropDesc->GetEnumDescriptor();

        // dwMask represent a mask allowing or rejecting values 0 to 31
        if(*lArg>=0 && *lArg<32)
        {
            if(!((pEnumDesc->dwMask >> (DWORD)*lArg)&1))
            {
                if(!hr)
                {
                    ids = IDS_ES_ENTER_PROPER_VALUE;
                    hr = E_INVALIDARG;
                }
            }
            else
            {
                hr = S_OK;
            }
        }
        else
        {
            // If it's not in the mask, look for the value in the array
            WORD i;

            for(i=0; i<pEnumDesc->cEnums; i++)
            {
                if(pEnumDesc->aenumpairs[i].iVal == *lArg)
                {
                    break;
                }
            }
            if(i == pEnumDesc->cEnums)
            {
                if(!hr)
                {
                    ids = IDS_ES_ENTER_PROPER_VALUE;
                    hr = E_INVALIDARG;
                }
            }
            else
            {
                hr = S_OK;
            }
        }
    }

Cleanup:
    if(hr && ids==IDS_ES_ENTER_PROPER_VALUE)
    {
        RRETURN1(pObject->SetErrorInfoPBadValue(bpp.dispid, ids), E_INVALIDARG);
    }
    else if(hr)
    {
        RRETURN1(pObject->SetErrorInfoPBadValue(
            bpp.dispid,
            ids,
            lMinValue,
            lMaxValue), E_INVALIDARG);
    }
    else
    {
        return S_OK;
    }
}

//+---------------------------------------------------------------
//
//  Member:     CBase::GetNumberProperty, public
//
//  Synopsis:   Helper for getting number values properties
//
//----------------------------------------------------------------
HRESULT NUMPROPPARAMS::GetNumberProperty(void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;
    long num;

    if(!pv)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    hr = GetNumber(pObject, pSubObject, &num);
    if(hr)
    {
        goto Cleanup;
    }

    SetNumberOfType(pv, VARENUM(vt), num);

Cleanup:
    RRETURN(pObject->SetErrorInfoPGet(hr, bpp.dispid));
}

//+---------------------------------------------------------------
//
//  Member:     SetNumberProperty, public
//
//  Synopsis:   Helper for setting number values properties
//
//----------------------------------------------------------------
HRESULT NUMPROPPARAMS::SetNumberProperty(long lValueNew, CBase* pObject, CVoid* pSubObject, BOOL fValidate /*=TRUE*/, WORD wFlags) const
{
    CBase::CLock Lock(pObject);
    HRESULT hr;
    long    lValueOld;
    BOOL    fOldPresent;

    // Get the current value of the property
    if(!pSubObject)
    {
        pSubObject = pObject;
    }

    hr = GetNumber(pObject, pSubObject, &lValueOld, &fOldPresent);
    if(hr)
    {
        goto Cleanup;
    }

    // Validate the old value, just for kicks...

    // If the old and new values are the same, then git out of here
    if(lValueNew == lValueOld)
    {
        goto Cleanup;
    }

    // Make sure the new value is OK
    if(fValidate)
    {
        // Validate using propdesc encoded parser rules 
        hr = ValidateNumberProperty(&lValueNew, pObject);
        if(hr)
        {
            return hr; // Error info set in validate
        }
    }

    // Stuff the new value in
    hr = SetNumber(pObject, pSubObject, lValueNew, wFlags);
    if(hr)
    {
        goto Cleanup;
    }

    // Tell everybody about the new value
    hr = pObject->OnPropertyChange(bpp.dispid, bpp.dwFlags);

    // If anybody complained, then revert

    // BUGBUG: should we delte the undo thing-a-ma-jig if the OnPropChange fails?
    if(hr)
    {
        SetNumber(pObject, pSubObject, lValueOld, wFlags);
        pObject->OnPropertyChange(bpp.dispid, bpp.dwFlags);
    }

Cleanup:
    RRETURN(pObject->SetErrorInfoPSet(hr, bpp.dispid));
}

HRESULT NUMPROPPARAMS::SetEnumStringProperty(BSTR bstrNew, CBase* pObject, CVoid* pSubObject, WORD wFlags) const
{
    HRESULT hr = E_INVALIDARG;
    CBase::CLock Lock(pObject);
    long lNewValue, lOldValue;
    BOOL fOldPresent;

    hr = GetNumber(pObject, pSubObject, &lOldValue, &fOldPresent);
    if(hr)
    {
        goto Cleanup;
    }

    hr = LookupEnumString(this, (LPTSTR)bstrNew, &lNewValue);
    if(hr)
    {
        goto Cleanup;
    }

    if(lNewValue == lOldValue)
    {
        goto Cleanup;
    }

    hr = ValidateNumberProperty(&lNewValue, pObject);
    if(hr)
    {
        goto Cleanup;
    }

// wlw note
//#ifndef NO_EDIT
//    if(pObject->QueryCreateUndo(TRUE, FALSE))
//    {
//        CVariant varOld;
//
//        if(fOldPresent)
//        {
//            ENUMDESC* pEnumDesc;
//            const TCHAR* pchOldValue;
//
//            if(bpp.dwPPFlags & PROPPARAM_ANUMBER)
//            {
//                pEnumDesc = *(ENUMDESC**)((BYTE*)(this+1) + sizeof(DWORD_PTR));
//            }
//            else
//            {
//                pEnumDesc = (ENUMDESC*)lMax;
//            }
//
//            pchOldValue = pEnumDesc->StringPtrFromEnum(lOldValue);
//
//            V_VT(&varOld) = VT_BSTR;
//            hr = FormsAllocString(pchOldValue, &V_BSTR(&varOld));
//            if(hr)
//            {
//                goto Cleanup;
//            }
//        }
//        else
//        {
//            V_VT(&varOld) = VT_NULL;
//        }
//
//        hr = pObject->CreatePropChangeUndo(bpp.dispid, &varOld, NULL);
//        if(hr)
//        {
//            goto Cleanup;
//        }
//    }
//#endif // NO_EDIT

    hr = SetNumber(pObject, pSubObject, lNewValue, wFlags);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pObject->OnPropertyChange(bpp.dispid, bpp.dwFlags);

Cleanup:
    RRETURN(pObject->SetErrorInfoPGet(hr, bpp.dispid));
}

HRESULT NUMPROPPARAMS::GetEnumStringProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;
    VARIANT varValue;
    PROPERTYDESC* ppdPropDesc = ((PROPERTYDESC*)this) - 1;

    if(!pbstr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(!pSubObject)
    {
        pSubObject = pObject;
    }


    hr = ppdPropDesc->HandleNumProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16), 
        (void*)&varValue, pObject, pSubObject);
    if(!hr)
    {
        *pbstr = V_BSTR(&varValue);
    }

Cleanup:
    RRETURN(pObject->SetErrorInfoPGet(hr, bpp.dispid));
}



STDMETHODIMP CBase::put_StyleComponent(BSTR v)
{
    GET_THUNK_PROPDESC;
    return put_StyleComponentHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_StyleComponentHelper(BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return pPropDesc->GetBasicPropParams()->SetStyleComponentProperty(v, this, ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::put_Url(BSTR v)
{
    GET_THUNK_PROPDESC;
    return put_UrlHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_UrlHelper(BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return pPropDesc->GetBasicPropParams()->SetUrlProperty(v, this,
        ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::put_String(BSTR v)
{
    GET_THUNK_PROPDESC;
    return put_StringHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_StringHelper(BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    CVoid* pSubObj = (pPropDesc->GetBasicPropParams()->dwPPFlags&PROPPARAM_ATTRARRAY) ?
        (ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray())) : CVOID_CAST(this);

    switch(pPropDesc->GetBasicPropParams()->wInvFunc)
    {
    case IDX_GS_BSTR:
        return pPropDesc->GetBasicPropParams()->SetStringProperty(v, this, pSubObj);
    case IDX_GS_PropEnum:
        return pPropDesc->GetNumPropParams()->SetEnumStringProperty(v, this, pSubObj);
    }
    return S_OK;
}

STDMETHODIMP CBase::put_Short(short v)
{
    GET_THUNK_PROPDESC;
    return put_ShortHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_ShortHelper(short v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return pPropDesc->GetNumPropParams()->SetNumberProperty(v, this, ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::put_Long(long v)
{
    GET_THUNK_PROPDESC;
    return put_LongHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_LongHelper(long v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    CVoid* pSubObj = (pPropDesc->GetBasicPropParams()->dwPPFlags&PROPPARAM_ATTRARRAY) ?
        (ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray())) : CVOID_CAST(this);

    return pPropDesc->GetNumPropParams()->SetNumberProperty(v, this, pSubObj);
}

STDMETHODIMP CBase::put_Bool(VARIANT_BOOL v)
{
    GET_THUNK_PROPDESC;
    return put_BoolHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_BoolHelper(VARIANT_BOOL v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    CVoid* pSubObj = (pPropDesc->GetBasicPropParams()->dwPPFlags&PROPPARAM_ATTRARRAY) ?
        (ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray())) : CVOID_CAST(this);

    return pPropDesc->GetNumPropParams()->SetNumberProperty(v, this, pSubObj);
}

STDMETHODIMP CBase::put_Variant(VARIANT v)
{
    GET_THUNK_PROPDESC;
    return put_VariantHelper(v, pPropDesc);
}

STDMETHODIMP CBase::put_VariantHelper(VARIANT v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return (pPropDesc->*pPropDesc->pfnHandleProperty)(
        HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, this, ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::get_Url(BSTR* p)
{
    GET_THUNK_PROPDESC;
    return get_UrlHelper(p, pPropDesc);
}

STDMETHODIMP CBase::get_UrlHelper(BSTR* p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return pPropDesc->GetBasicPropParams()->GetUrlProperty(p, this,
        ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::get_StyleComponent(BSTR* p)
{
    GET_THUNK_PROPDESC;
    return get_StyleComponentHelper(p, pPropDesc);
}

STDMETHODIMP CBase::get_StyleComponentHelper(BSTR* p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    return pPropDesc->GetBasicPropParams()->GetStyleComponentProperty(p, this,
        ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CBase::get_Property(void* p)
{
    GET_THUNK_PROPDESC;
    return get_PropertyHelper(p, pPropDesc);
}

STDMETHODIMP CBase::get_PropertyHelper(void* p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr)
{
    Assert(pPropDesc);

    CVoid* pSubObj = (pPropDesc->GetBasicPropParams()->dwPPFlags&PROPPARAM_ATTRARRAY) ?
        (ppAttr?CVOID_CAST(ppAttr):CVOID_CAST(GetAttrArray())) : CVOID_CAST(this);

    switch(pPropDesc->GetBasicPropParams()->wInvFunc)
    {
    case IDX_G_VARIANT:
    case IDX_GS_VARIANT:
        return (pPropDesc->*pPropDesc->pfnHandleProperty)(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16), p, this, pSubObj);
    case IDX_G_VARIANTBOOL:
    case IDX_GS_VARIANTBOOL:
    case IDX_G_long:
    case IDX_GS_long:
    case IDX_G_short:
    case IDX_GS_short:
        return pPropDesc->GetNumPropParams()->GetNumberProperty(p, this, pSubObj);
    case IDX_G_BSTR:
    case IDX_GS_BSTR:
        return pPropDesc->GetBasicPropParams()->GetStringProperty((BSTR*)p, this, pSubObj);
    case IDX_G_PropEnum:
    case IDX_GS_PropEnum:
        return pPropDesc->GetNumPropParams()->GetEnumStringProperty((BSTR*)p, this, pSubObj);
    }
    return S_OK;
}
