
#include "stdafx.h"

//+------------------------------------------------------------------------
//
//  Function:   ::ParseTextDecorationProperty
//
//  Synopsis:   Parses a text-decoration string in CSS format and sets the
//              appropriate sub-properties.
//
//-------------------------------------------------------------------------
HRESULT ParseTextDecorationProperty(CAttrArray** ppAA, LPCTSTR pcszTextDecoration, WORD wFlags)
{
    TCHAR* pszString;
    TCHAR* pszCopy;
    TCHAR* pszNextToken;
    HRESULT hr = S_OK;
    BOOL fInvalidValues = FALSE;
    BOOL fInsideParens;
    VARIANT v;

    Assert(ppAA && "No (CAttrArray*) pointer!");

    if(!pcszTextDecoration)
    {
        pcszTextDecoration = _T("");
    }

    v.vt = VT_I4;
    v.lVal = 0;

    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszTextDecoration)+1];
    if(!pszCopy)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    _tcscpy(pszCopy,  pcszTextDecoration);

    // Loop through the tokens in the string (parenthesis parsing is for future
    // text-decoration values that might have parameters).
    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        fInsideParens = FALSE;
        while(_istspace(*pszString))
        {
            pszString++;
        }

        while(*pszNextToken && (fInsideParens || !_istspace(*pszNextToken)))
        {
            if(*pszNextToken == _T('('))
            {
                fInsideParens = TRUE;
            }
            if(*pszNextToken == _T(')'))
            {
                fInsideParens = FALSE;
            }
            pszNextToken++;
        }

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        if(!StrCmpIC(pszString, _T("none")))
        {
            v.lVal = TD_NONE; // "none" clears all the other properties (unlike the other properties)
        }
        else if(!StrCmpIC( pszString, _T("underline")))
        {
            v.lVal |= TD_UNDERLINE;
        }
        else if(!StrCmpIC( pszString, _T("overline")))
        {
            v.lVal |= TD_OVERLINE;
        }
        else if(!StrCmpIC( pszString, _T("line-through")))
        {
            v.lVal |= TD_LINETHROUGH;
        }
        else if(!StrCmpIC( pszString, _T("blink")))
        {
            v.lVal |= TD_BLINK;
        }
        else
        {
            fInvalidValues = TRUE;
        }
    }

    hr = CAttrArray::Set(ppAA, DISPID_A_TEXTDECORATION, &v,
        (PROPERTYDESC*)&s_propdescCStyletextDecoration, CAttrValue::AA_StyleAttribute, wFlags);

Cleanup:
    delete[] pszCopy;
    RRETURN1(fInvalidValues?(hr?hr:E_INVALIDARG):hr, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:   ::ParseTextAutospaceProperty
//
//  Synopsis:   Parses a text-autospace string in CSS format and sets the
//              appropriate sub-properties.
//
//-------------------------------------------------------------------------
HRESULT ParseTextAutospaceProperty(CAttrArray** ppAA, LPCTSTR pcszTextAutospace, WORD wFlags)
{
    TCHAR* pszTokenBegin;
    TCHAR* pszCopy;
    TCHAR* pszTokenEnd;
    HRESULT hr = S_OK;
    BOOL fInvalidValues = FALSE;
    VARIANT v;

    Assert(ppAA && "No (CAttrArray*) pointer!");

    if(!pcszTextAutospace)
    {
        pcszTextAutospace = _T("");
    }

    v.vt = VT_I4;
    v.lVal = 0;

    pszCopy = pszTokenBegin = pszTokenEnd = new TCHAR[_tcslen(pcszTextAutospace)+1];
    if(!pszCopy)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    _tcscpy(pszCopy,  pcszTextAutospace);

    for(pszTokenBegin=pszTokenEnd=pszCopy;
        pszTokenBegin&&*pszTokenBegin; 
        pszTokenBegin=pszTokenEnd)
    {
        while(_istspace(*pszTokenBegin))
        {
            pszTokenBegin++;
        }

        pszTokenEnd = pszTokenBegin;

        while(*pszTokenEnd && !_istspace(*pszTokenEnd))
        {
            pszTokenEnd++;
        }

        if(*pszTokenEnd)
        {
            *pszTokenEnd++ = _T('\0');
        }

        if(StrCmpIC(pszTokenBegin, _T("ideograph-numeric")) == 0)
        {
            v.lVal |= TEXTAUTOSPACE_NUMERIC;
        }
        else if(StrCmpIC(pszTokenBegin, _T("ideograph-space")) == 0)
        {
            v.lVal |= TEXTAUTOSPACE_SPACE;
        }
        else if(StrCmpIC(pszTokenBegin, _T("ideograph-alpha")) == 0)
        {
            v.lVal |= TEXTAUTOSPACE_ALPHA;
        }
        else if(StrCmpIC(pszTokenBegin, _T("ideograph-parenthesis")) == 0)
        {
            v.lVal |= TEXTAUTOSPACE_PARENTHESIS;
        }
        else if(StrCmpIC(pszTokenBegin, _T("none")) == 0)
        {
            v.lVal = TEXTAUTOSPACE_NONE;
        }
        else
        {
            fInvalidValues = TRUE;
        }
    }

    hr = CAttrArray::Set(ppAA, DISPID_A_TEXTAUTOSPACE, &v,
        (PROPERTYDESC*)&s_propdescCStyletextAutospace, CAttrValue::AA_StyleAttribute, wFlags);

Cleanup:
    delete[] pszCopy;
    RRETURN1(fInvalidValues?(hr?hr:E_INVALIDARG):hr, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:   ::ParseListStyleProperty
//
//  Synopsis:   Parses a list-style string in CSS format and sets the
//              appropriate sub-properties.
//
//-------------------------------------------------------------------------
HRESULT ParseListStyleProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszListStyle)
{
    TCHAR* pszString;
    TCHAR* pszCopy;
    TCHAR* pszNextToken;
    HRESULT hrResult = S_OK;
    BOOL fInsideParens;

    if(!pcszListStyle || !*pcszListStyle)
    {
        if(*ppAA)
        {
            (*ppAA)->FindSimpleAndDelete(DISPID_A_LISTSTYLETYPE, CAttrValue::AA_StyleAttribute);
            (*ppAA)->FindSimpleAndDelete(DISPID_A_LISTSTYLEPOSITION, CAttrValue::AA_StyleAttribute);
            (*ppAA)->FindSimpleAndDelete(DISPID_A_LISTSTYLEIMAGE, CAttrValue::AA_StyleAttribute);
        }
        return S_OK;
    }

    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszListStyle)+1];
    if(!pszCopy)
    {
        hrResult = E_OUTOFMEMORY;
        goto Cleanup;
    }

    _tcscpy(pszCopy, pcszListStyle);

    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        fInsideParens = FALSE;
        while(_istspace(*pszString))
        {
            pszString++;
        }

        while(*pszNextToken && (fInsideParens || !_istspace(*pszNextToken)))
        {
            if(*pszNextToken == _T('('))
            {
                fInsideParens = TRUE;
            }
            if(*pszNextToken == _T(')'))
            {
                fInsideParens = FALSE;
            }
            pszNextToken++;
        }

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        // Try type
        if(s_propdescCStylelistStyleType.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA))
        {
            // Failed: try position
            if(s_propdescCStylelistStylePosition.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA))
            {
                // Failed: try image
                if(s_propdescCStylelistStyleImage.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA))
                {
                    hrResult = E_INVALIDARG;
                }
            }
        }
    }
Cleanup:
    delete[] pszCopy;
    RRETURN1(hrResult, E_INVALIDARG);
}

///+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleStyleProperty, public
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
HRESULT PROPERTYDESC::HandleStyleProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
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
                hr = VariantChangeTypeSpecial(&varDest, (VARIANT*)pv,  VT_BSTR);
                if(hr)
                {
                    goto Cleanup;
                }
                pv = V_BSTR(&varDest);
            }
        default:
            Assert(PROPTYPE(dwOpCode) == 0); // assumed native long
            break;
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
        case HANDLEPROP_AUTOMATION:
            if(pv && *(TCHAR*)pv)
            {
                AssertSz(FALSE, "must improve");
            }
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
            if(PROPTYPE(dwOpCode) == PROPTYPE_VARIANT)
            {
                hr = WriteStyleToBSTR(pObject, 
                    ((CElement*)pObject)->GetInLineStyleAttrArray(), 
                    &((VARIANT*)pv)->bstrVal, TRUE);
                ((VARIANT*)pv)->vt = VT_BSTR;
            }
            else if(PROPTYPE(dwOpCode) == PROPTYPE_BSTR)
            {
                hr = WriteStyleToBSTR(pObject, 
                    ((CElement*)pObject)->GetInLineStyleAttrArray(), 
                    (BSTR*)pv, TRUE);
            }
            else
            {
                BSTR bstr = NULL;

                Assert(PROPTYPE(dwOpCode) == 0);

                hr = WriteStyleToBSTR(pObject, 
                    ((CElement*)pObject)->GetInLineStyleAttrArray(), 
                    &bstr, TRUE);
                if(!hr)
                {
                    hr = ((CString*)pv)->Set(bstr);
                    FormsFreeString(bstr);
                }
            }
            if(hr)
            {
                goto Cleanup;
            }

            break;

        case HANDLEPROP_STREAM:
            Assert(PROPTYPE(dwOpCode) == PROPTYPE_LPWSTR);
            {
                BSTR bstrTemp;
                hr = WriteStyleToBSTR(pObject, ((CElement*)pObject)->GetInLineStyleAttrArray(), &bstrTemp, TRUE);
                if(!hr)
                {
                    if(*bstrTemp)
                    {
                        hr = ((IStream*)(void*)pv)->Write(bstrTemp, _tcslen(bstrTemp)*sizeof(TCHAR), NULL);
                    }
                    FormsFreeString(bstrTemp);
                }
            }

            if(hr)
            {
                goto Cleanup;
            }
            break;

        case HANDLEPROP_COMPARE:
            {
                CElement* pElem = (CElement*)pObject;
                CAttrArray* pAA = pElem->GetInLineStyleAttrArray();
                // Check for presence of Attributs and Expandos
                hr = pAA&&pAA->HasAnyAttribute(TRUE) ? S_FALSE : S_OK;
            }
            break;

        default:
            hr = E_FAIL;
            Assert(FALSE && "Invalid operation");
            break;
        }
    }

Cleanup:
    if(varDest.vt != VT_EMPTY)
    {
        VariantClear(&varDest);
    }
    RRETURN1(hr, S_FALSE);
}

HRESULT BASICPROPPARAMS::GetStyleComponentProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const
{
    VARIANT varValue;
    HRESULT hr;
    PROPERTYDESC* ppdPropDesc = ((PROPERTYDESC*)this) - 1;

    hr = ppdPropDesc->HandleStyleComponentProperty(HANDLEPROP_VALUE|(PROPTYPE_VARIANT<<16), 
        (void*)&varValue, pObject, pSubObject);
    if(!hr)
    {
        *pbstr = V_BSTR(&varValue);
    }

    RRETURN(hr);
}

HRESULT BASICPROPPARAMS::SetStyleComponentProperty(BSTR bstr, CBase* pObject, CVoid* pSubObject, WORD wFlags) const
{
    HRESULT hr;
    CBase::CLock Lock(pObject);
    DWORD dwOpCode = HANDLEPROP_SET|HANDLEPROP_AUTOMATION;
    PROPERTYDESC* ppdPropDesc = ((PROPERTYDESC*)this) - 1;

    if(wFlags & CAttrValue::AA_Extra_Important)
    {
        dwOpCode |= HANDLEPROP_IMPORTANT;
    }
    if(wFlags & CAttrValue::AA_Extra_Implied)
    {
        dwOpCode |= HANDLEPROP_IMPLIED;
    }

    SETPROPTYPE(dwOpCode, PROPTYPE_LPWSTR);

    hr = ppdPropDesc->HandleStyleComponentProperty(dwOpCode, (void*)bstr, pObject, pSubObject);

    RRETURN1(hr, E_INVALIDARG);
}

HRESULT BASICPROPPARAMS::GetStyleComponentBooleanProperty(VARIANT_BOOL* p, CBase* pObject, CVoid* pSubObject) const
{
    VARIANT varValue;
    HRESULT hr;
    PROPERTYDESC* ppdPropDesc = ((PROPERTYDESC*)this) - 1;

    hr = ppdPropDesc->HandleStyleComponentProperty(HANDLEPROP_VALUE|(PROPTYPE_VARIANT<<16), 
        (void*)&varValue, pObject, pSubObject);
    if(!hr)
    {
        Assert(varValue.vt == VT_BOOL);
        *p = varValue.boolVal;
    }
    RRETURN(hr);
}

HRESULT BASICPROPPARAMS::SetStyleComponentBooleanProperty(VARIANT_BOOL v, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr;
    CBase::CLock Lock(pObject);
    DWORD dwOpCode = HANDLEPROP_SET|HANDLEPROP_AUTOMATION;
    PROPERTYDESC* ppdPropDesc = ((PROPERTYDESC*)this) - 1;
    VARIANT var;

    var.vt = VT_BOOL;
    var.boolVal = v;

    SETPROPTYPE(dwOpCode, PROPTYPE_VARIANT);

    hr = ppdPropDesc->HandleStyleComponentProperty(dwOpCode, (void*)&var, pObject, pSubObject);
    RRETURN1(hr, E_INVALIDARG);
}

//+---------------------------------------------------------------
//
//  Member:     PROPERTYDESC::HandleStyleComponentProperty, public
//
//  Synopsis:   Helper for getting/setting url style sheet properties...
//              url(string)
//
//  Arguments:  dwOpCode        -- encodes the incoming type (PROPTYPE_FLAG) in the upper WORD and
//                                 the opcode in the lower WORD (HANDLERPROP_FLAGS)
//                                 PROPTYPE_EMPTY means the 'native' type (long in this case)
//              pv              -- points to the 'media' the value is stored for the get and set
//              pObject         -- object owns the property
//              pSubObject      -- subobject storing the property (could be the main object)
//
//----------------------------------------------------------------
HRESULT PROPERTYDESC::HandleStyleComponentProperty(DWORD dwOpCode, void* pv, CBase* pObject, CVoid* pSubObject) const
{
    HRESULT hr = S_OK;
    VARIANT varDest;
    size_t nLenIn = (size_t) -1;
    DWORD dispid = GetBasicPropParams()->dispid;
    BSTR bstrTemp; // Used by some of the stream writers
    BOOL fTDPropertyValue=FALSE; // If this is a SET of a text-decoration sub-property, this is the value
    WORD wFlags = 0;

    if(ISSET(dwOpCode))
    {
        Assert(!(ISSTREAM(dwOpCode))); // we can't do this yet...
        switch(dispid)
        {
        case DISPID_A_TEXTDECORATIONNONE:
        case DISPID_A_TEXTDECORATIONUNDERLINE:
        case DISPID_A_TEXTDECORATIONOVERLINE:
        case DISPID_A_TEXTDECORATIONLINETHROUGH:
        case DISPID_A_TEXTDECORATIONBLINK:
            Assert(PROPTYPE(dwOpCode)==PROPTYPE_VARIANT && "Text-decoration subproperties must take variants!");
            Assert(V_VT((VARIANT*)pv)==VT_BOOL && "Text-decoration subproperties must take BOOLEANs!");
            fTDPropertyValue = !!((VARIANT*)pv)->boolVal;
            break;

        default:
            switch(PROPTYPE(dwOpCode))
            {
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

                //intentional fallthrough
            case PROPTYPE_LPWSTR:
                switch(dispid)
                {
                case DISPID_A_BACKGROUNDIMAGE:
                case DISPID_A_LISTSTYLEIMAGE:
                case DISPID_A_FONTFACESRC:
                    nLenIn = ValidStyleUrl((TCHAR*)pv);
                    if(OPCODE(dwOpCode) == HANDLEPROP_VALUE)
                    {
                        if(!nLenIn && _tcsicmp((TCHAR*)pv, _T("none")))
                        {
                            hr = E_INVALIDARG;
                            goto Cleanup;
                        }
                    }
                    break;

                case DISPID_A_BEHAVIOR:
                    nLenIn = pv ? _tcslen((TCHAR*)pv) : 0;
                    break;
                }
                break;

            default:
                Assert(FALSE); // We shouldn't get here.
            }
            break;
        }

        switch(dispid)
        {
        case DISPID_A_LISTSTYLEIMAGE:
        case DISPID_A_BACKGROUNDIMAGE:
        case DISPID_A_FONTFACESRC:
        case DISPID_A_BEHAVIOR:
            if(nLenIn && (nLenIn!=(size_t)-1))
            {
                if(DISPID_A_BEHAVIOR == dispid)
                {
                    hr = HandleStringProperty(dwOpCode, (TCHAR*)pv, pObject, pSubObject);
                }
                else
                {
                    TCHAR* pch = (TCHAR*)pv;
                    TCHAR* psz = pch + 4;
                    TCHAR* quote = NULL;
                    TCHAR* pszEnd;
                    TCHAR terminator;

                    while(_istspace(*psz))
                    {
                        psz++;
                    }
                    if(*psz==_T('\'') || *psz==_T('"'))
                    {
                        quote = psz++;
                    }
                    nLenIn--; // Skip back over the ')' character - we know there is one, because ValidStyleUrl passed this string.
                    pszEnd = pch + nLenIn - 1;
                    while(_istspace(*pszEnd) && (pszEnd>psz))
                    {
                        pszEnd--;
                    }
                    if(quote && (*pszEnd==*quote))
                    {
                        pszEnd--;
                    }
                    terminator = *(pszEnd+1);
                    *(pszEnd+1) = _T('\0');
                    hr = HandleStringProperty(dwOpCode, psz, pObject, pSubObject);
                    *(pszEnd+1) = terminator;
                }
            }
            else
            {
                if(!pv || !*(TCHAR*)pv)
                {
                    // Empty string - delete the entry.
                    CAttrArray** ppAA = (CAttrArray**)pSubObject;

                    if(*ppAA)
                    {
                        (*ppAA)->FindSimpleAndDelete(dispid, CAttrValue::AA_StyleAttribute, NULL);
                    }
                }
                else if(!_tcsicmp((TCHAR*)pv, _T("none")))
                {
                    hr = HandleStringProperty(dwOpCode, (void*)_T(""), pObject, pSubObject);
                }
                else
                {
                    hr = E_INVALIDARG;
                }
            }
            break;

        case DISPID_A_BACKGROUND:
            hr = ParseBackgroundProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode),
                (TCHAR*)pv, (OPCODE(dwOpCode)==HANDLEPROP_VALUE));
            break;
        case DISPID_A_FONT:
            if(pv && FindSystemFontByName((TCHAR*)pv)!=sysfont_non_system)
            {
                hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
            }
            else
            {
                hr = ParseFontProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            }
            break;
        case DISPID_A_LAYOUTGRID:
            hr = ParseLayoutGridProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            break;
        case DISPID_A_TEXTAUTOSPACE:
            if(dwOpCode & HANDLEPROP_IMPORTANT)
            {
                wFlags |= CAttrValue::AA_Extra_Important;
            }
            if(dwOpCode & HANDLEPROP_IMPLIED)
            {
                wFlags |= CAttrValue::AA_Extra_Implied;
            }

            hr = ParseTextAutospaceProperty((CAttrArray**)pSubObject, (LPCTSTR)pv, wFlags);
            break;
        case DISPID_A_TEXTDECORATION:
            if(dwOpCode & HANDLEPROP_IMPORTANT)
            {
                wFlags |= CAttrValue::AA_Extra_Important;
            }
            if(dwOpCode & HANDLEPROP_IMPLIED)
            {
                wFlags |= CAttrValue::AA_Extra_Implied;
            }

            hr = ParseTextDecorationProperty((CAttrArray**)pSubObject, (LPCTSTR)pv, wFlags);
            break;
        case DISPID_A_TEXTDECORATIONNONE:
        case DISPID_A_TEXTDECORATIONUNDERLINE:
        case DISPID_A_TEXTDECORATIONOVERLINE:
        case DISPID_A_TEXTDECORATIONLINETHROUGH:
        case DISPID_A_TEXTDECORATIONBLINK:
            {
                VARIANT v;

                v.vt = VT_I4;
                v.lVal = 0;
                if(*((CAttrArray**)pSubObject))
                {
                    // See if we already have a text-decoration value
                    CAttrValue* pAV = (*((CAttrArray**)pSubObject))->Find(DISPID_A_TEXTDECORATION, CAttrValue::AA_Attribute);
                    if(pAV)
                    {
                        // We do!  Copy its value into our working variant
                        v.lVal = pAV->GetLong();
                    }
                }
                switch(dispid)
                {
                case DISPID_A_TEXTDECORATIONNONE:
                    if(fTDPropertyValue)
                    {
                        v.lVal = TD_NONE; // "none" clears all the other properties (unlike the other properties)
                    }
                    else
                    {
                        v.lVal &= ~TD_NONE;
                    }
                    break;
                case DISPID_A_TEXTDECORATIONUNDERLINE:
                    if(fTDPropertyValue)
                    {
                        v.lVal |= TD_UNDERLINE;
                    }
                    else
                    {
                        v.lVal &= ~TD_UNDERLINE;
                    }
                    break;
                case DISPID_A_TEXTDECORATIONOVERLINE:
                    if(fTDPropertyValue)
                    {
                        v.lVal |= TD_OVERLINE;
                    }
                    else
                    {
                        v.lVal &= ~TD_OVERLINE;
                    }
                    break;
                case DISPID_A_TEXTDECORATIONLINETHROUGH:
                    if(fTDPropertyValue)
                    {
                        v.lVal |= TD_LINETHROUGH;
                    }
                    else
                    {
                        v.lVal &= ~TD_LINETHROUGH;
                    }
                    break;
                case DISPID_A_TEXTDECORATIONBLINK:
                    if(fTDPropertyValue)
                    {
                        v.lVal |= TD_BLINK;
                    }
                    else
                    {
                        v.lVal &= ~TD_BLINK;
                    }
                    break;
                }
                if(dwOpCode & HANDLEPROP_IMPORTANT)
                {
                    wFlags |= CAttrValue::AA_Extra_Important;
                }
                if ( dwOpCode & HANDLEPROP_IMPLIED)
                {
                    wFlags |= CAttrValue::AA_Extra_Implied;
                }
                hr = CAttrArray::Set((CAttrArray**)pSubObject, DISPID_A_TEXTDECORATION, &v,
                    (PROPERTYDESC*)&s_propdescCStyletextDecoration, CAttrValue::AA_StyleAttribute, wFlags);
            }
            dispid = DISPID_A_TEXTDECORATION; // This is so we call OnPropertyChange for the right property below.
            break;

        case DISPID_A_MARGIN:
        case DISPID_A_PADDING:
            hr = ParseExpandProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv, dispid, TRUE);
            break;

        case DISPID_A_BORDERCOLOR:
        case DISPID_A_BORDERWIDTH:
        case DISPID_A_BORDERSTYLE:
            hr = ParseExpandProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv, dispid, FALSE);
            break;

        case DISPID_A_STYLETEXT:
            {
                LPTSTR lpszStyleText = (TCHAR*)pv;
                CAttrArray** ppAA = (CAttrArray**)pSubObject;

                if(*ppAA)
                {
                    (*ppAA)->Free();
                }

                if(lpszStyleText && *lpszStyleText)
                {
                    CStyle* pStyle = DYNCAST(CStyle, pObject);

                    Assert(pStyle);
                    pStyle->MaskPropertyChanges(TRUE);
                    AssertSz(FALSE, "must improve");
                    pStyle->MaskPropertyChanges(FALSE);
                }
            }
            break;

        case DISPID_A_BORDERTOP:
        case DISPID_A_BORDERRIGHT:
        case DISPID_A_BORDERBOTTOM:
        case DISPID_A_BORDERLEFT:
            hr = ParseAndExpandBorderSideProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv, dispid);
            break;

        case DISPID_A_BORDER:
            hr = ParseBorderProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            break;

        case DISPID_A_LISTSTYLE:
            hr = ParseListStyleProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            break;

        case DISPID_A_BACKGROUNDPOSITION:
            hr = ParseBackgroundPositionProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            break;

        case DISPID_A_CLIP:
            hr = ParseClipProperty((CAttrArray**)pSubObject, pObject, OPERATION(dwOpCode), (TCHAR*)pv);
            break;

        default:
            Assert("Attempting to set an unknown type of CStyleComponent!");
        }

        if(hr)
        {
            goto Cleanup;
        }
        else
        {
            // Note that dispid reflects the property that changed, not what was set -
            // e.g., textDecorationUnderline has been changed to textDecoration.
            if(dwOpCode & HANDLEPROP_AUTOMATION)
            {
                CBase::CLock Lock(pObject);
                hr = pObject->OnPropertyChange(dispid, GetdwFlags());
            }
        }
    }
    else
    {
        // GET value from data
        switch(OPCODE(dwOpCode))
        {
        case HANDLEPROP_STREAM:
            {
                IStream* pis = (IStream*)pv;

                switch(dispid)
                {
                case DISPID_A_LISTSTYLEIMAGE:
                case DISPID_A_BACKGROUNDIMAGE:
                case DISPID_A_BEHAVIOR:
                case DISPID_A_FONTFACESRC:
                    if((*(CAttrArray**)pSubObject)->Find(dispid, CAttrValue::AA_Attribute))
                    {
                        BSTR bstrSub;

                        hr = HandleStringProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_BSTR<<16), 
                            &bstrSub, pObject, pSubObject);
                        if(hr == S_OK)
                        {
                            if(bstrSub && *bstrSub)
                            {   // This is a normal url.
                                hr = pis->Write(_T("url("), 4*sizeof(TCHAR), NULL);
                                if(!hr)
                                {
                                    hr = pis->Write(bstrSub, FormsStringLen(bstrSub)*sizeof(TCHAR), NULL);
                                    if(!hr)
                                    {
                                        hr = pis->Write(_T(")"), 1*sizeof(TCHAR), NULL);
                                    }
                                }
                            }
                            else
                            {   // We only get here if a NULL string was stored in the array; i.e., the value is NONE.
                                hr = pis->Write(_T("none"), 4*sizeof(TCHAR), NULL);
                            }
                            FormsFreeString(bstrSub);
                        }
                    }
                    break;

                case DISPID_A_BACKGROUND:
                    hr = WriteBackgroundStyleToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_TEXTAUTOSPACE:
                    // We need to cook up this property.
                    hr = WriteTextAutospaceToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_TEXTDECORATION:
                    // We need to cook up this property.
                    hr = WriteTextDecorationToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_BORDERTOP:
                case DISPID_A_BORDERRIGHT:
                case DISPID_A_BORDERBOTTOM:
                case DISPID_A_BORDERLEFT:
                    hr = WriteBorderSidePropertyToBSTR(dispid, *(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_BORDER:
                    hr = WriteBorderToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_MARGIN:
                case DISPID_A_PADDING:
                case DISPID_A_BORDERCOLOR:
                case DISPID_A_BORDERWIDTH:
                case DISPID_A_BORDERSTYLE:
                    hr = WriteExpandedPropertyToBSTR(dispid, *(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_LISTSTYLE:
                    hr = WriteListStyleToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_BACKGROUNDPOSITION:
                    hr = WriteBackgroundPositionToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(hr == S_OK)
                    {
                        hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        FormsFreeString(bstrTemp);
                    }
                    break;

                case DISPID_A_FONT:
                    if((*(CAttrArray**)pSubObject)->Find(DISPID_A_FONT, CAttrValue::AA_Attribute))
                    {
                        hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
                    }
                    else
                    {
                        // We need to cook up a "font" property.
                        hr = WriteFontToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                        if(!hr)
                        {
                            if(*bstrTemp)
                            {
                                hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                            }
                            FormsFreeString(bstrTemp);
                        }
                    }
                    break;
                case DISPID_A_LAYOUTGRID:
                    // We need to cook up a "layout grid" property.
                    hr = WriteLayoutGridToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(!hr)
                    {
                        if(*bstrTemp)
                        {
                            hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        }
                        FormsFreeString(bstrTemp);
                    }
                    break;
                case DISPID_A_CLIP:
                    // We need to cook up a "clip" property with the "rect" shape.
                    hr = WriteClipToBSTR(*(CAttrArray**)pSubObject, &bstrTemp);
                    if(!hr)
                    {
                        if(*bstrTemp)
                        {
                            hr = pis->Write(bstrTemp, FormsStringLen(bstrTemp)*sizeof(TCHAR), NULL);
                        }
                        FormsFreeString(bstrTemp);
                    }
                    break;
                case DISPID_A_STYLETEXT:
                    hr = WriteStyleToBSTR(pObject, *(CAttrArray**)pSubObject, &bstrTemp, FALSE);
                    if(!hr)
                    {
                        if(*bstrTemp)
                        {
                            hr = pis->Write(bstrTemp, _tcslen(bstrTemp)*sizeof(TCHAR), NULL);
                        }
                        FormsFreeString(bstrTemp);
                    }
                    break;
                }
            }
            break;
        default:
            {
                BSTR* pbstr;
                switch(PROPTYPE(dwOpCode))
                {
                case PROPTYPE_VARIANT:
                    V_VT((VARIANT*)pv) = VT_BSTR;
                    pbstr = &(((VARIANT*)pv)->bstrVal);
                    break;
                case PROPTYPE_BSTR:
                    pbstr = (BSTR*)pv;
                    break;
                default:
                    Assert("Can't get anything but a VARIANT or BSTR for style component properties!");
                    hr = S_FALSE;
                    goto Cleanup;
                }
                switch(dispid)
                {
                case DISPID_A_BACKGROUND:
                    hr = WriteBackgroundStyleToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;
                case DISPID_A_LISTSTYLEIMAGE:
                case DISPID_A_BACKGROUNDIMAGE:
                case DISPID_A_FONTFACESRC:
                case DISPID_A_BEHAVIOR:
                    {
                        CString cstr;
                        if((*(CAttrArray**)pSubObject)->Find(dispid, CAttrValue::AA_Attribute))
                        {
                            BSTR bstrSub;
                            hr = HandleStringProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_BSTR<<16), 
                                &bstrSub, pObject, pSubObject);
                            if(hr == S_OK)
                            {
                                if(bstrSub && *bstrSub)
                                {
                                    // CONSIDER (alexz) using Format, to remove the memallocs here
                                    if(dispid != DISPID_A_BEHAVIOR)
                                    {
                                        cstr.Set(_T("url("));
                                    }

                                    cstr.Append(bstrSub);

                                    if(dispid != DISPID_A_BEHAVIOR)
                                    {
                                        cstr.Append(_T(")"));
                                    }
                                }
                                else
                                {
                                    // We only get here if a NULL string was stored in the array; i.e., the value is NONE.
                                    cstr.Set(_T("none"));
                                }
                                FormsFreeString(bstrSub);
                            }
                        }
                        hr = cstr.AllocBSTR(pbstr);
                    }
                    break;

                case DISPID_A_TEXTAUTOSPACE:
                    hr = WriteTextAutospaceToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;
                case DISPID_A_TEXTDECORATION:
                    hr = WriteTextDecorationToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_TEXTDECORATIONNONE:
                case DISPID_A_TEXTDECORATIONUNDERLINE:
                case DISPID_A_TEXTDECORATIONOVERLINE:
                case DISPID_A_TEXTDECORATIONLINETHROUGH:
                case DISPID_A_TEXTDECORATIONBLINK:
                    if(PROPTYPE(dwOpCode) != PROPTYPE_VARIANT)
                    {
                        Assert("Can't get/set text-decoration subproperties as anything but VARIANTs!");
                        hr = S_FALSE;
                        goto Cleanup;
                    }

                    V_VT((VARIANT*)pv) = VT_BOOL;
                    ((VARIANT*)pv)->boolVal = 0;

                    if(*((CAttrArray**)pSubObject))
                    {
                        // See if we already have a text-decoration value
                        CAttrValue* pAV = (*((CAttrArray**)pSubObject))->Find(DISPID_A_TEXTDECORATION, CAttrValue::AA_Attribute);
                        if(pAV)
                        {   // We do!  Copy its value into our working variant
                            long lVal = pAV->GetLong();

                            switch(dispid)
                            {
                            case DISPID_A_TEXTDECORATIONNONE:
                                lVal &= TD_NONE;
                                break;
                            case DISPID_A_TEXTDECORATIONUNDERLINE:
                                lVal &= TD_UNDERLINE;
                                break;
                            case DISPID_A_TEXTDECORATIONOVERLINE:
                                lVal &= TD_OVERLINE;
                                break;
                            case DISPID_A_TEXTDECORATIONLINETHROUGH:
                                lVal &= TD_LINETHROUGH;
                                break;
                            case DISPID_A_TEXTDECORATIONBLINK:
                                lVal &= TD_BLINK;
                                break;
                            }
                            if(lVal)
                            {
                                ((VARIANT*)pv)->boolVal = -1;
                            }
                        }
                    }
                    break;

                case DISPID_A_BORDERTOP:
                case DISPID_A_BORDERRIGHT:
                case DISPID_A_BORDERBOTTOM:
                case DISPID_A_BORDERLEFT:
                    hr = WriteBorderSidePropertyToBSTR(dispid, *(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_BORDER:
                    hr = WriteBorderToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_MARGIN:
                case DISPID_A_PADDING:
                case DISPID_A_BORDERCOLOR:
                case DISPID_A_BORDERWIDTH:
                case DISPID_A_BORDERSTYLE:
                    hr = WriteExpandedPropertyToBSTR(dispid, *(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_LISTSTYLE:
                    hr = WriteListStyleToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_BACKGROUNDPOSITION:
                    hr = WriteBackgroundPositionToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_FONT:
                    if((*(CAttrArray**)pSubObject)->Find(DISPID_A_FONT, CAttrValue::AA_Attribute))
                    {
                        hr = HandleStringProperty(dwOpCode, pv, pObject, pSubObject);
                    }
                    else
                    {
                        // We need to cook up a "font" property.
                        hr = WriteFontToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    }
                    break;

                case DISPID_A_LAYOUTGRID:
                    hr = WriteLayoutGridToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                case DISPID_A_STYLETEXT:
                    hr = WriteStyleToBSTR(pObject, *(CAttrArray**)pSubObject, pbstr, FALSE);
                    break;

                case DISPID_A_CLIP:
                    // We need to cook up a "clip" property with the "rect" shape.
                    hr = WriteClipToBSTR(*(CAttrArray**)pSubObject, pbstr);
                    break;

                default:
                    Assert("Unrecognized type being handled by CStyleUrl handler!" && FALSE);
                    break;
                }
                if(hr == S_FALSE)
                {
                    hr = FormsAllocString(_T(""), pbstr);
                }
            }
            break;
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Function:   ::ParseBackgroundPositionProperty
//
//  Synopsis:   Parses a background-position string in CSS format and sets
//              the appropriate sub-properties.
//
//-------------------------------------------------------------------------
HRESULT ParseBackgroundPositionProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBackgroundPosition)
{
    TCHAR* pszString;
    TCHAR* pszCopy;
    TCHAR* pszNextToken;
    BOOL fInsideParens;
    BOOL fXIsSet = FALSE;
    BOOL fYIsSet = FALSE;
    PROPERTYDESC* pPropertyDesc;
    TCHAR* pszLastXToken = NULL;
    HRESULT hr = S_OK;

    if(!pcszBackgroundPosition || !(*pcszBackgroundPosition))
    {
        // Empty value must set the properties to 0%
        s_propdescCStylebackgroundPositionX.a.TryLoadFromString(dwOpCode, _T("0%"), pObject, ppAA);
        s_propdescCStylebackgroundPositionY.a.TryLoadFromString(dwOpCode, _T("0%"), pObject, ppAA);
        return S_OK;
    }


    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszBackgroundPosition)+1];
    if(!pszCopy)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    _tcscpy(pszCopy, pcszBackgroundPosition);

    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        fInsideParens = FALSE;
        while(_istspace(*pszString))
        {
            pszString++;
        }

        while(*pszNextToken && (fInsideParens || !_istspace(*pszNextToken)))
        {
            if(*pszNextToken == _T('('))
            {
                fInsideParens = TRUE;
            }
            if(*pszNextToken == _T(')'))
            {
                fInsideParens = FALSE;
            }
            pszNextToken++;
        }

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        if(fXIsSet && !fYIsSet)
        {
            pPropertyDesc = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionY;
        }
        else // If X and Y have both either been set or both not set or just Y set, then try X first.
        {
            pPropertyDesc = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionX;
        }

        if(pPropertyDesc->TryLoadFromString(dwOpCode,pszString, pObject, ppAA))
        {
            // Failed: try the other propdesc, it might be a enum in that direction
            if(fXIsSet && !fYIsSet)
            {
                pPropertyDesc = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionX;
            }
            else
            {
                pPropertyDesc = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionY;
            }

            if(S_OK == pPropertyDesc->TryLoadFromString(dwOpCode, pszString, pObject, ppAA))
            {
                if(fXIsSet && !fYIsSet)
                {
                    fXIsSet = TRUE;
                    if(S_OK == s_propdescCStylebackgroundPositionY.a.TryLoadFromString(
                        dwOpCode, pszLastXToken, pObject, ppAA))
                    {
                        fYIsSet = TRUE;
                    }
                }
                else
                {
                    fYIsSet = TRUE;
                }
            }
            else
            {
                hr = E_INVALIDARG;
                goto Cleanup;
            }
        }
        else
        {
            if(fXIsSet && !fYIsSet)
            {
                fYIsSet = TRUE;
            }
            else
            {
                fXIsSet = TRUE;
                pszLastXToken = pszString;
            }
        }
        if(fXIsSet && fYIsSet)
        {
            goto Cleanup; // We're done - we've set both values.
        }
    }

    if(!fXIsSet)
    {
        s_propdescCStylebackgroundPositionX.a.TryLoadFromString(dwOpCode, _T("50%"), pObject, ppAA);
    }
    if(!fYIsSet)
    {
        s_propdescCStylebackgroundPositionY.a.TryLoadFromString(dwOpCode, _T("50%"), pObject, ppAA);
    }

Cleanup:
    delete[] pszCopy;
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::ParseClipProperty
//
//  Synopsis:   Parses a "clip" string in CSS format (from the positioning
//              specification) and sets the appropriate sub-properties.
//
//-------------------------------------------------------------------------
HRESULT ParseClipProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszClip)
{
    HRESULT hr = E_INVALIDARG;
    size_t nLen = pcszClip ? _tcslen(pcszClip) : 0;

    if(!pcszClip)
    {
        return S_OK;
    }

    if((nLen>6) && !_tcsnicmp(pcszClip, 4, _T("rect"), 4) && (pcszClip[nLen-1]==_T(')')))
    {
        // skip the "rect"
        LPCTSTR pcszProps = pcszClip + 4;

        // Terminate the string
        ((TCHAR*)pcszClip)[nLen-1] = 0;

        // skip any whitespace
        while(*pcszProps && *pcszProps==_T(' '))
        {
            pcszProps++;
        }

        // there better be a string, a '('
        if(*pcszProps != _T('('))
        {
            goto Cleanup;
        }

        // skip the '('
        pcszProps++;

        hr = ParseExpandProperty(ppAA, pObject, dwOpCode, pcszProps, DISPID_A_CLIP, TRUE);

        // Restore the string
        ((TCHAR*)pcszClip)[nLen-1] = _T(')');
    }

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}