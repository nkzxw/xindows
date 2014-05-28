
#include "stdafx.h"
#include "UnitValue.h"

#define CVIX(off)   (off-UNIT_POINT)

CUnitValue::TypeDesc CUnitValue::TypeNames[] =
{
    { _T("null"),   UNIT_NULLVALUE,     0,  0     }, // Never used
    { _T("pt"),     UNIT_POINT,         +3, 1000  },
    { _T("pc"),     UNIT_PICA,          +2, 100   },
    { _T("in"),     UNIT_INCH,          +3, 1000  },
    { _T("cm"),     UNIT_CM,            +3, 1000  },
    { _T("mm"),     UNIT_MM,            +2, 100   }, // Chosen to match HIMETRIC
    { _T("em"),     UNIT_EM,            +2, 100   },
    { _T("en"),     UNIT_EN,            +2, 100   },
    { _T("ex"),     UNIT_EX,            +2, 100   },
    { _T("px"),     UNIT_PIXELS,        0,  1     },
    { _T("%"),      UNIT_PERCENT,       +2, 100   },
    { _T("*"),      UNIT_TIMESRELATIVE, +2, 100   },
    { _T("float"),  UNIT_FLOAT,         +4, 10000 },
};

#define LOCAL_BUF_COUNT     (pdlLength+1)
HRESULT CUnitValue::NumberFromString(LPCTSTR pStr, const PROPERTYDESC* pPropDesc)
{
    BOOL fIsSigned = FALSE;
    long lNewValue = 0;
    LPCTSTR pStartPoint = NULL;
    UNITVALUETYPE uvt = UNIT_INTEGER;
    WORD i,j,k;
    HRESULT hr = S_OK;
    TCHAR tcValue[LOCAL_BUF_COUNT];  
    WORD wShift;
    NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(pPropDesc + 1);
    long lExponent = 0;
    BOOL fNegativeExponent = FALSE;

    enum ParseState
    {
        Starting,
        AfterPoint,
        AfterE,
        InExponent,
        InTypeSpecifier,
        ParseState_Last_enum
    } State = Starting;

    UINT uNumChars =0;
    UINT uPrefixDigits = 0;
    BOOL fSeenDigit = FALSE;

    // Find the unit specifier
    for(i=0; uNumChars<LOCAL_BUF_COUNT; i++)
    {
        switch(State)
        {
        case Starting:
            if(i==0 && (pStr[i]==_T('-') || pStr[i]==_T('+')))
            {
                tcValue[uNumChars++] = pStr[i];
                uPrefixDigits++;
                fIsSigned = TRUE;
            }
            else if(!_istdigit(pStr[i]))
            {
                if(pStr[i] == _T('.'))
                {
                    State = AfterPoint;
                }
                else if(pStr[i]==_T('\0') || pStr[i]==_T('\"')  ||pStr[i]==_T('\''))
                {
                    goto Convert;
                }
                else
                {
                    // such up white space, and treat whatevers left as the type
                    while(_istspace(pStr[i]))
                    {
                        i++;
                    }

                    pStartPoint = pStr + i;
                    State = (pStr[i]==_T('e') || pStr[i]==_T('E')) ? AfterE : InTypeSpecifier;
                }
            }
            else
            {
                fSeenDigit = TRUE;
                tcValue[uNumChars++] = pStr[i]; uPrefixDigits++;
            }
            break;

        case AfterPoint:
            if(!_istdigit(pStr[i]))
            {
                if(pStr[i] == _T('\0'))
                {
                    goto Convert;
                }
                else if(!_istspace(pStr[i]))
                {
                    pStartPoint = pStr + i;
                    State = (pStr[i]==_T('e') || pStr[i]==_T('E')) ? AfterE : InTypeSpecifier;
                }
            }
            else
            {
                fSeenDigit = TRUE;
                tcValue[uNumChars++] = pStr [i];
            }
            break;

        case AfterE:
            if(pStr[i]==_T('-') || pStr[i]==_T('+') || _istdigit(pStr[i]))
            {   // Looks like scientific notation to me.
                if(_istdigit(pStr[i]))
                {
                    lExponent = pStr [i] - _T('0');
                }
                else if(pStr[i] == _T('-'))
                {
                    fNegativeExponent = TRUE;
                }
                State = InExponent;
                break;
            }
            // Otherwise, this must just be a regular old "em" or "en" type specifier-
            // Set the state and let's drop back into this switch with the same char...
            State = InTypeSpecifier;
            i--;
            break;

        case InExponent:
            if(_istdigit(pStr[i]))
            {
                lExponent *= 10;
                lExponent += pStr[i] - _T('0');
                break;
            }
            while(_istspace(pStr[i]))
            {
                i++;
            }
            if(pStr[i] == _T('\0'))
            {
                goto Convert;
            }
            State = InTypeSpecifier; // otherwise, fall through to type handler
            pStartPoint = pStr + i;

        case InTypeSpecifier:
            if(_istspace(pStr[i]))
            {
                if(pPropDesc->IsStyleSheetProperty())
                {
                    while(_istspace(pStr[i]))
                    {
                        i++;
                    }
                    if(pStr[i] != _T('\0'))
                    {
                        hr = E_INVALIDARG;
                        goto Cleanup;
                    }
                }
                goto CompareTypeNames;
            }
            if(pStr[i] == _T('\0'))
            {
CompareTypeNames:
                for(j=0; j<ARRAYSIZE(TypeNames); j++)
                {
                    if(TypeNames[j].pszName)
                    {
                        int iLen = _tcslen(TypeNames[j].pszName);

                        if(_tcsnipre(TypeNames[j].pszName, iLen, pStartPoint, -1))
                        {
                            if(pPropDesc->IsStyleSheetProperty())
                            {
                                if(pStartPoint[iLen] && !_istspace(pStartPoint[ iLen ]))
                                {
                                    continue;
                                }
                            }
                            break;
                        }
                    }
                }
                if(j == ARRAYSIZE(TypeNames))
                {
                    if((ppp->bpp.dwPPFlags&PP_UV_UNITFONT) == PP_UV_UNITFONT)
                    {
                        hr = E_INVALIDARG;
                        goto Cleanup;
                    }
                    // Ignore invalid unit specifier
                    goto Convert;
                }
                uvt = TypeNames[j].uvt;
                goto Convert;
            }
            break;
        }
    }

Convert:
    if(!fSeenDigit && uvt!=UNIT_TIMESRELATIVE)
    {
        if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
        {
            long lEnum;
            hr = LookupEnumString(ppp, (LPTSTR)pStr, &lEnum);
            if(hr == S_OK)
            {
                SetValue(lEnum, UNIT_ENUM);
                goto Cleanup;
            }
        }

        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(fIsSigned && uvt==UNIT_INTEGER && ((ppp->bpp.dwPPFlags&PP_UV_UNITFONT)!=PP_UV_UNITFONT))
    {
        uvt = UNIT_RELATIVE;
    }

    // If the validation for this attribute does not require that we distinguish
    // between integer and relative, just store integer

    // Currently the only legal relative type is a Font size, if this changes either
    // add a separte PROPPARAM_ or do something else
    if(uvt==UNIT_RELATIVE && !(ppp->bpp.dwPPFlags&PROPPARAM_FONTSIZE))
    {
        uvt = UNIT_INTEGER;
    }

    // If a unit was supplied that is not valid for this propdesc,
    // drop back to the default (pixels if no notpresent and notassigned defaults)
    if((IsScalerUnit(uvt) && !(ppp->bpp.dwPPFlags&PROPPARAM_LENGTH)) ||
        (uvt==UNIT_PERCENT && !(ppp->bpp.dwPPFlags&PROPPARAM_PERCENTAGE)) ||
        (uvt==UNIT_TIMESRELATIVE && !(ppp->bpp.dwPPFlags&PROPPARAM_TIMESRELATIVE)) ||
        ((uvt==UNIT_INTEGER) && !(ppp->bpp.dwPPFlags&PROPPARAM_FONTSIZE)))
    {
        // If no units where specified and a default unit is specified, use it
        if((uvt!=UNIT_INTEGER) && pPropDesc->IsStyleSheetProperty()) // Stylesheet unit values only:
        {   // We'll default unadorned numbers to "px", but we won't change "100%" to "100px" if percent is not allowed.
            hr = E_INVALIDARG;
            goto Cleanup;
        }

        CUnitValue uv;
        uv.SetRawValue((long)pPropDesc->ulTagNotPresentDefault);
        uvt = uv.GetUnitType();
        if(uvt == UNIT_NULLVALUE)
        {
            uv.SetRawValue((long)pPropDesc->ulTagNotAssignedDefault);
            uvt = uv.GetUnitType();
            if(uvt == UNIT_NULLVALUE)
            {
                uvt = UNIT_PIXELS;
            }
        }
    }

    // Check the various types, don't do the shift for the various types
    switch(uvt)
    {
    case UNIT_INTEGER:
    case UNIT_RELATIVE:
    case UNIT_ENUM:
        wShift = 0;
        break;

    default:
        wShift = TypeNames[uvt].wShiftOffset;
        break;
    }

    if(lExponent && !fNegativeExponent && ((uPrefixDigits+wShift)<uNumChars))
    {
        long lAdditionalShift = uNumChars - (uPrefixDigits + wShift);
        if(lAdditionalShift > lExponent)
        {
            lAdditionalShift = lExponent;
        }
        wShift += (WORD)lAdditionalShift;
        lExponent -= lAdditionalShift;
    }
    // uPrefixDigits tells us how may characters there are before the point;
    // Assume we're always shifting to the right of the point
    k = (uPrefixDigits+wShift<LOCAL_BUF_COUNT) ? uPrefixDigits+wShift : LOCAL_BUF_COUNT-1;
    tcValue[k] = _T('\0');
    for(j=(WORD)uNumChars; j<k; j++)
    {
        tcValue[j] = _T('0');
    }

    // Skip leading zeros, they confuse StringToLong
    for(pStartPoint=tcValue; *pStartPoint==_T('0')&&*(pStartPoint+1)!=_T('\0'); pStartPoint++);

    if(*pStartPoint)
    {
        hr = ttol_with_error(pStartPoint, &lNewValue);
        // returns LONG_MIN or LONG_MIN on overflow
        if(hr)
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
    }
    else if(uvt != UNIT_TIMESRELATIVE)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if((uvt==UNIT_TIMESRELATIVE) && (lNewValue==0))
    {
        lNewValue = 1 * TypeNames[UNIT_TIMESRELATIVE].wScaleMult;
    }

    if(fNegativeExponent)
    {
        while(lNewValue && lExponent-->0)
        {
            lNewValue /= 10;
        }
    }
    else
    {
        while(lExponent-- > 0)
        {
            if(lNewValue > LONG_MAX/10)
            {
                lNewValue = LONG_MAX;
                break;
            }
            lNewValue *= 10;
        }
    }

    SetValue(lNewValue, uvt);

Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

CUnitValue::ConvertTable CUnitValue::BasicConversions[6][6] =
{
    {  // From UNIT_POINT
        { 1,1 },                // To UNIT_POINT
        { 1,12 },               // To UNIT_PICA
        { 1, 72 },              // To UNIT_INCH
        { 1*254, 72*100 },      // To UNIT_CM
        { 1*2540, 72*100 },     // To UNIT_MM
        { 20,1 },               // To UNIT_EM 
    },
    { // From UNIT_PICA
        { 12,1 },               // To UNIT_POINT
        { 1,1 },                // To UNIT_PICS
        { 12, 72 },             // To UNIT_INCH
        { 12*254, 72*100 },     // To UNIT_CM
        { 12*2540, 72*100 },    // To UNIT_MM
        { 20*12, 1 },           // To UNIT_EM 
    },
    { // From UNIT_INCH
        { 72,1 },               // To UNIT_POINT
        { 72,12 },              // To UNIT_PICA
        { 1, 1 },               // To UNIT_INCH
        { 254, 100 },           // To UNIT_CM
        { 2540, 100 },          // To UNIT_MM
        { 1440,1 },             // To UNIT_EM 
    },
    { // From UNIT_CM
        { 72*100,1*254 },        // To UNIT_POINT
        { 72*100,12*254 },       // To UNIT_PICA
        { 100, 254 },            // To UNIT_INCH
        { 1, 1 },                // To UNIT_CM
        { 10, 1 },               // To UNIT_MM
        { 20*72*100, 254 },      // To UNIT_EM 
        },
    { // From UNIT_MM
        { 72*100,1*2540 },       // To UNIT_POINT
        { 72*100,12*2540 },      // To UNIT_PICA
        { 100, 2540 },           // To UNIT_INCH
        { 10, 1 },               // To UNIT_CM
        { 1,1 },                 // To UNIT_MM
        { 20*72*100, 2540 },     // To UNIT_EM 
    },
    { // From UNIT_EM
        { 1,20 },                // To UNIT_POINT
        { 1, 20*12 },            // To UNIT_PICA
        { 1, 20*72 },            // To UNIT_INCH
        { 254, 20*72*100 },      // To UNIT_CM
        { 2540, 20*72*100 },     // To UNIT_MM
        { 1,1 },                 // To UNIT_EM 
    }
};

long CUnitValue::ConvertTo(long lValue, UNITVALUETYPE uvtFromUnits, UNITVALUETYPE uvtTo, 
                           DIRECTION direction, long lFontHeight/*=1*/)
{
    UNITVALUETYPE uvtToUnits;
    long lFontMul = 1, lFontDiv = 1;

    if(uvtFromUnits == uvtTo)
    {
        return lValue;
    }

    if(uvtFromUnits == UNIT_PIXELS)
    {
        int logPixels = (direction==DIRECTION_CX) ? _afxGlobalData._sizePixelsPerInch.cx :
            _afxGlobalData._sizePixelsPerInch.cy;

        // Convert to inches ( stored precision )
        lValue = MulDivQuick(lValue, TypeNames[UNIT_INCH].wScaleMult, logPixels);

        uvtFromUnits = UNIT_INCH;
        uvtToUnits = uvtTo;
    }
    else if(uvtTo == UNIT_PIXELS)
    {
        uvtToUnits = UNIT_INCH;
    }
    else
    {
        uvtToUnits = uvtTo;
    }

    if(uvtToUnits == UNIT_EM)
    {
        lFontDiv = lFontHeight;
    }
    if(uvtFromUnits == UNIT_EM)
    {
        lFontMul = lFontHeight;
    }

    if(uvtToUnits == UNIT_EX)
    {
        lFontDiv = lFontHeight * 2;
        uvtToUnits = UNIT_EM;
    }
    if(uvtFromUnits == UNIT_EX)
    {
        lFontMul = lFontHeight / 2;
        uvtFromUnits = UNIT_EM;
    }

    // Note that we perform two conversions in one here to avoid losing
    // intermediate accuracy for conversions to pixels, i.e. we're converting
    // units to inches and from inches to pixels.
    if(uvtTo == UNIT_PIXELS)
    {
        int logPixels = (direction==DIRECTION_CX) ? _afxGlobalData._sizePixelsPerInch.cx :
            _afxGlobalData._sizePixelsPerInch.cy;

        lValue = MulDivQuick(
            lValue,
            lFontMul*BasicConversions[CVIX(uvtFromUnits)][CVIX(uvtToUnits)].lMul*
            TypeNames[uvtToUnits].wScaleMult*logPixels,
            lFontDiv*BasicConversions[CVIX(uvtFromUnits)][CVIX(uvtToUnits)].lDiv*
            TypeNames[uvtFromUnits].wScaleMult*TypeNames[UNIT_INCH].wScaleMult);
    }
    else
    {
        lValue = MulDivQuick(
            lValue,
            lFontMul*BasicConversions[CVIX(uvtFromUnits)][CVIX(uvtToUnits)].lMul*
            TypeNames[uvtToUnits].wScaleMult,
            lFontDiv*BasicConversions[CVIX(uvtFromUnits)][CVIX(uvtToUnits)].lDiv*
            TypeNames[uvtFromUnits].wScaleMult);
    }        

    return lValue;
}

// Set the Value from the current system-wide measurement units - currently HiMetric for
// measurement type, but retains the previous storage unit type. If you pass in
// a transform, I'll convert to Window units ( assume stored value in Document units )
BOOL CUnitValue::SetHimetricMeasureValue(long lNewValue, DIRECTION direction, long lFontHeight)
{
    UNITVALUETYPE uvtToUnits = GetUnitType();
    long lToValue;

    // Check illegal conversions
    if(uvtToUnits==UNIT_TIMESRELATIVE || uvtToUnits==UNIT_PERCENT)
    {
        return FALSE;
    }

    if(uvtToUnits ==  UNIT_NULLVALUE)
    {
        // Not had a unit type yet
        uvtToUnits = UNIT_PIXELS;
    }

    lToValue = ConvertTo(lNewValue, UNIT_MM, uvtToUnits, direction, lFontHeight);

    SetValue(lToValue, uvtToUnits);

    return TRUE;
}

// Set the Percent Value - Don't need a transform because we assume lNewValue is in the
// same transform as lPercentOfValue
BOOL CUnitValue::SetPercentValue(long lNewValue, DIRECTION direction, long lPercentOfValue)
{
    UNITVALUETYPE uvtToUnits = GetUnitType();
    long lUnitValue = GetUnitValue();

    // Set the internal percentage to reflect what percentage of lMaxValue
    if(lPercentOfValue == 0)
    {
        lUnitValue = 0;
    }
    else if(uvtToUnits == UNIT_PERCENT)
    {
        lUnitValue = MulDivQuick(lNewValue, 100*TypeNames[UNIT_PERCENT].wScaleMult, lPercentOfValue);
    }
    else
    {
        lUnitValue = MulDivQuick(lNewValue, TypeNames[UNIT_TIMESRELATIVE].wScaleMult, lPercentOfValue);
    }
    SetValue(lUnitValue, uvtToUnits);
    return TRUE;
}

float CUnitValue::GetFloatValueInUnits(UNITVALUETYPE uvtTo, DIRECTION dir, long lFontHeight/*=1*/)
{
    // Convert the unitvalue into a floating point number in units uvt
    long lEncodedValue = ConvertTo(GetUnitValue(), GetUnitType(), uvtTo, dir, lFontHeight);

    return (float)lEncodedValue/(float)TypeNames[uvtTo].wScaleMult;
}

HRESULT CUnitValue::SetFloatValueKeepUnits(float fValue, UNITVALUETYPE uvt, long lCurrentPixelValue,
                                           DIRECTION dir, long lFontHeight)
{
    long lNewValue;

    // Set the new value to the equivalent of : fValue in uvt units
    // If the current value is a percent use the lCurrentPixelValue to
    // work out the new size

    // There are a number of restrictions on the uvt coming in. Since it's either
    // Document Units when called from CUnitMeasurement::SetValue or
    // PIXELS when called from CUnitMeasurement::SetPixelValue we restrict our
    // handling to the possible doc unit types - which are all the "metric" types
    lNewValue = (long)(fValue * (float)TypeNames[uvt].wScaleMult);

    if(uvt == GetUnitType())
    {
        SetValue(lNewValue, uvt);
    }
    else if(uvt==UNIT_NULLVALUE || !IsScalerUnit(uvt))
    {
        return E_INVALIDARG;
    }
    else if(GetUnitType() == UNIT_NULLVALUE)
    {
        // If we're current NUll, just set to the new value
        SetValue(lNewValue, uvt);
    }
    else if(IsScalerUnit(GetUnitType()))
    {
        // If the conversion is to/from metric units, just convert units
        SetValue(ConvertTo(lNewValue, uvt, GetUnitType(), dir, lFontHeight), GetUnitType());
    }
    else
    {
        // unit value holds a relative unit,
        // Convert the fValue,uvt to pixels
        lNewValue = ConvertTo(lNewValue, uvt, UNIT_PIXELS, dir, lFontHeight);

        if(GetUnitType() == UNIT_PERCENT)
        {
            SetValue(
                MulDivQuick(lNewValue, 
                100*TypeNames[UNIT_PERCENT].wScaleMult, 
                (!!lCurrentPixelValue)?lCurrentPixelValue:1), // don't pass in 0 for divisor
                UNIT_PERCENT);
        }
        else
        {
            // Cannot keep units without loss of precision, over-ride
            SetValue(lNewValue, UNIT_PIXELS);
        }
    }
    return S_OK;
}

BOOL CUnitValue::IsScalerUnit(UNITVALUETYPE uvt)
{
    switch(uvt)
    {
    case UNIT_POINT:
    case UNIT_PICA:
    case UNIT_INCH:
    case UNIT_CM:
    case UNIT_MM:
    case UNIT_PIXELS:
    case UNIT_EM:
    case UNIT_EN:
    case UNIT_EX:
    case UNIT_FLOAT:
        return TRUE;

    default:
        return FALSE;
    }
}

// We usually do not append default unit types to the string we return. Setting fAlwaysAppendUnit
//   to true forces the unit to be always appended to the number
HRESULT CUnitValue::FormatBuffer(LPTSTR szBuffer, UINT uMaxLen, const PROPERTYDESC* pPropDesc,
                                 BOOL fAlwaysAppendUnit/*=FALSE*/) const
{
    HRESULT         hr = S_OK;
    UNITVALUETYPE   uvt = GetUnitType();
    long            lUnitValue = GetUnitValue();
    TCHAR           ach[20];
    int             nLen, i, nOutLen=0;
    BOOL            fAppend = TRUE;

    switch(uvt)
    {
    case UNIT_ENUM:
        {
            ENUMDESC* pEnumDesc;
            LPCTSTR pcszEnum;

            Assert((((NUMPROPPARAMS*)(pPropDesc+1))->bpp.dwPPFlags&PROPPARAM_ENUM) && "Not an enum property!");

            hr = E_INVALIDARG;
            pEnumDesc = pPropDesc->GetEnumDescriptor();
            if(pEnumDesc)
            {
                pcszEnum = pEnumDesc->StringPtrFromEnum(lUnitValue);
                if(pcszEnum)
                {
                    _tcsncpy(szBuffer, pcszEnum, uMaxLen);
                    szBuffer[uMaxLen-1] = _T('\0');
                    hr = S_OK;
                }
            }
        }
        break;

    case UNIT_INTEGER:
        hr = Format(_afxGlobalData._hInstResource, 0, szBuffer, uMaxLen, _T("<0d>"), lUnitValue);
        break;

    case UNIT_RELATIVE:
        if(lUnitValue >= 0)
        {
            szBuffer[nOutLen++] = _T('+');
        }
        hr = Format(_afxGlobalData._hInstResource, 0, szBuffer+nOutLen, uMaxLen-nOutLen, _T("<0d>"), lUnitValue);
        break;

    case UNIT_NULLVALUE:
        // BUGBUG Need to chage save code to NOT comapre the lDefault directly
        // But to go through the handler
        szBuffer[0] = 0;
        hr = S_OK;
        break;

    case UNIT_TIMESRELATIVE:
        _tcscpy(szBuffer, TypeNames[uvt].pszName);
        hr = Format(_afxGlobalData._hInstResource, 0, szBuffer+_tcslen(szBuffer), uMaxLen-_tcslen(szBuffer),
            _T("<0d>"), lUnitValue);
        break;

    default:
        hr = Format(_afxGlobalData._hInstResource, 0, ach, ARRAYSIZE(ach), _T("<0d>"), lUnitValue);
        if(hr)
        {
            goto Cleanup;
        }

        nLen = _tcslen(ach);

        LPTSTR pStr = ach;

        if(ach[0] == _T('-'))
        {
            szBuffer[nOutLen++] = _T('-');
            pStr++;
            nLen--;
        }

        if(nLen > TypeNames[uvt].wShiftOffset)
        {
            for(i=0; i<(nLen-TypeNames[uvt].wShiftOffset); i++)
            {
                szBuffer[nOutLen++] = *pStr++;
            }
            szBuffer[nOutLen++] = _T('.');
            for(i=0; i<TypeNames[uvt].wShiftOffset; i++)
            {
                szBuffer[nOutLen++] = *pStr++;
            }
        }
        else
        {
            szBuffer[nOutLen++] = _T('0');
            szBuffer[nOutLen++] = _T('.');
            for(i=0; i<(TypeNames[uvt].wShiftOffset-nLen); i++)
            {
                szBuffer[nOutLen++] = _T('0');
            }
            for(i=0; i<nLen; i++)
            {
                szBuffer[nOutLen++] = *pStr++;
            }
        }

        // Strip trailing 0 digits. If there's at least one trailing digit put a point
        for(i=nOutLen-1; ; i--)
        {
            if(szBuffer[i] == _T('.'))
            {
                nOutLen--;
                break;
            }
            else if(szBuffer[i] == _T('0'))
            {
                nOutLen--;
            }
            else
            {
                break;
            }
        }

        // Append the type prefix, unless it's the default and not forced
        if(uvt == UNIT_FLOAT)
        {
            fAppend = FALSE;
        }
        else if(!fAlwaysAppendUnit && !( pPropDesc->GetPPFlags()&PROPPARAM_LENGTH))
        {
            CUnitValue uvDefault;
            uvDefault.SetRawValue(pPropDesc->ulTagNotPresentDefault);
            UNITVALUETYPE uvtDefault = uvDefault.GetUnitType();

            if(uvt==uvtDefault || (uvtDefault==UNIT_NULLVALUE && uvt==UNIT_PIXELS))
            {
                fAppend = FALSE;
            }
        }

        if(fAppend)
        {
            _tcscpy(szBuffer+nOutLen, TypeNames[uvt].pszName);
        }
        else
        {
            szBuffer[nOutLen] = _T('\0');
        }

        break;
    }
Cleanup:
    RRETURN(hr);
}

long CUnitValue::SetValue(long lVal, UNITVALUETYPE uvt) 
{
    if(lVal > (LONG_MAX>>NUM_TYPEBITS))
    {
        lVal = LONG_MAX>>NUM_TYPEBITS;
    }
    else if(lVal < (LONG_MIN>>NUM_TYPEBITS))
    {
        lVal = LONG_MIN>>NUM_TYPEBITS;
    }
    return _lValue = (lVal<<NUM_TYPEBITS)|uvt; 
}

// Get the Value in the cuurent system-wide measurement units - currently HiMetric for
// measurement type, just the value for UNIT_INTEGER or UNIT_RELATIVE. If you pass in
// a trannsform, I'll convert to Document units ( assume stored value in Window units )
long CUnitValue::GetMeasureValue(DIRECTION direction, UNITVALUETYPE uvtConvertToUnits, long lFontHeight) const
{
    UNITVALUETYPE uvtFromUnits = GetUnitType();
    long lValue = GetUnitValue();
    return ConvertTo(lValue, uvtFromUnits, uvtConvertToUnits, direction, lFontHeight);
}

long CUnitValue::GetPercentValue(DIRECTION direction, long lPercentOfValue) const
{
    //  lPercentOfValue is the 100%value, expressed in Window Himmetric coords
    UNITVALUETYPE uvtFromUnits = GetUnitType() ;
    long lUnitValue = GetUnitValue();

    if(uvtFromUnits == UNIT_PERCENT)
    {
        return  MulDivQuick(lPercentOfValue, lUnitValue, 100*TypeNames[UNIT_PERCENT].wScaleMult);
    }
    else
    {
        return MulDivQuick(lPercentOfValue, lUnitValue, TypeNames[UNIT_TIMESRELATIVE].wScaleMult);
    }
}

HRESULT CUnitValue::ConvertToUnitType(UNITVALUETYPE uvtConvertTo, long lCurrentPixelSize, 
                                      DIRECTION dir, long lFontHeight/*=1*/)
{
    AssertSz(GetUnitType()!=UNIT_PERCENT, "Cannot handle percent units! Contact RGardner!");

    // Convert the unit value to the new type but keep the
    // absolute value of the units the same
    // e.e. 20.34 pels -> 20.34 in
    if(GetUnitType() == uvtConvertTo)
    {
        return S_OK;
    }

    if(GetUnitType() == UNIT_NULLVALUE)
    {
        // The current value is Null, treat it as a pixel unit with lCurrentPixelSize value
        SetValue(lCurrentPixelSize, UNIT_PIXELS);
        // fall into one of the cases below
    }

    if(uvtConvertTo == UNIT_NULLVALUE)
    {
        SetValue(0, UNIT_NULLVALUE);
    }
    else if(IsScalerUnit(GetUnitType()) && IsScalerUnit(uvtConvertTo))
    {
        // Simply converting between scaler units e.g. 20mm => 2px
        SetValue(ConvertTo(GetUnitValue(), GetUnitType(), uvtConvertTo, dir, lFontHeight), uvtConvertTo);
    }
    else if(IsScalerUnit(GetUnitType()))
    {
        // Convert from a scaler type to a non-scaler type e.g. 20.3in => %
        // Have no reference for conversion, max it out
        switch(uvtConvertTo)
        {
        case UNIT_PERCENT:
            SetValue(100*TypeNames[UNIT_PERCENT].wScaleMult, UNIT_PERCENT);
            break;

        case UNIT_TIMESRELATIVE:
            SetValue(1*TypeNames[UNIT_PERCENT].wScaleMult, UNIT_TIMESRELATIVE);
            break;

        default:
            AssertSz(FALSE, "Invalid type passed to ConvertToUnitType()");
            return E_INVALIDARG;
        }
    }
    else if(IsScalerUnit(uvtConvertTo))
    {
        // Convert from a non-scaler to a scaler type use the current pixel size
        // e.g. We know that 20% is equivalent to 152 pixels. So we convert
        // the current pixel value to the new metric unit
        SetValue(ConvertTo(lCurrentPixelSize, UNIT_PIXELS, uvtConvertTo, dir, lFontHeight), uvtConvertTo);
    }
    else
    {
        // Convert between non-scaler types e,g, 20% => *
        // Since we have only two non-sclaer types right now:-
        switch(uvtConvertTo)
        {
        case UNIT_PERCENT:  // From UNIT_TIMESRELATIVE
            // 1* == 100%
            SetValue(
                MulDivQuick(GetUnitValue(),
                100*TypeNames[UNIT_PERCENT].wScaleMult,
                TypeNames[UNIT_TIMESRELATIVE].wScaleMult),
                uvtConvertTo);
            break;

        case UNIT_TIMESRELATIVE: // From UNIT_PERCENT
            // 100% == 1*
            SetValue(MulDivQuick(GetUnitValue(),
                TypeNames[UNIT_TIMESRELATIVE].wScaleMult,
                TypeNames[UNIT_PERCENT].wScaleMult*100),
                uvtConvertTo);
            break;

        default:
            AssertSz(FALSE, "Invalid type passed to ConvertToUnitType()");
            return E_INVALIDARG;
        }
    }
    return S_OK;
}

HRESULT CUnitValue::SetFloatUnitValue(float fValue)
{
    // Set the value from fValue
    // but keep the internaly stored units intact
    UNITVALUETYPE uvType;

    if((uvType=GetUnitType()) == UNIT_NULLVALUE)
    {
        // This case can only happen if SetUnitValue is called when the current value is NULL
        // so force a unit type
        uvType = UNIT_PIXELS;
    }

    // Convert the fValue into a CUnitValue with current unit
    SetValue((long)(fValue*(float)TypeNames[uvType].wScaleMult), uvType);
    return S_OK;
}

HRESULT CUnitValue::FromString(LPCTSTR pStr, const PROPERTYDESC* pPropDesc)
{
    NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(pPropDesc + 1);

    // if string is empty and boolean is an allowed type make it so...
    // Note the special handling of the table relative "*" here
    while(_istspace(*pStr))
    {
        pStr++;
    }

    if(_istdigit(pStr[0]) || pStr[0]==_T('+') || pStr[0]==_T('-') ||
        pStr[0]==_T('*') || pStr[0]==_T('.'))
    {
        // Assume some numerical value followed by an optional unit
        // specifier
        RRETURN1(NumberFromString(pStr, pPropDesc), E_INVALIDARG);
    }
    else
    {
        // Assume an enum
        if(ppp->bpp.dwPPFlags & PROPPARAM_ENUM)
        {
            long lNewValue;

            if(S_OK == LookupEnumString(ppp, (LPTSTR)pStr, &lNewValue))
            {
                SetValue(lNewValue, UNIT_ENUM);
                return S_OK;
            }
        }
        return E_INVALIDARG;
    }
}

extern HRESULT WriteText(IStream*, TCHAR*);
HRESULT CUnitValue::Persist(IStream* pStream, const PROPERTYDESC* pPropDesc) const
{
    TCHAR           cBuffer[30];
    int             iBufSize = ARRAYSIZE(cBuffer);
    LPTSTR          pBuffer = (LPTSTR)&cBuffer;
    HRESULT         hr;
    UNITVALUETYPE   uvt = GetUnitType();

    if(uvt==UNIT_PERCENT)
    {
        *pBuffer++ = _T('\"');
        iBufSize--;
    }

    hr = FormatBuffer(pBuffer, iBufSize, pPropDesc);
    if(hr)
    {
        goto Cleanup;
    }

    // and finally the close quotes
    if(uvt==UNIT_PERCENT)
    {
        _tcscat(cBuffer, _T("\""));
    }

    hr = WriteText(pStream, cBuffer);

Cleanup:
    RRETURN(hr);
}

HRESULT CUnitValue::IsValid(const PROPERTYDESC* pPropDesc) const
{
    HRESULT hr = S_OK;
    BOOL fTestAnyHow = FALSE;
    UNITVALUETYPE uvt = GetUnitType();
    long lUnitValue = GetUnitValue();
    NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(pPropDesc + 1);

    if((ppp->lMin==0) && (lUnitValue<0))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    switch(uvt)
    {
    case UNIT_TIMESRELATIVE:
        if(!(ppp->bpp.dwPPFlags&PROPPARAM_TIMESRELATIVE))
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        break;

    case UNIT_ENUM:
        if(!(ppp->bpp.dwPPFlags&PROPPARAM_ENUM) &&
            !(ppp->bpp.dwPPFlags&PROPPARAM_ANUMBER))
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        break;

    case UNIT_INTEGER:
        if(!(ppp->bpp.dwPPFlags&PROPPARAM_FONTSIZE) && 
            !(ppp->bpp.dwPPFlags&PROPPARAM_ANUMBER))
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        break;

    case UNIT_RELATIVE:
        if(ppp->bpp.dwPPFlags&PROPPARAM_FONTSIZE)
        {
            // Netscape treats any FONTSIZE as valid . A value greater than
            // 6 gets treated as seven, less than -6 gets treated as -6.
            // 8 and 9 get treated as "larger" and "smaller".
            goto Cleanup;
        }
        else if(!(ppp->bpp.dwPPFlags&PROPPARAM_ANUMBER))
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        break;

    case UNIT_PERCENT:
        if(!(ppp->bpp.dwPPFlags&PROPPARAM_PERCENTAGE))
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        else
        {
            long lMin = ((CUnitValue*)&(ppp->lMin))->GetUnitValue();
            long lMax = ((CUnitValue*)&(ppp->lMax))->GetUnitValue();

            lUnitValue = lUnitValue / TypeNames[UNIT_PERCENT].wScaleMult;

            if(lUnitValue<lMin || lUnitValue>lMax)
            {
                hr = E_INVALIDARG;
                goto Cleanup;
            }
        }
        break;

    case UNIT_NULLVALUE: // Always valid
        break;

    case UNIT_PIXELS: // pixels are HTML and CSS valid
        fTestAnyHow = TRUE;
        // fall through

    case UNIT_FLOAT:
    default: // other Units of measurement are only CSS valid
        if(!(ppp->bpp.dwPPFlags&PROPPARAM_LENGTH) && !fTestAnyHow)
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
        else
        {
            // it is a length but sometimes we do not want negatives
            long lMin = ((CUnitValue*)&(ppp->lMin))->GetUnitValue();
            long lMax = ((CUnitValue*)&(ppp->lMax))->GetUnitValue();
            if(lUnitValue<lMin || lUnitValue>lMax)
            {
                hr = E_INVALIDARG;
                goto Cleanup;
            }
        }

        break;
    }
Cleanup:
    RRETURN1(hr, E_INVALIDARG);
}

BOOL CUnitValue::SetFromHimetricValue(long lNewValue, DIRECTION direction, 
                                      long lPercentOfValue, long lFontHeight)
{
    UNITVALUETYPE uvtToUnits = GetUnitType();

    if(uvtToUnits==UNIT_PERCENT || uvtToUnits==UNIT_TIMESRELATIVE)
    {
        return SetPercentValue(lNewValue, direction, lPercentOfValue);
    }
    else
    {
        return SetHimetricMeasureValue(lNewValue, direction, lFontHeight);
    }
}

long CUnitValue::GetPixelValueCore(CTransform* pTransform, DIRECTION direction,
                                   long lPercentOfValue, long lFontHeight/*=1*/) const
{
    UNITVALUETYPE uvtFromUnits = GetUnitType();
    long lRetValue;

    switch(uvtFromUnits)
    {
    case UNIT_TIMESRELATIVE:
        lRetValue = MulDivQuick(GetUnitValue(), lPercentOfValue,
            TypeNames[UNIT_TIMESRELATIVE].wScaleMult);
        break;

    case UNIT_NULLVALUE:
        return 0L;

    case UNIT_PERCENT:
        lRetValue = GetPercentValue(direction, lPercentOfValue);
        break;

    case UNIT_EX:
    case UNIT_EM:
        // turn the twips into pixels
        lRetValue = ConvertTo(GetUnitValue(), uvtFromUnits, UNIT_PIXELS, direction, lFontHeight);

        // Still need to scale since zooming has been taken out and thus ConvertTo() no longer
        // scales (zooms) for printing (IE5 8543, IE4 11376).  Formerly:  pTransform = NULL;
        break;

    case UNIT_EN:
        lRetValue = MulDivQuick(GetUnitValue(), lFontHeight, TypeNames[UNIT_EN].wScaleMult);
        break;

    case UNIT_ENUM:
        // BUGBUG We need to put proper handling for enum in the Apply() so that
        // no one calls GetPixelValue on (for example) the margin CUV objects if
        // they're set to "auto".
        Assert("Shouldn't call GetPixelValue() on enumerated unit values!");
        return 0;

    case UNIT_INTEGER:
    case UNIT_RELATIVE:
        // BUGBUG this doesn't seem right; nothing's being done to convert the
        // retrieved value to anything resembling pixels.
        return GetUnitValue();

    case UNIT_FLOAT:
        lRetValue = MulDivQuick(GetUnitValue(), lFontHeight, TypeNames[UNIT_FLOAT].wScaleMult);
        break;

    default:
        lRetValue = GetMeasureValue(direction, UNIT_PIXELS);
        break;
    }
    // Finally convert from Window to Doc units, assuming return value is a length
    // For conversions to HIMETRIC, involve the transform
    if(pTransform)
    {
        BOOL fRelative = (UNIT_PERCENT==uvtFromUnits || UNIT_TIMESRELATIVE==uvtFromUnits);

        if(direction == DIRECTION_CX)
        {
            lRetValue = pTransform->DocPixelsFromWindowX(lRetValue, fRelative);
        }
        else
        {
            lRetValue = pTransform->DocPixelsFromWindowY(lRetValue, fRelative);
        }
    }
    return lRetValue;
}