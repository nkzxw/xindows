
#include "stdafx.h"
#include "StringUtil.h"

//+---------------------------------------------------------------
//
//  Function:   GetResource
//
//  Synopsis:   Loads any kind of resource.
//
//  Arguments:  [hinst] -- instance of the module with the resource
//              [lpstrId] -- the identifier of the resource
//              [lpstrType] -- the identifier of the resource type
//              [pcbSize] -- points to the out param: the number of bytes of resource data to load
//
//  Returns:    lpvBuf if the resource was successfully loaded, NULL otherwise
//
//  Notes:      This function combines Windows' FindResource, LoadResource,
//              LockResource.
//
//----------------------------------------------------------------
LPVOID GetResource(HINSTANCE hinst, LPCTSTR lpstrId, LPCTSTR lpstrType, ULONG* pcbSize)
{
    LPVOID  lpv;
    HGLOBAL hgbl;
    HRSRC   hrsrc;

    hrsrc = FindResource(hinst, lpstrId, lpstrType);
    if(!hrsrc)
    {
        return NULL;
    }

    hgbl = LoadResource(hinst, hrsrc);
    if(!hgbl)
    {
        return NULL;
    }

    lpv = LockResource(hgbl);
    if(pcbSize)
    {
        *pcbSize = lpv ? ::SizeofResource(hinst, hrsrc) : 0;
    }

    FreeResource(hgbl);

    return lpv;
}

//+------------------------------------------------------------------------
//  Format and VFormat
//
//  Format and VFormat are similiar to the sprintf family.
//  Format fixes two problems exhibited by sprintf.  First, the
//  order of argument substitution in sprintf is specified by
//  function caller. To support localization, the order of argument
//  substitution should be specified in the format string.  Second,
//  the range of formatting options are limited and cannot be
//  extended.
//
//  A format string consists of characters to be copied to the
//  output string and argument substitutions.  Arguments can be
//  referenced more than once.  This allows for very nice handling
//  of plurals.  For example: "There <0p/are/is/> <0d> toaster<0p//s/>
//  flying across my screen."
//
//  BiDi consideration: Arabic has a "dual" case between singular
//                      and plural (separate conjugation for a pair
//                      of things). We should find out if that poses
//                      a problem for Format.
//
//  FORMAT SPECIFIERS
//
//   "<" arg format options ">"
//
//      arg     - A single digit specifying the argument number.
//      handler - A single character specifying argument
//                type and display format.
//      options - Formatting options specific to format type.
//
//  NOTES
//
//      Two '<' characters in a row will print out as one '<'
//      character and does not mark the beginning of a format
//      specification.  Ex.: "The symbol << means less than."
//
//
//  FORMATS
//
//  d   Format long as decimal.
//      Options:
//          u       Argument is unsigned.
//          <digit> Field width.
//
//  s   Format zero terminated TCHAR *.
//      Options:
//          <none>
//
//  p   Format long as plural.
//      Options:
//          delimiter singular delimiter plural delimiter.
//      Example:
//          "There <0p/is/are/> <0d> toaster<0p//s/> flying across my screen"
//
//  i   Format resource string. First argugment is HINSTANCE, second
//      argument is string id.
//      Options:
//          <none>
//
//  g   Format GUID.
//      Options:
//          <none>
//
//  x   Format long as 8 character hex.
//      Options:
//          <none>
//
//  c   Format long as 6 character lowercase hex.  Option is for color.
//      Options:
//          <none>
//
//  C   Format long as 3 character lowercase hex.  Option is for color.
//      Options:
//          <none>
//
//  CALLING SEQUENCE
//
//      HRESULT VFormat(DWORD dwOptions,
//          void *pvOutput, int cchOutput,
//          TCHAR *pchFmt,
//          void *pvArgs);
//
//      HRESULT Format(DWORD dwOptions,
//          void *pvOutput, int cchOutput,
//          TCHAR *pchFmt,
//          ...);
//
//          dwOptions
//              Flags taken from the FMT_OPTIONS enumeration.
//
//              FMT_OUT_ALLOC
//                  Specifies that the pvOutput parameter is a pointer
//                  to a TCHAR*, and that the cchOutput parameter specifies
//                  the minimum number of characters to allocate for an
//                  output message buffer. The function allocates a buffer
//                  large enough to hold the formatted message, and
//                  places a pointer to the allocated buffer at the address
//                  specified by pvOutput. The caller should use the
//                  delete [] operator to free the buffer when it is no
//                  longer needed.
//
//              FMT_ARG_ARRAY
//                  Specifies that the pvArg parameter is NOT a va_list
//                  structure, but instead is just a pointer to an array
//                  of 32-bit values that represent the arguments.
//
//              FMT_EXTRA_NULL_MASK
//                  Add (dwOptions & FMT_EXTRA_NULL_MASK) null terminators to
//                  the end of the string.  This is useful for parsing
//                  multi-part strings where each part is separated
//                  by a null character.
//
//          pvOutput
//              Points to a buffer for the formatted (and null-terminated)
//              string. If dwOptions includes FMT_OUT_ALLOC, the function
//              allocates a buffer via operator new, and places the
//              address of the buffer at the address specified in pvOutput.
//
//          cchOutput
//              If the FMT_OUT_ALLOC flag is not set, this parameter
//              specifies the maximum number of characters that can be
//              stored in the output buffer. If the FMT_OUT_ALLOC flag is
//              set, this parameter specifies the minimum number of
//              characters to allocate for an output buffer.
//
//          pchFmt
//              Specifies the format string. Use MAKEINTRESOURCE(ids) to
//              specify the id of a string resource.
//
//          pvArgs
//              Pointe to 32-bit values thare are used as insert values
//              in the formatted string. By default this parameter is of
//              type va_list *.  If FMT_ARG_ARRAY is set, then pvArgs
//              is a pointer to an array of 32-bit values.
//
//-------------------------------------------------------------------------

//+------------------------------------------------------------------------
//
//  Function:   LoadString
//
//  Synopsis:   Get pointer to string resource. The returned
//              pointer does not need to be freed and is not
//              null terminated.
//
//  Arguments:  [hinst] - load from this module.
//              [ids] - string id.
//              [pcch] - number of characters in string.
//              [ppsz] - the string
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT LoadString(HINSTANCE hinst, UINT ids, int* pcch, TCHAR** ppsz)
{
    BYTE*   pb;
    int     i;

    pb = (BYTE*)GetResource(hinst, MAKEINTRESOURCE(((ids/16)+1)), RT_STRING, NULL);

    if(pb)
    {
        for(i=ids&0xF; --i>=0; )
        {
            pb += *(WORD*)pb*sizeof(TCHAR) + sizeof(TCHAR);
        }

        *pcch = *(WORD*)pb;
        *ppsz = (TCHAR*)(pb + sizeof(TCHAR));
    }
    else
    {
        *pcch = 0;
        *ppsz = NULL;
    }
    return *pcch!=0?S_OK:E_FAIL;
}

//+------------------------------------------------------------------------
//
//  Class:      OutputStream
//
//  Synopsis:   Abstracts writing to allocated buffer or fixed buffer.
//
//-------------------------------------------------------------------------
#define CCH_OUT_GROW    128
class OutputStream
{
private:
    DECLARE_MEMALLOC_NEW_DELETE();

public:
    HRESULT Init(BOOL fFixed, void* pvOutput, int cchOutput);
    HRESULT Put(TCHAR ch);
    HRESULT Put(TCHAR* sz);
    void    Nuke();

private:
    int     _cch;
    int     _cchAlloc;
    BOOL    _fAlloc;
    TCHAR*  _pch;
    union
    {
        TCHAR** _ppchAlloc;
        TCHAR*  _pchBase;
    };
};

//+------------------------------------------------------------------------
//
//  Member:     OutputStream::Init
//
//  Synopsis:   Initialize the output stream.
//
//  Arguments:  fAlloc - true if should allocate output buffer.
//              pvOutput - if fAlloc then place to store output buffer pointer,
//                  else pointer to acutal buffer.
//              cchOutput - if fAlloc then minimum allocation size
//                  else size of outpu buffer.
//
//-------------------------------------------------------------------------
HRESULT OutputStream::Init(BOOL fAlloc, void* pvOutput, int cchOutput)
{
    HRESULT hr = NOERROR;

    _fAlloc = fAlloc;
    if(fAlloc)
    {
        if(cchOutput <= 0)
        {
            cchOutput = CCH_OUT_GROW;
        }
        _ppchAlloc = (TCHAR**)pvOutput;
        _pch = *_ppchAlloc = (TCHAR*)MemAlloc(sizeof(TCHAR)*cchOutput);
        if(_pch)
        {
            _cch = _cchAlloc = cchOutput;
        }
        else
        {
            _cch = 0;
            hr = E_OUTOFMEMORY;
        }
    }
    else
    {
        _cch = cchOutput;
        _pch = _pchBase = (TCHAR*)pvOutput;
        _pch[_cch-1] = 0;
    }
    return hr;
}

//+------------------------------------------------------------------------
//
//  Member:     OutputStream::Put
//
//  Synopsis:   Put a character on the output stream.
//
//-------------------------------------------------------------------------
HRESULT OutputStream::Put(TCHAR ch)
{
    if(--_cch >= 0)
    {
        *_pch++ = ch;
    }
    else if(!_fAlloc)
    {
        Assert(0 && "Format: pvOutput too small.");
        return E_FAIL;
    }
    else
    {
        TCHAR* pch = *_ppchAlloc;
        *_ppchAlloc = (TCHAR*)MemAlloc(sizeof(TCHAR)*(_cchAlloc+CCH_OUT_GROW));
        if(!*_ppchAlloc)
        {
            MemFree(pch);
            _cch = 0;
            return E_OUTOFMEMORY;
        }
        else
        {
            memcpy(*_ppchAlloc, pch, _cchAlloc*sizeof(TCHAR));
            MemFree(pch);
            _pch = *_ppchAlloc + _cchAlloc;
            _cchAlloc += CCH_OUT_GROW;
            _cch = CCH_OUT_GROW - 1;
            *_pch++ = ch;
        }
    }
    return NOERROR;
}

//+------------------------------------------------------------------------
//
//  Member:     OutputStream::Put
//
//  Synopsis:   Put a string on the output stream.
//
//-------------------------------------------------------------------------
HRESULT OutputStream::Put(TCHAR* sz)
{
    HRESULT hr = NOERROR;
    while(hr==NOERROR && *sz)
    {
        hr = Put(*sz++);
    }
    return hr;
}

//+------------------------------------------------------------------------
//
//  Member:     OutputStream::Nuke
//
//  Synopsis:   Nuke the output stream. Called in case of error.
//
//-------------------------------------------------------------------------
void OutputStream::Nuke()
{
    if(!_fAlloc)
    {
        *_pchBase = 0;
    }
    else if(*_ppchAlloc)
    {
        MemFree(*_ppchAlloc);
        *_ppchAlloc = 0;
    }
}

//+------------------------------------------------------------------------
//
//  Function:   FormatDecimal
//
//  Synopsis:   Format adwArg[0] as decimal to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatDecimal(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    BOOL    fUnsigned = FALSE;
    TCHAR   szBuf[33];
    int     cWidth = 0;
    int     i;

    if(**ppchFmt == TEXT('u'))
    {
        fUnsigned = TRUE;
        *ppchFmt += 1;
    }
    if(InRange(**ppchFmt, _T('0'), _T('9')))
    {
        cWidth = **ppchFmt - _T('0');
        *ppchFmt += 1;
    }

    if(fUnsigned)
    {
        _ultot(adwArg[0], szBuf, 10);
    }
    else
    {
        _ltot(adwArg[0], szBuf, 10);
    }

    if(cWidth > 0)
    {
        for(i=cWidth-_tcslen(szBuf); --i>=0;)
        {
            pOutput->Put(_T("0"));
        }
    }

    return pOutput->Put(szBuf);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatHex
//
//  Synopsis:   Format adwArg[0] as hex to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatHex(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    TCHAR   szBuf[33];
    int     i;

    _ultot(adwArg[0], szBuf, 16);
    for(i=8-_tcsclen(szBuf); i>0; i--)
    {
        pOutput->Put(_T('0'));
    }
    return pOutput->Put(szBuf);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatColor
//
//  Synopsis:   Format adwArg[0] as a 6-digit hex to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatColor(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    TCHAR   szBuf[33];
    long    lRGB = (long)adwArg[0];
    long    lBGR = ((lRGB&0xffL)<<16)|(lRGB&0xff00L)|((lRGB&0xff0000L)>>16);

    pOutput->Put(_T('#'));
    _ultot(lBGR, szBuf, 16);
    CharLower(szBuf);
    for(int i=6-_tcsclen(szBuf); i>0; i--)
    {
        pOutput->Put(_T('0'));
    }
    return pOutput->Put(szBuf);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatPound3
//
//  Synopsis:   Format adwArg[0] as a 3-digit hex to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatPound3(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    TCHAR   szBuf[33];
    long    lRGB = (long)adwArg[0];
    long    lBGR = ((lRGB&0xf)<<8)|((lRGB&0xf00)>>4)|((lRGB&0xf0000)>>16);

    pOutput->Put(_T('#'));
    _ultot(lBGR, szBuf, 16);
    CharLower(szBuf);
    for(int i=3-_tcsclen(szBuf); i>0; i--)
    {
        pOutput->Put(_T('0'));
    }
    return pOutput->Put(szBuf);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatString
//
//  Synopsis:   Format adwArg[0] as string ptr to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatString(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    return pOutput->Put((TCHAR*)adwArg[0]);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatGuid
//
//  Synopsis:   Format adwArg[0] as guid ptr to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatGuid(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    TCHAR szBuf[40];
    Verify(StringFromGUID2(*(CLSID*)adwArg[0], szBuf, ARRAYSIZE(szBuf)) > 0);
    return pOutput->Put(szBuf);
}

//+------------------------------------------------------------------------
//
//  Function:   FormatRsrc
//
//  Synopsis:   Format resource string to the output stream.
//              adwArg[0] is the HINSTANCE and adArg[1] is the string id.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatRsrc(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    HRESULT hr;
    int     cch;
    TCHAR*  psz;

    hr = LoadString((HINSTANCE)adwArg[0], adwArg[1], &cch, &psz);
    Assert(!hr);
    while(hr==NOERROR && --cch>=0)
    {
        hr = pOutput->Put(*psz++);
    }
    return hr;
}

//+------------------------------------------------------------------------
//
//  Function:   FormatPlural
//
//  Synopsis:   Format adwArg[0] as plural to the output stream.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
static HRESULT FormatPlural(TCHAR** ppchFmt, OutputStream* pOutput, DWORD_PTR* adwArg)
{
    HRESULT hr = NOERROR;
    TCHAR*  pch = *ppchFmt;
    TCHAR   chDelim;
    TCHAR   ch;

    chDelim = *pch++;
    if(adwArg[0] == 1)
    {
        while(*pch && (ch=*pch++)!=chDelim)
        {
            hr = pOutput->Put(ch);
            if(hr)
            {
                return hr;
            }

        }
        while(*pch && *pch++!=chDelim);
    }
    else
    {
        while(*pch && *pch++!=chDelim);

        while(*pch && (ch=*pch++)!=chDelim)
        {
            hr = pOutput->Put(ch);
            if(hr)
            {
                return hr;
            }
        }
    }
    *ppchFmt = pch;
    return NOERROR;
}

//+------------------------------------------------------------------------
//
//  Function:   Format
//
//  Synopsis:   Replacement for sprintf and its friends.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
HRESULT Format(HINSTANCE hInstResource, DWORD dwOptions, void* pvOutput, int cchOutput, TCHAR* pchFmt, ...)
{
    HRESULT hr;
    va_list arg;

    va_start(arg, pchFmt);
    hr = VFormat(hInstResource, dwOptions, pvOutput, cchOutput, pchFmt, &arg);
    va_end(arg);
    return hr;
}

//+------------------------------------------------------------------------
//
//  Function:   Format
//
//  Synopsis:   Replacement for vsprintf and its friends.
//              See block comment at top of this file for more discussion.
//
//-------------------------------------------------------------------------
HRESULT VFormat(HINSTANCE hInstResource, DWORD dwOptions, void* pvOutput, int cchOutput, TCHAR* pchFmt, void* pvArgs)
{
    HRESULT     hr = NOERROR;
    int         i;
    int         cch;
    TCHAR *     pchFmtEnd;      // pointer location past end of format string.
    DWORD_PTR*  adwArg;         // pointer to argument array
    DWORD_PTR   adwArgBuf[10];  // buffer for va_list args
    HRESULT (*pfn)(TCHAR**, OutputStream*, DWORD_PTR*);
    OutputStream output;

    // Step 1: Setup argument array pointer, pdwArgs.
    if(dwOptions & FMT_ARG_ARRAY)
    {
        adwArg = (DWORD_PTR*)pvArgs;
    }
    else
    {
        // Fetch arguments from variable length argument list.
        // Two assumptons are made here that are not generally true
        // of floating point types:
        //  - sizeof(argument) == sizeof(DWORD_PTR)
        //  - va_arg(arg, DWORD_PTR) does the right thing for all types.
        adwArg = adwArgBuf;
        for(i=0; i<ARRAYSIZE(adwArgBuf); i++)
        {
            adwArgBuf[i] = va_arg(*(va_list*)pvArgs, DWORD_PTR);
        }
    }

    // Step 2: Setup source string pointers, pchFmt & pchFmtEnd.
    if(!IS_INTRESOURCE(pchFmt))
    {
        pchFmtEnd = pchFmt + _tcslen(pchFmt);
    }
    else
    {
        hr = LoadString(hInstResource, (UINT)(UINT_PTR)pchFmt, &cch, &pchFmt);
        if(hr)
        {
            return hr;
        }

        pchFmtEnd = pchFmt + cch;
    }

    // Step 3: Do the formatting.
    hr = output.Init(dwOptions&FMT_OUT_ALLOC, pvOutput, cchOutput);
    if(hr)
    {
        return hr;
    }

    while(pchFmt < pchFmtEnd)
    {
        if(*pchFmt != TEXT('<'))
        {
            hr = output.Put(*pchFmt++);
            if(hr)
            {
                goto Cleanup;
            }
        }
        else
        {
            pchFmt += 1;
            if(!InRange(*pchFmt, _T('0'), _T('9')))
            {
                hr = output.Put(*pchFmt++);
                if(hr)
                {
                    goto Cleanup;
                }
            }
            else
            {
                i = *pchFmt++ - TEXT('0');
                if(InRange(*pchFmt, _T('0'), _T('9')))
                {
                    i = i*10 + *pchFmt++ - TEXT('0');
                }
                Assert(i < ARRAYSIZE(adwArgBuf));

                switch(*pchFmt++)
                {
                case 'd': pfn = FormatDecimal; break;
                case 's': pfn = FormatString;  break;
                case 'p': pfn = FormatPlural;  break;
                case 'i': pfn = FormatRsrc;    break;
                case 'g': pfn = FormatGuid;    break;
                case 'x': pfn = FormatHex;     break;
                case 'c': pfn = FormatColor;   break;
                case 'C': pfn = FormatPound3;  break;
                default:
                    Assert(0 && "Format: Unknown format type.");
                    hr = E_FAIL;
                    goto Cleanup;
                    break;
                }
                hr = pfn(&pchFmt, &output, &adwArg[i]);
                if(hr)
                {
                    goto Cleanup;
                }
                if(*pchFmt++ != TEXT('>'))
                {
                    Assert(0 && "Format: Unknown option.");
                    hr = E_FAIL;
                    goto Cleanup;
                }
            }
        }
    }

    for(i=dwOptions&FMT_EXTRA_NULL_MASK; i>=0; i--)
    {
        hr = output.Put((TCHAR)0);
        if(hr)
        {
            goto Cleanup;
        }
    }

    return S_OK;

Cleanup:
    output.Nuke();
    return hr;
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsAllocString
//
//  Synopsis:   Allocs a BSTR and initializes it from a string.  If the
//              initializer is NULL or the empty string, the resulting bstr is
//              NULL.
//
//  Arguments:  [pch]   -- String to initialize BSTR.
//              [pBSTR] -- The result.
//
//  Returns:    HRESULT.
//
//  Modifies:   [pBSTR]
//
//  History:    5-06-94   adams   Created
//
//----------------------------------------------------------------------------
HRESULT FormsAllocStringW(LPCWSTR pch, BSTR* pBSTR)
{
    HRESULT hr;

    Assert(pBSTR);
    if(!pch || !*pch)
    {
        *pBSTR = NULL;
        return S_OK;
    }
    *pBSTR = SysAllocString(pch);
    hr = *pBSTR ? S_OK : E_OUTOFMEMORY;
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsAllocStringLen
//
//  Synopsis:   Allocs a BSTR of [uc] + 1 OLECHARS, and
//              initializes it from an optional string.  If [uc] == 0, the
//              resulting bstr is NULL.
//
//  Arguments:  [pch]   -- String to initialize.
//              [uc]    -- Count of characters of string.
//              [pBSTR] -- The result.
//
//  Returns:    HRESULT.
//
//  Modifies:   [pBSTR].
//
//  History:    5-06-94   adams   Created
//
//----------------------------------------------------------------------------
HRESULT FormsAllocStringLenW(LPCWSTR pch, UINT uc, BSTR* pBSTR)
{
    HRESULT hr;

    Assert(pBSTR);
    if(uc == 0)
    {
        *pBSTR = NULL;
        return S_OK;
    }

    *pBSTR = SysAllocStringLen(pch, uc);
    hr = *pBSTR ? S_OK : E_OUTOFMEMORY;
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsReAllocString
//
//  Synopsis:   Allocates a BSTR initialized from a string; if successful,
//              frees the original string and replaces it.
//
//  Arguments:  [pBSTR] -- String to reallocate.
//              [pch]   -- Initializer.
//
//  Returns:    HRESULT.
//
//  Modifies:   [pBSTR].
//
//  History:    5-06-94   adams   Created
//
//----------------------------------------------------------------------------
HRESULT FormsReAllocStringW(BSTR* pBSTR, LPCWSTR pch)
{
    Assert(pBSTR);

    return SysReAllocString(pBSTR, pch)?S_OK:E_OUTOFMEMORY;
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsReAllocStringLen
//
//  Synopsis:   Allocates a BSTR of [uc] + 1 OLECHARs and optionally
//              initializes it from a string; if successful, frees the original
//              string and replaces it.
//
//  Arguments:  [pBSTR] -- String to reallocate.
//              [pch]   -- Initializer.
//              [uc]    -- Count of characters.
//
//  Returns:    HRESULT.
//
//  Modifies:   [pBSTR].
//
//  History:    5-06-94   adams   Created
//
//----------------------------------------------------------------------------
HRESULT FormsReAllocStringLenW(BSTR* pBSTR, LPCWSTR pch, UINT uc)
{
    Assert(pBSTR);

    return SysReAllocStringLen(pBSTR, pch, uc)?S_OK:E_OUTOFMEMORY;
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringLen
//
//  Synopsis:   Returns the length of the BSTR.
//
//  History:    5-06-94   adams   Created
//              6-30-95   andrewl Changed BSTR to const BSTR
//
//----------------------------------------------------------------------------
UINT FormsStringLen(const BSTR bstr)
{
    return bstr?SysStringLen((BSTR)bstr):0;
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringByteLen
//
//  Synopsis:   Returns the length of a BSTR in bytes.
//
//  History:    5-06-94   adams   Created
//              6-30-95   andrewl Changed BSTR to const BSTR
//
//----------------------------------------------------------------------------
UINT FormsStringByteLen(const BSTR bstr)
{
    return bstr?SysStringByteLen((BSTR)bstr):0;
}