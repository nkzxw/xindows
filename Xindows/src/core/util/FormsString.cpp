
#include "stdafx.h"
#include "FormsString.h"

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringCmp
//
//  Synopsis:   As per _tcscmp, checking for NULL bstrs.
//
//  History:    5-06-94   adams   Created
//              25-Jun-94 doncl   changed from _tc to wc
//
//----------------------------------------------------------------------------
int FormsStringCmp(LPCTSTR bstr1, LPCTSTR bstr2)
{
    return _tcscmp(STRVAL(bstr1), STRVAL(bstr2));
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringNCmp
//
//  Synopsis:   As per _tcsncmp, checking for NULL bstrs.
//
//  History:    5-06-94   adams   Created
//              25-Jun-94 doncl   changed from _tc to wc
//
//----------------------------------------------------------------------------
int FormsStringNCmp(LPCTSTR bstr1, int cch1, LPCTSTR bstr2, int cch2)
{
    return _tcsncmp(STRVAL(bstr1), STRVAL(bstr2), min(cch1, cch2));
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringICmp
//
//  Synopsis:   As per wcsicmp, checking for NULL bstrs.
//
//  History:    5-06-94   adams   Created
//              25-Jun-94 doncl   changed from _tc to wc
//              15-Aug-94 doncl   changed from wcsicmp to _tcsicmp
//
//----------------------------------------------------------------------------
int FormsStringICmp(LPCTSTR bstr1, LPCTSTR bstr2)
{
    return _tcsicmp(STRVAL(bstr1), STRVAL(bstr2));
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringNICmp
//
//  Synopsis:   As per wcsnicmp, checking for NULL bstrs.
//
//  History:    5-06-94   adams   Created
//              25-Jun-94 doncl   changed from _tc to wc
//              15-Aug-94 doncl   changed from wcsnicmp to _tcsnicmp
//
//----------------------------------------------------------------------------
int FormsStringNICmp(LPCTSTR bstr1, int cch1, LPCTSTR bstr2, int cch2)
{
    return _tcsnicmp(STRVAL(bstr1), cch1, STRVAL(bstr2), min(cch1, cch2));
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsStringCmpCase
//
//----------------------------------------------------------------------------
int FormsStringCmpCase(LPCTSTR bstr1, LPCTSTR bstr2, BOOL fCaseSensitive)
{
    return (fCaseSensitive)?
        _tcscmp(STRVAL(bstr1), STRVAL(bstr2)):_tcsicmp(STRVAL(bstr1), STRVAL(bstr2));
}

//+-------------------------------------------------------------------------
// Function:    FormsSplitAtDelimiter
//
// Synopsis:    split a name into its head component (everything before the first
//              dot), and the tail component (the rest).
//
// Arguments:	bstrName    name to be split
//              pbstrHead   where to store head component
//              pbstrTail   where to store the rest
//              fFirst      TRUE - split at first delimiter, FALSE - at last
//              tchDelim    delimiter character (defaults to _T('.'))
//----------------------------------------------------------------------------
void FormsSplitAtDelimiter(LPCTSTR bstrName, BSTR* pbstrHead, BSTR* pbstrTail,
                           BOOL fFirst, TCHAR tchDelim)
{
    if(FormsIsEmptyString(bstrName))
    {
        *pbstrHead = NULL;
        *pbstrTail = NULL;
    }
    else
    {
        TCHAR* ptchDelim = fFirst ?  _tcschr((LPTSTR)bstrName, tchDelim)
            : _tcsrchr((LPTSTR)bstrName, tchDelim);

        if(ptchDelim)
        {
            FormsAllocStringLen(bstrName, ptchDelim-bstrName, pbstrHead);
            FormsAllocString(ptchDelim+1, pbstrTail);
        }
        else if(fFirst)
        {
            FormsAllocString(bstrName, pbstrHead);
            *pbstrTail = NULL;
        }
        else
        {
            *pbstrHead = NULL;
            FormsAllocString(bstrName, pbstrTail);
        }
    }
}

//+------------------------------------------------------------------------
//
//  Function:   TaskAllocString
//
//  Synopsis:   Allocates a string copy that can be passed across an interface
//              boundary, using the standard memory allocation conventions.
//
//              The inline function TaskFreeString is provided for symmetry.
//
//  Arguments:  pstrSrc    String to copy
//              ppstrDest  Copy of string is returned in *ppstr
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT TaskAllocString(const TCHAR* pstrSrc, TCHAR** ppstrDest)
{
    TCHAR* pstr;
    size_t cb;

    cb = (_tcsclen(pstrSrc)+1) * sizeof(TCHAR);
    *ppstrDest = pstr = (TCHAR*)CoTaskMemAlloc(cb);
    if(!pstr)
    {
        return E_OUTOFMEMORY;
    }
    else
    {
        memcpy(pstr, pstrSrc, cb);
        return S_OK;
    }
}

//+------------------------------------------------------------------------
//
//  Function:   TaskReplaceString
//
//  Synopsis:   Replaces a string copy that can be passed across an interface
//              boundary, using the standard memory allocation conventions.
//
//              The inline function TaskFreeString is provided for symmetry.
//
//  Arguments:  pstrSrc    String to copy. May be NULL.
//              ppstrDest  Copy of string is returned in *ppstrDest,
//                         previous string is freed on success
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT TaskReplaceString(const TCHAR* pstrSrc, TCHAR** ppstrDest)
{
    TCHAR* pstr;
    HRESULT hr;

    if(pstrSrc)
    {
        hr = TaskAllocString(pstrSrc, &pstr);
        if(hr)
        {
            RRETURN(hr);
        }
    }
    else
    {
        pstr = NULL;
    }

    CoTaskMemFree(*ppstrDest);
    *ppstrDest = pstr;

    return S_OK;
}

BOOL _tcsequal(const TCHAR* string1, const TCHAR* string2)
{
    // This function optimizes the case where all we want to find
    // is if the strings are equal and do not care about the relative
    // ordering of the strings.
    while(*string1)
    {
        if(*string1 != *string2)
        {
            return FALSE;
        }

        string1 += 1;
        string2 += 1;

    }
    return (*string2)?(FALSE):(TRUE);
}

BOOL _tcsiequal(const TCHAR* string1, const TCHAR* string2)
{
    // This function optimizes the case where all we want to find
    // is if the strings are equal and do not care about the relative
    // ordering of the strings (or CaSe).
    while(*string1)
    {
        TCHAR ch1 = *string1;
        TCHAR ch2 = *string2;

        CharLowerBuff(&ch1, 1);
        CharLowerBuff(&ch2, 1);

        if(ch1 != ch2)
        {
            return FALSE;
        }
        string1 += 1;
        string2 += 1;
    }
    return (*string2)?(FALSE):(TRUE);
}

BOOL _tcsnpre(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2)
{
    if(cch1 == -1)
    {
        cch1 = _tcslen(string1);
    }
    if(cch2 == -1)
    {
        cch2 = _tcslen(string2);
    }
    return (cch1<=cch2 && CompareString(_afxGlobalData._lcidUserDefault,
        0, string1, cch1, string2, cch1)==CSTR_EQUAL);
}

BOOL _tcsnipre(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2)
{
    if(cch1 == -1)
    {
        cch1 = _tcslen(string1);
    }
    if(cch2 == -1)
    {
        cch2 = _tcslen(string2);
    }

    return (cch1<=cch2 && CompareString(_afxGlobalData._lcidUserDefault,
        NORM_IGNORECASE|NORM_IGNOREWIDTH|NORM_IGNOREKANATYPE,
        string1, cch1, string2, cch1)==CSTR_EQUAL);
}

BOOL _7csnipre(const TCHAR* pch1, int cch1, const TCHAR* pch2, int cch2)
{
    if(cch1 == -1)
    {
        cch1 = _tcslen(pch1);
    }
    if(cch2 == -1)
    {
        cch2 = _tcslen(pch2);
    }
    if(cch1 <= cch2)
    {
        while(cch1)
        {
            TCHAR ch1 = *pch1>=_T('a')&&*pch1<=_T('z') ? 
                *pch1+_T('A')-_T('a') : *pch1;
            TCHAR ch2 = *pch2>=_T('a')&&*pch2<=_T('z') ?
                *pch2+_T('A')-_T('a') : *pch2;

            if(ch1 != ch2)
            {
                return FALSE;
            }

            pch1 += 1;
            pch2 += 1;
            cch1--;
        }
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

const TCHAR* __cdecl _tcsistr(const TCHAR* tcs1,const TCHAR* tcs2)
{
    const TCHAR* cp;
    int cc,count;
    int n2Len = _tcslen(tcs2);
    int n1Len = _tcslen(tcs1);

    if(n1Len >= n2Len)
    {
        for(cp=tcs1,count=n1Len-n2Len; count>=0; cp++,count--)
        {
            cc = CompareString(_afxGlobalData._lcidUserDefault,
                NORM_IGNORECASE|NORM_IGNOREWIDTH|NORM_IGNOREKANATYPE,
                cp, n2Len,tcs2, n2Len);
            if(cc > 0)
            {
                cc -= 2;
            }
            if(!cc)
            {
                return cp;
            }
        }
    }
    return NULL;
}

int __cdecl _tcsncmp(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2)
{
    int cc;

    cc = CompareString(
        _afxGlobalData._lcidUserDefault,
        0,
        string1, cch1,
        string2, cch2);

    if(cc > 0)
    {
        return (cc-2);
    }
    else
    {
        return _NLSCMPERROR;
    }
}

int __cdecl _tcsnicmp(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2)
{
    int cc;

    cc = CompareString(_afxGlobalData._lcidUserDefault,
        NORM_IGNORECASE|NORM_IGNOREWIDTH|NORM_IGNOREKANATYPE,
        string1, cch1, string2, cch2);

    if(cc > 0)
    {
        return (cc-2);
    }
    else
    {
        return _NLSCMPERROR;
    }
}

/* flag values */
#define FL_UNSIGNED   1       /* wcstoul called */
#define FL_NEG        2       /* negative sign found */
#define FL_OVERFLOW   4       /* overflow occured */
#define FL_READDIGIT  8       /* we've read at least one correct digit */

static HRESULT PropertyStringToLong(LPCTSTR nptr, TCHAR** endptr, int ibase,
                                    int flags, unsigned long* plNumber)
{
    const TCHAR*    p;
    TCHAR           c;
    unsigned long   number;
    unsigned        digval;
    unsigned long   maxval;

    *plNumber = 0;                  /* on error result is 0 */

    p = nptr;                       /* p is our scanning pointer */
    number = 0;                     /* start with zero */

    c = *p++;                       /* read char */
    while(_istspace(c))
    {
        c = *p++;                   /* skip whitespace */
    }

    if(c == '-') 
    {
        flags |= FL_NEG;			/* remember minus sign */
        c = *p++;
    }
    else if(c == '+')
    {
        c = *p++;					/* skip sign */
    }

    if(ibase<0 || ibase==1 || ibase>36) 
    {
        /* bad base! */
        if(endptr)
        {
            /* store beginning of string in endptr */
            *endptr = (TCHAR*)nptr;
        }
        return E_INVALIDARG;        /* return 0 */
    }
    else if(ibase == 0) 
    {
        /* determine base free-lance, based on first two chars of
        string */
        if(c != L'0')
        {
            ibase = 10;
        }
        else if(*p==L'x' || *p==L'X')
        {
            ibase = 16;
        }
        else
        {
            ibase = 8;
        }
    }

    if(ibase == 16) 
    {
        /* we might have 0x in front of number; remove if there */
        if(c==L'0' && (*p==L'x' || *p==L'X')) 
        {
            ++p;
            c = *p++;				/* advance past prefix */
        }
    }

    /* if our number exceeds this, we will overflow on multiply */
    maxval = ULONG_MAX / ibase;

    for(;;)
    {      /* exit in middle of loop */
        /* convert c to value */
        if(_istdigit(c))
        {
            digval = c - L'0';
        }
        else if(_istalpha(c))
        {
            if(ibase > 10)
            {
                if(c>='a' && c<='z')
                {
                    digval = (unsigned)c - 'a' + 10;
                }
                else
                {
                    digval = (unsigned)c - 'A' + 10;
                }
            }
            else
            {
                return E_INVALIDARG; /* return 0 */
            }
        }
        else
        {
            break;
        }

        if(digval >= (unsigned)ibase)
        {
            break;					/* exit loop if bad digit found */
        }

        /* record the fact we have read one digit */
        flags |= FL_READDIGIT;

        /* we now need to compute number = number * base + digval,
        but we need to know if overflow occured.  This requires
        a tricky pre-check. */
        if(number<maxval || (number==maxval && (unsigned long)digval<=ULONG_MAX%ibase)) 
        {
            /* we won't overflow, go ahead and multiply */
            number = number*ibase + digval;
        }
        else 
        {
            /* we would have overflowed -- set the overflow flag */
            flags |= FL_OVERFLOW;
        }

        c = *p++;					/* read next digit */
    }

    --p;							/* point to place that stopped scan */

    if(!(flags & FL_READDIGIT)) 
    {
        number = 0L;				/* return 0 */

        /* no number there; return 0 and point to beginning of
        string */
        if(endptr)
        {
            /* store beginning of string in endptr later on */
            p = nptr;
        }

        return E_INVALIDARG;		/*Return error not a number*/
    }
    else if((flags&FL_OVERFLOW) ||
        (!(flags&FL_UNSIGNED) &&
        (((flags&FL_NEG)&&(number>-LONG_MIN)) ||
        (!(flags&FL_NEG)&&(number>LONG_MAX)))))
    {
        /* overflow or signed overflow occurred */
        //errno = ERANGE;
        if(flags & FL_UNSIGNED)
        {
            number = ULONG_MAX;
        }
        else if(flags & FL_NEG)
        {
            number = (unsigned long)(-LONG_MIN);
        }
        else
        {
            number = LONG_MAX;
        }
    }

    if(endptr != NULL)
    {
        /* store pointer to char that stopped the scan */
        *endptr = (TCHAR*)p;
    }

    if(flags & FL_NEG)
    {
        /* negate result if there was a neg sign */
        number = (unsigned long)(-(long)number);
    }

    *plNumber = number; 
    return S_OK;                    /* done. */
}

HRESULT ttol_with_error(LPCTSTR pStr, long* plValue)
{
    // Always do base 10 regardless of contents of 
    RRETURN1(PropertyStringToLong(pStr, NULL, 10, 0, (unsigned long*)plValue), E_INVALIDARG);
}