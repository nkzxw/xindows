
#ifndef __XINDOWS_CORE_UTIL_FORMSSTRING_H__
#define __XINDOWS_CORE_UTIL_FORMSSTRING_H__

int     FormsStringCmp(LPCTSTR bstr1, LPCTSTR bstr2);
int     FormsStringNCmp(LPCTSTR bstr1, int cch1, LPCTSTR bstr2, int cch2);
int     FormsStringICmp(LPCTSTR bstr1, LPCTSTR bstr2);
int     FormsStringNICmp(LPCTSTR bstr1, int cch1, LPCTSTR bstr2, int cch2);
int     FormsStringCmpCase(LPCTSTR bstr1, LPCTSTR bstr2, BOOL fCaseSensitive);
void    FormsSplitAtDelimiter(LPCTSTR bstrName, BSTR* pbstrHead, BSTR* pbstrTail,
                              BOOL fFirst=TRUE, TCHAR tchDelim=_T('.'));

//+-------------------------------------------------------------
//
//  Function:   STRVAL
//
//  Synopsis:   Returns the string unless it is NULL, in which
//              case the empty string is returned.
//
//--------------------------------------------------------------
inline LPCTSTR STRVAL(LPCTSTR lpcwsz)
{
    return lpcwsz?lpcwsz:_afxGlobalData._Zero.ach;
}

//+-------------------------------------------------------------
//
//  Function:   FormsIsEmptyString
//
//  Synopsis:   Returns TRUE if the string is empty, FALSE otherwise.
//
//--------------------------------------------------------------
inline BOOL FormsIsEmptyString(LPCTSTR lpcwsz)
{
    return !(lpcwsz&&lpcwsz[0]);
}

//+-------------------------------------------------------------
//
//  Function:   FormsFreeString
//
//  Synopsis:   Frees a BSTR.
//
//--------------------------------------------------------------
inline void FormsFreeString(BSTR bstr)
{
    SysFreeString(bstr);
}

//+-------------------------------------------------------------
//
//  Function:   FormsSplitFirstComponent
//
//  Synopsis:   Split a string into the first component (up to the first
//              delimiter) and everything else.
//
//--------------------------------------------------------------
inline void FormsSplitFirstComponent(LPCTSTR bstrName, BSTR* pbstrHead,
                                     BSTR* pbstrTail, TCHAR tchDelim=_T('.'))
{
    FormsSplitAtDelimiter(bstrName, pbstrHead, pbstrTail, TRUE, tchDelim);
}

//+-------------------------------------------------------------
//
//  Function:   FormsSplitLastComponent
//
//  Synopsis:   Split a string into everything else and the last component
//              (after the first delimiter).
//
//--------------------------------------------------------------
inline void FormsSplitLastComponent(LPCTSTR bstrName, BSTR* pbstrHead,
                                    BSTR* pbstrTail, TCHAR tchDelim=_T('.'))
{
    FormsSplitAtDelimiter(bstrName, pbstrHead, pbstrTail, FALSE, tchDelim);
}

HRESULT TaskAllocString(const TCHAR* pstrSrc, TCHAR** ppstrDest);
HRESULT TaskReplaceString(const TCHAR* pstrSrc, TCHAR** ppstrDest);
inline void TaskFreeString(void* pstr) { CoTaskMemFree(pstr); }


BOOL _tcsequal(const TCHAR* string1, const TCHAR* string2);
BOOL _tcsiequal(const TCHAR* string1, const TCHAR* string2);
BOOL _tcsnpre(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2);
BOOL _tcsnipre(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2);
BOOL _7csnipre(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2);

const TCHAR* __cdecl _tcsistr(const TCHAR* wcs1,const TCHAR* wcs2);
int _cdecl _tcsncmp(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2);
int _cdecl _tcsnicmp(const TCHAR* string1, int cch1, const TCHAR* string2, int cch2);

// Special implementation _ttol except it returns an error if the string to
// convert isn't a number.
HRESULT ttol_with_error(LPCTSTR pStr, long* plValue);

#endif //__XINDOWS_CORE_UTIL_FORMSSTRING_H__