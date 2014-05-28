
#include "stdafx.h"
#include "String.h"

//+---------------------------------------------------------------------------
//
//  Member:     CString::Set, public
//
//  Synopsis:   Allocates memory for a string and initializes it from a given
//              string.
//
//  Arguments:  [pch] -- String to initialize with. Can be NULL
//              [uc]  -- Number of characters to allocate.
//
//  Returns:    HRESULT
//
//  Notes:      The total number of characters allocated is uc+1, to allow
//              for a NULL terminator. If [pch] is NULL, then the string is
//              uninitialized except for the NULL terminator, which is at
//              character position [uc].
//
//  Mac note:   The total number of characters allocated is increased by the
//              size of a pointer.  This pointer, located prior to the length
//              field, will point to a memory allocation that contains the
//              CHAR or WCHAR string alternative to the _pch string.
//
//----------------------------------------------------------------------------
HRESULT CString::Set(LPCTSTR pch, UINT uc)
{
    if(pch == _pch)
    {
        if(uc == Length())
        {
            return S_OK;
        }
        // when the ptrs are the same the length can only be
        // different if the ptrs are NULL.  this is a hack used
        // internally to implement realloc type expansion
        Assert(pch==NULL && _pch==NULL);
    }

    Free();

    BYTE* p = new BYTE[sizeof(TCHAR)+sizeof(TCHAR)*uc+sizeof(UINT)];
    if(p)
    {
        *(UINT*)(p) = uc * sizeof(TCHAR);
        _pch = (TCHAR*)(p+sizeof(UINT));
        if(pch)
        {
            _tcsncpy(_pch, pch, uc);
        }

        ((TCHAR*)_pch)[uc] = 0;
    }
    else
    {
        return E_OUTOFMEMORY;
    }
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Set, public
//
//  Synopsis:   Allocates a string and initializes it
//
//  Arguments:  [pch] -- String to initialize from
//
//----------------------------------------------------------------------------
HRESULT CString::Set(LPCTSTR pch)
{
    RRETURN(Set(pch, pch?_tcsclen(pch):0));
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Set, public
//
//  Synopsis:   Allocates a string and initializes it
//
//  Arguments:  [cstr] -- String to initialize from
//
//----------------------------------------------------------------------------
HRESULT CString::Set(const CString& cstr)
{
    RRETURN(Set(cstr, cstr.Length()));
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::TakeOwnership, public
//
//  Synopsis:   Takes the ownership of a string from another CString class.
//
//  Arguments:  [cstr] -- Class to take string from
//
//  Notes:      This method just transfers a string from one CString class to
//              another. The class which is the source of the transfer has
//              a NULL value at the end of the operation.
//
//----------------------------------------------------------------------------
void CString::TakeOwnership(CString& cstr)
{
    _Free();
    _pch = cstr._pch;
    cstr._pch = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::SetBSTR, public
//
//  Synopsis:   Allocates a string and initializes it from a BSTR
//
//  Arguments:  [bstr] -- Initialization string
//
//  Notes:      This method is more efficient than Set(LPCWSTR pch) because
//              of the length-prefix on BSTRs.
//
//----------------------------------------------------------------------------
HRESULT CString::SetBSTR(const BSTR bstr)
{
    RRETURN(Set(bstr, bstr?SysStringLen(bstr):0));
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Length, public
//
//  Synopsis:   Returns the length of the string associated with this class
//
//----------------------------------------------------------------------------
UINT CString::Length() const
{
    if(_pch)
    {
        return (((UINT*)_pch)[-1])/sizeof(TCHAR);
    }
    else
    {
        return 0;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::ReAlloc, public
//
//  Synopsis:   Reallocate the string to a different size buffer.
//              The length of the string is not affected, but it is allocated
//              within a larger (or maybe smaller) memory allocation.
//
//----------------------------------------------------------------------------
HRESULT CString::ReAlloc(UINT uc)
{
    HRESULT hr;
    TCHAR*  pchOld;
    UINT    ubOld;

    Assert(uc >= Length());     // Disallowed to allocate a too-short buffer.

    if(uc)
    {
        pchOld = _pch;          // Save pointer to old string.
        _pch = 0;               // Prevent Set from Freeing the string.
        ubOld = pchOld ?        // Save old length
            *(UINT*)(((BYTE*)pchOld)-sizeof(UINT)) : 0;

        hr = Set(pchOld, uc);   // Alloc new; Copy old string.
        if(hr)
        {
            _pch = pchOld;
            RRETURN(hr);
        }
        *(UINT*)(((BYTE *)_pch)-sizeof(UINT)) = ubOld; // Restore length.

        if(pchOld)
        {
            delete[] (((BYTE*)pchOld)-sizeof(UINT));
        }
    }

    // else if uc == 0, then, since we have already checked that uc >= Length,
    // length must == 0.
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Append
//
//  Synopsis:   Append chars to the end of the string, reallocating & updating
//              its length.
//
//----------------------------------------------------------------------------
HRESULT CString::Append(LPCTSTR pch, UINT uc)
{
    HRESULT hr = S_OK;
    UINT    ucOld, ucNew;
    BYTE*   p;

    if(uc)
    {
        ucOld = Length();
        ucNew = ucOld + uc;
        hr = ReAlloc(ucNew);
        if(hr)
        {
            goto Cleanup;
        }
        _tcsncpy(_pch+ucOld, pch, uc);
        ((TCHAR*)_pch)[ucNew] = 0;
        p = ((BYTE*)_pch - sizeof(UINT));
        *(UINT*)p = ucNew * sizeof(TCHAR);
    }
Cleanup:
    RRETURN(hr);
}

HRESULT CString::Append(LPCTSTR pch)
{
    RRETURN(Append(pch, pch?_tcsclen(pch):0));
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::SetLengthNoAlloc, public
//
//  Synopsis:   Sets the internal length of the string. Note.  There is no
//              verification that the length that is set is within the allocated
//              range of the string. If the caller sets the length too large,
//              this blasts a null byte into memory.
//
//----------------------------------------------------------------------------
HRESULT CString::SetLengthNoAlloc(UINT uc)
{
    if(_pch)
    {
        BYTE* p = ( (BYTE*)_pch - sizeof(UINT));
        *(UINT*)p = uc * sizeof(TCHAR); // Set the length prefix.
        ((TCHAR*)_pch)[uc] = 0;         // Set null terminator
        return S_OK;
    }
    else 
    {
        return E_POINTER;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::AllocBSTR, public
//
//  Synopsis:   Allocates a BSTR and initializes it with the string that is
//              associated with this class.
//
//  Arguments:  [pBSTR] -- Place to put new BSTR. This pointer should not
//                         be pointing to an existing BSTR.
//
//
//  Mac note:   Mac BSTRs are ANSI so we need to convert before allocating the
//              BSTR.
//----------------------------------------------------------------------------
HRESULT CString::AllocBSTR(BSTR* pBSTR) const
{
    if(!_pch)
    {
        *pBSTR = 0;
        return S_OK;
    }

    *pBSTR = SysAllocStringLen(*this, Length());
    RRETURN(*pBSTR?S_OK:E_OUTOFMEMORY);
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::TrimTrailingWhitespace, public
//
//  Synopsis:   Removes any trailing whitespace in the string that is
//              associated with this class.
//
//  Arguments:  None.
//
//----------------------------------------------------------------------------
HRESULT CString::TrimTrailingWhitespace()
{
    if(!_pch)
    {
        return S_OK;
    }

    UINT ucNewLength = Length();

    for(; ucNewLength>0; ucNewLength--)
    {
        if(!_istspace(((TCHAR*)_pch)[ucNewLength-1]))
        {
            break;
        }
        ((TCHAR*)_pch)[ucNewLength-1] = (TCHAR)0;
    }

    BYTE* p = ((BYTE*)_pch - sizeof(UINT));
    *(UINT*)p = ucNewLength * sizeof(TCHAR);
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::_Free, private
//
//  Synopsis:   Frees any memory held onto by this class.
//
//----------------------------------------------------------------------------
void CString::_Free()
{
    if(_pch)
    {
        delete[] (((BYTE*)_pch)-sizeof(UINT));
    }
    _pch = 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Clone
//
//  Synopsis:   Make copy of current string
//
//----------------------------------------------------------------------------
HRESULT CString::Clone(CString** ppCStr) const
{
    HRESULT hr;
    Assert(ppCStr);
    *ppCStr = new CString;
    hr = *ppCStr ? S_OK : E_OUTOFMEMORY;
    if(hr)
    {
        goto Cleanup;
    }

    hr = (*ppCStr)->Set(*this);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Compare
//
//  Synopsis:   Case insensitive comparison of 2 strings
//
//----------------------------------------------------------------------------
BOOL CString::Compare(const CString* pCStr) const
{
    return (!_tcsicmp(*pCStr, *this));
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::ComputeCrc
//
//  Synopsis:   Computes a hash of the string.
//
//----------------------------------------------------------------------------
WORD CString::ComputeCrc() const
{
    WORD            wHash=0;
    const TCHAR*    pch;
    int             i;

    pch=*this;
    for(i=Length(); i>0; i--,pch++)
    {
        wHash = wHash << 7^wHash >> (16-7)^(TCHAR)CharUpper((LPTSTR)((DWORD_PTR)(*pch)));
    }

    return wHash;
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Load
//
//  Synopsis:   Initializes the CString from a stream
//
//----------------------------------------------------------------------------
HRESULT CString::Load(IStream* pstm)
{
    DWORD   cch;
    HRESULT hr;

    hr = pstm->Read(&cch, sizeof(DWORD), NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(cch == 0xFFFFFFFF)
    {
        Free();
    }
    else
    {
        hr = Set(NULL, cch);
        if(hr)
        {
            goto Cleanup;
        }

        if(cch)
        {
            hr = pstm->Read(_pch, cch*sizeof(TCHAR), NULL);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::Save
//
//  Synopsis:   Writes the contents of the CString to a stream
//
//----------------------------------------------------------------------------
HRESULT CString::Save(IStream* pstm) const
{
    DWORD   cch = _pch ? Length() : 0xFFFFFFFF;
    HRESULT hr;

    hr = pstm->Write(&cch, sizeof(DWORD), NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(cch && cch!=0xFFFFFFFF)
    {
        hr = pstm->Write(_pch, cch*sizeof(TCHAR), NULL);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CString::GetSaveSize
//
//  Synopsis:   Returns the number of bytes which will be written by
//              CString::Save
//
//----------------------------------------------------------------------------
ULONG CString::GetSaveSize() const
{
    return (sizeof(DWORD)+(_pch?(Length()*sizeof(TCHAR)):0));
}



//+---------------------------------------------------------------------------
//
//  Function:   MbcsFromUnicode
//
//  Synopsis:   Converts a string to MBCS from Unicode.
//
//  Arguments:  [pstr]  -- The buffer for the MBCS string.
//              [cch]   -- The size of the MBCS buffer, including space for
//                              NULL terminator.
//
//              [pwstr] -- The Unicode string to convert.
//              [cwch]  -- The number of characters in the Unicode string to
//                              convert, including NULL terminator.  If this
//                              number is -1, the string is assumed to be
//                              NULL terminated.  -1 is supplied as a
//                              default argument.
//
//  Returns:    If [pstr] is NULL or [cch] is 0, 0 is returned.  Otherwise,
//              the number of characters converted, including the terminating
//              NULL, is returned (note that converting the empty string will
//              return 1).  If the conversion fails, 0 is returned.
//
//  Modifies:   [pstr].
//
//----------------------------------------------------------------------------
int MbcsFromUnicode(LPSTR pstr, int cch, LPCWSTR pwstr, int cwch)
{
    int ret;

    Assert(cch >= 0);
    if(!pstr || cch==0)
    {
        return 0;
    }

    Assert(pwstr);
    Assert(cwch==-1 || cwch>0);

    ret = WideCharToMultiByte(CP_ACP, 0, pwstr, cwch, pstr, cch, NULL, NULL);
    return ret;
}

//+---------------------------------------------------------------------------
//
//  Function:   UnicodeFromMbcs
//
//  Synopsis:   Converts a string to Unicode from MBCS.
//
//  Arguments:  [pwstr] -- The buffer for the Unicode string.
//              [cwch]  -- The size of the Unicode buffer, including space for
//                              NULL terminator.
//
//              [pstr]  -- The MBCS string to convert.
//              [cch]  -- The number of characters in the MBCS string to
//                              convert, including NULL terminator.  If this
//                              number is -1, the string is assumed to be
//                              NULL terminated.  -1 is supplied as a
//                              default argument.
//
//  Returns:    If [pwstr] is NULL or [cwch] is 0, 0 is returned.  Otherwise,
//              the number of characters converted, including the terminating
//              NULL, is returned (note that converting the empty string will
//              return 1).  If the conversion fails, 0 is returned.
//
//  Modifies:   [pwstr].
//
//----------------------------------------------------------------------------
int UnicodeFromMbcs(LPWSTR pwstr, int cwch, LPCSTR pstr, int cch)
{
    int ret;

    Assert(cwch >= 0);

    if(!pwstr || cwch==0)
    {
        return 0;
    }

    Assert(pstr);
    Assert(cch==-1 || cch>0);

    ret = MultiByteToWideChar(CP_ACP, 0, pstr, cch, pwstr, cwch);

    return ret;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStr::Free
//
//  Synopsis:   Frees string if alloc'd and initializes to NULL.
//
//----------------------------------------------------------------------------
void CConvertStr::Free()
{
    if(_pstr!=_ach && HIWORD64(_pstr)!=0)
    {
        delete[] _pstr;
    }

    _pstr = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStrW::Free
//
//  Synopsis:   Frees string if alloc'd and initializes to NULL.
//
//----------------------------------------------------------------------------
void CConvertStrW::Free()
{
    if(_pwstr!=_awch && HIWORD64(_pwstr)!=0)
    {
        delete[] _pwstr;
    }

    _pwstr = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrInW::Init
//
//  Synopsis:   Converts a LPSTR function argument to a LPWSTR.
//
//  Arguments:  [pstr] -- The function argument.  May be NULL or an atom
//                              (HIWORD64(pwstr) == 0).
//
//              [cch]  -- The number of characters in the string to
//                          convert.  If -1, the string is assumed to be
//                          NULL terminated and its length is calculated.
//
//  Modifies:   [this]
//
//----------------------------------------------------------------------------
void CStrInW::Init(LPCSTR pstr, int cch)
{
    int cchBufReq;

    _cwchLen = 0;

    // Check if string is NULL or an atom.
    if(HIWORD64(pstr) == 0)
    {
        _pwstr = (LPWSTR) pstr;
        return;
    }

    Assert(cch==-1 || cch>0);

    // Convert string to preallocated buffer, and return if successful.
    //
    // Since the passed in buffer may not be null terminated, we have
    // a problem if cch==ARRAYSIZE(_awch), because MultiByteToWideChar
    // will succeed, and we won't be able to null terminate the string!
    // Decrease our buffer by one for this case.
    _cwchLen = MultiByteToWideChar(CP_ACP, 0, pstr, cch, _awch, ARRAYSIZE(_awch)-1);

    if(_cwchLen > 0)
    {
        // Some callers don't NULL terminate.
        //
        // We could check "if (-1 != cch)" before doing this,
        // but always doing the null is less code.
        _awch[_cwchLen] = 0;

        if(_awch[_cwchLen-1] == 0)
        {
            _cwchLen--; // account for terminator
        }

        _pwstr = _awch;
        return;
    }

    // Alloc space on heap for buffer.
    cchBufReq = MultiByteToWideChar(CP_ACP, 0, pstr, cch, NULL, 0);

    // Again, leave room for null termination
    cchBufReq++;

    Assert(cchBufReq > 0);
    _pwstr = new WCHAR[cchBufReq];
    if(!_pwstr)
    {
        // On failure, the argument will point to the empty string.
        _awch[0] = 0;
        _pwstr = _awch;
        return;
    }

    Assert(HIWORD64(_pwstr));
    _cwchLen = MultiByteToWideChar(CP_ACP, 0, pstr, cch, _pwstr, cchBufReq);

    // Again, make sure we're always null terminated
    Assert(_cwchLen < cchBufReq);
    _pwstr[_cwchLen] = 0;

    if(0 == _pwstr[_cwchLen-1]) // account for terminator
    {
        _cwchLen--;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrIn::CStrIn
//
//  Synopsis:   Inits the class.
//
//  NOTE:       Don't inline this function or you'll increase code size
//              by pushing -1 on the stack for each call.
//
//----------------------------------------------------------------------------
CStrIn::CStrIn(LPCWSTR pwstr) : CConvertStr(CP_ACP)
{
    Init(pwstr, -1);
}

CStrIn::CStrIn(UINT uCP, LPCWSTR pwstr) : CConvertStr(uCP==1200?CP_ACP:uCP)
{
    Init(pwstr, -1);
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrIn::Init
//
//  Synopsis:   Converts a LPWSTR function argument to a LPSTR.
//
//  Arguments:  [pwstr] -- The function argument.  May be NULL or an atom
//                              (HIWORD64(pwstr) == 0).
//
//              [cwch]  -- The number of characters in the string to
//                          convert.  If -1, the string is assumed to be
//                          NULL terminated and its length is calculated.
//
//  Modifies:   [this]
//
//----------------------------------------------------------------------------
void CStrIn::Init(LPCWSTR pwstr, int cwch)
{
    int cchBufReq;

    _cchLen = 0;

    // Check if string is NULL or an atom.
    if(HIWORD64(pwstr) == 0)
    {
        _pstr = (LPSTR) pwstr;
        return;
    }

    if(cwch == 0)
    {
        *_ach = '\0';
        _pstr = _ach;
        return;
    }

    Assert(cwch==-1 || cwch>0);
    // Convert string to preallocated buffer, and return if successful.
    _cchLen = WideCharToMultiByte(_uCP, 0, pwstr, cwch, _ach, ARRAYSIZE(_ach)-1, NULL, NULL);

    if(_cchLen > 0)
    {
        // This is DBCS safe since byte before _cchLen is last character
        _ach[_cchLen] = 0;
        // BUGBUG DBCS REVIEW: this may not be safe if the last character
        // was a multibyte character...
        if(_ach[_cchLen-1]==0)
        {
            _cchLen--; // account for terminator
        }
        _pstr = _ach;
        return;
    }

    // Alloc space on heap for buffer.
    cchBufReq = WideCharToMultiByte(_uCP, 0, pwstr, cwch, NULL, 0, NULL, NULL);

    Assert(cchBufReq > 0);

    cchBufReq++; // may need to append NUL

    Trace((_T("CStrIn: Allocating buffer for argument (_uCP=%ld,cwch=%ld,pwstr=%lX,cchBufReq=%ld)\n"),
        _uCP, cwch, pwstr, cchBufReq));

    _pstr = new char[cchBufReq];
    if(!_pstr)
    {
        // On failure, the argument will point to the empty string.
        Trace0("CStrIn: No heap space for wrapped function argument.\n");
        _ach[0] = 0;
        _pstr = _ach;
        return;
    }

    Assert(HIWORD64(_pstr));
    _cchLen = WideCharToMultiByte(_uCP, 0, pwstr, cwch, _pstr, cchBufReq, NULL, NULL);

    // Again, make sure we're always null terminated
    Assert(_cchLen < cchBufReq);
    _pstr[_cchLen] = 0;
    if(0 == _pstr[_cchLen-1]) // account for terminator
    {
        _cchLen--;
    }
}



//+---------------------------------------------------------------------------
//
//  Member:     CStrOut::CStrOut
//
//  Synopsis:   Allocates enough space for an out buffer.
//
//  Arguments:  [pwstr]   -- The Unicode buffer to convert to when destroyed.
//                              May be NULL.
//
//              [cwchBuf] -- The size of the buffer in characters.
//
//  Modifies:   [this].
//
//----------------------------------------------------------------------------
CStrOut::CStrOut(LPWSTR pwstr, int cwchBuf) : CConvertStr(CP_ACP)
{
    Assert(cwchBuf >= 0);

    _pwstr = pwstr;
    _cwchBuf = cwchBuf;

    if(!pwstr)
    {
        _cwchBuf = 0;
        _pstr = NULL;
        return;
    }

    Assert(HIWORD64(pwstr));

    // Initialize buffer in case Windows API returns an error.
    _ach[0] = 0;

    // Use preallocated buffer if big enough.
    if(cwchBuf*2 <= ARRAYSIZE(_ach))
    {
        _pstr = _ach;
        return;
    }

    // Allocate buffer.
    Trace0("CStrOut: Allocating buffer for wrapped function argument.\n");
    _pstr = new char[cwchBuf*2];
    if(!_pstr)
    {
        // On failure, the argument will point to a zero-sized buffer initialized
        // to the empty string.  This should cause the Windows API to fail.
        Trace0("CStrOut: No heap space for wrapped function argument.\n");
        Assert(cwchBuf > 0);
        _pwstr[0] = 0;
        _cwchBuf = 0;
        _pstr = _ach;
        return;
    }

    Assert(HIWORD64(_pstr));
    _pstr[0] = 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrOut::ConvertIncludingNul
//
//  Synopsis:   Converts the buffer from MBCS to Unicode
//
//  Return:     Character count INCLUDING the trailing '\0'
//
//----------------------------------------------------------------------------
int CStrOut::ConvertIncludingNul()
{
    int cch;

    if(!_pstr)
    {
        return 0;
    }

    cch = MultiByteToWideChar(_uCP, 0, _pstr, -1, _pwstr, _cwchBuf);

    Free();
    return cch;
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrOut::ConvertExcludingNul
//
//  Synopsis:   Converts the buffer from MBCS to Unicode
//
//  Return:     Character count EXCLUDING the trailing '\0'
//
//----------------------------------------------------------------------------
int CStrOut::ConvertExcludingNul()
{
    int ret = ConvertIncludingNul();
    if(ret > 0)
    {
        ret -= 1;
    }
    return ret;
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrOut::~CStrOut
//
//  Synopsis:   Converts the buffer from MBCS to Unicode.
//
//  Note:       Don't inline this function, or you'll increase code size as
//              both ConvertIncludingNul() and CConvertStr::~CConvertStr will be
//              called inline.
//
//----------------------------------------------------------------------------
CStrOut::~CStrOut()
{
    ConvertIncludingNul();
}