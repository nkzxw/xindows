
#ifndef __XINDOWSUTIL_STRING_STRING_H__
#define __XINDOWSUTIL_STRING_STRING_H__

/*
Use this macro to avoid initialization of embedded
objects when parent object zeros out the memory
*/
#define CSTR_NOINIT     ((float)0.0)

/*
This class defines a length prefix 0 terminated string object. It points
to the beginning of the characters so the pointer returned can be used in
normal string operations taking into account that of course that it can
contain any binary value.
*/
class XINDOWS_PUBLIC CString
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    /*
    Default constructor
    */
    CString()
    {
        _pch = 0;
    }

    /*
    Special constructor to AVOID construction for embedded
    objects...
    */
    CString(float num)
    {
        Assert(_pch == 0);
    }
    /*
    Destructor will free data
    */
    ~CString()
    {
        _Free();
    }

    operator LPTSTR() const { return _pch; }
    HRESULT Set(LPCTSTR pch);
    HRESULT Set(LPCTSTR pch, UINT uc);

    HRESULT SetBSTR(const BSTR bstr);

    HRESULT Set(const CString& cstr);
    void    TakeOwnership(CString& cstr);
    UINT    Length() const;

    // Update the internal length indication without
    // changin any alocation.
    HRESULT SetLengthNoAlloc(UINT uc);

    // Reallocate the string to a larger size, length unchanged.
    HRESULT ReAlloc(UINT uc); 

    HRESULT Append(LPCTSTR pch);
    HRESULT Append(LPCTSTR pch, UINT uc);

    void Free()
    {
        _Free();
        _pch = 0;
    }

    TCHAR* TakePch() { TCHAR* pch = _pch; _pch = NULL; return(pch); }

    HRESULT AllocBSTR(BSTR* pBSTR) const;

    HRESULT TrimTrailingWhitespace();

private:
    void    _Free();
    LPTSTR  _pch;

    NO_COPY(CString);

public:
    HRESULT Clone(CString** ppCStr) const;
    BOOL    Compare(const CString* pCStr) const;
    WORD    ComputeCrc() const;
    BOOL    IsNull(void) const { return _pch==NULL?TRUE:FALSE; }
    HRESULT Save(IStream* pstm) const;
    HRESULT Load(IStream* pstm);
    ULONG   GetSaveSize() const;
};

//+-----------------------------------------------------------
//
//  Helper class:     CStringNullTerminator
//
//------------------------------------------------------------
class XINDOWS_PUBLIC CStringNullTerminator
{
public:
    CStringNullTerminator(LPTSTR pch)
    {
        _pch = pch;
        if(_pch)
        {
            _ch = *_pch;
            *_pch = 0;
        }
    }
    ~CStringNullTerminator()
    {
        if(_pch)
        {
            *_pch = _ch;
        }
    }

    LPTSTR  _pch;
    TCHAR   _ch;
};


#ifndef HIWORD64
#ifdef _WIN64
#define HIWORD64(p)     ((ULONG_PTR)(p)>>16)
#else
#define HIWORD64        HIWORD
#endif
#endif

//+---------------------------------------------------------------------------
//
//  Class:      CConvertStr (CStr)
//
//  Purpose:    Base class for conversion classes.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CConvertStr
{
public:
    operator char*();

protected:
    CConvertStr(UINT uCP);
    ~CConvertStr();
    void Free();

    UINT    _uCP;
    LPSTR   _pstr;
    char    _ach[MAX_PATH*sizeof(WCHAR)];
};

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStr::CConvertStr
//
//  Synopsis:   ctor.
//
//----------------------------------------------------------------------------
inline CConvertStr::CConvertStr(UINT uCP)
{
    _uCP = uCP;
    _pstr = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStr::~CConvertStr
//
//  Synopsis:   dtor.
//
//----------------------------------------------------------------------------
inline CConvertStr::~CConvertStr()
{
    Free();
}

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStr::operator char *
//
//  Synopsis:   Returns the string.
//
//----------------------------------------------------------------------------
inline CConvertStr::operator char*()
{
    return _pstr;
}



//+---------------------------------------------------------------------------
//
//  Class:      CStrIn (CStrI)
//
//  Purpose:    Converts string function arguments which are passed into
//              a Windows API.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CStrIn : public CConvertStr
{
public:
    CStrIn(LPCWSTR pwstr);
    CStrIn(LPCWSTR pwstr, int cwch);
    CStrIn(UINT uCP, LPCWSTR pwstr);
    CStrIn(UINT uCP, LPCWSTR pwstr, int cwch);
    int strlen();

protected:
    CStrIn();
    void Init(LPCWSTR pwstr, int cwch);

    int _cchLen;
};

//+---------------------------------------------------------------------------
//
//  Member:     CStrIn::CStrIn
//
//  Synopsis:   Inits the class with a given length
//
//----------------------------------------------------------------------------
inline CStrIn::CStrIn(LPCWSTR pwstr, int cwch) : CConvertStr(CP_ACP)
{
    Init(pwstr, cwch);
}

inline CStrIn::CStrIn(UINT uCP, LPCWSTR pwstr, int cwch) : CConvertStr(uCP==1200?CP_ACP:uCP)
{
    Init(pwstr, cwch);
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrIn::CStrIn
//
//  Synopsis:   Initialization for derived classes which call Init.
//
//----------------------------------------------------------------------------
inline CStrIn::CStrIn() : CConvertStr(CP_ACP)
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrIn::strlen
//
//  Synopsis:   Returns the length of the string in characters, excluding
//              the terminating NULL.
//
//----------------------------------------------------------------------------
inline int CStrIn::strlen()
{
    return _cchLen;
}



//+---------------------------------------------------------------------------
//
//  Class:      CStrOut (CStrO)
//
//  Purpose:    Converts string function arguments which are passed out
//              from a Windows API.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CStrOut : public CConvertStr
{
public:
    CStrOut(LPWSTR pwstr, int cwchBuf);
    ~CStrOut();

    int BufSize();
    int ConvertIncludingNul();
    int ConvertExcludingNul();

private:
    LPWSTR  _pwstr;
    int     _cwchBuf;
};

//+---------------------------------------------------------------------------
//
//  Member:     CStrOut::BufSize
//
//  Synopsis:   Returns the size of the buffer to receive an out argument,
//              including the terminating NULL.
//
//----------------------------------------------------------------------------
inline int CStrOut::BufSize()
{
    return _cwchBuf*sizeof(WCHAR);
}



//+---------------------------------------------------------------------------
//
//  Class:      CConvertStrW (CStr)
//
//  Purpose:    Base class for multibyte conversion classes.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CConvertStrW
{
public:
    operator WCHAR*();

protected:
    CConvertStrW();
    ~CConvertStrW();
    void Free();

    LPWSTR  _pwstr;
    WCHAR   _awch[MAX_PATH*sizeof(WCHAR)];
};

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStrW::CConvertStrW
//
//  Synopsis:   ctor.
//
//----------------------------------------------------------------------------
inline CConvertStrW::CConvertStrW()
{
    _pwstr = NULL;
}


//+---------------------------------------------------------------------------
//
//  Member:     CConvertStrW::~CConvertStrW
//
//  Synopsis:   dtor.
//
//----------------------------------------------------------------------------
inline CConvertStrW::~CConvertStrW()
{
    Free();
}

//+---------------------------------------------------------------------------
//
//  Member:     CConvertStrW::operator WCHAR *
//
//  Synopsis:   Returns the string.
//
//----------------------------------------------------------------------------
inline CConvertStrW::operator WCHAR*()
{
    return _pwstr;
}



//+---------------------------------------------------------------------------
//
//  Class:      CStrInW (CStrI)
//
//  Purpose:    Converts multibyte strings into UNICODE
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CStrInW : public CConvertStrW
{
public:
    CStrInW(LPCSTR pstr) { Init(pstr, -1); }
    CStrInW(LPCSTR pstr, int cch) { Init(pstr, cch); }
    CStrInW();

    int strlen();
    void Init(LPCSTR pstr, int cch);

protected:
    int _cwchLen;
};

//+---------------------------------------------------------------------------
//
//  Member:     CStrInW::CStrInW
//
//  Synopsis:   Initialization for derived classes which call Init.
//
//----------------------------------------------------------------------------
inline CStrInW::CStrInW()
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CStrInW::strlen
//
//  Synopsis:   Returns the length of the string in characters, excluding
//              the terminating NULL.
//
//----------------------------------------------------------------------------
inline int CStrInW::strlen()
{
    return _cwchLen;
}



//+---------------------------------------------------------------------------
//
//  Class:      CStrOutW (CStrO)
//
//  Purpose:    Converts returned unicode strings into ANSI.  Used for [out]
//				params (so we initialize with a buffer that should later be
//				filled with the correct ansi data)
//			
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CStrOutW : public CConvertStrW
{
public:
    CStrOutW(LPSTR pstr, int cchBuf);
    ~CStrOutW();

    int BufSize();
    int ConvertIncludingNul();
    int ConvertExcludingNul();

private:
    LPSTR   _pstr;
    int     _cchBuf;
};

//+---------------------------------------------------------------------------
//
//  Member:     CStrOutW::BufSize
//
//  Synopsis:   Returns the size of the buffer to receive an out argument,
//              including the terminating NULL.
//
//----------------------------------------------------------------------------
inline int CStrOutW::BufSize()
{
    return _cchBuf;
}

#endif //__XINDOWSUTIL_STRING_STRING_H__