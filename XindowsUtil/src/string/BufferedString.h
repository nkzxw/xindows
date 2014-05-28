
#ifndef __XINDOWSUTIL_STRING_BUFFERSTRING_H__
#define __XINDOWSUTIL_STRING_BUFFERSTRING_H__

//+---------------------------------------------------------------------------
//
// Contents     Class definition for a buffered, appendable string class
//
// Classes      CBufferedString
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CBufferedString
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    CBufferedString::CBufferedString() 
    {
        _pchBuf = NULL;
        _cchBufSize = 0;
        _cchIndex =0;
    }
    CBufferedString::~CBufferedString() { Free(); }

    /*
    Creates a buffer and initializes it with the supplied TCHAR*
    */
    HRESULT Set(const TCHAR* pch=NULL);
    void Free() { delete[] _pchBuf; }

    /*
    Adds at the end of the buffer, growing it it necessary.
    */
    HRESULT QuickAppend(const TCHAR* pch) { return (QuickAppend(pch, _tcslen(pch))); }
    HRESULT QuickAppend(const TCHAR* pch, ULONG cch);

    /*
    Returns current size of the buffer
    */
    UINT    Size() { return _pchBuf?_cchBufSize:0; }

    /*
    Returns current length of the buffer string
    */
    UINT    Length() { return _pchBuf?_cchIndex:0; }

    operator LPTSTR() const { return _pchBuf; }

    TCHAR*  _pchBuf;        // Actual buffer
    UINT    _cchBufSize;    // Size of _pchBuf
    UINT    _cchIndex;      // Length of _pchBuf
};

#endif //__XINDOWSUTIL_STRING_BUFFERSTRING_H__