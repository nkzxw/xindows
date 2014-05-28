
#ifndef __XINDOWSUTIL_STRING_STRINGUTIL_H__
#define __XINDOWSUTIL_STRING_STRINGUTIL_H__

//+-------------------------------------------------------------
//
//  Formatting Swiss Army Knife
//
//--------------------------------------------------------------
enum FMT_OPTIONS
{
    FMT_EXTRA_NULL_MASK = 0x0F,
    FMT_OUT_ALLOC       = 0x10,
    FMT_OUT_XTRATERM    = 0x20,
    FMT_ARG_ARRAY       = 0x40
};
XINDOWS_PUBLIC HRESULT VFormat(HINSTANCE hInstResource, DWORD dwOptions, void* pvOutput, int cchOutput, TCHAR* pchFmt, void* pvArgs);
XINDOWS_PUBLIC HRESULT Format(HINSTANCE hInstResource, DWORD dwOptions, void* pvOutput, int cchOutput, TCHAR* pchFmt, ...);

//+-------------------------------------------------------------
//
//  BSTR utilities
//
//--------------------------------------------------------------
#ifdef _UNICODE
XINDOWS_PUBLIC HRESULT FormsAllocStringW(LPCWSTR pch, BSTR* pBSTR);
XINDOWS_PUBLIC HRESULT FormsAllocStringLenW(LPCWSTR pch, UINT uc, BSTR* pBSTR);
XINDOWS_PUBLIC HRESULT FormsReAllocStringW(BSTR* pBSTR, LPCWSTR pch);
XINDOWS_PUBLIC HRESULT FormsReAllocStringLenW(BSTR* pBSTR, LPCWSTR pch, UINT uc);
#else
XINDOWS_PUBLIC HRESULT FormsAllocStringW(LPCTSTR pch, BSTR* pBSTR);
XINDOWS_PUBLIC HRESULT FormsAllocStringLenW(LPCTSTR pch, UINT uc, BSTR* pBSTR);
XINDOWS_PUBLIC HRESULT FormsReAllocStringW(BSTR* pBSTR, LPCWSTR pch);
XINDOWS_PUBLIC HRESULT FormsReAllocStringLenW(BSTR* pBSTR, LPCSTR pch, UINT uc);
#endif // UNICODE

XINDOWS_PUBLIC UINT    FormsStringLen(const BSTR bstr);
XINDOWS_PUBLIC UINT    FormsStringByteLen(const BSTR bstr);

#define FormsAllocString        FormsAllocStringW
#define FormsAllocStringLen     FormsAllocStringLenW
#define FormsReAllocString      FormsReAllocStringW
#define FormsReAllocStringLen   FormsReAllocStringLenW

#endif //__XINDOWSUTIL_STRING_STRINGUTIL_H__