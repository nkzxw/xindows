
#include "stdafx.h"
#include "testutil.h"

#define MAKE_HTKEY(i)   ((void*)((((DWORD)(i)*4567)<<2)|4))
#define MAKE_HTVAL(k)   ((void*)(~(DWORD)MAKE_HTKEY(k)))

CHtPvPv* phtable = NULL;

BOOL TestHTInsert(int i)
{
    void* pvKey = MAKE_HTKEY(i);
    void* pvVal = MAKE_HTVAL(i);
    Verify(phtable->Insert(pvKey, pvVal) == S_OK);
    Verify(phtable->Lookup(pvKey) == pvVal);
    return TRUE;
}

BOOL TestHTRemove(int i)
{
    void* pvKey = MAKE_HTKEY(i);
    void* pvVal = MAKE_HTVAL(i);
    Verify(phtable->Remove(pvKey) == pvVal);
    Verify(phtable->Remove(pvKey) == NULL);
    return TRUE;
}

BOOL TestHTVerify(int i, int n)
{
    void*   pvKey;
    void*   pvVal; 
    int     j;

    Verify((int)phtable->GetCount() == (n-i));

    for(j=i; j<n; ++j)
    {
        pvKey = MAKE_HTKEY(j);
        pvVal = MAKE_HTVAL(j);
        Verify(phtable->Lookup(pvKey) == pvVal);
    }

    return TRUE;
}

XINDOWS_PUBLIC HRESULT TestHtPvPv()
{
    CHtPvPv*    pht = new CHtPvPv;
    int         cLim, cEntMax;
    int         i/*, j*/;

    // Insufficient memory, don't crash.
    if(!pht)
    {
        return S_FALSE;
    }

    cLim    = 256;
    cEntMax = 383;
    phtable = pht;

    // Insert elements from 0 to cLim
    for(i=0; i<cLim; ++i)
    {
        if(!TestHTInsert(i)) return S_FALSE;
        if(!TestHTVerify(0, i+1)) return S_FALSE;
    }

    // Remove elements from 0 to cLim
    for(i=0; i<cLim; ++i)
    {
        if(!TestHTRemove(i)) return S_FALSE;
        if(!TestHTVerify(i+1, cLim)) return S_FALSE;
    }

    // Rehash and make sure number of deleted entries is now zero
    //phtable->Rehash(phtable->GetMaxCount());
    //Verify(phtable->GetDelCount() == 0);
    //if(!TestHTVerify(0, 0)) return S_FALSE;

    //// Insert elements from 0 to cLim/2
    //for(i=0; i<cLim/2; ++i)
    //{
    //    if(!TestHTInsert(i)) return S_FALSE;
    //    if(!TestHTVerify(0, i + 1)) return S_FALSE;
    //}

    //// Insert elements from cLim/2 to cLim and remove elements from 0 to cLim/2
    //for(i=cLim/2,j=0; i<cLim; ++i,++j)
    //{
    //    if(!TestHTInsert(i)) return S_FALSE;
    //    if(!TestHTVerify(j, i+1)) return S_FALSE;
    //    if(!TestHTRemove(j)) return S_FALSE;
    //    if(!TestHTVerify(j+1, i+1)) return S_FALSE;
    //}

    //// Insert two elements from cLim to cLim*2, remove one element from cLim/2 to cLim
    //for(i=cLim,j=cLim/2; i<cLim*2; i+=2,++j)
    //{
    //    if(!TestHTRemove(j)) return S_FALSE;
    //    if(!TestHTVerify(j+1, i)) return S_FALSE;
    //    if(!TestHTInsert(i)) return S_FALSE;
    //    if(!TestHTVerify(j+1, i+1)) return S_FALSE;
    //    if(!TestHTInsert(i+1)) return S_FALSE;
    //    if(!TestHTVerify(j+1, i+2)) return S_FALSE;
    //}

    //// Rehash and make sure number of deleted entries is now zero
    //phtable->Rehash(phtable->GetMaxCount());
    //Verify(phtable->GetDelCount() == 0);

    //if(!TestHTVerify(cLim, cLim*2)) return S_FALSE;

    //// Insert elements from cLim*2, remove two elements from cLim
    //for(i=cLim*2,j=cLim; i<cLim*3; ++i,j+=2)
    //{
    //    if(!TestHTInsert(i)) return S_FALSE;
    //    if(!TestHTVerify(j, i+1)) return S_FALSE;
    //    if(!TestHTRemove(j)) return S_FALSE;
    //    if(!TestHTVerify(j+1, i+1)) return S_FALSE;
    //    if(!TestHTRemove(j+1)) return S_FALSE;
    //    if(!TestHTVerify(j+2, i+1)) return S_FALSE;
    //}

    //// Rehash and make sure number of deleted entries is now zero
    //phtable->Rehash(phtable->GetMaxCount());
    //Verify(phtable->GetDelCount() == 0);
    //if(!TestHTVerify(0, 0)) return S_FALSE;

    delete pht;

    return S_OK;
}





static int Test1Format(HRESULT hrExpected, TCHAR* pszExpected, TCHAR* pszFmt, ...)
{
    HRESULT hr;
    int     cErrors = 0;
    va_list arg;
    TCHAR   szBuf[256];
    TCHAR*  pszOutput;

    va_start(arg, pszFmt);
    hr = VFormat(_afxGlobalData._hInstResource, 0, szBuf, ARRAYSIZE(szBuf), pszFmt, &arg);
    va_end(arg);
    if(hr != hrExpected)
    {
        cErrors += 1;
        Trace3("Error for buf '%ls'\n  Expected=%08x\n Actual  =%08x\n", pszFmt, hrExpected, hr);
    }
    else if(!hrExpected && _tcscmp(szBuf, pszExpected))
    {
        cErrors += 1;
        Trace3("Error for buf '%ls'\n  Expected '%ls'\n  Actual   '%ls'\n", pszFmt, pszExpected, szBuf);
    }

    va_start(arg, pszFmt);
    hr = VFormat(_afxGlobalData._hInstResource, FMT_OUT_ALLOC, &pszOutput, 10, pszFmt,&arg);
    va_end(arg);
    if(hr != hrExpected)
    {
        cErrors += 1;
        Trace3("Error for alloc 'l%s'\n  Expected=%08x\n Actual  =%08x\n", pszFmt, hrExpected, hr);
    }
    else if(!hrExpected && _tcscmp(pszOutput, pszExpected))
    {
        cErrors += 1;
        Trace3("Error for alloc '%ls'\n  Expected '%ls'\n  Actual   '%ls'\n", pszFmt, pszExpected, pszOutput);
    }
    MemFree(pszOutput);
    return cErrors;
}

int TestFormat()
{
    int cErrors = 0;

    cErrors += Test1Format(0, TEXT("Now is the time."), TEXT("Now is the time."));

    cErrors += Test1Format(0, TEXT("1 2 3 4 5"),
        TEXT("<0d> <1d> <2d> <3d> <4d>"),
        (long)1, (long)2, (long)3, (long)4, (long)5);

    cErrors += Test1Format(0, TEXT("1 2 3 4 5"),
        TEXT("<0s> <1s> <2s> <3s> <4s>"),
        TEXT("1"), TEXT("2"), TEXT("3"), TEXT("4"), TEXT("5"));

    cErrors += Test1Format(0, TEXT("<0d>"),
        TEXT("<<<0d>d>"),
        (long)0);

    cErrors += Test1Format(0, TEXT("aba"),
        TEXT("<0s><1s><0s>"),
        TEXT("a"), TEXT("b"));

    cErrors += Test1Format(0, TEXT("-1 -2 -3 -4 -5"),
        TEXT("<0d> <1d> <2d> <3d> <4d>"),
        (long)-1, (long)-2, (long)-3, (long)-4, (long)-5);

    cErrors += Test1Format(0, TEXT("There is 1 toaster flying across my screen"),
        TEXT("There <0p/is/are/> <0d> toaster<0p//s/> flying across my screen"),
        (long)1);

    cErrors += Test1Format(0, TEXT("There are 2 toasters flying across my screen"),
        TEXT("There <0p/is/are/> <0d> toaster<0p//s/> flying across my screen"),
        (long)2);

    cErrors += Test1Format(E_FAIL, TEXT("-1 -2 -3 -4 -5"),
        TEXT("<0q> <1d> <2d> <3d> <4d>"),
        (long)-1, (long)-2, (long)-3, (long)-4, (long)-5);

    cErrors += Test1Format(E_FAIL, TEXT("-1 -2 -3 -4 -5"),
        TEXT("<0dfoobar> <1d> <2d> <3d> <4d>"),
        (long)-1, (long)-2, (long)-3, (long)-4, (long)-5);

    return cErrors;
}

XINDOWS_PUBLIC void TestStringUtil()
{
    int nErrors = TestFormat();
}

XINDOWS_PUBLIC void TestButtonUtil(HDC hdc)
{
    CEnsureThreadState ets;

    CDrawInfo di;
    di._hdc = hdc;
    di._dxrInch = di._dyrInch = 96;
    di._dxtInch = di._dytInch = 96;
    CUtilityButton util;
    RECT rc = { 100, 100, 120, 120 };
    SIZEL sz = { 20, 20 };
    util.DrawButton(&di, NULL, BG_LEFT, TRUE, TRUE, TRUE, rc, sz, 0);
}

XINDOWS_PUBLIC void TestVariant()
{
    CVariant vt(VT_NULL);
    vt.CoerceVariantArg(VT_BSTR);
}