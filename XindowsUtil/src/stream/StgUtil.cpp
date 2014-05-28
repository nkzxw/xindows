
#include "stdafx.h"
#include "StreamDef.h"

//+---------------------------------------------------------------
//
//  Function:   CreateStorageOnHGlobal
//
//  Synopsis:   Creates an IStorage on a global memory handle
//
//  Arguments:  [hgbl] -- memory handle to create storage on
//              [ppStg] -- where the storage is returned
//
//  Returns:    Success if the storage could be successfully created
//              on the memory handle.
//
//  Notes:      This helper function combines CreateILockBytesOnHGlobal
//              and StgCreateDocfileOnILockBytes.  hgbl may be NULL in
//              which case a global memory handle will be automatically
//              allocated.
//
//----------------------------------------------------------------
HRESULT CreateStorageOnHGlobal(HGLOBAL hgbl, LPSTORAGE* ppStg)
{
    HRESULT hr;
    LPLOCKBYTES pLockBytes;

    hr = CreateILockBytesOnHGlobal(hgbl, TRUE, &pLockBytes);

    if(!hr)
    {
        hr = StgCreateDocfileOnILockBytes(
            pLockBytes,
            STGM_CREATE|STGM_READWRITE|STGM_SHARE_EXCLUSIVE,
            0,
            ppStg);
        pLockBytes->Release();
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   GetBStrFromStream
//
//  Synopsis:   Given a stream, this function allocates and returns a bstr
//              representing its contents.
//
//----------------------------------------------------------------------------
HRESULT GetBStrFromStream(IStream* pIStream, BSTR* pbstr, BOOL fStripTrailingCRLF)
{
    HRESULT hr;
    HGLOBAL hHtmlText = 0;
    TCHAR*  pstrWide = NULL;

    *pbstr = NULL;

    hr = GetHGlobalFromStream(pIStream, &hHtmlText);
    if(hr)
    {
        goto Cleanup;
    }

    pstrWide = (TCHAR*)GlobalLock(hHtmlText);

    if(!pstrWide)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    if(fStripTrailingCRLF)
    {
        // Remove trailing cr/lf's
        TCHAR* pstr = pstrWide + _tcslen(pstrWide);

        while((pstr-- > pstrWide) && (*pstr=='\r' || *pstr=='\n')) ;

        *(pstr+1) = 0;
    }

    hr = FormsAllocString(pstrWide, pbstr);

    GlobalUnlock(hHtmlText);

Cleanup:
    RRETURN(hr);
}