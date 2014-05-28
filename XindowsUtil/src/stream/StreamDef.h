
#ifndef __XINDOWSUTIL_STREAM_STREAMDEF_H__
#define __XINDOWSUTIL_STREAM_STREAMDEF_H__

//+---------------------------------------------------------------------
//
//  IStorage and IStream Helper functions
//
//----------------------------------------------------------------------

// LARGE_INTEGER sign conversions are a pain without this
union LARGEINT
{
    LONGLONG        i64;
    ULONGLONG       ui64;
    LARGE_INTEGER   li;
    ULARGE_INTEGER  uli;
};

const LARGEINT LI_ZERO = { 0 };

XINDOWS_PUBLIC HRESULT CreateStorageOnHGlobal(HGLOBAL hgbl, LPSTORAGE* ppStg);
XINDOWS_PUBLIC HRESULT GetBStrFromStream(IStream* pIStream, BSTR* pbstr, BOOL fStripTrailingCRLF);

#endif //__XINDOWSUTIL_STREAM_STREAMDEF_H__