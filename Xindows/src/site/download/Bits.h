
#ifndef __XINDOWS_SITE_DOWNLOAD_BITS_H__
#define __XINDOWS_SITE_DOWNLOAD_BITS_H__

#include "Dwn.h"

// CBitsInfo ------------------------------------------------------------------
class CBitsInfo : public CDwnInfo
{
    typedef CDwnInfo super;
    friend class CBitsCtx;
    friend class CBitsLoad;

public:

    DECLARE_MEMCLEAR_NEW_DELETE()

    CBitsInfo(UINT dt) : CDwnInfo() { _dt = dt; }
    virtual        ~CBitsInfo();
    virtual UINT    GetType()                    { return (DWNCTX_BITS); }
    virtual DWORD   GetProgSinkClass()           { return (_dwClass); }
    virtual HRESULT Init(DWNLOADINFO* pdli);
    HRESULT         SetFile(LPCTSTR pch)         { RRETURN(_cstrFile.Set(pch)); }
    virtual HRESULT GetFile(LPTSTR* ppch);
    HRESULT         NewDwnCtx(CDwnCtx** ppDwnCtx);
    HRESULT         NewDwnLoad(CDwnLoad** ppDwnLoad);
    HRESULT         OnLoadFile(LPCTSTR pszFile, HANDLE* phLock, BOOL fIsTemp);
    void            OnLoadDwnStm(CDwnStm* pDwnStm);
    virtual void    OnLoadDone(HRESULT hr);
    virtual BOOL    AttachEarly(UINT dt, DWORD dwRefresh, DWORD dwFlags, DWORD dwBindf);
    HRESULT         GetStream(IStream** ppStream);

private:
    // Data members
    UINT        _dt;
    CString     _cstrFile;
    HANDLE      _hLock;
    CDwnStm*    _pDwnStm;
    BOOL        _fIsTemp;
    DWORD       _dwClass;
};

// CBitsLoad ---------------------------------------------------------------
class CBitsLoad : public CDwnLoad
{
    typedef CDwnLoad super;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

private:
    virtual ~CBitsLoad();
    CBitsInfo* GetBitsInfo()   { return ((CBitsInfo*)_pDwnInfo); }
    virtual HRESULT Init(DWNLOADINFO* pdli, CDwnInfo* pDwnInfo);        
    virtual HRESULT OnBindHeaders();
    virtual HRESULT OnBindData();

    // Data members
    BOOL        _fGotFile;
    BOOL        _fGotData;
    CDwnStm*    _pDwnStm;
    IStream*    _pStmFile;
};

#endif //__XINDOWS_SITE_DOWNLOAD_BITS_H__