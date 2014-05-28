
#include "stdafx.h"
#include "FatStg.h"

class FatStream : public IStream
{
    friend HRESULT CreateStreamOnFile(LPCTSTR, DWORD, LPSTREAM* ppstrm);
    friend HRESULT CloseStreamOnFile(LPSTREAM pStm);

public:
    DECLARE_MEMALLOC_NEW_DELETE();
    DECLARE_STANDARD_IUNKNOWN(FatStream);

    // *** IStream methods ***
    STDMETHOD(Read)(void* pv, ULONG cb, ULONG* pcbRead);
    STDMETHOD(Write)(void const* pv, ULONG cb, ULONG* pcbWritten);
    STDMETHOD(Seek)(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition);
    STDMETHOD(SetSize)(ULARGE_INTEGER libNewSize);
    STDMETHOD(CopyTo)(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten);
    STDMETHOD(Commit)(DWORD grfCommitFlags);
    STDMETHOD(Revert)(void);
    STDMETHOD(LockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(UnlockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(Stat)(STATSTG* pstatstg, DWORD grfStatFlag);
    STDMETHOD(Clone)(IStream** ppstm);

private:
    FatStream()
    {
        _ulRefs = 1;
        _hfile = INVALID_HANDLE_VALUE;
    }
    ~FatStream();

    HRESULT Init(HANDLE hfile, LPCTSTR pchFile);

    HANDLE  _hfile;
    CString _cstrFileName;
};

FatStream::~FatStream()
{
    if(_hfile != INVALID_HANDLE_VALUE)
    {
        CloseHandle(_hfile); 
    }

    if(_cstrFileName.Length()) 
    {
        DeleteFile(_cstrFileName);
    }
}

//+---------------------------------------------------------------
//
//  Function:   CreateStreamOnFile, public
//
//  Synopsis:   Provides an IStream interface to a DOS file
//
//  Arguments:  [pchFile] -- the DOS file
//              [dwDesiredAccess] -- see CreateFile
//              [dwShareMode] -- see CreateFile
//              [dwCreationDistribution] -- see CreateFile
//              [ppstrm] -- where the opened stream is returned
//
//  Returns:    Success if the stream interface could be provided
//
//----------------------------------------------------------------
HRESULT CreateStreamOnFile(LPCTSTR pchFile, DWORD dwSTGM, LPSTREAM* ppstrm)
{
    DWORD       dwDesiredAccess = 0;
    DWORD       dwShareMode = 0;
    DWORD       dwCreationDistribution = 0;
    HANDLE      hfile;
    FatStream*  pStmFat = NULL;
    HRESULT     hr = S_OK;

    Assert(pchFile!=NULL && _tcsclen(pchFile)!=0);

    if(dwSTGM & STGM_READWRITE)
    {
        dwDesiredAccess |= (GENERIC_READ|GENERIC_WRITE);
    }
    else if(dwSTGM & STGM_WRITE)
    {
        dwDesiredAccess |= GENERIC_WRITE;
    }
    else
    {
        dwDesiredAccess |= GENERIC_READ;
    }

    if(dwSTGM & STGM_SHARE_DENY_NONE)
    {
        dwShareMode |= (FILE_SHARE_READ|FILE_SHARE_WRITE);
    }
    else if(dwSTGM & STGM_SHARE_DENY_WRITE)
    {
        dwShareMode |= FILE_SHARE_READ;
    }
    else if(dwSTGM & STGM_SHARE_DENY_READ)
    {
        dwShareMode |= FILE_SHARE_WRITE;
    }

    if(dwSTGM & STGM_CREATE)
    {
        dwCreationDistribution |= CREATE_ALWAYS;
    }
    else
    {
        dwCreationDistribution |= OPEN_EXISTING;
    }

    hfile = CreateFile(
        pchFile,
        dwDesiredAccess,
        dwShareMode,
        NULL,
        dwCreationDistribution,
        FILE_ATTRIBUTE_NORMAL,
        NULL);

    if(hfile == INVALID_HANDLE_VALUE)
    {
        hr = GetLastWin32Error();
        goto Cleanup;
    }

    pStmFat = new FatStream();
    if(!pStmFat)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = pStmFat->Init(hfile, dwSTGM&STGM_DELETEONRELEASE?pchFile:NULL);
    if(hr)
    {
        goto Error;
    }

    *ppstrm = pStmFat;

Cleanup:
    RRETURN(hr);

Error:
    delete pStmFat;
    goto Cleanup;
}

HRESULT CloseStreamOnFile(LPSTREAM pStm)
{
    HRESULT hr = S_OK;
    FatStream* pStmFat = DYNCAST(FatStream, pStm);

    if(pStmFat->_hfile == INVALID_HANDLE_VALUE)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    if(!CloseHandle(pStmFat->_hfile))
    {
        hr = GetLastWin32Error();
    }

    pStmFat->_hfile = INVALID_HANDLE_VALUE;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Init
//
//  Synopsis:   Initialize the class.
//
//----------------------------------------------------------------
HRESULT FatStream::Init(HANDLE hfile, LPCTSTR pchFile)
{
    HRESULT hr = S_OK;

    _hfile = hfile;

    if(pchFile)
    {
        hr = _cstrFileName.Set(pchFile);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::QueryInterface
//
//  Synopsis:   method from IUnknown interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::QueryInterface(REFIID iid, void** ppv)
{
    if(iid==IID_IUnknown || iid==IID_IStream)
    {
        *ppv = (IStream*)this;
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    ((IUnknown*)*ppv)->AddRef();
    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Read
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
    ULONG cbTemp;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        RRETURN(E_FAIL);
    }

    RRETURN(ReadFile(
        _hfile,
        pv,
        cb,
        pcbRead?pcbRead:&cbTemp,
        NULL)?S_OK:GetLastWin32Error());
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Write
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Write(void const* pv, ULONG cb, ULONG* pcbWritten)
{
    BOOL fSuccess;
    DWORD cbWritten;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        RRETURN(E_FAIL);
    }

    fSuccess = WriteFile(_hfile, pv, (DWORD)cb, &cbWritten, NULL);
    if(pcbWritten != NULL)
    {
        *pcbWritten = cbWritten;
    }

    RRETURN(fSuccess?S_OK:GetLastWin32Error());
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Seek
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
{
    DWORD newpos;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        RRETURN(E_FAIL);
    }

    if(dlibMove.HighPart!=0 && dlibMove.HighPart!=-1)
    {
        RRETURN(E_FAIL);
    }

    newpos = SetFilePointer(_hfile, (LONG)dlibMove.LowPart, NULL, (int)dwOrigin);

    if(plibNewPosition != NULL)
    {
        ULISet32(*plibNewPosition, (LONG)newpos);
    }

    RRETURN(newpos==-1 ? E_FAIL : S_OK);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::SetSize
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::SetSize(ULARGE_INTEGER libNewSize)
{
    DWORD newpos;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        RRETURN(E_FAIL);
    }

    if(libNewSize.HighPart != 0)
    {
        RRETURN(E_FAIL);
    }

    newpos = SetFilePointer(_hfile, (LONG)libNewSize.LowPart, NULL, FILE_BEGIN);
    if(newpos == -1)
    {
        RRETURN(GetLastWin32Error());
    }

    if(!SetEndOfFile(_hfile))
    {
        RRETURN(GetLastWin32Error());
    }

    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::CopyTo
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten)
{
    DWORDLONG   cbRequest   = *(LONGLONG*)&cb;
    DWORDLONG   cbActual    = 0;
    HRESULT     hr          = S_OK;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    while(cbRequest)
    {
        BYTE    buf[4096];
        ULONG   cbAttempt = min(cbRequest, (DWORDLONG)sizeof(buf));
        ULONG   cbRead;

        hr = Read(buf, cbAttempt, &cbRead);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pstm->Write(buf, cbRead, NULL);
        if(hr)
        {
            goto Cleanup;
        }

        cbActual += cbRead;
        cbRequest -= cbRead;

        if(cbAttempt > cbRead)
        {
            break;
        }
    }

    if(pcbRead)
    {
        *pcbRead = *(ULARGE_INTEGER*)&cbActual;
    }
    if(pcbWritten)
    {
        *pcbWritten = *(ULARGE_INTEGER*)&cbActual;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Commit
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Commit(DWORD grfCommitFlags)
{
    RRETURN(STG_E_INVALIDFUNCTION);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Revert
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Revert(void)
{
    RRETURN(STG_E_INVALIDFUNCTION);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::LockRegion
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
    RRETURN(STG_E_INVALIDFUNCTION);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::UnlockRegion
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
    RRETURN(STG_E_INVALIDFUNCTION);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Stat
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Stat(STATSTG* pstatstg, DWORD grfStatFlag)
{
    HRESULT hr = S_OK;

    if(_hfile == INVALID_HANDLE_VALUE)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    if(pstatstg != NULL)
    {
        pstatstg->pwcsName = NULL;
        if(grfStatFlag != STATFLAG_NONAME)
        {
            int cchFileName = _cstrFileName.Length();
            if(cchFileName)
            {
                pstatstg->pwcsName = (LPOLESTR)CoTaskMemAlloc(sizeof(TCHAR)*(cchFileName+1));
                if(!pstatstg->pwcsName)
                {
                    hr = E_OUTOFMEMORY;
                    goto Cleanup;
                }
                _tcscpy(pstatstg->pwcsName, _cstrFileName);
            }
        }

        pstatstg->type = STGTY_STREAM;
        ULISet32(pstatstg->cbSize, (LONG)GetFileSize(_hfile, NULL));
        pstatstg->grfLocksSupported = 0; // no locking supported
    }

Cleanup:
    RRETURN (hr);
}

//+---------------------------------------------------------------
//
//  Member:     FatStream::Clone
//
//  Synopsis:   method of IStream interface
//
//----------------------------------------------------------------
STDMETHODIMP FatStream::Clone(IStream** ppstm)
{
    RRETURN(E_NOTIMPL);
}