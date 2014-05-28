
#include "stdafx.h"
#include "Image.h"

#include "ArtPlayer.h"

CImgTask::~CImgTask()
{
    if(_pImgInfo)
    {
        _pImgInfo->SubRelease();
    }

    if(_pDwnBindData)
    {
        _pDwnBindData->Release();
    }

    if(!_fComplete)
    {
        FreeGifAnimData(&_gad, (CImgBitsDIB*)_pImgBits);
        if(_pArtPlayer)
        {
            delete _pArtPlayer;
        }
        if(_pImgBits)
        {
            delete _pImgBits;
        }
    }
}

void CImgTask::Init(CImgInfo* pImgInfo, MIMEINFO* pmi, CDwnBindData* pDwnBindData)
{
    _colorMode  = pImgInfo->GetColorMode();
    _pmi        = pmi;
    _lTrans     = -1;
    _ySrcBot    = -2;

    _pImgInfo = pImgInfo;
    _pImgInfo->SubAddRef();

    _pDwnBindData = pDwnBindData;
    _pDwnBindData->AddRef();
}

void CImgTask::Run()
{
    GetImgTaskExec()->RunTask(this);
}

BOOL CImgTask::Read(void* pv, ULONG cb, ULONG* pcbRead, ULONG cbMinReq)
{
    ULONG   cbReq = cb, cbGot, cbTot = 0;
    HRESULT hr    = S_OK;

    if(cbMinReq==0 || cbMinReq>cb)
    {
        cbMinReq = cb;
    }

    for(;;)
    {
        if(_fTerminate)
        {
            hr = E_ABORT;
            break;
        }

        hr = _pDwnBindData->Read(pv, cbReq, &cbGot);
        if(hr)
        {
            break;
        }

        cbTot += cbGot;
        cbReq -= cbGot;
        pv = (BYTE*)pv + cbGot;

        if(cbReq == 0)
        {
            break;
        }

        if(!cbGot || IsTimeout())
        {
            if(_pDwnBindData->IsEof() || (!cbGot && cbTot>=cbMinReq))
            {
                break;
            }

            GetImgTaskExec()->YieldTask(this, !cbGot);
        }
    }

    if(pcbRead)
    {
        *pcbRead = cbTot;
    }

    return (hr==S_OK && cbTot>0);
}

void CImgTask::Terminate()
{
    if(!_fTerminate)
    {
        _fTerminate = TRUE;
        SetBlocked(FALSE);
    }
}

void CImgTask::OnProg(BOOL fLast, ULONG ulBits, BOOL fAll, LONG yBot)
{
    DWORD dwTick = GetTickCount();

    _yTopProg = Union(_yTopProg, _yBotProg, fAll, yBot);
    _yBotProg = yBot;

    if(fLast || (dwTick-_dwTickProg>1000))
    {
        _dwTickProg = dwTick;

        _pImgInfo->OnTaskProg(this, ulBits, _yTopProg==-1, _yBotProg);

        _yTopProg = _yBotProg;
    }
}

void CImgTask::OnAnim()
{
    _pImgInfo->OnTaskAnim(this);
}

void CImgTask::Exec()
{
    BOOL fNonProgressive = FALSE;

    _dwTickProg = GetTickCount();

    Decode(&fNonProgressive);
    if(_pImgInfo->TstFlags(DWNF_MIRRORIMAGE))
    {
        if(_pImgBits)
        {
            ((CImgBitsDIB*)(_pImgBits))->SetMirrorStatus(TRUE);
        }
    }

    if(_pImgBits && _ySrcBot>-2)
    {
        _fComplete = _pImgInfo->OnTaskBits(this, _pImgBits,
            &_gad, _pArtPlayer, _lTrans, _ySrcBot, fNonProgressive);
    }

    if(_ySrcBot==-2 || (_fComplete && _ySrcBot!=-1))
    {
        OnProg(TRUE, _ySrcBot==-2?IMGBITS_NONE:IMGBITS_PARTIAL, TRUE, _yBot);
    }

    if(_fComplete && _ySrcBot==-1 && !_pDwnBindData->IsEof())
    {
        BYTE ab[512];
        // The image is fully decoded but the binding hasn't reached EOF.
        // Attempt to read the final EOF to allow the binding to complete
        // normally.  This will make sure that HTTP downloads don't delete
        // the cache file just because a decoder (such as the BMP decoder)
        // knows how many bytes to expect based on the header and doesn't
        // stick around until it sees EOF.
        Read(ab, sizeof(ab), NULL);
    }

    if(!_pDwnBindData->IsEof())
    {
        // Looks like the decoder didn't like the data.  Since it is still
        // flowing and we don't need any more of it, abort the binding.
        _pDwnBindData->Terminate(E_ABORT);
    }

    _fTerminate = TRUE;
}
 
BOOL CImgTask::DoTaskGetReport(CArtPlayer* pArtPlayer)
{
    BOOL fResult = FALSE;

    if((pArtPlayer==_pArtPlayer) && _pImgBits)
    {
        fResult = _pArtPlayer->GetArtReport((CImgBitsDIB**)_pImgBits,
            _yHei, _colorMode);
    }

    return fResult;
}