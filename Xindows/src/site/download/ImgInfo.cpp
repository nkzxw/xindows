
#include "stdafx.h"
#include "Image.h"

#include "ArtPlayer.h"

CImgBits* g_pImgBitsNotLoaded = NULL;
CImgBits* g_pImgBitsMissing = NULL;

// CImgInfo -------------------------------------------------------------------
HRESULT CImgInfo::Init(DWNLOADINFO* pdli)
{
    HRESULT hr;

    hr = super::Init(pdli);
    if(hr)
    {
        goto Cleanup;
    }

    _fNoOptimize = pdli->pDwnDoc->GetDownf() & DWNF_NOOPTIMIZE;
    SetFlags(IMGBITS_NONE);

Cleanup:
    RRETURN(hr);
}

CImgInfo::~CImgInfo()
{
    Assert(_pImgTask == NULL);
    Assert(_cLoad == 0);

    if(GetDwnInfoLock() == NULL)
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

void CImgInfo::Passivate()
{
    CImgTask* pImgTask;

    super::Passivate();

    pImgTask = (CImgTask*)InterlockedExchangePointer((void**)&_pImgTask, NULL);

    if(pImgTask)
    {
        pImgTask->Terminate();
        pImgTask->Release();
    }
}

HRESULT CImgInfo::NewDwnCtx(CDwnCtx** ppDwnCtx)
{
    *ppDwnCtx = new CImgCtx;
    RRETURN(*ppDwnCtx ? S_OK : E_OUTOFMEMORY);
}

HRESULT CImgInfo::NewDwnLoad(CDwnLoad ** ppDwnLoad)
{
    *ppDwnLoad = new CImgLoad;
    RRETURN(*ppDwnLoad ? S_OK : E_OUTOFMEMORY);
}

CArtPlayer* CImgInfo::GetArtPlayer()
{
    CArtPlayer* pArtPlayer = NULL;

    EnterCriticalSection();

    if(_pArtPlayer)
    {
        pArtPlayer = _pArtPlayer;
    }
    else if(_pImgTask)
    {
        pArtPlayer = _pImgTask->GetArtPlayer();
    }

    LeaveCriticalSection();

    return pArtPlayer;
}

void CImgInfo::Signal(WORD wChg, BOOL fInvalAll, int yBot)
{
    if(_pDwnCtxHead)
    {
        EnterCriticalSection();

        for(CDwnCtx* pDwnCtx=_pDwnCtxHead; pDwnCtx; pDwnCtx=pDwnCtx->GetDwnCtxNext())
        {
            ((CImgCtx*)pDwnCtx)->Signal(wChg, fInvalAll, yBot);
        }

        LeaveCriticalSection();
    }
}

void CImgInfo::Reset()
{
    Assert(EnteredCriticalSection());

    _ySrcBot = 0;

    Signal(IMGCHG_VIEW, TRUE, 0);

    if(_gad.pgf)
    {
        FreeGifAnimData(&_gad, (CImgBitsDIB*)_pImgBits);
        memset(&_gad, 0, sizeof(GIFANIMDATA));

        //$BUGBUG (lmollico): we should free the pImgAnimState list
        ClrFlags(IMGANIM_ANIMATED);
    }

    if(_pArtPlayer)
    {
        delete _pArtPlayer;
        _pArtPlayer = NULL;

        // ??? $BUGBUG (lmollico): we should free the pImgAnimState list
        ClrFlags(IMGANIM_ANIMATED);
    }

    if(_pImgBits)
    {
        delete _pImgBits;
        _pImgBits = NULL;
    }
}

void CImgInfo::Abort(HRESULT hrErr, CDwnLoad** ppDwnLoad)
{
    Assert(EnteredCriticalSection());

    if(TstFlags(DWNLOAD_LOADING))
    {
        Signal(IMGCHG_VIEW, TRUE, 0);
    }

    if(_pImgTask)
    {
        _pImgTask->Terminate();
    }

    super::Abort(hrErr, ppDwnLoad);
}

void CImgInfo::DrawFrame(
         HDC hdc,
         IMGANIMSTATE* pImgAnimState,
         RECT* prcDst,
         RECT* prcSrc,
         RECT* prcDstFull,
         DWORD dwFlags)
{
    GIFFRAME* pgf;
    GIFFRAME* pgfDraw;
    GIFFRAME* pgfDrawNext;
    POINT pt;
    HRGN hrgnClipOld = NULL;
    int iOldClipKind = -1;
    RECT rcSrc;
    LONG xDstWid;
    LONG yDstHei;

    pt = _afxGlobalData._Zero.pt;

    Assert(!prcSrc || prcDstFull);

    if(prcSrc == NULL)
    {
        rcSrc.left   = 0;
        rcSrc.top    = 0;
        rcSrc.right  = _xWid;
        rcSrc.bottom = _yHei;
        prcSrc       = &rcSrc;

        if(!prcDstFull)
        {
            prcDstFull = prcDst;
        }
    }

    Assert(pImgAnimState != NULL);

    EnterCriticalSection();

    if(GetArtPlayer())
    {
        DrawImage(hdc, prcDst, prcSrc, SRCCOPY, dwFlags);
    }
    else
    {
        xDstWid = prcDstFull->right - prcDstFull->left;
        yDstHei = prcDstFull->bottom - prcDstFull->top;

        ComputeFrameVisibility(pImgAnimState, _xWid, _yHei, xDstWid, yDstHei);

        pgfDraw = pImgAnimState->pgfDraw;
        pgfDrawNext = pgfDraw->pgfNext;

        {
            hrgnClipOld = CreateRectRgnIndirect(prcDst);
            if(hrgnClipOld == NULL)
            {
                goto Cleanup;
            }

            iOldClipKind = GetClipRgn(hdc, hrgnClipOld);
            if(iOldClipKind == -1)
            {
                goto Cleanup;
            }

            GetViewportOrgEx(hdc, &pt);
            pt.x += prcDstFull->left;
            pt.y += prcDstFull->top;
        }

        // Now, draw the frames from this iteration
        for(pgf=pImgAnimState->pgfFirst; pgf!=pgfDrawNext; pgf=pgf->pgfNext)
        {
            if(pgf->bRgnKind != NULLREGION)
            {
                {
                    OffsetRgn(pgf->hrgnVis, pt.x, pt.y);
                    // BUGWIN16: We don't have ExtSelectClipRgn - sure to cause some bugs here !!
                    if(iOldClipKind == 1)
                    {
                        ExtSelectClipRgn(hdc, pgf->hrgnVis, RGN_AND);
                    }
                    else
                    {
                        SelectClipRgn(hdc, pgf->hrgnVis);
                    }
                }

                pgf->pibd->StretchBltOffset(hdc, prcDst, prcSrc, pgf->left, pgf->top, SRCCOPY, dwFlags);

                {
                    if(iOldClipKind == 1)
                    {
                        SelectClipRgn(hdc, hrgnClipOld);
                    }
                    else
                    {
                        SelectClipRgn(hdc, NULL);
                    }
                }
            }
            if(pgf->hrgnVis != NULL)
            {
                DeleteRgn(pgf->hrgnVis);
                pgf->hrgnVis = NULL;
                pgf->bRgnKind = NULLREGION;
            }
        } // for each frame in the preceding iteration
    }

Cleanup:
    if(hrgnClipOld)
    {
        DeleteRgn(hrgnClipOld);
    }

    LeaveCriticalSection();
}

//+------------------------------------------------------------------------
//
//  Member:     CImgInfo::NextFrame
//
//  Synopsis:   returns TRUE if there is a Frame to Draw
//
//-------------------------------------------------------------------------
BOOL CImgInfo::NextFrame(IMGANIMSTATE* pImgAnimState, DWORD dwCurTimeMS, DWORD* pdwFrameTimeMS)
{
    int iLateTimeMS;
    BOOL fResult = TRUE;

    *pdwFrameTimeMS = 0xFFFFFFFF;
    BOOL fCritical = FALSE;

    // CHROME
    // Check for NULL before accessing members. Check was previously after
    // member access
    if(pImgAnimState == NULL)
    {
        return FALSE;
    }

    pImgAnimState->fLoop = FALSE;
    pImgAnimState->fStop = FALSE;

    // if it still isn't our time, update our countdown and return
    if(dwCurTimeMS < pImgAnimState->dwNextTimeMS)
    {
        *pdwFrameTimeMS = pImgAnimState->dwNextTimeMS - dwCurTimeMS;
        return FALSE;
    }

    iLateTimeMS = dwCurTimeMS - pImgAnimState->dwNextTimeMS;

    // See if this is an ART SlideShow    
    CArtPlayer* pArtPlayer = GetArtPlayer();

    if((!TstFlags(DWNLOAD_COMPLETE|DWNLOAD_STOPPED)) || (pArtPlayer))
    {
        EnterCriticalSection();
        fCritical = TRUE;
    }

    if(pArtPlayer)
    {
        if(_pImgBits)
        {
            fResult = pArtPlayer->GetArtReport((CImgBitsDIB**)&_pImgBits, _yHei, GetColorMode());
        }
        else if (_pImgTask)
        {
            fResult = _pImgTask->DoTaskGetReport(pArtPlayer);
        }
        else
        {
            fResult = FALSE;
        }

        // catch the case where we're even late for the next frame
        if(iLateTimeMS > (int)pArtPlayer->_uiUpdateRate)
        {
            *pdwFrameTimeMS = 1;
        }
        else
        {
            *pdwFrameTimeMS = pArtPlayer->_uiUpdateRate - iLateTimeMS;
        }

        pImgAnimState->dwNextTimeMS = dwCurTimeMS + *pdwFrameTimeMS;

        // If the show has been stopped and rewound, stop the timer
        if(pArtPlayer->_fRewind && !pArtPlayer->_fPlaying)
        {
            pArtPlayer->_fRewind = FALSE;
            pImgAnimState->fStop = TRUE;
        }

        // If the show is done, rewind it and stop the timer
        if(pArtPlayer->_fIsDone &&
            (pArtPlayer->_ulCurrentTime==pArtPlayer->_ulAvailPlayTime))
        {
            pArtPlayer->DoPlayCommand(IDM_IMGARTREWIND);
            pImgAnimState->fStop = TRUE;
        }
    }
    else
    {
        // okay, its time to move on to the next frame
        if(pImgAnimState->pgfDraw->pgfNext != NULL)
        {
            pImgAnimState->pgfDraw = pImgAnimState->pgfDraw->pgfNext;

            // catch the case where we're even late for the next frame
            if(iLateTimeMS > (int)pImgAnimState->pgfDraw->uiDelayTime)
            {
                *pdwFrameTimeMS = 1;
            }
            else
            {
                *pdwFrameTimeMS = pImgAnimState->pgfDraw->uiDelayTime - iLateTimeMS;
            }

            pImgAnimState->dwNextTimeMS = dwCurTimeMS + *pdwFrameTimeMS;
        }
        else if(!TstFlags(IMGLOAD_COMPLETE))
        {
            fResult = FALSE;
        }
        else if(_gad.fLooped &&
            (_gad.cLoops==0 || pImgAnimState->dwLoopIter<_gad.cLoops))
        {
            // we're looped...
            pImgAnimState->fLoop = TRUE;
            pImgAnimState->dwLoopIter++;
            pImgAnimState->pgfDraw = pImgAnimState->pgfFirst;

            // catch the case where we're even late for the next frame
            if(iLateTimeMS > (int)pImgAnimState->pgfDraw->uiDelayTime)
            {
                *pdwFrameTimeMS = 1;
            }
            else
            {
                *pdwFrameTimeMS = pImgAnimState->pgfDraw->uiDelayTime - iLateTimeMS;
            }

            pImgAnimState->dwNextTimeMS = dwCurTimeMS + *pdwFrameTimeMS;
        }
        // CHROME
        else
        {
            // It appears fStop was not being set on completion
            // This was causing Chrome serious performance problems.
            // Note: no check of Chrome hosted as no access to the
            // document object
            pImgAnimState->fStop = TRUE;
            fResult = FALSE;
        }
    }

    if(fCritical)
    {
        LeaveCriticalSection();
    }

    return fResult;
}

void CImgInfo::InitImgAnimState(IMGANIMSTATE* pImgAnimState)
{
    Assert(pImgAnimState);

    EnterCriticalSection();

    memset(pImgAnimState, 0, sizeof(IMGANIMSTATE));

    // See if this is an ART SlideShow    
    CArtPlayer* pArtPlayer = GetArtPlayer();

    if(pArtPlayer)
    {
        pImgAnimState->dwNextTimeMS = GetTickCount() + pArtPlayer->_uiUpdateRate;
    }
    else
    {
        pImgAnimState->dwLoopIter = 0;
        pImgAnimState->fLoop = FALSE;

        if(_gad.pgf)
        {
            pImgAnimState->pgfFirst = _gad.pgf; // the first image is there
            pImgAnimState->pgfDraw = _gad.pgf;
        }
        else if(_pImgTask)
        {
            pImgAnimState->pgfFirst = _pImgTask->GetPgf(); // the first image is there
            pImgAnimState->pgfDraw = _pImgTask->GetPgf();
        }

        if(pImgAnimState->pgfDraw)
        {
            pImgAnimState->dwNextTimeMS = GetTickCount() + pImgAnimState->pgfDraw->uiDelayTime;
        }
        else
        {
            pImgAnimState->dwNextTimeMS = (DWORD)-1;
        }
    }

    LeaveCriticalSection();
}

// Callbacks ------------------------------------------------------------------
void CImgInfo::OnLoadTask(CImgLoad* pImgLoad, CImgTask* pImgTask)
{
    EnterCriticalSection();

    if(pImgLoad == _pDwnLoad)
    {
        Assert(!TstFlags(IMGLOAD_COMPLETE));
        Assert(!_pImgTask);

        _pImgTask = pImgTask;
        _pImgTask->AddRef();
    }

    LeaveCriticalSection();
}

void CImgInfo::OnLoadDone(HRESULT hrErr)
{
    Assert(EnteredCriticalSection());

    if(!_pImgTask && TstFlags(DWNLOAD_LOADING))
    {
        UpdFlags(DWNLOAD_MASK,
            TstFlags(DWNF_DOWNLOADONLY)?DWNLOAD_COMPLETE:DWNLOAD_ERROR);
        Signal(IMGCHG_VIEW|IMGCHG_COMPLETE, TRUE, 0);
    }
}

void CImgInfo::OnTaskSize(CImgTask* pImgTask, LONG xWid, LONG yHei, long lTrans, MIMEINFO* pmi)
{
    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        if(xWid!=_xWid || yHei!=_yHei || lTrans!=_lTrans)
        {
            _xWid   = xWid;
            _yHei   = yHei;
            _lTrans = lTrans;
            _pmi    = pmi;
            Signal(IMGCHG_SIZE, FALSE, 0);
        }
    }

    LeaveCriticalSection();
}

void CImgInfo::OnTaskTrans(CImgTask* pImgTask, long lTrans)
{
    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        if(lTrans != _lTrans)
        {
            _lTrans = lTrans;
            Signal(IMGCHG_VIEW, FALSE, 0);
        }
    }

    LeaveCriticalSection();
}

void CImgInfo::OnTaskProg(CImgTask* pImgTask, ULONG ulBits, BOOL fAll, LONG yBot)
{
    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        UpdFlags(IMGBITS_MASK, ulBits);
        Signal(IMGCHG_VIEW, fAll, yBot);
    }

    LeaveCriticalSection();
}

void CImgInfo::OnTaskAnim(CImgTask* pImgTask)
{
    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        SetFlags(IMGANIM_ANIMATED);
        super::Signal(IMGCHG_ANIMATE);
    }

    LeaveCriticalSection();
}

BOOL CImgInfo::OnTaskBits(
          CImgTask* pImgTask,
          CImgBits* pImgBits,
          GIFANIMDATA* pgad,
          CArtPlayer* pArtPlayer,
          LONG lTrans,
          LONG ySrcBot,
          BOOL fNonProgressive)
{
    BOOL fResult = FALSE;

    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        if(TstFlags(IMGLOAD_LOADING))
        {
            WORD wSig = IMGCHG_COMPLETE | (TstFlags(IMGANIM_ANIMATED)?IMGCHG_ANIMATE:0);

            UpdFlags(IMGLOAD_MASK|IMGBITS_MASK, IMGLOAD_COMPLETE|IMGBITS_TOTAL);
            Signal(wSig|(fNonProgressive?IMGCHG_VIEW:0), FALSE, 0);
        }

        if(pgad)
        {
            _gad = *pgad;
        }

        if(pArtPlayer)
        {
            _pArtPlayer = pArtPlayer;

            // Keep dynamic art out of cache
            memset(&_ftLastMod, 0, sizeof(_ftLastMod));
        }

        _ySrcBot    = ySrcBot;
        _lTrans     = lTrans;
        _pImgBits   = pImgBits;
        fResult     = TRUE;

        if(_ySrcBot==-1 && !pImgBits->IsTransparent())
        {
            UpdFlags(IMGTRANS_MASK, IMGTRANS_OPAQUE);
        }

        if(!_fNoOptimize)
        {
            _pImgBits->Optimize();
        }
    }

    LeaveCriticalSection();

    return fResult;
}

void CImgInfo::OnTaskDone(CImgTask* pImgTask)
{
    CDwnLoad* pDwnLoad = NULL;

    EnterCriticalSection();

    if(pImgTask == _pImgTask)
    {
        if(TstFlags(IMGLOAD_LOADING))
        {
            UpdFlags(IMGLOAD_MASK, IMGLOAD_ERROR);
            Signal(IMGCHG_VIEW|IMGCHG_COMPLETE, TRUE, 0);
        }

        _pImgTask = NULL;
        pDwnLoad  = _pDwnLoad;
#if (SPVER>=0x01)
        if(pDwnLoad)
        {
            pDwnLoad->AddRef();
        }
#endif
    }
    else
    {
        pImgTask = NULL;
    }

    LeaveCriticalSection();

    if(pImgTask)
    {
        pImgTask->Release();
    }

    if(pDwnLoad)
    {
        pDwnLoad->OnDone(S_OK);
#if (SPVER>=0x01)
        pDwnLoad->Release();
#endif
    }
}

HRESULT CImgInfo::DrawImage(HDC hdc, RECT* prcDst, RECT* prcSrc, DWORD dwRop, DWORD dwFlags)
{
    RECT rcSrc;
    BOOL fCritical = FALSE;

    if(!TstFlags(DWNLOAD_COMPLETE|DWNLOAD_STOPPED))
    {
        EnterCriticalSection();
        fCritical = TRUE;
    }

    if(prcSrc == NULL)
    {
        rcSrc.left   = 0;
        rcSrc.top    = 0;
        rcSrc.right  = _xWid;
        rcSrc.bottom = _yHei;
        prcSrc       = &rcSrc;
    }

    if(_pImgBits)
    {
        if(_gad.pgf)
        {
            _gad.pgf->pibd->StretchBltOffset(hdc, prcDst, prcSrc, _gad.pgf->left, _gad.pgf->top, dwRop, dwFlags);
        }
        else
        {
            _pImgBits->StretchBlt(hdc, prcDst, prcSrc, dwRop, dwFlags);
        }
    }
    else if(_pImgTask)
    {
        CImgTask* pImgTask = _pImgTask;
        pImgTask->SubAddRef();

        if(fCritical)
        {
            LeaveCriticalSection();
            fCritical = FALSE;
        }

        pImgTask->BltDib(hdc, prcDst, prcSrc, dwRop, dwFlags);
        pImgTask->SubRelease();
    }

    if(fCritical)
    {
        LeaveCriticalSection();
    }

    return S_OK;
}

HRESULT CImgInfo::SaveAsBmp(IStream* pStm, BOOL fFileHeader)
{
    HRESULT hr = S_OK;

    EnterCriticalSection();

    if(_pImgBits)
    {
        hr = _pImgBits->SaveAsBmp(pStm, fFileHeader);
    }

    LeaveCriticalSection();

    RRETURN(hr);
}

// Caching --------------------------------------------------------------------
ULONG CImgInfo::ComputeCacheSize()
{
    ULONG cb = 0;

    if(_gad.pgf)
    {
        GIFFRAME* pgf = _gad.pgf;

        for(; pgf; pgf=pgf->pgfNext)
        {
            cb += pgf->pibd->CbTotal();
        }
    }
    else if(_pImgBits)
    {
        cb = _pImgBits->CbTotal();
    }

    return cb;
}

BOOL CImgInfo::AttachEarly(UINT dt, DWORD dwRefresh, DWORD dwFlags, DWORD dwBindf)
{
    // In order to attach to an existing CImgInfo, the following must match:
    //      _cstrUrl            (Already checked by caller)
    //      _dwRefresh
    //      DWNF_COLORMODE
    //      DWNF_DOWNLOADONLY
    //      DWNF_FORCEDITHER
    //      BINDF_OFFLINEOPERATION
    Assert(dt == DWNCTX_IMG);

    return(GetRefresh() == dwRefresh
        && GetFlags(DWNF_COLORMODE|DWNF_DOWNLOADONLY|DWNF_FORCEDITHER)==
        (dwFlags&(DWNF_COLORMODE|DWNF_DOWNLOADONLY|DWNF_FORCEDITHER))
        && (GetBindf()&BINDF_OFFLINEOPERATION)==(dwBindf&BINDF_OFFLINEOPERATION));
}

BOOL CImgInfo::CanAttachLate(CDwnInfo* pDwnInfo)
{
    CImgInfo* pImgInfo = (CImgInfo*)pDwnInfo;

    return (pImgInfo->GetFlags(DWNF_COLORMODE|DWNF_DOWNLOADONLY|DWNF_FORCEDITHER) ==
        GetFlags(DWNF_COLORMODE|DWNF_DOWNLOADONLY|DWNF_FORCEDITHER));
}

void CImgInfo::AttachLate(CDwnInfo* pDwnInfo)
{
    CImgInfo* pImgInfo = (CImgInfo*)pDwnInfo;

    Assert(_pDwnInfoLock == NULL);

    _xWid         = pImgInfo->_xWid;
    _yHei         = pImgInfo->_yHei;
    _ySrcBot      = pImgInfo->_ySrcBot;
    _pImgBits     = pImgInfo->_pImgBits;
    _lTrans       = pImgInfo->_lTrans;
    _gad          = pImgInfo->_gad;
    _pArtPlayer   = pImgInfo->_pArtPlayer;
    _pmi          = pImgInfo->_pmi;
    _dwSecFlags   = pImgInfo->_dwSecFlags;
    _pDwnInfoLock = pDwnInfo;
    _pDwnInfoLock->SubAddRef();

    UpdFlags(IMGLOAD_MASK|IMGBITS_MASK|IMGTRANS_MASK|IMGANIM_MASK,
        pDwnInfo->GetFlags(IMGLOAD_MASK|IMGBITS_MASK|IMGTRANS_MASK|IMGANIM_MASK));

    Signal(TstFlags(IMGANIM_ANIMATED) ? IMGCHG_SIZE|IMGCHG_VIEW|IMGCHG_ANIMATE|IMGCHG_COMPLETE : 
        IMGCHG_SIZE|IMGCHG_VIEW|IMGCHG_COMPLETE, TRUE, 0);
}