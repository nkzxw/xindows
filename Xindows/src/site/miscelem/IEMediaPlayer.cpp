
#include "stdafx.h"
#include "IEMediaPlayer.h"

#include <strmif.h>
#include <control.h>
#include <evcode.h>

#define DEFAULT_VIDEOWIDTH  32 // bugbug default size?
#define DEFAULT_VIDEOHEIGHT 32

const CLSID CLSID_FilterGraph =
{ 0xe436ebb3, 0x524f, 0x11ce, { 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70 } };

CIEMediaPlayer::CIEMediaPlayer()
{
    _ulRefs = 1;                // born with 1
    _fState = IEMM_Uninitialized;
    _pGraph = NULL;
    _pchURL = NULL;
    _hwndOwner = NULL;
    _fHasAudio = FALSE;
    _fHasVideo = FALSE;

    _fDataDownloaded = FALSE;
    _fRestoreVolume = FALSE;
    _lLoopCount = 1;
    _lPlaysDone = 0;
    _lOriginalVol = 1000;       // init it out of range 
    _lOriginalBal = -100000;    // ditto

    _xWidth = DEFAULT_VIDEOWIDTH; // bugbug default size?
    _yHeight = DEFAULT_VIDEOHEIGHT;
}

CIEMediaPlayer::~CIEMediaPlayer()
{
    if(_fRestoreVolume)
    {
        SetVolume(_lOriginalVol);
        SetBalance(_lOriginalBal);
        Stop();
    }

    DeleteContents();
}

HRESULT CIEMediaPlayer::QueryInterface(REFIID riid, LPVOID* ppv)
{ 
    if(riid == IID_IUnknown)
    {
        *ppv = (IUnknown*)this;
        AddRef();
        return S_OK;
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
}

void CIEMediaPlayer::DeleteContents(void)
{
    if(_pGraph) 
    {
        if(_hwndOwner)
        {
            HRESULT hr;
            IVideoWindow* pVW = NULL;
            hr = _pGraph->QueryInterface(IID_IVideoWindow, (void**)&pVW);
            if(OK(hr))
            {
                pVW->put_MessageDrain((OAHWND)NULL);
                pVW->Release();
                _hwndOwner = NULL;
            }
        }
        _pGraph->Release();
        _pGraph = NULL;
    }

    if(_pchURL)
    {
        MemFreeString(_pchURL);
        _pchURL = NULL;
    }

    _fState = IEMM_Uninitialized;
}

HRESULT CIEMediaPlayer::Initialize(void) 
{
    HRESULT hr;             // return code

    if(_pGraph)             // already initialized
    {
        _pGraph->Release(); // go away
        _pGraph = NULL;
    }


    hr = CoCreateInstance(
        CLSID_FilterGraph,  // get this documents graph object
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_IGraphBuilder,
        (void**)&_pGraph);

    if(FAILED(hr)) 
    {
        DeleteContents();
        return hr;
    }

    _fState = IEMM_Initialized;
    return S_OK;
}

HRESULT CIEMediaPlayer::SetURL(const TCHAR* pchURL)
{
    HRESULT hr = S_OK, hr2 = S_OK;
    IVideoWindow* pVW;

    if(!pchURL)
    {
        return ERROR_INVALID_PARAMETER;
    }

    MemReplaceString(pchURL, &_pchURL);

    if(_pGraph)
    {
        Stop();                 // We already have a graph built so call stop in case it's running
    }

    hr = Initialize();          // This will release the Graph and CoCreateInstance() a new one.

    if(FAILED(hr))
    {
        goto Failed;
    }

    _fDataDownloaded = TRUE;    // we we're passed in a valid URL

    // Build the graph.
    //
    // This won't return until the the file type is sniffed and the appropriate
    // graph is built
    //
    hr = _pGraph->RenderFile(_pchURL, NULL);
    if(SUCCEEDED(hr))
    {
        _fState = IEMM_Stopped;
    }
    else
    {
        _fState = IEMM_Aborted;
        goto Failed;
    }

    // Need to check to see if there's a Video renderer interface and
    // shut it off if it's there. Our default state is to show no window until
    // someone sets our window position
    //
    // For BGSound this allows video files to be used without having a video window
    //  pop up on us.
    // For DYNSRC a video window size will be set at which point we'll 
    //  turn the thing on again.
    hr = _pGraph->QueryInterface(IID_IVideoWindow, (void**)&pVW);
    if(OK(hr)) 
    {
        long lVisible;

        // if this fails then we have an audio only stream
        hr2 = pVW->get_Visible(&lVisible);  
        if(hr2 == S_OK)
        {
            SIZE size;
            GetSize(&size);             // this will get the size of the video source
                                        // and cache the results for later

            hr2 = pVW->put_AutoShow(0); // turn off the auto show of the video window
            _fHasVideo = TRUE;
        }
        else
        {
            _xWidth = 0;
            _yHeight = 0;
        }
        pVW->Release();
    }

    _fUseSegments = FALSE;

    IMediaSeeking* pIMediaSeeking;
    hr = _pGraph->QueryInterface(IID_IMediaSeeking, (void**)&pIMediaSeeking);
    if(SUCCEEDED(hr))
    {
        // See if Segment seeking is supported (for Seamless looping)
        if(pIMediaSeeking)
        {
            DWORD dwCaps = AM_SEEKING_CanDoSegments;
            _fUseSegments = (S_OK == pIMediaSeeking->CheckCapabilities(&dwCaps));
        } 
        pIMediaSeeking->Release();
    }

    IBasicAudio* pIBa;
    long lOriginalVolume, lOriginalBalance;

    lOriginalVolume = lOriginalBalance = 0;

    hr = _pGraph->QueryInterface(IID_IBasicAudio, (void**)&pIBa);
    if(SUCCEEDED(hr))
    {
        hr2 = pIBa->get_Volume(&lOriginalVolume);
        if(hr2 == S_OK)
        {
            _fHasAudio = TRUE;
            pIBa->get_Balance(&lOriginalBalance);
        }
        pIBa->Release();
    }

    // save away the original volume so that we can restore it on our way out
    if(_fHasAudio && _lOriginalVol>0)
    {
        _lOriginalBal = lOriginalBalance;
        _lOriginalVol = lOriginalVolume;
    }

Failed:
    return hr;
}

HRESULT CIEMediaPlayer::SetVideoWindow(HWND hwnd)
{
    IVideoWindow* pVW = NULL;
    HRESULT hr  = S_OK;

    if(!_fHasVideo)
    {
        return S_FALSE;
    }

    if(!_pGraph)
    {
        return S_FALSE;
    }

    if(hwnd)
    {
        _hwndOwner = hwnd;

        hr = _pGraph->QueryInterface(IID_IVideoWindow, (void**)&pVW);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = pVW->put_Owner((OAHWND)hwnd);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = pVW->put_MessageDrain((OAHWND)hwnd);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = pVW->put_WindowStyle(WS_CHILDWINDOW);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = pVW->put_BackgroundPalette(-1); // OATRUE
        if(FAILED(hr))
        {
            goto Cleanup;
        }
    }

Cleanup:
    if(pVW)
    {
        pVW->Release();
    }

    RRETURN(hr);
}

HRESULT CIEMediaPlayer::SetWindowPosition(RECT* prc)
{
    HRESULT hr;
    IVideoWindow* pVW = NULL;

    Assert(prc);

    if(!_fHasVideo)
    {
        return S_FALSE;
    }

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IVideoWindow, (void**)&pVW);
    if(FAILED(hr))
    {
        goto Cleanup;
    }

    hr = pVW->SetWindowPosition(
        prc->left,
        prc->top,
        prc->right-prc->left,
        prc->bottom-prc->top);

Cleanup:
    if(pVW)
    {
        pVW->Release();
    }

    RRETURN(hr);

}

HRESULT CIEMediaPlayer::SetVisible(BOOL fVisible)
{
    HRESULT hr;
    IVideoWindow* pVW = NULL;

    if(!_fHasVideo)
    {
        return S_FALSE;
    }

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IVideoWindow, (void**)&pVW);
    if(OK(hr))
    {
        if(fVisible)
        {
            hr = pVW->put_Visible(-1);  // OATRUE
        }
        else 
        {
            hr = pVW->put_Visible(0);   // OAFALSE
        }
    }

    if(pVW)
    {
        pVW->Release();
    }

    return hr;
}

HRESULT CIEMediaPlayer::GetSize(SIZE* psize)
{
    long    lWidth = 0;
    long    lHeight = 0;
    HRESULT hr = S_OK;
    IBasicVideo* pBV = NULL;

    if(!_pGraph)
    {
        return E_FAIL;
    }

    if(_xWidth==DEFAULT_VIDEOWIDTH && _yHeight==DEFAULT_VIDEOHEIGHT)
    {
        hr = _pGraph->QueryInterface(IID_IBasicVideo, (void**)&pBV);

        if(OK(hr))
        {
            hr = pBV->get_SourceWidth(&lWidth);
        }

        if(OK(hr))
        {
            hr = pBV->get_SourceHeight(&lHeight);
        }

        if(OK(hr))
        {
            _xWidth = lWidth;
            _yHeight = lHeight;
        }
    }

    psize->cx = _xWidth;
    psize->cy = _yHeight;

    if(pBV)
    {
        pBV->Release();
    }

    return hr;
}

HRESULT CIEMediaPlayer::SetNotifyWindow(HWND hwnd, long lmsg, long lParam)
{
    HRESULT hr = S_OK;
    IMediaEventEx* pMvEx = NULL;

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IMediaEventEx, (void**)&pMvEx);
    if(FAILED(hr))
    {
        goto Failed;
    }

    hr = pMvEx->SetNotifyWindow((OAHWND)hwnd, lmsg, lParam);

    pMvEx->Release();

Failed:
    return hr;
}

HRESULT CIEMediaPlayer::SetLoopCount(long uLoopCount)
{
    _lLoopCount = uLoopCount;
    _lPlaysDone = 0;
    return S_OK;
}

long CIEMediaPlayer::GetVolume(void)
{
    HRESULT hr;
    IBasicAudio* pIBa;
    long lTheVolume=E_FAIL; // BUGBUG Arye: shouldn't this be zero?

    if(_pGraph)
    {
        hr = _pGraph->QueryInterface(IID_IBasicAudio, (void**)&pIBa);
        if(SUCCEEDED(hr) && pIBa)
        {
            hr = pIBa->get_Volume(&lTheVolume);
            pIBa->Release();
        }
    }
    return lTheVolume;
}

HRESULT CIEMediaPlayer::SetVolume(long lVol)
{
    HRESULT hr = S_OK;
    IBasicAudio* pIBa;

    if(lVol<-10000 && lVol>0)
    {
        return ERROR_INVALID_PARAMETER;
    }

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IBasicAudio, (void**)&pIBa);
    if(SUCCEEDED(hr) && pIBa)
    {
        hr = pIBa->put_Volume(lVol);
        pIBa->Release();
    }
    _fRestoreVolume = TRUE;

    return hr;
}

long CIEMediaPlayer::GetBalance(void)
{
    HRESULT hr;
    IBasicAudio* pIBa;
    long lTheBal=E_FAIL;

    if(_pGraph)
    {
        hr = _pGraph->QueryInterface(IID_IBasicAudio, (void**)&pIBa);
        if(SUCCEEDED(hr) && pIBa)
        {
            hr = pIBa->get_Balance(&lTheBal);
            pIBa->Release();
        }
    }

    return lTheBal;
}   

HRESULT CIEMediaPlayer::SetBalance(long lBal)
{
    HRESULT hr = S_OK;
    IBasicAudio* pIBa;

    if(lBal<-10000 && lBal>10000)
    {
        return ERROR_INVALID_PARAMETER;
    }

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IBasicAudio, (void**)&pIBa);
    if(SUCCEEDED(hr) && pIBa)
    {
        hr = pIBa->put_Balance(lBal);
        pIBa->Release();
    }

    _fRestoreVolume = TRUE;

    return hr;
}

int CIEMediaPlayer::GetStatus(void)
{
    return _fState;
}

HRESULT CIEMediaPlayer::Play()
{
    HRESULT hr = S_OK;

    if(CanPlay())
    {
        IMediaControl* pMC = NULL;

        // Obtain the interface to our filter graph
        hr = _pGraph->QueryInterface(IID_IMediaControl, (void**)&pMC);
        if(SUCCEEDED(hr) && pMC)
        {
            if(_fUseSegments)
            {
                // If we're using seamless looping, we need to 1st set seeking flags
                Seek(0);
            }

            // Ask the filter graph to play 
            hr = pMC->Run();

            if(SUCCEEDED(hr))
            {
                _fState = IEMM_Playing;
                if(_lLoopCount > 0)
                {
                    _lPlaysDone++;
                }
            }
            else
            {
                pMC->Stop(); // some filters in the graph may have started so we better stop them
            }
            pMC->Release();
        }
        else
        {
            hr = S_FALSE;
        }
    }

    RRETURN1(hr, S_FALSE);
}

HRESULT CIEMediaPlayer::Pause()
{
    HRESULT hr = S_OK;

    if(CanPause())
    {
        IMediaControl* pMC;
        // Obtain the interface to our filter graph
        hr = _pGraph->QueryInterface(IID_IMediaControl, (void**)&pMC);

        // Ask the filter graph to pause
        if(SUCCEEDED(hr))
        {
            pMC->Pause();
        }

        pMC->Release();

        _fState = IEMM_Paused;
    }
    else
    {
        hr = S_FALSE;
    }

    RRETURN1(hr, S_FALSE);
}

HRESULT CIEMediaPlayer::Abort()
{
    // Must stop play first.
    Stop();

    _fState = IEMM_Aborted;

    return S_OK; // must not fail
}

HRESULT CIEMediaPlayer::Seek(ULONG uPosition)
{
    HRESULT hr = S_OK;
    IMediaSeeking* pIMediaSeeking=NULL;
    if(_pGraph)
    {
        hr = _pGraph->QueryInterface(IID_IMediaSeeking, (void**)&pIMediaSeeking);
    }

    if(SUCCEEDED(hr) && pIMediaSeeking)
    {
        LONGLONG llStop;

        hr = pIMediaSeeking->GetPositions(NULL, &llStop);
        if(SUCCEEDED(hr) && (llStop>uPosition))
        {
            long lSegmentSeek = 0L;
            LONGLONG llPosition = (LONGLONG) uPosition;

            SetSegmentSeekFlags(&lSegmentSeek); // in case we're seamless looping
            hr = pIMediaSeeking->SetPositions(
                &llPosition,
                AM_SEEKING_AbsolutePositioning|lSegmentSeek,
                &llStop,
                AM_SEEKING_NoPositioning);
        }
        pIMediaSeeking->Release();
    }

    RRETURN1(hr, S_FALSE);
}

void CIEMediaPlayer::SetSegmentSeekFlags(LONG* plSegmentSeek)
{
    if(plSegmentSeek)
    {
        *plSegmentSeek = 
            _fUseSegments&&_lLoopCount!=1 ?
            ((_lLoopCount==-1) ||
            (_lLoopCount>_lPlaysDone+1) ?
            AM_SEEKING_NoFlush|AM_SEEKING_Segment : AM_SEEKING_NoFlush)
            : 0L;
    }
}

HRESULT CIEMediaPlayer::Stop()
{
    HRESULT hr = S_OK;

    if(CanStop())
    {
        IMediaControl* pMC;

        // Obtain the interface to our filter graph
        hr = _pGraph->QueryInterface(IID_IMediaControl, (void**)&pMC);
        if(SUCCEEDED(hr))
        {
            // Stop the filter graph
            hr = pMC->Stop();
            // Release the interface
            pMC->Release();

            // set the flags
            _fState = IEMM_Stopped;
        }
    }
    else
    {
        hr = S_FALSE;
    }

    RRETURN1(hr, S_FALSE);
}

// ======================================================================
//
// If the event handle is valid, ask the graph
// if anything has happened. eg the graph has stopped...
// ======================================================================
HRESULT CIEMediaPlayer::NotifyEvent(void) 
{
    HRESULT hr = S_OK;
    long lEventCode;
    LPARAM lParam1, lParam2;
    IMediaEvent* pME = NULL;

    if(!_pGraph)
    {
        return E_FAIL;
    }

    hr = _pGraph->QueryInterface(IID_IMediaEvent, (void**)&pME); 
    if(FAILED(hr))
    {
        goto GN_Failed;
    }

    hr = pME->GetEvent(&lEventCode, &lParam1, &lParam2, 0);
    if(FAILED(hr))
    {
        goto GN_Failed;
    }

    if(lEventCode==EC_COMPLETE || lEventCode==EC_END_OF_SEGMENT)
    {
        // Do we need to loop?
        if(_lLoopCount==-1 || _lLoopCount>_lPlaysDone)
        {
            Seek(0); // we're still playing so seek back to the begining and we'll keep going
            if(_lLoopCount >0)
            {
                _lPlaysDone++;
            }
        }
        else    
        {
            // we're done stop the graph
            Stop();
            _fState = IEMM_Completed;
        }

    }
    else if((lEventCode==EC_ERRORABORT) || (lEventCode==EC_USERABORT)) 
    {
        Stop();
    }

GN_Failed:
    if(pME) pME->Release();

    RRETURN1(hr, S_FALSE);
}
