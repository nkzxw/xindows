
#ifndef __XINDOWS_SITE_MISCELEM_IEMEDIAPLAYER_H__
#define __XINDOWS_SITE_MISCELEM_IEMEDIAPLAYER_H__

interface IGraphBuilder;

class CIEMediaPlayer : public IUnknown
{
protected:
    void DeleteContents(void);
    void StopThread();

    // Implementation
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    DECLARE_STANDARD_IUNKNOWN(CIEMediaPlayer);

    // Attributes
    enum state
    {
        IEMM_Uninitialized, 
        IEMM_Initialized,
        IEMM_Downloading, 
        IEMM_Stopped, 
        IEMM_Paused, 
        IEMM_Playing,
        IEMM_Completed, 
        IEMM_Aborted
    };

    BOOL CanPlay() { return _pGraph && _fDataDownloaded && (_fState==IEMM_Stopped || _fState==IEMM_Paused || _fState==IEMM_Completed); }
    BOOL CanStop() { return _pGraph && _fDataDownloaded && (_fState==IEMM_Playing || _fState==IEMM_Paused); }
    BOOL CanPause(){ return _pGraph && _fDataDownloaded && (_fState==IEMM_Playing || _fState==IEMM_Stopped); }
    BOOL CanSeek() { return _pGraph && _fDataDownloaded && (_fState==IEMM_Stopped || _fState==IEMM_Completed); }
    BOOL IsInitialized(){ return _fState!=IEMM_Uninitialized; }

    BOOL IsAudio() { return _fHasAudio && !_fHasVideo; }

    virtual ~CIEMediaPlayer();
    CIEMediaPlayer();

    int     GetStatus();
    HRESULT SetLoopCount(long);
    HRESULT SetVisible(BOOL);
    HRESULT SetVolume(long);
    HRESULT SetBalance(long);
    long    GetVolume();
    long    GetBalance();

    HRESULT Initialize(void);
    HRESULT SetURL(const TCHAR*);
    HRESULT Play();
    HRESULT Pause();
    HRESULT Stop();
    HRESULT Abort();
    HRESULT Seek(ULONG);
    void    SetSegmentSeekFlags(LONG* plSegmentSeek);

    HRESULT SetVideoWindow(HWND hwnd);
    HRESULT SetWindowPosition(RECT* prc);
    HRESULT GetSize(SIZE* psize);

    HRESULT NotifyEvent(void);

    HRESULT SetNotifyWindow(HWND, long, long);

protected:
    state   _fState;
    long    _lLoopCount;
    long    _lPlaysDone;
    long    _lOriginalVol;
    long    _lOriginalBal;
    long    _xWidth;
    long    _yHeight;

    BOOL    _fRestoreVolume: 1;
    BOOL    _fDataDownloaded: 1;
    BOOL    _fHasAudio: 1;
    BOOL    _fHasVideo: 1;
    BOOL    _fUseSegments: 1;

    TCHAR*  _pchURL;
    HWND    _hwndOwner;
    IGraphBuilder* _pGraph;
};

#endif //__XINDOWS_SITE_MISCELEM_IEMEDIAPLAYER_H__