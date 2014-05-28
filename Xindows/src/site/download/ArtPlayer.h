
#ifndef __XINDOWS_SITE_DOWNLOAD_ARTPLAYER_H__
#define __XINDOWS_SITE_DOWNLOAD_ARTPLAYER_H__

class CArtPlayer
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    ~CArtPlayer();

    BOOL    QueryPlayState(int iCommand);
    void    DoPlayCommand(int iCommand);

    BOOL    GetArtReport(CImgBitsDIB** ppibd, LONG _yHei, LONG _colorMode);

    RECT        _rcUpdateRect;
    UINT        _uiUpdateRate;
    ULONG       _ulCurrentTime;
    ULONG       _ulAvailPlayTime;
    DWORD_PTR   _dwShowHandle;
    BOOL        _fTemporalART:1;
    BOOL        _fHasSound:1;
    BOOL        _fDynamicImages:1;
    BOOL        _fPlaying:1;
    BOOL        _fPaused:1;
    BOOL        _fRewind:1;
    BOOL        _fIsDone:1;
    BOOL        _fUpdateImage:1;
    BOOL        _fInPlayer:1;
};

#endif //__XINDOWS_SITE_DOWNLOAD_ARTPLAYER_H__