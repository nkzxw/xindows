
#ifndef __XINDOWS_SITE_DOWNLOAD_IMGANIM_H__
#define __XINDOWS_SITE_DOWNLOAD_IMGANIM_H__

CImgAnim* GetImgAnim();
CImgAnim* CreateImgAnim();

enum
{
    ANIMSYNC_GETIMGCTX,
    ANIMSYNC_GETHWND,
    ANIMSYNC_TIMER,
    ANIMSYNC_INVALIDATE
};

enum ANIMSTATE
{
    ANIMSTATE_PLAY,
    ANIMSTATE_PAUSE,
    ANIMSTATE_STOP
};

class CAnimSync
{
public:

    DECLARE_MEMALLOC_NEW_DELETE()

    CAnimSync() {}

    typedef void (*ASCALLBACK)(void* pvObj, DWORD dwReason,
        void* pvArg, void** ppvDataOut, IMGANIMSTATE* pAnimState);

    BOOL IsEmpty();
    CImgCtx* GetImgCtx();

    HRESULT Register(void* pvObj, DWORD_PTR dwDocId, DWORD_PTR dwImgId,
        CAnimSync::ASCALLBACK pfnCallback, void* pvArg);
    void Unregister(void* pvObj);

    void OnTimer(DWORD* pdwFrameTimeMS);
    void Update(HWND* pHwnd);
    void Invalidate();

    IMGANIMSTATE   _imgAnimState;
    DWORD_PTR      _dwDocId;
    DWORD_PTR      _dwImgId;
    ANIMSTATE      _state;
    BOOL           _fInvalidated : 1;

private:
    struct CLIENT
    {
        void*      pvObj;
        ASCALLBACK pfnCallback;
        void*      pvArg;
    };

    CStackDataAry<CLIENT, 1> _aryClients;
};

class CImgAnim
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    CImgAnim();
    ~CImgAnim();

    IMGANIMSTATE* GetImgAnimState(LONG lCookie);

    HRESULT RegisterForAnim(void* pObj, DWORD_PTR dwDocId, DWORD_PTR dwImgId,
        CAnimSync::ASCALLBACK callback, void* puData, LONG* plCookie);
    void UnregisterForAnim(void* pObj, LONG lCookie);
    void OnTimer();
    void ProgAnim(LONG lCookie);

    void SetAnimState(DWORD_PTR dwDocId, ANIMSTATE state);

    void StartAnim(LONG lCookie);
    void StopAnim(LONG lCookie);

private:
    CAnimSync* GetAnimSync(LONG lCookie);
    void SetInterval(DWORD dwInterval);

    HRESULT FindOrCreateAnimSync(DWORD_PTR dwDocId, DWORD_PTR dwImgId, LONG* plCookie, CAnimSync** ppAnimSync);
    void CleanupAnimSync(LONG lCookie);

    CPtrAry<CAnimSync*> _aryAnimSync;
    DWORD _dwInterval;
};

#endif //__XINDOWS_SITE_DOWNLOAD_IMGANIM_H__