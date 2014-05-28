
#ifndef __XINDOWS_CORE_BASE_SERVER_H__
#define __XINDOWS_CORE_BASE_SERVER_H__

class CInPlace;

enum OLE_SERVER_STATE
{
    OS_PASSIVE,
    OS_LOADED,      // handler but no server
    OS_RUNNING,     // server running, invisible
    OS_INPLACE,     // server running, inplace-active, no U.I.
    OS_UIACTIVE,    // server running, inplace-active, w/ U.I.
    OS_OPEN         // server running, open-edited
};

//+---------------------------------------------------------------
//
//  Flag values for CServer::CLASSDESC::_dwFlags
//
//----------------------------------------------------------------
enum SERVERDESC_FLAG
{
    SERVERDESC_INVAL_ON_RESIZE      = 1                              << 1,
    SERVERDESC_CREATE_UNDOMGR       = SERVERDESC_INVAL_ON_RESIZE     << 1,
    SERVERDESC_ACTIVATEONENTRY      = SERVERDESC_CREATE_UNDOMGR      << 1,
    SERVERDESC_DEACTIVATEONLEAVE    = SERVERDESC_ACTIVATEONENTRY     << 1,
    SERVERDESC_ACTIVATEONDRAG       = SERVERDESC_DEACTIVATEONLEAVE   << 1,
    SERVERDESC_SUPPORT_DRAG_DROP    = SERVERDESC_ACTIVATEONDRAG      << 1,
    SERVERDESC_DRAGDROP_DESIGNONLY  = SERVERDESC_SUPPORT_DRAG_DROP   << 1,
    SERVERDESC_DRAGDROP_EVENTONLY   = SERVERDESC_DRAGDROP_DESIGNONLY << 1,
    SERVERDESC_HAS_MENU             = SERVERDESC_DRAGDROP_EVENTONLY  << 1,
    SERVERDESC_HAS_TOOLBAR          = SERVERDESC_HAS_MENU            << 1,
    SERVERDESC_UIACTIVE_DESIGNONLY  = SERVERDESC_HAS_TOOLBAR         << 1,
    SERVERDESC_LAST                 = SERVERDESC_UIACTIVE_DESIGNONLY,
    SERVERDESC_MAX                  = LONG_MAX // needed to force enum to be dword
};

//+---------------------------------------------------------------------------
//
//  Flag values for CServer::OnPropertyChange
//
//----------------------------------------------------------------------------
enum SERVERCHNG_FLAG
{
    SERVERCHNG_NOPROPCHANGE = 1 << 1,
    SERVERCHNG_NOVIEWCHANGE = SERVERCHNG_NOPROPCHANGE << 1,
    SERVERCHNG_NODATACHANGE = SERVERCHNG_NOVIEWCHANGE << 1,
    SERVERCHNG_LAST         = SERVERCHNG_NODATACHANGE,
    SERVERCHNG_MAX          = LONG_MAX // needed to force enum to be dword
};

//+---------------------------------------------------------------
//
//  Flag values for CServer::CLock
//
//----------------------------------------------------------------
enum SERVERLOCK_FLAG
{
    SERVERLOCK_STABILIZED       = 1,
    SERVERLOCK_TRANSITION       = SERVERLOCK_STABILIZED    << 1,
    SERVERLOCK_PROPNOTIFY       = SERVERLOCK_TRANSITION    << 1,
    SERVERLOCK_INONPROPCHANGE   = SERVERLOCK_PROPNOTIFY    << 1,
    // general purpose recursion blocking flag
    SERVERLOCK_BLOCKRECURSION   = SERVERLOCK_INONPROPCHANGE << 1,
    SERVERLOCK_BLOCKPAINT       = SERVERLOCK_BLOCKRECURSION << 1,
    SERVERLOCK_IGNOREERASEBKGND = SERVERLOCK_BLOCKPAINT     << 1,
    SERVERLOCK_VIEWCHANGE       = SERVERLOCK_IGNOREERASEBKGND << 1,
    SERVERLOCK_LAST             = SERVERLOCK_VIEWCHANGE,
    SERVERLOCK_MAX              = LONG_MAX // Force enum to 32 bits
};

class NOVTABLE CServer : public CBase
{
    DECLARE_CLASS_TYPES(CServer, CBase)

//private:
public:
    DECLARE_MEMCLEAR_NEW_DELETE();

public:
    NV_DECLARE_TEAROFF_METHOD(GetWindow , getwindow , (HWND* lphwnd));
    STDMETHOD(SetObjectRects)(LPCRECT lprcPosRect, LPCRECT lprcClipRect);
    STDMETHOD(OnWindowMessage)(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult);
    NV_DECLARE_TEAROFF_METHOD(GetDropTarget, getdroptarget, (IDropTarget** ppDropTarget));
    STDMETHOD(TranslateAccelerator)(LPMSG lpmsg);

    // IDropTarget methods
    STDMETHOD(DragEnter)(IDataObject* pIDataSource, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
    STDMETHOD(DragOver)(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
    STDMETHOD(DragLeave)(BOOL fDrop);
    STDMETHOD(Drop)(IDataObject* pIDataSource, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);

    // CBase overrides
    DECLARE_PRIVATE_QI_FUNCS(CBase)
    ULONG SubAddRef()       { return CBase::SubAddRef(); }
    ULONG SubRelease()      { return CBase::SubRelease();}
    ULONG GetRefs()         { return CBase::GetRefs();   }
    ULONG GetObjectRefs()   { return CBase::GetObjectRefs(); }

    HRESULT FireEvent(DISPID dispidEvent, DISPID dispidProp,
        IDispatch* pEventObject, BYTE* pbTypes, ...);

    virtual HRESULT ActivateInPlace(HWND hWndSite, LPMSG lpmsg);
    virtual HRESULT DeactivateInPlace();

    // helper functions for IOleInPlaceSiteWindowless
    HRESULT GetDC(LPRECT prc, DWORD dwFlags, HDC* phDC);
    HRESULT ReleaseDC(HDC hDC);
    HRESULT InvalidateRect(LPCRECT prc, BOOL fErase);
    HRESULT InvalidateRgn(HRGN hrgn, BOOL fErase);
    HRESULT Scroll(int dx, int dy, LPCRECT prcScroll, LPCRECT prcClip);
    HRESULT AdjustRect(LPRECT prc);

    HRESULT SetCapture(BOOL fCaptured);     // Capture or release the mouse.
    BOOL    GetCapture();                   // Have we have captured the mouse?
    HRESULT SetFocus(BOOL fFocus);          // Set focus to self or release.
    BOOL    GetFocus();                     // Helper function to determine if client site has focus

    // Window management
    HWND            GetHWND();
    virtual HRESULT AttachWin(HWND hWndParent, RECT* prc, HWND* phWnd);
    virtual void    DetachWin();

    // UI Active Border implementation
    BOOL            OnNCSetCursor(HWND hwnd, int nHitTest, UINT msg);
    LONG            OnNCHitTest(POINTS);

    // Window procedure
    virtual void    OnPaint();
    virtual BOOL    OnEraseBkgnd(HDC hdc);
    virtual void    OnDestroy();
    HRESULT         OnDefWindowMessage(UINT, WPARAM, LPARAM, LRESULT*);

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

    // Helper functions for translate accelerator
    virtual HRESULT DoTranslateAccelerator(LPMSG lpmsg) { return E_NOTIMPL; }

    // Inline access to internal state
    BOOL TestLock(WORD wLockFlags);
    WORD Unlock(WORD wLockFlags);
    void Relock(WORD wLockFlags);

    CServer();
    virtual ~CServer();

    virtual HRESULT RunningToInPlace(HWND hWndSite, LPMSG lpmsg);
    virtual HRESULT InPlaceToRunning();

    virtual HRESULT EnsureInPlaceObject();

    // Helper methods for derived classes
    virtual HPALETTE GetPalette(HDC hdc=0, BOOL* pfHtPal=NULL);

    virtual HRESULT OnPropertyChange(DISPID dispidProperty, DWORD dwFlags);


    SIZEL           _sizel;             // Our Extent
    CInPlace*       _pInPlace;          // Pointer to our inplace object.
    WORD            _wLockFlags;        // Used by CLock class.
    static ATOM     s_atomWndClass;

    //+----------------------------------------------------------
    //
    //  CServer::CLock
    //
    //-----------------------------------------------------------
    class CLock
    {
    public:
        CLock(CServer* pServer, WORD wLockFlags);
        CLock(CServer* pServer);
        ~CLock();

    private:
        CServer* _pServer;
        WORD _wLockFlags;
    };

    //+----------------------------------------------------------
    //
    //  CServer::CLASSDESC
    //
    //-----------------------------------------------------------
    struct CLASSDESC
    {
        CBase::CLASSDESC _classdescBase;

        BOOL TestFlag(SERVERDESC_FLAG dw) const { return (_classdescBase._dwFlags&dw)!=0; }
    };

    const CLASSDESC* ServerDesc() const { return (const CLASSDESC*)BaseDesc(); }
};

//+------------------------------------------------------------------------
//
//  Member:     CServer::TestLock
//
//  Synopsis:   Returns true if given resource is locked.
//
//  Returns:    BOOL
//
//-------------------------------------------------------------------------
inline BOOL CServer::TestLock(WORD wLockFlags)
{
    return wLockFlags&_wLockFlags;
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::Unlock
//
//  Synopsis:   Unlock resource temporarily.
//
//  Returns:    BOOL
//
//-------------------------------------------------------------------------
inline WORD CServer::Unlock(WORD wLockFlags)
{
    WORD wLockFlagsOld;

    wLockFlagsOld = _wLockFlags;
    _wLockFlags &= ~wLockFlags;
    return wLockFlagsOld;
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::Relock
//
//  Synopsis:   Relock resource that was temporarily unlocked.
//
//-------------------------------------------------------------------------
inline void CServer::Relock(WORD wLockFlags)
{
    _wLockFlags = wLockFlags;
}

#endif //__XINDOWS_CORE_BASE_SERVER_H__