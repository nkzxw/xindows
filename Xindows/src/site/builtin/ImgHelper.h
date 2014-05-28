
#ifndef __XINDOWS_SITE_BUILTIN_IMGHELPER_H__
#define __XINDOWS_SITE_BUILTIN_IMGHELPER_H__

class CImageLayout;
class CIEMediaPlayer;
class CBitsCtx;
class CDwnChan;

#define DEF_HSPACE  3

#define DEFINE_IMG_FIRE_BOOL(eventName)         BOOL Fire_##eventName();
#define DEFINE_IMG_FIRE_VOID(eventName)         void Fire_##eventName();
#define DEFINE_IMG_GETAA(returnType, propName)  returnType GetAA##propName () const;
#define DEFINE_IMG_SETAA(paraType, propName)    HRESULT SetAA##propName (paraType pv);

#define IMGF_REQUEST_RESIZE     0x00000001
#define IMGF_INVALIDATE_FRAME   0x00000002

CImgBits* GetPlaceHolderBitmap(BOOL fMissing);
void GetPlaceHolderBitmapSize(BOOL fMissing, SIZE* pSize);
void DrawPlaceHolder(HDC hdc,
                     RECT rectImg,
                     TCHAR* lpString,
                     CODEPAGE codepage,
                     LCID lcid,
                     SHORT sBaselineFont,
                     SIZE* psizeGrab,
                     BOOL fMissing,
                     COLORREF fgColor,
                     COLORREF bgColor,
                     SIZE* psizePrint,
                     BOOL fRTL,
                     DWORD dwFlags);
int DrawTextInCodePage(UINT uCP, HDC hDC, LPCWSTR lpString, int nCount, LPRECT lpRect, UINT uFormat);


class CImgHelper : public CVoid
{
    DECLARE_CLASS_TYPES(CImgHelper, CVoid)

public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    CImgHelper ();
    CImgHelper (CDocument* pDoc, CSite* pSite);

    LONG    GetAAHspace();

    HRESULT EnterTree();
    void    ExitTree(CNotification* pnf);

    // IOleCommandTarget methods
    STDMETHOD(QueryStatus) (
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext);

    STDMETHOD(Exec) (
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut);

    // Context menu gelper
    HRESULT ShowImgContextMenu(CMessage* pMessage);

    // InvokeEx to validate ready state.
    NV_DECLARE_TEAROFF_METHOD(ContextThunk_InvokeExReady, invokeexready, (
        DISPID      id,
        LCID        lcid,
        WORD        wFlags,
        DISPPARAMS* pdp,
        VARIANT*    pvarRes,
        EXCEPINFO*  pei,
        IServiceProvider* pSrvProvider));
    void    Passivate();
    void    CleanupImage();

    void    GetRectImg(CRect* prectImg);

    void    Notify(CNotification* pNF);

    void    SetImgAnim(BOOL fDisplayNoneOrHidden);

    void    SetActivity();
    void    SetVideo();
    void    Replay();

    HRESULT SetReadyStateImg(long readyStateScript);
    void    OnReadyStateChange();

    HRESULT ShowTooltip(CMessage* pmsg, POINT pt);

    HRESULT GetFile(TCHAR** ppchFile);
    HRESULT GetUrl(TCHAR** ppchUrl);

    HRESULT PromptSaveAs(TCHAR* pchFileName=NULL, int cchFileName=0);

    // img & bits ctx
    HRESULT FetchAndSetImgCtx(const TCHAR* pchUrl, DWORD dwSetFlags);
    void    SetImgCtx(CImgCtx* pImgCtx, DWORD dwSetFlags);
    void    SetBitsCtx(CBitsCtx* pBitsCtx);
    void    OnDwnChan(CDwnChan* pDwnChan);
    void    OnAnimSync(DWORD dwReason, void* puData, void** pDataOut, IMGANIMSTATE* pAnimState);

    static void OnAnimSyncCallback(
        void*   pvObj,
        DWORD   dwReason,
        void*   pvArg,
        void**  ppvDataOut,
        IMGANIMSTATE* pAnimState)
    {
        ((CImgHelper*)pvObj)->OnAnimSync(dwReason, pvArg, ppvDataOut, pAnimState);
    }

    static void CALLBACK OnDwnChanCallback(void* pvObj, void* pvArg)
    {
        ((CImgHelper*)pvArg)->OnDwnChan((CDwnChan*)pvObj);
    }

    //--------------------------------------------------------------
    // Property bag and class descriptor
    //--------------------------------------------------------------
    void GetMarginInfo(
        CParentInfo* ppri,
        LONG* plLMargin,
        LONG* plTMargin,
        LONG* plRMargin,
        LONG* plBMargin);

    void CalcSize(CCalcInfo* pci, SIZE* psize);

    //base implementation prototypes
    NV_DECLARE_TEAROFF_METHOD(put_src, PUT_src, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_src, GET_src, (BSTR* p));

    HRESULT STDMETHODCALLTYPE get_height(long* p);
    HRESULT STDMETHODCALLTYPE put_width(long v);
    HRESULT STDMETHODCALLTYPE get_width(long* p);

    NV_DECLARE_TEAROFF_METHOD(get_onload, GET_onload, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onload, PUT_onload, (VARIANT));

    // Attr array prop set helpers, called internally only.
    HRESULT putWidth(long l);
    HRESULT putHeight(long l);
    HRESULT GetHeight(long* pl);
    HRESULT GetWidth(long* pl);

    NV_DECLARE_TEAROFF_METHOD(get_readyState, get_readyState, (BSTR* pBSTR));
    NV_DECLARE_TEAROFF_METHOD(get_readyState, get_readyState, (VARIANT* pReadyState));

    CImgCtx*    GetImgCtx() { return (_pImgCtxPending ? _pImgCtxPending : _pImgCtx); }
    HRESULT     SetImgSrc(DWORD dwSetFlags);
    HRESULT     SetImgDynsrc();

    void        InvalidateFrame();

    CTreeNode*  GetFirstBranch ()
    {
        Assert(_pOwner);
        return _pOwner->GetFirstBranch();
    }

    CImageLayout* Layout ();

    CDocument* Doc()
    {
        Assert(_pOwner);
        return _pOwner->Doc();
    }

    HRESULT EnsureInMarkup()
    {
        Assert(_pOwner);
        RRETURN(_pOwner->EnsureInMarkup());
    }

    void ResizeElement(DWORD grfFlags=0)
    {
        Assert(_pOwner);
        _pOwner->ResizeElement(grfFlags);
    }

    void RemeasureElement(DWORD grfFlags=0)
    {
        Assert(_pOwner);
        _pOwner->RemeasureElement(grfFlags);
    }

    BOOL IsEditable(BOOL fContainerOnly=FALSE)
    {
        Assert(_pOwner);
        return _pOwner->IsEditable(fContainerOnly);
    }

    BOOL IsAligned()
    {
        Assert(_pOwner);
        return _pOwner->IsAligned();
    }

    BOOL HasLayout()
    {
        Assert(_pOwner);
        return _pOwner->HasLayout();
    }

    BOOL IsInMarkup() const
    {
        Assert(_pOwner);
        return _pOwner->IsInMarkup();
    }

    BOOL IsInPrimaryMarkup() const
    {
        Assert(_pOwner);
        return _pOwner->IsInPrimaryMarkup();
    }

    BOOL IsOpaque();
    void Draw(CFormDrawInfo* pDI);

    BOOL IsHSpaceDefined() const;

    DEFINE_IMG_GETAA(LPCTSTR, alt);

    DEFINE_IMG_GETAA(LPCTSTR, src);
    DEFINE_IMG_SETAA(LPCTSTR, src);

    DEFINE_IMG_GETAA(CUnitValue, border);
    DEFINE_IMG_GETAA(long, vspace);
    DEFINE_IMG_GETAA(long, hspace);
    DEFINE_IMG_GETAA(LPCTSTR, lowsrc);
    DEFINE_IMG_GETAA(LPCTSTR, dynsrc);

    DEFINE_IMG_GETAA(htmlStart, start);

    DEFINE_IMG_GETAA(VARIANT_BOOL, complete);
    DEFINE_IMG_SETAA(VARIANT_BOOL, complete);

    DEFINE_IMG_GETAA(long, loop);
    DEFINE_IMG_GETAA(LPCTSTR, onload);
    DEFINE_IMG_GETAA(LPCTSTR, onerror);
    DEFINE_IMG_GETAA(LPCTSTR, onabort);
    DEFINE_IMG_GETAA(LPCTSTR, name);
    DEFINE_IMG_GETAA(LPCTSTR, title);
    DEFINE_IMG_GETAA(VARIANT_BOOL, cache);

    DEFINE_IMG_FIRE_VOID(onerror);
    DEFINE_IMG_FIRE_VOID(onload);
    DEFINE_IMG_FIRE_VOID(onabort);
    DEFINE_IMG_FIRE_VOID(onreadystatechange);
    DEFINE_IMG_FIRE_BOOL(onbeforecopy);
    DEFINE_IMG_FIRE_BOOL(onbeforepaste);
    DEFINE_IMG_FIRE_BOOL(onbeforecut);
    DEFINE_IMG_FIRE_BOOL(oncut);
    DEFINE_IMG_FIRE_BOOL(oncopy);
    DEFINE_IMG_FIRE_BOOL(onpaste);

    CImgCtx*    _pImgCtx;
    CImgCtx*    _pImgCtxPending;
    CBitsCtx*   _pBitsCtx;

public:
    HWND            _hwnd;
    CSite*          _pOwner;
    CIEMediaPlayer* _pVideoObj;
    unsigned        _readyStateImg : 3;
    unsigned        _readyStateFired : 3;
    BOOL            _fCreatedWithNew:1; // TRUE if the image was created with JScript new operator
                                        // and is invisible, parented to the root site.
    BOOL            _fSizeInCtor:1;     // TRUE if the image width and height are specified while creating
    BOOL            _fVideoPositioned : 1;

private:
    BOOL    _fResizePending : 1;
    BOOL    _fIsInPlace : 1;
    BOOL    _fIsActive : 1;
    BOOL    _fSizeChange : 1;
    BOOL    _fStopped : 1;
    BOOL    _fAnimated : 1;
    BOOL    _fDisplayNoneOrHidden : 1;
    BOOL    _fExpandAltText : 1;
    BOOL    _fHideForSecurity : 1;
    BOOL    _fNeedSize : 1;
    BOOL    _fCache : 1;

    LONG    _lCookie;

    HBITMAP _hbmCache;
    LONG    _xWidCache;
    LONG    _yHeiCache;
    LONG    _colorMode;

    HRESULT CacheImage(HDC hdc, CRect* prcDst, SIZE* pSize, DWORD dwFlags, ULONG ulState);
    void    DrawCachedImage(HDC hdc, CRect* prcDst, DWORD dwFlags, ULONG ulState);
};

#endif //__XINDOWS_SITE_BUILTIN_IMGHELPER_H__