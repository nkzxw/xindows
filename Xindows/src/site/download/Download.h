
#ifndef __XINDOWS_SITE_DOWNLOAD_DOWNLOAD_H__
#define __XINDOWS_SITE_DOWNLOAD_DOWNLOAD_H__

#include <IImgCtx.h>

// Declarations ---------------------------------------------------------
class CImgTask;
class CDocument;
class CDwnCtx;
class CDwnInfo;
class CDwnLoad;
class CDwnBindData;
class CDwnDoc;
class CPostItem;
class CArtPlayer;

// Definitions ----------------------------------------------------------
#define DWNLOAD_NOTLOADED       IMGLOAD_NOTLOADED
#define DWNLOAD_LOADING         IMGLOAD_LOADING
#define DWNLOAD_STOPPED         IMGLOAD_STOPPED
#define DWNLOAD_ERROR           IMGLOAD_ERROR
#define DWNLOAD_COMPLETE        IMGLOAD_COMPLETE
#define DWNLOAD_MASK            IMGLOAD_MASK

#define DWNCHG_COMPLETE         IMGCHG_COMPLETE

#define DWNCTX_HTM              0 // DWNTYPE_HTM
#define DWNCTX_IMG              1 // DWNTYPE_IMG
#define DWNCTX_BITS             2 // DWNTYPE_BITS
#define DWNCTX_FILE             3 // DWNTYPE_FILE
#define DWNCTX_MAX              4

#define DWNF_COLORMODE          DWN_COLORMODE        // 0x0000002F
#define DWNF_DOWNLOADONLY       DWN_DOWNLOADONLY     // 0x00000040
#define DWNF_FORCEDITHER        DWN_FORCEDITHER      // 0x00000080
#define DWNF_RAWIMAGE           DWN_RAWIMAGE         // 0x00000100
#define DWNF_MIRRORIMAGE        DWN_MIRRORIMAGE      // 0x00000200

// Note: This block of flags overlaps the state flags on purpose as they are
//       not needed at the same time.
#define DWNF_GETCONTENTTYPE     0x00010000
#define DWNF_GETREFRESH         0x00020000
#define DWNF_GETMODTIME         0x00040000
#define DWNF_GETFILELOCK        0x00080000
#define DWNF_GETFLAGS           0x00100000           // both reqflags and security flags
#define DWNF_ISDOCBIND          0x00200000
#define DWNF_NOAUTOBUFFER       0x00400000
#define DWNF_IGNORESECURITY     0x00800000
#define DWNF_GETSTATUSCODE      0x01000000
#define DWNF_HANDLEECHO         0x02000000
#define DWNF_GETSECCONINFO      0x04000000
#define DWNF_NOOPTIMIZE         0x08000000

#define DWNF_STATE              0xFFFF0000

#define SZ_DWNBINDINFO_OBJECTPARAM      _T("__DWNBINDINFO")
#define SZ_HTMLLOADOPTIONS_OBJECTPARAM  _T("__HTMLLOADOPTIONS")

// Types ----------------------------------------------------------------
#define PFNDWNCHAN PFNIMGCTXCALLBACK
typedef class CImgTask* (NEWIMGTASKFN)();

struct MIMEINFO
{
    CLIPFORMAT      cf;
    LPCTSTR         pch;
    NEWIMGTASKFN*   pfnImg;
    int             ids;
};

struct GIFFRAME;

struct IMGANIMSTATE
{
    DWORD       dwLoopIter;     // Current iteration of looped animation, not actually used for Netscape compliance reasons
    GIFFRAME*   pgfFirst;       // First frame
    GIFFRAME*   pgfDraw;        // Last frame we need to draw
    DWORD       dwNextTimeMS;   // Time to display pgfDraw->pgfNext, or next iteration
    BOOL        fLoop;
    BOOL        fStop;
};

class CTreePos;

struct HTMPASTEINFO
{
    CODEPAGE    cp;             // Source CodePage for pasting

    int         cbSelBegin;
    int         cbSelEnd;

    CTreePos*   ptpSelBegin;
    CTreePos*   ptpSelEnd;

    CString     cstrSourceUrl;
};

struct DWNLOADINFO
{
    CDwnBindData*       pDwnBindData;   // Bind data already in progress
    CDwnDoc*            pDwnDoc;        // CDwnDoc for binding
    IInternetSession*   pInetSess;      // Internet session to use if possible
    LPCTSTR             pchUrl;         // Bind data provided by URL
    IStream*            pstm;           // Bind data provided by IStream
    IMoniker*           pmk;            // Bind data provided by Moniker
    IBindCtx*           pbc;            // IBindCtx to use for binding
    BOOL                fForceInet;     // Force use of pInetSess
    BOOL                fClientData;    // Keep document open
    BOOL                fResynchronize; // Don't believe dwRefresh from pDwnDoc
    BOOL                fUnsecureSource;// Assume source is not https-secure
    DWORD               dwProgClass;    // used by CBitsCtx to override PROGSINK_CLASS (see CBitsInfo::Init)
};

struct DWNDOCINFO
{
    CDwnBindData*       pDwnBindData;   // CDwnBindData awaiting callback
    void*               pvArg;          // Argument for callback
};
DECLARE_CDataAry(CAryDwnDocInfo, DWNDOCINFO);


// Prototypes -----------------------------------------------------------------
HRESULT             NewDwnCtx(UINT dt, BOOL fLoad, DWNLOADINFO* pdli, CDwnCtx** ppDwnCtx);
int                 GetDefaultColorMode();
MIMEINFO*           GetMimeInfoFromClipFormat(CLIPFORMAT cf);
MIMEINFO*           GetMimeInfoFromMimeType(const TCHAR* pchMime);
int                 NormalizerCR(BOOL* pfEndCR, LPTSTR pchStart, LPTSTR* ppchEnd);
int                 NormalizerChar(LPTSTR pchStart, LPTSTR* ppchEnd, BOOL* pfAscii=NULL);
IInternetSession*   TlsGetInternetSession();


class CDwnChan : public CBaseFT
{
    typedef CBaseFT super;

private:
    DECLARE_MEMCLEAR_NEW_DELETE()

public:
    CDwnChan(CRITICAL_SECTION* pcs=NULL);
    void SetCallback(PFNDWNCHAN pfnCallback, void* pvCallback);
    void Disconnect();

#ifdef _DEBUG
    virtual char* GetOnMethodCallName() = 0;
#endif

protected:
    virtual void Passivate();
    void Signal();
    NV_DECLARE_ONCALL_METHOD(OnMethodCall, onmethodcall, (DWORD_PTR dwContext));

private:
    BOOL            _fSignalled;
    PFNDWNCHAN      _pfnCallback;
    void*           _pvCallback;
    THREADSTATE*    _pts;
};

class CDwnCtx : public CDwnChan
{
    typedef CDwnChan super;
    friend class CDwnInfo;

private:
    DECLARE_MEMCLEAR_NEW_DELETE()

public:
#ifdef _DEBUG
    void        EnterCriticalSection();
    void        LeaveCriticalSection();
    BOOL        EnteredCriticalSection();
#endif

    LPCTSTR     GetUrl();
    MIMEINFO*   GetMimeInfo();
    HRESULT     GetFile(LPTSTR* ppch);
    FILETIME    GetLastMod();
    DWORD       GetSecFlags();
    HRESULT     SetProgSink(IProgSink* pProgSink);
    ULONG       GetState(BOOL fClear=FALSE);
    void        SetLoad(BOOL fLoad, DWNLOADINFO* pdli, BOOL fReload);
    CDwnCtx*    GetDwnCtxNext() { return(_pDwnCtxNext); }
    CDwnLoad*   GetDwnLoad(); // Get AddRef'd CDwnLoad

protected:
    virtual void Passivate();
    CDwnInfo* GetDwnInfo() { return(_pDwnInfo); }
    void SetDwnInfo(CDwnInfo* pDwnInfo) { _pDwnInfo = pDwnInfo; }
    void Signal(WORD wChg);

    // Data members
    CDwnCtx*    _pDwnCtxNext;
    CDwnInfo*   _pDwnInfo;
    IProgSink*  _pProgSink;
    WORD        _wChg;
    WORD        _wChgReq;
    BOOL        _fLoad;
};


// flags for CopyOriginalSource
#define HTMSRC_FIXCRLF      0x1
#define HTMSRC_MULTIBYTE    0x2
#define HTMSRC_PRETRANSFORM 0x04

// flags for DrawImage
#define DRAWIMAGE_NHPALETTE 0x01    // DrawImage - set if being drawn from a non-halftone palette
#define DRAWIMAGE_NOTRANS   0x02    // DrawImage - set to ignore transparency
#define DRAWIMAGE_MASKONLY  0x04    // DrawImage - set to draw the transparency mask only


class CImgCtx : public CDwnCtx, public IImgCtx
{
    typedef CDwnCtx super;
    friend class CImgInfo;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    // IUnknown members
    STDMETHOD(QueryInterface)(REFIID, void**);
    STDMETHOD_(ULONG,AddRef)();
    STDMETHOD_(ULONG,Release)();

    // IImgCtx members
    STDMETHOD(Load)(LPCWSTR pszUrl, DWORD dwFlags);
    STDMETHOD(SelectChanges)(ULONG ulChgOn, ULONG ulChgOff, BOOL fSignal);
    STDMETHOD(SetCallback)(PFNIMGCTXCALLBACK pfnCallback, void* pvCallback);
    STDMETHOD(Disconnect)();
    STDMETHOD(GetUpdateRects)(RECT* prc, RECT* prectImg, LONG* pcrc);
    STDMETHOD(GetStateInfo)(ULONG* pulState, SIZE* psize, BOOL fClearChanges);
    STDMETHOD(GetPalette)(HPALETTE* phpal);
    STDMETHOD(Draw)(HDC hdc, LPRECT prcBounds);
    STDMETHOD(Tile)(HDC hdc, POINT* pptOrg, RECT* prcClip, SIZE* psizePrint);
    STDMETHOD(StretchBlt)(HDC hdc, int dstX, int dstY, int dstXE, int dstYE, int srcX, int srcY, int srcXE, int srcYE, DWORD dwROP);

    // CImgCtx members
    CImgCtx();
    ULONG   GetState() { return(super::GetState(FALSE)); }
    ULONG   GetState(BOOL fClear, SIZE* psize);
    HRESULT SaveAsBmp(IStream* pStm, BOOL fFileHeader);
    void    Tile(HDC hdc, POINT* pptOrg, RECT* prcClip, SIZE* psizePrint, COLORREF crBack, IMGANIMSTATE* pImgAnimState, DWORD dwFlags);
    BOOL    NextFrame(IMGANIMSTATE* pImgAnimState, DWORD dwCurTimeMS, DWORD* pdwFrameTimeMS);
    void    DrawFrame(HDC hdc, IMGANIMSTATE* pImgAnimState, RECT* prcDst, RECT* prcSrc, RECT* prcDestFull, DWORD dwFlags=0);
    void    InitImgAnimState(IMGANIMSTATE* pImgAnimState);
    DWORD_PTR GetImgId() { return (DWORD_PTR) GetImgInfo(); }
    HRESULT DrawEx(HDC hdc, LPRECT prcBounds, DWORD dwFlags); // allows dwFlags to be supplied

    CArtPlayer* GetArtPlayer();

#ifdef _DEBUG
    virtual char* GetOnMethodCallName() { return "CImgCtx::OnMethodCall"; }
#endif

protected:
    CImgInfo* GetImgInfo() { return((CImgInfo*)GetDwnInfo()); }

    void Init(CImgInfo* pImgInfo);
    void Signal(WORD wChg, BOOL fInvalAll, int yBot);
    void TileFast(HDC hdc, RECT* prc, LONG xSrcOrg, LONG ySrcOrg, BOOL fOpaque, COLORREF crBack, IMGANIMSTATE* pImgAnimState, DWORD dwFlags);
    void TileSlow(HDC hdc, RECT* prc, LONG xSrcOrg, LONG ySrcOrg, SIZE* psizePrint, BOOL fOpaque, COLORREF crBack, IMGANIMSTATE* pImgAnimState, DWORD dwFlags);

    // Data members
    LONG _yTop;
    LONG _yBot;
};


class CBitsCtx : public CDwnCtx
{
    typedef CDwnCtx super;
    friend class CBitsInfo;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    // CBitsCtx methods
    ULONG   GetState() { return(super::GetState(FALSE)); }
    void    SelectChanges(ULONG ulChgOn, ULONG ulChgOff, BOOL fSignal);
    HRESULT GetStream(IStream** ppStream);

#ifdef _DEBUG
    virtual char* GetOnMethodCallName() { return "CBitsCtx::OnMethodCall"; }
#endif
};


// CDwnDoc --------------------------------------------------------------------
class CDwnDoc : public CDwnChan, public IAuthenticate, public IWindowForBindingUI
{
    typedef CDwnChan super;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CDwnDoc();
    virtual ~CDwnDoc();

    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID iid, LPVOID* ppv);
    STDMETHOD_(ULONG,AddRef)()  { return(super::AddRef()); }
    STDMETHOD_(ULONG,Release)() { return(super::Release()); }

    // CDwnDoc methods
    void            SetBindf(DWORD dwBindf)         { _dwBindf = dwBindf; }
    void            SetDocBindf(DWORD dwBindf)      { _dwDocBindf = dwBindf; }
    void            SetLoadf(DWORD dwLoadf)         { _dwLoadf = dwLoadf; }
    void            SetDownf(DWORD dwDownf)         { _dwDownf = dwDownf; }
    void            SetRefresh(DWORD dwRefresh)     { _dwRefresh = dwRefresh; }
    void            SetDocCodePage(CODEPAGE cp)     { _cpDoc = cp; }
    void            SetURLCodePage(CODEPAGE cp)     { _cpURL = cp; }
    HRESULT         SetAcceptLanguage(LPCTSTR psz)  { RRETURN(SetString(&_cstrAcceptLang, psz)); }
    HRESULT         SetUserAgent(LPCTSTR psz)       { RRETURN(SetString(&_cstrUserAgent,  psz)); }
    HRESULT         SetAuthorColors(LPCTSTR pchColors, int cchColors=-1);
    void            AddBytesRead(DWORD dw)          { _dwRead += dw; }
    void            TakeRequestHeaders(BYTE** ppb, ULONG* pcb);
    BYTE*           GetRequestHeaders()             { return _pbRequestHeaders; }
    ULONG           GetRequestHeadersLength()       { return _cbRequestHeaders; }

    DWORD           GetBindf()                      { return(_dwBindf); }
    DWORD           GetDocBindf()                   { return(_dwDocBindf); }
    DWORD           GetLoadf()                      { return(_dwLoadf); }
    DWORD           GetDownf()                      { return(_dwDownf); }
    DWORD           GetRefresh()                    { return(_dwRefresh); }
    DWORD           GetDocCodePage()                { return(_cpDoc); }
    DWORD           GetURLCodePage()                { return(_cpURL); }
    LPCTSTR         GetAcceptLanguage()             { return(_cstrAcceptLang); }
    LPCTSTR         GetUserAgent()                  { return(_cstrUserAgent); }
    HRESULT         GetColors(CColorInfo* pCI);
    DWORD           GetBytesRead()                  { return(_dwRead); }
    CDocument*      GetCDoc()                       { return(_pDoc); }

    void            SetDownloadNotify(IDownloadNotify* pDownloadNotify);
    IDownloadNotify*GetDownloadNotify()             { return(_pDownloadNotify); }

    static HRESULT  Load(IStream* pstm, CDwnDoc** ppDwnDoc);
    static HRESULT  Save(CDwnDoc* pDwnDoc, IStream* pstm);
    static ULONG    GetSaveSize(CDwnDoc* pDwnDoc);

    void            SetDoc(CDocument* pDoc);
    void            Disconnect();

    BOOL            IsDocThread()                   { return(_dwThreadId == GetCurrentThreadId()); }
    HRESULT         AddDocThreadCallback(CDwnBindData* pDwnBindData, void* pvArg);
    HRESULT         QueryService(BOOL fBindOnApt, REFGUID rguid, REFIID riid, void** ppvObj);

    void            PreventAuthorPalette()          { _fGotAuthorPalette = TRUE; }
    BOOL            WantAuthorPalette()             { return !_fGotAuthorPalette; }
    BOOL            GotAuthorPalette()              { return _fGotAuthorPalette&&(_ape!=NULL); }

#ifdef _DEBUG
    virtual char*  GetOnMethodCallName()            { return("CDwnDoc::OnMethodCall"); }
#endif

    // IAuthenticate methods
    STDMETHOD(Authenticate)(HWND* phwnd, LPWSTR* ppszUsername, LPWSTR* ppszPassword);

    // IWindowForBindingUI methods
    STDMETHOD(GetWindow)(REFGUID rguidReason, HWND* phwnd);

protected:
    static HRESULT  SetString(CString* pcstr, LPCTSTR psz);
    void OnDocThreadCallback();
    static void CALLBACK OnDocThreadCallback(void* pvObj, void* pvArg)
    {
        ((CDwnDoc*)pvArg)->OnDocThreadCallback();
    }

    // Data members
    CDocument*          _pDoc;              // The document itself
    DWORD               _dwThreadId;        // Thread Id of the document
    DWORD               _dwBindf;           // See BINDF_* flags
    DWORD               _dwDocBindf;        // Bindf to be used for doc load only
    DWORD               _dwLoadf;           // See DLCTL_* flags
    DWORD               _dwDownf;           // See DWNF_* flags
    DWORD               _dwRefresh;         // Refresh level
    DWORD               _cpDoc;             // Codepage (Document)
    DWORD               _cpURL;             // Codepage (URL)
    DWORD               _dwRead;            // Number of bytes downloaded
    CString             _cstrAcceptLang;    // Accept language for headers
    CString             _cstrUserAgent;     // User agent for headers
    CAryDwnDocInfo      _aryDwnDocInfo;     // Bindings waiting for callback
    BOOL                _fCallbacks;        // DwnDoc is accepting callbacks
    BOOL                _fGotAuthorPalette; // Have we got the author defined palette (or is too late)
    UINT                _cpe;               // Number of author defined colors
    PALETTEENTRY*       _ape;               // Author defined palette
    ULONG               _cbRequestHeaders;  // Byte length of the following field
    BYTE*               _pbRequestHeaders;  // Raw request headers to use (ASCII)
    IDownloadNotify*    _pDownloadNotify;   // Free-threaded callback to notify of downloads (see dwndoc.cxx)
};

DEFINE_GUID(IID_IDwnBindInfo,  0x3050f3c3, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);

// CDwnBindInfo ---------------------------------------------------------------
class CDwnBindInfo :
    public CBaseFT,
    public IBindStatusCallback,
    public IServiceProvider,
    public IHttpNegotiate,
    public IInternetBindInfo
{
    typedef CBaseFT super;
    friend HRESULT CreateDwnBindInfo(IUnknown*, IUnknown**);

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CDwnBindInfo();

    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID iid, LPVOID* ppv);
    STDMETHOD_(ULONG,AddRef)();
    STDMETHOD_(ULONG,Release)();

    // IBindStatusCallback methods
    STDMETHOD(OnStartBinding)(DWORD dwReserved, IBinding* pib);
    STDMETHOD(GetPriority)(LONG* pnPriority);
    STDMETHOD(OnLowResource)(DWORD dwReserved);
    STDMETHOD(OnProgress)(ULONG ulProgress, ULONG ulProgressMax,  ULONG ulStatusCode,  LPCWSTR szStatusText);
    STDMETHOD(OnStopBinding)(HRESULT hresult, LPCWSTR szError);
    STDMETHOD(GetBindInfo)(DWORD* grfBINDF, BINDINFO* pbindinfo);
    STDMETHOD(OnDataAvailable)(DWORD grfBSCF, DWORD dwSize, FORMATETC* pformatetc, STGMEDIUM* pstgmed);
    STDMETHOD(OnObjectAvailable)(REFIID riid, IUnknown* punk);

    // IInternetBindInfo methods
    STDMETHOD(GetBindString)(ULONG ulStringType, LPOLESTR* ppwzStr, ULONG cEl, ULONG* pcElFetched);

    // IServiceProvider methods
    STDMETHOD(QueryService)(REFGUID rguidService, REFIID riid, void** ppvObj);

    // IHttpNegotiate methods

    STDMETHOD(BeginningTransaction)(LPCWSTR szURL, LPCWSTR szHeaders, DWORD dwReserved, LPWSTR* pszAdditionalHeaders);
    STDMETHOD(OnResponse)(DWORD dwResponseCode, LPCWSTR szResponseHeaders, LPCWSTR szRequestHeaders, LPWSTR* pszAdditionalRequestHeaders);

    // CDwnBindInfo methods
    void            SetDwnDoc(CDwnDoc* pDwnDoc);
    void            SetIsDocBind()          { _fIsDocBind = TRUE; }
    CDwnDoc*        GetDwnDoc()             { return(_pDwnDoc); }
    virtual UINT    GetScheme();
    LPCTSTR         GetContentType()        { return(_cstrContentType); }
    LPCTSTR         GetCacheFilename()      { return(_cstrCacheFilename); }
    virtual BOOL    IsBindOnApt()           { return(TRUE); }

#ifdef _DEBUG
    virtual BOOL    CheckThread()           { return(TRUE); }
#endif

protected:
    virtual ~CDwnBindInfo();
    static BOOL CanMarshalIID(REFIID riid);
    static HRESULT ValidateMarshalParams(REFIID riid, void* pvInterface,
        DWORD dwDestContext, void* pvDestContext, DWORD mshlflags);

    // Data members
    CDwnDoc*    _pDwnDoc;           // CDwnDoc for this binding
    BOOL        _fIsDocBind;        // Binding for document itself
    CString     _cstrContentType;   // Content type of this binding
    CString     _cstrCacheFilename; // cache file name for this file
};

#endif //__XINDOWS_SITE_DOWNLOAD_DOWNLOAD_H__