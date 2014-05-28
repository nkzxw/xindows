
#ifndef __XINDOWS_SITE_DOWNLOAD_IMAGE_H__
#define __XINDOWS_SITE_DOWNLOAD_IMAGE_H__

#include "Dwn.h"

// Forward --------------------------------------------------------------------
class   CImgInfo;
class   CImgTask;
class   CImgTaskExec;
class   CImgLoad;
class   CArtPlayer;
struct  GIFFRAME;
struct  GIFANIMDATA;

// Definitions ----------------------------------------------------------------
enum
{
    gifNoneSpecified    = 0,    // no disposal method specified
    gifNoDispose        = 1,    // do not dispose, leave the bits there
    gifRestoreBkgnd     = 2,    // replace the image with the background color
    gifRestorePrev      = 3     // replace the image with the previous pixels
};

#define TRANSF_TRANSPARENT  0x01        // Image is marked transparent
#define TRANSF_TRANSMASK    0x02        // Attempted to create an hbmMask

#define dwGIFVerUnknown     ((DWORD)0)  // unknown version of GIF file
#define dwGIFVer87a         ((DWORD)87) // GIF87a file format
#define dwGIFVer89a         ((DWORD)89) // GIF89a file format.

// Globals --------------------------------------------------------------------
extern WORD         g_wIdxBgColor;
extern WORD         g_wIdxFgColor;
extern WORD         g_wIdxTrans;
extern RGBQUAD      g_rgbBgColor;
extern RGBQUAD      g_rgbFgColor;
extern PALETTEENTRY g_peVga[16];
extern BYTE*        g_pInvCMAP;
extern CImgBits*    g_pImgBitsNotLoaded;
extern CImgBits*    g_pImgBitsMissing;

// Types ----------------------------------------------------------------------
struct GIFFRAME
{
    GIFFRAME *      pgfNext;
    CImgBitsDIB*    pibd;
    HRGN            hrgnVis;            // region describing currently visible portion of the frame
    BYTE            bDisposalMethod;    // see enum above
    BYTE            bTransFlags;        // see TRANSF_ flags below
    BYTE            bRgnKind;           // region type for hrgnVis
    UINT            uiDelayTime;        // frame duration, in ms
    long            left;               // bounds relative to the GIF logical screen 
    long            top; 
    long            width;
    long            height;
};

struct GIFANIMDATA
{
    BYTE        fAnimated;          // TRUE if cFrames and pgf define a GIF animation
    BYTE        fLooped;            // TRUE if we've seen a Netscape loop block
    BYTE        fHasTransparency;   // TRUE if a frame is transparent, or if a frame does
                                    // not cover the entire logical screen.
    BYTE        bAlign;             // Reserved
    UINT        cLoops;             // A la Netscape, we will treat this as 
                                    // "loop forever" if it is zero.
    DWORD       dwGIFVer;           // GIF Version <see defines above> we need to special case 87a backgrounds
    GIFFRAME*   pgf;                // animation frame entries
};

struct FIBERINFO
{
    void*       pvFiber;
    void*       pvMain;
    CImgTask*   pImgTask;
};

// Functions ------------------------------------------------------------------
void*   pCreateDitherData(int xsize);
int     x_ComputeConstrainMap(
                  int cEntries,
                  PALETTEENTRY* pcolors,
                  int transparent,
                  int* pmapconstrained);
void    x_ColorConstrain(
                 unsigned char HUGEP* psrc,
                 unsigned char HUGEP* pdst,
                 int* pmapconstrained,
                 long xsize);
void    x_DitherRelative(
                 BYTE* pbSrc,
                 BYTE* pbDst,
                 PALETTEENTRY* pe,
                 int xsize,
                 int ysize,
                 int transparent,
                 int* v_rgb_mem,
                 int yfirst,
                 int ylast);
HRESULT x_Dither(
                 unsigned char* pdata,
                 PALETTEENTRY* pe,
                 int xsize,
                 int ysize,
                 int transparent);

void    CalcStretchRect(
                    RECT* prectStretch,
                    LONG xWid,
                    LONG yHei,
                    LONG xDispWidth,
                    LONG yDispHeight,
                    GIFFRAME* pgf);
void    getPassInfo(
                    int logicalRowX,
                    int height,
                    int* pPassX,
                    int* pRowX,
                    int* pBandX);
int     Union(int _yTop, int _yBot, BOOL fInvalidateAll, int yBot);
void    ComputeFrameVisibility(
                   IMGANIMSTATE* pImgAnimState,
                   LONG xWidth,
                   LONG yHeight,
                   LONG xDispWidth,
                   LONG yDispHeight);
void    FreeGifAnimData(GIFANIMDATA* pgad, CImgBitsDIB* pImgBits);
HBITMAP ComputeTransMask(HBITMAP hbmDib, BOOL fPal, BYTE bTrans);
HRESULT StartImgTask(CImgTask* pImgTask);
void    KillImgTaskExec();
BOOL    IsPluginImgFormat(BYTE* pb, UINT cb);

// CImgInfo -------------------------------------------------------------------
class CImgInfo : public CDwnInfo
{
    typedef CDwnInfo super;
    friend class CImgCtx;
    friend class CImgLoad;
    friend class CImgTask;
    friend class CImgTaskExec;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    virtual        ~CImgInfo();
    virtual HRESULT Init(DWNLOADINFO* pdli);
    virtual void    Passivate();

    HRESULT         DrawImage(HDC hdc, RECT* prcDst, RECT* prcSrc, DWORD dwROP, DWORD dwFlags);
    void            InitImgAnimState(IMGANIMSTATE * pImgAnimState);
    BOOL            NextFrame(IMGANIMSTATE* pImgAnimState, DWORD dwCurTimeMS, DWORD* pdwFrameTimeMS);
    void            DrawFrame(HDC hdc, IMGANIMSTATE* pImgAnimState, RECT* prcDst, RECT* prcSrc, RECT* prcDestFull, DWORD dwFlags);
    int             GetColorMode() { return ((int)GetFlags(DWNF_COLORMODE)); }
    HRESULT         SaveAsBmp(IStream* pStm, BOOL fFileHeader);
    virtual HRESULT NewDwnCtx(CDwnCtx** ppDwnCtx);
    virtual HRESULT NewDwnLoad(CDwnLoad** ppDwnLoad);
    CArtPlayer*     GetArtPlayer();

    void            GetImageAndMask(CImgBits** ppImgBits);
    LONG            GetTrans() { return _lTrans; }

protected:
    UINT            GetType() { return(DWNCTX_IMG); }
    DWORD           GetProgSinkClass() { return PROGSINK_CLASS_MULTIMEDIA; }
    void            Signal(WORD wChg, BOOL fInvalAll, int yBot);
    void            OnLoadTask(CImgLoad* pImgLoad, CImgTask* pImgTask);
    virtual void    OnLoadDone(HRESULT hrErr);
    void            OnTaskSize(CImgTask* pImgTask, LONG xWid, LONG yHei, LONG lTrans, MIMEINFO* pmi);
    void            OnTaskTrans(CImgTask* pImgTask, LONG lTrans);
    void            OnTaskProg(CImgTask* pImgTask, ULONG ulState, BOOL fAll, LONG yBot);
    void            OnTaskAnim(CImgTask* pImgTask);
    BOOL            OnTaskBits(CImgTask* pImgTask, CImgBits* pImgBits, GIFANIMDATA* pgad, CArtPlayer* pArtPlayer, LONG lTrans, LONG ySrcBot, BOOL fNonProgressive);
    void            OnTaskDone(CImgTask* pImgTask);
    virtual void    Abort(HRESULT hrErr, CDwnLoad** ppDwnLoad);
    virtual void    Reset();

    virtual BOOL    AttachEarly(UINT dt, DWORD dwRefresh, DWORD dwFlags, DWORD dwBindf);
    virtual BOOL    CanAttachLate(CDwnInfo* pDwnInfo);
    virtual void    AttachLate(CDwnInfo* pDwnInfo);
    virtual ULONG   ComputeCacheSize();

    // Data members
    CImgTask*       _pImgTask;
    LONG            _xWid;
    LONG            _yHei;
    LONG            _ySrcBot;

    CImgBits*      _pImgBits;

    LONG            _lTrans;
    GIFANIMDATA     _gad;

    unsigned        _fNoOptimize; // not optimize if imginfo created by external component
    CArtPlayer*     _pArtPlayer;
};


// CImgTaskExec ---------------------------------------------------------------
class CImgTaskExec : public CDwnTaskExec
{
    typedef CDwnTaskExec super;
    friend class CImgTask;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CImgTaskExec(CRITICAL_SECTION* pcs) : super(pcs) {}
    HRESULT             RequestCoInit();

protected:
    virtual HRESULT     ThreadInit();
    virtual void        ThreadTerm();
    virtual void        ThreadExit();
    virtual void        ThreadTimeout();

    FIBERINFO*          GetFiber(CImgTask* pImgTask);
    void                YieldTask(CImgTask* pImgTask, BOOL fBlock);
    void                RunTask(CImgTask* pImgTask);
    void                AssignFiber(FIBERINFO* pfi);
    static void WINAPI  FiberProc(void* pv);

    // Data members
    void*       _pvFiberMain;
    FIBERINFO   _afi[5];
    BOOL        _fCoInit;
};

// CImgTask -------------------------------------------------------------------
class CImgTask : public CDwnTask
{
    typedef CDwnTask super;
    friend class CImgTaskExec;

private:
    DECLARE_MEMCLEAR_NEW_DELETE()

public:
    void            Init(CImgInfo* pImgInfo, MIMEINFO* pmi, CDwnBindData* pDwnBindData);
    void            Exec();
    BOOL            Read(void* pv, ULONG cb, ULONG* pcbRead=NULL, ULONG cbMinReq=0);
    virtual void    Terminate();

    void            OnSize(LONG xWid, LONG yHei, LONG lTrans);
    void            OnTrans(LONG lTrans);
    void            OnProg(BOOL fLast, ULONG ulBits, BOOL fAll, LONG yBot);
    void            OnAnim();
    LPCTSTR         GetUrl();
    GIFFRAME*       GetPgf() { return _gad.pgf; }

    virtual void    BltDib(HDC hdc, RECT* prcDst, RECT* prcSrc, DWORD dwRop, DWORD dwFlags) {}

    BOOL            IsFullyAvail() { return _pDwnBindData->IsFullyAvail(); }

    void            GetImageAndMask(CImgBits** ppImgBits);

    CArtPlayer*     GetArtPlayer() { return _pArtPlayer; }
    BOOL            DoTaskGetReport(CArtPlayer* pArtPlayer);

protected:
    ~CImgTask();
    virtual void    Run();
    virtual void    Decode(BOOL* pfNonProgressive) = 0;
    CImgTaskExec*   GetImgTaskExec() { return ((CImgTaskExec*)_pDwnTaskExec); }

    // Data members
    CImgInfo*       _pImgInfo;
    CDwnBindData*   _pDwnBindData;
    FIBERINFO*      _pfi;
    BOOL            _fTerminate;
    BOOL            _fWaitForFiber;
    LONG            _xWid;
    LONG            _yHei;
    LONG            _ySrcBot;
    LONG            _lTrans;

    CImgBits*       _pImgBits;

    GIFANIMDATA     _gad;
    CArtPlayer*    _pArtPlayer;
    PALETTEENTRY    _ape[256];
    LONG            _yBot;
    BOOL            _fComplete;
    LONG            _colorMode;
    MIMEINFO*      _pmi;
    DWORD           _dwTickProg;
    LONG            _yTopProg;
    LONG            _yBotProg;
};

// CImgLoad ----------------------------------------------------------------
class CImgLoad : public CDwnLoad
{
    typedef CDwnLoad super;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    // CDwnLoad methods
    virtual        ~CImgLoad();
    virtual void    Passivate();
    virtual HRESULT Init(DWNLOADINFO* pdli, CDwnInfo* pDwnInfo);
    virtual HRESULT OnBindHeaders();
    virtual HRESULT OnBindMime(MIMEINFO* pmi);
    virtual HRESULT OnBindData();
    virtual void    OnBindDone(HRESULT hrErr);

    // CImgLoad methods
    CImgInfo*       GetImgInfo()    { return (CImgInfo*)_pDwnInfo; }
    int             GetColorMode()  { return GetImgInfo()->GetColorMode(); }

protected:
    // Data members
    CImgTask* _pImgTask;
};

// Inlines --------------------------------------------------------------------
inline void CImgTask::OnSize(LONG xWid, LONG yHei, LONG lTrans)
{
    _pImgInfo->OnTaskSize(this, xWid, yHei, lTrans, _pmi);
}

inline void CImgTask::OnTrans(LONG lTrans)
{
    _pImgInfo->OnTaskTrans(this, lTrans);
}

inline LPCTSTR CImgTask::GetUrl()
{
    return _pImgInfo->GetUrl();
}

inline void CImgInfo::GetImageAndMask(CImgBits** ppImgBits)
{
    EnterCriticalSection();

    if(_pImgTask)
    {
        _pImgTask->GetImageAndMask(ppImgBits);
    }
    else
    {
        *ppImgBits = _pImgBits;
    }

    LeaveCriticalSection();
}

inline void CImgTask::GetImageAndMask(CImgBits** ppImgBits)
{
    EnterCriticalSection();

    *ppImgBits = _pImgBits;

    LeaveCriticalSection();
}

// lookup closest entry in g_hpalHalftone to (r,g,b)
inline BYTE RGB2Index(BYTE r, BYTE g, BYTE b)
{ 
    return g_pInvCMAP[((((r>>3)<<5)+(g>>3))<<5)+(b>>3)];
}

STDAPI DecodeImage(IStream* pStream, IMapMIMEToCLSID* pMap, IUnknown* pEventSink);
STDAPI IdentifyMIMEType(const BYTE* pbBytes, ULONG nBytes, UINT* pnFormat);
STDAPI CreateDDrawSurfaceOnDIB(HBITMAP hbmDib, IDirectDrawSurface** ppSurface);


#endif //__XINDOWS_SITE_DOWNLOAD_IMAGE_H__