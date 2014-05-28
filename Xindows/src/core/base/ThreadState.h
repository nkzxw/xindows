
#ifndef __XINDOWS_CORE_BASE_THREADSTATE_H__
#define __XINDOWS_CORE_BASE_THREADSTATE_H__

class CCharFormat;
class CParaFormat;
class CFancyFormat;
class CAttrArray;
class CScrollbarController;
class CLSCache;
class CDocument;
class CAryTimers;
class CAryCalls;
class CTask;
class CErrorInfo;
class CAryBBCR;
class CTooltip;
class CTimerCtx;
class CImgAnim;
struct BCR;
struct OPTIONSETTINGS;

template<class T> class CDataCache;

typedef CDataCache<CCharFormat>     CCharFormatCache;
typedef CDataCache<CParaFormat>     CParaFormatCache;
typedef CDataCache<CFancyFormat>    CFancyFormatCache;
typedef CDataCache<CAttrArray>      CStyleExpandoCache;


class CTlsDocAry : public CPtrAry<CDocument*>
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();
    CTlsDocAry() : CPtrAry<CDocument*>() {}
};

class CTlsOptionAry : public CPtrAry<OPTIONSETTINGS*>
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();
    CTlsOptionAry() : CPtrAry<OPTIONSETTINGS*>() {}
};



//+----------------------------------------------------------------------------
//
//  Struct:     THREADSTATE
//
//  Synopsis:   Contains all per-thread variables
//
//              Add per-thread variables to the structure beneath the chain
//              pointers. When referencing external classes or structs, prepend
//              "class" or "struct" to the variable rather than including the
//              necessary header file. Lastly, if the variable requires an
//              initial state other than zero, modify DllThreadAttach and
//              DllThreadDetach to call the appropriate Init/Deinit routines.
//
//              Related variables should be grouped together (though not all have
//              been done this way yet) into a struct. This makes managing the
//              namespace easier and keeps clear what is used by what.
//
//  NOTE:       The lifetime of these structures is tied to that of the contained
//              hwndGlobalWindow. That is, the per-thread instance is destroyed
//              when that window is closed (sent WM_CLOSE). Do not delete instances
//              of this structure directly unless hwndGlobalWindow is NULL.
//
//              Use the TLS macro to access these variables, e.g., foo = TLS(bar);
//
//-----------------------------------------------------------------------------
struct THREADSTATE
{
    DECLARE_MEMCLEAR_NEW_DELETE()

    THREADSTATE*            ptsNext;            // THREADSTATE chain pointer

    struct DLL
    {
        DWORD               idThread;           // ID of owning thread
        LONG                lObjCount;          // Per-thread count of (external) references
    } dll;                                      // DLL management related variables

    struct GWND
    {
        HWND                hwndGlobalWindow;   // "Global" window (global per-thread)
        UINT                uID;                // Timer ID counter
        CAryTimers*         paryTimers;         // Dynamic array of timers
        BOOL                fMethodCallPosted;  // Flag to indicate method PostMessage was called
        CAryCalls*          paryCalls;          // Dynamic array of method calls
        void*               pvCapture;          // Object which owns mouse capture
        PFN_VOID_ONMESSAGE  pfnOnMouseMessage;  // Message handler on object which owns mouse
        HWND                hwndCapture;        // Window which owns mouse
        HHOOK               hhookGWMouse;       // Active mouse hook
        CRITICAL_SECTION    cs;                 // Thread synchronization
    } gwnd;                                     // Global window related variables

    CCharFormatCache*       _pCharFormatCache;
    CParaFormatCache*       _pParaFormatCache;
    CFancyFormatCache*      _pFancyFormatCache;
    CStyleExpandoCache*     _pStyleExpandoCache;
    LONG                    _ipfDefault;
    const CParaFormat*      _ppfDefault;
    LONG                    _iffDefault;
    const CFancyFormat*     _pffDefault;

    struct TASK
    {
        CTask*              ptaskHead;          // Linked list of tasks
        CTask*              ptaskNext;          // Next task to run
        CTask*              ptaskNextInLoop;    // Next task to run in loop
        CTask*              ptaskCur;           // Currently running task
        DWORD               dwTickRun;          // Timeout interval for running task
        DWORD               cUnblocked;         // Count of unblocked tasks
        DWORD               cInterval;          // Count of interval tasks
        DWORD               dwTickTimer;        // Last timer interval set
        BOOL                fSuspended;         // Suspended due to re-entrancy
    } task;

    IDataObject*            pDataClip;          // Pointer to transfer object
    HDC                     hdcDesktop;         // DC compatible with the desktop DC
    CErrorInfo*             pEI;                // Per-thread IErrorInfo object
    BCR*                    pbcr;               // Brush cache
    int                     ibcrNext;           // Rotating index to pick brush to discard
    CAryBBCR*               paryBBCR;           // Bitmap brush cache array
    CTooltip*               pTT;                // Current tooltip object
    CScrollbarController*   pSBC;               // Scrollbar controller
    IInternetSession*       pInetSess;          // Cached internet session
    CTimerCtx*              pTimerCtx;          // CTimerMan lacky for this thread

    CImgAnim*               pImgAnim;           // ImgAnim for this thread

    struct SCROLLTIMEINFO
    {
        LONG    lRepeatDelay;
        LONG    lRepeatRate;
        LONG    lFocusRate;
        LONG    iTargetScrolls;                 // Number of scroll operations during 
                                                // the last smooth scrolling.
        DWORD   dwTargetTime;                   // Total time the scroll operation took
        DWORD   dwTimeEndLastScroll;            // DEBUG ONLY: need to time in the ship build as well
    } scrollTimeInfo;

    CTlsDocAry _paryDoc;                        // Thread Docs will add and remove themselves from the array

    struct OPTIONSETTINGSINFO
    {
        CTlsOptionAry pcache;                   // Registry settings
    } optionSettingsInfo;

    CLSCache*   _pLSCache;                      // the lineservices cache

    long        nUndoState;                     // enum for normal, undo, redo (enum in undo.hxx)
};

#endif //__XINDOWS_CORE_BASE_THREADSTATE_H__