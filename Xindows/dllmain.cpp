// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"

extern HRESULT  InitBrushCache(THREADSTATE* pts);
extern void     DeinitBrushCache(THREADSTATE* pts);
extern HRESULT  InitBmpBrushCache(THREADSTATE* pts);
extern void     DeinitBmpBrushCache(THREADSTATE* pts);
extern HRESULT  InitGlobalWindow(THREADSTATE* pts);
extern void     DeinitGlobalWindow(THREADSTATE* pts);
extern void     DeinitTooltip(THREADSTATE* pts);
extern HRESULT  InitTaskManager(THREADSTATE* pts);
extern void     DeinitTaskManager(THREADSTATE* pts);
extern void     DeinitFormatCache(THREADSTATE* pts);

extern void     DeinitTearOffCache();
extern void     DeinitWindowClasses();
extern void     InitFormClipFormats();
extern void     InitColorTranslation();
extern void     ClearFaceCache();
extern HRESULT  InitPalette();
extern BOOL     InitImageUtil();
extern void     DeinitPalette();

extern void     InitFontCache();
extern void     DeinitTextSubSystem();
extern void     DeinitDownload();
extern void     DeinitImageSizeCache();

extern HRESULT  InitScrollbar(THREADSTATE* pts);
extern void     DeinitScrollbar(THREADSTATE* pts);
extern HRESULT  InitSystemMetricValues(THREADSTATE* pts);
extern void     DeinitSystemMetricValues(THREADSTATE* pts);
extern void     DeinitTimerCtx(THREADSTATE* pts);

extern void     OnSettingsChangeAllDocs(BOOL fNeedLayout);

extern void     DeinitImgAnim(THREADSTATE* pts);

extern void     DeinitFontLinking(THREADSTATE* pts);

extern HRESULT  InitLSCache(THREADSTATE* pts);
extern void     DeinitLSCache(THREADSTATE* pts);


// Thread attach/passivate/detach routines
typedef HRESULT (*PFN_HRESULT_INIT)(THREADSTATE* pts);
typedef void    (*PFN_VOID_DEINIT)(THREADSTATE* pts);

static const DWORD TLS_NULL = ((DWORD)-1);  // NULL TLS value (defined as 0xFFFFFFFF)

HRESULT DllThreadAttach();
void    DllThreadDetach(THREADSTATE* pts);
void    DllThreadPassivate(THREADSTATE* pts);
BOOL    DllProcessAttach();
void    DllProcessDetach();
void    DllThreadEmergencyDetach();



BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
    BOOL fOk = TRUE;

	switch(ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		{
            _afxGlobalData._dwTls = TLS_NULL;
            _afxGlobalData._pts = NULL;

			_afxGlobalData._hInstThisModule = hModule;
            _afxGlobalData._hInstResource = NULL;

            _afxGlobalData._lcidUserDefault = 0;
            _afxGlobalData._cpDefault = 0;
            _afxGlobalData._fFarEastWinNT = FALSE;
            _afxGlobalData._fBidiSupport = FALSE;
            _afxGlobalData._fComplexScriptInput = FALSE;

            _afxGlobalData._fHighContrastMode = FALSE;

            _afxGlobalData._fSystemFontsNeedRefreshing = TRUE;
            _afxGlobalData._fNoDisplayChange = TRUE;

            _afxGlobalData._sizePixelsPerInch.cx = 96;
            _afxGlobalData._sizePixelsPerInch.cy = 96;

            fOk = DllProcessAttach();

            CSelectTracker::InitMetrics();
            CGetBlockFmtCommand::Init();
            /*CComposeSettingsCommand::Init(); wlw note*/
		}
        break;
	case DLL_THREAD_DETACH:
        {
            DllThreadEmergencyDetach();
        }
        break;
	case DLL_PROCESS_DETACH:
        {
            DllProcessDetach();
        }
		break;
	}
	return fOk;
}



//+------------------------------------------------------------------
//
//  Function:   DllThreadAttach
//
//  Synopsis:   Prepare per-thread variables
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//-------------------------------------------------------------------
HRESULT DllThreadAttach()
{
    // Initialization routines (called in order)
    static const PFN_HRESULT_INIT s_apfnInit[] =
    {
        InitGlobalWindow,       // Must occur first
        InitSystemMetricValues,
        InitBrushCache,
        InitBmpBrushCache,
        InitScrollbar,
        InitTaskManager,
        InitLSCache
    };
    THREADSTATE*    pts;
    int             cInit;
    HRESULT         hr = S_OK;

    Trace1("DllThreadAttach called - TID 0x%08x\n", GetCurrentThreadId());

    // Allocate per-thread variables
    Assert(!TlsGetValue(_afxGlobalData._dwTls));
    pts = new THREADSTATE;
    if(!pts || !TlsSetValue(_afxGlobalData._dwTls, pts))
    {
        if(pts)
        {
            delete pts;
            pts = NULL;
        }
        hr = E_OUTOFMEMORY;
        goto Error;
    }

    MemSetName((pts, "THREADSTATE %08x", GetCurrentThreadId()));

    // Initialize the structure
    pts->dll.idThread = GetCurrentThreadId();

    // Run thread initialization routines
    for(cInit=0; cInit<ARRAYSIZE(s_apfnInit); cInit++)
    {
        hr = (*s_apfnInit[cInit])(pts);
        if(FAILED(hr))
        {
            goto Error;
        }
    }

Cleanup:
    // If successful, insert the per-thread state structure into the global chain
    if(!hr)
    {
        LOCK_GLOBALS;
        pts->ptsNext = _afxGlobalData._pts;
        _afxGlobalData._pts = pts;
    }
    Assert(SUCCEEDED(hr) && "Thread initialization failed");
    return hr;

Error:
    DllThreadDetach(pts);
    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Function:   DeinitOptionSettings
//
//  Synopsis:   Frees up memory stored in the TLS(optionSettingsInfo) struct.
//
//  Notes:      Called by the DllThreadDetach code.
//
//----------------------------------------------------------------------------
void DeinitOptionSettings(THREADSTATE* pts)
{
    int c, n;
    OPTIONSETTINGS** ppOS;
    CODEPAGESETTINGS** ppCS;

    // Free all entries in the options cache
    for(c=pts->optionSettingsInfo.pcache.Size(),
        ppOS=pts->optionSettingsInfo.pcache; c;
        c--,ppOS++)
    {
        (*ppOS)->cstrUserStylesheet.Free();

        // Free all entries in the codepage settings cache
        for(ppCS=(*ppOS)->aryCodepageSettingsCache,
            n=(*ppOS)->aryCodepageSettingsCache.Size(); n;
            n--,ppCS++)
        {
            MemFree(*ppCS);
        }

        delete (*ppOS);
    }

    pts->optionSettingsInfo.pcache.DeleteAll();
}

//+------------------------------------------------------------------
//
//  Function:   DllThreadDetach
//
//  Synopsis:   Release/clean-up per-thread variables
//
//  NOTE:   Since DllThreadDetach is called when DllThreadAttach fails
//          all Deinitxxxx routines must be robust. They must work correctly
//          even if the corresponding Initxxxx routine was not first called.
//
//-------------------------------------------------------------------
void DllThreadDetach(THREADSTATE* pts)
{
    // Deinitialization routines (called in order)
    static const PFN_VOID_DEINIT s_apfnDeinit[] =
    {
        DeinitOptionSettings,
        DeinitTaskManager,
        /*DeinitCommitHolder, wlw note*/
        DeinitTooltip,
        DeinitBrushCache,
        DeinitBmpBrushCache,
        DeinitScrollbar,
        DeinitTimerCtx,
        DeinitSystemMetricValues,
        DeinitFormatCache,
        /*DeinitContextMenus, wlw note*/
        DeinitImgAnim,
        DeinitFontLinking,
        DeinitLSCache,
        DeinitGlobalWindow      // Must occur last
    };
    THREADSTATE**   ppts;
    int             cDeinit;

    if(!pts)
    {
        return;
    }

    Assert(pts->dll.idThread == GetCurrentThreadId());

#ifdef _DEBUG
    Trace1("DllThreadDetach called - TID 0x%08x\n", GetCurrentThreadId());
    if(pts->dll.lObjCount)
    {
        Trace2("Thread (TID=0x%08x) terminated with primary object count=%d\n",
            GetCurrentThreadId(), pts->dll.lObjCount);
    }
#endif

    // Deinitialize the per-thread variables
    for(cDeinit=0; cDeinit<ARRAYSIZE(s_apfnDeinit); cDeinit++)
    {
        (*s_apfnDeinit[cDeinit])(pts);
    }

    ClearErrorInfo(pts);

    // Remove the per-thread structure from the global list
    LOCK_GLOBALS;
    for(ppts=&_afxGlobalData._pts; *ppts&&*ppts!=pts; ppts=&((*ppts)->ptsNext));
    if(*ppts)
    {
        *ppts = pts->ptsNext;
    }
#ifdef _DEBUG
    else
    {
        Trace1("THREADSTATE not on global chain - TID=0x%08x\n", GetCurrentThreadId());
    }
#endif

    // Disconnect the memory from the thread and delete it
    TlsSetValue(_afxGlobalData._dwTls, NULL);
    delete pts;

    return;
}

//+------------------------------------------------------------------
//
//  Function:   DllThreadPassivate
//
//  Synopsis:   This function is called when the per-thread object count,
//              dll.lObjCount, transitions to zero.  Deinit things here that
//              cannot be handled at process detach time.
//
//              This function can be called zero, one, or more times during
//              the time the DLL is loaded. Every function called from here
//              here should be prepared to handle this.
//
//-------------------------------------------------------------------
void DllThreadPassivate(THREADSTATE* pts)
{
    // Passivate per-thread objects
    // These include:
    //  a) Per-thread OLE error object (held by OLE automation)
    //  b) OLE clipboard object (only one thread should have anything on the clipboard)
    //  c) Per-thread ITypeInfo caches
    //  d) Per-thread picture helper
    //  e) Per-thread default IFont objects
    //
    // NOTE: This code contains one possible race condition: It is possible for
    //       the contents of the clipboard to be changed between the call to
    //       OleIsCurrentClipboard and OleFlushClipboard. If that occurs, this
    //       will flush the wrong object.
    //       GaryBu and I (BrendanD) discussed this case and felt it was not
    //       sufficiently important to warrent implementing a more complete
    //       (and costly) solution.
    Assert(pts);
    Trace1("DllThreadPassivate called - TID 0x%08x\n", GetCurrentThreadId());

    // Tell ole to clear its error info object if one exists.  If we
    // did not load OLEAUT32 then one could not exist.
    Verify(OK(SetErrorInfo(NULL, NULL)));

    if(pts->pDataClip && !OleIsCurrentClipboard(pts->pDataClip))
    {
        DEBUG_ONLY(HRESULT hr =) OleFlushClipboard();
#ifdef _DEBUG
        if(hr)
        {
            Assert(!pts->pDataClip && "Clipboard data should be flushed now");
        }
#endif
    }

    ClearInterface(&pts->pInetSess);
}

//+------------------------------------------------------------------
//
//  Function:   DllProcessAttach
//
//  Synopsis:   Initialize the DLL.
//
//  NOTE:       Even though DllMain is *not* called with DLL_THREAD_ATTACH
//              for the primary thread (since it is assumed during
//              DLL_PROCESS_ATTACH processing), this routines does not call
//              DllThreadAttach. Instead, all entry points are protected by a
//              call to EnsureThreadState which will, if necessary, call
//              DllThreadAttach. To call DllThreadAttach from here, might create
//              an unnecessary instance of the per-thread state structure.
//
//-------------------------------------------------------------------
BOOL DllProcessAttach()
{
    TCHAR* pch = NULL;

    // Allocate per-thread storage index
    _afxGlobalData._dwTls = TlsAlloc();
    if(_afxGlobalData._dwTls == TLS_NULL)
    {
        goto Error;
    }

    memset(&_afxGlobalData._Zero, 0, sizeof(_afxGlobalData._Zero));

    GetModuleFileName(_afxGlobalData._hInstThisModule, _afxGlobalData._szModulePath, MAX_PATH);
    lstrcpy(_afxGlobalData._szModuleFolder, _afxGlobalData._szModulePath);
    PathRemoveFileSpec(_afxGlobalData._szModuleFolder);
    lstrcat(_afxGlobalData._szModuleFolder, _T("\\"));

    TCHAR szResPath[MAX_PATH];
    lstrcpy(szResPath, _afxGlobalData._szModuleFolder);
    lstrcat(szResPath, _T("XindowsRes.dll"));
    _afxGlobalData._hInstResource = ::LoadLibrary(szResPath);

    const BOOL fFarEastLCID = IsFarEastLCID(GetSystemDefaultLCID());
    _afxGlobalData._fFarEastWinNT = fFarEastLCID;

    HKL aHKL[32];
    UINT uKeyboards = GetKeyboardLayoutList(32, aHKL);
    // check all keyboard layouts for existance of a RTL language.
    // bounce out when the first one is encountered.
    for(UINT i=0; i<uKeyboards; i++)
    {
        if(IsBidiLCID(LOWORD(aHKL[i])))
        {
            _afxGlobalData._fBidiSupport = TRUE;
            _afxGlobalData._fComplexScriptInput = TRUE;
            break;
        }

        if(IsComplexLCID(LOWORD(aHKL[i])))
        {
            _afxGlobalData._fComplexScriptInput = TRUE;
        }
    }

    // Register common clipboard formats used in Forms^3.  These are
    // available in the g_acfCommon array.
    RegisterClipFormats();

    // Now that the g_acfCommon are registered, we can initialize
    // the clip format array for the form and all the controls.
    InitFormClipFormats();

    // Initialize the global halftone palette
    if(FAILED(InitPalette()))
    {
        goto Error;
    }

    // Initialize the font cache. We do this here instead
    // of in the InitTextSubSystem() function because that
    // is very slow and we don't want to do it until
    // we know we're loading text. We need to do this here
    // because the registry loading code will initialize
    // the font face atom table.
    InitFontCache();

    return TRUE;

Error:
    DllProcessDetach();
    return FALSE;
}

//+------------------------------------------------------------------
//
//  Function:   DllProcessDetach
//
//  Synopsis:   Deinitialize the DLL.
//
//              This function can be called on partial initialization
//              of the DLL.  Every function called from here must be
//              capable being called without a corresponding call to
//              an initialization function.
//
//-------------------------------------------------------------------
void DllProcessDetach()
{
    if(_afxGlobalData._pts)
    {	
        // Our last, best hope for avoiding a crash it to try to take down the global
        // window on this thread and pray it was the last one.
        DllThreadEmergencyDetach();
    }

    DeinitWindowClasses();
    DeinitTextSubSystem();
    DeinitDownload();
    DeinitPalette();
    DeinitImageSizeCache();

    DeinitTearOffCache();

#ifdef _DEBUG
    if(_afxGlobalData._pts != NULL)
    {
        Trace0("One or more THREADSTATE's exist on DLL unload.");
    }
#endif //_DEBUG

    // Free per-thread storage index
    if(_afxGlobalData._dwTls != TLS_NULL)
    {
        TlsFree(_afxGlobalData._dwTls);
    }
}

void DllThreadEmergencyDetach()
{
    // This gets called by DLL_THREAD_DETACH.  We need to be careful to only do the
    // absolute minimum here to make sure we don't crash in the future.
    THREADSTATE* pts = (THREADSTATE*)TlsGetValue(_afxGlobalData._dwTls);

    if(pts && pts->dll.idThread==GetCurrentThreadId() && pts->gwnd.hwndGlobalWindow)
    {
        EnterCriticalSection(&pts->gwnd.cs);
        DestroyWindow(pts->gwnd.hwndGlobalWindow);
        pts->gwnd.hwndGlobalWindow = NULL;
        LeaveCriticalSection(&pts->gwnd.cs);
    }
}


//+---------------------------------------------------------------------------
//
//  Function:   DllUpdateSettings
//
//  Synopsis:   Updated cached system settings.  Called when the DLL is
//              attached and in response to WM_WININICHANGE, WM_DEVMODECHANGE,
//              and so on.
//
//----------------------------------------------------------------------------
void DllUpdateSettings(UINT msg)
{
    _afxGlobalData._fSystemFontsNeedRefreshing = TRUE;

    if(msg==WM_SYSCOLORCHANGE || msg==WM_DISPLAYCHANGE)
    {
        _afxGlobalData._fNoDisplayChange = 
            (_afxGlobalData._fNoDisplayChange && (msg!=WM_DISPLAYCHANGE));

        // On syscolor change, we need to update color table
        InitColorTranslation();
        InitPalette();
        InitImageUtil();
    }
    // When the fonts available have changed, we need to
    // recheck the system for font faces.
    else if(msg == WM_FONTCHANGE)
    {
        ClearFaceCache();
    }

    InitSystemMetricValues(GetThreadState());

    OnSettingsChangeAllDocs(WM_SYSCOLORCHANGE==msg||WM_FONTCHANGE==msg||(WM_USER+338)==msg);
}