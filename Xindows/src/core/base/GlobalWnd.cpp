
#include "stdafx.h"

#define WM_METHODCALL   (WM_APP+2)

struct CALLENTRY
{
    void*           pvObject;
    PFN_VOID_ONCALL pfnOnCall;
    DWORD_PTR       dwContext;
#ifdef _DEBUG
    char*           pszOnCall;
#endif
};

DECLARE_CDataAry(CAryCalls, CALLENTRY);

static ATOM s_atomGlobalWndClass = NULL;

// Mouse capture
LRESULT CALLBACK GWMouseProc(int nCode, WPARAM wParam, LPARAM lParam);

extern void DllUpdateSettings(UINT msg);

#ifdef _DEBUG
char* DecodeMessage(UINT msg)
{
    switch(msg)
    {
    case WM_NULL:               return("WM_NULL");
    case WM_CREATE:             return("WM_CREATE");
    case WM_DESTROY:            return("WM_DESTROY");
    case WM_MOVE:               return("WM_MOVE");
    case WM_SIZE:               return("WM_SIZE");
    case WM_ACTIVATE:           return("WM_ACTIVATE");
    case WM_SETFOCUS:           return("WM_SETFOCUS");
    case WM_KILLFOCUS:          return("WM_KILLFOCUS");
    case WM_ENABLE:             return("WM_ENABLE");
    case WM_SETREDRAW:          return("WM_SETREDRAW");
    case WM_SETTEXT:            return("WM_SETTEXT");
    case WM_GETTEXT:            return("WM_GETTEXT");
    case WM_GETTEXTLENGTH:      return("WM_GETTEXTLENGTH");
    case WM_PAINT:              return("WM_PAINT");
    case WM_CLOSE:              return("WM_CLOSE");
    case WM_QUERYENDSESSION:    return("WM_QUERYENDSESSION");
    case WM_QUERYOPEN:          return("WM_QUERYOPEN");
    case WM_ENDSESSION:         return("WM_ENDSESSION");
    case WM_QUIT:               return("WM_QUIT");
    case WM_ERASEBKGND:         return("WM_ERASEBKGND");
    case WM_SYSCOLORCHANGE:     return("WM_SYSCOLORCHANGE");
    case WM_SHOWWINDOW:         return("WM_SHOWWINDOW");
    case WM_WININICHANGE:       return("WM_WININICHANGE");
    case WM_DEVMODECHANGE:      return("WM_DEVMODECHANGE");
    case WM_ACTIVATEAPP:        return("WM_ACTIVATEAPP");
    case WM_FONTCHANGE:         return("WM_FONTCHANGE");
    case WM_TIMECHANGE:         return("WM_TIMECHANGE");
    case WM_CANCELMODE:         return("WM_CANCELMODE");
    case WM_SETCURSOR:          return("WM_SETCURSOR");
    case WM_MOUSEACTIVATE:      return("WM_MOUSEACTIVATE");
    case WM_CHILDACTIVATE:      return("WM_CHILDACTIVATE");
    case WM_QUEUESYNC:          return("WM_QUEUESYNC");
    case WM_GETMINMAXINFO:      return("WM_GETMINMAXINFO");
    case WM_PAINTICON:          return("WM_PAINTICON");
    case WM_ICONERASEBKGND:     return("WM_ICONERASEBKGND");
    case WM_NEXTDLGCTL:         return("WM_NEXTDLGCTL");
    case WM_SPOOLERSTATUS:      return("WM_SPOOLERSTATUS");
    case WM_DRAWITEM:           return("WM_DRAWITEM");
    case WM_MEASUREITEM:        return("WM_MEASUREITEM");
    case WM_DELETEITEM:         return("WM_DELETEITEM");
    case WM_VKEYTOITEM:         return("WM_VKEYTOITEM");
    case WM_CHARTOITEM:         return("WM_CHARTOITEM");
    case WM_SETFONT:            return("WM_SETFONT");
    case WM_GETFONT:            return("WM_GETFONT");
    case WM_SETHOTKEY:          return("WM_SETHOTKEY");
    case WM_GETHOTKEY:          return("WM_GETHOTKEY");
    case WM_QUERYDRAGICON:      return("WM_QUERYDRAGICON");
    case WM_COMPAREITEM:        return("WM_COMPAREITEM");
    case WM_COMPACTING:         return("WM_COMPACTING");
    case WM_COMMNOTIFY:         return("WM_COMMNOTIFY");
    case WM_WINDOWPOSCHANGING:  return("WM_WINDOWPOSCHANGING");
    case WM_WINDOWPOSCHANGED:   return("WM_WINDOWPOSCHANGED");
    case WM_POWER:              return("WM_POWER");
    case WM_COPYDATA:           return("WM_COPYDATA");
    case WM_CANCELJOURNAL:      return("WM_CANCELJOURNAL");
    case WM_NOTIFY:             return("WM_NOTIFY");
    case WM_INPUTLANGCHANGEREQUEST: return("WM_INPUTLANGCHANGEREQUEST");
    case WM_INPUTLANGCHANGE:    return("WM_INPUTLANGCHANGE");
    case WM_TCARD:              return("WM_TCARD");
    case WM_HELP:               return("WM_HELP");
    case WM_USERCHANGED:        return("WM_USERCHANGED");
    case WM_NOTIFYFORMAT:       return("WM_NOTIFYFORMAT");
    case WM_CONTEXTMENU:        return("WM_CONTEXTMENU");
    case WM_STYLECHANGING:      return("WM_STYLECHANGING");
    case WM_STYLECHANGED:       return("WM_STYLECHANGED");
    case WM_DISPLAYCHANGE:      return("WM_DISPLAYCHANGE");
    case WM_GETICON:            return("WM_GETICON");
    case WM_SETICON:            return("WM_SETICON");
    case WM_NCCREATE:           return("WM_NCCREATE");
    case WM_NCDESTROY:          return("WM_NCDESTROY");
    case WM_NCCALCSIZE:         return("WM_NCCALCSIZE");
    case WM_NCHITTEST:          return("WM_NCHITTEST");
    case WM_NCPAINT:            return("WM_NCPAINT");
    case WM_NCACTIVATE:         return("WM_NCACTIVATE");
    case WM_GETDLGCODE:         return("WM_GETDLGCODE");
    case WM_SYNCPAINT:          return("WM_SYNCPAINT");
    case WM_NCMOUSEMOVE:        return("WM_NCMOUSEMOVE");
    case WM_NCLBUTTONDOWN:      return("WM_NCLBUTTONDOWN");
    case WM_NCLBUTTONUP:        return("WM_NCLBUTTONUP");
    case WM_NCLBUTTONDBLCLK:    return("WM_NCLBUTTONDBLCLK");
    case WM_NCRBUTTONDOWN:      return("WM_NCRBUTTONDOWN");
    case WM_NCRBUTTONUP:        return("WM_NCRBUTTONUP");
    case WM_NCRBUTTONDBLCLK:    return("WM_NCRBUTTONDBLCLK");
    case WM_NCMBUTTONDOWN:      return("WM_NCMBUTTONDOWN");
    case WM_NCMBUTTONUP:        return("WM_NCMBUTTONUP");
    case WM_NCMBUTTONDBLCLK:    return("WM_NCMBUTTONDBLCLK");
    case WM_KEYDOWN:            return("WM_KEYDOWN");
    case WM_KEYUP:              return("WM_KEYUP");
    case WM_CHAR:               return("WM_CHAR");
    case WM_DEADCHAR:           return("WM_DEADCHAR");
    case WM_SYSKEYDOWN:         return("WM_SYSKEYDOWN");
    case WM_SYSKEYUP:           return("WM_SYSKEYUP");
    case WM_SYSCHAR:            return("WM_SYSCHAR");
    case WM_SYSDEADCHAR:        return("WM_SYSDEADCHAR");
    case WM_IME_STARTCOMPOSITION:   return("WM_IME_STARTCOMPOSITION");
    case WM_IME_ENDCOMPOSITION: return("WM_IME_ENDCOMPOSITION");
    case WM_IME_COMPOSITION:    return("WM_IME_COMPOSITION");
    case WM_INITDIALOG:         return("WM_INITDIALOG");
    case WM_COMMAND:            return("WM_COMMAND");
    case WM_SYSCOMMAND:         return("WM_SYSCOMMAND");
    case WM_TIMER:              return("WM_TIMER");
    case WM_HSCROLL:            return("WM_HSCROLL");
    case WM_VSCROLL:            return("WM_VSCROLL");
    case WM_INITMENU:           return("WM_INITMENU");
    case WM_INITMENUPOPUP:      return("WM_INITMENUPOPUP");
    case WM_MENUSELECT:         return("WM_MENUSELECT");
    case WM_MENUCHAR:           return("WM_MENUCHAR");
    case WM_ENTERIDLE:          return("WM_ENTERIDLE");
    case WM_CTLCOLORMSGBOX:     return("WM_CTLCOLORMSGBOX");
    case WM_CTLCOLOREDIT:       return("WM_CTLCOLOREDIT");
    case WM_CTLCOLORLISTBOX:    return("WM_CTLCOLORLISTBOX");
    case WM_CTLCOLORBTN:        return("WM_CTLCOLORBTN");
    case WM_CTLCOLORDLG:        return("WM_CTLCOLORDLG");
    case WM_CTLCOLORSCROLLBAR:  return("WM_CTLCOLORSCROLLBAR");
    case WM_CTLCOLORSTATIC:     return("WM_CTLCOLORSTATIC");
    case WM_MOUSEMOVE:          return("WM_MOUSEMOVE");
    case WM_LBUTTONDOWN:        return("WM_LBUTTONDOWN");
    case WM_LBUTTONUP:          return("WM_LBUTTONUP");
    case WM_LBUTTONDBLCLK:      return("WM_LBUTTONDBLCLK");
    case WM_RBUTTONDOWN:        return("WM_RBUTTONDOWN");
    case WM_RBUTTONUP:          return("WM_RBUTTONUP");
    case WM_RBUTTONDBLCLK:      return("WM_RBUTTONDBLCLK");
    case WM_MBUTTONDOWN:        return("WM_MBUTTONDOWN");
    case WM_MBUTTONUP:          return("WM_MBUTTONUP");
    case WM_MBUTTONDBLCLK:      return("WM_MBUTTONDBLCLK");
    case WM_MOUSEWHEEL:         return("WM_MOUSEWHEEL");
    case WM_PARENTNOTIFY:       return("WM_PARENTNOTIFY");
    case WM_ENTERMENULOOP:      return("WM_ENTERMENULOOP");
    case WM_EXITMENULOOP:       return("WM_EXITMENULOOP");
    case WM_NEXTMENU:           return("WM_NEXTMENU");
    case WM_SIZING:             return("WM_SIZING");
    case WM_CAPTURECHANGED:     return("WM_CAPTURECHANGED");
    case WM_MOVING:             return("WM_MOVING");
    case WM_POWERBROADCAST:     return("WM_POWERBROADCAST");
    case WM_DEVICECHANGE:       return("WM_DEVICECHANGE");
    case WM_MDICREATE:          return("WM_MDICREATE");
    case WM_MDIDESTROY:         return("WM_MDIDESTROY");
    case WM_MDIACTIVATE:        return("WM_MDIACTIVATE");
    case WM_MDIRESTORE:         return("WM_MDIRESTORE");
    case WM_MDINEXT:            return("WM_MDINEXT");
    case WM_MDIMAXIMIZE:        return("WM_MDIMAXIMIZE");
    case WM_MDITILE:            return("WM_MDITILE");
    case WM_MDICASCADE:         return("WM_MDICASCADE");
    case WM_MDIICONARRANGE:     return("WM_MDIICONARRANGE");
    case WM_MDIGETACTIVE:       return("WM_MDIGETACTIVE");
    case WM_MDISETMENU:         return("WM_MDISETMENU");
    case WM_ENTERSIZEMOVE:      return("WM_ENTERSIZEMOVE");
    case WM_EXITSIZEMOVE:       return("WM_EXITSIZEMOVE");
    case WM_DROPFILES:          return("WM_DROPFILES");
    case WM_MDIREFRESHMENU:     return("WM_MDIREFRESHMENU");
    case WM_IME_SETCONTEXT:     return("WM_IME_SETCONTEXT");
    case WM_IME_NOTIFY:         return("WM_IME_NOTIFY");
    case WM_IME_CONTROL:        return("WM_IME_CONTROL");
    case WM_IME_COMPOSITIONFULL:    return("WM_IME_COMPOSITIONFULL");
    case WM_IME_SELECT:         return("WM_IME_SELECT");
    case WM_IME_CHAR:           return("WM_IME_CHAR");
    case WM_IME_KEYDOWN:        return("WM_IME_KEYDOWN");
    case WM_IME_KEYUP:          return("WM_IME_KEYUP");
    case WM_MOUSEHOVER:         return("WM_MOUSEHOVER");
    case WM_MOUSELEAVE:         return("WM_MOUSELEAVE");
    case WM_CUT:                return("WM_CUT");
    case WM_COPY:               return("WM_COPY");
    case WM_PASTE:              return("WM_PASTE");
    case WM_CLEAR:              return("WM_CLEAR");
    case WM_UNDO:               return("WM_UNDO");
    case WM_RENDERFORMAT:       return("WM_RENDERFORMAT");
    case WM_RENDERALLFORMATS:   return("WM_RENDERALLFORMATS");
    case WM_DESTROYCLIPBOARD:   return("WM_DESTROYCLIPBOARD");
    case WM_DRAWCLIPBOARD:      return("WM_DRAWCLIPBOARD");
    case WM_PAINTCLIPBOARD:     return("WM_PAINTCLIPBOARD");
    case WM_VSCROLLCLIPBOARD:   return("WM_VSCROLLCLIPBOARD");
    case WM_SIZECLIPBOARD:      return("WM_SIZECLIPBOARD");
    case WM_ASKCBFORMATNAME:    return("WM_ASKCBFORMATNAME");
    case WM_CHANGECBCHAIN:      return("WM_CHANGECBCHAIN");
    case WM_HSCROLLCLIPBOARD:   return("WM_HSCROLLCLIPBOARD");
    case WM_QUERYNEWPALETTE:    return("WM_QUERYNEWPALETTE");
    case WM_PALETTEISCHANGING:  return("WM_PALETTEISCHANGING");
    case WM_PALETTECHANGED:     return("WM_PALETTECHANGED");
    case WM_HOTKEY:             return("WM_HOTKEY");
    case WM_PRINT:              return("WM_PRINT");
    case WM_PRINTCLIENT:        return("WM_PRINTCLIENT");
    case WM_USER:               return("WM_USER");
    case WM_USER+1:             return("WM_USER+1");
    case WM_USER+2:             return("WM_USER+2");
    case WM_USER+3:             return("WM_USER+3");
    case WM_USER+4:             return("WM_USER+4");
    }

    return("");
}

char* DecodeWindowClass(HWND hwnd)
{
    static char ach[40];
    ach[0] = 0;
    GetClassNameA(hwnd, ach, sizeof(ach));
    return ach;
}
#endif

//+-------------------------------------------------------------------------
//
//  Function:   FormsTrackPopupMenu
//
//  Synopsis:   Allows windowless controls to use the global window for
//              command routing of popup menu selection.
//
//--------------------------------------------------------------------------
HRESULT FormsTrackPopupMenu(HMENU hMenu, UINT fuFlags, int x, int y, HWND hwndMessage, int* piSelection)
{
    BOOL    fAvail;
    HRESULT hr = S_OK;
    MSG     msg;
    HWND    hwnd;

    hwnd = (hwndMessage) ? (hwndMessage) : (TLS(gwnd.hwndGlobalWindow));

    if(::TrackPopupMenu(hMenu, fuFlags, x, y, 0, hwnd, (RECT*)NULL))
    {
        // The menu popped up and the item was chosen.  Peek messages
        // until the specified command was found.
        fAvail = PeekMessage(&msg, hwnd, WM_COMMAND, WM_COMMAND, PM_REMOVE);

        if(fAvail)
        {
            *piSelection = GET_WM_COMMAND_ID(msg.wParam, msg.lParam);
            hr = S_OK;
        }
        else
        {
            // No WM_COMMAND was available, so this means that the
            // menu was brought down
            hr = S_FALSE;
        }
    }
    else
    {
        hr = GetLastWin32Error();
    }

    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Function:   _GWPostMethodCallEx
//
//  Synopsis:   Call given method on given object a later time.
//              It is the caller's responsiblity to insure that the
//              object remains valid until the call is made or the
//              posted call is killed.
//
//  Arguments:  pvObject    The object
//              pfnOnCall   Method to call.
//              dwContext   Context for method
//
//--------------------------------------------------------------------------
HRESULT _GWPostMethodCallEx(
        THREADSTATE*    pts,
        void*           pvObject,
        PFN_VOID_ONCALL pfnOnCall,
        DWORD_PTR       dwContext,
        BOOL            fIgnoreContext
#ifdef _DEBUG
        , char*         pszOnCall
#endif // _DEBUG
        )
{
    HRESULT     hr;
    CALLENTRY*  pce;
    int         c;

    EnterCriticalSection(&pts->gwnd.cs);

    for(c=(*(pts->gwnd.paryCalls)).Size(),pce=(*(pts->gwnd.paryCalls)); c>0; c--,pce++)
    {
        if(pce->pvObject==pvObject && pce->pfnOnCall==pfnOnCall &&
            pce->dwContext==dwContext && !fIgnoreContext)
        {
            hr = S_OK;
            goto Cleanup;
        }
    }

    c = pts->gwnd.paryCalls->Size();

    hr = (*(pts->gwnd.paryCalls)).EnsureSize(c+1);
    if(hr)
    {
        goto Cleanup;
    }

    pce = &(*(pts->gwnd.paryCalls))[c];

    pce->pvObject = pvObject;
    pce->pfnOnCall = pfnOnCall;
    pce->dwContext = dwContext;

#ifdef _DEBUG
    pce->pszOnCall = pszOnCall;
#endif

    if(!pts->gwnd.fMethodCallPosted)
    {
        if(!PostMessage(pts->gwnd.hwndGlobalWindow, WM_METHODCALL, 0, 0))
        {
            hr = GetLastWin32Error();
            goto Cleanup;
        }

        pts->gwnd.fMethodCallPosted = TRUE;
    }

    (*(pts->gwnd.paryCalls)).SetSize(c+1);

Cleanup:
    LeaveCriticalSection(&pts->gwnd.cs);

    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Function:   GWKillMethodCallEx
//
//  Synopsis:   Kill method call posted with GWPostMethodCall.
//
//  Arguments:  pvObject    The object
//              pfnOnCall   Method to call.  If null, kills all calls
//                          for pvObject.
//              dwContext   Context.  If zero, kills all calls for pvObject
//                          and pfnOnCall.
//
//--------------------------------------------------------------------------
void GWKillMethodCallEx(THREADSTATE* pts, void* pvObject, PFN_VOID_ONCALL pfnOnCall, DWORD_PTR dwContext)
{
    CALLENTRY*  pce;
    int         c;

    // Handle pts being NULL for
    Assert(pts);

    // check for no calls before entering critical section
    if(!pts || (pts->gwnd.paryCalls->Size()==0))
    {
        return;
    }

    EnterCriticalSection(&pts->gwnd.cs);

    // Null pointer instead of deleting array element to
    // handle calls to this function when OnMethodCall
    // is iterating over the array.
    for(c=(*(pts->gwnd.paryCalls)).Size(),pce=(*(pts->gwnd.paryCalls)); c>0; c--,pce++)
    {
        if(pce->pvObject == pvObject)
        {
            if(pfnOnCall == NULL)
            {
                pce->pvObject = NULL;
            }
            else if(pce->pfnOnCall == pfnOnCall)
            {
                if(pce->dwContext==dwContext || dwContext==0)
                {
                    pce->pvObject = NULL;
                }
            }
        }
    }

    LeaveCriticalSection(&pts->gwnd.cs);
}

void GWKillMethodCall(void* pvObject, PFN_VOID_ONCALL pfnOnCall, DWORD_PTR dwContext)
{
    GWKillMethodCallEx(GetThreadState(), pvObject, pfnOnCall, dwContext);
}

#ifdef _DEBUG
BOOL GWHasPostedMethod(void* pvObject)
{
    THREADSTATE*    pts = (THREADSTATE*)(TlsGetValue(_afxGlobalData._dwTls));
    CALLENTRY*      pce;
    int             c;
    BOOL            fRet = FALSE;

    if(!pts)
    {
        return FALSE;
    }

    EnterCriticalSection(&pts->gwnd.cs);

    // Null pointer instead of deleting array element to
    // handle calls to this function when OnMethodCall
    // is iterating over the array.
    for(c=(*(pts->gwnd.paryCalls)).Size(),pce=(*(pts->gwnd.paryCalls));
        c>0;
        c--,pce++)
    {
        if(pce->pvObject == pvObject)
        {
            fRet = TRUE;
            break;
        }
    }

    LeaveCriticalSection(&pts->gwnd.cs);

    return fRet;
}
#endif // _DEBUG

//+-------------------------------------------------------------------------
//
//  Funciton:   GlobalWndOnMethodCall
//
//  Synopsis:   Handles deferred method calls.
//
//--------------------------------------------------------------------------
void GlobalWndOnMethodCall()
{
    THREADSTATE*    pts = GetThreadState();
    CALLENTRY*      pce;
    CALLENTRY*      pceLive;
    int             i;
    int             c;

    // Re-enable posting of layout messages because the function call
    // could cause us to enter a modal loop and then OnMethodCall would
    // not process further messages.  This will happen if the function
    // call fires an event procedure that calls Userform.Show in VB.
    EnterCriticalSection(&pts->gwnd.cs);

    pts->gwnd.fMethodCallPosted = FALSE;

    for(i=0; i<(*(pts->gwnd.paryCalls)).Size(); i++)
    {
        // Pointer into array is fetched at every iteration in order
        // to handle calls to GWPostMethodCall during the loop.

        pce = &(*(pts->gwnd.paryCalls))[i];
        if(pce->pvObject)
        {
            CALLENTRY ce = *pce;
            pce->pvObject = NULL;

            LeaveCriticalSection(&pts->gwnd.cs);

            CALL_METHOD((CVoid*)ce.pvObject, ce.pfnOnCall, (ce.dwContext));

            EnterCriticalSection(&pts->gwnd.cs);

            // DO NOT USE pce after the method call - if the object calls
            // GWPostMethodCall it may cause gwnd.paryCalls to do a ReAlloc,
            // which invalidates pce!
        }
    }

    pceLive = (*(pts->gwnd.paryCalls));
    i = 0;
    for(c=(*(pts->gwnd.paryCalls)).Size(),pce=(*(pts->gwnd.paryCalls));
        c>0;
        c--,pce++)
    {
        if(pce->pvObject)
        {
            *pceLive++ = *pce;
            i++;
        }
    }

    (*(pts->gwnd.paryCalls)).SetSize(i);

    LeaveCriticalSection(&pts->gwnd.cs);
}

//+-------------------------------------------------------------------------
//
//  Function:   _GWSetCapture
//
//  Synopsis:   Capture the mouse.
//
//--------------------------------------------------------------------------
HRESULT GWSetCapture(void* pvObject, PFN_VOID_ONMESSAGE pfnOnMouseMessage, HWND hwnd)
{
    THREADSTATE* pts;

    Assert(pvObject);

    pts = GetThreadState();

    if(pvObject != pts->gwnd.pvCapture)
    {
        if(pts->gwnd.pvCapture)
        {
            CALL_METHOD((CVoid*)pts->gwnd.pvCapture, pts->gwnd.pfnOnMouseMessage, (WM_CAPTURECHANGED, 0, 0));
        }
        pts->gwnd.pvCapture = pvObject;
        pts->gwnd.pfnOnMouseMessage = pfnOnMouseMessage;
        pts->gwnd.hwndCapture = hwnd;
    }

    if(GetCapture() != TLS(gwnd.hwndGlobalWindow))
    {
        SetCapture(TLS(gwnd.hwndGlobalWindow));

        if(!pts->gwnd.hhookGWMouse)
        {
            pts->gwnd.hhookGWMouse = SetWindowsHookEx(
                WH_MOUSE,
                GWMouseProc,
                (HINSTANCE)NULL,
                GetCurrentThreadId());
        }
    }

    return S_OK;
}

//+-------------------------------------------------------------------------
//
//  Function:   _GWReleaseCapture
//
//  Synopsis:   Release the mouse capture.
//
//--------------------------------------------------------------------------
void GWReleaseCapture(void* pvObject)
{
    THREADSTATE* pts;

    Assert(pvObject);

    pts = GetThreadState();

    if(pts->gwnd.pvCapture == pvObject)
    {
        CVoid* pvCapture = (CVoid*)pts->gwnd.pvCapture;
        pts->gwnd.pvCapture = NULL;
        CALL_METHOD(pvCapture, pts->gwnd.pfnOnMouseMessage, (WM_CAPTURECHANGED, 0, 0));

        if(pts->gwnd.hhookGWMouse)
        {
            UnhookWindowsHookEx(pts->gwnd.hhookGWMouse);
            pts->gwnd.hhookGWMouse = NULL;
        }

        if(GetCapture() == TLS(gwnd.hwndGlobalWindow))
        {
            ReleaseCapture();
        }
    }
}

//+-------------------------------------------------------------------------
//
//  Function:   _GWGetCapture
//
//  Synopsis:   Return the object that has captured the mouse.
//
//--------------------------------------------------------------------------
BOOL GWGetCapture(void* pvObject)
{
    if(GetCapture() == TLS(gwnd.hwndGlobalWindow))
    {
        return pvObject==TLS(gwnd.pvCapture);
    }
    return FALSE;
}

//+-------------------------------------------------------------------------
//
//  Function:   _GWMouseProc
//
//  Synopsis:   Mouse proc for the global window.  This mouse hook is installed
//              under WinNT when the global window has mouse capture and removed
//              when it releases capture.  Under windows 95, the global window
//              receives WM_CAPTURECHANGED messages and this hook isn't necessary.
//              Under WinNT this mouse proc simulates the WM_CAPTURECHANGED.
//
//  BUGBUG - This global mouse proc should be modified  so that WM_CAPTURECHANGED
//           is sent to all forms3 windows when mouse capture changes.
//--------------------------------------------------------------------------
LRESULT CALLBACK GWMouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if(nCode < 0) /* do not process the message */
    {
        return CallNextHookEx(TLS(gwnd.hhookGWMouse), nCode, wParam, lParam);
    }

    if(nCode == HC_ACTION)
    {
        MOUSEHOOKSTRUCT* pmh = (MOUSEHOOKSTRUCT*)lParam;

        if(wParam>=WM_MOUSEFIRST && wParam<=WM_MOUSELAST &&
            TLS(gwnd.pvCapture) && pmh->hwnd!=TLS(gwnd.hwndGlobalWindow))
        {
            GWReleaseCapture(TLS(gwnd.pvCapture));
        }
    }

    return CallNextHookEx(TLS(gwnd.hhookGWMouse), nCode, wParam, lParam);
}

//+-------------------------------------------------------------------------
//
//  Method:     OnTimer
//
//  Synopsis:   Handles timer event from the global window.
//
//--------------------------------------------------------------------------
static void CALLBACK OnTimer(HWND hwnd, UINT id)
{
    TIMERENTRY*     pte;
    int             c;
    THREADSTATE*    pts = GetThreadState();

    for(c=pts->gwnd.paryTimers->Size(),pte=*pts->gwnd.paryTimers; c>0; c--,pte++)
    {
        if(pte->idTimerAlias == id)
        {
            CALL_METHOD((CVoid*)pte->pvObject, pte->pfnOnTick, (pte->idTimer));
            break;
        }
    }
}

//+-------------------------------------------------------------------------
//
//  Function:   GlobalWndProc
//
//  Synopsis:   Window proc for the global window
//
//--------------------------------------------------------------------------
LRESULT CALLBACK GlobalWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    THREADSTATE*    pts;
    LRESULT         lRet = 0;
    LONG            lDllCount;
    BOOL            fCallDefWndProc = FALSE;

    // (paulnel) 74803 - turn off mirroring if the system supports it. This causes all
    //           sorts of problems in Trident.
    if(msg == WM_NCCREATE)
    {
        DWORD dwExStyles;
        if((dwExStyles=GetWindowLong(hwnd, GWL_EXSTYLE)) & WS_EX_LAYOUTRTL)
        {
            SetWindowLong(hwnd, GWL_EXSTYLE, dwExStyles&~WS_EX_LAYOUTRTL);
        }
    }

    pts = (THREADSTATE*)TlsGetValue(_afxGlobalData._dwTls);
    lDllCount = pts ? pts->dll.lObjCount : 0;

    // We need to guard against a Release() call made during
    // message processing causing the ref count to go to zero 
    // and the TLS getting blown away.  This will be done
    // by manually adjusting the counter (to be fast) and 
    // only call DecrementObjectCount if necessary.
    if(lDllCount)
    {
        Verify(++pts->dll.lObjCount > 0);
    }

    // Note: When adding new messages to this loop do the 
    //       following:
    //  - Handle the message
    //  - set lRet to be the lResult to return (if not 0)
    //  - set fCallDefWndProc to TRUE if DefWindowProc() should
    //    be called (default is to NOT call DefWindowProc())
    //  - use 'break' to exit the switch statement
    switch(msg)
    {
    case WM_TIMER:
        OnTimer(hwnd, (UINT)wParam);
        break;

    case WM_METHODCALL:
        GlobalWndOnMethodCall();
        break;    

    case WM_MOUSEMOVE:
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_LBUTTONDBLCLK:
    case WM_RBUTTONDOWN:
    case WM_RBUTTONUP:
    case WM_RBUTTONDBLCLK:
    case WM_MBUTTONDOWN:
    case WM_MBUTTONUP:
    case WM_MBUTTONDBLCLK:
        if(pts && pts->gwnd.pvCapture)
        {
            POINT  pt;
            pt.x = MAKEPOINTS(lParam).x;
            pt.y = MAKEPOINTS(lParam).y;
            ScreenToClient(pts->gwnd.hwndCapture, &pt);
            lRet = CALL_METHOD(
                (CVoid*)pts->gwnd.pvCapture,
                pts->gwnd.pfnOnMouseMessage,
                (msg,
                wParam,
                MAKELONG(pt.x, pt.y)));
        }
        else
        {
            fCallDefWndProc = TRUE;
        }
        break;

    case WM_CAPTURECHANGED:
        if(pts && pts->gwnd.pvCapture)
        {
            CVoid* pvCapture = (CVoid*)pts->gwnd.pvCapture;
            pts->gwnd.pvCapture = 0;
            lRet = CALL_METHOD(
                pvCapture,
                pts->gwnd.pfnOnMouseMessage,
                (msg,
                wParam,
                lParam));
        }
        else
        {
            fCallDefWndProc = TRUE;
        }
        break;

        // case WM_WININICHANGE: obsolete, should not be used anymore
        // replaced with WM_SETTINGCHANGE (which has the same ID, but uses wParam)
    case WM_SETTINGCHANGE:
    case WM_SYSCOLORCHANGE:
    case WM_DEVMODECHANGE:
    case WM_FONTCHANGE:
    case WM_DISPLAYCHANGE:
    case WM_USER+338: // sent by properties change dialog
        Trace1("Processing system change %d\n", msg);
        DllUpdateSettings(msg);
        break;

    default:
        fCallDefWndProc = TRUE;
        goto Cleanup;
    }

Cleanup:
    if(lDllCount && pts->dll.lObjCount==1)
    {
        // TLS about to go away.  Let the Passivates occur
        // and then say we handled the message.  Since
        // DecrementObjectCount plays with the secondary
        // count as well we need to increment it as well.
        DecrementObjectCount();
        lRet = 0;
    }
    else
    {
        if(lDllCount)
        {
            Verify(--pts->dll.lObjCount >= 0);
        }
        if(fCallDefWndProc)
        {
            lRet = DefWindowProc(hwnd, msg, wParam, lParam);
        }
    }

    return lRet;
}

//+-------------------------------------------------------------------------
//
//  Function:   DeinitGlobalWindow
//
//--------------------------------------------------------------------------
void DeinitGlobalWindow(THREADSTATE* pts)
{
    Assert(pts);

#ifdef _DEBUG
    if(pts->gwnd.paryTimers)
    {
        Assert((*(pts->gwnd.paryTimers)).Size() == 0);
    }
    if(pts->gwnd.paryCalls)
    {
        for(int i=0; i<(*(pts->gwnd.paryCalls)).Size(); i++)
        {
            Assert((*(pts->gwnd.paryCalls))[i].pvObject == NULL);
        }
    }
#endif

    if(pts->gwnd.paryCalls)
    {
        (*(pts->gwnd.paryCalls)).DeleteAll();
    }

    // Delete per-thread dynamic arrays
    delete pts->gwnd.paryTimers;
    delete pts->gwnd.paryCalls;

    if(pts->gwnd.hwndGlobalWindow)
    {
        if(pts->dll.idThread == GetCurrentThreadId())
        {
            Verify(DestroyWindow(pts->gwnd.hwndGlobalWindow));
        }

        // BUGBUG: if we're on the wrong thread we can't destroy the window
        pts->gwnd.hwndGlobalWindow = NULL;

        DeleteCriticalSection(&pts->gwnd.cs);
    }
}

//+-------------------------------------------------------------------------
//
//  Function:   InitGlobalWindow
//
//  Synopsis:   Initializes and creates the global hwnd.
//
//--------------------------------------------------------------------------
HRESULT InitGlobalWindow(THREADSTATE* pts)
{
    HRESULT hr = S_OK;
    TCHAR* pszBuf;

    Assert(pts);
    
    // Create the per-thread "global" window
    if(!s_atomGlobalWndClass)
    {
        hr = RegisterWindowClass(
            _T("Hidden"),
            GlobalWndProc,
            0,
            NULL,
            NULL,
            &s_atomGlobalWndClass);
        if(hr)
        {
            goto Error;
        }
    }

    pszBuf = (TCHAR*)(DWORD_PTR)s_atomGlobalWndClass;

    pts->gwnd.hwndGlobalWindow = CreateWindow(
        pszBuf,
        NULL,
        WS_POPUP,
        0, 0, 0, 0,
        NULL,
        NULL,
        _afxGlobalData._hInstThisModule,
        NULL);
    if(pts->gwnd.hwndGlobalWindow == NULL)
    {
        hr = GetLastWin32Error();
        goto Error;
    }

    InitializeCriticalSection(&pts->gwnd.cs);

    // Allocate per-thread dynamic arrays
    pts->gwnd.paryTimers = new CAryTimers;
    pts->gwnd.paryCalls = new CAryCalls;
    if(!pts->gwnd.paryTimers || !pts->gwnd.paryCalls)
    {
        hr = E_OUTOFMEMORY;
        goto Error;
    }

Error:
    RRETURN(hr);
}
