
#include "stdafx.h"
#include "WindowsUtil.h"

static int  s_catomWndClass;
static ATOM s_aatomWndClass[32];

//+---------------------------------------------------------------------------
//
//  Function:   Register window class
//
//  Synopsis:   Register a window class.
//
//  Arguments:  pstrClass   String to use as part of class name.
//              pfnWndProc  The window procedure.
//              style       Window style flags.
//              pstrBase    Base class name, can be null.
//              ppfnBase    Base class window proc, can be null.
//              patom       Name of registered class.
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT RegisterWindowClass(
            TCHAR*      pstrClass,
            LRESULT     (CALLBACK *pfnWndProc)(HWND, UINT, WPARAM, LPARAM),
            UINT        style,
            TCHAR*      pstrBase,
            WNDPROC*    ppfnBase,
            ATOM*       patom,
            HICON       hIconSm /*=NULL*/)
{
    TCHAR       achClass[64];
    WNDCLASSEX  wc;

    LOCK_GLOBALS; // Guard access to globals (the passed atom and the atom array)

    // If another thread registered the class before this one obtained the global
    // lock, return immediately
    if(*patom)
    {
        RRETURN(S_OK);
    }

    Assert(*patom == 0);
    Assert(s_catomWndClass < ARRAYSIZE(s_aatomWndClass));

    Verify(OK(Format(_afxGlobalData._hInstResource, 0, achClass, ARRAYSIZE(achClass), _T("Xindows_<0s>"), pstrClass)));

    if(pstrBase)
    {
        if(!GetClassInfoEx(NULL, pstrBase, &wc))
        {
            goto Error;
        }

        *ppfnBase = wc.lpfnWndProc;
    }
    else
    {
        memset(&wc, 0, sizeof(wc));
    }

    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = (WNDPROC)pfnWndProc;
    wc.lpszClassName = achClass;
    wc.style |= style;
    wc.hInstance = _afxGlobalData._hInstThisModule;
    wc.hIconSm = hIconSm;

    *patom = RegisterClassEx(&wc);
    if(!*patom)
    {
        goto Error;
    }
    s_aatomWndClass[s_catomWndClass++] = *patom;

    return S_OK;

Error:
    RRETURN(GetLastWin32Error());
}

//+---------------------------------------------------------------------------
//
//  Function:   DeinitWindowClasses
//
//  Synopsis:   Unregister any window classes we have registered.
//
//----------------------------------------------------------------------------
void DeinitWindowClasses()
{
    while(--s_catomWndClass >= 0)
    {
        Verify(UnregisterClass((TCHAR*)(DWORD_PTR)s_aatomWndClass[s_catomWndClass], _afxGlobalData._hInstThisModule));
    }
}

//+------------------------------------------------------------------------
//
//  Function:   UpdateChildTree
//
//  Synopsis:   Calls UpdateWindow for a window, then recursively calls
//              UpdateWindow its children.
//
//  Arguments:  [hWnd]      Window to update, along with its children
//
//-------------------------------------------------------------------------
void UpdateChildTree(HWND hWnd)
{
    // The RedrawWindow call seems to miss the hWnd you actually pass in, or
    // else doesn't validate the update region after it has been redrawn, thus
    // the need for the UpdateWindow() call.
    if(hWnd)
    {
        UpdateWindow(hWnd);
        RedrawWindow(hWnd, NULL, NULL, RDW_UPDATENOW|RDW_ALLCHILDREN|RDW_INTERNALPAINT);
    }
}

//+------------------------------------------------------------------------
//
//  Function:   SetCursorIDC
//
//  Synopsis:   Set the cursor.  If the cursor id matches a standard
//              Windows cursor, then the cursor is loaded from the system.
//              Otherwise, the cursor is loaded from this instance.
//
//  Arguments:  idr - IDC_xxx from Windows.
//
//-------------------------------------------------------------------------
HCURSOR SetCursorIDC(LPCTSTR idr)
{
    HCURSOR hcursor, hcurRet;

    // No support for loading cursor by string id.
    Assert(IS_INTRESOURCE(idr));

    // We assume that if it's greater than IDC_ARROW,
    // then it's a system cursor.
    hcursor = LoadCursorA(
        ((DWORD_PTR)idr>=(DWORD_PTR)IDC_ARROW)?NULL:_afxGlobalData._hInstResource,
        (char*)idr);

    Assert(hcursor && "Failed to load cursor");

    hcurRet = GetCursor();
    if(hcursor && hcurRet!=hcursor)
    {
        // BUGBUG(sujalp): The windows SetCursor() call here has an *ugly* flash
        // in the incorrect position when the cursor's are changing. (Bug29467).
        // (This might be related to windows first showing the new cursor and then
        // setting its hotspot). To avoid this flash, we hide the cursor just
        // before changing the cursor and we enable it after changing the cursor.
        ShowCursor(FALSE);
        hcurRet = SetCursor(hcursor);
        ShowCursor(TRUE);
    }

    return hcurRet;
}

//+------------------------------------------------------------------------
//
//  Member:     CCurs::CCurs
//
//  Synopsis:   Constructor.  Loads the specified cursor and pushes it
//              on top of the cursor stack.  If the cursor id matches a
//              standard Windows cursor, then the cursor is loaded from
//              the system.  Otherwise, the cursor is loaded from this
//              instance.
//
//  Arguments:  idr - The resource id
//
//-------------------------------------------------------------------------
CCurs::CCurs(LPCTSTR idr)
{
    _hcrsOld = SetCursorIDC(idr);
    _hcrs    = GetCursor();
}

//+------------------------------------------------------------------------
//
//  Member:     CCurs::~CCurs
//
//  Synopsis:   Destructor.  Pops the cursor specified in the constructor
//              off the top of the cursor stack.  If the active cursor has
//              changed in the meantime, through some other mechanism, then
//              the old cursor is not restored.
//
//-------------------------------------------------------------------------
CCurs::~CCurs()
{
    if(GetCursor() == _hcrs)
    {
        ShowCursor(FALSE);
        SetCursor(_hcrsOld);
        ShowCursor(TRUE);
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   VBShiftState
//
//  Synopsis:   Helper function, returns shift state for KeyDown/KeyUp events
//
//  Arguments:  (none)
//
//  Returns:    USHORT
//
//  Notes:      This function maps the keystate supplied by Windows to
//              1, 2 and 4 (which are numbers from VB)
//
//----------------------------------------------------------------------------
short VBShiftState()
{
    short sKeyState = 0;

    if(GetKeyState(VK_SHIFT) & 0x8000)
    {
        sKeyState |= VB_SHIFT;
    }

    if(GetKeyState(VK_CONTROL) & 0x8000)
    {
        sKeyState |= VB_CONTROL;
    }

    if(GetKeyState(VK_MENU) & 0x8000)
    {
        sKeyState |= VB_ALT;
    }

    return sKeyState;
}

short VBShiftState(DWORD grfKeyState)
{
    short sKeyState = 0;

    if(grfKeyState & MK_SHIFT)
    {
        sKeyState |= VB_SHIFT;
    }

    if(grfKeyState & MK_CONTROL)
    {
        sKeyState |= VB_CONTROL;
    }

    if(grfKeyState & MK_ALT)
    {
        sKeyState |= VB_ALT;
    }

    return sKeyState;
}

//+---------------------------------------------------------------------------
//
//  Function:   VBButtonState
//
//  Synopsis:   Helper function, returns button state for Mouse events
//
//  Arguments:  w -- word containing mouse ButtonState
//
//  Returns:    short
//
//  Notes:      This function maps the buttonstate supplied by Windows to
//              1, 2 and 4 (which are numbers from VB)
//
//----------------------------------------------------------------------------
short VBButtonState(WPARAM w)
{
    short sButtonState = 0;

    if(w & MK_LBUTTON)
    {
        sButtonState |= VB_LBUTTON;
    }
    if(w & MK_RBUTTON)
    {
        sButtonState |= VB_RBUTTON;
    }
    if(w & MK_MBUTTON)
    {
        sButtonState |= VB_MBUTTON;
    }

    return sButtonState;
}
