
#include "stdafx.h"
#include "Tooltip.h"

//+-------------------------------------------------------------------------
//
//  Function:   EnsureInit
//
//  Synopsis:   Static helper for initializing the tooltip code.
//
//--------------------------------------------------------------------------
static HRESULT EnsureInit(void)
{
    THREADSTATE* pts;
    HRESULT hr = S_OK;

    pts = GetThreadState();

    if(pts->pTT == NULL)
    {
        pts->pTT = new CTooltip;
        if(!pts->pTT)
        {
            Assert(0 && "Failed to create Tooltip object.");
            hr = E_OUTOFMEMORY;
        }
    }

    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     FormsShowTooltip
//
//  Synopsis:   Displays the Tooltip.
//
//  Arguments:  szText      Text to display in the Tooltip
//              hwnd        Owner hwnd
//              msg         Message passed to tooltip for precessing
//              prc         Coords of rectangle in which the mouse
//                          should reside to display the Tooltip.  If the
//                          mouse moves out of this rect, the Tooltip will
//                          be brought down.
//              dwCookie    Any value that is used to identify a tooltip
//                          source.  If subsequent calls to FormsShowTooltip
//                          have the same dwCookie, the tooltip isn't
//                          redisplayed (avoiding flashing).  This value
//                          could be a pointer to the source's this ptr,
//                          an index, etc.
//              fRTL        COMPLEXSCRIPT flag indicating if the element
//                          lays out weak/neutral characters right to left
//
//--------------------------------------------------------------------------
void FormsShowTooltip(TCHAR* szText, HWND hwnd, MSG msg, RECT* prc, DWORD_PTR dwCookie, BOOL fRTL)
{
    HRESULT hr;

    hr = EnsureInit();
    if(hr)
    {
        return;
    }

    TLS(pTT)->Show(szText, hwnd, msg,  prc, dwCookie, fRTL);
}

//+-------------------------------------------------------------------------
//
//  Method:     FormsHideTooltip
//
//  Synopsis:   Forces the Tooltip to be brought down.
//
//  Arguments:  [fReset]    If true, cause the tooltip code to reset
//                          the cookie value.  Use FALSE, for example, when
//                          you dismiss the tooltip due to a click.  In this
//                          case you wouldn't want the tooltip to reappear
//                          if the mouse is still in the region.
//
//--------------------------------------------------------------------------
void FormsHideTooltip(BOOL fReset)
{
    if(TLS(pTT))
    {
        TLS(pTT)->Hide();
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     FormsTooltipMessage
//
//  Synopsis:   Allows the tooltip code to dismiss the tooltip based on
//              a window message.  Centralizing the code here keeps clients
//              from having to remember the dismissal rules themselves.
//
//  Returns:    TRUE if the message type is one that would dismiss the
//              tooltip.
//
//--------------------------------------------------------------------------
BOOL FormsTooltipMessage(UINT msg, WPARAM wParam, LPARAM lParam)
{
    BOOL fHide = FALSE;

    switch(msg)
    {
    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_MBUTTONDOWN:
    case WM_KEYUP:
    case WM_KEYDOWN:
    case WM_CONTEXTMENU:
    case WM_MOUSEWHEEL:
        fHide = TRUE;
        break;

    case WM_SETCURSOR:
        switch(HIWORD(lParam))
        {
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_MBUTTONDOWN:
            fHide = TRUE;
            break;

        default:
            break;
        }

    default:
        break;
    }

    if(fHide)
    {
        FormsHideTooltip(FALSE);
    }

    return fHide;
}

//+-------------------------------------------------------------------------
//
//  Method:     CTooltip::CTooltip
//
//  Synopsis:   Constructor
//
//--------------------------------------------------------------------------
CTooltip::CTooltip()
{
    Trace0("CTooltip::CTooltip\n");
}

//+-------------------------------------------------------------------------
//
//  Method:     CTooltip::~CTooltip
//
//  Synopsis:   Destructor
//
//--------------------------------------------------------------------------
CTooltip::~CTooltip()
{
    Trace0("CTooltip::~CTooltip\n");

    if(_hwnd != NULL)
    {
        // bugbug: the presence of some third party controls (see 6383)
        // kill the tooltip window ahead of time.  hence we are conservative
        // here by ensuring that the window still is one.
        Verify(!IsWindow(_hwnd) || DestroyWindow(_hwnd));
        _hwnd = NULL;
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     CTooltip::Init
//
//  Synopsis:   2nd phase constructor
//
//--------------------------------------------------------------------------
HRESULT CTooltip::Init(HWND hwnd)
{
    HRESULT hr = S_OK;

    Trace0("CTooltip::Init\n");

    InitCommonControls();

    _hwnd = CreateWindowEx(
        0,
        TOOLTIPS_CLASS,
        NULL,
        TTS_NOPREFIX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        10,
        10,
        hwnd,
        NULL,
        _afxGlobalData._hInstThisModule,
        NULL);

    if(_hwnd == NULL)
    {
        hr = GetLastWin32Error();
        goto Cleanup;
    }

    SendMessage(_hwnd, TTM_SETMAXTIPWIDTH, 0, (LPARAM)(INT)3*_afxGlobalData._sizePixelsPerInch.cx);

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CTooltip::Show
//
//  Synopsis:   Displays the Tooltip.
//
//  Arguments:  szText      Text to display in the Tooltip
//              hwnd        Owner hwnd
//              msg         Message passed to tooltip for precessing
//              prc         Coords of rectangle in which the mouse
//                          should reside to display the Tooltip.  If the
//                          mouse moves out of this rect, the Tooltip will
//                          be brought down.
//              dwCookie    Any value that is used to identify a tooltip
//                          source.  If subsequent calls to FormsShowTooltip
//                          have the same dwCookie, the tooltip isn't
//                          redisplayed (avoiding flashing).  This value
//                          could be a pointer to the source's this ptr,
//                          an index, etc.
//              fRTL        COMPLEXSCRIPT flag indicating if the element
//                          lays out weak/neutral characters right to left
//
//--------------------------------------------------------------------------
void CTooltip::Show(TCHAR* szText, HWND hwnd, MSG msg, RECT* prc, DWORD_PTR dwCookie, BOOL fRTL)
{
    Trace0("CTooltip::Show\n");

    // If owner window is diffrent, 
    // create new tooltip window
    if(_hwndOwner != hwnd)
    {
        if(IsWindow(_hwnd))
        {
            DestroyWindow(_hwnd);
        }

        if(Init(hwnd) != S_OK)
        {
            return;
        }
    }

    // if tool source is different, cookie is different,
    // tooltip text is different, or rectangle is different,
    // delete old tool and create new tool.
    // BUGBUG: Since pElement is used as the cookie, and multiline elements
    // have disjoint rects, the tooltip will disappear between lines.
    // This doesn't look too annoying. (jbeda)
    if(_hwndOwner!=hwnd || _dwCookie!=dwCookie 
        || _tcscmp(_cstrText, szText) || !EqualRect(prc, &_ti.rect))
    {
        if(_ti.cbSize == sizeof(_ti)-sizeof(void*))
        {
            SendMessage(_hwnd, TTM_DELTOOL, 0, (LPARAM)&_ti);
        }

        memset(&_ti, 0, sizeof(_ti));

        // Update text
        _cstrText.Set(szText, _tcslen(szText));

        _ti.cbSize = sizeof(_ti) - sizeof(void*); // wlw note: ms tooltip bug
        _ti.uFlags = (fRTL ? TTF_RTLREADING : 0);
        _ti.hwnd = hwnd;
        _ti.uId = (UINT_PTR)dwCookie;
        _ti.rect.left = prc->left;
        _ti.rect.top = prc->top;
        _ti.rect.right = prc->right;
        _ti.rect.bottom = prc->bottom;
        _ti.lpszText = _cstrText;
        SendMessage(_hwnd, TTM_ADDTOOL, 0, (LPARAM)&_ti);
    }

    SendMessage(_hwnd, TTM_RELAYEVENT, 0, (LPARAM)&msg);
    SendMessage(_hwnd, TTM_ACTIVATE, (WPARAM)(BOOL)TRUE, 0);

    _hwndOwner = hwnd;
    _dwCookie = dwCookie;
}

//+-------------------------------------------------------------------------
//
//  Method:     CTooltip::Hide
//
//  Synopsis:   Hides the Tooltip window.
//
//--------------------------------------------------------------------------
void CTooltip::Hide()
{
    Trace0("CTooltip::Hide\n");
    SendMessage(_hwnd, TTM_ACTIVATE, (WPARAM)(BOOL)FALSE, 0);
}



void DeinitTooltip(THREADSTATE* pts)
{
    Assert(pts);

    if(pts->pTT)
    {
        delete pts->pTT;
        pts->pTT = NULL;
    }
}