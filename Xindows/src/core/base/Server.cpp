
#include "stdafx.h"
#include "Server.h"

ATOM CServer::s_atomWndClass;

//+---------------------------------------------------------------
//
//  Member:     CServer::GetWindow, IOleWindow
//
//  Synopsis:   Method of IOleWindow interface
//
//---------------------------------------------------------------
HRESULT CServer::GetWindow(HWND* lphwnd)
{
    if(_pInPlace == NULL)
    {
        RRETURN_NOTRACE(E_FAIL);
    }

    *lphwnd = _pInPlace->_hwnd;

    RRETURN_NOTRACE((_pInPlace->_hwnd ? S_OK : E_FAIL));
}

//+---------------------------------------------------------------
//
//  Member:     CServer::SetObjectRects, IOleInPlaceObject
//
//  Synopsis:   Method of IOleInPlaceObject interface
//
//  Notes:      This method does a Move window on the child
//              window to put it in its new position.
//
//---------------------------------------------------------------
HRESULT CServer::SetObjectRects(LPCRECT prcPos, LPCRECT prcClip)
{
    HRESULT     hr = S_OK;
    RECT        rcWnd;
    RECT        rcVisible;

    if(!prcPos)
    {
        RRETURN(E_INVALIDARG);
    }

    Assert(_pInPlace);
    Assert(prcPos->left <= prcPos->right);
    Assert(prcPos->top <= prcPos->bottom);

    // Handle bogus call from MFC container with NULL clip rectangle...
    if(!prcClip)
    {
        prcClip = prcPos;
    }

    Trace((_T("%08lx SetObjectRects pos=%ld %ld %ld %ld; clip=%ld %ld %ld %ld\n"),
        this, *prcPos, *prcClip));

    // Check for the container requesting an infinite scale factor.
    if((_sizel.cx==0 && prcPos->right-prcPos->left!=0) ||
        (_sizel.cy==0 && prcPos->bottom-prcPos->top!=0))
    {
        Assert(0 && "Host error: Infinite scale factor. Not a Forms?error.");
        RRETURN(E_FAIL);
    }

    // Update the inplace RECTs and offsets
    {
        Assert(_pInPlace->_hwnd);

        _pInPlace->_ptWnd = *(POINT*)&prcPos->left;

        CopyRect(&_pInPlace->_rcPos, prcPos);
        OffsetRect(&_pInPlace->_rcPos,  -_pInPlace->_ptWnd.x, -_pInPlace->_ptWnd.y);

        CopyRect(&_pInPlace->_rcClip, prcClip);
        OffsetRect(&_pInPlace->_rcClip, -_pInPlace->_ptWnd.x, -_pInPlace->_ptWnd.y);

        CopyRect(&rcWnd, prcPos);

        IntersectRect(&rcVisible, &rcWnd, prcClip);
        if(EqualRect(&rcVisible, &rcWnd))
        {
            if(_pInPlace->_fUsingWindowRgn)
            {
                SetWindowRgn(_pInPlace->_hwnd, NULL, TRUE);
                _pInPlace->_fUsingWindowRgn = FALSE;
            }
        }
        else 
        {
            _pInPlace->_fUsingWindowRgn = TRUE;
            OffsetRect(&rcVisible, -rcWnd.left, -rcWnd.top);
            SetWindowRgn(_pInPlace->_hwnd, CreateRectRgnIndirect(&rcVisible), TRUE);
        }

        // we go to some trouble here to make sure we aren't calling SetWindowPos
        // with invalid area in our window, because Windows sends a WM_ERASEBKGND
        // message to our window and all child windows if it finds a completely
        // invalid window at the time that SetWindowPos is called.  This causes
        // truly disastrous flashing to occur when we are resizing the window.
        HRGN hrgnUpdate = ::CreateRectRgnIndirect(&_afxGlobalData._Zero.rc);
        if(hrgnUpdate)
        {
            int result = ::GetUpdateRgn(_pInPlace->_hwnd, hrgnUpdate, FALSE);
            if(result != ERROR && result != NULLREGION)
            {
                ::ValidateRgn(_pInPlace->_hwnd, hrgnUpdate);
            }
            else
            {
                ::DeleteObject(hrgnUpdate);
                hrgnUpdate = NULL;
            }
        }

        Trace((_T("%08lx SetObjectRects > SetWindowPos %ld %ld %ld %ld\n"), this, rcWnd));

        SetWindowPos(_pInPlace->_hwnd, NULL, rcWnd.left, rcWnd.top,
            rcWnd.right-rcWnd.left, rcWnd.bottom-rcWnd.top, SWP_NOZORDER|SWP_NOACTIVATE);

        // restore previously invalid area
        if(hrgnUpdate)
        {
            if(_pInPlace && _pInPlace->_hwnd)
            {
                ::InvalidateRgn(_pInPlace->_hwnd, hrgnUpdate, FALSE);
            }
            ::DeleteObject(hrgnUpdate);
        }
    }

    RRETURN(hr);
}

//+------------------------------------------------------------------
//
//  Member:     CServer::OnWindowMessage
//
//  Synopsis:   Handles windows messages.
//
//  Arguments:  msg     the message identifier
//              wParam  the first message parameter
//              lParam  the second message parameter
//
//  Returns:    LRESULT as in WNDPROCs
//
//-------------------------------------------------------------------
HRESULT CServer::OnWindowMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult)
{
    HRESULT hr = S_OK;
    POINT pt;

    *plResult = 0;

    // We should always be stabilized when handling a window message.
    // This check insures that derived classes are doing the right thing.
    Assert(TestLock(SERVERLOCK_STABILIZED));

    // Events fired by the derived class implementation of OnWindowMessage
    // may have caused this object to leave the inplace state.  Bail out
    // here if it looks like this happened.  We check for _pInPlace instead
    // of _state < OS_INPLACE because the window procedure can be called
    // before we are officially in the inplace state.
    if(!_pInPlace)
    {
        return S_OK;
    }

    // Process the message.
    switch(msg)
    {
    case WM_DESTROY:
        OnDestroy();
        break;

    case WM_NCHITTEST:
        *plResult = OnNCHitTest(MAKEPOINTS(lParam));
        break;

    case WM_SETCURSOR:
        if(LOWORD(lParam) != HTCLIENT)
        {
            if(_pInPlace->_hwnd &&
                !OnNCSetCursor((HWND)wParam, LOWORD(lParam), HIWORD(lParam)))
            {
                *plResult = DefWindowProc(_pInPlace->_hwnd, msg, wParam, lParam);
            }
        }
        else
        {
            HWND hwnd;

            GetCursorPos(&pt);

            hwnd = GetHWND();
            if(hwnd)
            {
                ScreenToClient(hwnd, &pt);

                // Give the container a chance to set the cursor.
                OnDefWindowMessage(msg, wParam?wParam:(WPARAM)hwnd, lParam, plResult);
            }
        }
        break;

    case WM_MOUSEMOVE:
        break;

    case WM_ERASEBKGND:
        *plResult = OnEraseBkgnd((HDC)wParam);
        break;

    case WM_PAINT:
        OnPaint();
        break;

    case WM_SETFOCUS:
    case WM_KILLFOCUS:
        {
            _pInPlace->_fFocus = (msg == WM_SETFOCUS);
        }
        break;

        // OCX containers (e.g. VB4) will forward WM_PALETTECHANGED and WM_QUERYNEWPALETTE
        // on to us to properly realize our palette for controls in cases where they are
        // windowed.  This is semantically equivalent to the code in MinFrameWndProc in
        // minfr.cxx.
    case WM_PALETTECHANGED:
        Assert(_pInPlace);
        if((HWND)wParam == _pInPlace->_hwnd)
        {
            break;
        }
        // **** FALL THRU ****
    case WM_QUERYNEWPALETTE:
        {
            HDC hdc;
            Assert(_pInPlace);
            hdc = ::GetDC(_pInPlace->_hwnd);
            if(hdc)
            {
                BOOL fInvalidate = FALSE;
                HPALETTE hpal;

                hpal = GetPalette();
                if(hpal)
                {
                    HPALETTE hpalOld = SelectPalette(hdc, hpal, (msg==WM_PALETTECHANGED));
                    fInvalidate = RealizePalette(hdc) || (msg==WM_PALETTECHANGED);
                    SelectPalette(hdc, hpalOld, TRUE);

                    if(fInvalidate)
                    {
                        if (_pInPlace->_hwnd)
                        {
                            RedrawWindow(_pInPlace->_hwnd, NULL, NULL, RDW_INVALIDATE|RDW_ALLCHILDREN);
                        }
                        else
                        {
                            InvalidateRect(NULL, TRUE);
                        }
                    }
                }
                ::ReleaseDC(_pInPlace->_hwnd, hdc);
                *plResult = !!hpal;
            }
            break;
        }

    default:
        hr = OnDefWindowMessage(msg, wParam, lParam, plResult);
        break;
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::GetDropTarget
//
//  Synopsis:   returns IDropTarget interface.
//
//----------------------------------------------------------------------------
HRESULT CServer::GetDropTarget(IDropTarget** ppDropTarget)
{
    HRESULT hr;

    if(!ServerDesc()->TestFlag(SERVERDESC_SUPPORT_DRAG_DROP))
    {
        *ppDropTarget = NULL;
        hr = E_NOTIMPL;
    }
    else
    {
        *ppDropTarget = new CDropTarget(this);
        hr = *ppDropTarget ? S_OK : E_OUTOFMEMORY;
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CServer::TranslateAccelerator, IOleInPlaceActiveObject
//
//  Synopsis:   Method of IOleInPlaceActiveObject interface
//
//---------------------------------------------------------------
HRESULT CServer::TranslateAccelerator(LPMSG lpmsg)
{
    HRESULT         hr = S_FALSE;
    CLock           Lock(this);

    if(lpmsg->message>=WM_KEYFIRST && lpmsg->message<=WM_KEYLAST)
    {
        hr = DoTranslateAccelerator(lpmsg);
    }

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::DragEnter, IDropTarget
//
//----------------------------------------------------------------------------
STDMETHODIMP CServer::DragEnter(IDataObject* pIDataSource, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    HRESULT hr = S_OK;

    if(pIDataSource==NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // DragLeave will not always be called, since we bubble up to
    // our container's DragEnter if we decide we don't like the
    // data object offered.  As a result, we need to clear any
    // pointers hanging around from the last drag-drop. (chrisz)
    ReplaceInterface(&_pInPlace->_pDataObj, pIDataSource);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::DragOver, IDropTarget
//
//----------------------------------------------------------------------------
STDMETHODIMP CServer::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::DragLeave, IDropTarget
//
//----------------------------------------------------------------------------
STDMETHODIMP CServer::DragLeave(BOOL fDrop)
{
    ClearInterface(&_pInPlace->_pDataObj);
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::Drop, IDropTarget
//
//----------------------------------------------------------------------------
STDMETHODIMP CServer::Drop(IDataObject* pDataObject, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::FireEvent, public
//
//  Synopsis:   Fires an event out the primary dispatch event connection point,
//              but only after the control is initialized (or fully loaded).
//
//  Arguments:  [dispidEvent]   -- DISPID of event to fire
//              [dispidProp]    -- Dispid of prop storing event function
//              [pbTypes]       -- Pointer to array giving the types of params
//              [...]           -- Parameters
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CServer::FireEvent(DISPID dispidEvent, DISPID dispidProp,
                           IDispatch* pEventObject, BYTE * pbTypes, ...)
{
    va_list valParms;
    HRESULT hr = S_OK;

    va_start(valParms, pbTypes);
    hr = FireEventV(dispidEvent, dispidProp, pEventObject, NULL, pbTypes, valParms);
    va_end(valParms);

    RRETURN(hr);
}

//+------------------------------------------------------------------
//
//  Member:     CServer::PrivateQueryInterface, public
//
//  Synopsis:   QueryInterface on the private unknown for CServer
//
//  Arguments:  [riid] -- Interface to return
//              [ppv]  -- New interface returned here
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------
STDMETHODIMP CServer::PrivateQueryInterface(REFIID iid, void** ppv)
{
    *ppv = NULL;
    switch(iid.Data1)
    {
        QI_INHERITS((IPrivateUnknown*)this, IUnknown)
        QI_TEAROFF((CBase*)this, IOleCommandTarget, NULL)
    default:
        if(iid == IID_IConnectionPointContainer)
        {
            *((IConnectionPointContainer**)ppv) = new CConnectionPointContainer(this, NULL);

            if(!*ppv)
            {
                RRETURN(E_OUTOFMEMORY);
            }
        }
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    (*(IUnknown**)ppv)->AddRef();

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::ActivateInPlace, CServer
//
//  Synopsis:   In-place activates the object
//
//  Returns:    Success if we in-place activated properly
//
//  Notes:      This method implements the standard in-place activation
//              protocol and creates the in-place window.
//
//----------------------------------------------------------------------------
HRESULT CServer::ActivateInPlace(HWND hWndSite, LPMSG lpmsg)
{
    HRESULT hr;
    RECT    rcPos;
    RECT    rcVisible;
    BOOL    fUsingWindowlessSite = FALSE;
    BOOL    fNoRedraw = FALSE;
    WORD    wLockFlags;

    hr = EnsureInPlaceObject();
    if(hr)
    {
        RRETURN(hr);
    }

    Assert(_pInPlace);
    _pInPlace->_fDeactivating = FALSE;

    // Check for the container requesting an infinite scale factor.
    if((_sizel.cx==0 && _pInPlace->_rcPos.right-_pInPlace->_rcPos.left!=0) ||
        (_sizel.cy==0 && _pInPlace->_rcPos.bottom-_pInPlace->_rcPos.top!=0))
    {
        Assert(0 && "Host error: Infinite scale factor. Not a Forms?error.");
        hr = E_FAIL;
        goto Error;
    }

    Trace((_T("%08lx ActivateInPlace context pos=%ld %ld %ld %ld; clip=%ld %ld %ld %ld\n"),
        this, _pInPlace->_rcPos, _pInPlace->_rcClip));

    Trace3("%08lx ActivateInPlace extent=%ld %ld", this, _sizel.cx, _sizel.cy);

    rcPos = _pInPlace->_rcPos;

    _pInPlace->_ptWnd  = *(POINT*)&_pInPlace->_rcPos;
    OffsetRect(&_pInPlace->_rcPos, -_pInPlace->_ptWnd.x, -_pInPlace->_ptWnd.y);
    OffsetRect(&_pInPlace->_rcClip, -_pInPlace->_ptWnd.x, -_pInPlace->_ptWnd.y);

    hr = AttachWin(hWndSite, &rcPos, &_pInPlace->_hwnd);
    if(hr)
    {
        goto Error;
    }

    IntersectRect(&rcVisible, &_pInPlace->_rcPos, &_pInPlace->_rcClip);
    if(!EqualRect(&rcVisible, &_pInPlace->_rcPos))
    {
        _pInPlace->_fUsingWindowRgn = TRUE;
        SetWindowRgn(_pInPlace->_hwnd, CreateRectRgnIndirect(&rcVisible), FALSE);
    }

    wLockFlags = Unlock(SERVERLOCK_TRANSITION);

    Relock(wLockFlags);

    if(!fNoRedraw)
    {
        InvalidateRect(NULL, TRUE);
    }

Cleanup:
    RRETURN(hr);

Error:
    DeactivateInPlace();
    goto Cleanup;
}

//+---------------------------------------------------------------
//
//  Member:     CServer::DeactivateInPlace, CServer
//
//  Synopsis:   In-place deactivates the object
//
//  Returns:    Success except for catastophic circumstances
//
//  Notes:      This method "undoes" everything done in ActivateInPlace
//              including destroying the inplace active window.
//
//---------------------------------------------------------------
HRESULT CServer::DeactivateInPlace()
{
    Assert(_pInPlace);

    DetachWin();

    delete _pInPlace;
    _pInPlace = NULL;

    return S_OK;
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::GetDC
//
//  Synopsis:   Gets a DC for the caller, type of DC depends on flags.
//
//  Arguments:  [prc]     -- param as per IOleInPlaceSiteWindowless::GetDC
//              [dwFlags] --                -do-
//              [phDC]    --                -do-
//
//  Returns:    HRESULT
//
//  History:    28-Mar-95   SumitC      Created
//
//  Notes:
//
//---------------------------------------------------------------------
HRESULT CServer::GetDC(LPRECT prc, DWORD dwFlags, HDC* phDC)
{
    HRESULT hr = S_OK;

    if(phDC == NULL)
    {
        return E_POINTER;
    }

    *phDC = NULL;

    if(_pInPlace == NULL)
    {
        return E_FAIL;
    }

    _pInPlace->_fIPNoDraw = _pInPlace->_fIPPaintBkgnd = _pInPlace->_fIPOffScreen = FALSE;

    if(_pInPlace->_hwnd)
    {
        // get the window dc
        *phDC = ::GetDC(_pInPlace->_hwnd);
        if(*phDC == NULL)
        {
            RRETURN(GetLastWin32Error());
        }

        HPALETTE hpal = GetPalette(*phDC);

        if(dwFlags & OLEDC_OFFSCREEN)
        {
            // build an offscreen buffer to return
            Assert(_pInPlace->_pOSC == NULL);
            _pInPlace->_pOSC = new COffScreenContext
                (*phDC,
                prc->right-prc->left,
                prc->bottom-prc->top,
                hpal,
                ((dwFlags>>16)&OFFSCR_BPP)
                |(dwFlags&OFFSCR_SURFACE)
                |(dwFlags&OFFSCR_3DSURFACE));

            if(_pInPlace->_pOSC == NULL)
            {
                ::ReleaseDC(_pInPlace->_hwnd, *phDC);
                RRETURN(E_OUTOFMEMORY);
            }

            *phDC = _pInPlace->_pOSC->GetDC(prc);
            _pInPlace->_fIPOffScreen = TRUE;
        }

        if(dwFlags & OLEDC_NODRAW)
        {
            _pInPlace->_fIPNoDraw = TRUE;
        }
        else if(dwFlags & OLEDC_PAINTBKGND)
        {
            _pInPlace->_fIPPaintBkgnd = TRUE;
        }
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::ReleaseDC
//
//  Synopsis:   Releases a DC obtained via GetDC above.
//
//  Arguments:  [hDC] -- param as per IOleInPlaceSiteWindowless::ReleaseDC
//
//  Returns:    HRESULT
//
//---------------------------------------------------------------------
HRESULT CServer::ReleaseDC(HDC hDC)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);

    // Get our palette out of the DC so that it doesn't stay locked
    SelectPalette(hDC, (HPALETTE)GetStockObject(DEFAULT_PALETTE), TRUE);

    if(_pInPlace->_hwnd)
    {
        if(_pInPlace->_fIPOffScreen)
        {
            Assert(_pInPlace->_pOSC);

            ::ReleaseDC(_pInPlace->_hwnd, _pInPlace->_pOSC->ReleaseDC(_pInPlace->_hwnd, !_pInPlace->_fIPNoDraw));

            delete _pInPlace->_pOSC;
            _pInPlace->_pOSC = NULL;
        }
        else
        {
            if(::ReleaseDC(_pInPlace->_hwnd, hDC) == 0)
            {
                hr = GetLastWin32Error();
            }
        }
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::InvalidateRect
//
//  Synopsis:
//
//  Arguments:  [prc]    -- param as per IOleInPlaceSiteWindowless::InvalidateRect
//              [fErase] --             -do-
//
//  Returns:    HRESULT
//
//---------------------------------------------------------------------
HRESULT CServer::InvalidateRect(LPCRECT prc, BOOL fErase)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);

    if(_pInPlace->_hwnd)
    {
        if(::InvalidateRect(_pInPlace->_hwnd, prc, fErase) == 0)
        {
            hr = GetLastWin32Error();
        }
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::InvalidateRgn
//
//  Synopsis:
//
//  Arguments:  [hrgn]   -- param as per IOleInPlaceSiteWindowless::InvalidateRgn
//              [fErase] --                 -do-
//
//  Returns:    HRESULT
//
//---------------------------------------------------------------------
HRESULT CServer::InvalidateRgn(HRGN hrgn, BOOL fErase)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);

    if(_pInPlace->_hwnd)
    {
        ::InvalidateRgn(_pInPlace->_hwnd, hrgn, fErase); // always returns TRUE
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::Scroll
//
//  Synopsis:
//
//  Arguments:  [dx]        --
//              [dy]        --
//              [prcScroll] --
//              [prcClip]   --
//
//  Returns:    HRESULT
//
//  History:    28-Mar-95   SumitC      Created
//
//  Notes:
//
//---------------------------------------------------------------------
HRESULT CServer::Scroll(int dx, int dy, LPCRECT prcScroll, LPCRECT prcClip)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);

    if(::ScrollWindowEx(_pInPlace->_hwnd,
        dx,
        dy,
        prcScroll,
        prcClip,
        NULL,
        NULL,
        SW_INVALIDATE) == ERROR)
    {
        hr = GetLastWin32Error();
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::AdjustRect
//
//  Synopsis:
//
//  Arguments:  [prc] -- param as per IOleInPlaceSiteWindowless::AdjustRect
//
//  Returns:    HRESULT
//
//---------------------------------------------------------------------
HRESULT CServer::AdjustRect(LPRECT prc)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);

    IntersectRect(prc, prc, &_pInPlace->_rcClip);

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::SetCapture, public
//
//  Synopsis:   fCaptured --> Capture the mouse to this server
//              !fCaptured --> This server will not have the capture.
//
//---------------------------------------------------------------------
HRESULT CServer::SetCapture(BOOL fCaptured)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace->_hwnd);
    if(fCaptured)
    {
        // Capture the focus.
        ::SetCapture(_pInPlace->_hwnd);
    }
    else if(::GetCapture() == _pInPlace->_hwnd)
    {
        // Release capture if we have it.
        ::ReleaseCapture();
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::GetCapture, public
//
//  Synopsis:   Find out if we have captured the mouse.
//
//---------------------------------------------------------------------
BOOL CServer::GetCapture()
{
    Assert(_pInPlace);
    Assert(_pInPlace->_hwnd);

    return ::GetCapture()==_pInPlace->_hwnd;
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::SetFocus, public
//
//  Synopsis:   Set the focus to this server.
//
//---------------------------------------------------------------------
HRESULT CServer::SetFocus(BOOL fFocus)
{
    HRESULT hr = S_OK;

    Assert(_pInPlace);
    Assert(_pInPlace->_hwnd);
    
    if(fFocus)
    {
        // Capture the focus.
        ::SetFocus(_pInPlace->_hwnd);
    }

    RRETURN(hr);
}

//+--------------------------------------------------------------------
//
//  Member:     CServer::GetFocus, public
//
//  Synopsis:   Determines whether the client site has focus
//
//---------------------------------------------------------------------
BOOL CServer::GetFocus()
{
    return (_pInPlace->_hwnd==::GetFocus());
}

//+---------------------------------------------------------------
//
//  Member:     CServer::GetHWND()
//
//  Synopsis:   Get window used by this CServer.
//
//---------------------------------------------------------------
HWND CServer::GetHWND()
{
    HWND hwnd= _pInPlace->_hwnd;
    return hwnd;
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::AttachWin
//
//  Synopsis:   Creates the in-place window for the object.  The base
//              class does nothing here; any derived object which can
//              be in-place or UI Activated should override this method
//
//  Returns:    HWND
//
//-------------------------------------------------------------------------
HRESULT CServer::AttachWin(HWND hWndParent, RECT* prc, HWND* phWnd)
{
    HRESULT hr = S_OK;
    HWND    hwnd;
    TCHAR*  pszBuf;

    if(!s_atomWndClass)
    {
        hr = RegisterWindowClass(
            _T("Server"),
            CServer::WndProc,
            CS_DBLCLKS,
            NULL, NULL,
            &CServer::s_atomWndClass);
        if(hr)
        {
            hwnd = NULL;
            goto Cleanup;
        }
    }

    Assert(phWnd);

    pszBuf = (TCHAR*)(DWORD_PTR)s_atomWndClass; 

    hwnd = CreateWindowEx(
        0,
        pszBuf,
        NULL,
        WS_CHILD|WS_CLIPCHILDREN|WS_CLIPSIBLINGS,
        prc->left, prc->top,
        prc->right-prc->left, prc->bottom-prc->top,
        hWndParent,
        0, // no child identifier - shouldn't send WM_COMMAND
        _afxGlobalData._hInstThisModule,
        this);
    if(!hwnd)
    {
        hr = GetLastWin32Error();
        DetachWin();

        goto Cleanup;
    }


    SetWindowPos(_pInPlace->_hwnd, NULL, 0, 0, 0, 0,
        SWP_SHOWWINDOW|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOREDRAW|SWP_NOACTIVATE);

Cleanup:
    Assert(hr || hwnd);
    *phWnd = hwnd;
    RRETURN(hr);
}

//+--------------------------------------------------------------
//
//  Member:     CServer::DetachWin, CServer
//
//  Synopsis:   Detaches the child's in-place
//              window from the current parent.
//
//  Notes:      This destroys the _hwnd of the server.
//              If the derived class does anything
//              other than create a Window on AttachWin,
//              it must over-ride this function.
//              If the derived class destroys the window
//              on detach, it must set _hwnd = NULL
//
//---------------------------------------------------------------
void CServer::DetachWin()
{
    Assert(_pInPlace);
    Assert(IsWindow(_pInPlace->_hwnd));
    Verify(DestroyWindow(_pInPlace->_hwnd));

    _pInPlace->_hwnd = NULL;
}

//+----------------------------------------------------------------
//
//  Member:     CServer::OnNCSetCursor
//
//  Synopsis:   Set the cursor for the non-licent area.
//
//-----------------------------------------------------------------
BOOL CServer::OnNCSetCursor(HWND hwnd, int nHitTest, UINT msg)
{
    TCHAR* idc;

    // Set cursor for non-client area.  We don't let the default
    // window procedure handle this because it gives the container
    // a chance to override what we really want.
    if(hwnd!=_pInPlace->_hwnd || msg!=WM_MOUSEMOVE)
    {
        return FALSE;
    }

    switch(nHitTest)
    {
    case HTCAPTION:
        idc = IDC_SIZEALL;
        break;

    case HTLEFT:
    case HTRIGHT:
        idc = IDC_SIZEWE;
        break;

    case HTTOP:
    case HTBOTTOM:
        idc = IDC_SIZENS;
        break;

    case HTTOPLEFT:
    case HTBOTTOMRIGHT:
        idc = IDC_SIZENWSE;
        break;

    case HTTOPRIGHT:
    case HTBOTTOMLEFT:
        idc = IDC_SIZENESW;
        break;

    default:
        return FALSE;
    }

    SetCursorIDC(idc);

    return TRUE;
}

//+----------------------------------------------------------------
//
//  Member:     CServer::OnNCHitTest
//
//  Synopsis:   Returns the hit code for the border part at the given
//              point.
//
//-----------------------------------------------------------------
LONG CServer::OnNCHitTest(POINTS pts)
{
    RECT    rc;
    POINT   pt;

    pt.x = pts.x;
    pt.y = pts.y;

    // If it clearly is in the client area, return that fact.
    GetClientRect(_pInPlace->_hwnd , &rc);
    ScreenToClient(_pInPlace->_hwnd, &pt);
    if(PtInRect(&rc, pt))
    {
        return HTCLIENT;
    }

    return HTNOWHERE;
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::OnPaint
//
//  Synopsis:   Draws the client area when a window is present
//
//-------------------------------------------------------------------------
void CServer::OnPaint()
{
}

//+-----------------------------------------------------------------------
//
//  Member:     CServer::OnEraseBkgnd
//
//  Synopsis:   Erases the background when a window is present
//
//------------------------------------------------------------------------
BOOL CServer::OnEraseBkgnd(HDC hdc)
{
    return FALSE;
}

//+-----------------------------------------------------------------------
//
//  Member:    CServer::OnDestroy
//
//  Synopsis:  Deregister Drag and Drop
//
//------------------------------------------------------------------------
void CServer::OnDestroy( )
{
    Assert(_pInPlace->_hwnd);
    RevokeDragDrop(_pInPlace->_hwnd);
}

//+-----------------------------------------------------------------------
//
//  Member:     CServer::OnDefWindowMessage
//
//  Synopsis:   Default handling for window messages.
//
//  Arguments:  msg      the message identifier
//              wParam   the first message parameter
//              lParam   the second message parameter
//              plResult return value from window proc
//
//------------------------------------------------------------------------
HRESULT CServer::OnDefWindowMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult)
{
    HRESULT hr = S_OK;

    // Events fired by the derived class implementation of OnWindowMessage
    // may have caused this object to leave the inplace state.  Bail out
    // here if it looks like this happened.  We check for _pInPlace instead
    // of _state < OS_INPLACE because the window procedure can be called
    // before we are officially in the inplace state.
    if(!_pInPlace)
    {
        return S_OK;
    }

    if(_pInPlace->_hwnd)
    {
        *plResult = DefWindowProc(_pInPlace->_hwnd, msg, wParam, lParam);      
    }
    else
    {
        *plResult = 0;
    }

    RRETURN(hr);
}

//+-----------------------------------------------------------------------
//
//  Member:     CServer::WndProc
//
//  Synopsis:   Window procedure for use by derived class.
//              This function maintains the relationship between
//              CServer and HWND and delegates all other functionality
//              to the CServer::OnWindowMessage virtual method.
//
//  Arguments:  hwnd     the window
//              msg      the message identifier
//              wParam   the first message parameter
//              lParam   the second message parameter
//              plResult the window procedure return value
//
//  Returns:    S_FALSE if caller should delegate to the default window proc
//
//------------------------------------------------------------------------
LRESULT CALLBACK CServer::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    CServer* pServer;
    LRESULT lResult = 0;

    // a-msadek; Trident window should not be mirrored if hosted by a mirrored process
    if(msg == WM_NCCREATE)
    {
        DWORD dwExStyles;
        if((dwExStyles=GetWindowLong(hwnd, GWL_EXSTYLE)) & WS_EX_LAYOUTRTL)
        {
            SetWindowLong(hwnd, GWL_EXSTYLE, dwExStyles&~WS_EX_LAYOUTRTL);
        }
    }

    // If creating, then establish the connection between the HWND and
    // the CServer.  Otherwise, fetch the pointer to the CServer.
    if(msg == WM_NCCREATE)
    {
        pServer = (CServer*)((LPCREATESTRUCTW)lParam)->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pServer);
        pServer->_pInPlace->_hwnd = hwnd; 
        pServer->PrivateAddRef();
    }
    else
    {
        pServer = (CServer*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    }

    if(pServer)
    {
        Assert(pServer->_pInPlace);
        Assert(pServer->_pInPlace->_hwnd == hwnd);

        // Give the derived class a chance to handle the window message.
        pServer->OnWindowMessage(msg, wParam, lParam, &lResult);

        // If destroying, break the connection between the HWND and CServer.
        // The call to release might destroy the server, so it must come
        // after the OnWindowMessage virtual function call.
        if(msg == WM_NCDESTROY)
        {
            SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
            pServer->_pInPlace->_hwnd = NULL;
            pServer->PrivateRelease();
        }
    }

    return lResult;
}

//+---------------------------------------------------------------
//
//  Member:     CServer::CServer
//
//  Synopsis:   Constructor for CServer object
//
//  Notes:      To create a properly initialized object you must
//              call the Init method immediately after construction.
//
//---------------------------------------------------------------
CServer::CServer() : CBase()
{
    Trace0("constructing CServer\n");

    IncrementObjectCount();

    _sizel.cx = _sizel.cy = 1;
}

//+---------------------------------------------------------------
//
//  Member:     CServer::~CServer
//
//  Synopsis:   Destructor for the CServer object
//
//  Notes:      The destructor is called as a result of the servers
//              reference count going to 0.  It ensure the object
//              is in a passive state and releases the data/view and inplace
//              subobjects objects.
//
//---------------------------------------------------------------
CServer::~CServer(void)
{
    // Release interface pointers and strings

    // Delete here just in case we created it without going inplace.
    delete _pInPlace;
    _pInPlace = NULL;

    DecrementObjectCount();
    
    Trace0("destructed CServer\n");
}

//+---------------------------------------------------------------
//
//  Member:     CServer::RunningToInPlace, protected
//
//  Synopsis:   Effects the direct Running to inplace state transition
//
//  Returns:    Success iff the object is in the inplace state.  On failure
//              the object will be in a consistent Running state.
//
//  Notes:      This transition invokes the ActivateInPlace method.  Containers
//              will typically override this method in order to additionally
//              inplace activate any inside-out embeddings that are visible.
//
//---------------------------------------------------------------
HRESULT CServer::RunningToInPlace(HWND hWndSite, LPMSG lpmsg)
{
    HRESULT hr;

    hr = ActivateInPlace(hWndSite, lpmsg);
    if(hr)
    {
        goto Error;
    }

Error:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CServer::InPlaceToRunning, protected
//
//  Synopsis:   Effects the direct inplace to Running state transition
//
//  Returns:    Success under all but catastrophic circumstances.
//
//  Notes:      This transition invokes the DeactivateInPlace method.
//
//              Containers will typically override this method in
//              order to additionally inplace deactivate any inplace-active
//              embeddings.
//
//              This method is called as the result of a DoVerb(HIDE...)
//
//---------------------------------------------------------------
HRESULT CServer::InPlaceToRunning(void)
{
    HRESULT hr;

    hr = DeactivateInPlace();

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::EnsureInPlaceObject, CServer
//
//  Synopsis:   Creates the InPlace object when needed.
//
//  Arguments:  (none)
//
//----------------------------------------------------------------------------
HRESULT CServer::EnsureInPlaceObject()
{
    if(!_pInPlace)
    {
        _pInPlace = new CInPlace();
        if(!_pInPlace)
        {
            RRETURN(E_OUTOFMEMORY);
        }
    }

    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     CServer::GetPalette, public
//
//  Synopsis:   Returns the document palette
//
//  Notes:      Returns the document palette.  This implementation is
//              really lame and doesn't cache the result from the ambient.
//              it is really expected to be overridden by the derived
//              class.
//
//---------------------------------------------------------------
HPALETTE CServer::GetPalette(HDC hdc, BOOL* pfHtPal)
{
    CVariant var;
    HPALETTE hpal = GetDefaultPalette();

    if(hdc && hpal)
    {
        SelectPalette(hdc, hpal, TRUE);
        RealizePalette(hdc);
    }

    if(pfHtPal)
    {
        *pfHtPal = FALSE;
    }

    return hpal;
}

//+---------------------------------------------------------------------------
//
//  Member:     CServer::OnPropertyChange
//
//  Synopsis:   Fires property change event, and then OnDataChange
//
//  Arguments:  [dispidProperty] -- PROPID of property that changed
//              [dwFlags]        -- Flags to inhibit behavior
//
//  Notes:      The [dwFlags] parameter has the following values that can be
//              OR'd together:
//
//              SERVERCHNG_NOPROPCHANGE -- Inhibits the OnChanged notification
//                 through the PropNotifySink.
//              SERVERCHNG_NOVIEWCHANGE -- Inhibits the OnViewChange notification.
//              SERVERCHNG_NODATACHANGE -- Inhibits the OnDataChange notification.
//
//----------------------------------------------------------------------------
HRESULT CServer::OnPropertyChange(DISPID dispidProperty, DWORD dwFlags)
{
    if(TestLock(SERVERLOCK_PROPNOTIFY))
    {
        return S_OK;
    }

    if(!(dwFlags & SERVERCHNG_NOPROPCHANGE))
    {
        FireOnChanged(dispidProperty);
    }

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::CLock::CLock
//
//  Synopsis:   Lock resources in CServer object.
//
//-------------------------------------------------------------------------
CServer::CLock::CLock(CServer* pServer, WORD wLockFlags)
{
    _pServer = pServer;
    _wLockFlags = pServer->_wLockFlags;
    pServer->_wLockFlags |= wLockFlags | SERVERLOCK_STABILIZED;

    pServer->PrivateAddRef();
}

CServer::CLock::CLock(CServer* pServer)
{
    _pServer = pServer;
    _wLockFlags = pServer->_wLockFlags;
    pServer->_wLockFlags |= SERVERLOCK_STABILIZED;

    pServer->PrivateAddRef();
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::CLock::~CLock
//
//  Synopsis:   Unlock resources in CServer object.
//
//-------------------------------------------------------------------------
CServer::CLock::~CLock()
{
    _pServer->_wLockFlags = _wLockFlags;

    _pServer->PrivateRelease();
}