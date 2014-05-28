
#include "stdafx.h"
#include "View.h"

#include "../display/DispRoot.h"
#include "../layout/Layout.h"

//+---------------------------------------------------------------------------
//
//  Member:     Getxxxxx
//
//  Synopsis:   Return _pv properly type-casted
//
//----------------------------------------------------------------------------
CElement* CViewTask::GetElement() const
{
    return DYNCAST(CElement, _pElement);
}

CLayout* CViewTask::GetLayout() const
{
    return DYNCAST(CLayout, _pLayout);
}

DISPID CViewTask::GetEventDispID() const
{
    return (DISPID)_dispidEvent;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetSourceIndex
//
//  Synopsis:   Return the source index for the object associated with a task
//
//----------------------------------------------------------------------------
long CViewTask::GetSourceIndex() const
{
    Assert(_vtt == VTT_LAYOUT);
    return GetLayout()->ElementOwner()->GetSourceIndex();
}



//+---------------------------------------------------------------------------
//
//  Member:     QueryInterface
//
//  Synopsis:   Implementation of QueryInterface
//
//----------------------------------------------------------------------------
HRESULT CViewDispClient::QueryInterface(REFIID iid, LPVOID* ppv)
{
    if(!ppv)
	{
		RRETURN(E_INVALIDARG);
	}

    *ppv = NULL;

	if(iid == IID_IUnknown)
	{
		(*ppv) = (IDispObserver*)this;
	}
	else if(iid == IID_IDispObserver)
	{
		(*ppv) = (IDispObserver*)this;
	}

    if(*ppv)
    {
        ((IUnknown*)*ppv)->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

//+---------------------------------------------------------------------------
//
//  Member:     Release
//
//  Synopsis:   Release reference
//
//----------------------------------------------------------------------------
ULONG CViewDispClient::Release()
{
    if(--_cRefs == 0)
    {
        return 0;
    }

    return _cRefs;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetOwner
//
//  Synopsis:   Return display node owner
//
//----------------------------------------------------------------------------
void CViewDispClient::GetOwner(CDispNode* pDispNode, void** ppv)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    Assert(pDispNode);
    Assert(pDispNode == View()->_pDispRoot);
    Assert(ppv);
    *ppv = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClient
//
//  Synopsis:   Draw client content
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClient(
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    CDispSurface*	pDispSurface,
    CDispNode*		pDispNode,
    void*			cookie,
    void*			pClientData,
    DWORD			dwFlags)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClientBackground
//
//  Synopsis:   Draw background
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClientBackground(
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    CDispSurface*	pDispSurface,
    CDispNode*		pDispNode,
    void*			pClientData,
    DWORD			dwFlags)
{
    Assert(pDispNode);
    Assert(pDispNode == View()->_pDispRoot);
    Assert(pClientData);

    CFormDrawInfo* pDI = (CFormDrawInfo*)pClientData;
    CSetDrawSurface sds(pDI, prcBounds, prcRedraw, pDispSurface);

    View()->RenderBackground(pDI);
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClientBorder
//
//  Synopsis:   Draw border
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClientBorder(
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    CDispSurface*	pDispSurface,
    CDispNode*		pDispNode,
    void*			pClientData,
    DWORD			dwFlags)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClientScrollbar
//
//  Synopsis:   Draw scrollbar
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClientScrollbar(
    int				iDirection,
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    long			contentSize,
    long			containerSize,
    long			scrollAmount,
    CDispSurface*	pDispSurface,
    CDispNode*		pDispNode,
    void*			pClientData,
    DWORD			dwFlags)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClientScrollbarFiller
//
//  Synopsis:   Draw scrollbar
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClientScrollbarFiller(
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    CDispSurface*	pDispSurface,
    CDispNode*		pDispNode,
    void*			pClientData,
    DWORD			dwFlags)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestScrollbar
//
//  Synopsis:   Hit test the scrollbar
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::HitTestScrollbar(
    int				iDirection,
    const POINT*	pptHit,
    CDispNode*		pDispNode,
    void*			pClientData)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestScrollbarFiller
//
//  Synopsis:   Hit test the scrollbar filler
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::HitTestScrollbarFiller(
    const POINT*	pptHit,
    CDispNode*		pDispNode,
    void*			pClientData)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestContent
//
//  Synopsis:   Hit test the display node
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::HitTestContent(
    const POINT*	pptHit,
    CDispNode*		pDispNode,
    void*			pClientData)
{
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestFuzzy
//
//  Synopsis:   Hit test the display node
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::HitTestFuzzy(
    const POINT*	pptHitInParentCoords,
    CDispNode*		pDispNode,
    void*			pClientData)
{
    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestBorder
//
//  Synopsis:   Hit test the display node
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::HitTestBorder(
    const POINT*	pptHit,
    CDispNode*		pDispNode,
    void*			pClientData)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     ProcessDisplayTreeTraversal
//
//  Synopsis:   Process display tree traversal
//
//----------------------------------------------------------------------------
BOOL CViewDispClient::ProcessDisplayTreeTraversal(void* pClientData)
{
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetZOrderForSelf
//
//  Synopsis:   Return z-index
//
//----------------------------------------------------------------------------
LONG CViewDispClient::GetZOrderForSelf()
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetZOrderForChild
//
//  Synopsis:   Return z-index for child node
//
//----------------------------------------------------------------------------
LONG CViewDispClient::GetZOrderForChild(void* cookie)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     CompareZOrder
//
//  Synopsis:   Compare the z-order of a display node with this display node
//
//----------------------------------------------------------------------------
LONG CViewDispClient::CompareZOrder(CDispNode* pDispNode1, CDispNode* pDispNode2)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return -1;
}

//+---------------------------------------------------------------------------
//
//  Member:     HandleViewChange
//
//  Synopsis:   Respond to a change in position
//
//----------------------------------------------------------------------------
void CViewDispClient::HandleViewChange(
    DWORD		flags,
    const RECT*	prcClient,
    const RECT*	prcClip,
    CDispNode*	pDispNode)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     GetFilter
//
//  Synopsis:   Get filter interface.
//
//----------------------------------------------------------------------------
CDispFilter* CViewDispClient::GetFilter()
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     NotifyScrollEvent
//
//  Synopsis:   Respond to a change in scroll position
//
//----------------------------------------------------------------------------
void CViewDispClient::NotifyScrollEvent(RECT* prcScroll, SIZE* psizeScrollDelta)
{
    if(prcScroll && psizeScrollDelta)
    {
        View()->ScrollInvalid((CRect&)*prcScroll, (CSize&)*psizeScrollDelta);
    }

    // the invalid area that was already published to Windows may need to be
    // adjusted too.  Get the update region for our window, mask it by the
    // rect that is being scrolled, and adjust it.
    HWND hwnd = Doc()->GetHWND();
    if(hwnd != NULL)
    {
        HRGN hrgnUpdate = ::CreateRectRgnIndirect(&_afxGlobalData._Zero.rc);
        if(hrgnUpdate != NULL)
        {
            if(::GetUpdateRgn(hwnd, hrgnUpdate, FALSE) != NULLREGION)
            {
                CRegion rgnUpdate(hrgnUpdate);
                rgnUpdate.Offset(*psizeScrollDelta);
                rgnUpdate.Intersect(*prcScroll);
                View()->Invalidate(rgnUpdate);
                ::ValidateRect(hwnd, prcScroll);
            }
            
            ::DeleteObject(hrgnUpdate);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetClientLayersInfo
//
//  Synopsis:   Return rendering layer information flags
//
//----------------------------------------------------------------------------
DWORD CViewDispClient::GetClientLayersInfo(CDispNode* pDispNodeFor)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
    return 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawClientLayers
//
//  Synopsis:   Render a layer
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawClientLayers(
    const RECT*		prcBounds,
    const RECT*		prcRedraw,
    CDispSurface*	pSurface,
    CDispNode*		pDispNode,
    void*			cookie,
    void*			pClientData,
    DWORD			dwFlags)
{
    AssertSz(FALSE, "Illegal CDispClient method called on view");
}

//+---------------------------------------------------------------------------
//
//  Member:     Invalidate
//
//  Synopsis:   Forward invalidate to appropriate receiver
//
//  Arguments:  prcInvalid - The invalid rect (if rgnInvalid is NULL)
//              rgnInvalid - The invalid region
//              flags      - Flags to force synchronous redraw and invalidation of child windows
//
//----------------------------------------------------------------------------
void CViewDispClient::Invalidate(
    const RECT*	prcInvalid,
    HRGN		rgnInvalid,
    DWORD		flags)
{
    BOOL fSynchronousRedraw = (((DWORD)flags)&0x01) && !(View()->GetLayoutFlags()&LAYOUT_FORCE);
    BOOL fInvalChildWindows = ((DWORD)flags) & 0x02;

    if(rgnInvalid)
    {
        View()->Invalidate(rgnInvalid, fSynchronousRedraw, fInvalChildWindows);
    }
    else if(prcInvalid)
    {
        View()->Invalidate((CRect*)prcInvalid, fSynchronousRedraw, fInvalChildWindows);
    }
    else if(fSynchronousRedraw)
    {
        View()->PostRenderView(TRUE);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetClientDC
//
//  Synopsis:   Get a DC to be used for scrolling and drawing
//
//  Arguments:  prc - Bounds of area to be painted
//
//  Returns:    HDC obtained
//
//----------------------------------------------------------------------------
HDC CViewDispClient::GetClientDC(const RECT* prc)
{
    Assert(Doc());
    Assert(Doc()->GetHWND());
    HDC hdc = ::GetDCEx(Doc()->GetHWND(), NULL, DCX_CACHE|DCX_CLIPSIBLINGS);

    Doc()->GetPalette(hdc);

    return hdc;
}

//+---------------------------------------------------------------------------
//
//  Member:     ReleaseClientDC
//
//  Synopsis:   Release DC obtained through GetClientDC
//
//  Arguments:  HDC - DC to be released
//
//----------------------------------------------------------------------------
void CViewDispClient::ReleaseClientDC(HDC hdc)
{
    Assert(Doc());
    Assert(Doc()->GetHWND());
    Verify(::ReleaseDC(Doc()->GetHWND(), hdc));
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawSynchronous
//
//  Synopsis:   Draw the display tree synchronously
//
//  Arguments:  hrgn    - Region to draw
//              hdc     - DC to draw into
//              prcClip - Clip rect
//
//----------------------------------------------------------------------------
void CViewDispClient::DrawSynchronous(
    HRGN		hrgn,
    HDC			hdc,
    const RECT*	prcClip)
{
    // BUGBUG: Remove the arguments from this method - They should not be used (brendand)
    Assert(!hrgn && !hdc && !prcClip);

    View()->PostRenderView(TRUE);
}



//+---------------------------------------------------------------------------
//
//  Member:     CView/~CView
//
//  Synopsis:   Constructor/Destructor
//
//----------------------------------------------------------------------------
CView::CView()
{
    ClearRanges();
}

CView::~CView()
{
    Assert(_pDispRoot == NULL);
    Assert(_aryTaskMisc.Size() == 0);
    Assert(_aryTaskLayout.Size() == 0);
    Assert(_aryTaskEvent.Size() == 0);
    Assert(_client._cRefs   == 0);
    Assert(!HasInvalid());
}

//+---------------------------------------------------------------------------
//
//  Member:     Initialize
//
//  Synopsis:   Initialize the view
//
//  Arguments:  pDoc - Pointer to owning CDoc
//
//----------------------------------------------------------------------------
void CView::Initialize(CDocument* pDoc)
{
    Assert(pDoc);
    Assert(!_pDoc);
    Assert(!_pDispRoot);

    _pDoc = pDoc;
}

//+---------------------------------------------------------------------------
//
//  Member:     Activate
//
//  Synopsis:   Activate the view
//
//----------------------------------------------------------------------------
HRESULT CView::Activate()
{
    Assert(IsInitialized());
    Assert(!_pDispRoot);
    Assert(!_pLayout);
    Assert(!_pDispRoot);

    Trace0("View : Activate\n");

    _pDispRoot = new CDispRoot(&_client, &_client);

    if(_pDispRoot)
    {
        // Mark the view active and open it
        SetFlag(VF_ACTIVE);

        Verify(OpenView());

        // Initialize the display tree
        _pDispRoot->SetOwned();
        _pDispRoot->SetLayerType(DISPNODELAYER_FLOW);
        _pDispRoot->SetBackground(TRUE);
        _pDispRoot->SetOpaque(TRUE);

        RefreshSettings();

        //  Enable the recalc engine
        //  NOTE: This is a workaround that stops recalc from running until
        //        the view is around.  When the OM is robust enough to handle
        //        no view documents then this should be removed.
        /*_pDoc->suspendRecalc(FALSE); wlw note*/
    }

    RRETURN(_pDispRoot?S_OK:E_FAIL);
}

//+---------------------------------------------------------------------------
//
//  Member:     Deactivate
//
//  Synopsis:   Deactivate a view
//
//----------------------------------------------------------------------------
void CView::Deactivate()
{
    Trace0("View : Deactivate\n");

    Unload();

    if(_pDispRoot)
    {
        Assert(_pDispRoot->GetObserver() == (IDispObserver*)(&_client));

        _pDispRoot->Destroy();
        _pDispRoot = NULL;
    }

    _grfFlags = 0;

    Assert(!_grfLocks);
    Assert(_client._cRefs == 0);
}

//+---------------------------------------------------------------------------
//
//  Member:     Unload
//
//  Synopsis:   Unload a view
//
//----------------------------------------------------------------------------
void CView::Unload()
{
    Assert(!IsActive() || _pDispRoot);

    Trace0("View : Unload");

    if(IsActive())
    {
        EndDeferSetWindowPos(0, TRUE);
        EndDeferSetWindowRgn(0, TRUE);
        EndDeferSetObjectRects(0, TRUE);

        EnsureDisplayTreeIsOpen();

        DeleteAdorners();
        _aryTaskMisc.DeleteAll();
        _aryTaskLayout.DeleteAll();
        _aryTaskEvent.DeleteAll();

        if(_pDispRoot)
        {
            CDispNode* pDispNode;

            Assert(_pDispRoot->CountChildren() <= 1);

            pDispNode = _pDispRoot->GetFirstChildNode();

            if(pDispNode)
            {
                pDispNode->ExtractFromTree();
            }
        }

        ClearRanges();
        ClearInvalid();
        CloseView(LAYOUT_SYNCHRONOUS);
    }

    Assert(!HasTasks());
    Assert(!HasInvalid());

    _pLayout = NULL;
    _grfFlags &= VF_ACTIVE|VF_TREEOPEN;
    _grfLayout = 0;
}

//+---------------------------------------------------------------------------
//
//  Member:     EnsureView
//
//  Synopsis:   Ensure the view is ready for rendering
//
//  Arguments:  grfLayout - Collection of LAYOUT_xxxx flags
//
//  Returns:    TRUE if processing completed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::EnsureView(DWORD grfLayout)
{
    Assert(IsInitialized());

    if(IsActive())
    {
		Assert(!IsLockSet(VL_ENSUREINPROGRESS)
			|| (IsLockSet(VL_UPDATEINPROGRESS)&&(grfLayout&LAYOUT_INPAINT)));
		Assert(!IsLockSet(VL_RENDERINPROGRESS));

        Trace0("View : EnsureView - Enter\n");

        // Note if recursively entered and return
        // (If the recursion occured from a WM_PAINT generated by an earlier call
        // to EnsureView, return success (so rendering continues).
        // All other recursion fails (WM_PAINT processing can recover from failure).)
        if(IsLockSet(VL_ENSUREINPROGRESS) || IsLockSet(VL_RENDERINPROGRESS))
        {
            BOOL fReturn;

            fReturn =
                ((grfLayout&LAYOUT_INPAINT)&&IsLockSet(VL_UPDATEINPROGRESS)&&!IsLockSet(VL_ACCUMULATINGINVALID) ? TRUE : FALSE);

			Trace1("View : EnsureView  - Exit(%S), Skipped processing\n", fReturn?_T("TRUE"):_T("FALSE"));

            return fReturn;
        }

        // Add flags left by asynchronous requests
        Assert(!(grfLayout & LAYOUT_TASKFLAGS));
        Assert(!(_grfLayout & LAYOUT_TASKFLAGS));

        grfLayout |= _grfLayout;
        _grfLayout = 0;

        // If processing synchronously, delete any outstanding asynchronous requests
        // (Asynchronous requests remove the posted message when dispatched)
        CloseView(grfLayout);
        grfLayout &= ~LAYOUT_SYNCHRONOUS;

        // If there is no work to do, exit immediately
        BOOL fSizingNeeded = !IsSized(GetRootLayout());

		if(!fSizingNeeded
			&& !(_grfFlags&(VF_TREEOPEN|VF_NEEDSRECALC|VF_FORCEPAINT))
			&& !HasTasks()
			&& Doc()->_aryPendingFilterElements.Size()==0
			&& !HasInvalid()
			&& !(grfLayout&LAYOUT_SYNCHRONOUSPAINT)
			&& !(grfLayout&LAYOUT_FORCE))
        {
			Assert((grfLayout&~(LAYOUT_NOBACKGROUND
				| LAYOUT_INPAINT
				| LAYOUT_DEFEREVENTS
				| LAYOUT_DEFERENDDEFER
				| LAYOUT_DEFERINVAL
				| LAYOUT_DEFERPAINT)) == 0);
			Assert(!IsDisplayTreeOpen());

            Trace0("View : EnsureView - Exit(TRUE)\n");

            return TRUE;
        }

        // Ensure the view
        {
            CView::CLock lockEnsure(this, VL_ENSUREINPROGRESS);

            // Update the view (if requested)
            {
                CView::CLock lockUpdate(this, VL_UPDATEINPROGRESS);

                // Perform layout and accumulate invalid rectangles/region
                {
                    CView::CLock lockInvalid(this, VL_ACCUMULATINGINVALID);

                    // Process tasks in the following order:
                    //     1) Ensure the display tree is open (so it can be changed)
                    //     2) Transition any waiting objects
                    //     3) Execute pending recalc tasks (which could post events and layout tasks)
                    //     4) Execute pending asynchronous events (which could post layout tasks)
                    //     5) Ensure focus is up-to-date (which may post layout tasks)
                    //     6) Ensure the root layout is correctly sized (which may override pending layout tasks)
                    //     7) Execute pending measure/positioning layout tasks (which can size and move adorners)
                    //     8) Update adorners
                    //     9) Execute pending adorner layout tasks
                    //    10) Close the display tree (which may generate invalid regions and deferred requests)
                    //
                    // BUGBUG: This routine needs to first process foreground tasks (and it should process all foreground
                    //         tasks regardless how much time has passed?) and then background tasks. If too much time
                    //         has passed and tasks remain, it should then post a background view closure to complete the
                    //         work. (brendand)
                    {
                        CView::CLock lockLayout(this, VL_TASKSINPROGRESS);

                        EnsureDisplayTreeIsOpen();

                        if(!IsLockSet(VL_TASKCANRUNSCRIPT))
                        {
                            CView::CLock lockRecalc(this, VL_TASKCANRUNSCRIPT);

                            ExecuteEventTasks(grfLayout);
                        }

                        _pDoc->ExecuteFilterTasks();

                        EnsureFocus();

                        EnsureSize(grfLayout);
                        grfLayout &= ~LAYOUT_FORCE;

                        ExecuteLayoutTasks(grfLayout|LAYOUT_MEASURE);

                        ExecuteLayoutTasks(grfLayout|LAYOUT_POSITION);

                        UpdateAdorners(grfLayout);

                        ExecuteLayoutTasks(grfLayout|LAYOUT_ADORNERS);

#ifdef _DEBUG
                        // Make sure that no tasks were left lying around
                        for(int i=0; i<_aryTaskLayout.Size(); i++)
                        {
                            CViewTask vt = _aryTaskLayout[i];
                            Assert(vt.IsFlagSet(LAYOUT_TASKDELETED));
                        }
#endif

                        _aryTaskLayout.DeleteAll();

                        CloseDisplayTree();
                        Assert(!IsDisplayTreeOpen());
                    }

                    // Process remaining deferred requests and, if necessary, adjust HWND z-ordering
                    EndDeferSetWindowPos(grfLayout);

                    EndDeferSetWindowRgn(grfLayout);

                    EndDeferSetObjectRects(grfLayout);

                    if(IsFlagSet(VF_DIRTYZORDER))
                    {
                        ClearFlag(VF_DIRTYZORDER);
                        FixWindowZOrder();
                    }
                }

                // Publish the accumulated invalid rectangles/region
                PublishInvalid(grfLayout);

                // If requested and not in WM_PAINT handling, render the view (by forcing a WM_PAINT)
				if(!(grfLayout&(LAYOUT_INPAINT|LAYOUT_DEFERPAINT))
					&& (IsFlagSet(VF_FORCEPAINT) || grfLayout&LAYOUT_SYNCHRONOUSPAINT))
				{
					Trace0("View : EnsureView  - Calling UpdateForm\n");

                    Doc()->UpdateForm();
                }
                ClearFlag(VF_FORCEPAINT);
            }

            // Update the caret position
            if(GetRootLayout() && (fSizingNeeded || HasDirtyRange()))
            {
                CCaret* pCaret = Doc()->_pCaret;

                if(pCaret)
                {
                    LONG cp = pCaret->GetCp();

                    if(fSizingNeeded
                        || (cp>=_cpStartMeasured && cp<=_cpEndMeasured)
                        || (cp>=_cpStartTranslated && cp<=_cpEndTranslated))
                    {
                        pCaret->UpdateCaret(FALSE, FALSE);            
                    }
                }
            }

            ClearRanges();
        }
    }

    Trace0("View : EnsureView - Exit(TRUE)\n");

    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetViewPosition, SetViewPosition
//              
//  Synopsis:   Get/Set the position at which display tree content will be rendered
//              in its host coordinates (used for device offsets when printing).
//              
//  Arguments:  pt      point in host coordinates
//              
//----------------------------------------------------------------------------
void CView::GetViewPosition(CPoint* ppt)
{
    Assert(ppt);

    *ppt = _pDispRoot ? _pDispRoot->GetRootPosition() : _afxGlobalData._Zero.pt;
}

void CView::SetViewPosition(const POINT& pt)
{
    if(IsActive())
    {
        OpenView();
        _pDispRoot->SetRootPosition(pt);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     SetViewOffset
//              
//  Synopsis:   Set the view offset, which shifts displayed content (used by
//              printing to display a series of pages with content effectively
//              scrolled between pages).
//              
//  Arguments:  sizeOffset      offset where positive values display content
//                              farther to the right and bottom
//              
//----------------------------------------------------------------------------
BOOL CView::SetViewOffset(const SIZE& sizeOffset)
{
    if(IsActive())
    {
        OpenView();
        return _pDispRoot->SetContentOffset(sizeOffset);
    }
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     EraseBackground
//              
//  Synopsis:   Draw background and border
//              
//  Arguments:  pDI      - Draw context
//              hrgnDraw - Region to draw (if not NULL)
//              prcDraw  - Rect to draw (if not NULL)
//              
//----------------------------------------------------------------------------
void CView::EraseBackground(
    CFormDrawInfo*	pDI,
    HRGN			hrgnDraw,
    const RECT*		prcDraw)
{
    if(IsActive())
    {
        Assert(!IsLockSet(VL_RENDERINPROGRESS));
        Assert(hrgnDraw || prcDraw);

#ifdef _DEBUG
        if(hrgnDraw || prcDraw)
        {
            CRegion rgn(prcDraw, hrgnDraw);

			Trace((_T("View : EraseBackground - HDC(0x%x), HRGN(0x%x), rc(%d, %d, %d, %d)\n"),
				pDI->_hdc,
				rgn.GetRegionAlias(),
				rgn.GetBounds().left,
				rgn.GetBounds().top,
				rgn.GetBounds().right,
				rgn.GetBounds().bottom));
		}
        else
        {
			Trace1("View : EraseBackground - HDC(0x%x), No HRGN or rectangle was passed!\n", pDI->_hdc);
        }
#endif

        Assert(pDI->_hdc != NULL);
        
        if(pDI->_hdc != NULL)
        {
            CView::CLock lock(this, VL_RENDERINPROGRESS);
    
            CPaintCaret         hc(_pDoc->_pCaret);
            CDispDrawContext*   pDispContext = _pDispRoot->GetDrawContext();
            HDC                 hdc = pDI->_hdc;
            POINT               ptOrg = _afxGlobalData._Zero.pt;
    
            _pDispRoot->SetDestination(hdc, NULL);
            pDI->_hdc = NULL;

            ::GetViewportOrgEx(hdc, &ptOrg);

            _pDispRoot->EraseBackground(pDispContext, (void*)pDI, hrgnDraw, prcDraw);

            ::SetViewportOrgEx(hdc, ptOrg.x, ptOrg.y, NULL);
    
            pDI->_hdc = hdc;
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RenderBackground
//
//  Synopsis:   Render the background of the view
//
//  Arguments:  pDI - Current CFormDrawInfo
//
//----------------------------------------------------------------------------
void CView::RenderBackground(CFormDrawInfo* pDI)
{
    if(IsActive())
    {
        Assert(pDI);
        Assert(pDI->GetDC());

        PatBltBrush(pDI->GetDC(), &pDI->_rcClip, PATCOPY, Doc()->_pOptionSettings->crBack());
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RenderElement
//
//  Synopsis:   Render the element (and all its children) onto the passed HDC
//
//  Arguments:  pElement - the element to render
//              hdc - HDC on to which to render the element
//
//  Notes:      For now this will assume no clip region and draw at 0, 0
//              Later on we may want to allow the caller to specify an offset
//              and a clip rect/region.
//
//----------------------------------------------------------------------------
void CView::RenderElement(CElement* pElement, HDC hdc)
{
    Assert(pElement != NULL);
    Assert(hdc != NULL);
    
    if(IsActive() && hdc!=NULL)
    {
        CLayout* pLayout = pElement->GetUpdatedNearestLayout(); // find the appropriate layout

		Trace3("View : RenderElement - HDC(0x%x), Element(0x%x, %S)\n",
			hdc,
			pElement,
			pElement->TagName());

		if(!IsSized(GetRootLayout())
			|| (_grfFlags&(VF_TREEOPEN|VF_NEEDSRECALC))
			|| HasTasks() || (_grfLayout&LAYOUT_FORCE))
		{
            if(IsLockSet(VL_ACCUMULATINGINVALID) || IsLockSet(VL_RENDERINPROGRESS))
            {
                pElement->Invalidate();
                return;
            }
            
            EnsureView(LAYOUT_SYNCHRONOUS|LAYOUT_DEFEREVENTS|LAYOUT_DEFERPAINT);
        }

        CView::CLock lock(this, VL_RENDERINPROGRESS);

        CDispSurface* pSurface = new CDispSurface(hdc);
        CPoint pt(0, 0);
        CFormDrawInfo DI;
        CDispNode* pDispNode = pLayout->GetElementDispNode();

        DI.Init(pElement, hdc);
        DI._hdc = NULL; // do not mess with this until michaelw has dealt with brendan in a kind manner
        // BUGBUG (michaelw) We should probably pull the current clip region from the DC and pass it in.
        // (donmarsh) -- No, filters need to process these pixels unclipped!
        _pDispRoot->DrawNode(pDispNode, pSurface, pt, NULL, &DI);

        delete pSurface;
    }
}

#ifdef _DEBUG
//+------------------------------------------------------------------------
//
//  Function:   DumpRegion
//
//  Synopsis:   Write region to debug output
//
//-------------------------------------------------------------------------
void DumpRegion(HRGN hrgn)
{
	struct
	{
		RGNDATAHEADER rdh;
		RECT arc[128];
	} data;

	if(::GetRegionData(hrgn, ARRAYSIZE(data.arc), (RGNDATA*)&data) < 0)
	{
		Trace1("HRGN(0x%x): Buffer too small\n", hrgn);
	}
	else
	{
		Trace((_T("HRGN(0x%x) iType(%d) nCount(%d) nRgnSize(%d) rcBound(%d,%d,%d,%d)\n"),
			hrgn,
			data.rdh.iType,
			data.rdh.nCount,
			data.rdh.nRgnSize,
			data.rdh.rcBound.left,
			data.rdh.rcBound.top,
			data.rdh.rcBound.right,
			data.rdh.rcBound.bottom));

		for(DWORD i=0; i<data.rdh.nCount; i++)
		{
			Trace((_T("    rc(%d,%d,%d,%d)\n"),
				data.arc[i].left,
				data.arc[i].top,
				data.arc[i].right,
				data.arc[i].bottom));
		}
	}
}
#endif

//+---------------------------------------------------------------------------
//
//  Member:     RenderView
//
//  Synopsis:   Render the view onto the passed HDC
//              
//  Arguments:  pDI      - Draw context
//              hrgnDraw - Region to draw (if not NULL)
//              prcDraw  - Rect to draw (if not NULL)
//
//----------------------------------------------------------------------------
void CView::RenderView(
    CFormDrawInfo*	pDI,
    HRGN			hrgnDraw,
    const RECT*		prcDraw)
{
    if(IsActive())
    {
        Assert(!IsLockSet(VL_ACCUMULATINGINVALID));
        Assert(!IsLockSet(VL_RENDERINPROGRESS));
        Assert(hrgnDraw || prcDraw);

        Trace0("View : RenderView - Enter\n");

#ifdef _DEBUG
        if(hrgnDraw || prcDraw)
        {
            CRegion rgn(prcDraw, hrgnDraw);

			Trace((_T("View : RenderView - HDC(0x%x), HRGN(0x%x), rc(%d, %d, %d, %d)\n"),
				pDI->_hdc,
				rgn.GetRegionAlias(),
				rgn.GetBounds().left,
				rgn.GetBounds().top,
				rgn.GetBounds().right,
				rgn.GetBounds().bottom));
        }
        else
        {
			Trace1("View : RenderView - HDC(0x%x), No HRGN or rectangle was passed!\n", pDI->_hdc);
        }
#endif

        if(IsLockSet(VL_ACCUMULATINGINVALID) || IsLockSet(VL_RENDERINPROGRESS))
        {
            if(IsLockSet(VL_ACCUMULATINGINVALID))
            {
                SetFlag(VF_FORCEPAINT);
            }
        }
        else if(pDI->_hdc != NULL)
        {
            CPaintCaret pc(_pDoc->_pCaret);
            CView::CLock lock(this, VL_RENDERINPROGRESS);
            
            CDispDrawContext* pDispContext = _pDispRoot->GetDrawContext();
            HDC hdc = pDI->_hdc;

            _pDispRoot->SetDestination(hdc, NULL);
            pDI->_hdc = NULL;

#ifdef _DEBUG
            {
                HRGN hrgn = ::CreateRectRgn(0, 0, 0, 0);
                RECT rc;

                if(hrgn)
                {
                    switch(::GetClipRgn(hdc, hrgn))
                    {
                    case 1:
                        Trace0("View: RenderView - ******** Start HDC(0x%x) clipping ********\n");
                        DumpRegion(hrgn);
                        Trace0("View: RenderView - ******** After HDC(0x%x) clipping ********\n");
                        break;

                    case 0:
                        switch(::GetClipBox(hdc, &rc))
                        {
                        case SIMPLEREGION:
							Trace((_T("View: RenderView - HDC(0x%x) has SIMPLE clip region rc(%d,%d,%d,%d)\n"),
								hdc,
								rc.left,
								rc.top,
								rc.right,
								rc.bottom));
                            break;

                        case COMPLEXREGION:
							Trace((_T("View: RenderView - HDC(0x%x) has COMPLEX clip region rc(%d,%d,%d,%d)\n"),
								hdc,
								rc.left,
								rc.top,
								rc.right,
								rc.bottom));
                            break;

                        default:
                            Trace1("View: RenderView - HDC(0x%x) has NO clip region\n", hdc);
                            break;
                        }
                        break;

                    default:
                        Trace1("View: RenderView - HDC(0x%x) Aaack! Error obtaining clip region\n", hdc);
                        break;
                    }

                    ::DeleteObject(hrgn);
                }
            }
#endif

            // Update the buffer size
            // (We need to set this information each time in case something has changed.
            // Fortunately, it won't do any real work until it absolutely has to)
			_pDispRoot->SetOffscreenBufferInfo(
				_pDoc->GetPalette(),
				_pDoc->_bufferDepth,
				(_pDoc->_cSurface>0),
				(_pDoc->_c3DSurface>0),
				WantOffscreenBuffer(),
				AllowOffscreenBuffer());

            _pDispRoot->DrawRoot(pDispContext, (void*)pDI, hrgnDraw, prcDraw);
            pDI->_hdc = hdc;

            SetFlag(VF_HASRENDERED);
        }

        Trace0("View : RenderView - Exit\n");
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RefreshSettings
//
//  Synopsis:   Refresh any view state derived from external states
//              (e.g., registry settings)
//
//----------------------------------------------------------------------------
void CView::RefreshSettings()
{
    if(IsActive())
    {
        Assert(_pDispRoot);

		_pDispRoot->SetCanSmoothScroll(AllowSmoothScrolling());
		_pDispRoot->SetCanScrollDC(AllowScrolling());
		_pDispRoot->SetOffscreenBufferInfo(
			_pDoc->GetPalette(), 
			_pDoc->_bufferDepth, 
			(_pDoc->_cSurface>0), 
			(_pDoc->_c3DSurface>0), 
			WantOffscreenBuffer(),
			AllowOffscreenBuffer());
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     HitScrollInset
//
//  Synopsis:   Find the top-most scroller that contains the given hit point
//              near its edge, and is capable of scrolling in that direction.
//
//  Arguments:  pptHit       - Point to test
//              pdwScrollDir - Returns the direction(s) that this scroller
//                             can scroll in
//
//  Returns:    CDispScroller that can scroll, NULL otherwise
//
//----------------------------------------------------------------------------
CDispScroller* CView::HitScrollInset(CPoint* pptHit, DWORD* pdwScrollDir)
{
    return IsActive()?_pDispRoot->HitScrollInset(pptHit, pdwScrollDir):FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestPoint
//
//  Synopsis:   Initiate a hit-test through the display tree
//
//  Arguments:  pMessage   - CMessage for which to hit test
//              ppTreeNode - Where to return CTreeNode of hit element
//              grfFlags   - HT_xxxx flags
//
//  Returns:    HTC_xxx value
//
//----------------------------------------------------------------------------
HTC CView::HitTestPoint(CMessage* pMessage, CTreeNode** ppTreeNode, DWORD grfFlags)
{
    if(!IsActive())
	{
		return HTC_NO;
	}

	CDispNode* pDispNode;
	POINT ptContent = _afxGlobalData._Zero.pt;
	HTC htc = HitTestPoint(
		pMessage->pt,
		COORDSYS_CONTAINER,
		NULL,
		grfFlags,
		&pMessage->resultsHitTest,
		ppTreeNode,
		ptContent,
		&pDispNode);

    // Save content coordinate point and associated CDispNode
    pMessage->SetContentPoint(ptContent, pDispNode);

    return htc;
}

//+---------------------------------------------------------------------------
//
//  Member:     HitTestPoint
//
//  Synopsis:   Initiate a hit-test through the display tree
//
//  Arguments:  pt         - POINT to hit test
//              cs         - Coordinate system for the point
//              pElement   - The CElement start hittesting at (if NULL, start at the root)
//              grfFlags   - HT_xxxx flags
//              phtr       - Pointer to a HITTESTRESULTS
//              ppTreeNode - Where to return CTreeNode of hit element
//              ptContent  - Hit tested point in content coordinates
//              ppDispNode - Display node that was hit
//
//  Returns:    HTC_xxx value
//
//----------------------------------------------------------------------------
HTC CView::HitTestPoint(
    const POINT&		pt,
    COORDINATE_SYSTEM	cs,
    CElement*			pElement,
    DWORD				grfFlags,
    HITTESTRESULTS*		pHTRslts,
    CTreeNode**			ppTreeNode,
    POINT&				ptContent,
    CDispNode**			ppDispNode)
{
    Assert(ppTreeNode);
    Assert(!IsLockSet(VL_ACCUMULATINGINVALID));

    if(!IsActive() || IsLockSet(VL_ACCUMULATINGINVALID))
	{
		return HTC_NO;
	}

    CPoint			ptHit(pt);
    CHitTestInfo	hti;
    CDispNode*		pDispNode = NULL;

    // BUGBUG: The fuzzy border should be used whenever the container of is
    //         in edit-mode. For now, only use it when the document itself
    //         is in design-mode. (donmarsh)
    long cFuzzyBorder = 0;

    // Ensure the view is up-to-date so the hit-test is accurate
	if(!IsSized(GetRootLayout())
		|| (_grfFlags&(VF_TREEOPEN|VF_NEEDSRECALC))
		|| HasTasks() || (_grfLayout&LAYOUT_FORCE))
    {
        if(IsLockSet(VL_ENSUREINPROGRESS) || IsLockSet(VL_RENDERINPROGRESS))
		{
			return HTC_NO;
		}

        EnsureView(LAYOUT_DEFEREVENTS|LAYOUT_DEFERPAINT);
    }

    // Construct the default CHitTestInfo
    hti._htc			= HTC_NO;
    hti._phtr			= pHTRslts;
    hti._pNodeElement	= *ppTreeNode;
    hti._pDispNode		= NULL;
    hti._ptContent		= _afxGlobalData._Zero.pt;
    hti._grfFlags		= grfFlags;

    // Determine the starting display node
    // NOTE: This must be obtained after calling EnsureView since pending layout
    //       can change the display nodes associated with a element
    if(pElement)
    {
        CLayout* pLayout = pElement->GetUpdatedNearestLayout();

        if(pLayout)
        {
            pDispNode = pLayout->GetElementDispNode(pElement);
        }
    }

    if(!pDispNode)
    {
        pDispNode = _pDispRoot;
    }

    // Find the hit
    pDispNode->HitTest(&ptHit, cs, &hti, !!(grfFlags&HT_VIRTUALHITTEST), cFuzzyBorder);
    
    // Save content coordinate point and associated CDispNode
    if(!hti._pNodeElement)
    {
        Assert(!hti._pDispNode);

        CElement* pElement = _pDoc->GetPrimaryElementClient();

        if(pElement)
        {
            hti._htc = HTC_YES;
            hti._pNodeElement = pElement->GetFirstBranch();

            if(hti._phtr)
            {
                hti._phtr->_fWantArrow = TRUE;
            }
        }
    }

    ptContent	= hti._ptContent;
    *ppDispNode	= hti._pDispNode;
    *ppTreeNode	= hti._pNodeElement;

    return hti._htc;
}

//+---------------------------------------------------------------------------
//
//  Member:     Invalidate
//
//  Synopsis:   Invalidate the view
//
//  Arguments:  prcInvalid         - Invalid rectangle
//              rgn                - Invalid region
//              fSynchronousRedraw - TRUE to redraw synchronously
//              fInvalChildWindows - TRUE to invalidate child windows
//
//----------------------------------------------------------------------------
void CView::Invalidate(
    const CRect*	prcInvalid,
    BOOL			fSynchronousRedraw,
    BOOL			fInvalChildWindows,
    BOOL			fPostRender)
{
    if(IsActive())
    {
        if(!prcInvalid)
        {
            prcInvalid = (CRect*)&_pDispRoot->GetBounds();
        }

		Trace((_T("View : Invalidate - rc(%d, %d, %d, %d)\n"),
			prcInvalid->left,
			prcInvalid->top,
			prcInvalid->right,
			prcInvalid->bottom));

        // Maintain a small number of invalid rectangles
        if(_cInvalidRects < MAX_INVALID)
        {
            // Ignore the rectangle if contained within another
			int i;
            for(i=0; i<_cInvalidRects; i++)
            {
                if(_aryInvalidRects[i].Contains(*prcInvalid))
				{
					break;
				}
            }

            if(i >= _cInvalidRects)
            {
                _aryInvalidRects[_cInvalidRects++] = *prcInvalid;
            }
        }
        // If too many arrive, union the rectangle into that which results in the least growth
        else
        {
            CRect	rc;
            long	i, iBest;
            long	c, cBest;

            iBest = 0;
            cBest = MINLONG;

            for(i=0; i<MAX_INVALID; i++)
            {
                rc = _aryInvalidRects[i];
                c  = rc.FastArea();

                rc.Union(*prcInvalid);
                c -= rc.FastArea();

                if(c > cBest)
                {
                    iBest = i;
                    cBest = c;

                    if(!cBest)
					{
						break;
					}
                }
            }

            Assert(iBest>=0 && iBest<MAX_INVALID);

            if(cBest)
            {
                _aryInvalidRects[iBest].Union(*prcInvalid);
            }
        }

        // Note if child HWNDs need invalidation
        if(fInvalChildWindows)
        {
            SetFlag(VF_INVALCHILDWINDOWS);
        }

        // If not actively accumulating invalid rectangles/region,
        // ensure that the view will eventually render
        if(fPostRender)
		{
			PostRenderView(fSynchronousRedraw);
		}
    }
}

void CView::Invalidate(
    const CRegion&	rgn,
    BOOL			fSynchronousRedraw,
    BOOL			fInvalChildWindows)
{
    if(IsActive())
    {
        // If the region is a rectangle, forward and return
        if(!rgn.IsRegion())
        {
            Invalidate(&rgn.AsRect(), fSynchronousRedraw, fInvalChildWindows);
            return;
        }

        Trace((_T("View : Invalidate - HRGN(0x%x, (%d, %d, %d, %d))\n"),
               rgn.GetRegionAlias(),
               rgn.GetBounds().left,
               rgn.GetBounds().top,
               rgn.GetBounds().right,
               rgn.GetBounds().bottom));

        // Collect the invalid region
        _rgnInvalid.Union(rgn);

        // Note if child HWNDs need invalidation
        if(fInvalChildWindows)
        {
            SetFlag(VF_INVALCHILDWINDOWS);
        }

        // If not actively accumulating invalid rectangles/region,
        // ensure that the view will eventually render
        PostRenderView(fSynchronousRedraw);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     InvalidateBorder
//
//  Synopsis:   Invalidate a border around the edge of the view
//
//  Arguments:  cBorder - Border width to invalidate
//
//-----------------------------------------------------------------------------
void CView::InvalidateBorder(long cBorder)
{
    if(IsActive() && cBorder)
    {
        CSize size;

        GetViewSize(&size);

        CRect rc(size);
        CRect rcBorder;

        rcBorder = rc;
        rcBorder.right = rcBorder.left + cBorder;
        Invalidate(&rcBorder);

        rcBorder = rc;
        rcBorder.bottom = rcBorder.top + cBorder;
        Invalidate(&rcBorder);

        rcBorder = rc;
        rcBorder.left = rcBorder.right - cBorder;
        Invalidate(&rcBorder);

        rcBorder = rc;
        rcBorder.top = rcBorder.bottom - cBorder;
        Invalidate(&rcBorder);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     Notify
//
//  Synopsis:   Respond to a notification
//
//  Arguments:  pnf - Notification sent
//
//-----------------------------------------------------------------------------
void CView::Notify(CNotification* pnf)
{
    if(IsActive())
    {
        switch(pnf->Type())
        {
        case NTYPE_MEASURED_RANGE:
            AccumulateMeasuredRange(pnf->Cp(0), pnf->Cch(LONG_MAX));
            break;

        case NTYPE_TRANSLATED_RANGE:
            AccumulateTranslatedRange(pnf->DataAsSize(), pnf->Cp(0), pnf->Cch(LONG_MAX));
            break;

        case NTYPE_ELEMENT_SIZECHANGED:
        case NTYPE_ELEMENT_POSITIONCHANGED:
            if(HasAdorners())
            {
                long iAdorner;
                BOOL fShapeChange = pnf->IsType(NTYPE_ELEMENT_SIZECHANGED);

                Assert(!pnf->Element()->IsPositionStatic());

				for(iAdorner=GetAdorner(pnf->Element());
					iAdorner>=0; iAdorner=GetAdorner(pnf->Element(), iAdorner+1))
				{
                    if(fShapeChange)
                    {
                        _aryAdorners[iAdorner]->ShapeChanged();
                    }
                    else
                    {
                        _aryAdorners[iAdorner]->PositionChanged();
                    }
                }
            }
            break;

        case NTYPE_ELEMENT_ENSURERECALC:
            EnsureSize(_grfLayout);
            _grfLayout &= ~LAYOUT_FORCE;
            break;

        case NTYPE_ELEMENT_RESIZE:
        case NTYPE_ELEMENT_RESIZEANDREMEASURE:
            Verify(OpenView());
            if(pnf->Element() == Doc()->GetPrimaryElementClient())
            {
                pnf->Element()->DirtyLayout(pnf->LayoutFlags());
            }
            ClearFlag(VF_SIZED);
            break;

        case NTYPE_DISPLAY_CHANGE:
        case NTYPE_VIEW_ATTACHELEMENT:
        case NTYPE_VIEW_DETACHELEMENT:
            Verify(OpenView());
            ClearFlag(VF_SIZED);
            ClearFlag(VF_ATTACHED);
            break;
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     AddLayoutTask
//
//  Synopsis:   Add a layout view task
//
//  Arguments:  pLayout   - CLayout to invoke
//              grfLayout - Collection of LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
HRESULT CView::AddLayoutTask(CLayout* pLayout, DWORD grfLayout)
{
    Assert(!((grfLayout & ~(LAYOUT_FORCE|LAYOUT_SYNCHRONOUSPAINT))&LAYOUT_NONTASKFLAGS));
    Assert(grfLayout & (LAYOUT_MEASURE|LAYOUT_POSITION|LAYOUT_ADORNERS));

    if(!IsActive())
	{
		return S_OK;
	}

    HRESULT hr = AddTask(pLayout, CViewTask::VTT_LAYOUT, (grfLayout&LAYOUT_TASKFLAGS));

    if(SUCCEEDED(hr))
    {
        grfLayout &= ~LAYOUT_FORCE;
        _grfLayout |= grfLayout&LAYOUT_NONTASKFLAGS;

        if(_grfLayout & LAYOUT_SYNCHRONOUS)
        {
            EnsureView(LAYOUT_SYNCHRONOUS|LAYOUT_SYNCHRONOUSPAINT);
        }
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     RemoveEventTasks
//
//  Synopsis:   Remove all event tasks for the passed element
//
//  Arguments:  pElement - The element whose tasks are to be removed
//
//----------------------------------------------------------------------------
void CView::RemoveEventTasks(CElement*  pElement)
{
    Assert(IsActive() || !_aryTaskEvent.Size());
    Assert(pElement);
    Assert(pElement->_fHasPendingEvent);

    if(IsActive())
    {
        int cTasks = _aryTaskEvent.Size();
        int iTask;

        for(iTask=0; iTask<cTasks;)
        {
            if(_aryTaskEvent[iTask]._pElement == pElement)
            {
                _aryTaskEvent.Delete(iTask);
                cTasks--;
            }
            else
            {
                iTask++;
            }
        }
    }

    pElement->_fHasPendingEvent = FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     RemoveAdorners
//
//  Synopsis:   Remove all adorners associated with an element
//
//  Arguments:  pElement - CElement associated with the adorner
//
//----------------------------------------------------------------------------
void CView::RemoveAdorners(CElement* pElement)
{
    Assert(IsActive() || !_aryAdorners.Size());

    if(IsActive() && _aryAdorners.Size())
    {
        long iAdorner;

		for(iAdorner=GetAdorner(pElement); iAdorner>=0; iAdorner=GetAdorner(pElement, iAdorner))
		{
            RemoveAdorner(_aryAdorners[iAdorner], FALSE);
        }

        if(pElement->HasLayout())
        {
            pElement->GetCurLayout()->SetIsAdorned(FALSE);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     EndDeferred
//
//  Synopsis:   End all deferred requests
//              NOTE: This routine is meant only for callers outside of CView.
//                    CView should call the underlying routines directly.
//
//----------------------------------------------------------------------------
void CView::EndDeferred()
{
    Assert(!IsLockSet(VL_ENSUREINPROGRESS));
    Assert(!IsLockSet(VL_RENDERINPROGRESS));

    EndDeferSetWindowPos();
    EndDeferSetWindowRgn();
    EndDeferSetObjectRects();

    if(IsFlagSet(VF_DIRTYZORDER))
    {
        ClearFlag(VF_DIRTYZORDER);
        FixWindowZOrder();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     DeferSetObjectRects, EndDeferSetObjectRects, SetObjectRects
//
//  Synopsis:   Collect, defer, and execute SetObjectRects requests
//
//----------------------------------------------------------------------------
void CView::DeferSetObjectRects(
    IOleInPlaceObject*	pInPlaceObject,
    RECT*				prcObj,
    RECT*				prcClip,
    HWND				hwnd,
    BOOL				fInvalidate)
{
    Assert(IsActive() || !_arySor.Size());

    if(!IsActive() || pInPlaceObject==NULL)
	{
		return;
	}

    Trace((_T("defer SOR %x Client: %ld %ld %ld %ld  Clip: %ld %ld %ld %ld  inv: %d\n"),
                    hwnd,
                    prcObj->top, prcObj->bottom, prcObj->left, prcObj->right,
                    prcClip->top, prcClip->bottom, prcClip->left, prcClip->right,
                    fInvalidate));

    // Scan array for matching entry
    // If one is found, update it; Otherwise, append a new entry
    int i;
    SOR* psor;
    for(i=_arySor.Size(),psor=&(_arySor[0]); i>0; i--,psor++)
    {
        if(psor->hwnd==hwnd && psor->pInPlaceObject==pInPlaceObject)
        {
            psor->rc              = *prcObj;
            psor->rcClip          = *prcClip;
            psor->fInvalidate     = psor->fInvalidate || fInvalidate;
            break;
        }
    }

    // No match was found, append a new entry
    if(i <= 0)
    {
        if(_arySor.EnsureSize(_arySor.Size()+1) == S_OK)
        {
            psor = &_arySor[_arySor.Size()];
            psor->pInPlaceObject  = pInPlaceObject;
            pInPlaceObject->AddRef();
            psor->rc              = *prcObj;
            psor->rcClip          = *prcClip;
            psor->hwnd            = hwnd;
            psor->fInvalidate     = fInvalidate;
            _arySor.SetSize(_arySor.Size()+1);
        }
        else
        {
            CServer::CLock Lock(_pDoc, SERVERLOCK_BLOCKPAINT|SERVERLOCK_IGNOREERASEBKGND);
            SetObjectRectsHelper(pInPlaceObject, prcObj, prcClip, hwnd, fInvalidate);
        }
    }
}

void CView::EndDeferSetObjectRects(DWORD grfLayout, BOOL fIgnore)
{
    SOR* psor;
    int i;

    Assert(fIgnore || _pDispRoot);
    Assert(IsActive() || !_arySor.Size());

    if(grfLayout & LAYOUT_DEFERENDDEFER)
	{
		return;
	}

    if(!IsActive() || !_arySor.Size())
	{
		return;
	}

    {
        CServer::CLock Lock(_pDoc, SERVERLOCK_IGNOREERASEBKGND);

        for(i=_arySor.Size(),psor=&(_arySor[0]); i>0; i--,psor++)
        {
            if(!fIgnore)
            {
				SetObjectRectsHelper(psor->pInPlaceObject,
					&psor->rc, &psor->rcClip, psor->hwnd, psor->fInvalidate);
            }

            psor->pInPlaceObject->Release();
        }
    }

    _arySor.DeleteAll();
}

// used to enumerate children of a suspected clipping outer window
struct INNERWINDOWTESTSTRUCT
{
    HWND hwndParent;
    CView* pView;
    CRect rc;
};

// this callback checks each of the children of a particular
// outer window.  If the child is positioned in the given place
// (namely the client rect), we'll mark the parent as a clipping
// outer window.  It's called from SetObjectRectsHelper, below.
BOOL CALLBACK TestInnerWindow(HWND hwndChild, LPARAM lparam)
{
    INNERWINDOWTESTSTRUCT* piwts = (INNERWINDOWTESTSTRUCT*)lparam;
    CRect rcActual;

    piwts->pView->GetHWNDRect(hwndChild, &rcActual);
    if(rcActual == piwts->rc)
    {
        piwts->pView->AddClippingOuterWindow(piwts->hwndParent);
        return FALSE;
    }
    return TRUE;
}

void CView::SetObjectRectsHelper(
    IOleInPlaceObject*	pInPlaceObject,
    RECT*				prcObj,
    RECT*				prcClip,
    HWND				hwnd,
    BOOL				fInvalidate)
{
    CRect rcWndBefore, rcWndAfter;
    HRGN hrgnBefore=NULL, hrgnAfter;

    if(hwnd)
    {
        hrgnBefore = ::CreateRectRgnIndirect(&_afxGlobalData._Zero.rc);
        if(hrgnBefore)
		{
			::GetWindowRgn(hwnd, hrgnBefore);
		}
        GetHWNDRect(hwnd, &rcWndBefore);
    }

	Trace((_T("SOR %x Client: %ld %ld %ld %ld  Clip: %ld %ld %ld %ld  inv: %d\n"),
		hwnd,
		prcObj->top, prcObj->bottom, prcObj->left, prcObj->right,
		prcClip->top, prcClip->bottom, prcClip->left, prcClip->right,
		fInvalidate));

	pInPlaceObject->SetObjectRects(prcObj, prcClip);

    // MFC controls change the HWND to the rcClip instead of the rcObj,
    // but they don't change the Window region accordingly.  We try to
    // detect that, and make the right change for them (bug 75218).

    // only do this if it looks like SetObjectRects moved the position
    if(hwnd)
	{
		GetHWNDRect(hwnd, &rcWndAfter);
	}
    if(hwnd && rcWndBefore!=rcWndAfter)
    {
        hrgnAfter = ::CreateRectRgnIndirect(&_afxGlobalData._Zero.rc);
        if(hrgnBefore && hrgnAfter)
        {
            ::GetWindowRgn(hwnd, hrgnAfter);

            // if the position moved but the region didn't, move the region
            if(EqualRgn(hrgnBefore, hrgnAfter))
            {
                CSize sizeOffset = rcWndBefore.TopLeft() - rcWndAfter.TopLeft();
                ::OffsetRgn(hrgnAfter, sizeOffset.cx, sizeOffset.cy);
                ::SetWindowRgn(hwnd, hrgnAfter, FALSE);

				Trace3("move window rgn for %x by %ld %ld\n",
					hwnd, sizeOffset.cy, sizeOffset.cx);
			}
			else
            {
                ::DeleteObject(hrgnAfter);
            }
        }

        // if SetObjectRects moved the position to the clipping rect, we
        // may want to mark the hwnd so that we move it we can move it directly
        // to the clipping rect in the future.
        if(rcWndAfter==*prcClip && rcWndAfter!=*prcObj)
        {
            INNERWINDOWTESTSTRUCT iwts;

            iwts.hwndParent = hwnd;
            iwts.pView = this;
            iwts.rc = *prcObj;
            
            EnumChildWindows(hwnd, TestInnerWindow, (LPARAM)&iwts);
        }
    }

    if(hrgnBefore)
    {
        ::DeleteObject(hrgnBefore);
    }

    if(hwnd && fInvalidate)
    {
        ::InvalidateRect(hwnd, NULL, FALSE);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetHWNDRect, AddClippingOuterWindow, RemoveClippingOuterWindow,
//              IndexOfClippingOuterWindow
//
//  Synopsis:   Some controls (MFC controls, prominently) implement clipping
//              by a pair of windows:  an outer window set to the clip rect
//              and an inner window set to the content rect.  These routines
//              help detect such HWNDs, so that we can move them to their
//              real positions and avoid flicker.
//
//----------------------------------------------------------------------------
void CView::GetHWNDRect(HWND hwnd, CRect* prcWndActual) const
{    
    CInPlace* pInPlace = Doc()->_pInPlace;
    if(pInPlace)
    {
        HWND hwndParent = pInPlace->_hwnd;
        Assert(hwndParent);
        ::GetWindowRect(hwnd, prcWndActual);
        CPoint ptActual(prcWndActual->TopLeft());
        ::ScreenToClient(hwndParent, &ptActual);
        prcWndActual->MoveTo(ptActual); // change to parent window coords.
    }
    else
    {
        *prcWndActual = _afxGlobalData._Zero.rc;
    }
}

void CView::AddClippingOuterWindow(HWND hwnd)
{
    if(!IsClippingOuterWindow(hwnd))
    {
        _aryClippingOuterHwnd.AppendIndirect(&hwnd);
        Trace1("hwnd %x marked as MFC control", hwnd);
    }
}

void CView::RemoveClippingOuterWindow(HWND hwnd)
{
    _aryClippingOuterHwnd.DeleteByValueIndirect(&hwnd);
}

int CView::IndexOfClippingOuterWindow(HWND hwnd)
{
    return _aryClippingOuterHwnd.FindIndirect(&hwnd);
}

//+---------------------------------------------------------------------------
//
//  Member:     DeferSetWindowPos, EndDeferSetWindowPos
//
//  Synopsis:   Collect, defer, and execute SetWindowPos requests
//
//----------------------------------------------------------------------------
void CView::DeferSetWindowPos(
    HWND		hwnd,
    const RECT*	prc,
    UINT		uFlags,
    const RECT*	prcInvalid)
{
    Assert(IsActive() || !_aryWndPos.Size());

    if(!IsActive())
	{
		return;
	}

#ifdef _DEBUG
	if(prcInvalid)
	{
		Trace((_T("defer SWP %x Rect: %ld %ld %ld %ld  Flags: %x  Inval: %ld %ld %ld %ld\n"),
			hwnd,
			prc->top, prc->bottom, prc->left, prc->right,
			uFlags,
			prcInvalid->top, prcInvalid->bottom, prcInvalid->left, prcInvalid->right));
	}
	else
	{
		Trace((_T("defer SWP %x Rect: %ld %ld %ld %ld  Flags: %x  Inval: null\n"),
			hwnd,
			prc->top, prc->bottom, prc->left, prc->right,
			uFlags));
	}
#endif

    if(hwnd)
    {
        WND_POS* pWndPos;
        int i;

        // always no z-order change... this is handled by FixWindowZOrder
        uFlags |= SWP_NOZORDER;

        // Scan array for matching entry
        // If one is found, update it; Otherwise, append a new entry
        for(i=_aryWndPos.Size(),pWndPos=&(_aryWndPos[0]); i>0; i--,pWndPos++)
        {
            if(pWndPos->hwnd == hwnd)
            {
                pWndPos->rc = *prc;

                // Reset flags
                // (Always keep the flags that include either SWP_SHOWWINDOW or SWP_HIDEWINDOW)
                if((uFlags&(SWP_SHOWWINDOW|SWP_HIDEWINDOW)) || !(pWndPos->uFlags&(SWP_SHOWWINDOW|SWP_HIDEWINDOW)))
                {
                    //  NOTE: If ever not set, SWP_NOREDRAW is ignored
                    if((uFlags&SWP_NOREDRAW) && !(pWndPos->uFlags&SWP_NOREDRAW))
                    {
                        uFlags &= ~SWP_NOREDRAW;
                    }

                    pWndPos->uFlags = uFlags;
                }

                break;
            }
        }

        // No match was found, append a new entry
        if(i <= 0)
        {
            pWndPos = _aryWndPos.Append();

            if(pWndPos != NULL)
            {
                pWndPos->hwnd = hwnd;
                pWndPos->rc = *prc;
                pWndPos->uFlags = uFlags;
            }
        }
    }
}

void CView::EndDeferSetWindowPos(DWORD grfLayout, BOOL fIgnore)
{
    Assert(fIgnore || _pDispRoot);
    Assert(IsActive() || !_aryWndPos.Size());

    if(grfLayout & LAYOUT_DEFERENDDEFER)
	{
		return;
	}

    if(!IsActive() || !_aryWndPos.Size())
	{
		return;
	}

    if(!fIgnore)
    {
        CServer::CLock Lock(_pDoc, SERVERLOCK_IGNOREERASEBKGND);

        WND_POS*	pWndPos;
        HDWP		hdwpShowHide = NULL;
        HDWP		hdwpOther    = NULL;
        HDWP		hdwp;
        int			cShowHideWindows = 0;
        int			i;

        // Since positioning requests and SWP_SHOW/HIDEWINDOW requests cannot be safely mixed,
        // process first all positioning requests followed by all SWP_SHOW/HIDEWINDOW requests
        for(i=_aryWndPos.Size(),pWndPos=&(_aryWndPos[0]); i>0; i--,pWndPos++)
        {
            if(pWndPos->uFlags & (SWP_SHOWWINDOW|SWP_HIDEWINDOW))
            {
                cShowHideWindows++;
            }
        }

        // Allocate deferred structures large enough for each set of requests
        if(cShowHideWindows > 0)
        {
            hdwpShowHide = ::BeginDeferWindowPos(cShowHideWindows);
            if(!hdwpShowHide)
			{
				goto Cleanup;
			}
        }

        if(cShowHideWindows < _aryWndPos.Size())
        {
            hdwpOther = ::BeginDeferWindowPos(_aryWndPos.Size()-cShowHideWindows);
            if(!hdwpOther)
			{
				goto Cleanup;
			}
        }

        // Collect and issue the requests
        Assert(cShowHideWindows<=0 || hdwpShowHide!=NULL);
        Assert(cShowHideWindows>=_aryWndPos.Size() || hdwpOther!=NULL);

        for(i=_aryWndPos.Size(),pWndPos=&(_aryWndPos[0]); i>0; i--,pWndPos++)
        {
            // The window cached by DeferSetWindowPos may be destroyed
            // before we come here.
            if(!IsWindow(pWndPos->hwnd))
			{
				continue;
			}

			Trace((_T("SWP %x Rect: %ld %ld %ld %ld  Flags: %x\n"),
				pWndPos->hwnd,
				pWndPos->rc.top, pWndPos->rc.bottom, pWndPos->rc.left, pWndPos->rc.right,
				pWndPos->uFlags));

			if(pWndPos->uFlags & (SWP_SHOWWINDOW|SWP_HIDEWINDOW))
			{
				hdwp = ::DeferWindowPos(hdwpShowHide,
					pWndPos->hwnd,
					NULL,
					pWndPos->rc.left,
					pWndPos->rc.top,
					pWndPos->rc.Width(),
					pWndPos->rc.Height(),
					pWndPos->uFlags);

                if(!hdwp)
				{
					goto Cleanup;
				}

                hdwpShowHide = hdwp;
            }
            else
            {
				hdwp = ::DeferWindowPos(hdwpOther,
					pWndPos->hwnd,
					NULL,
					pWndPos->rc.left,
					pWndPos->rc.top,
					pWndPos->rc.Width(),
					pWndPos->rc.Height(),
					pWndPos->uFlags);

                if(!hdwp)
				{
					goto Cleanup;
				}

                hdwpOther = hdwp;
            }
        }

    Cleanup:
        if(hdwpOther != NULL)
        {
            ::EndDeferWindowPos(hdwpOther);
        }

        if(hdwpShowHide != NULL)
        {
            ::EndDeferWindowPos(hdwpShowHide);
        }

        _aryWndPos.DeleteAll();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     DeferSetWindowRgn, EndDeferWindowRgn
//
//  Synopsis:   Collect, defer, and execute SetWindowRgn requests
//
//----------------------------------------------------------------------------
void CView::DeferSetWindowRgn(
    HWND		hwnd,
    const RECT*	prc,
    BOOL		fRedraw)
{
    Assert(prc);
    Assert(IsActive() || !_aryWndRgn.Size());

    if(!IsActive())
	{
		return;
	}

    Trace((_T("defer SWRgn %x Rect: %ld %ld %ld %ld  Redraw: %d\n"),
            hwnd, prc->top, prc->bottom, prc->left, prc->right, fRedraw));

    int i;
    WND_RGN* pwrgn;

    // Scan array for matching entry
    // If one is found, update it; Otherwise, append a new entry
    for(i=_aryWndRgn.Size(),pwrgn=&(_aryWndRgn[0]); i>0; i--,pwrgn++)
    {
        if(pwrgn->hwnd == hwnd)
        {
            pwrgn->rc = *prc;
            pwrgn->fRedraw = pwrgn->fRedraw || fRedraw;
            break;
        }
    }

    // No match was found, append a new entry
    if(i <= 0)
    {
        if(_aryWndRgn.EnsureSize(_aryWndRgn.Size()+1) == S_OK)
        {
            pwrgn = &_aryWndRgn[_aryWndRgn.Size()];

            pwrgn->hwnd    = hwnd;
            pwrgn->rc      = *prc;
            pwrgn->fRedraw = fRedraw;

            _aryWndRgn.SetSize(_aryWndRgn.Size()+1);
        }
        else
        {
            ::SetWindowRgn(hwnd, ::CreateRectRgnIndirect(prc), fRedraw);

            Trace((_T("SWRgn %x Rect: %ld %ld %ld %ld  Redraw: %d\n"),
                hwnd, prc->top, prc->bottom, prc->left, prc->right, fRedraw));
        }
    }
}

void CView::EndDeferSetWindowRgn(DWORD grfLayout, BOOL fIgnore)
{
    Assert(fIgnore || _pDispRoot);
    Assert(IsActive() || !_aryWndRgn.Size());

    if(grfLayout & LAYOUT_DEFERENDDEFER)
	{
		return;
	}

    if(!IsActive() ||  !_aryWndRgn.Size())
	{
		return;
	}

    if(!fIgnore)
    {
        WND_RGN*    pwrgn;
        int         i;

        for(i=_aryWndRgn.Size(),pwrgn=&(_aryWndRgn[0]); i>0; i--,pwrgn++)
        {
            ::SetWindowRgn(pwrgn->hwnd, ::CreateRectRgnIndirect(&pwrgn->rc), pwrgn->fRedraw);

			Trace((_T("SWRgn %x Rect: %ld %ld %ld %ld  Redraw: %d\n"),
				pwrgn->hwnd,
				pwrgn->rc.top, pwrgn->rc.bottom, pwrgn->rc.left, pwrgn->rc.right,
				pwrgn->fRedraw));
		}
	}

    _aryWndRgn.DeleteAll();
}

//+---------------------------------------------------------------------------
//
//  Member:     SetFocus
//
//  Synopsis:   Set or clear the association between the focus adorner and
//              an element
//
//  Arguments:  pElement  - Element which has focus (may be NULL)
//              iDivision - Subdivision of the element which has focus
//
//----------------------------------------------------------------------------
void CView::SetFocus(CElement* pElement, long iDivision)
{
    if(!IsActive())
	{
		return;
	}

    // BUGBUG (donmarsh) -- we fail if we're in the middle of drawing.  Soon
    // we will have a better mechanism for deferring requests that are made
    // in the middle of rendering.
    if(!OpenView())
	{
		return;
	}

    if(pElement)
    {
        CTreeNode* pNode = pElement->GetFirstBranch();

        if(!pNode || pNode->IsVisibilityHidden() || pNode->IsDisplayNone())
        {
            pElement = NULL;
        }
    }

    if(pElement)
    {
        if(!_pFocusAdorner)
        {
            _pFocusAdorner = new CFocusAdorner(this);

            if(_pFocusAdorner && !SUCCEEDED(AddAdorner(_pFocusAdorner)))
            {
                _pFocusAdorner->Destroy();
                _pFocusAdorner = NULL;
            }
        }

        if(_pFocusAdorner)
        {
            _pFocusAdorner->SetElement(pElement, iDivision);
        }
    }
    else if(_pFocusAdorner)
    {
        RemoveAdorner(_pFocusAdorner, FALSE);
        _pFocusAdorner = NULL;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     InvalidateFocus
//
//  Synopsis:   Invalidate the focus shape (if any)
//
//----------------------------------------------------------------------------
void CView::InvalidateFocus() const
{
    if(IsActive() && _pFocusAdorner)
    {
        _pFocusAdorner->Invalidate();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetViewSize
//
//  Synopsis:   Returns the current size of the view
//
//  Arguments:  psize - Returns the size
//
//----------------------------------------------------------------------------
void CView::GetViewSize(CSize* psize) const
{
    Assert(psize);

    if(IsActive())
    {
        _pDispRoot->GetSize(psize);
    }
    else
    {
        *psize = _afxGlobalData._Zero.size;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     SetViewSize
//
//  Synopsis:   Set the size of the view
//
//  Arguments:  size - New view size
//
//----------------------------------------------------------------------------
void CView::SetViewSize(const CSize& size)
{
    Assert(!IsLockSet(VL_RENDERINPROGRESS));

    if(IsActive())
    {
        CSize sizeCurrent;

        GetViewSize(&sizeCurrent);

        if(size != sizeCurrent)
        {
            if(OpenView())
            {
                // Set the new size
                _pDispRoot->SetRootSize(size, FALSE);
                ClearFlag(VF_SIZED);

                // Queue a resize event
                if(Doc()->_fFiredOnLoad)
                {
                    CLayout* pLayout = GetRootLayout();

                    if(pLayout && !pLayout->IsDisplayNone())
                    {
                       AddEventTask(pLayout->ElementOwner(), DISPID_EVMETH_ONRESIZE);
                    }
                }
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     OpenView
//
//  Synopsis:   Prepare the view for changes (e.g., open the display tree)
//
//  Returns:    TRUE if the view was successfully opened, FALSE if we are in
//              the middle of rendering
//
//----------------------------------------------------------------------------
BOOL CView::OpenView(BOOL fBackground, BOOL fPostClose, BOOL fResetTree)
{
    if(!IsActive())
	{
		return TRUE;
	}

	Assert((!IsLockSet(VL_ENSUREINPROGRESS)||IsLockSet(VL_TASKSINPROGRESS)) && !IsLockSet(VL_RENDERINPROGRESS));
	Assert(!IsLockSet(VL_TASKSINPROGRESS) || IsFlagSet(VF_TREEOPEN));

	if((IsLockSet(VL_ENSUREINPROGRESS)&&!IsLockSet(VL_TASKSINPROGRESS)) || IsLockSet(VL_RENDERINPROGRESS))
	{
		return FALSE;
	}

    if(fResetTree)
    {
        CloseDisplayTree();
    }

    EnsureDisplayTreeIsOpen();

    if(fPostClose && !IsLockSet(VL_TASKSINPROGRESS))
    {
        PostCloseView(fBackground);
    }

    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CloseView
//
//  Synopsis:   Perform all work necessary to close the view
//
//  Arguments:  grfLayout - Collection of LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
void CView::CloseView(DWORD grfLayout)
{
    if(IsFlagSet(VF_PENDINGENSUREVIEW))
    {
        if(grfLayout & LAYOUT_SYNCHRONOUS)
        {
			GWKillMethodCall(this, ONCALL_METHOD(CView, EnsureViewCallback, ensureviewcallback), 0);
			ClearFlag(VF_PENDINGENSUREVIEW);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     PostCloseView
//
//  Synopsis:   Ensure a close view event is posted
//
//              The view is closed whenever EnsureView is called. EnsureView is
//              called for WM_PAINT handling, from the global window (for queued
//              background work or all work on pages without a frame rate), and
//              in response to the draw timer.
//
//              Background closures always occur through a message posted to
//              the global window.
//
//  Arguments:  fBackground - Close the view in the background
//
//----------------------------------------------------------------------------
void CView::PostCloseView(BOOL fBackground)
{
	if(IsActive() && !IsFlagSet(VF_PENDINGENSUREVIEW) && !IsLockSet(VL_ACCUMULATINGINVALID))
	{
        // BUGBUG: Add support for other kinds of tasks (e.g., background) and wire in support for the
        //         document frame rate/draw timer (brendand)

        HRESULT hr;

		hr = GWPostMethodCall(this,
			ONCALL_METHOD(CView, EnsureViewCallback, ensureviewcallback),
            0, FALSE, "CView::EnsureViewCallback");
		SetFlag(VF_PENDINGENSUREVIEW);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     PostRenderView
//
//  Synopsis:   Ensure rendering will take place
//
//  Arguments:  fSynchronousRedraw - Render synchronously (if possible)
//
//----------------------------------------------------------------------------
void CView::PostRenderView(BOOL fSynchronousRedraw)
{
    // For normal documents, ensure rendering will occur
    // (Assume rendering will occur for print documents)
    if(IsActive())
    {
        Trace0("View : PostRenderView\n");

        // If a synchronouse render is requested, then
        //     1) If about to render, ensure rendering occurs after publishing the invalid rectangles/region
        //     2) If not rendering, force a immediate WM_PAINT
        //     3) Otherwise, drop the request
        //        (Requests that arrive while rendering are illegal)
        if(fSynchronousRedraw)
        {
            if(IsLockSet(VL_ACCUMULATINGINVALID))
            {
                Trace0("View : PostRenderView - Setting VF_FORCEPAINT\n");

                SetFlag(VF_FORCEPAINT);
            }

            else if(!IsLockSet(VL_UPDATEINPROGRESS) && !IsLockSet(VL_RENDERINPROGRESS))
            {
                Trace0("View : PostRenderView - Calling DrawSynchronous\n");

                DrawSynchronous();
            }
        }
        // For asynchronous requests, just note that the view needs closing
        else
        {
            Trace0("View : PostRenderView - Calling PostCloseView\n");

            PostCloseView();
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawSynchronous
//
//  Synopsis:   Force a synchronous draw of the accumulated invalid areas
//
//  Arguments:  hdc - HDC into which to render
//
//----------------------------------------------------------------------------
void CView::DrawSynchronous()
{
    Assert(!IsLockSet(VL_ACCUMULATINGINVALID) && !IsLockSet(VL_UPDATEINPROGRESS) && !IsLockSet(VL_RENDERINPROGRESS));

    Trace0("View : DrawSynchronous\n");

    if(HasInvalid())
    {
        Trace((_T("View : DrawSynchronous - HasInvalid of (%d, %d, %d, %d)\n"),
               _rgnInvalid.GetBounds().left,
               _rgnInvalid.GetBounds().top,
               _rgnInvalid.GetBounds().right,
               _rgnInvalid.GetBounds().bottom));

        Trace0("View : DrawSynchronous - Calling UpdateForm\n");

        // Ensure all nested ActiveX/APPLETs receive pending SetObjectRects
        // before receiving a WM_PAINT
        // (This is necessary since Windows sends WM_PAINT to child HWNDs
        // before their parent HWND. Since we delay sending SetObjectRects
        // until we receive our own WM_PAINT, not forcing it through now
        // causes it to come after the child HWND has painted.)
        EnsureView();

        Doc()->UpdateForm();
    }

    Trace0("View : DrawSynchronous - Exit\n");
}

//+---------------------------------------------------------------------------
//
//  Member:     EnsureFocus
//
//  Synopsis:   Ensure the focus adorner is properly initialized
//
//----------------------------------------------------------------------------
void CView::EnsureFocus()
{
    if(IsActive() && _pFocusAdorner)
    {
        _pFocusAdorner->EnsureFocus();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     EnsureSize
//
//  Synopsis:   Ensure the passed (top) layout is in-sync with the current
//              view size
//
//  Arguments:  grfLayout - Current LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
void CView::EnsureSize(DWORD grfLayout)
{
    if(IsActive())
    {
        CLayout* pLayout = GetRootLayout();

        if(pLayout && !pLayout->IsDisplayNone())
        {
            // If LAYOUT_FORCE is set, re-size and re-attach the layout
            if(grfLayout & LAYOUT_FORCE)
            {
                _pLayout = NULL;
                ClearFlag(VF_SIZED);
                ClearFlag(VF_ATTACHED);
            }

            // If the current layout is not sized, size it
            if(!IsSized(pLayout))
            {
                CSize size;
                CCalcInfo CI(&Doc()->_dci, pLayout);

                GetViewSize(&size);

                CI.SizeToParent((SIZE*)&size);
                CI._grfLayout = grfLayout | LAYOUT_MEASURE;

                // NOTE: Used only by CDoc::GetDocCoords
                if(IsFlagSet(VF_FULLSIZE))
                {
                    CI._smMode = SIZEMODE_FULLSIZE;
                }

                pLayout->_fSizeThis = TRUE;
                pLayout->CalcSize(&CI, (SIZE*)&size);

                SetFlag(VF_SIZED);

                if(_pLayout != pLayout)
                {
                    _pLayout = pLayout;
                    ClearFlag(VF_ATTACHED);
                }
            }

            // If the current layout is not attached, attach it
            if(!IsAttached(pLayout))
            {
                CDispNode* pDispNode;

                // Remove the previous top-most layout (if it exists)
                Assert(_pDispRoot->CountChildren() <= 1);
                pDispNode = _pDispRoot->GetFirstChildNode();

                if(pDispNode)
                {
                    pDispNode->ExtractFromTree();
                }

                // Insert the current top-most layout
                pDispNode = pLayout->GetElementDispNode();

                if(pDispNode)
                {
                    pDispNode->SetPosition(_afxGlobalData._Zero.pt);
                    _pDispRoot->InsertChildInFlow(pDispNode);
                    SetFlag(VF_ATTACHED);
                }
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     EnsureViewCallback
//
//  Synopsis:   Process a posted call to EnsureView
//
//  Arguments:  grfLayout - Collection of LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
void CView::EnsureViewCallback(DWORD_PTR dwContext)
{
    DWORD grfLayout = (DWORD)dwContext;
    // We need to ref the document because all external holders of the doc
    // could go away during EnsureView.  E.g., EnsureView may cause embedded
    // controls to transition state; that gives external code a chance to
    // execute, and they may release all refs to the doc.
    CServer::CLock Lock(_pDoc);

    Assert(IsInitialized());
    Assert(IsActive());

    // GlobalWndOnMethodCall() will have cleared the callback entry in its array;
    // now clear the flag to keep its meaning sync'ed up.
    ClearFlag(VF_PENDINGENSUREVIEW);

    Verify(EnsureView(grfLayout));
}

//+---------------------------------------------------------------------------
//
//  Member:     FixWindowZOrder
//
//  Synopsis:   Fix Z order of child windows.
//
//  Arguments:  none
//              
//----------------------------------------------------------------------------
void CView::FixWindowZOrder()
{
    if(IsActive())
    {
        CWindowOrderInfo windowOrderInfo;
        _pDispRoot->TraverseTreeTopToBottom((void*)&windowOrderInfo);
        windowOrderInfo.SetWindowOrder();
    }
}

CView::CWindowOrderInfo::CWindowOrderInfo()
{
    _hdwp = ::BeginDeferWindowPos(30);
    _hwndAfter = HWND_TOP;
}

void CView::CWindowOrderInfo::AddWindow(HWND hwnd)
{
    if(_hdwp)
	{
		_hdwp = ::DeferWindowPos(
			_hdwp,
			hwnd,
			_hwndAfter,
			0, 0, 0, 0,
			SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
    }
    else
	{
		::SetWindowPos(
			hwnd,
			_hwndAfter,
			0, 0, 0, 0,
			SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
	}

    _hwndAfter = hwnd;
}

void CView::CWindowOrderInfo::SetWindowOrder()
{
    if(_hdwp)
    {
        ::EndDeferWindowPos(_hdwp);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetRootLayout
//
//  Synopsis:   Return the top-most layout of the associated CMarkup
//
//  Returns:    CLayout of the top-most element (if available), NULL otherwise
//
//----------------------------------------------------------------------------
CLayout* CView::GetRootLayout() const
{
    CElement* pElement = _pDoc->GetPrimaryElementClient();
    return pElement?pElement->GetUpdatedLayout():NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     ScrollInvalid
//
//  Synopsis:   Scroll any outstanding invalid rectangles/region that
//              intersects the passed scrolled rectangle
//
//  Arguments:  rcScroll        - Rectangle that scrolled
//              sizeScrollDelta - Amount of the scroll
//
//----------------------------------------------------------------------------
void CView::ScrollInvalid(const CRect& rcScroll, const CSize& sizeScrollDelta)
{
    Trace((_T("View : ScrollInvalid - rc(%d, %d, %d, %d), delta(%d, %d)\n"),
           rcScroll.left,
           rcScroll.top,
           rcScroll.right,
           rcScroll.bottom,
           sizeScrollDelta.cx,
           sizeScrollDelta.cy));

    if(HasInvalid())
    {
        CRegion rgnAdjust(rcScroll);

        MergeInvalid();

        rgnAdjust.Intersect(_rgnInvalid);

        if(!rgnAdjust.IsEmpty())
        {
            _rgnInvalid.Subtract(rcScroll);
            rgnAdjust.Offset(sizeScrollDelta);
            rgnAdjust.Intersect(rcScroll);
            _rgnInvalid.Union(rgnAdjust);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     ClearInvalid
//
//  Synopsis:   Clear accumulated invalid rectangles/region
//
//----------------------------------------------------------------------------
void CView::ClearInvalid()
{
    if(IsActive())
    {
        ClearFlag(VF_INVALCHILDWINDOWS);
        _cInvalidRects = 0;
        _rgnInvalid.SetEmpty();
    }

    Assert(!HasInvalid());
}

//+---------------------------------------------------------------------------
//
//  Member:     PublishInvalid
//
//  Synopsis:   Push accumulated invalid rectangles/region into the
//              associated device (HWND)
//
//  Arguments:  grfLayout - Current LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
void CView::PublishInvalid(DWORD grfLayout)
{
    if(IsActive())
    {
        CPoint pt;

        MergeInvalid();

        if(!(grfLayout & LAYOUT_DEFERINVAL))
        {
            GetViewPosition(&pt);

            if(_rgnInvalid.IsRegion())
            {
                if(pt != _afxGlobalData._Zero.pt)
                {
                    _rgnInvalid.Offset(pt.AsSize());
                }

                HRGN hrgn;
                
                hrgn = _rgnInvalid.OrphanRegion();

                Doc()->Invalidate(NULL, NULL, hrgn, (IsFlagSet(VF_INVALCHILDWINDOWS)?INVAL_CHILDWINDOWS:0));

                ::DeleteObject(hrgn);
            }
            else if(!_rgnInvalid.IsEmpty())
            {
                CRect rcInvalid(_rgnInvalid.AsRect());

                rcInvalid.OffsetRect(pt.AsSize());

                Doc()->Invalidate(&rcInvalid, NULL, NULL, (IsFlagSet(VF_INVALCHILDWINDOWS)?INVAL_CHILDWINDOWS:0));
            }

            ClearInvalid();
        }
    }

    Assert(grfLayout&LAYOUT_DEFERINVAL || !HasInvalid());
}

//+---------------------------------------------------------------------------
//
//  Member:     MergeInvalid
//
//  Synopsis:   Merge the collected invalid rectangles/region into a
//              single CRegion object
//
//----------------------------------------------------------------------------
void CView::MergeInvalid()
{
    if(_cInvalidRects)
    {
        // Create the invalid region
        // (If too many invalid rectangles arrived during initial load,
        //  union them all into a single rectangle)
        if(_cInvalidRects==MAX_INVALID && !IsFlagSet(VF_HASRENDERED))
        {
            CRect rcUnion(_aryInvalidRects[0]);
            CRect rcView;

            for(int i=1; i<_cInvalidRects; i++)
            {
                rcUnion.Union(_aryInvalidRects[i]);
            }

            GetViewRect(&rcView);
            rcUnion.IntersectRect(rcView);

            _rgnInvalid.Union(rcUnion);
        }
        else
        {
            for(int i=0; i<_cInvalidRects; i++)
            {
                _rgnInvalid.Union(_aryInvalidRects[i]);
            }
        }

        _cInvalidRects = 0;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     OpenDisplayTree
//
//  Synopsis:   Open the display tree (if it exists)
//
//----------------------------------------------------------------------------
void CView::OpenDisplayTree()
{
    Assert(IsActive() || !IsFlagSet(VF_TREEOPEN));

	if(IsActive())
	{
		Assert((IsFlagSet(VF_TREEOPEN) && _pDispRoot->DisplayTreeIsOpen())
			|| (!IsFlagSet(VF_TREEOPEN) && !_pDispRoot->DisplayTreeIsOpen()));

		if(!IsFlagSet(VF_TREEOPEN))
		{
			_pDispRoot->OpenDisplayTree();
			SetFlag(VF_TREEOPEN);
		}
	}
}

//+---------------------------------------------------------------------------
//
//  Member:     CloseDisplayTree
//
//  Synopsis:   Close the display tree
//
//----------------------------------------------------------------------------
void CView::CloseDisplayTree()
{
    Assert(IsActive() || !IsFlagSet(VF_TREEOPEN));

    if(IsActive())
    {
        Assert((IsFlagSet(VF_TREEOPEN) && _pDispRoot->DisplayTreeIsOpen())
            || (!IsFlagSet(VF_TREEOPEN) && !_pDispRoot->DisplayTreeIsOpen()));

        if(IsFlagSet(VF_TREEOPEN))
        {
            CServer::CLock lock(Doc(), SERVERLOCK_IGNOREERASEBKGND);
            _pDispRoot->CloseDisplayTree();
            ClearFlag(VF_TREEOPEN);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     WantOffscreenBuffer
//
//  Synopsis:   Determine if an offscreen buffer should be used
//
//  Returns:    TRUE if an offscreen buffer should be used, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::WantOffscreenBuffer() const
{
    return !((_pDoc->_dwFlagsHostInfo&DOCHOSTUIFLAG_DISABLE_OFFSCREEN)
            || !IsFlagSet(VF_HASRENDERED));
}

//+---------------------------------------------------------------------------
//
//  Member:     AllowOffscreenBuffer
//
//  Synopsis:   Determine if an offscreen buffer can be used
//              NOTE: This should override all other checks (e.g., WantOffscreenBuffer)
//
//  Returns:    TRUE if an offscreen buffer can be used, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::AllowOffscreenBuffer() const
{
	return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     AllowScrolling
//
//  Synopsis:   Determine if scrolling of the DC is allowed
//
//  Returns:    TRUE if scrolling is allowed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::AllowScrolling() const
{
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     AllowSmoothScrolling
//
//  Synopsis:   Determine if smooth scrolling is allowed
//
//  Returns:    TRUE if smooth scrolling is allowed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::AllowSmoothScrolling() const
{
    return _pDoc->_pOptionSettings->fSmoothScrolling&&AllowScrolling();
}

#ifdef _DEBUG
//+---------------------------------------------------------------------------
//
//  Member:     IsDisplayTreeOpen
//
//  Synopsis:   Check if the display tree is open
//
//----------------------------------------------------------------------------
BOOL CView::IsDisplayTreeOpen() const
{
    return (IsActive() && _pDispRoot && _pDispRoot->DisplayTreeIsOpen());
}
#endif

//+---------------------------------------------------------------------------
//
//  Member:     AddTask
//
//  Synopsis:   Add a task to the view-task queue
//
//  Arguments:  pv        - Object to call
//              vtt       - Task type
//              grfLayout - Collection of LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
HRESULT CView::AddTask(
    void* pv,
    CViewTask::VIEWTASKTYPE vtt,
    DWORD grfLayout,
    LONG lData)
{
    if(!IsActive())
	{
		return S_OK;
	}

    CViewTask vt(pv, vtt, lData, grfLayout);
    CAryVTasks* pTaskList = GetTaskList(vtt);
    int i;
    HRESULT hr;

    AssertSz(_pDoc, "View used while CDoc is not inplace active");
    AssertSz(_pDispRoot, "View used while CDoc is not inplace active");
    Assert(!(grfLayout & LAYOUT_NONTASKFLAGS));
    Assert(!(grfLayout & LAYOUT_TASKDELETED));

    i = vtt!=CViewTask::VTT_EVENT ? FindTask(pTaskList, vt) : -1;

    if(i < 0)
    {
        hr = pTaskList->AppendIndirect(&vt);

        if(SUCCEEDED(hr))
        {
            PostCloseView(vt.IsFlagSet(LAYOUT_BACKGROUND));
        }
    }
    else
    {
        CViewTask* pvt = &(*pTaskList)[i];
        
        pvt->AddFlags(grfLayout);
        Assert(!(pvt->IsFlagsSet(LAYOUT_BACKGROUND|LAYOUT_PRIORITY)));

        hr = S_OK;
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     ExecuteTask
//
//  Synopsis:   Execute a single task from the task queue
//
//  Arguments:  pv        - Object to call
//              vtt       - Task type
//
//----------------------------------------------------------------------------
void CView::ExecuteTask(void* pv, CViewTask::VIEWTASKTYPE vtt)
{
    CViewTask* pvt = GetTask(pv, vtt);

    Assert(IsActive() || !pvt);
    Assert(!pvt || !pvt->IsFlagSet(LAYOUT_TASKDELETED));

    if(pvt)
    {
        CViewTask vtTask = *pvt;

        switch(vtt)
        {
        case CViewTask::VTT_LAYOUT:
            vtTask.GetLayout()->DoLayout(GetLayoutFlags());
            Assert(FindTask((CAryVTasks*)&_aryTaskLayout, vtTask) < 0);
            break;
#ifdef _DEBUG
        default:
            Assert(FALSE);
            break;
#endif
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     ExecuteLayoutTasks
//
//  Synopsis:   Execute pending layout tasks in document source order
//
//  Arguments:  grfLayout - Collections of LAYOUT_xxxx flags
//
//  Returns:    TRUE if all tasks were processed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::ExecuteLayoutTasks(DWORD grfLayout)
{
    Assert(IsActive());

    CViewTask* pvtTask;
    DWORD fLayout = (grfLayout & (LAYOUT_MEASURE|LAYOUT_POSITION|LAYOUT_ADORNERS));

    while(pvtTask=GetNextTaskInSourceOrder(CViewTask::VTT_LAYOUT, fLayout))
    {
        Assert(pvtTask->IsType(CViewTask::VTT_LAYOUT));
        Assert(pvtTask->GetLayout());
        Assert(!(pvtTask->GetFlags() & LAYOUT_NONTASKFLAGS));
        Assert(!pvtTask->IsFlagSet(LAYOUT_TASKDELETED));

        // Mark task has having completed the current LAYOUT_MEASURE/POSITION/ADORNERS pass
        pvtTask->_grfLayout &= ~fLayout;

        // Execute the task
        pvtTask->GetLayout()->DoLayout(grfLayout);
    }

    if(fLayout & LAYOUT_ADORNERS)
    {
        _aryTaskLayout.DeleteAll();
    }

    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     ExecuteEventTasks
//
//  Synopsis:   Execute pending event-firing tasks in receiving order
//
//  Arguments:  grfLayout - Collections of LAYOUT_xxxx flags
//
//  Returns:    TRUE if all tasks were processed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CView::ExecuteEventTasks(DWORD grfLayout)
{
    Assert(IsActive());

    if(grfLayout & LAYOUT_DEFEREVENTS)
	{
		return (_aryTaskEvent.Size()==0);
	}

    CElement* pElement;
    CAryVTaskEvent aryTaskEvent;
    CViewTask* pvtTask;
    int cTasks;

    aryTaskEvent.Copy(_aryTaskEvent, FALSE);
    _aryTaskEvent.DeleteAll();

    // Ensure the elements stay for the duration of all pending events
    // (The reference count is increment once per-pending event)
    cTasks = aryTaskEvent.Size();
	for(pvtTask=&((CAryVTaskEvent&)aryTaskEvent)[0]; cTasks; pvtTask++,cTasks--)
	{
        Assert(pvtTask->GetElement());

        pElement = pvtTask->GetElement();
        pElement->AddRef();
        pElement->_fHasPendingEvent = FALSE;
    }

    // Fire the events and release the hold on the elements
    cTasks = aryTaskEvent.Size();
    for(pvtTask=&((CAryVTaskEvent&)aryTaskEvent)[0]; cTasks; pvtTask++,cTasks--)
    {
        Assert(pvtTask->IsType(CViewTask::VTT_EVENT));
        Assert(pvtTask->GetElement());
        Assert(!pvtTask->IsFlagSet(LAYOUT_TASKDELETED));

        pElement = pvtTask->GetElement();

        switch(pvtTask->GetEventDispID())
        {
        case DISPID_EVMETH_ONRESIZE:
            if(_pDoc->_fFiredOnLoad)
			{
                //if((pElement->Tag()==ETAG_BODY) && _pDoc->_pOmWindow)
                //{
                //    _pDoc->_pOmWindow->Fire_onresize();
                //}
                //else
                //{
                //    pElement->Fire_onresize();
                //}
            }
            break;

        default:
            break;
        }

        pElement->Release();
    }

    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     FindTask
//
//  Synopsis:   Locate a task in the task array
//
//  Arguments:  vt  - Task to remove
//
//  Returns:    If found, index of task in task array
//              Otherwise, -1
//
//----------------------------------------------------------------------------
int CView::FindTask(const CAryVTasks* pTaskList, const CViewTask& vt) const
{
    int i;

    Assert(IsActive() || !(pTaskList->Size()));

    if(!IsActive())
	{
		return -1;
	}

    // If the _pv field is supplied, look for an exact match
    if(vt.GetObject())
    {
		for(i=pTaskList->Size()-1;
			i >= 0&&!((CAryVTasks&)(*pTaskList))[i].IsFlagSet(LAYOUT_TASKDELETED)
			&&vt!=(const CViewTask&)(((CAryVTasks&)(*pTaskList))[i]); i--) ;
    }
    // Otherwise, match only on task type
    else
    {
		for(i=pTaskList->Size()-1;
			i>=0&& !((CAryVTasks&)(*pTaskList))[i].IsFlagSet(LAYOUT_TASKDELETED)
			&& !vt.IsType(((CAryVTasks&)(*pTaskList))[i]); i--) ;
    }

    return i;
}

//+---------------------------------------------------------------------------
//
//  Member:     GetTask
//
//  Synopsis:   Retrieve the a specific task from the appropriate task queue
//
//  Arguments:  vt - CViewTask to retrieve
//
//  Returns:    CViewTask pointer if found, NULL otherwise
//
//----------------------------------------------------------------------------
CViewTask* CView::GetTask(const CViewTask& vt) const
{
    CAryVTasks* paryTasks = GetTaskList(vt.GetType());
    int iTask = FindTask(paryTasks, vt);

    Assert(IsActive() || iTask<0);
    Assert(iTask<0 || !((*paryTasks)[iTask]).IsFlagSet(LAYOUT_TASKDELETED));

    return (iTask>=0 ? &((*paryTasks)[iTask]) : NULL);
}

//+---------------------------------------------------------------------------
//
//  Member:     GetNextTask
//
//  Synopsis:   Retrieve the next task of the specified task type
//
//  Arguments:  vtt       - VIEWTASKTYPE to search for
//              grfLayout - LAYOUT_MEASURE/POSITION/ADORNERS filter
//
//  Returns:    CViewTask pointer if found, NULL otherwise
//
//----------------------------------------------------------------------------
CViewTask* CView::GetNextTask(CViewTask::VIEWTASKTYPE vtt, DWORD grfLayout) const
{
    CViewTask*	pvtTask = NULL;
    CAryVTasks*	pTaskList = GetTaskList(vtt);
    int			cTasks;

    cTasks = pTaskList->Size();

    Assert(IsActive() || !cTasks);

    if(cTasks)
    {
        for(pvtTask=&((CAryVTasks&)(*pTaskList))[0];
            cTasks&&(pvtTask->IsFlagSet(LAYOUT_TASKDELETED)||!pvtTask->IsFlagsSet(grfLayout)||!pvtTask->IsType(vtt));
            pvtTask++,cTasks--) ;
	}

    Assert(cTasks<=0 || !pvtTask->IsFlagSet(LAYOUT_TASKDELETED));

    return (cTasks>0 ? pvtTask : NULL);
}

//+---------------------------------------------------------------------------
//
//  Member:     GetNextTaskInSourceOrder
//
//  Synopsis:   Retrieve the next task, in document source order, of the
//              specified task type
//
//  Arguments:  vtt       - VIEWTASKTYPE to search for
//              grfLayout - LAYOUT_MEASURE/POSITION/ADORNERS filter
//
//  Returns:    CViewTask pointer if found, NULL otherwise
//
//----------------------------------------------------------------------------
CViewTask* CView::GetNextTaskInSourceOrder(CViewTask::VIEWTASKTYPE vtt, DWORD grfLayout) const
{
    CViewTask*	pvtTask = NULL;
    CAryVTasks*	pTaskList = GetTaskList(vtt);
    int			cTasks;

    cTasks = pTaskList->Size();

    Assert(IsActive() || !cTasks);

    if(cTasks > 1)
    {
        CViewTask* pvt;
        int si, siTask;

        siTask = INT_MAX;

        for(pvt=&((CAryVTasks&)(*pTaskList))[0]; cTasks; pvt++,cTasks--)
        {
			if(!pvt->IsFlagSet(LAYOUT_TASKDELETED) && pvt->IsFlagsSet(grfLayout) && pvt->IsType(vtt))
			{
                Assert(pvt->GetObject());

                si = pvt->GetSourceIndex();
                if(si < siTask)
                {
                    siTask  = si;
                    pvtTask = pvt;
                }
            }
        }
    }
	else if(cTasks
		&& !((CAryVTasks&)(*pTaskList))[0].IsFlagSet(LAYOUT_TASKDELETED)
		&& ((CAryVTasks&)(*pTaskList))[0].IsFlagsSet(grfLayout)
		&& ((CAryVTasks&)(*pTaskList))[0].IsType(vtt))
    {
        Assert(cTasks == 1);
        pvtTask = &((CAryVTasks&)(*pTaskList))[0];
    }

    Assert(!pvtTask || !pvtTask->IsFlagSet(LAYOUT_TASKDELETED));

    return pvtTask;
}

//+---------------------------------------------------------------------------
//
//  Member:     RemoveTask
//
//  Synopsis:   Remove a task from the view-task queue
//
//              NOTE: This routine is safe to call with non-existent tasks
//
//  Arguments:  iTask - Index of task to remove, -1 the task is non-existent
//
//----------------------------------------------------------------------------
void CView::RemoveTask(CAryVTasks* pTaskList, int iTask)
{
    Assert(IsActive() || !(pTaskList->Size()));

    if(iTask>=0 && iTask<pTaskList->Size())
    {
        if(pTaskList!=&_aryTaskLayout || !IsLockSet(VL_TASKSINPROGRESS))
        {
            pTaskList->Delete(iTask);
        }
        else
        {
            (*pTaskList)[iTask]._grfLayout |= LAYOUT_TASKDELETED;
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     AddAdorner
//
//  Synopsis:   Add a CAdorner to the list of adorners
//
//              Adorners are kept in order by their start cp. If two (or more)
//              have the same start cp, they are sorted by their end cp.
//
//  Arguments:  pAdorner - CAdorner to add
//
//----------------------------------------------------------------------------
HRESULT CView::AddAdorner(CAdorner* pAdorner)
{
    long cpStartNew, cpEndNew;
    long iAdorner, cAdorners;
    HRESULT hr = E_FAIL;

    if(IsActive())
    {
        Assert(pAdorner);

        pAdorner->GetRange(&cpStartNew, &cpEndNew);

        iAdorner = 0;
        cAdorners = _aryAdorners.Size();

        if(cAdorners)
        {
            CAdorner** ppAdorner;
            long cpStart, cpEnd;

            for(ppAdorner=&(_aryAdorners[0]); iAdorner<cAdorners; iAdorner++,ppAdorner++)
            {
                (*ppAdorner)->GetRange(&cpStart, &cpEnd);

                if(cpStartNew<cpStart || (cpStartNew==cpStart && cpEndNew<cpEnd))
				{
					break;
				}
            }
        }

        Assert(iAdorner <= cAdorners);

        hr = _aryAdorners.InsertIndirect(iAdorner, &pAdorner);

        if(SUCCEEDED(hr))
        {
            CElement* pElement = pAdorner->GetElement();

            Verify(OpenView());
            Assert(cpEndNew >= cpStartNew);
            AccumulateMeasuredRange(cpStartNew, (cpEndNew-cpStartNew));


            if(pElement && pElement->HasLayout())
            {
                pElement->GetCurLayout()->SetIsAdorned(TRUE);
            }
        }
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     DeleteAdorners
//
//  Synopsis:   Delete and remove all adorners
//
//----------------------------------------------------------------------------
void CView::DeleteAdorners()
{
    long iAdorner  = 0;
    long cAdorners = _aryAdorners.Size();

    Assert(IsActive() || !cAdorners);

    if(cAdorners)
    {
        CAdorner** ppAdorner;

        for(ppAdorner=&(_aryAdorners[0]); iAdorner<cAdorners; iAdorner++,ppAdorner++)
        {
            (*ppAdorner)->Destroy();
        }

        _aryAdorners.DeleteAll();

        _pFocusAdorner = NULL;
    }

    Assert(!_pFocusAdorner);
}

//+---------------------------------------------------------------------------
//
//  Member:     GetAdorner
//
//  Synopsis:   Find the next adorner associated with the passed element
//
//  Arguments:  pElement  - CElement associated with the adorner
//              piAdorner - Index at which to begin the search
//
//  Returns:    Index of adorner if found, -1 otherwise
//
//----------------------------------------------------------------------------
long CView::GetAdorner(CElement* pElement, long iAdorner) const
{
    Assert(IsActive() || !_aryAdorners.Size());

    if(!IsActive())
	{
		return -1;
	}

    long cAdorners = _aryAdorners.Size();

    Assert(iAdorner >= 0);

    if(iAdorner < cAdorners)
    {
        CAdorner** ppAdorner;

		for(ppAdorner=&(((CAryAdorners &)_aryAdorners)[iAdorner]); iAdorner<cAdorners; iAdorner++,ppAdorner++)
		{
            if((*ppAdorner)->GetElement() == pElement)
			{
				break;
			}
        }
    }

    return (iAdorner<cAdorners ? iAdorner : -1);
}

//+---------------------------------------------------------------------------
//
//  Member:     RemoveAdorner
//
//  Synopsis:   Remove the specified adorner
//
//  Arguments:  pAdorner - CAdorner to remove
//
//----------------------------------------------------------------------------
void CView::RemoveAdorner(CAdorner* pAdorner, BOOL fCheckForLast)
{
    Assert(IsActive() || !_aryAdorners.Size());

    if(IsActive())
    {
        CElement* pElement = pAdorner->GetElement();

        Assert(IsInState(VS_OPEN));

        if(_aryAdorners.DeleteByValueIndirect(&pAdorner))
        {
            pAdorner->Destroy();
        }

        if(fCheckForLast && pElement && pElement->HasLayout())
        {
            pElement->GetCurLayout()->SetIsAdorned(HasAdorners(pElement));
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     AccumulateMeasuredRanged
//
//  Synopsis:   Accumulate the measured range
//
//  Arguments:  cp  - First character of the range
//              cch - Number of characters in the range
//
//----------------------------------------------------------------------------
void CView::AccumulateMeasuredRange(long cp, long cch)
{
    if(IsActive())
    {
        long cpEnd = cp + cch;

        Assert(IsInState(VS_OPEN));

        if(_cpStartMeasured < 0)
        {
            _cpStartMeasured = cp;
            _cpEndMeasured   = cpEnd;
        }
        else
        {
            _cpStartMeasured = min(_cpStartMeasured, cp);
            _cpEndMeasured   = max(_cpEndMeasured, cpEnd);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     AccumulateTranslatedRange
//
//  Synopsis:   Accumulate the translated range
//
//  Arguments:  size - Amount of translation
//              cp   - First character of the range
//              cch  - Number of characters in the range
//
//----------------------------------------------------------------------------
void CView::AccumulateTranslatedRange(const CSize& size, long cp, long cch)
{
    if(IsActive())
    {
        long cpEnd = cp + cch;

        Assert(IsInState(VS_OPEN));

        if(_cpStartTranslated < 0)
        {
            _cpStartTranslated = cp;
            _cpEndTranslated = cpEnd;
            _sizeTranslated = size;
        }
        else
        {
            _cpStartTranslated = min(_cpStartTranslated, cp);
            _cpEndTranslated = max(_cpEndTranslated, cpEnd);

            if(_sizeTranslated != size)
            {
                _sizeTranslated = _afxGlobalData._Zero.size;
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     UpdateAdorners
//
//  Synopsis:   Notify the adorners within the measured and translated ranges
//              possible size/position changes
//
//  Arguments:  grfLayout - Current LAYOUT_xxxx flags
//
//----------------------------------------------------------------------------
void CView::UpdateAdorners(DWORD grfLayout)
{
    Assert(IsActive() || !_aryAdorners.Size());

    if(!IsActive())
	{
		return;
	}

    Assert(_cpStartMeasured<0 || _cpEndMeasured>=0);
    Assert(_cpStartMeasured <= _cpEndMeasured);
    Assert(_cpStartTranslated<0 || _cpEndTranslated>=0);
    Assert(_cpStartTranslated <= _cpEndTranslated);

    // If adorners and a dirty range exists, update the affected adorners
    if(_aryAdorners.Size() && HasDirtyRange())
    {
        long        iAdorner;
        long        cAdorners;
        long        cpStart, cpEnd;
        long        cpFirst, cpLast;
        CAdorner**  ppAdorner;

        _aryAdorners[0]->GetRange(&cpStart, &cpEnd);
        cpFirst = cpStart;

        _aryAdorners[_aryAdorners.Size()-1]->GetRange(&cpStart, &cpEnd);
        cpLast = cpEnd;

        // If a measured range intersects the adorners, update the affected adorners
		if(HasMeasuredRange()
			&& ((cpFirst>=_cpStartMeasured && cpFirst<_cpEndMeasured)
			|| (cpLast>=_cpStartMeasured && cpLast<=_cpEndMeasured)))
		{
            // Find the first adorner to intersect the range
            cpStart = LONG_MAX;
            cpEnd   = LONG_MIN;

            iAdorner  = 0;
            cAdorners = _aryAdorners.Size();

			for(ppAdorner=&(_aryAdorners[iAdorner]); iAdorner<cAdorners; iAdorner++,ppAdorner++)
			{
                (*ppAdorner)->GetRange(&cpStart, &cpEnd);

                if(cpEnd >= _cpStartMeasured)
				{
					break;
				}
            }

            // Notify all adorners that intersect the range
            cAdorners -= iAdorner;

            if(cAdorners)
            {
                while(cpStart < _cpEndMeasured)
                {
                    (*ppAdorner)->ShapeChanged();
                    (*ppAdorner)->PositionChanged();

                    cAdorners--;
                    if(!cAdorners)
					{
						break;
					}

                    ppAdorner++;
                    (*ppAdorner)->GetRange(&cpStart, &cpEnd);
                }
            }
        }

        // If a translated range intersects the adorners, update the affected adorners
		if(HasTranslatedRange()
			&& ((cpFirst>=_cpStartTranslated &&cpFirst<_cpEndTranslated)
			|| (cpLast>=_cpStartTranslated && cpLast<=_cpEndTranslated)))
        {
            CSize* psizeTranslated;

            // If a measured range exists,
            // shrink the translated range such that they do not overlap
            if(HasMeasuredRange())
            {
                if(_cpEndMeasured <= _cpEndTranslated)
                {
                    _cpStartTranslated = max(_cpStartTranslated, _cpEndMeasured);
                }

                if(_cpStartMeasured >= _cpStartTranslated)
                {
                    _cpEndTranslated = min(_cpEndTranslated, _cpStartMeasured);
                }

                // marka - Translated can be < Measured range
				if((_cpEndTranslated<_cpEndMeasured) && (_cpStartTranslated>_cpStartMeasured))
                {
                    _cpEndTranslated = 0;
                    _cpStartTranslated = 0;
                }
            }

            if(_cpStartTranslated < _cpEndTranslated)
            {
                // Find the first adorner to intersect the range
                cpStart = LONG_MAX;
                cpEnd = LONG_MIN;

                iAdorner = 0;
                cAdorners = _aryAdorners.Size();

                psizeTranslated = _sizeTranslated!=_afxGlobalData._Zero.size ? &_sizeTranslated : NULL;

				for(ppAdorner=&(_aryAdorners[iAdorner]); iAdorner<cAdorners; iAdorner++,ppAdorner++)
                {
                    (*ppAdorner)->GetRange(&cpStart, &cpEnd);

                    if(cpStart >= _cpStartTranslated)
					{
						break;
					}
                }

                // Notify all adorners that intersect the range
                cAdorners -= iAdorner;

                if(cAdorners)
                {
                    while(cpStart < _cpEndTranslated)
                    {
                        (*ppAdorner)->PositionChanged(psizeTranslated);

                        cAdorners--;
                        if(!cAdorners)
						{
							break;
						}

                        ppAdorner++;
                        (*ppAdorner)->GetRange(&cpStart, &cpEnd);
                    }
                }
            }
        }
    }
}