
#include "stdafx.h"
#include "FlowLayout.h"

#include "../display/DispItemPlus.h"
#include "../display/DispScroller.h"

#include "../text/LSMeasurer.h"

#define DO_PROFILE			0
#define MAX_RECURSION_LEVEL	40

CFlowLayout::CFlowLayout(CElement* pElementFlowLayout) : CLayout(pElementFlowLayout)
{
	_xMax = _xMin = -1;
	_fCanHaveChildren = TRUE;
	Assert(ElementContent());
	ElementContent()->_fOwnsRuns = TRUE;
}

CFlowLayout::~CFlowLayout()
{
	Assert(_fDetached);
}

void CFlowLayout::DoLayout(DWORD grfLayout)
{
	Assert(grfLayout&(LAYOUT_MEASURE|LAYOUT_POSITION|LAYOUT_ADORNERS));

	CElement* pElementOwner = ElementOwner();

	// If this layout is no longer needed, ignore the request and remove it
	if(pElementOwner->HasLayout() && !pElementOwner->NeedsLayout())
	{
		ElementOwner()->DestroyLayout();
	}
	// Hidden layout should just accumulate changes
	// (It will be measured when re-shown)
	else if(!IsDisplayNone() || Tag()==ETAG_BODY)
	{
		CCalcInfo CI(this);
		CSize size;

		GetSize(&size);

		CI._grfLayout |= grfLayout;

		// If requested, measure
		if(grfLayout & LAYOUT_MEASURE)
		{
			// we want to do this each time inorder to
			// properly pick up things like opacity.
			if(_fForceLayout)
			{
				CI._grfLayout |= LAYOUT_FORCE;
			}

			EnsureDispNode(&CI, !!(CI._grfLayout & LAYOUT_FORCE));

			if(IsDirty() || (CI._grfLayout&(LAYOUT_FORCE|LAYOUT_FORCEFIRSTCALC)))
			{
				CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);

				CalcTextSize(&CI, &size);

				if(CI._grfLayout & LAYOUT_FORCE)
				{
					SizeDispNode(&CI, size);
					SizeContentDispNode(CSize(_dp.GetMaxWidth(), _dp.GetHeight()));

					if(pElementOwner->IsAbsolute())
					{
						ElementOwner()->SendNotification(NTYPE_ELEMENT_SIZECHANGED);
					}
				}
			}

			Reset(FALSE);
		}
		_fForceLayout = FALSE;

		// Process outstanding layout requests (e.g., sizing positioned elements, adding adorners)
		if(HasRequestQueue())
		{
			long xParentWidth;
			long yParentHeight;

			_dp.GetViewWidthAndHeightForChild(
				&CI,
				&xParentWidth,
				&yParentHeight,
				CI._smMode==SIZEMODE_MMWIDTH);

			ProcessRequests(&CI, CSize(xParentWidth, yParentHeight));
		}

		Assert(!HasRequestQueue() || GetView()->HasLayoutTask(this));
	}
	else
	{
		FlushRequests();
		RemoveLayoutRequest();
		Assert(!HasRequestQueue() || GetView()->HasLayoutTask(this));
	}
}

void CFlowLayout::Notify(CNotification* pnf)
{
	DWORD dwData;
	BOOL fHandle = TRUE;

	//  Respond to the notification if:
	//      a) The channel is enabled
	//      b) The text is not sizing
	//      c) Or it is a position/z-change notification
	if(IsListening() && (!TestLock(CElement::ELEMENTLOCK_SIZING)
		|| IsPositionNotification(pnf)))
	{
		BOOL fRangeSet = FALSE;

		// For notifications originating from this element,
		// set the range to the content rather than that in the container (if they differ)
		if(pnf->Element()==ElementOwner() && ElementOwner()->HasSlaveMarkupPtr())
		{
			pnf->SetTextRange(GetContentFirstCp(), GetContentTextLength());
			fRangeSet = TRUE;
		}

#ifdef _DEBUG
		long cp = _dtr._cp;
		long cchNew = _dtr._cchNew;
		long cchOld = _dtr._cchOld;

		if(pnf->IsTextChange())
		{
			Assert(pnf->Cp(GetContentFirstCp()) >= 0);
		}
#endif

		// If the notification is already "handled",
		// Make adjustments to cached values for any text changes
		if(pnf->IsHandled() && pnf->IsTextChange())
		{
			_dtr.Adjust(pnf, GetContentFirstCp());

			GetDisplay()->Notify(pnf);
		}
		// If characters or an element are invalidating,
		// then immediately invalidate the appropriate rectangles
		else if(IsInvalidationNotification(pnf))
		{
			// Invalidate the entire layout if the associated element initiated the request
			if(ElementOwner()==pnf->Element() || pnf->IsType(NTYPE_ELEMENT_INVAL_Z_DESCENDANTS))
			{
				Invalidate();
			}
			// Otherwise, isolate the appropriate range and invalidate the derived rectangles
			else
			{
				long cpFirst = GetContentFirstCp();
				long cpLast  = GetContentLastCp();
				long cp      = pnf->Cp(cpFirst) - cpFirst;
				long cch     = pnf->Cch(cpLast);

				Assert(pnf->IsType(NTYPE_ELEMENT_INVALIDATE)
					|| pnf->IsType(NTYPE_CHARS_INVALIDATE));
				Assert(cp >= 0);
				Assert(cch >= 0);

				// Obtain the rectangles if the request range is measured
				if(IsRangeBeforeDirty(cp, cch) && (cp+cch)<=GetDisplay()->_dcpCalcMax)
				{
					CDataAry<RECT> aryRects;
					CTreeNode* pNotifiedNode = pnf->Element()->GetFirstBranch();
					CTreeNode* pRelativeNode;
					CRelDispNodeCache* pRDNC;

					GetDisplay()->RegionFromRange(&aryRects, pnf->Cp(cpFirst), cch);

					// If the notified element is relative or is contained
					// within a relative element (i.e. what we call "inheriting"
					// relativeness), then we need to find the element that's responsible
					// for the relativeness, and invalidate its dispnode.
					if(pNotifiedNode->IsInheritingRelativeness())
					{
						pRDNC = _dp.GetRelDispNodeCache();
						if(pRDNC) 
						{
							// BUGBUG: this assert is legit; remove the above if clause
							// once OnPropertyChange() is modified to not fire invalidate
							// when its dwFlags have remeasure in them.  The problem is
							// that OnPropertyChange is invalidating when we've been
							// asked to remeasure, so the dispnodes/reldispnodcache may
							// not have been created yet.
							Assert(pRDNC && "Must have a RDNC if one of our descendants inherited relativeness!");                       
							// Find the element that's causing the notified element
							// to be relative.  Search up to the flow layout owner.
							pRelativeNode = pNotifiedNode->GetCurrentRelativeNode(ElementOwner());
							// Tell the relative dispnode cache to invalidate the
							// requested region of the relative element
							pRDNC->Invalidate(pRelativeNode->Element(), &aryRects[0], aryRects.Size());
						}
					}
					else
					{
						Invalidate(&aryRects[0], aryRects.Size());
					}
				}
				//  Otherwise, if a dirty region exists, extend the dirty region to encompass it
				//  NOTE: Requests within unmeasured regions are handled automatically during
				//        the measure
				else if(IsDirty())
				{
					_dtr.Accumulate(pnf, GetContentFirstCp(), GetContentLastCp(), FALSE);
				}
			}
		}
		// Handle z-order and position change notifications
		else if(IsPositionNotification(pnf))
		{
			fHandle = HandlePositionNotification(pnf);
		}
		// Handle translated ranges
		else if(pnf->IsType(NTYPE_TRANSLATED_RANGE))
		{
			Assert(pnf->IsDataValid());
			HandleTranslatedRange(pnf->DataAsSize());
		}
		// Handle z-parent changes
		else if(pnf->IsType(NTYPE_ZPARENT_CHANGE))
		{
			if(!ElementOwner()->IsPositionStatic())
			{
				ElementOwner()->ZChangeElement();
			}

			else if(_fContainsRelative)
			{
				ZChangeRelDispNodes();
			}
		}
		// Handle changes to CSS display and visibility
		else if(pnf->IsType(NTYPE_DISPLAY_CHANGE) || pnf->IsType(NTYPE_VISIBILITY_CHANGE))
		{
			HandleVisibleChange(pnf->IsType(NTYPE_VISIBILITY_CHANGE));

			if(_fContainsRelative)
			{
				if(pnf->IsType(NTYPE_VISIBILITY_CHANGE))
				{
					_dp.EnsureDispNodeVisibility();
				}
				else
				{
					_dp.HandleDisplayChange();
				}
			}
		}
		// Insert adornments as needed
		else if(pnf->IsType(NTYPE_ELEMENT_ADD_ADORNER))
		{
			fHandle = HandleAddAdornerNotification(pnf);
		}
		// Otherwise, accumulate the information
		else if(pnf->IsTextChange() || pnf->IsLayoutChange())
		{
			long cpFirst     = GetContentFirstCp();
			long cpLast      = GetContentLastCp();
			BOOL fIsAbsolute = FALSE;

			Assert(!IsLocked());

			// Always dirty the layout of resizing/morphing elements
            if(pnf->Element() && pnf->Element()!=ElementOwner()
                && (pnf->IsType(NTYPE_ELEMENT_RESIZE)
                || pnf->IsType(NTYPE_ELEMENT_RESIZEANDREMEASURE)))
            {
				pnf->Element()->DirtyLayout(pnf->LayoutFlags());

				// For absolute elements, simply note the need to re-calc them
				if(pnf->Element()->IsAbsolute())
				{
					fIsAbsolute = TRUE;

					QueueRequest(CRequest::RF_MEASURE, pnf->Element());
				}
			}

			// Otherwise, collect the range covered by the notification
			if(((!fIsAbsolute
				|| pnf->IsType(NTYPE_ELEMENT_RESIZEANDREMEASURE))
				|| pnf->Element()==ElementOwner())
				&& pnf->Cp(cpFirst)>=0)
			{
				// Accumulate the affected range
				_dtr.Accumulate(
                    pnf,
					cpFirst,
					cpLast,
					(pnf->Element()==ElementOwner()&&!fRangeSet));

				// Content's are dirtied so reset the minmax flag on the display
				_dp._fMinMaxCalced = FALSE;

				// Mark forced layout if requested
				if(pnf->IsFlagSet(NFLAGS_FORCE))
				{
					if(pnf->Element() == ElementOwner())
					{
						_fForceLayout = TRUE;
					}
					else
					{
						_fDTRForceLayout = TRUE;
					}
				}

				// Invalidate cached min/max values when content changes size
				if(!_fPreserveMinMax
					&& _fMinMaxValid 
					&& (pnf->IsType(NTYPE_ELEMENT_MINMAX)
					|| (_fContentsAffectSize
					&& (pnf->IsTextChange()
					|| pnf->IsType(NTYPE_ELEMENT_RESIZE)
					|| pnf->IsType(NTYPE_ELEMENT_REMEASURE)
					|| pnf->IsType(NTYPE_ELEMENT_RESIZEANDREMEASURE)
					|| pnf->IsType(NTYPE_CHARS_RESIZE)))))
				{
					ResetMinMax();
				}
			}

			// If the layout is transitioning to dirty for the first time and
			// is not set to get called by its containing layout engine (via CalcSize),
			// post a layout request
			// (For purposes of posting a layout request, transitioning to dirty
			//  means that previously no text changes were recorded and no absolute
			//  descendents needed sizing and now at least of those states is TRUE)
			if(fIsAbsolute
				|| (!pnf->IsFlagSet(NFLAGS_DONOTLAYOUT)
				&& !_fSizeThis
				&& IsDirty()))
			{
				PostLayoutRequest(pnf->LayoutFlags()|LAYOUT_MEASURE);
			}
#ifdef _DEBUG
			else if(!pnf->IsFlagSet(NFLAGS_DONOTLAYOUT) && !_fSizeThis)
			{
				Assert(!IsDirty()
					|| !GetView()->IsActive()
					|| IsDisplayNone()
					|| GetView()->HasLayoutTask(this));
			}
#endif
		}

		// Reset the range if previously set
		if(fRangeSet)
		{
			pnf->ClearTextRange();
		}
	}

	switch(pnf->Type())
	{
	case NTYPE_DOC_STATE_CHANGE_1:
		pnf->Data(&dwData);
		break;

	case NTYPE_ELEMENT_EXITTREE_1:
		Reset(TRUE);
		break;

	case NTYPE_ZERO_GRAY_CHANGE:
		HandleZeroGrayChange(pnf);
		break; 

	case NTYPE_RANGE_ENSURERECALC:
	case NTYPE_ELEMENT_ENSURERECALC:
		fHandle = pnf->Element() != ElementOwner();
		// If the request is for this element and layout is dirty,
		// convert the pending layout call to a dirty range in the parent layout
		// (Processing the pending layout call immediately could result in measuring
		//  twice since the parent may be dirty as well - Converting it into a dirty
		//  range in the parent is only slightly more expensive than processing it
		//  immediately and prevents the double measuring, keeping things in the
		//  right order)
		if(pnf->Element()==ElementOwner() && IsDirty())
		{
			ElementOwner()->ResizeElement();
		}
		// Otherwise, calculate up through the requesting element
		else
		{
			CView* pView = Doc()->GetView();
			long cpFirst = GetContentFirstCp();
			long cpLast = GetContentLastCp();
			long cp;
			long cch;

			// If the requesting element is the element owner,
			// calculate up through the end of the available content
			if(pnf->Element() == ElementOwner())
			{
				cp = GetDisplay()->_dcpCalcMax;
				cch = cpLast - (cpFirst + cp);
			}
			// Otherwise, calculate up through the element
			else
			{
				ElementOwner()->SendNotification(NTYPE_ELEMENT_ENSURERECALC);

				cp = pnf->Cp(cpFirst) - cpFirst;
				cch = pnf->Cch(cpLast);
			}

			if(pView->IsActive())
			{
				CView::CEnsureDisplayTree edt(pView);

				if(!IsRangeBeforeDirty(cp, cch) || GetDisplay()->_dcpCalcMax<=(cp+cch))
				{
					GetDisplay()->WaitForRecalc((cpFirst+cp+cch), -1);
				}

				if(pnf->IsType(NTYPE_ELEMENT_ENSURERECALC) && pnf->Element()!=ElementOwner())
				{
					ProcessRequest(pnf->Element() );
				}
			}
		}
		break;
	}

	// Handle the notification
	if(fHandle && pnf->IsFlagSet(NFLAGS_ANCESTORS))
	{
		pnf->SetHandler(ElementOwner());
	}
}

//+----------------------------------------------------------------------------
//
//  Member:     Reset
//
//  Synopsis:   Reset the channel (clearing any dirty state)
//
//-----------------------------------------------------------------------------
void CFlowLayout::Reset(BOOL fForce)
{
	super::Reset(fForce);
	CancelChanges();
}

void CFlowLayout::Listen(BOOL fListen)
{
	if((unsigned)fListen != _fListen)
	{
		if(_fListen)
		{
			Reset(TRUE);
		}

		_fListen = (unsigned)fListen;
	}
}

BOOL CFlowLayout::IsListening()
{
	return !!_fListen;
}

LONG CFlowLayout::GetMaxLineWidth()
{
	return _dp.GetMaxPixelWidth()-_dp.GetCaret();
}

void CFlowLayout::Detach()
{
	// flushes the region cache and rel line cache.
	_dp.Detach();
	super::Detach();
}

HRESULT CFlowLayout::AllCharsAffected()
{
	CNotification nf;

	nf.CharsResize(
		GetContentFirstCp(),
		GetContentTextLength(),
		GetFirstBranch());
	GetContentMarkup()->Notify(nf);
	return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CommitChanges
//
//  Synopsis:   Commit any outstanding text changes to the display
//
//-------------------------------------------------------------------------
void CFlowLayout::CommitChanges(CCalcInfo* pci)
{
	long cp;
	long cchOld;
	long cchNew;
	BOOL fForce = !!_fDTRForceLayout;

	// Ignore unnecessary or recursive requests
	if(!IsDirty() || (IsDisplayNone() && (Tag()!=ETAG_BODY)))
	{
		goto Cleanup;
	}

	// Reset dirty state (since changes they are now being handled)
	cp = Cp() + GetContentFirstCp();
	cchOld = CchOld();
	cchNew = CchNew();

	CancelChanges();

	DEBUG_ONLY(Lock());

	// Recalculate the display to account for the pending changes
	{
		CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);
		CSaveCalcInfo sci(pci, this);

		if(fForce)
		{
			pci->_grfLayout |= LAYOUT_FORCE;
		}

		GetDisplay()->UpdateView(pci, cp, cchOld, cchNew);
	}

	// Fire a "content changed" notification (to our host)
	OnTextChange();

	DEBUG_ONLY(Unlock());

Cleanup:
	return;
}

//+------------------------------------------------------------------------
//
//  Member:     CancelChanges
//
//  Synopsis:   Cancel any outstanding text changes to the display
//
//-------------------------------------------------------------------------
void CFlowLayout::CancelChanges()
{
	if(IsDirty())
	{
		_dtr.Reset();
	}
	_fDTRForceLayout = FALSE;
}

//+------------------------------------------------------------------------
//
//  Member:     IsCommitted
//
//  Synopsis:   Verify that all changes are committed
//
//-------------------------------------------------------------------------
BOOL CFlowLayout::IsCommitted()
{
	return !IsDirty();
}

void CFlowLayout::ViewChange(BOOL fUpdate)
{
}

HRESULT CFlowLayout::WaitForParentToRecalc(CElement* pElement, CCalcInfo* pci)
{
	HRESULT hr = S_OK;
	LONG cpElementStart, cchElement;
	LONG cpThisStart, cchThis;
	CTreePos* ptpElement;
	CTreePos* ptpThis;

	pElement->GetTreeExtent(&ptpElement, NULL);
	cpElementStart = ptpElement->GetCp();
	cchElement = pElement->GetElementCch() + 2;

	ElementOwner()->GetTreeExtent(&ptpThis, NULL);
	cpThisStart = ptpThis->GetCp();
	cchThis = ElementOwner()->GetElementCch() + 2;

	if(cpElementStart<cpThisStart || cpElementStart+cchElement>cpThisStart+cchThis)
	{
		hr = S_FALSE;
		goto Cleanup;
	}
	if(cchElement < 1)
	{
		hr = S_FALSE;
		goto Cleanup;
	}

	//BUGBUG MERGEFUN (carled) this should be rewritten to take a ptp and thus
	// avoid the expensive GetCp call above.  The need to do this will be driven
	// by the perf numbers
	hr = WaitForParentToRecalc(cpElementStart+cchElement, -1, pci);
	if(hr)
	{
		goto Cleanup;
	}

Cleanup:
	RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     WaitForParentToRecalc
//
//  Synopsis:   Waits for a this site to finish recalcing upto cpMax/yMax.
//              If first waits for all txtsites above this to finish recalcing.
//
//  Params:     [cpMax]: The cp to calculate too
//              [yMax]:  The y position to calculate too
//              [pci]:   The CCalcInfo
//
//  Return:     HRESULT
//
//--------------------------------------------------------------------------
HRESULT CFlowLayout::WaitForParentToRecalc(
	LONG cpMax,		//@parm Position recalc up to (-1 to ignore)
	LONG yMax,		//@parm ypos to recalc up to (-1 to ignore)
	CCalcInfo* pci)
{
	HRESULT hr = S_OK;

	Assert(!TestLock(CElement::ELEMENTLOCK_RECALC));

	if(!TestLock(CElement::ELEMENTLOCK_SIZING))
	{
#ifdef _DEBUG
		// BUGBUG(sujalp): We should never recurse when we are not SIZING.
		// This code to catch the case in which we recurse when we are not
		// SIZING.
		CElement::CLock LockRecalc(ElementOwner(), CElement::ELEMENTLOCK_RECALC);
#endif
		ElementOwner()->SendNotification(NTYPE_ELEMENT_ENSURERECALC);
	}

	if(!_dp.WaitForRecalc(cpMax, yMax, pci))
	{
		hr = S_FALSE;
		goto Cleanup;
	}

Cleanup:
	RRETURN1(hr, S_FALSE);
}

void CFlowLayout::SizeContentDispNode(const SIZE& size, BOOL fInvalidateAll)
{
	if(!_dp._fHasMultipleTextNodes)
	{
		super::SizeContentDispNode(size, fInvalidateAll);
	}
	else
	{
		CDispNode* pDispNode = GetFirstContentDispNode();

		// we better have a content dispnode
		Assert(pDispNode);

		while(pDispNode)
		{
			if(pDispNode->GetDispClient() == this)
			{
				pDispNode->SetSize(CSize(size.cx, size.cy-(pDispNode->GetPosition()).y), fInvalidateAll);
			}
			pDispNode = pDispNode->GetNextSiblingNode(TRUE);
		}
	}
}

void CFlowLayout::RegionFromElement(
									CElement*		pElement,
									CDataAry<RECT>*	paryRects,
									RECT*			prcBound,
									DWORD			dwFlags)
{
	Assert(pElement);
	Assert(paryRects);

	if(!pElement || !paryRects)
	{
		return;
	}

	// Is the element passed the same element with the owner?
	if(_pElementOwner == pElement)
	{
		// call CLayout implementation.
		super::RegionFromElement(pElement, paryRects, prcBound, dwFlags);
	}
	else
	{
		// BUGBUG: [FerhanE]
		//      Change this part to support relative coordinates, after Srini implements that,
		//
		// Delegate the call to the CDisplay implementation
		_dp.RegionFromElement(
			pElement,          // the element
			paryRects,         // rects returned here
			NULL,              // offset the rects returned
			NULL,              // ask RFE to get CFormDrawInfo
			dwFlags,           // coord w/ respect to the client rc.
			-1,                // Get the complete focus
			-1,                //
			prcBound);         // bounds of the element!
	}
}

CLayout* CFlowLayout::GetFirstLayout(DWORD_PTR* pdw, BOOL fBack, BOOL fRaw)
{
	Assert(!fRaw);

	if(ElementOwner()->GetFirstBranch())
	{
		CChildIterator* pLayoutIterator = new CChildIterator(
			ElementOwner(),
			NULL,
			CHILDITERATOR_USELAYOUT);
		*pdw = (DWORD_PTR)pLayoutIterator;

		return *pdw==NULL?NULL:CFlowLayout::GetNextLayout(pdw, fBack, fRaw);
	}
	else
	{
		// If CTxtSite is not in the tree, no need to walk through
		// CChildIterator
		* pdw = 0;
		return NULL;
	}
}

//+------------------------------------------------------------------------
//
//  Member:     GetNextLayout
//
//  Synopsis:   Enumeration method to loop thru children
//
//  Arguments:  [pdw]       cookie to be used in further enum
//              [fBack]     go from back
//
//  Returns:    Layout
//
//-------------------------------------------------------------------------
CLayout* CFlowLayout::GetNextLayout(DWORD_PTR* pdw, BOOL fBack, BOOL fRaw)
{
	CLayout* pLayout = NULL;

	Assert(!fRaw);

	{
		CChildIterator* pLayoutWalker = (CChildIterator*)(*pdw);
		if(pLayoutWalker)
		{
			CTreeNode* pNode = fBack ? pLayoutWalker->PreviousChild()
				: pLayoutWalker->NextChild();
			pLayout = pNode ? pNode->GetUpdatedLayout() : NULL;
		}
	}
	return pLayout;
}

//+---------------------------------------------------------------------------
//
// Member:      ContainsChildLayout
//
//----------------------------------------------------------------------------
BOOL CFlowLayout::ContainsChildLayout(BOOL fRaw)
{
	Assert(!fRaw);

	{
		DWORD_PTR dw;
		CLayout* pLayout = GetFirstLayout(&dw, FALSE, fRaw);
		ClearLayoutIterator(dw, fRaw);
		return pLayout?TRUE:FALSE;
	}
}

//+---------------------------------------------------------------------------
//
// Member:      ClearLayoutIterator
//
//----------------------------------------------------------------------------
void CFlowLayout::ClearLayoutIterator(DWORD_PTR dw, BOOL fRaw)
{
	if(!fRaw)
	{
		CChildIterator* pLayoutWalker = (CChildIterator*)dw;
		if(pLayoutWalker)
		{
			delete pLayoutWalker;
		}
	}
}

HRESULT CFlowLayout::GetElementsInZOrder(
	CPtrAry<CElement*>*	paryElements,
	CElement*			pElementThis,
	RECT*				prcDraw,
	HRGN				hrgn,
	BOOL				fIncludeNotVisible/*==FALSE*/)
{
	CLayout*	pLayout;
	DWORD_PTR	dw = 0;
	BOOL		fRegionFound = FALSE;
	HRESULT		hr = S_OK;
	CDocument*  pDoc;

	Assert(pElementThis);

	// First, add all our direct children that intersect the draw rect
	// (This includes in-flow sites and those with relative positioning)
	if(pElementThis == ElementOwner())
	{
		for(pLayout=GetFirstLayout(&dw); pLayout; pLayout=GetNextLayout(&dw))
		{
			// If this site is positioned then it will be added in the lower loop.
			if(pLayout->GetFirstBranch()->IsPositionStatic())
			{
				hr = paryElements->Append(pLayout->ElementOwner());
				if(hr)
				{
					goto Cleanup;
				}
			}

		}
	}

	// Next, add absolutely positioned elements for which this site is the
	// RenderParent.
	pDoc = Doc();
	if(pDoc->_fRegionCollection && ElementOwner()->IsZParent())
	{
		long		cpElemStart;
		long		cpMax = GetDisplay()->GetMaxCpCalced();
		long		lIndex;
		long		lArySize;
		CElement*	pElement;
		CTreeNode*	pNode;
		CCollectionCache* pCollectionCache;

		hr = pDoc->PrimaryMarkup()->EnsureCollectionCache(CMarkup::REGION_COLLECTION);
		if(!hr)
		{
			pCollectionCache = pDoc->PrimaryMarkup()->CollectionCache();

			lArySize = pCollectionCache->SizeAry(CMarkup::REGION_COLLECTION);

			for(lIndex=0; lIndex<lArySize; lIndex++)
			{
				hr = pCollectionCache->GetIntoAry(CMarkup::REGION_COLLECTION, lIndex, &pElement);

				Assert(pElement);
				pNode = pElement->GetFirstBranch();

				cpElemStart = pElement->GetFirstCp();

				if(!pElement->IsPositionStatic() && pElementThis==pNode->ZParent())
				{
					CLayout* pLayoutParent;

					pLayout = pElement->GetUpdatedLayout();

					if(pLayout)
					{
						if(!fIncludeNotVisible && !pLayout->ElementOwner()->IsVisible(FALSE))
						{
							continue;
						}
					}
					else
					{
						const CCharFormat* pCF = pNode->GetCharFormat();

						// BUGBUG -- will this mess up visibility:visible
						// for child elements of pElement? (lylec)
						if(pCF->IsVisibilityHidden() || pCF->IsDisplayNone())
						{
							continue;
						}
					}

					pLayoutParent = pElement->GetUpdatedParentLayout();

					// (srinib) Table lies to its parent layout that it is calced until
					// it sees an end table. So tablecell's containing positioned elements
					// have to take care of this issue, because they may not be calced yet.

					// If the parent layout has not calc the current element
					if(cpMax<=cpElemStart ||
						(pLayoutParent &&
						pLayoutParent->IsFlowLayout() &&
						pLayoutParent->IsFlowLayout()->GetDisplay()->GetMaxCpCalced()<=cpElemStart))
					{
						continue;
					}

					fRegionFound = TRUE;

					hr = paryElements->Append(pElement);
					if(hr)
					{
						goto Cleanup;
					}
				}
			}
		}
	}

	if(fRegionFound)
	{
		qsort(*paryElements, paryElements->Size(),
			sizeof(CElement*), CompareElementsByZIndex);
	}

Cleanup:
	ClearLayoutIterator(dw, FALSE);
	RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     SetZOrder
//
//  Synopsis:   set z order for site
//
//  Arguments:  [pLayout]   set z order for this layout
//              [zorder]    to set
//              [fUpdate]   update windows and invalidate
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CFlowLayout::SetZOrder(CLayout* pLayout, LAYOUT_ZORDER zorder, BOOL fUpdate)
{
	HRESULT hr = S_OK;

	if(Doc()->TestLock(FORMLOCK_ARYSITE))
	{
		hr = E_FAIL;
		goto Cleanup;
	}

	if(fUpdate)
	{
		Doc()->FixZOrder();

		Invalidate();
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CFlowLayout::ScrollRangeIntoView(
	long		cpMin,
	long		cpMost,
	SCROLLPIN	spVert,
	SCROLLPIN	spHorz,
	BOOL		fScrollBits)
{
	extern void BoundingRectForAnArrayOfRectsWithEmptyOnes(RECT* prcBound, CDataAry<RECT>*paryRects);

	HRESULT hr = S_OK;

	if(_pDispNode && cpMin>=0)
	{
		CStackDataAry<RECT, 5>	aryRects;
		CRect					rc;
		CCalcInfo				CI(this);

		hr = WaitForParentToRecalc(cpMost, -1, &CI);
		if(hr)
		{
			goto Cleanup;
		}

		_dp.RegionFromElement(
			ElementOwner(),    // the element
			&aryRects,         // rects returned here
			NULL,              // offset the rects
			NULL,              // ask RFE to get CFormDrawInfo
			RFE_SCROLL_INTO_VIEW, // coord w/ respect to the display and not the client rc
			cpMin,             // give me the rects from here ..
			cpMost,            // ... till here
			NULL);             // dont need bounds of the element!

		// Calculate and return the total bounding rect
		BoundingRectForAnArrayOfRectsWithEmptyOnes(&rc, &aryRects);

		if(spVert == SP_TOPLEFT)
		{
			// Though RegionFromElement has already called WaitForRecalc,
			// it calls it until top is recalculated. In order to scroll
			// ptStart to the to of the screen, we need to wait until
			// another screen is recalculated.
			if(!_dp.WaitForRecalc(-1, rc.top+_dp.GetViewHeight()))
			{
				hr = E_FAIL;
				goto Cleanup;
			}
		}

		ScrollRectIntoView(rc, spVert, spHorz, fScrollBits);
		hr = S_OK;
	}
	else
	{
		hr = super::ScrollRangeIntoView(cpMin, cpMost, spVert, spHorz, fScrollBits);
	}


Cleanup:
	RRETURN1(hr, S_FALSE);
}

HTC CFlowLayout::BranchFromPoint(
								 DWORD				dwFlags,
								 POINT				pt,
								 CTreeNode**		ppNodeBranch,
								 HITTESTRESULTS*	presultsHitTest,
								 BOOL				fNoPseudoHit,
								 CDispNode*			pDispNode)
{
	HTC			htc = HTC_YES;
	CLinePtr	rp(&_dp);
	CTreePos*	ptp = NULL;
	LONG		cp, ili, yLine;
	DWORD		dwCFPFlags = 0;
	CRect		rc(pt, pt);
	long		iliStart = -1;
	long		iliFinish = -1;


	dwCFPFlags |= (dwFlags&HT_ALLOWEOL) ? CDisplay::CFP_ALLOWEOL : 0;
	dwCFPFlags |= !(dwFlags&HT_DONTIGNOREBEFOREAFTER) ? CDisplay::CFP_IGNOREBEFOREAFTERSPACE : 0;
	dwCFPFlags |= !(dwFlags&HT_NOEXACTFIT) ? CDisplay::CFP_EXACTFIT : 0;
	dwCFPFlags |= fNoPseudoHit ? CDisplay::CFP_NOPSEUDOHIT : 0;

	Assert(ElementOwner()->IsVisible(FALSE) || ElementOwner()->Tag()==ETAG_BODY);

	*ppNodeBranch = NULL;

	// if the current layout has multiple text nodes then compute the
	// range of lines the belong to the current dispNode
	if(_dp._fHasMultipleTextNodes && pDispNode)
	{
		GetTextNodeRange(pDispNode, &iliStart, &iliFinish);
	}

	if(pDispNode == GetElementDispNode())
	{
		GetClientRect(&rc); 
		rc.MoveTo(pt);
	}

	ili = _dp.LineFromPos(rc, &yLine, &cp,
		CDisplay::LFP_ZORDERSEARCH|
		CDisplay::LFP_IGNORERELATIVE|
		CDisplay::LFP_IGNOREALIGNED|
		(fNoPseudoHit?CDisplay::LFP_EXACTLINEHIT:0),
		iliStart,
		iliFinish);
	if(ili < 0)
	{
		htc = HTC_NO;
		goto Cleanup;
	}

	if((cp=_dp.CpFromPointEx(ili, yLine, cp, pt, &rp, &ptp, NULL, dwCFPFlags,
		&presultsHitTest->_fRightOfCp, &presultsHitTest->_fPseudoHit,
		&presultsHitTest->_cchPreChars, NULL)) == -1) // fExactFit=TRUE to look at whole characters
	{
		htc = HTC_NO;
		goto Cleanup;
	}

	if(cp<GetContentFirstCp() || cp>GetContentLastCp())
	{
		htc = HTC_NO;
		goto Cleanup;
	}

	presultsHitTest->_cpHit  = cp;
	presultsHitTest->_iliHit = rp;
	presultsHitTest->_ichHit = rp.RpGetIch();

    if(pDispNode)
    {
        pt.y += pDispNode->GetPosition().y;
    }

    htc = BranchFromPointEx(pt, rp, ptp, NULL, ppNodeBranch, presultsHitTest->_fPseudoHit,
        &presultsHitTest->_fWantArrow,
        !(dwFlags&HT_DONTIGNOREBEFOREAFTER));

Cleanup:
	if(htc != HTC_YES)
	{
		presultsHitTest->_fWantArrow = TRUE;
	}
	return htc;
}

HTC CFlowLayout::BranchFromPointEx(
                                   POINT        pt,
                                   CLinePtr&    rp,
                                   CTreePos*    ptp,
                                   CTreeNode*   pNodeRelative,  // (IN) non-NULL if we are hit-testing a relative element (NOT its flow position)
                                   CTreeNode**  ppNodeBranch,   // (OUT) returns branch that we hit
                                   BOOL         fPseudoHit,     // (IN) if true, text was NOT hit (CpFromPointEx() figures this out)
                                   BOOL*        pfWantArrow,    // (OUT) 
                                   BOOL         bIgnoreBeforeAfter)
{
    const CCharFormat* pCF;
    CTreeNode*  pNode = NULL;
    CElement *  pElementStop = NULL;
    HTC         htc = HTC_YES;
    BOOL        fVisible = TRUE;
    Assert(ptp);

    // If we are on a line which contains an table, and we are not ignoring before and
    // aftrespace, then we want to hit that table...
    if(!bIgnoreBeforeAfter)
    {
        CLine* pli = rp.CurLine();
        if(pli && pli->_fSingleSite && pli->_cch==rp.RpGetIch())
        {
            rp.RpBeginLine();
            _dp.FormattingNodeForLine(_dp.GetFirstCp()+rp.GetCp(), NULL, pli->_cch, NULL, &ptp, NULL);
        }
    }

    // Get the branch corresponding to the cp hit.
    pNode = ptp->GetBranch();

    // If we hit the white space around text, then find the block element
    // that we hit. For example, if we hit in the padding of a block
    // element then it is hit.
    if(bIgnoreBeforeAfter && fPseudoHit)
    {
        CStackDataAry<RECT, 1>  aryRects;
        CMarkup*    pMarkup = GetContentMarkup();
        LONG        cpClipStart;
        LONG        cpClipFinish;
        DWORD       dwFlags = RFE_HITTEST;

        // If a relative node is passed in, the point is already in the
        // co-ordinate system established by the relative element, so
        // pass RFE_IGNORE_RELATIVE to ignore the relative top left when
        // computing the region.
        if(pNodeRelative)
        {
            dwFlags |= RFE_IGNORE_RELATIVE;
            pNode   =  pNodeRelative;
        }

        cpClipStart = cpClipFinish = GetContentFirstCp();
        rp.RpBeginLine();
        cpClipStart += rp.GetCp();
        rp.RpEndLine();
        cpClipFinish += rp.GetCp();

        // walk up the tree and find the block element that we hit.
        while(pNode && !SameScope(pNode, ElementContent()))
        {
            if(!pNodeRelative)
            {
                pNode = pMarkup->SearchBranchForBlockElement(pNode, this);
            }

            if(!pNode)
            {
                // this is bad, somehow, a pNodeRelative was passed in for
                // an element that is not under this flowlayout. How to interpret
                // this? easiest (and safest) is to just bail
                htc = HTC_NO;
                goto Cleanup;
            }

            _dp.RegionFromElement(pNode->Element(), &aryRects, NULL,
                NULL, dwFlags, cpClipStart, cpClipFinish);

            if(PointInRectAry(pt, aryRects))
            {
                break;
            }
            else if(pNodeRelative)
            {
                htc = HTC_NO;
            }

            if(pNodeRelative || SameScope(pNode, ElementContent()))
            {
                break;
            }

            pNode = pNode->Parent();
        }

        *pfWantArrow = TRUE;
    }

    Assert(pNode);

    // pNode now points to the element we hit, but it might be
    // hidden.  If it's hidden, we need to walk up the parent
    // chain until we find an ancestor that isn't hidden,
    // or until we hit the layout owner or a relative element
    // (we want testing to stop on relative elements because
    // they exist in a different z-plane).  Note that
    // BranchFromPointEx may be called for hidden elements that
    // are inside a relative element.
    pElementStop = pNodeRelative ? pNodeRelative->Element() : ElementContent();

    while(DifferentScope(pNode, pElementStop))
    {
        pCF = pNode->GetCharFormat();
        if(pCF->IsDisplayNone() || pCF->IsVisibilityHidden())
        {
            fVisible = FALSE;
            pNode = pNode->Parent();
        }
        else
        {
            break;
        }
    }

    Assert(pNode);

    // if we hit the layout element and it is a pseudo hit or
    // if the element hit is not visible then consider it a miss

    // We want to show an arrow if we didn't hit text (fPseudoHit TRUE) OR
    // if we did hit text, but it wasn't visible (as determined by the loop
    // above which set fVisible FALSE).
    if(fPseudoHit || !fVisible)
    {
        // If we walked all the way up to the container, then we want
        // to return HTC_NO so the display tree will call HitTestContent
        // on the container's dispnode (i.e. "the background"), which will
        // return HTC_YES.

        // BUGBUG: relative elements need to be looked at

        // If it's relative, then htc was set earlier, so don't
        // touch it now.
        if(!pNodeRelative)
        {
            htc = SameScope(pNode, ElementContent()) ? HTC_NO : HTC_YES;
        }

        *pfWantArrow  = TRUE;
    }
    else
    {
        *pfWantArrow  = !!fPseudoHit;
    }

Cleanup:
    *ppNodeBranch = pNode;

    return htc;
}

void CFlowLayout::Draw(CFormDrawInfo* pDI, CDispNode* pDispNode)
{
	// CFlowLayout and text rendering know how to deal with device coordinates,
	// which are necessary for Win 9x because of the 16-bit coordinate limitation
	// in GDI.
	pDI->SetDeviceCoordinateMode();    

	GetDisplay()->Render(pDI, pDI->_rc, pDI->_rcClip, pDispNode);
}

void CFlowLayout::ShowSelected(CTreePos* ptpStart, CTreePos* ptpEnd, BOOL fSelected,  BOOL fLayoutCompletelyEnclosed, BOOL fFireOM)
{
	Assert(ptpStart && ptpEnd && ptpStart->GetMarkup()==ptpStart->GetMarkup());
	CElement* pElement = ElementOwner();

	// If this has a slave markup, but the selection is in the main markup, then
	// select the element (as opposed to part or all of its content)
	if(pElement->HasSlaveMarkupPtr()
		&& ptpStart->GetMarkup()!=ElementOwner()->GetSlaveMarkupPtr())
	{
		SetSelected(fSelected, TRUE);
	}
	else
	{
		if((pElement->_etag==ETAG_BUTTON) || (pElement->_etag==ETAG_TEXTAREA))
		{
			if((fSelected&&fLayoutCompletelyEnclosed) || (!fSelected&&!fLayoutCompletelyEnclosed))
			{
				SetSelected(fSelected, TRUE );
			}
			else
			{
				_dp.ShowSelected(ptpStart, ptpEnd, fSelected);
			}
		}
		else
		{
			_dp.ShowSelected(ptpStart, ptpEnd, fSelected);
		}
	}
}

HRESULT CFlowLayout::NotifyKillSelection()
{
	CDocument* pDoc = Doc();

	// If our doc is not the only doc, then we need notify
	// the other docs so that they can kill their selections
	if(pDoc != Doc())
	{
		CNotification nf;

		nf.KillSelection(pDoc->GetPrimaryElementClient());
		pDoc->BroadcastNotify(&nf);
	}
	return S_OK;
}

void CFlowLayout::ResizePercentHeightSites()
{
	CNotification	nf;
	CLayout*		pLayout;
	DWORD_PTR		dw;
	BOOL			fFoundAtLeastOne = FALSE;

	// If no contained sites are affected, immediately return
	if(!_fChildHeightPercent)
	{
		return;
	}

	// Otherwise, iterate through all sites, sending an ElementResize notification for those affected
	// (Also, note that resizing a percentage height site cannot affect min/max values)
	// NOTE: With "autoclear", the min/max can vary after resizing percentage sized
	//       descendents. However, the calculated min/max values, which used for table
	//       sizing, should take into account those changes since doing so would likely
	//       break how tables layout relative to Netscape. (brendand)
	Assert(!_fPreserveMinMax);
	_fPreserveMinMax = TRUE;

	for(pLayout=GetFirstLayout(&dw); pLayout; pLayout=GetNextLayout(&dw))
	{
		if(pLayout->PercentHeight())
		{
			nf.ElementResize(pLayout->ElementOwner(), NFLAGS_CLEANCHANGE);
			GetContentMarkup()->Notify(nf);

			fFoundAtLeastOne = TRUE;
		}
	}
	ClearLayoutIterator(dw, FALSE);

	_fPreserveMinMax = FALSE;

	// clear the flag if there was no work done.  oppurtunistic cleanup
	_fChildHeightPercent = fFoundAtLeastOne ;
}

void CFlowLayout::ResetMinMax()
{
	_fMinMaxValid = FALSE;
	_dp._fMinMaxCalced = FALSE;
	MarkHasAlignedLayouts(FALSE);
}

BOOL CFlowLayout::GetSiteWidth(
							   CLayout*		pLayout,
							   CCalcInfo*	pci,
							   BOOL			fBreakAtWord,
							   LONG			xWidthMax,
							   LONG*		pxWidth,
							   LONG*		pyHeight,
							   INT*			pxMinSiteWidth)

{
	CDocument* pDoc = Doc();
	LONG xLeftMargin, xRightMargin;
	LONG yTopMargin, yBottomMargin;
	SIZE sizeObj;

	Assert(pLayout && pxWidth);

	*pxWidth = 0;

	if(pxMinSiteWidth)
	{
		*pxMinSiteWidth = 0;
	}

	if(pyHeight)
	{
		*pyHeight = 0;
	}

	if(pLayout->IsDisplayNone())
	{
		return FALSE;
	}

	// get the margin info for the site
	pLayout->GetMarginInfo(pci, &xLeftMargin, &yTopMargin, &xRightMargin, &yBottomMargin);

	// measure the site
	if(pDoc->_lRecursionLevel == MAX_RECURSION_LEVEL)
	{
		AssertSz(0, "Max recursion level reached!");
		sizeObj.cx = 0;
		sizeObj.cy = 0;
	}
	else
	{
		LONG lRet;

		pDoc->_lRecursionLevel++;
		lRet = MeasureSite(
            pLayout,
			pci,
			xWidthMax - xLeftMargin - xRightMargin,
			fBreakAtWord,
			&sizeObj,
			pxMinSiteWidth);
		pDoc->_lRecursionLevel--;

		if(lRet)
		{
			return TRUE;
		}
	}

	// Propagate the _fAutoBelow bit, if the child is auto positioned or
	// non-zparent children have auto positioned children
	if(!_fAutoBelow)
	{
		const CFancyFormat* pFF = pLayout->GetFirstBranch()->GetFancyFormat();

		if(pFF->IsAutoPositioned() || (!pFF->IsZParent()
			&& (pLayout->_fContainsRelative || pLayout->_fAutoBelow)))
		{
			_fAutoBelow = TRUE;
		}
	}

	// not adjust the size and proposed x pos to include margins
	*pxWidth = max(0L, xLeftMargin+xRightMargin+sizeObj.cx);

	if(pxMinSiteWidth)
	{
		*pxMinSiteWidth += max(0L, xLeftMargin+xRightMargin);
	}

	if (pyHeight)
	{
		*pyHeight = max(0L, sizeObj.cy+yTopMargin+yBottomMargin);
	}

	return FALSE;
}

int CFlowLayout::MeasureSite(
							 CLayout*	pLayout,
							 CCalcInfo*	pci,
							 LONG		xWidthMax,
							 BOOL		fBreakAtWord,
							 SIZE*		psizeObj,
							 int*		pxMinWidth)
{
	CSaveCalcInfo sci(pci);
	LONG lRet = 0;

	Assert(pci->_smMode != SIZEMODE_SET);

	if(!pLayout->ElementOwner()->IsInMarkup())
	{
		psizeObj->cx = psizeObj->cy = 0;
		return lRet;
	}

	// if the layout we are measuring (must be a child of ours)
	// is percent sized, then we should take this oppurtunity
	// to set some work-flags 
	{
		const CFancyFormat* pFF = pLayout->GetFirstBranch()->GetFancyFormat();

		if(pFF->_fHeightPercent)
		{
			_fChildHeightPercent = TRUE;
		}

		if(pFF->_fWidthPercent)
		{
			_fChildWidthPercent = TRUE;
		}
	}

	if(fBreakAtWord)
	{
		long xParentWidth;
		long yParentHeight;

		_dp.GetViewWidthAndHeightForChild(
			pci,
			&xParentWidth,
			&yParentHeight);

		// Set the appropriate parent width
		pci->SizeToParent(xParentWidth, yParentHeight);

		// set available size in sizeObj
		psizeObj->cx = xWidthMax;
		psizeObj->cy = pci->_sizeParent.cy;

		// Ensure the available size does not exceed that of the view
		// (For example, when word-breaking is disabled, the available size
		//  is set exceedingly large. However, percentage sized sites should
		//  still not grow themselves past the view width.)
		if(pci->_smMode==SIZEMODE_NATURAL && pLayout->PercentSize())
		{
			if(pci->_sizeParent.cx < psizeObj->cx)
			{
				psizeObj->cx = pci->_sizeParent.cx;
			}
			if(pci->_sizeParent.cy < psizeObj->cy)
			{
				psizeObj->cy = pci->_sizeParent.cy;
			}
		}

		// If the site is absolutely positioned, only use SIZEMODE_NATURAL
		if(pLayout->ElementOwner()->IsAbsolute())
		{
			pci->_smMode = SIZEMODE_NATURAL;
		}

		// Mark the site for sizing if it is already marked
		// or it is percentage sized and the view size has changed and
		// the site doesn't already know whether to resize.
		if(!pLayout->ElementOwner()->TestClassFlag(CElement::ELEMENTDESC_NOPCTRESIZE))
		{
			pLayout->_fSizeThis = pLayout->_fSizeThis || pLayout->PercentSize();
		}

		pLayout->CalcSize(pci, psizeObj);

		if(pci->_smMode==SIZEMODE_MMWIDTH && pxMinWidth)
		{
			*pxMinWidth = psizeObj->cy;
		}
	}
	else
	{
		pLayout->GetSize(psizeObj);
	}

	return lRet;
}

DWORD CFlowLayout::CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault)
{
	CSaveCalcInfo sci(pci, this);
	CScopeFlag csfCalcing(this);
	BOOL fNormalMode = (pci->_smMode==SIZEMODE_NATURAL
		|| pci->_smMode==SIZEMODE_SET
		|| pci->_smMode==SIZEMODE_FULLSIZE);
	BOOL fRecalcText = FALSE;
	BOOL fWidthChanged, fHeightChanged;
	CSize sizeOriginal;
	DWORD grfReturn;

	Assert(!IsDisplayNone() || Tag()==ETAG_BODY);

	Listen();

	Assert(pci);
	Assert(psize);

	// Set default return values and initial state
	GetSize(&sizeOriginal);

	if(_fForceLayout)
	{
		pci->_grfLayout |= LAYOUT_FORCE;
		_fForceLayout = FALSE;
	}

	grfReturn = (pci->_grfLayout & LAYOUT_FORCE);

	if(pci->_grfLayout & LAYOUT_FORCE)
	{
		_fSizeThis         = TRUE;
		_fAutoBelow        = FALSE;
		_fPositionSet      = FALSE;
		_fContainsRelative = FALSE;
	}
	fWidthChanged  = (fNormalMode ? psize->cx != sizeOriginal.cx : FALSE);
	fHeightChanged = (fNormalMode ? psize->cy != sizeOriginal.cy : FALSE);

	// If height has changed, mark percentage sized children as in need of sizing
	// (Width changes cause a full re-calc and thus do not need to resize each
	//  percentage-sized site)
	if(fNormalMode && !fWidthChanged && !ContainsVertPercentAttr() && fHeightChanged)
	{
		long fContentsAffectSize = _fContentsAffectSize;

		_fContentsAffectSize = FALSE;
		ResizePercentHeightSites();
		_fContentsAffectSize = fContentsAffectSize;
	}

	// For changes which invalidate the entire layout, dirty all of the text
	fRecalcText = (fNormalMode &&
        (IsDirty() || _fSizeThis || fWidthChanged || fHeightChanged))
		|| (pci->_grfLayout&LAYOUT_FORCE)
		|| (pci->_smMode==SIZEMODE_MMWIDTH&&!_fMinMaxValid)
		|| (pci->_smMode==SIZEMODE_MINWIDTH&&!_fMinMaxValid);

	// Cache sizes and recalculate the text (if required)
	if(fRecalcText)
	{
		CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);

		// If dirty, ensure display tree nodes exist
		if(_fSizeThis && fNormalMode
			&& (EnsureDispNode(pci, (grfReturn&LAYOUT_FORCE))==S_FALSE))
		{
			grfReturn |= LAYOUT_HRESIZE|LAYOUT_VRESIZE;
		}

		// Calculate the text
		CalcTextSize(pci, psize, psizeDefault);

		// For normal modes, cache values and request layout
		if(fNormalMode)
		{
			grfReturn |= LAYOUT_THIS |
				(psize->cx!=sizeOriginal.cx?LAYOUT_HRESIZE:0) |
				(psize->cy != sizeOriginal.cy?LAYOUT_VRESIZE:0);

			// If size changes occurred, size the display nodes
			if(_pDispNode && (grfReturn&(LAYOUT_FORCE|LAYOUT_HRESIZE|LAYOUT_VRESIZE)))
			{
				SizeDispNode(pci, *psize);
				SizeContentDispNode(CSize(_dp.GetMaxWidth(), _dp.GetHeight()));
			}

			// Mark the site clean
			_fSizeThis = FALSE;
		}
		// For min/max mode, cache the values and note that they are now valid
		else if(pci->_smMode == SIZEMODE_MMWIDTH)
		{
			_xMax = psize->cx;
			_xMin = psize->cy;
			_fMinMaxValid = TRUE;
		}
		else if(pci->_smMode == SIZEMODE_MINWIDTH)
		{
			_xMin = psize->cx;
		}
	}

	// If any absolutely positioned sites need sizing, do so now
	if(pci->_smMode==SIZEMODE_NATURAL && HasRequestQueue())
	{
		long xParentWidth;
		long yParentHeight;

		_dp.GetViewWidthAndHeightForChild(
			pci,
			&xParentWidth,
			&yParentHeight,
			pci->_smMode==SIZEMODE_MMWIDTH);

		ProcessRequests(pci, CSize(xParentWidth, yParentHeight));
	}

	// Lastly, return the requested size
	switch(pci->_smMode)
	{
	case SIZEMODE_SET:
	case SIZEMODE_NATURAL:
	case SIZEMODE_FULLSIZE:
		Assert(!_fSizeThis);
		GetSize(psize);
		Reset(FALSE);
		Assert(!HasRequestQueue() || GetView()->HasLayoutTask(this));
		break;

	case SIZEMODE_MMWIDTH:
		Assert(_fMinMaxValid);
		psize->cx = _xMax;
		psize->cy = _xMin;
		if(!fRecalcText && psizeDefault)
		{
			GetSize(psizeDefault);
		}
		break;

	case SIZEMODE_MINWIDTH:
		psize->cx = _xMin;
		psize->cy = 0;
		break;
	}

	return grfReturn;
}

//+-------------------------------------------------------------------------
//
//  Method:     CFlowLayout::CalcTextSize
//
//  Synopsis:   Calculate the size of the contained text
//
//--------------------------------------------------------------------------
void CFlowLayout::CalcTextSize(
						  CCalcInfo*	pci,
						  SIZE*			psize,
						  SIZE*			psizeDefault)
{
	BOOL	fNormalMode = (pci->_smMode==SIZEMODE_NATURAL || pci->_smMode==SIZEMODE_SET);
	BOOL	fFullRecalc;
	CRect	rcView;
	long	cxView, cyView;
	long	cxAdjustment = 0;
	long	cyAdjustment = 0;
	long	yViewHeightOld = _dp.GetViewHeight();

	// Hidden layouts should just accumulate changes, and
	// are measured when unhidden.
	Assert(!IsDisplayNone() || Tag()==ETAG_BODY);

	_dp._fRTL = (GetFirstBranch()->GetCascadedBlockDirection() == styleDirRightToLeft);

	// Adjust the incoming size for max/min width requests
	if(pci->_smMode == SIZEMODE_MMWIDTH)
	{
		psize->cx = lMaximumWidth;
		psize->cy = lMaximumWidth;
	}
	else if(pci->_smMode == SIZEMODE_MINWIDTH)
	{
		psize->cx = 1;
		psize->cy = lMaximumWidth;
	}

	// Construct the "view" rectangle from the available size
	// Also, determine the amount of space to allow for things external to the view
	// (For sites which size to their content, calculate the full amount;
	//  otherwise, simply take that which is left over after determining the
	//  view size. Additionally, ensure their view size is never less than 1
	//  pixel, since recalc will not take place for smaller sizes.)
	if(_fContentsAffectSize)
	{
		long lMinimum = (pci->_smMode==SIZEMODE_MINWIDTH ? 1 : 0);

		rcView.top    = 0;
		rcView.left   = 0;
		rcView.right  = 0x7FFFFFFF;
		rcView.bottom = 0x7FFFFFFF;

		GetClientRect(&rcView, CLIENTRECT_USERECT, pci);

		cxAdjustment  = 0x7FFFFFFF - (rcView.right-rcView.left);
		cyAdjustment  = 0x7FFFFFFF - (rcView.bottom-rcView.top);

		rcView.right  = rcView.left + max(lMinimum, psize->cx-cxAdjustment);
		rcView.bottom = rcView.top  + max(lMinimum, psize->cy-cyAdjustment);
	}
	else
	{
		rcView.top    = 0;
		rcView.left   = 0;
		rcView.right  = psize->cx;
		rcView.bottom = psize->cy;

		GetClientRect(&rcView, CLIENTRECT_USERECT, pci);
	}

	cxView = max(0L, rcView.right-rcView.left);
	cyView = max(0L, rcView.bottom-rcView.top);

	if(!_fContentsAffectSize)
	{
		cxAdjustment = psize->cx - cxView;
		cyAdjustment = psize->cy - cyView;
	}

	// Determine if a full recalc of the text is necessary
	// NOTE: SetViewSize must always be called first
	fFullRecalc = _dp.SetViewSize(rcView)
		|| (ContainsVertPercentAttr()
		&& _dp.GetViewHeight()!=yViewHeightOld)
		|| !fNormalMode
		|| (pci->_grfLayout&LAYOUT_FORCE);


	if(fFullRecalc)
	{
		CSaveCalcInfo sci(pci, this);
		BOOL fWordWrap = _dp.GetWordWrap();

		if(_fDTRForceLayout)
		{
			pci->_grfLayout |= LAYOUT_FORCE;
		}

		// If the text will be fully recalc'd, cancel any outstanding changes
		CancelChanges();

		if(pci->_smMode!=SIZEMODE_MMWIDTH && pci->_smMode!=SIZEMODE_MINWIDTH)
		{
			long xParentWidth;
			long yParentHeight;

			_dp.GetViewWidthAndHeightForChild(
				pci,
				&xParentWidth,
				&yParentHeight,
				pci->_smMode==SIZEMODE_MMWIDTH);

			// BUGBUG(SujalP, SriniB and BrendanD): These 2 lines are really needed here
			// and in all places where we instantiate a CI. However, its expensive right
			// now to add these lines at all the places and hence we are removing them
			// from here right now. We have also removed them from CDisplay::UpdateView()
			// for the same reason. (Bug#s: 58809, 62517, 62977)
			pci->SizeToParent(xParentWidth, yParentHeight);
		}

		if(pci->_smMode == SIZEMODE_MMWIDTH)
		{
			_dp.SetWordWrap(FALSE);
		}

		if(fNormalMode
			&& !(pci->_grfLayout&LAYOUT_FORCE)
			&& !ContainsHorzPercentAttr() // if we contain a horz percent attr, then RecalcLineShift() is insufficient; we need to do a full RecalcView()
			&& _fMinMaxValid
			&& _dp._fMinMaxCalced
			&& cxView>=_dp._xMaxWidth
			&& !_fSizeToContent)
		{
			Assert(_dp._xWidthView == cxView);
			Assert(!_fChildWidthPercent);
			Assert(!ContainsChildLayout());

			_dp.RecalcLineShift(pci, pci->_grfLayout);
		}
		else
		{
			_fAutoBelow        = FALSE;
			_fContainsRelative = FALSE;
			_dp.RecalcView(pci, fFullRecalc);
		}

		// Inval since we are doing a full recalc
		Invalidate();

		if(fNormalMode)
		{
			_dp._fMinMaxCalced = FALSE;
		}

		if(pci->_smMode == SIZEMODE_MMWIDTH)
		{
			_dp.SetWordWrap(fWordWrap);
		}
	}
	// If only a partial recalc is necessary, commit the changes
	else if(!IsCommitted())
	{
		Assert(pci->_smMode != SIZEMODE_MMWIDTH);
		Assert(pci->_smMode != SIZEMODE_MINWIDTH);

		CommitChanges(pci);
	}

	// Propagate CCalcInfo state as appropriate
	if(fNormalMode)
	{
		CLine*		pli;
		unsigned	i;

		// For normal calculations, determine the baseline of the "first" line
		// (skipping over any aligned sites at the beginning of the text)
		pci->_yBaseLine = 0;

		for(i=0; i<_dp.Count(); i++)
		{
			pli = _dp.Elem(i);
			if(!pli->IsFrame())
			{
				// BUGBUG: Is this correct given the new meaning of the CLine members? (brendand)
				pci->_yBaseLine = pli->_yHeight - pli->_yDescent;
				break;
			}
		}
	}

	// Determine the size from the calculated text
	if(pci->_smMode != SIZEMODE_SET)
	{
		switch(pci->_smMode)
		{
		case SIZEMODE_FULLSIZE:
		case SIZEMODE_NATURAL:
			if(_fContentsAffectSize || pci->_smMode==SIZEMODE_FULLSIZE)
			{
				_dp.GetSize(psize);
				((CSize*)psize)->Max(((CRect &)rcView).Size());
			}
			else
			{
				psize->cx = cxView;
				psize->cy = cyView;
			}
			break;

		case SIZEMODE_MMWIDTH:
			{
				psize->cx = _dp.GetWidth();
				psize->cy = _dp._xMinWidth;

				if(psizeDefault)
				{
					psizeDefault->cx = _dp.GetWidth() + cxAdjustment;
					psizeDefault->cy = _dp.GetHeight() + cyAdjustment;
				}

				break;
			}
		case SIZEMODE_MINWIDTH:
			{

				psize->cx = _dp.GetWidth();
				psize->cy = _dp.GetHeight();
				_dp.FlushRecalc();
				break;
			}

#ifdef _DEBUG
		default:
			AssertSz(0, "CFlowLayout::CalcTextSize: Unknown SIZEMODE_xxxx");
			break;
#endif
		}

		psize->cx += cxAdjustment;
		psize->cy += (pci->_smMode==SIZEMODE_MMWIDTH ? cxAdjustment : cyAdjustment);
	}
}

void CFlowLayout::GetPositionInFlow(CElement* pElement, CPoint* ppt)
{
	CLinePtr rp(&_dp);
	CTreePos* ptpStart;

	Assert(pElement);
	Assert(ppt);

	if(pElement->IsRelative() && !pElement->NeedsLayout())
	{
		GetFlowPosition(pElement, ppt);
	}
	else
	{
		ppt->x = ppt->y = 0;

		// get the tree extent of the element of the layout passed in
		pElement->GetTreeExtent(&ptpStart, NULL);

		if(_dp.RenderedPointFromTp(ptpStart->GetCp(), ptpStart, FALSE, *ppt, &rp, TA_TOP) < 0)
		{
			return;
		}

		// if the NODE is RTL, mirror the point
		if(_dp.IsRTL())
		{
			ppt->x = -ppt->x;
		}

		if(pElement->NeedsLayout())
		{
			ppt->y += pElement->GetLayoutPtr()->GetYProposed();
		}
		ppt->y += rp->GetYTop();
	}
}

HRESULT CFlowLayout::GetChildElementTopLeft(POINT& pt, CElement* pChild)
{
	Assert(pChild && !pChild->NeedsLayout());

	// handle a couple special cases. we won't hit
	// these when coming in from the OM, but if this fx is
	// used internally, we might. so here they are
	switch(pChild->Tag())
	{
	case ETAG_AREA:
		{
			AssertSz(FALSE, "must improve");
		}
		break;

	default:
		{
			CTreePos*	ptpStart;
			CTreePos*	ptpEnd;
			LONG		cpStart;
			ELEMENT_TAG	etag = ElementOwner()->Tag();

			pt.x = pt.y = -1;

			// get the extent of this element
			pChild->GetTreeExtent(&ptpStart, &ptpEnd);

			if(!ptpStart || !ptpEnd)
			{
				goto Cleanup;
			}

			cpStart = ptpStart->GetCp();

			{
				CStackDataAry<RECT, 1> aryRects;

				_dp.RegionFromElement(pChild, &aryRects, NULL, NULL, RFE_ELEMENT_RECT, cpStart, cpStart);

				if(aryRects.Size())
				{
					pt.x = aryRects[0].left;
					pt.y = aryRects[0].top;
				}
			}

			// if we are for a table cell, then we need to adjust for the cell insets,
			// in case the content is vertically aligned.
			if((etag==ETAG_TD) || (etag==ETAG_TH) || (etag==ETAG_CAPTION))
			{
				CDispNode* pDispNode = GetElementDispNode();
				if(pDispNode && pDispNode->HasInset())
				{
					const CSize& sizeInset = pDispNode->GetInset();
					pt.x += sizeInset.cx;
					pt.y += sizeInset.cy;
				}
			}
		}
		break;
	}

Cleanup:
	return S_OK;
}

CDispNode* CFlowLayout::GetElementDispNode(CElement* pElement) const
{
	return (!pElement
		|| pElement==ElementOwner()
		? super::GetElementDispNode(pElement)
		: pElement->IsRelative()
		? ((CFlowLayout*)this)->_dp.FindElementDispNode(pElement)
		: NULL);
}

void CFlowLayout::SetElementDispNode(CElement* pElement, CDispNode* pDispNode)
{
	if(!pElement || pElement==ElementOwner())
	{
		super::SetElementDispNode(pElement, pDispNode);
	}
	else
	{
		Assert(pElement->IsRelative());
		Assert(!pElement->HasLayout());

		_dp.SetElementDispNode(pElement, pDispNode);
	}
}

void CFlowLayout::GetContentSize(CSize* psize, BOOL fActualSize) const
{
	if(fActualSize)
	{
		psize->cx = _dp.GetWidth();
		psize->cy = _dp.GetHeight();
	}
	else
	{
		super::GetContentSize(psize, fActualSize);
	}
}

void CFlowLayout::GetTextNodeRange(CDispNode* pDispNode, long* piliStart, long* piliFinish)
{
	Assert(pDispNode);
	Assert(piliStart);
	Assert(piliFinish);

	*piliStart = 0;
	*piliFinish = GetDisplay()->LineCount();

	// First content disp node does not have a cookie
	if(pDispNode != GetFirstContentDispNode())
	{
		*piliStart = (LONG)(LONG_PTR)pDispNode->GetExtraCookie();
	}

	pDispNode = pDispNode->GetNextSiblingNode(TRUE);

	while(pDispNode)
	{
		if(this == pDispNode->GetDispClient())
		{
			*piliFinish = (LONG)(LONG_PTR)pDispNode->GetExtraCookie();
			break;
		}
		pDispNode = pDispNode->GetNextSiblingNode(TRUE);
	}
}

#define CX_CONTEXTMENUOFFSET    2
#define CY_CONTEXTMENUOFFSET    2
STDMETHODIMP CFlowLayout::HandleMessage(CMessage* pMessage)
{
	HRESULT		hr = S_FALSE;
	CDocument*  pDoc = Doc();

	BOOL		fLbuttonDown;
	BOOL		fInBrowseMode = !ElementOwner()->IsEditable();
	CDispNode*	pDispNode = GetElementDispNode();
	BOOL		fIsScroller   = (pDispNode && pDispNode->IsScroller());


	// Prepare the message for this layout
	PrepareMessage(pMessage);

	// First, forward mouse messages to the scrollbars (if any)
	// (Keyboard messages are handled below and then notify the scrollbar)
	if(fIsScroller
		&& ((pMessage->htc==HTC_HSCROLLBAR && pMessage->pNodeHit->Element()==ElementOwner())
		|| (pMessage->htc==HTC_VSCROLLBAR && pMessage->pNodeHit->Element()==ElementOwner()))
		&& ((pMessage->message>=WM_MOUSEFIRST
		&& pMessage->message!=WM_MOUSEWHEEL
		&& pMessage->message<=WM_MOUSELAST)
		|| pMessage->message==WM_SETCURSOR))
	{
		hr = HandleScrollbarMessage(pMessage, ElementOwner());
		if(hr != S_FALSE)
		{
			goto Cleanup;
		}
	}

	// In Edit mode, if no element was hit, resolve to the closest element
	if(!fInBrowseMode
		&& !pMessage->pNodeHit
		&& ((pMessage->message>=WM_MOUSEFIRST
		&& pMessage->message<=WM_MOUSELAST)
		|| pMessage->message==WM_SETCURSOR
		|| pMessage->message==WM_CONTEXTMENU)
		// BUGBUG (MohanB) Not sure if it's okay to nuke this pElem check
		)
	{
		CTreeNode* pNode;
		HTC htc;

		pMessage->resultsHitTest._fWantArrow = FALSE;
		pMessage->resultsHitTest._fRightOfCp = FALSE;

		htc = BranchFromPoint(HT_DONTIGNOREBEFOREAFTER,
			pMessage->ptContent,
			&pNode,
			&pMessage->resultsHitTest);

		if(HTC_YES==htc && pNode)
		{
			pMessage->SetNodeHit(pNode);
		}
	}

	switch(pMessage->message)
	{
	case WM_LBUTTONDBLCLK:
	case WM_RBUTTONDBLCLK:
		hr = HandleButtonDblClk(pMessage);
		break;

	case WM_MOUSEMOVE:
		fLbuttonDown = !!(GetKeyState(VK_LBUTTON) & 0x8000);

		if(fLbuttonDown)
		{
			// if we came in with a lbutton down (lbuttonDown = TRUE) and now
			// it is up, it might be that we lost the mouse up event due to a
			// DoDragDrop loop. In this case we have to UI activate ourselves
			if(!(GetKeyState(VK_LBUTTON) & 0x8000))
			{
				hr = ElementOwner()->BecomeUIActive();
				if(hr)
				{
					goto Cleanup;
				}
			}
		}
		break;
	case WM_MOUSEWHEEL:
		hr = HandleMouseWheel(pMessage);
		break;

	case WM_CONTEXTMENU:
		switch(pDoc->GetSelectionType())
		{
		case SELECTION_TYPE_Control:
			// Pass on the message to the first element that is selected
			{
				IHTMLElement*	pIElemSelected	= NULL;
				CElement*		pElemSelected	= NULL;

				if(S_OK!=pDoc->GetCurrentMarkup()->GetElementSegment(0, &pIElemSelected)
					|| !pIElemSelected
					|| S_OK!=pIElemSelected->QueryInterface(CLSID_CElement, (void**)&pElemSelected))
				{
					AssertSz(FALSE, "Cannot get selected element");
					hr = S_OK;
					goto Cleanup;
				}
				// if the site-selected element is different from the owner, pass the message to it
				Assert(pElemSelected);
				if(pElemSelected != ElementOwner())
				{
					// Call HandleMessage directly because we do not bubbling here
					hr = pElemSelected->HandleMessage(pMessage);
				}
			}
			break;

		case SELECTION_TYPE_Selection:
			// Display special menu for text selection in browse mode
			// BUGBUG (MohanB) Should show default menu for HTMLAREA?
			if(!IsEditable(TRUE))
			{
				int cx, cy;

				cx = (short)LOWORD(pMessage->lParam);
				cy = (short)HIWORD(pMessage->lParam);

				if(cx==-1 && cy==-1) // SHIFT+F10
				{
					// Compute position at whcih to display the menu
					IMarkupPointer* pIStart	= NULL;
					IMarkupPointer* pIEnd	= NULL;
					CMarkupPointer* pStart	= NULL;
					CMarkupPointer* pEnd	= NULL;

					if(S_OK==pDoc->CreateMarkupPointer(&pStart)
						&& S_OK==pDoc->CreateMarkupPointer(&pEnd)
						&& S_OK==pStart->QueryInterface(IID_IMarkupPointer, (void**)&pIStart)
						&& S_OK==pEnd->QueryInterface(IID_IMarkupPointer, (void**)&pIEnd)
						&& S_OK==pDoc->GetCurrentMarkup()->MovePointersToSegment(0, pIStart, pIEnd))
					{
						CMarkupPointer* pmpSelMin;
						POINT ptSelMin;

						ReleaseInterface(pIStart);
						ReleaseInterface(pIEnd);

						if(OldCompare(pStart, pEnd) > 0)
						{
							pmpSelMin = pEnd;
						}
						else
						{
							pmpSelMin = pStart;
						}

						if(GetDisplay()->PointFromTp(
							pmpSelMin->GetCp(), NULL, FALSE, FALSE, ptSelMin, NULL, TA_BASELINE) != -1)
						{
							RECT rcWin;

							GetWindowRect(pDoc->InPlace()->_hwnd, &rcWin);
							cx = ptSelMin.x - GetXScroll() + rcWin.left - CX_CONTEXTMENUOFFSET;
							cy = ptSelMin.y - GetYScroll() + rcWin.top  - CY_CONTEXTMENUOFFSET;
						}

						ReleaseInterface(pStart);
						ReleaseInterface(pEnd);
					}
				}
				hr = ElementOwner()->OnContextMenu(cx, cy, CONTEXT_MENU_TEXTSELECT);
			}
			break;
		}
		// For other selection types, let super handle the message
		break;

	case WM_SETCURSOR:
		// Are we over empty region?
		hr = HandleSetCursor(
			pMessage,
			pMessage->resultsHitTest._fWantArrow&&fInBrowseMode);
		// BUGBUG (MohanB) Not sure if it's okay to nuke this pElem check
		break;

	case WM_SYSKEYDOWN:
		hr = HandleSysKeyDown(pMessage);
		break;
	}

	// Remember to call super
	if(hr == S_FALSE)
	{
		hr = super::HandleMessage(pMessage);
	}

Cleanup:
	RRETURN1(hr, S_FALSE);
}

HRESULT CFlowLayout::HandleButtonDblClk(CMessage* pMessage)
{
	// Repaint window to show any exposed portions
	ViewChange(TRUE);
	return S_OK;
}

// helper function: determine how many lines to scroll per mouse wheel
LONG WheelScrollLines()
{
    LONG uScrollLines = 3; // reasonable default

    ::SystemParametersInfo(
        SPI_GETWHEELSCROLLLINES,
        0,
        &uScrollLines,
        0);

    return uScrollLines;
}

// used by CTxtSite::HandleMouseWheel. The "magic" number is similar to
// the one used in CDisplay::VScroll(SB_LINEDOWN, fFixed) and
// CDisplay::VScroll(SB_PAGEDOWN, fFixed).
#define PERCENT_PER_LINE    50
#define PERCENT_PER_PAGE    875

HRESULT CFlowLayout::HandleMouseWheel(CMessage* pMessage)
{
	// WM_MOUSEWHEEL       - line scroll mode
	HRESULT	hr			= S_FALSE;
	BOOL	fControl	= (pMessage->dwKeyState&FCONTROL) ? (TRUE) : (FALSE);
	BOOL	fShift		= (pMessage->dwKeyState&FSHIFT) ? (TRUE) : (FALSE);
	BOOL	fEditable	= ElementOwner()->IsEditable(TRUE);
	short	zDelta;

	if(!fControl && !fShift)
	{
		// Do not scroll if vscroll is disallowed. This prevents content of
		// frames with scrolling=no does not get scrolled (IE5 #31515).
		CDispNodeInfo dni;
		GetDispNodeInfo(&dni);
		if(!dni.IsVScrollbarAllowed())
		{
			goto Cleanup;
		}

		// mousewheel scrolling, allow partial circle scrolling.
		zDelta = (short)HIWORD(pMessage->wParam);

		if(zDelta != 0)
		{
			long uScrollLines = WheelScrollLines();
			LONG yPercent = (uScrollLines>=0)
				? ((-zDelta*PERCENT_PER_LINE*uScrollLines)/WHEEL_DELTA)
				: ((-zDelta*PERCENT_PER_PAGE*abs(uScrollLines))/WHEEL_DELTA);

			if(ScrollByPercent(0, yPercent, MAX_SCROLLTIME))
			{
				hr = S_OK;
			}
		}
	}

Cleanup:
	RRETURN1(hr, S_FALSE);
}

// wlw note: should in textsite.cpp
WORD ConvVKey(WORD vKey)
{
    switch(vKey)
    {
    case VK_LEFT:
        return VK_DOWN;

    case VK_RIGHT:
        return VK_UP;

    case VK_UP:
        return VK_LEFT;

    case VK_DOWN:
        return VK_RIGHT;

    default:
        return vKey;
    }
}

HRESULT CFlowLayout::HandleSysKeyDown(CMessage* pMessage)
{
	HRESULT hr = S_FALSE;

	// BUGBUG: (anandra) Most of these should be handled as commands
	// not keydowns.  Use ::TranslateAccelerator to perform translation.
	if(_fVertical)
	{
		pMessage->wParam = ConvVKey((WORD)pMessage->wParam);
	}

	if(pMessage->wParam==VK_BACK && (pMessage->lParam&SYS_ALTERNATE))
	{
		Sound();
		hr = S_OK;
	}

	RRETURN1(hr, S_FALSE);
}

BOOL CFlowLayout::OnSetCursor(CMessage* pMessage)
{
	if(SameScope(ElementOwner(), pMessage->pNodeHit->GetUpdatedNearestLayoutNode())
		&& (HTC_YES==pMessage->htc))
	{
		return FALSE; // cursor will be set to text caret later
	}
	else
	{
		return super::OnSetCursor(pMessage);
	}
}

HRESULT CFlowLayout::HandleSetCursor(CMessage* pMessage, BOOL fIsOverEmptyRegion)
{
	HRESULT		hr = S_OK;
	LPCTSTR		idcNew = IDC_ARROW;
	RECT		rc;
	POINT		pt = pMessage->pt;
	BOOL		fEditable = ElementOwner()->IsEditable();
	CElement*	pElement = pMessage->pNodeHit->Element();
	CDocument*  pDoc = Doc();
	BOOL		fOverEditableElement = (fEditable && pElement && pDoc->IsElementSiteSelectable(pElement));

	// BUGBUG (MohanB) A hack to fix IE5 #60103; should be cleaned up in IE6.
	// MUSTFIX: We should set the default cursor (I-Beam for text, Arrow for the rest) only
	// after the message bubbles through all elements upto the root. This allows elements
	// (like anchor) which like to set non-default cursors over their content to do so.
	if(!fEditable && Tag()==ETAG_GENERIC)
	{
		return S_FALSE;
	}

	if(!fOverEditableElement)
	{
		GetClientRect(&rc);
	}

	Assert(pMessage->IsContentPointValid());
	if(fOverEditableElement || PtInRect(&rc, pMessage->ptContent))
	{
		if(fIsOverEmptyRegion)
		{
			idcNew = IDC_ARROW;
		}
		else
		{
			if(fEditable && (pMessage->htc>=HTC_TOPBORDER || pMessage->htc==HTC_EDGE))
			{
				idcNew = pDoc->GetCursorForHTC(pMessage->htc);
			}        
			else if(!pDoc->IsPointInSelection(pt))
			{
				// If CDoc is a HTML dialog, do not show IBeam cursor.
				if(fEditable || !( pDoc->_dwFlagsHostInfo&DOCHOSTUIFLAG_DIALOG)
					|| _fAllowSelectionInDialog)
				{
					// Adjust for Slave to make currency checkwork
					if(pElement && pElement->_etag==ETAG_TXTSLAVE)
					{
						pElement = pElement->MarkupMaster();
					}

					{
						idcNew = IDC_IBEAM;
					}
				}                    
			}  
			else if(pDoc->GetSelectionType() == SELECTION_TYPE_Control)
			{
				// We are in a selection. But the Adorners didn't set the HTC_CODE.
				// Set the caret to the size all - to indicate they can click down and drag.
				//
				// This is a little ambiguous - they can click in and UI activate to type
				// but they can also start a move we decided on the below.
				idcNew = IDC_SIZEALL;
			}
		}
	}

	ElementOwner()->SetCursorStyle(idcNew);

	RRETURN1(hr, S_FALSE);
}

BOOL CFlowLayout::IsElementBlockInContext(CElement* pElement)
{
	BOOL fRet = FALSE;

	if(pElement == ElementContent())
	{
		fRet = TRUE;
	}
	else if(!pElement->IsBlockElement() && !pElement->IsContainer() )
	{
		fRet = FALSE;
	}
	else if(!pElement->HasLayout())
	{
		fRet = TRUE;
	}
	else
	{
		BOOL fIsContainer = pElement->IsContainer();

		if(!fIsContainer)
		{
			fRet = TRUE;

			// God, I hate this hack ...
			if(pElement->Tag() == ETAG_FIELDSET)
			{
				CTreeNode* pNode = pElement->GetFirstBranch();

				if(pNode)
				{
					if(pNode->GetCascadeddisplay()!=styleDisplayBlock &&
						!pNode->GetCascadedwidth().IsNullOrEnum()) // IsWidthAuto
					{
						fRet = FALSE;
					}
				}
			}
		}
		else
		{
			// HACK ALERT!
			//
			// For display purposes, contianer elements in their parent context must
			// indicate themselves as block elements.  We do this only for container
			// elements who have been explicity marked as display block.
			if(fIsContainer)
			{
				CTreeNode* pNode = pElement->GetFirstBranch();

				if(pNode && pNode->GetCascadeddisplay() == styleDisplayBlock)
				{
					fRet = TRUE;
				}
			}
		}
	}

	return fRet;
}

HRESULT CFlowLayout::PreDrag(DWORD dwKeyState, IDataObject** ppDO, IDropSource** ppDS)
{
	HRESULT hr = S_OK;

	CDocument* pDoc = Doc();

	CSelDragDropSrcInfo* pDragInfo;

	// Setup some info for drag feedback
	Assert(!pDoc->_pDragDropSrcInfo);
	pDragInfo = new CSelDragDropSrcInfo(pDoc, ElementContent());

	if(!pDragInfo)
	{
		hr = E_OUTOFMEMORY;
		goto Cleanup;
	}
	hr = pDragInfo->GetDataObjectAndDropSource(ppDO, ppDS);
	if(hr)
	{
		goto Cleanup;
	}

	pDoc->_pDragDropSrcInfo = pDragInfo;

Cleanup:
	RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     PostDrag
//
//  Synopsis:   Handle the result of an OLE drag/drop operation
//
//  Arguments:  hrDrop      The hr that DoDragDrop came back with
//              dwEffect    The effect of the drag/drop
//
//-------------------------------------------------------------------------
HRESULT CFlowLayout::PostDrag(HRESULT hrDrop, DWORD dwEffect)
{
//#ifdef MERGEFUN // Edit team: figure out better way to send sel-change notifs
//	CCallMgr callmgr(GetPed());
//#endif
	HRESULT hr;

    /*CParentUndo pu(Doc()); wlw note*/

	hr = hrDrop;

	if(IsEditable())
	{
		/*pu.Start(IDS_UNDODRAGDROP); wlw note*/
	}

	if(hr == DRAGDROP_S_CANCEL)
	{
		hr = S_OK;
		goto Cleanup;
	}

	if(hr != DRAGDROP_S_DROP)
	{
		goto Cleanup;
	}

	hr = S_OK;

	switch(dwEffect)
	{
	case DROPEFFECT_NONE:
	case DROPEFFECT_COPY:
		Invalidate();
		break ;

	case DROPEFFECT_LINK:
		break;

	case 7:
		// dropEffect ALL - do the same thing as 3

	case 3: // BugBug - this is for TriEdit - faking out a position with Drag & Drop.
		{
			Assert(Doc()->_pDragDropSrcInfo);
			if(Doc()->_pDragDropSrcInfo)
			{
				CSelDragDropSrcInfo* pDragInfo;

				Assert(DRAGDROPSRCTYPE_SELECTION == Doc()->_pDragDropSrcInfo->_srcType);

				pDragInfo = DYNCAST(CSelDragDropSrcInfo, Doc()->_pDragDropSrcInfo);
				pDragInfo->PostDragSelect();
			}

		}
		break;

	case DROPEFFECT_MOVE:
		if(Doc()->_fSlowClick)
		{
			goto Cleanup;
		}

		Assert(Doc()->_pDragDropSrcInfo);
		if(Doc()->_pDragDropSrcInfo)
		{
			CSelDragDropSrcInfo* pDragInfo;

			Assert(DRAGDROPSRCTYPE_SELECTION == Doc()->_pDragDropSrcInfo->_srcType);


			pDragInfo = DYNCAST(CSelDragDropSrcInfo, Doc()->_pDragDropSrcInfo);
			pDragInfo->PostDragDelete();
		}
		break;

	default:
		Assert(FALSE && "Unrecognized drop effect");
		break;
	}

Cleanup:
	/*pu.Finish(hr); wlw note*/
	RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     Drop
//
//  Synopsis:
//
//----------------------------------------------------------------------------
#define DROPEFFECT_ALL (DROPEFFECT_NONE|DROPEFFECT_COPY|DROPEFFECT_MOVE|DROPEFFECT_LINK)

HRESULT CFlowLayout::Drop(
				  IDataObject*	pDataObj,
				  DWORD			grfKeyState,
				  POINTL		ptlScreen,
				  DWORD*		pdwEffect)
{
	CDocument*  pDoc			= Doc();
	DWORD	    dwAllowed		= *pdwEffect;
	TCHAR	    pszFileType[4]	= _T("");
	CPoint	    pt;
	HRESULT	    hr				= S_OK;

	// Can be null if dragenter was handled by script
	if(!_pDropTargetSelInfo)
	{
		*pdwEffect = DROPEFFECT_NONE ;
		return S_OK;
	}

	// Find out what the effect is and execute it
	// If our operation fails we return DROPEFFECT_NONE
	DragOver(grfKeyState, ptlScreen, pdwEffect);

	DropHelper(ptlScreen, dwAllowed, pdwEffect, pszFileType);


	if(Doc()->_fSlowClick && *pdwEffect==DROPEFFECT_MOVE)
	{
		*pdwEffect = DROPEFFECT_NONE ;
		goto Cleanup;
	}

	// We're all ok at this point. We delegate the handling of the actual
	// drop operation to the DropTarget.

	pt.x = ptlScreen.x;
	pt.y = ptlScreen.y;
	ScreenToClient(pDoc->_pInPlace->_hwnd, (POINT*)&pt);

	// We DON'T TRANSFORM THE POINT, AS MOVEPOINTERTOPOINT is in Global Coords
	hr = _pDropTargetSelInfo->Drop(this, pDataObj, grfKeyState, pt, pdwEffect);


Cleanup:
	// Erase any feedback that's showing.
	DragHide();

	Assert(_pDropTargetSelInfo);
	delete _pDropTargetSelInfo;
	_pDropTargetSelInfo = NULL;

	RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     DragLeave
//
//  Synopsis:
//
//----------------------------------------------------------------------------
HRESULT CFlowLayout::DragLeave()
{
//#ifdef MERGEFUN // Edit team: figure out better way to send sel-change notifs
//	CCallMgr callmgr(GetPed());
//#endif
	HRESULT hr = S_OK;

	if(!_pDropTargetSelInfo)
	{
		goto Cleanup;
	}

	hr = super::DragLeave();

Cleanup:
	if(_pDropTargetSelInfo)
	{
		delete _pDropTargetSelInfo;
		_pDropTargetSelInfo = NULL;
	}
	RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     ParseDragData
//
//  Synopsis:   Drag/drop helper override
//
//----------------------------------------------------------------------------
HRESULT CFlowLayout::ParseDragData(IDataObject* pDO)
{
	DWORD dwFlags = 0;
	HRESULT hr;
	CTreeNode* pNode = NULL;

	// Start with flags set to default values.
	Doc()->_fOKEmbed = FALSE;
	Doc()->_fOKLink = FALSE;
	Doc()->_fFromCtrlPalette = FALSE;

	if(!ElementOwner()->IsEditable() || !ElementOwner()->IsEnabled())
	{
		// no need to do anything else, bcos we're read only.
		hr = S_FALSE;
		goto Cleanup;
	}

	pNode = ElementOwner()->GetFirstBranch();
	if(!pNode)
	{
		hr = E_FAIL;
		goto Cleanup;
	}

	// Allow only plain text to be pasted in input text controls
	if(!pNode->SupportsHtml() && pDO->QueryGetData(&g_rgFETC[iAnsiFETC])!=NOERROR)
	{
		hr = S_FALSE;
		goto Cleanup;
	}

	{
		hr = CTextXBag::GetDataObjectInfo(pDO, &dwFlags);
		if(hr)
		{
			goto Cleanup;
		}
	}

	if(dwFlags & DOI_CANPASTEPLAIN)
	{
		hr = S_OK;
	}
	else
	{
		hr = super::ParseDragData(pDO);
	}

Cleanup:
	RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     DrawDragFeedback
//
//  Synopsis:
//
//----------------------------------------------------------------------------
void CFlowLayout::DrawDragFeedback()
{
	Assert(_pDropTargetSelInfo);

	_pDropTargetSelInfo->DrawDragFeedback();
}

//+------------------------------------------------------------------------
//
//  Member:     InitDragInfo
//
//  Synopsis:   Setup a struct to enable drawing of the drag feedback
//
//  Arguments:  pDO         The data object
//              ptlScreen   Screen loc of obj.
//
//  Notes:      This assumes that the DO has been parsed and
//              any appropriate data on the form has been set.
//
//-------------------------------------------------------------------------
HRESULT CFlowLayout::InitDragInfo(IDataObject* pDO, POINTL ptlScreen)
{
	CPoint pt;
	pt.x = ptlScreen.x;
	pt.y = ptlScreen.y;

	ScreenToClient(Doc()->_pInPlace->_hwnd, (CPoint*)&pt);

	// We DON'T TRANSFORM THE POINT, AS MOVEPOINTERTOPOINT is in Global Coords
	Assert(!_pDropTargetSelInfo);
	_pDropTargetSelInfo = new CDropTargetInfo(this, Doc(), pt);
	if(!_pDropTargetSelInfo)
	{
		RRETURN(E_OUTOFMEMORY);
	}

	return S_OK;

}

//+====================================================================================
//
// Method: DragOver
//
// Synopsis: Delegate to Layout::DragOver - unless we don't have a _pDropTargetSelInfo
//
//------------------------------------------------------------------------------------
HRESULT CFlowLayout::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
	// Can be null if dragenter was handled by script
	if(!_pDropTargetSelInfo)
	{
		*pdwEffect = DROPEFFECT_NONE ;
		return S_OK;
	}

	return super::DragOver(grfKeyState, pt, pdwEffect);
}

//+---------------------------------------------------------------------------
//
//  Member:     UpdateDragFeedback
//
//  Synopsis:
//
//----------------------------------------------------------------------------
HRESULT CFlowLayout::UpdateDragFeedback(POINTL ptlScreen)
{
	CSelDragDropSrcInfo* pDragInfo = NULL;
	CDocument* pDoc = Doc();
	CPoint pt;

	// Can be null if dragenter was handled by script
	if(!_pDropTargetSelInfo)
	{
		return S_OK;
	}

	pt.x = ptlScreen.x;
	pt.y = ptlScreen.y;

	ScreenToClient(pDoc->_pInPlace->_hwnd, (POINT*)&pt);

	// We DON'T TRANSFORM THE POINT, AS MOVEPOINTERTOPOINT is in Global Coords
	if((pDoc->_fIsDragDropSrc ) && (pDoc->_pDragDropSrcInfo) &&
		(pDoc->_pDragDropSrcInfo->_srcType==DRAGDROPSRCTYPE_SELECTION))
	{
		pDragInfo = DYNCAST(CSelDragDropSrcInfo, pDoc->_pDragDropSrcInfo);
	}
	_pDropTargetSelInfo->UpdateDragFeedback(this, pt, pDragInfo );

	Trace((_T("Update Drag Feedback: pt:%ld,%ld After Transform:%ld,%ld\n"),
		ptlScreen.x, ptlScreen.y, pt.x, pt.y));

	return S_OK;
}

CFlowLayout* CFlowLayout::GetNextFlowLayout(
	NAVIGATE_DIRECTION	iDir,
	POINT				ptPosition,
	CElement*			pElementLayout,
	LONG*				pcp,
	BOOL*				pfCaretNotAtBOL,
	BOOL*				pfAtLogicalBOL)
{
	CFlowLayout* pFlowLayout = NULL; // Stores the new txtsite found in the given dirn.

	Assert(pcp);
	Assert(!pElementLayout || pElementLayout->GetUpdatedParentLayout()==this);

	if(pElementLayout == NULL)
	{
		CLayout* pParentLayout = GetUpdatedParentLayout();
		// By default ask our parent to get the next flowlayout.
		if(pParentLayout)
		{
			pFlowLayout = pParentLayout->GetNextFlowLayout(iDir, ptPosition, ElementOwner(), pcp, pfCaretNotAtBOL, pfAtLogicalBOL);
		}
	}
	else
	{
		CTreePos* ptpStart; // extent of the element
		CTreePos* ptpFinish;

		// Start off with the txtsite found as being this one
		pFlowLayout = this;
		*pcp = 0;

		// find the extent of the element passed in.
		pElementLayout->GetTreeExtent(&ptpStart, &ptpFinish);

		switch(iDir)
		{
		case NAVIGATE_UP:
		case NAVIGATE_DOWN:
			{
				CLinePtr rp(&_dp); // The line in this site where the child lives
				POINT    pt; // The point where the site resides.

				// find the line where the given layout lives.
				if(_dp.RenderedPointFromTp(ptpStart->GetCp(), ptpStart, FALSE, pt, &rp, TA_TOP, NULL) < 0)
				{
					goto Cleanup;
				}

				// Now navigate from this line ... either up/down
				pFlowLayout = _dp.MoveLineUpOrDown(iDir, rp, ptPosition.x, pcp, pfCaretNotAtBOL, pfAtLogicalBOL);
				break;
			}
		case NAVIGATE_LEFT:
			// position ptpStart just before the child layout
			ptpStart = ptpStart->PreviousTreePos();

			if(ptpStart)
			{
				// Now let's get the txt site that's interesting to us
				pFlowLayout = ptpStart->GetBranch()->GetFlowLayout();

				// and the cp...
				*pcp = ptpStart->GetCp();
			}
			break;

		case NAVIGATE_RIGHT:
			// Position the ptpFinish just after the child layout.
			ptpFinish = ptpFinish->PreviousTreePos();

			if(ptpFinish)
			{
				// Now let's get the txt site that's interesting to us
				pFlowLayout = ptpFinish->GetBranch()->GetFlowLayout();

				// and the cp...
				*pcp = ptpFinish->GetCp();
				break;
			}
		}
	}

Cleanup:
	return pFlowLayout;
}

CFlowLayout* CFlowLayout::GetPositionInFlowLayout(
	NAVIGATE_DIRECTION	iDir,
	POINT				ptPosition,
	LONG*				pcp,
	BOOL*				pfCaretNotAtBOL,
	BOOL*				pfAtLogicalBOL)
{
	CFlowLayout* pFlowLayout = this; // The txtsite we found ... by default us

	Assert(pcp);

	switch(iDir)
	{
	case NAVIGATE_UP:
	case NAVIGATE_DOWN:
		{
			CPoint ptGlobal(ptPosition); // The desired position of the caret
			CPoint ptContent;
			CLinePtr rp(&_dp); // The line in which the point ptPosition is
			CRect rcClient; // Rect used to get the client rect

			// Be sure that the point is within this site's client rect
			RestrictPointToClientRect(&ptGlobal);
			ptContent = ptGlobal;
			TransformPoint(&ptContent, COORDSYS_GLOBAL, COORDSYS_CONTENT);

			// Construct a point within this site's client rect (based on
			// the direction we are traversing.
			GetClientRect(&rcClient);
			rcClient.MoveTo(ptContent);

			// Find the line within this txt site where we want to be placed.
			rp = _dp.LineFromPos(rcClient,
				(CDisplay::LFP_ZORDERSEARCH|
				CDisplay::LFP_IGNOREALIGNED|
				CDisplay::LFP_IGNORERELATIVE|
				(iDir==NAVIGATE_UP?CDisplay::LFP_INTERSECTBOTTOM:0)));

			if(rp < 0)
			{
				*pcp = 0;
			}
			else
			{
				// Found the line ... let's navigate to it.
				pFlowLayout = _dp.NavigateToLine(iDir, rp, ptGlobal, pcp, pfCaretNotAtBOL, pfAtLogicalBOL);
			}
			break;
		}

	case NAVIGATE_LEFT:
		{
			// We have come to this site while going left in a site outside this site.
			// So position ourselves just after the last valid character.
			*pcp = GetContentLastCp() - 1;
#ifdef _DEBUG
			{
				//CRchTxtPtr rtp(GetPed());
				//rtp.SetCp(*pcp);
				//Assert(WCH_TXTSITEEND == rtp._rpTX.GetChar());
			}
#endif
			break;
		}

	case NAVIGATE_RIGHT:
		// We have come to this site while going right in a site outside this site.
		// So position ourselves just before the first character.
		*pcp = GetContentFirstCp();
#ifdef _DEBUG
		{
			//CRchTxtPtr rtp(GetPed());
			//rtp.SetCp(*pcp);
			//Assert(IsTxtSiteBreak(rtp._rpTX.GetPrevChar()));
		}
#endif
		break;
	}

	return pFlowLayout;
}

LONG CFlowLayout::XCaretToRelative(LONG xCaret)
{
	return xCaret;
}

//+----------------------------------------------------------------------------
//
//  Member:     XCaretToAbsolute
//
//  Synopsis:   Converts a xCaret value from site relative coordinates to
//              global client window coordinates
//
//  Arguments:  [xCaretReally]: The xCaret position in site relative coords
//
//-----------------------------------------------------------------------------
LONG CFlowLayout::XCaretToAbsolute(LONG xCaretReally)
{
	return xCaretReally;
}

//+-------------------------------------------------------------------------
//
//  Method:     YieldCurrencyHelper
//
//  Synopsis:   Relinquish currency
//
//  Arguments:  pElemNew    New site that wants currency
//
//--------------------------------------------------------------------------
HRESULT CFlowLayout::YieldCurrencyHelper(CElement* pElemNew)
{
    HRESULT hr = S_OK;

    Assert(pElemNew != ElementOwner());

#ifndef NO_IME
    // Restore the IMC if we've temporarily disabled it.  See OnSetFocus().
    if(Doc()->_himcCache)
    {
        ImmAssociateContext(Doc()->_pInPlace->_hwnd, Doc()->_himcCache);
        Doc()->_himcCache = NULL;
    }
#endif

    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     BecomeCurrentHelper
//
//  Synopsis:   Force currency on the site.
//
//  Notes:      This is the method that external objects should call
//              to force sites to become current.
//
//--------------------------------------------------------------------------
HRESULT CFlowLayout::BecomeCurrentHelper(long lSubDivision, BOOL* pfYieldFailed, CMessage* pMessage)
{
	// CHROME
	// If we are chrome hosted then we have no HWND.
	// Therefore simply ask our container whether we have focus.
	BOOL fOnSetFocus = (::GetFocus()==Doc()->_pInPlace->_hwnd);
	HRESULT hr = S_OK;

	// BUG44836 -- moved here from OnSetFocus -- OnSetFocus() was calling
	// this fn more aggressively than we want.
	NotifyKillSelection();

	// Call OnSetFocus directly if the doc's window already had focus. We
	// don't need to do this if the window didn't have focus because when
	// we take focus in BecomeCurrent, the window message handler does this.
	if(fOnSetFocus)
	{
		// if our inplace window did have the focus, then fire the onfocus
		// only if onblur was previously fired and the body is becoming the
		// current site and currency did not change in onchange or onpropchange
		if(ElementOwner()==Doc()->_pElemCurrent && ElementOwner()==Doc()->GetPrimaryElementClient())
		{
			//Doc()->Fire_onfocus();
		}
	}

	RRETURN1(hr, S_FALSE);
}

LONG CFlowLayout::GetNestedElementCch(
									  CElement*		pElement,		// IN:  The nested element
									  CTreePos**	pptpLast,		// OUT: The pos beyond pElement 
									  LONG			cpLayoutLast,	// IN:  This layout's last cp
									  CTreePos*		ptpLayoutLast)	// IN:  This layout's last pos
{
	CTreePos*	ptpStart;
	CTreePos*	ptpLast;
	long		cpElemStart;
	long		cpElemLast;

	if(cpLayoutLast == -1)
	{
		cpLayoutLast = GetContentLastCp();
	}
	Assert(cpLayoutLast == GetContentLastCp());
	if(ptpLayoutLast == NULL)
	{
		ElementContent()->GetTreeExtent(NULL, &ptpLayoutLast);
	}
#ifdef _DEBUG
	{
		CTreePos* ptpLastDbg;
		ElementContent()->GetTreeExtent(NULL, &ptpLastDbg);
		Assert(ptpLayoutLast == ptpLastDbg);
	}
#endif

	pElement->GetTreeExtent(&ptpStart, &ptpLast);

	cpElemStart = ptpStart->GetCp();
	cpElemLast = ptpLast->GetCp();

	if(cpElemLast > cpLayoutLast)
	{
		if(pptpLast)
		{
			ptpLast = ptpLayoutLast->PreviousTreePos();
			while(!ptpLast->IsNode())
			{
				Assert(ptpLast->GetCch() == 0);
				ptpLast = ptpLast->PreviousTreePos();
			}
		}

		// for overlapping layout limit the range to
		// parent layout's scope.
		cpElemLast = cpLayoutLast - 1;
	}

	if(pptpLast)
	{
		*pptpLast = ptpLast;
	}

	return (cpElemLast-cpElemStart+1);
}

HRESULT CFlowLayout::Init()
{
	HRESULT hr;

	hr = super::Init();
	if(hr)
	{
		goto Cleanup;
	}

	if(!_dp.Init())
	{
		hr = E_FAIL;
		goto Cleanup;
	}

	_dp._dxCaret = ElementOwner()->IsEditable(TRUE) ? 1 : 0;

Cleanup:
	RRETURN(hr);
}


//+---------------------------------------------------------------------------
//
//  Member:     CFlowLayout::OnExitTree
//
//  Synopsis:   Deinitilialize the display on exit tree
//
//----------------------------------------------------------------------------
HRESULT CFlowLayout::OnExitTree()
{
	HRESULT hr;

	hr = super::OnExitTree();
	if(hr)
	{
		goto Cleanup;
	}

	_dp.FlushRecalc();

Cleanup:
	RRETURN(hr);
}

BOOL CFlowLayout::HitTestContent(const POINT* pptHit, CDispNode* pDispNode, void* pClientData)
{
	Assert(pptHit);
	Assert(pDispNode);
	Assert(pClientData);

	CHitTestInfo*	phti            = (CHitTestInfo*)pClientData;
	CDispNode*		pDispElement    = GetElementDispNode();
	BOOL			fHitTestContent = !pDispElement->IsContainer() || pDispElement!=pDispNode;

	// Skip nested markups if asked to do so
	if(phti->_grfFlags & HT_SKIPSITES)
	{
		CElement* pElement = ElementContent();
		if(pElement && pElement->HasSlaveMarkupPtr()
			&& !SameScope(pElement, phti->_pNodeElement))
		{
			phti->_htc = HTC_NO;
			goto Cleanup;
		}
	}

	if(fHitTestContent)
	{
		// If allowed, see if a child element is hit
		// NOTE: Only check content when the hit display node is a content node
		phti->_htc = BranchFromPoint(
			phti->_grfFlags,
			*pptHit,
			&phti->_pNodeElement,
			phti->_phtr,
			TRUE, // ignore pseudo hit's
			pDispNode);

		// BUGBUG (donmarsh) - BranchFromPoint was written before the Display Tree
		// was introduced, so it might return a child element that already rejected
		// the hit when it was called directly from the Display Tree.  Therefore,
		// if BranchFromPoint returned an element that has its own layout, and if
		// that layout has its own display node, we reject the hit.
		if(phti->_htc==HTC_YES
			&& phti->_pNodeElement!=NULL
			&& phti->_pNodeElement->Element()!=ElementOwner()
			&& phti->_pNodeElement->NeedsLayout())
		{
			phti->_htc = HTC_NO;
			phti->_pNodeElement = NULL;
		}
		// If we pseudo-hit a flow-layer dispnode that isn't the "bottom-most"
		// (i.e. "first") flow-layer dispnode for this element, then we pretend
		// we really didn't hit it.  This allows hit testing to "drill through"
		// multiple flow-layer dispnodes in order to support hitting through
		// non-text areas of display nodes generated by -ve margins.
		// Bug #
		else if(phti->_htc==HTC_YES 
			&& phti->_phtr->_fPseudoHit==TRUE
			&& pDispNode->GetLayerType()==DISPNODELAYER_FLOW
			&& pDispNode->GetPreviousSiblingNode(TRUE))
		{
			phti->_htc = HTC_NO;
		}

		if(phti->_htc == HTC_YES)
		{
			goto Finalize;
		}
	}

	// Do not call super if we are hit testing content and the current
	// element is a container. DisplayTree calls back with a HitTest
	// for the background after hittesting the -Z content.
	if(!fHitTestContent || !pDispElement->IsContainer())
	{
		// If no child and no peer was hit, use default handling
		super::HitTestContent(pptHit, pDispNode, pClientData);
	}
	goto Cleanup; // done

Finalize:
	// Save the point and CDispNode associated with the hit
	phti->_ptContent = *pptHit;
	phti->_pDispNode = pDispNode;

Cleanup:
	return (phti->_htc!=HTC_NO);
}


//+---------------------------------------------------------------------------
//
//  Member:     CFlowLayout::NotifyScrollEvent
//
//  Synopsis:   Respond to a change in the scroll position of the display node
//
//----------------------------------------------------------------------------
void CFlowLayout::NotifyScrollEvent(RECT* prcScroll, SIZE* psizeScrollDelta)
{
	// BUGBUG: Add 1st visible cp tracking here....(brendand)?
	super::NotifyScrollEvent(prcScroll, psizeScrollDelta);
}



const CLayout::LAYOUTDESC C1DLayout::s_layoutdesc =
{
	LAYOUTDESC_FLOWLAYOUT, // _dwFlags
};

//+---------------------------------------------------------------------------
//
//  Member:     C1DLayout::Init()
//
//  Synopsis:   Called when the element enters the tree. Super initializes
//              CDisplay.
//
//----------------------------------------------------------------------------
HRESULT C1DLayout::Init()
{
	HRESULT hr = super::Init();
	if(hr)
	{
		goto Cleanup;
	}

Cleanup:
	RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     C1DLayout::CalcSize
//
//----------------------------------------------------------------------------
DWORD C1DLayout::CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault)
{
	CSaveCalcInfo	sci(pci, this);
	CScopeFlag		csfCalcing(this);
	CSize			sizeOriginal;
	BOOL			fNormalMode = (pci->_smMode==SIZEMODE_NATURAL ||
		pci->_smMode==SIZEMODE_SET ||
		pci->_smMode==SIZEMODE_FULLSIZE);
	BOOL			fSizeChanged;
	DWORD			grfReturn;

	CDisplay* pdp = GetDisplay();
	Assert(pdp);

	Listen();

	Assert(pci);
	Assert(psize);

	GetSize(&sizeOriginal);

	if(_fForceLayout)
	{
		pci->_grfLayout |= LAYOUT_FORCE;
		_fForceLayout = FALSE;
	}

	// Set default return values and initial state
	grfReturn = (pci->_grfLayout & LAYOUT_FORCE);

	if(pci->_grfLayout & LAYOUT_FORCE)
	{
		_fSizeThis         = TRUE;
		_fAutoBelow        = FALSE;
		_fPositionSet      = FALSE;
		_fContainsRelative = FALSE;
	}

	fSizeChanged = (fNormalMode
		? ((psize->cx!=sizeOriginal.cx)||(psize->cy!=sizeOriginal.cy)) : FALSE);

	// Text recalc is required if:
	//  a) The site is marked for sizing
	//  b) The available width has changed (during normal calculations)
	//  c) The calculation is forced
	//  d) Changes against the text are outstanding
	//  e) The requested size is not already cached
	//
	if((fNormalMode&&_fSizeThis)
		|| (fNormalMode&&fSizeChanged)
		|| (pci->_grfLayout&LAYOUT_FORCE)
		|| !IsCommitted()
		|| (pci->_smMode==SIZEMODE_MMWIDTH&&!_fMinMaxValid)
		|| (pci->_smMode==SIZEMODE_MINWIDTH &&(!_fMinMaxValid||_xMin<0)))
	{
		// This is a helper routine specific to CFlowSite
		grfReturn = RecalcText(pci, psize, psizeDefault);
	}

	// Propagate SIZEMODE_NATURAL requests to contained sites
	if(pci->_smMode==SIZEMODE_NATURAL && HasRequestQueue())
	{
		long xParentWidth;
		long yParentHeight;

		pdp->GetViewWidthAndHeightForChild(
			pci,
			&xParentWidth,
			&yParentHeight,
			pci->_smMode==SIZEMODE_MMWIDTH);

		ProcessRequests(pci, CSize(xParentWidth, yParentHeight));
	}

	// Lastly, return the requested size
	switch(pci->_smMode)
	{
	case SIZEMODE_SET:
	case SIZEMODE_NATURAL:
	case SIZEMODE_FULLSIZE:
		Assert(!_fSizeThis);
		GetSize(psize);
		Reset(FALSE);
		Assert(!HasRequestQueue() || GetView()->HasLayoutTask(this));
		break;

	case SIZEMODE_MMWIDTH:
		Assert(_fMinMaxValid);
		psize->cx = _xMax;
		psize->cy = _xMin;
		break;

	case SIZEMODE_MINWIDTH:
		psize->cx = _xMin;
		psize->cy = 0;
		break;
	}

	return grfReturn;
}

//+---------------------------------------------------------------------------
//
//  Member:     C1DLayout::RecalcText
//
//  Synopsis:   Helper function for CalcSize
//
//----------------------------------------------------------------------------
DWORD C1DLayout::RecalcText(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault)
{
	CSaveCalcInfo	sci(pci, this);
	BOOL			fNormalMode = (pci->_smMode==SIZEMODE_NATURAL
		|| pci->_smMode==SIZEMODE_SET
		|| pci->_smMode==SIZEMODE_FULLSIZE);
	BOOL			fHeightChanged;
	CTreeNode*		pContext = GetFirstBranch();
	const CFancyFormat* pFF = pContext->GetFancyFormat();
	CUnitValue		uvWidth = pFF->_cuvWidth;
	CUnitValue		uvHeight = pFF->_cuvHeight;
	CUnitValue		uvBottom = pFF->_cuvBottom;
	BOOL			fPercentWidth = pFF->_fWidthPercent;
	BOOL			fPercentHeight = pFF->_fHeightPercent;
	styleOverflow	bOverflowX = (styleOverflow)pFF->_bOverflowX;
	styleOverflow	bOverflowY = (styleOverflow)pFF->_bOverflowY;
	styleStyleFloat	bFloat = (styleStyleFloat)pFF->_bStyleFloat;
	BOOL			fClipContentX = ((bOverflowX!=styleOverflowVisible) &&
		(bOverflowX!=styleOverflowNotSet));
	BOOL			fClipContentY = ((bOverflowY!=styleOverflowVisible) &&
		(bOverflowY!=styleOverflowNotSet));
	long			cxUser, cyUser, cxInit = 0;
	long			xLeft, xRight;
	long			yTop, yBottom;
	CSize			sizeOriginal;
	DWORD			grfReturn;
	long			lPadding[PADDING_MAX];

	Assert(GetDisplay());

	grfReturn = (pci->_grfLayout & LAYOUT_FORCE);

	GetSize(&sizeOriginal);

	// If dirty, ensure display tree nodes exist
	if(_fSizeThis && (EnsureDispNode(pci, (grfReturn&LAYOUT_FORCE))==S_FALSE))
	{
		grfReturn |= LAYOUT_HRESIZE|LAYOUT_VRESIZE;
	}

	// We need to get padding into the mixture for calculating
	// percentWidth and percentHeight to get the correct numbers
	long lcxPadding = 0, lcyPadding = 0;
	if(ElementOwner()->IsAbsolute() && (fPercentWidth || fPercentHeight))
	{
		CFlowLayout* pLayoutParent = pContext->Parent()->GetFlowLayout();
		CDisplay* pdpParent = pLayoutParent->GetDisplay();

		pdpParent->GetPadding(pci, lPadding, pci->_smMode==SIZEMODE_MMWIDTH);

		if(fPercentWidth)
		{
			lcxPadding = lPadding[PADDING_LEFT] + lPadding[PADDING_RIGHT];
		}

		if(fPercentHeight)
		{
			lcyPadding = lPadding[PADDING_TOP] + lPadding[PADDING_BOTTOM];
		}
	}

	// aligned & absolute layouts are sized based on parent's available width,
	// and inline elements are sized based on the current line wrapping width.
	// psize->cx holds the wrapping width.
	cxUser = (fPercentWidth&&!fNormalMode)
		? 0
		: uvWidth.XGetPixelValue(
		pci,
		pFF->_fAlignedLayout ? pci->_sizeParent.cx : psize->cx,
		GetFirstBranch()->GetFontHeightInTwips(&uvWidth));
	cyUser = (fPercentHeight && !fNormalMode)
		? 0
		: uvHeight.YGetPixelValue(
		pci,
		pci->_sizeParent.cy + lcyPadding,
		GetFirstBranch()->GetFontHeightInTwips(&uvHeight));

	xLeft = 0;
	xRight = psize->cx;

	yTop = 0;
	yBottom = psize->cy;

	// BUGBUG (a-pauln) Does this flag need to be set further down and with different cause?
	//                  What about fClipContent? Rename to _fContentsAffectHeight?

	// set the width and height for our text layout
	if(ElementOwner()->IsAbsolute() && (uvWidth.IsNullOrEnum() || uvHeight.IsNullOrEnum()))
	{
		CalcAbsoluteSize(pci, psize, &cxUser, &cyUser, &xLeft, &xRight, &yTop, &yBottom, fClipContentX, fClipContentY);
	}
	else
	{
		_fContentsAffectSize =  (uvHeight.IsNullOrEnum()
			|| uvWidth.IsNullOrEnum()
			|| !fClipContentX
			|| !fClipContentY);
	}

	// if contents affect size, start with 0 height and expand.
	cxInit =
		psize->cx = uvWidth.IsNullOrEnum()
		? max(0L, xRight-xLeft)
		: max(0L, cxUser);
	psize->cy = uvHeight.IsNullOrEnum()
		? 0
		: max(0L, cyUser);

	// Decide if our size changed
	// We used to have a flag for both width and height but
	// it is too fragile to predict if the width has changed
	// or not.  It is also too hard to know whether our
	// height has changed or not when _fContentsAffectSize.
	fHeightChanged = FALSE;

	if(fNormalMode)
	{
		fHeightChanged = _fContentsAffectSize ? TRUE : (cyUser!=sizeOriginal.cy);
	}

	// Mark percentage sized children as in need of sizing
	// CONSIDER: as an optimization, if we knew for sure the height
	//   did indeed change, we could not do this here and be sure
	//   that the full recalc would take care of all children.
	if(fNormalMode && !ContainsVertPercentAttr() && fHeightChanged)
	{
		long fContentsAffectSize = _fContentsAffectSize;

		_fContentsAffectSize = FALSE;
		ResizePercentHeightSites();
		_fContentsAffectSize = fContentsAffectSize;
	}

	// Cache sizes and recalculate the text (if required)
	// ** with lock **
	{
		CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);

		//
		// Auto-size to content if we are width auto and either float
		// left/right or absolutely positioned.
		if(uvWidth.IsNullOrEnum()
			&& (bFloat==styleStyleFloatLeft
			|| bFloat==styleStyleFloatRight
			|| !ElementOwner()->IsBlockElement()
			|| ElementOwner()->IsAbsolute()))
		{
			_fSizeToContent = TRUE;
		}

		// Calculate the text
		CalcTextSize(pci, psize);

		// For normal modes, cache values and request layout
		if(fNormalMode)
		{
			// For each direction,
			// If a explicit size exists and clipping, use the explicit size
			// Otherwise, adjust size to content
			if(fClipContentX)
			{
				psize->cx = uvWidth.IsNullOrEnum() ? psize->cx : cxUser;
			}
			else
			{
				if(_fSizeToContent)
				{
					CRect rc(*psize);
					CRect rcBorders;

					rc.right = psize->cx = _dp.GetWidth();

					_dp.SetViewSize(rc);

					SizeContentDispNode(*psize);

					_dp.RecalcLineShift(pci, 0);

					rcBorders.top    = 0;
					rcBorders.left   = 0;
					rcBorders.right  = 0x7FFFFFFF;
					rcBorders.bottom = 0x7FFFFFFF;

					GetClientRect(&rcBorders, CLIENTRECT_USERECT, pci);

					psize->cx += 0x7FFFFFFF - (rcBorders.right - rcBorders.left);
				}
				else if(!uvWidth.IsNullOrEnum() && psize->cx<cxUser)
				{
					psize->cx = max(0L, cxUser);
				}

			}

			if(fClipContentY)
			{
				psize->cy = uvHeight.IsNullOrEnum() ? psize->cy : cyUser;
			}
			else
			{
				if(!uvHeight.IsNullOrEnum() && (psize->cy < cyUser))
				{
					psize->cy = max(0L, cyUser);
				}
			}

		}

		_fSizeToContent = FALSE;
	}

	// Resize the child height percent sites again.  Do another
	// CalcTextSize to make sure they are actually sized.
	// ** No Lock **
	if(fNormalMode && _fContentsAffectSize)
	{
		if(ContainsVertPercentAttr() || _fChildHeightPercent)
		{
			SIZE    size;

			size.cx = cxInit;
			size.cy = psize->cy;
			{
				long fContentsAffectSize = _fContentsAffectSize;

				_fContentsAffectSize = FALSE;
				ResizePercentHeightSites();
				_fContentsAffectSize = fContentsAffectSize;
			}

			CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);

			CalcTextSize(pci, &size);

			psize->cy = max(psize->cy, size.cy);

			if (fClipContentY)
			{
				psize->cy = uvHeight.IsNullOrEnum() ? psize->cy : cyUser;
			}

			// Intent is to insure that both calculations return the same width.
			// This is (validly) not the case when we are absolutely positioned and have
			//  children dependent on our height and we are not clipping  and the normal
			//  width of content is less than the width of the percent sized child.
			//
			Assert((psize->cx>=size.cx)
				|| (ElementOwner()->IsAbsolute() 
				&& _fChildHeightPercent
				&& fNormalMode
				&& !fClipContentX));
		}
	}

	// Finish up the recalc ** with lock **
	{
		CElement::CLock Lock(ElementOwner(), CElement::ELEMENTLOCK_SIZING);

		if(fNormalMode)
		{
			grfReturn |= LAYOUT_THIS
				| (psize->cx != sizeOriginal.cx
				? LAYOUT_HRESIZE
				: 0)
				| (psize->cy != sizeOriginal.cy
				? LAYOUT_VRESIZE
				: 0);

			// If size changes occurred, size the display nodes
			if(grfReturn & (LAYOUT_FORCE|LAYOUT_HRESIZE|LAYOUT_VRESIZE))
			{
				SizeDispNode(pci, *psize);
                SizeContentDispNode(CSize(_dp.GetMaxWidth(), _dp.GetHeight()));
			}

			// Mark the site clean
			// (This is done last so that other code, such as GetClientRect,
			// will "do the right thing" while the size is being determined)
			//
			_fSizeThis = FALSE;
		}

		// For min/max mode, cache the values and note that they are now valid
		else if(pci->_smMode == SIZEMODE_MMWIDTH)
		{
			// If the user doesn't specify a width, or the width
			// is percentage size, we return min/max info
			// as calculated for any other regular CTxtSite (by
			// CalcTextSize)
			if(!uvWidth.IsNullOrEnum() && !fPercentWidth)
			{
				// We have two cases here:
				// 1) we are clipping - just use the size specified
				//      for both min/max
				// 2) we are growing but know our width so use
				//      max(psize->cy, cxUser) for both min/max
				//
                psize->cy = psize->cx =
                    fClipContentX ? max(0L, cxUser) : max(psize->cy, cxUser);
            }
            _xMax = psize->cx;
			_xMin = psize->cy;

			_fMinMaxValid = TRUE;
		}
		else if(pci->_smMode == SIZEMODE_MINWIDTH)
		{
			if(!uvWidth.IsNullOrEnum() && !fPercentWidth)
			{
				psize->cx = fClipContentX
					? max(0L, cxUser)
					: max(psize->cx, cxUser);
			}
			_xMin = psize->cx;
		}
	}
	return grfReturn;
}

//+---------------------------------------------------------------------------
//
//  Member:     C1DLayout::CalcAbsoluteSize
//
//  Synopsis:   Computes width and height of absolutely positioned element
//
//----------------------------------------------------------------------------
void C1DLayout::CalcAbsoluteSize(CCalcInfo * pci,
							SIZE* psize,
							long* pcxWidth,
							long* pcyHeight,
							long* pxLeft,
							long* pxRight,
							long* pyTop,
							long* pyBottom,
							BOOL  fClipContentX,
							BOOL  fClipContentY)
{
	CTreeNode*   pContext      = GetFirstBranch();
	styleDir     bDirection    = pContext->GetCascadedBlockDirection();
	CUnitValue   uvWidth       = pContext->GetCascadedwidth();
	CUnitValue   uvHeight      = pContext->GetCascadedheight();
	CUnitValue   uvLeft        = pContext->GetCascadedleft();
	CUnitValue   uvRight       = pContext->GetCascadedright();
	CUnitValue   uvTop         = pContext->GetCascadedtop();
	CUnitValue   uvBottom      = pContext->GetCascadedbottom();
	BOOL         fWidthAuto    = uvWidth.IsNullOrEnum();
	BOOL         fHeightAuto   = uvHeight.IsNullOrEnum();
	BOOL         fLeftAuto     = uvLeft.IsNullOrEnum();
	BOOL         fRightAuto    = uvRight.IsNullOrEnum();
	BOOL         fTopAuto      = uvTop.IsNullOrEnum();
	BOOL         fBottomAuto   = uvBottom.IsNullOrEnum();
	CFlowLayout* pLayoutParent = pContext->Parent()->GetFlowLayout();
	CDisplay*    pdpParent     = pLayoutParent->GetDisplay();
	long         lPadding[PADDING_MAX];
	POINT        pt;

	pt = _afxGlobalData._Zero.pt;

	_fContentsAffectSize =  (fHeightAuto
		&& fBottomAuto)
		|| (fWidthAuto
		&& ((bDirection==htmlDirLeftToRight
		&& fRightAuto)
		|| (bDirection==htmlDirRightToLeft
		&& fLeftAuto)))
		|| !fClipContentX
		|| !fClipContentY;

	// BUGBUG: The pci should pass the yOffset as well! (brendand)
	if ((fHeightAuto
		&& fBottomAuto)
		|| (fLeftAuto
		&& bDirection==htmlDirLeftToRight)
		|| (fRightAuto
		&& bDirection==htmlDirRightToLeft))
	{
		CLinePtr rp(pdpParent);
		CTreePos* ptpStart;
		CTreePos* ptpEnd;

		ElementOwner()->GetTreeExtent(&ptpStart, &ptpEnd);

		if(pdpParent->PointFromTp(ptpStart->GetCp(), ptpStart, FALSE, FALSE, pt, &rp, TA_TOP) < 0)
		{
			pt = _afxGlobalData._Zero.pt;
		}
		else if(bDirection == htmlDirRightToLeft)
		{
			pt.x = psize->cx - pt.x;
		}
	}

	if(fWidthAuto)
	{
		if(pci->_fUseOffset)
		{
			// measurer sets up _xLayoutOffsetInLine, use it instead of computing the
			// the xOffset. The xOffset computed through PointFromTp below, does not
			// take line shift (used for alignment) into account when the line containing
			// the current layout is being measured.
			pt.x = pci->_xOffsetInLine;
		}

		// pt.x includes the leading padding of the layout owner, but the available
		// size does not include the padding. So, subtract the leadding padding for
		// pt.x to make it relative to the leading padding(that is the zero of where
		// content actualy starts.
		pdpParent->GetPadding(pci, lPadding, pci->_smMode == SIZEMODE_MMWIDTH);

		pt.x -= (bDirection == htmlDirLeftToRight) ? lPadding[PADDING_LEFT] :
            lPadding[PADDING_RIGHT];

		*pxLeft = !fLeftAuto
			? uvLeft.XGetPixelValue(
			pci,
			psize->cx,
			GetFirstBranch()->GetFontHeightInTwips(&uvLeft))
			: 0;

		*pxRight = !fRightAuto
			? psize->cx - 
			uvRight.XGetPixelValue(
			pci,
			psize->cx,
			GetFirstBranch()->GetFontHeightInTwips(&uvRight))
			: psize->cx;
	}

	if(fHeightAuto)
	{
		*pyTop = !fTopAuto
			? uvTop.XGetPixelValue(
			pci,
			psize->cy,
			GetFirstBranch()->GetFontHeightInTwips(&uvTop))
			: pt.y;

		*pyBottom = !fBottomAuto
			? psize->cy - 
			uvBottom.XGetPixelValue(
			pci,
			psize->cy,
			GetFirstBranch()->GetFontHeightInTwips(&uvBottom))
			: psize->cy;
	}
}

//+---------------------------------------------------------------------------
//
//  Member:     Notify
//
//  Synopsis:   Handle notification
//
//----------------------------------------------------------------------------
void C1DLayout::Notify(CNotification* pNF)
{
	super::Notify(pNF);
}

void C1DLayout::ShowSelected(CTreePos* ptpStart, CTreePos* ptpEnd, BOOL fSelected,  BOOL fLayoutCompletelyEnclosed, BOOL fFireOM)
{
	Assert(ptpStart && ptpEnd && ptpStart->GetMarkup() == ptpStart->GetMarkup());
	CElement* pElement = ElementOwner();
	// For IE 5 we have decided to not draw the dithered selection at browse time for DIVs
	// people thought it weird.
	//
	// We do however draw it for XML Sprinkles.
    if(pElement->_etag==ETAG_GENERIC && 
        (fSelected && fLayoutCompletelyEnclosed) ||
        (!fSelected && ! fLayoutCompletelyEnclosed))
    {
        SetSelected(fSelected, TRUE);
    }
	else
	{
		_dp.ShowSelected( ptpStart, ptpEnd, fSelected);
	}
}



BOOL PointInRectAry(POINT pt, CStackDataAry<RECT, 1>& aryRects)
{
    for(int i=0; i<aryRects.Size(); i++)
    {
        RECT rc = aryRects[i];
        if(PtInRect(&rc, pt))
        {
            return TRUE;
        }
    }
    return FALSE;
}