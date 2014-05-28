
#include "stdafx.h"
#include "Adorner.h"

#include "View.h"
#include "../display/DispNode.h"
#include "../display/DispItemPlus.h"
#include "../display/DispRoot.h"
#include "../display/DispDefs.h"
#include "../layout/Layout.h"

//+====================================================================================
//
//  Member:     CAdorner, ~CAdorner
//
//  Synopsis:   Constructor/desctructor for CAdorner
//
//  Arguments:  pView    - Associated CView
//              pElement - Associated CElement
//
//------------------------------------------------------------------------------------
CAdorner::CAdorner(CView* pView, CElement* pElement)
{
    Assert(pView);

    _cRefs = 1;
    _pView = pView;
    _pDispNode = NULL;
    _pElement = (pElement&&pElement->_etag==ETAG_TXTSLAVE) ? pElement->MarkupMaster() : pElement;

    Assert(!_pElement || _pElement->IsInMarkup());
}

CAdorner::~CAdorner()
{
    Assert(!_cRefs);
    DestroyDispNode();
}

//+====================================================================================
//
// Method:EnsureDispNode
//
// Synopsis: Creates a DispItemPlus suitable for 'Adornment'
//
//------------------------------------------------------------------------------------
void CAdorner::EnsureDispNode()
{
    Assert(_pElement);
    Assert(_pView);
    Assert(_pView->IsInState(CView::VS_OPEN));

    CTreeNode* pTreeNode = _pElement->GetFirstBranch();

	if(!_pDispNode && pTreeNode)
	{
		_pDispNode = (CDispNode*)CDispRoot::CreateDispItemPlus(this,
			TRUE,
			FALSE,
			FALSE,
			DISPNODEBORDER_NONE,
			FALSE);

        if(_pDispNode)
        {
            _pDispNode->SetExtraCookie(GetDispCookie());
            _pDispNode->SetLayerType(GetDispLayer());
            _pDispNode->SetOwned(TRUE);
        }
    }

    if(_pDispNode && !_pDispNode->GetParentNode())
    {
        CNotification nf;

        nf.ElementAddAdorner(_pElement);
        nf.SetData((void*)this);

        _pElement->SendNotification(&nf);
    }
}

//+====================================================================================
//
// Method:GetBounds
//
// Synopsis: Return the bounds of the adorner in global coordinates
//
//------------------------------------------------------------------------------------
void CAdorner::GetBounds(CRect* prc) const
{
    Assert(prc);

    if(_pDispNode)
    {
        _pDispNode->GetClientRect(prc, CLIENTRECT_CONTENT);
        _pDispNode->TransformRect(prc, COORDSYS_CONTENT, COORDSYS_GLOBAL, FALSE);
    }
    else
    {
        *prc = _afxGlobalData._Zero.rc;
    }
}

//+====================================================================================
//
// Method:  GetRange
//
// Synopsis: Retrieve the cp range associated with an adorner
//
//------------------------------------------------------------------------------------
void CAdorner::GetRange(long* pcpStart, long* pcpEnd) const
{
    Assert(pcpStart);
    Assert(pcpEnd);

    if(_pElement && _pElement->GetFirstBranch())
    {
        long cch;

        _pElement->GetRange(pcpStart, &cch);

        *pcpEnd = *pcpStart + cch;
    }
    else
    {
        *pcpStart = *pcpEnd = 0;
    }
}

//+====================================================================================
//
// Method:  All CDispClient overrides
//
//------------------------------------------------------------------------------------
void CAdorner::GetOwner(CDispNode* pDispNode, void** ppv)
{
    Assert(pDispNode);
    Assert(ppv);
    *ppv = NULL;
}

void CAdorner::DrawClient(
      const RECT*	prcBounds,
	  const RECT*	prcRedraw,
	  CDispSurface*	pDispSurface,
	  CDispNode*	pDispNode,
	  void*			cookie,
	  void*			pClientData,
	  DWORD			dwFlags)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

void CAdorner::DrawClientBackground(
	 const RECT*    prcBounds,
	 const RECT*    prcRedraw,
	 CDispSurface*  pDispSurface,
	 CDispNode*     pDispNode,
	 void*          pClientData,
	 DWORD          dwFlags)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

void CAdorner::DrawClientBorder(
       const RECT*		prcBounds,
       const RECT*		prcRedraw,
       CDispSurface*	pDispSurface,
       CDispNode*		pDispNode,
       void*			pClientData,
       DWORD			dwFlags)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

void CAdorner::DrawClientScrollbar(
	   int			whichScrollbar,
	   const RECT*	prcBounds,
       const RECT*	prcRedraw,
	   LONG			contentSize,
	   LONG			containerSize,
	   LONG			scrollAmount,
	   CDispSurface* pSurface,
	   CDispNode*	pDispNode,
	   void*        pClientData,
	   DWORD        dwFlags)
{
	AssertSz(0, "Unexpected/Unimplemented method called in CAdorner");
}

void CAdorner::DrawClientScrollbarFiller(
       const RECT*		prcBounds,
       const RECT*		prcRedraw,
       CDispSurface*	pDispSurface,
       CDispNode*		pDispNode,
       void*			pClientData,
       DWORD			dwFlags)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

BOOL CAdorner::HitTestContent(
	  const POINT*	pptHit,
	  CDispNode*    pDispNode,
	  void*			pClientData)
{
	return FALSE;
}

BOOL CAdorner::HitTestFuzzy(
	  const POINT*	pptHitInParentCoords,
	  CDispNode*	pDispNode,
	  void*		    pClientData)
{
	return FALSE;
}

BOOL CAdorner::HitTestScrollbar(
	  int			whichScrollbar,
	  const POINT*	pptHit,
	  CDispNode*	pDispNode,
	  void*		pClientData)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
	return FALSE;
}

BOOL CAdorner::HitTestScrollbarFiller(
	  const POINT*	pptHit,
	  CDispNode*    pDispNode,
	  void*			pClientData)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
	return FALSE;
}

BOOL CAdorner::HitTestBorder(
	  const POINT*	pptHit,
	  CDispNode*    pDispNode,
	  void*			pClientData)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
	return FALSE;
}

LONG CAdorner::GetZOrderForSelf()
{
	return 0;
}

LONG CAdorner::GetZOrderForChild(void* cookie)
{
    return 0;
}

LONG CAdorner::CompareZOrder(CDispNode* pDispNode1, CDispNode* pDispNode2)
{
    Assert(pDispNode1);
    Assert(pDispNode2);
    Assert(pDispNode1 == _pDispNode);
    Assert(_pElement);

    CElement* pElement = ::GetDispNodeElement(pDispNode2);

    // Compare element z-order
    // If the same element is associated with both display nodes,
    // then the second display node is also for an adorner
    return (_pElement!=pElement ? _pElement->CompareZOrder(pElement) : 0);
}

CDispFilter* CAdorner::GetFilter()
{
    AssertSz(0, "Unexpected/Unimplemented method called in CAdorner");
    return NULL;
}

void CAdorner::HandleViewChange(
	  DWORD		    flags,
	  const RECT*	prcClient,	// global coordinates
	  const RECT*	prcClip,	// global coordinates
	  CDispNode*	pDispNode)
{
	AssertSz(0, "Unexpected/Unimplemented method called in CAdorner");
}

BOOL CAdorner::ProcessDisplayTreeTraversal(void* pClientData)
{
	return TRUE;
}

void CAdorner::NotifyScrollEvent(RECT* prcScroll, SIZE* psizeScrollDelta)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

DWORD CAdorner::GetClientLayersInfo(CDispNode* pDispNodeFor)
{
	return 0;
}

void CAdorner::DrawClientLayers(
	  const RECT*   prcBounds,
	  const RECT*   prcRedraw,
	  CDispSurface* pDispSurface,
	  CDispNode*    pDispNode,
	  void*         cookie,
	  void*         pClientData,
	  DWORD         dwFlags)
{
	AssertSz(0, "CAdorner- unexpected and unimplemented method called");
}

//+====================================================================================
//
//  Member:     DestroyDispNode
//
//  Synopsis:   Destroy the adorner display node (if any)
//
//------------------------------------------------------------------------------------
void CAdorner::DestroyDispNode()
{
    if(_pDispNode)
    {
        Assert(_pView);
        Assert(_pView->IsInState(CView::VS_OPEN));    
        _pDispNode->Destroy();
        _pDispNode = NULL;
    }
}



//+====================================================================================
//
//  Member:     CFocusAdorner, ~CFocusAdorner
//
//  Synopsis:   Constructor/desctructor for CFocusAdorner
//
//  Arguments:  pView    - Associated CView
//              pElement - Associated CElement
//
//------------------------------------------------------------------------------------
CFocusAdorner::CFocusAdorner(CView* pView) : CAdorner(pView)
{
    Assert(pView);
}

CFocusAdorner::~CFocusAdorner()
{
    Assert(_pView);

    delete _pShape;

    if(_pView->_pFocusAdorner == this)
    {
        _pView->_pFocusAdorner = NULL;
    }
}

//+====================================================================================
//
// Method:  EnsureDispNode
//
// Synopsis: Ensure the display node is created and properly inserted in the display tree
//
//------------------------------------------------------------------------------------
void CFocusAdorner::EnsureDispNode()
{
    Assert(_pElement);
    Assert(_pView->IsInState(CView::VS_OPEN));

    if(_pShape)
    {
        CTreeNode* pTreeNode = _pElement->GetFirstBranch();

        if(!_pDispNode && pTreeNode)
        {
            _dnl = _pElement->GetFirstBranch()->IsPositionStatic()
                        ? DISPNODELAYER_FLOW
                        : _pElement->GetFirstBranch()->GetCascadedzIndex()>=0?DISPNODELAYER_POSITIVEZ:DISPNODELAYER_NEGATIVEZ;
            _adl = _dnl==DISPNODELAYER_FLOW ? ADL_TOPOFFLOW : ADL_ONELEMENT;
        }

        if(!_pDispNode || !_pDispNode->GetParentNode())
        {
            super::EnsureDispNode();

            if(_pDispNode)
            {
                _pDispNode->SetAffectsScrollBounds(FALSE);
            }
        }
    }
    else
    {
        DestroyDispNode();
    }
}

//+====================================================================================
//
// Method:  EnsureFocus
//
// Synopsis: Ensure focus display node exists and is properly inserted in the display tree
//
//------------------------------------------------------------------------------------
void CFocusAdorner::EnsureFocus()
{
    Assert(_pElement);
    if(_pShape && (!_pDispNode || !_pDispNode->GetParentNode()))
    {
        EnsureDispNode();
    }
}

//+====================================================================================
//
// Method:  SetElement
//
// Synopsis: Set the associated element
//
//------------------------------------------------------------------------------------
void CFocusAdorner::SetElement(CElement* pElement, long iDivision)
{
    Assert(_pView->IsInState(CView::VS_OPEN));
    Assert(pElement);
    Assert(pElement->IsInMarkup());

    if(pElement->_etag == ETAG_TXTSLAVE)
    {
        pElement = pElement->MarkupMaster();
    }

    if(pElement!=_pElement || iDivision!=_iDivision)
    {
        _pElement = pElement;
        _iDivision = iDivision;

        DestroyDispNode();
        ShapeChanged();
    }

    Assert(_pElement);
}

//+====================================================================================
//
// Method: PositionChanged
//
// Synopsis: Hit the Layout for the size you should be and ask your adorner for it to give
//           you your position based on this
//
//------------------------------------------------------------------------------------
void CFocusAdorner::PositionChanged(const CSize* psize)
{
    Assert(_pElement);
    Assert(_pElement->GetFirstBranch());
    Assert(_pView->IsInState(CView::VS_OPEN));

    if(_pDispNode)
    {
        CLayout*	pLayout		= _pElement->GetUpdatedNearestLayout();
        CTreeNode*	pTreeNode	= _pElement->GetFirstBranch();
        BOOL		fRelative	= pTreeNode->GetCharFormat()->_fRelative;
        CDispNode*	pDispParent	= _pDispNode->GetParentNode();
        CDispNode*	pDispNode	= NULL;

        Assert(_pShape);

        // Get the display node which contains the element with focus
        // (If the focus display node is not yet anchored in the display tree, pretend the element
        //  is not correct as well. After the focus display node is anchored, this routine will
        //  get called again and can correctly associate the display nodes at that time.)
        if(pDispParent)
        {
            // BUGBUG: Move this logic down into GetElementDispNode (passing a flag so that GetElementDispNode
            //         can distinguish between "find nearest" and "find exact" with this call being a "find nearest"
            //         and virtually all others being a "find exact" (brendand)
            CElement* pDisplayElement = NULL;

            if(!pTreeNode->IsPositionStatic() || _pElement->HasLayout())
            {
                pDisplayElement = _pElement;
            }
            else if(!fRelative)
            {
                pDisplayElement = pLayout->ElementOwner();
            }
            else
            {
                CTreeNode* pDisplayNode = pTreeNode->GetCurrentRelativeNode(pLayout->ElementOwner());

                Assert(pDisplayNode); // This should never be NULL, but be safe anyway
                if(pDisplayNode)
                {
                    pDisplayElement = pDisplayNode->Element();
                }
            }

            Assert(pDisplayElement); // This should never be NULL, but be safe anyway
            if(pDisplayElement)
            {
                pDispNode = pLayout->GetElementDispNode(pDisplayElement);
            }
        }

        // Verify that the display node which contains the element with focus and the focus display node
        // are both correctly anchored in the display tree (that is, have a common parent)
        // (If they do not, this routine will get called again once both get correctly anchored
        //  after layout is completed)
        if(pDispNode)
        {
            CDispNode* pDispNodeTemp;

			for(pDispNodeTemp=pDispNode;
				pDispNodeTemp&&pDispNodeTemp!=pDispParent;
				pDispNodeTemp=pDispNodeTemp->GetParentNode()) ;

			if(!pDispNodeTemp)
            {
                pDispNode = NULL;
            }

            Assert(!pDispNode || pDispNodeTemp==pDispParent);
        }

        if(pDispNode)
        {
            if(!psize || _dnl!=DISPNODELAYER_FLOW)
            {
                CPoint ptFromOffset(_afxGlobalData._Zero.pt);
                CPoint ptToOffset(_afxGlobalData._Zero.pt);

                if(!_fTopLeftValid)
                {
                    CRect rc;
                    _pShape->GetBoundingRect(&rc);
                    _pShape->OffsetShape(-rc.TopLeft().AsSize());

                    _ptTopLeft = rc.TopLeft();

                    if(!_pElement->HasLayout() && fRelative)
                    {
                        CPoint ptOffset;

                        _pElement->GetUpdatedParentLayout()->GetFlowPosition(pDispNode, &ptOffset);

                        _ptTopLeft -= ptOffset.AsSize();
                    }

                    _fTopLeftValid = TRUE;
                }

                pDispNode->TransformPoint(&ptFromOffset, COORDSYS_CONTENT, COORDSYS_GLOBAL);
                _pDispNode->TransformPoint(&ptToOffset, COORDSYS_PARENT, COORDSYS_GLOBAL);

                _pDispNode->SetPosition(_ptTopLeft.AsSize()+(ptFromOffset-ptToOffset).AsPoint());
            }
            else
            {
                CPoint pt = _pDispNode->GetPosition();

                Assert(_fTopLeftValid);

                pt += *psize;

                _pDispNode->SetPosition(pt);
            }
        }
        // If the display node containing the element with focus is not correctly placed in the display
        // tree, remove the focus display node as well to prevent artifacts
        else
        {
            _pDispNode->ExtractFromTree();
        }
    }
}

//+====================================================================================
//
// Method: ShapeChanged
//
// Synopsis: Hit the Layout for the size you should be and ask your adorner for it to give
//           you your position based on this
//
//------------------------------------------------------------------------------------
void CFocusAdorner::ShapeChanged()
{
    Assert(_pView->IsInState(CView::VS_OPEN));
    Assert(_pElement);

    delete _pShape;
    _pShape = NULL;

    _fTopLeftValid = FALSE;

    CDocument* pDoc = _pView->Doc();
    CDocInfo dci(pDoc->_dci);
    CShape* pShape;

    if(_pElement->_etag==ETAG_DIV && DYNCAST(CDivElement, _pElement)->GetAAnofocusrect())
    {
        pShape = NULL;
    }
    else
    {
        _pElement->GetFocusShape(_iDivision, &dci, &pShape);
    }

    if(pShape)
    {
        CRect rc;

        pShape->GetBoundingRect(&rc);

        if(rc.IsEmpty())
        {
            delete pShape;
        }
        else
        {
            _pShape = pShape;
        }
    }
    
    EnsureDispNode();

    if(_pDispNode)
    {
        CRect rc;

        Assert(_pShape);

        _pShape->GetBoundingRect(&rc);
        _pDispNode->SetSize(rc.Size(), TRUE);
    }
}

//+====================================================================================
//
// Method: Draw
//
// Synopsis: Wraps call to IElementAdorner.Draw - to allow adorner
//           to draw.
//
//------------------------------------------------------------------------------------
void CFocusAdorner::DrawClient(
	   const RECT*		prcBounds,
	   const RECT*		prcRedraw,
	   CDispSurface*	pDispSurface,
	   CDispNode*		pDispNode,
	   void*			cookie,
	   void*			pClientData,
	   DWORD			dwFlags)
{
	Assert(_pElement);
    if(!_pElement->IsEditable(TRUE) && _pView->Doc()->HasFocus()
        && !(_pView->Doc()->_wUIState&UISF_HIDEFOCUS))
    {
		Assert(pClientData);

		CFormDrawInfo* pDI = (CFormDrawInfo*)pClientData;
        CSetDrawSurface sds(pDI, prcBounds, prcRedraw, pDispSurface);

        _pShape->DrawShape(pDI);
    }
}

//+====================================================================================
//
// Method: GetZOrderForSelf
//
// Synopsis: IDispClient - get z-order.
//
//------------------------------------------------------------------------------------
LONG CFocusAdorner::GetZOrderForSelf()
{
    Assert(_pElement);
    Assert(!_pElement->GetFirstBranch()->IsPositionStatic());
    Assert(_dnl != DISPNODELAYER_FLOW);
    return _pElement->GetFirstBranch()->GetCascadedzIndex();
}