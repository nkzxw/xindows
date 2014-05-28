
#include "stdafx.h"
#include "TreeNode.h"

BEGIN_TEAROFF_TABLE_NAMED(CTreeNode, s_apfnNodeVTable)
    TEAROFF_METHOD(CTreeNode, &GetInterface, getinterface, (REFIID riid, void** ppv))
    TEAROFF_METHOD_(CTreeNode, &PutRef, putref, ULONG, ())
    TEAROFF_METHOD_(CTreeNode, &RemoveRef, removeref, ULONG, ())
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
    TEAROFF_METHOD_NULL
END_TEAROFF_TABLE()

CTreeNode::CTreeNode(CTreeNode* pParent, CElement* pElement/*=NULL*/)
{
    Assert(_pNodeParent == NULL);
    Assert(_pElement == NULL);

    _iFF = _iCF = _iPF = -1;

    SetElement(pElement);
    SetParent(pParent);

    Assert(pElement && pElement->Doc() && pElement->Doc()->AreLookasidesClear(this, LOOKASIDE_NODE_NUMBER));
}

HRESULT CTreeNode::GetElementInterface(REFIID riid, void** ppVoid)
{
    CElement* pElement = Element();

    Assert(pElement);

    if(pElement->GetFirstBranch() == this)
    {
        return pElement->QueryInterface(riid, ppVoid);
    }
    else
    {
        return GetInterface(riid, ppVoid);
    }
}

HRESULT CTreeNode::GetInterface(REFIID iid, void** ppv)
{
    void*   pv = NULL;
    void*   pvWhack;
    void*   apfnWhack;
    HRESULT hr = E_NOINTERFACE;
    BOOL    fReuseTearoff = FALSE;
    const IID* piidDisp;

    *ppv = NULL;

    AssertSz(_pElement, "_pElement is NULL in CTreeNode::GetInterface -- VERY BAD!!!");

    // handle IUnknown when tearoff is already created
    if(iid==IID_IUnknown && HasPrimaryTearoff())
    {
        IUnknown* pTearoff;

        pTearoff = GetPrimaryTearoff();

        Assert(pTearoff);

        pTearoff->AddRef();

        *ppv = pTearoff;

        return S_OK;
    }

    if(iid == CLSID_CTreeNode)
    {
        *ppv = this;
        return S_OK;
    }
    else if(iid==CLSID_CElement || iid==CLSID_CTextSite)
    {
        return _pElement->QueryInterface(iid, ppv);
    }

    // Create a tearoff to return

    // Get the interface from the element
    hr = _pElement->PrivateQueryInterface(iid, &pv);
    if(hr)
    {
        pv = NULL;
        goto Cleanup;
    }

    // Whack in our node information, or the primary tearoff
    if(HasPrimaryTearoff())
    {
        pvWhack = (void*)GetPrimaryTearoff();
        apfnWhack = *(void**)pvWhack;
    }
    else
    {
        pvWhack = this;
        apfnWhack = (void*)s_apfnNodeVTable;
    }

    Assert(pvWhack);

    // InstallTearOffObject puts a pointer to an object
    // into pvObject2 of the tearoff.  This means that we
    // are assuming that when an element is QI'd, the interfaces
    // listed below will be returned as tearoffs.  Also, every
    // interface that uses the eax trick for passing context
    // through must be in this list.  Otherwise we will
    // create another tearoff pointing to the tearoff.
    piidDisp = _pElement->BaseDesc()->_piidDispinterface;
    piidDisp = piidDisp ? piidDisp : &IID_NULL;

    if( iid==*piidDisp                 ||
        iid==IID_IHTMLElement          ||
        iid==IID_IHTMLControlElement)
    {
        hr = InstallTearOffObject(pv, pvWhack, apfnWhack, QI_MASK);
        if(hr)
        {
            goto Cleanup;
        }

        *ppv = pv;
        fReuseTearoff = TRUE;
    }
    else
    {
        hr = CreateTearOffThunk(pv, *(void**)pv, NULL, ppv, pvWhack,
            apfnWhack, QI_MASK, NULL);
        ((IUnknown *)pv)->Release();
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    if(!HasPrimaryTearoff())
    {
        // This tearoff that we just created is now the primary tearoff
        SetPrimaryTearoff((IUnknown*)*ppv);
    }

    if(!fReuseTearoff)
    {
        ((IUnknown*)*ppv)->AddRef();
    }

    hr = S_OK;

Cleanup:
    if(hr && pv)
    {
        ((IUnknown*)pv)->Release();
    }

    RRETURN(hr);
}

ULONG CTreeNode::PutRef(void)
{
    Assert(!HasPrimaryTearoff());

    // We do this so that we know that the node will
    // die before the element.  If it was the other way
    // around, we wouldn't be able to get to the doc to
    // del our primary tearoff lookaside pointer
    Element()->AddRef();

    return 1;
}

ULONG CTreeNode::RemoveRef(void)
{
    CElement* pElement = Element();

    Assert(HasPrimaryTearoff());

    DelPrimaryTearoff();

    if(!_fInMarkup)
    {
        delete this;
    }

    // Release the ref that we put on the element above.
    pElement->Release();

    return 1;
}

ULONG CTreeNode::NodeAddRef(void)
{
    IUnknown* pTearoff = NULL;

    if(!HasPrimaryTearoff())
    {
        // Use GetInterface to create the primary interface
        Verify(!GetInterface(IID_IUnknown, (void**)&pTearoff));

        Assert(pTearoff);

        return 1;
    }
    else
    {
        pTearoff = GetPrimaryTearoff();

        Assert( pTearoff );

        return pTearoff->AddRef();
    }
}

ULONG CTreeNode::NodeRelease(void)
{
    IUnknown* pTearoff;

    Assert(HasPrimaryTearoff());

    pTearoff = GetPrimaryTearoff();

    Assert(pTearoff);

    return pTearoff->Release();
}

void CTreeNode::PrivateEnterTree()
{
    Assert(!_fInMarkup);
    _fInMarkup = TRUE;
}

void CTreeNode::PrivateExitTree()
{
    PrivateMakeDead();
    PrivateMarkupRelease();
}

void CTreeNode::PrivateMakeDead()
{
    Assert(_fInMarkup);
    _fInMarkup = FALSE;

    VoidCachedNodeInfo();

    SetParent(NULL);
}

void CTreeNode::PrivateMarkupRelease()
{
    if(!HasPrimaryTearoff())
    {
        delete this;
    }
}

void CTreeNode::SetElement(CElement* pElement)
{
    _pElement = pElement;
    if(pElement)
    {
        _etag = pElement->_etag;
    }
}

void CTreeNode::SetParent(CTreeNode* pNodeParent)
{
    _pNodeParent = pNodeParent;
}

CTreeNode* CTreeNode::NextBranch()
{
    Assert(IsInMarkup());

    CTreePos* ptpCurr = GetEndPos();

    if(ptpCurr->IsEdgeScope())
    {
        return NULL;
    }
    else
    {
        CElement* pElement = Element();
        CTreeNode* pNodeCurr;

        do
        {
            ptpCurr = ptpCurr->NextTreePos();

            Assert(ptpCurr->IsNode());

            pNodeCurr = ptpCurr->Branch();
        } while(pNodeCurr->Element() != pElement);

        return pNodeCurr;
    }
}

CTreeNode* CTreeNode::PreviousBranch()
{
    Assert(IsInMarkup());

    CTreePos* ptpCurr = GetBeginPos();

    if(ptpCurr->IsEdgeScope())
    {
        return NULL;
    }
    else
    {
        CElement* pElement = Element();
        CTreeNode* pNodeCurr;

        do
        {
            ptpCurr = ptpCurr->PreviousTreePos();

            Assert(ptpCurr->IsNode());

            pNodeCurr = ptpCurr->Branch();
        } while(pNodeCurr->Element() != pElement);

        return pNodeCurr;
    }
}

CDocument* CTreeNode::Doc()
{
    Assert(_pElement);
    return _pElement->Doc();
}

CRootElement* CTreeNode::IsRoot()
{
    Assert(!IsDead());
    return Element()->IsRoot();
}

CMarkup* CTreeNode::GetMarkup()
{
    Assert(!IsDead());
    return Element()->GetMarkup();
}

CRootElement* CTreeNode::MarkupRoot()
{
    Assert(!IsDead());
    return Element()->MarkupRoot();
}

// Does the element that this node points to have currency?
BOOL CTreeNode::HasCurrency()
{
    return (this && Element()->HasCurrency());
}

BOOL CTreeNode::IsContainer()
{
    return Element()->IsContainer();
}

CTreeNode* CTreeNode::GetContainerBranch()
{
    CTreeNode* pNode = this;

    for(; pNode; pNode=pNode->Parent())
    {
        if(pNode->IsContainer())
        {
            break;
        }
    }

    return pNode;
}

BOOL CTreeNode::SupportsHtml()
{
    CElement* pContainer = GetContainer();

    return (pContainer&&pContainer->HasFlag(TAGDESC_ACCEPTHTML));
}

CTreeNode* CTreeNode::Ancestor(ELEMENT_TAG etag)
{
    CTreeNode* context = this;

    while(context && context->Tag()!=etag)
    {
        context = context->Parent();
    }

    return context;
}

CTreeNode* CTreeNode::Ancestor(ELEMENT_TAG* arytag)
{
    CTreeNode*      context = this;
    ELEMENT_TAG     etag;
    ELEMENT_TAG*    petag;

    while(context)
    {
        etag = context->Tag();

        for(petag=arytag; *petag; petag++)
        {
            if(etag == *petag)
            {
                return context;
            }
        }

        context = context->Parent();
    }

    return context; // NULL context
}

//+---------------------------------------------------------------------------
//
// CTreeNode::RenderParent()
// CTreeNode::ZParent()
// CTreeNode::ClipParent()
//
// Parent accessor methods used for positioning support.
//
// The following chart defines the different parents that are used for
// positioning.  Each parent determines different parameters for any
// relatively positioned or absolutely positioned element.
//
//     PARENT        RELATIVE            ABSOLUTE               PARENT TYPE       USED IN
//     ------        --------            --------               -----------       -------
//
//  "ElementParent"  Coordinates       Not meaningful             Element         GetRenderPosition/PositionObjects (implicit)
//
//  "ParentSite"           Percent/Auto/CalcSize (both)             Site          Measuring (implicit)
//
//  "RenderParent"                Painting (both)                   Site          GetSiteDrawList/HitTestPoint
//
//  "ZParent"        Z-Order            Z-Order/Coordinates    Element or Site    GetElementsInZOrder/GetRenderPosition
//
//  "Clip Parent"  Clip Rect/Auto/SetPos   Clip Rect/SetPos     Absolute Site     SetPosition/PositionObjects
//
//
//  If the ZParent is a site, then it is the same as the RenderParent.
//
//----------------------------------------------------------------------------
CTreeNode* CTreeNode::ZParentBranch()
{
    CTreeNode* pNode;

    if(Element()->IsPositionStatic() && !GetCharFormat()->_fRelative)
    {
        return GetUpdatedParentLayoutNode();
    }

    pNode = this;
    if(pNode)
    {
        pNode = pNode->Parent();
    }
    for(; pNode; pNode=pNode->Parent())
    {
        if(pNode->Element()->IsZParent())
        {
            break;
        }
    }

    // If pNode is NULL then 'this' is the BODY or this node is not parented
    // into the main document tree.  Return the body branch for the first
    // case and NULL for the second.
    return (pNode ? pNode : ((Tag()==ETAG_BODY)?this:NULL));
}

CTreeNode* CTreeNode::RenderParentBranch()
{
    CTreeNode* pNode;

    if(Element()->IsPositionStatic())
    {
        return GetUpdatedParentLayoutNode();
    }


    pNode = this;
    if(pNode)
    {
        pNode = pNode->Parent();
    }
    for(; pNode; pNode=pNode->Parent())
    {
        if(pNode->NeedsLayout() && pNode->Element()->IsZParent())
        {
            break;
        }
    }

    return (pNode ? pNode : ((Tag()==ETAG_BODY)?this:NULL));
}

CTreeNode* CTreeNode::ClipParentBranch()
{
    CTreeNode* pNode;

    if(Element()->IsPositionStatic())
    {
        return GetUpdatedParentLayoutNode();
    }

    pNode = this;
    if(pNode)
    {
        pNode = pNode->Parent();
    }
    for(; pNode; pNode=pNode->Parent())
    {
        if(pNode->NeedsLayout() && pNode->Element()->IsClipParent())
        {
            break;
        }
    }

    Assert(pNode);

    return pNode;
}

CTreeNode* CTreeNode::ScrollingParentBranch()
{
    CTreeNode* pNode = NULL;

    if(Parent())
    {
        pNode = GetUpdatedParentLayoutNode();

        for(; pNode; pNode=pNode->GetUpdatedParentLayoutNode())
        {
            if(pNode->Element()->IsScrollingParent())
            {
                break;
            }
        }
    }

    return pNode;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTreeNode::GetFirstCommonAncestor, public
//
//  Synopsis:   Returns the first branch whose element is common to
//              both of the branches specified.
//
//  Arguments:  [pNode] -- branch to find common parent of with [this]
//              [pEltStop]  -- Stop walking tree if you hit this element.
//                             If NULL then search to root.
//
//  Returns:    Branch with common element from first starting point.
//
//----------------------------------------------------------------------------
CTreeNode* CTreeNode::GetFirstCommonAncestor(CTreeNode* pNodeTwo, CElement* pEltStop)
{
    CTreeNode* pNode;
    CElement* pElement;

    if(pNodeTwo->Element() == Element())
    {
        return this;
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        pElement->_fFirstCommonAncestor = 0;

        if(pElement == pEltStop)
        {
            break;
        }
    }

    for(pNode=pNodeTwo; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        pElement->_fFirstCommonAncestor = 1;

        if(pElement == pEltStop)
        {
            break;
        }
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        if(pElement->_fFirstCommonAncestor)
        {
            return pNode;
        }

        Assert(pElement != pEltStop);
    }

    return NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTreeNode::GetFirstCommonBlockOrLayoutAncestor, public
//
//  Synopsis:   Returns the first branch whose element is common to
//              both of the branches specified and is a block or layout element
//
//  Arguments:  [pNode] -- branch to find common parent of with [this]
//              [pEltStop]  -- Stop walking tree if you hit this element.
//                             If NULL then search to root.
//
//  Returns:    Branch with common element from first starting point.
//
//----------------------------------------------------------------------------
CTreeNode* CTreeNode::GetFirstCommonBlockOrLayoutAncestor(CTreeNode* pNodeTwo, CElement* pEltStop)
{
    CTreeNode* pNode;
    CElement* pElement;

    if(pNodeTwo->Element() == Element())
    {
        return this;
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        pElement->_fFirstCommonAncestor = 0;

        if(pElement == pEltStop)
        {
            break;
        }
    }

    for(pNode=pNodeTwo; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        pElement->_fFirstCommonAncestor = 1;

        if(pElement == pEltStop)
        {
            break;
        }
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        pElement = pNode->Element();

        if((pElement->HasLayout() || pElement->IsBlockElement())
            && pElement->_fFirstCommonAncestor)
        {
            return pNode;
        }

        Assert(pElement != pEltStop);
    }

    return NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTreeNode::GetFirstCommonAncestorNode, public
//
//  Synopsis:   Returns the first node that is common to
//              both of the branches specified.
//
//  Arguments:  [pNode] -- branch to find common parent of with [this]
//              [pEltStop]  -- Stop walking tree if you hit this element.
//                             If NULL then search to root.
//
//  Returns:    Branch with common node from first starting point.
//
//----------------------------------------------------------------------------
CTreeNode* CTreeNode::GetFirstCommonAncestorNode(CTreeNode* pNodeTwo, CElement* pEltStop)
{
    CTreeNode* pNode;

    if(this == pNodeTwo)
    {
        return this;
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        pNode->_fFirstCommonAncestorNode = 0;

        if(pNode->Element() == pEltStop)
        {
            break;
        }
    }

    for(pNode=pNodeTwo; pNode; pNode=pNode->Parent())
    {
        pNode->_fFirstCommonAncestorNode = 1;

        if(pNode->Element() == pEltStop)
        {
            break;
        }
    }

    for(pNode=this; pNode; pNode=pNode->Parent())
    {
        if(pNode->_fFirstCommonAncestorNode)
        {
            return pNode;
        }

        Assert(pNode->Element() != pEltStop);
    }

    return NULL;
}

// The following function is used by CTxtRange::GetExtendedSelectionInfo
CTreeNode* CTreeNode::SearchBranchForPureBlockElement(CFlowLayout* pFlowLayout)
{
    return pFlowLayout->GetContentMarkup()->SearchBranchForBlockElement(this, pFlowLayout);
}

//+----------------------------------------------------------------------------
//
//  Member:     CTreeNode::SearchBranchToFlowLayoutForTag
//
//  Synopsis:   Looks up the parent chain for the first element which
//              matches the tag.  No stopper element here, but stops
//              at the first text site it encounters.
//
//-----------------------------------------------------------------------------
CTreeNode* CTreeNode::SearchBranchToFlowLayoutForTag(ELEMENT_TAG etag)
{
    CTreeNode* pNode = this;

    do
    {
        if(pNode->Tag() == etag)
        {
            return pNode;
        }

        pNode = pNode->Parent();
    } while(pNode && !pNode->HasFlowLayout());

    return NULL;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::SearchBranchToRootForTag
//
//  Synopsis:   Looks up the parent chain for the first element which
//              matches the tag.  No stopper element here, goes all the
//              way up to the <HTML> tag.
//
//-----------------------------------------------------------------------------
CTreeNode* CTreeNode::SearchBranchToRootForTag(ELEMENT_TAG etag)
{
    CTreeNode* pNode = this;

    do
    {
        if(pNode->Tag() == etag)
        {
            return pNode;
        }
    } while((pNode=pNode->Parent()) != NULL);

    return NULL;
}

//+----------------------------------------------------------------------------
//
//  Member:     CTreeNode::SearchBranchToRootForScope
//
//  Synopsis:   Looks up the parent chain for the first element which
//              has the same scope as the given element.  Will not stop
//              until there is no parent.
//
//-----------------------------------------------------------------------------
CTreeNode* CTreeNode::SearchBranchToRootForScope(CElement* pElementFindMe)
{
    CTreeNode* pNode = this;

    do
    {
        if(pNode->Element() == pElementFindMe)
        {
            return pNode;
        }
    } while((pNode=pNode->Parent()) != NULL);

    return NULL;
}

//+----------------------------------------------------------------------------
//
//  Member:     CTreeNode::SearchBranchToRootForNode
//
//  Synopsis:   Looks up the parent chain for the given element. Will not stop
//              until there is no parent.
//
//-----------------------------------------------------------------------------
BOOL CTreeNode::SearchBranchToRootForNode(CTreeNode* pNodeFindMe)
{
    CTreeNode* pNode = this;

    do
    {
        if(pNode == pNodeFindMe)
        {
            return TRUE;
        }
    } while((pNode=pNode->Parent()) != NULL);

    return FALSE;
}

//+----------------------------------------------------------------------------
//
// Member:      GetCurrentRelativeNode
//
// Synopsis:    Get the node that is relative, which causes the current
//              chunk to be relative.
//
//-----------------------------------------------------------------------------
CTreeNode* CTreeNode::GetCurrentRelativeNode(CElement* pElementFL)
{
    const CFancyFormat* pFF;
    CTreeNode* pNodeStart = this;

    while(pNodeStart && DifferentScope(pElementFL, pNodeStart))
    {
        pFF = pNodeStart->GetFancyFormat();

        // BUGBUG (jbeda): I'm pretty sure this is wrong for some
        // overlapped cases

        // relatively positioned layout elements are to be ignored
        if(!pNodeStart->Element()->NeedsLayout() && pFF->_fRelative)
        {
            return pNodeStart;
        }

        pNodeStart = pNodeStart->Parent();
    }
    return NULL;
}

CLayout* CTreeNode::GetCurLayout()
{
    return Element()->GetCurLayout();
}

// CTreeNode - layout related functions
BOOL CTreeNode::HasLayout()
{
    return Element()->HasLayout();
}

CLayout* CTreeNode::GetCurNearestLayout()
{
    if(this)
    {
        CLayout* pLayout = GetCurLayout();

        return (pLayout?pLayout:GetCurParentLayout());
    }
    else
    {
        return NULL;
    }
}

CTreeNode* CTreeNode::GetCurNearestLayoutNode()
{
    return (HasLayout()?this:GetCurParentLayoutNode());
}

CElement* CTreeNode::GetCurNearestLayoutElement()
{
    return (HasLayout()?Element():GetCurParentLayoutElement());
}

CLayout* CTreeNode::GetCurParentLayout()
{
    CTreeNode* pTreeNode = GetCurParentLayoutNode();
    return  pTreeNode?pTreeNode->GetCurLayout():NULL;
}

CTreeNode* CTreeNode::GetCurParentLayoutNode()
{
    CTreeNode* pNode = this;

    if(pNode)
    {
        pNode = pNode->Parent();
    }

    while(pNode)
    {
        if(pNode->HasLayout())
        {
            return pNode;
        }

        pNode = pNode->Parent();
    }
    return NULL;
}

CElement* CTreeNode::GetCurParentLayoutElement()
{
    return GetCurParentLayoutNode()->SafeElement();
}

CLayout* CTreeNode::GetUpdatedLayout()
{
    return Element()->GetUpdatedLayout();
}

CLayout* CTreeNode::GetUpdatedLayoutPtr()
{
    return Element()->GetUpdatedLayoutPtr();
}

BOOL CTreeNode::NeedsLayout()
{
    return Element()->NeedsLayout();
}

CLayout* CTreeNode::GetUpdatedNearestLayout()
{
    if(this)
    {
        CLayout* pLayout = GetUpdatedLayout();

        return (pLayout?pLayout:GetUpdatedParentLayout());
    }
    else
    {
        return NULL;
    }
}

CTreeNode* CTreeNode::GetUpdatedNearestLayoutNode()
{
    return (NeedsLayout()?this:GetUpdatedParentLayoutNode());
}

CElement* CTreeNode::GetUpdatedNearestLayoutElement()
{
    return (NeedsLayout()?Element():GetUpdatedParentLayoutElement());
}

CLayout* CTreeNode::GetUpdatedParentLayout()
{
    CTreeNode* pTreeNode = GetUpdatedParentLayoutNode();
    return  pTreeNode?pTreeNode->GetUpdatedLayoutPtr():NULL;
}

CTreeNode* CTreeNode::GetUpdatedParentLayoutNode()
{
    CTreeNode* pNode = this;

    if(pNode)
    {
        pNode = pNode->Parent();
    }

    while(pNode)
    {
        if(pNode->NeedsLayout())
        {
            return pNode;
        }

        pNode = pNode->Parent();
    }
    return NULL;
}

CElement* CTreeNode::GetUpdatedParentLayoutElement()
{
    return GetUpdatedParentLayoutNode()->SafeElement();
}

CFlowLayout* CTreeNode::GetFlowLayout()
{
    CTreeNode* pNode = this;
    CFlowLayout* pFL;

    while(pNode)
    {
        pFL = pNode->HasFlowLayout();
        if(pFL)
        {
            return pFL;
        }
        pNode = pNode->Parent();
    }

    return NULL;
}

CTreeNode* CTreeNode::GetFlowLayoutNode()
{
    CTreeNode* pNode = this;

    while(pNode)
    {
        if(pNode->HasFlowLayout())
        {
            return pNode;
        }
        pNode = pNode->Parent();
    }

    return NULL;
}

CElement* CTreeNode::GetFlowLayoutElement()
{
    return GetFlowLayoutNode()->SafeElement();
}

CFlowLayout* CTreeNode::HasFlowLayout()
{
    return Element()->HasFlowLayout();
}

htmlBlockAlign CTreeNode::GetParagraphAlign(BOOL fOuter)
{
    return htmlBlockAlign(GetParaFormat()->GetBlockAlign(!fOuter));
}

htmlControlAlign CTreeNode::GetSiteAlign()
{
    return htmlControlAlign(GetFancyFormat()->_bControlAlign);
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedclearLeft
//
// Synopsis:
//
//-----------------------------------------------------------------------------
BOOL CTreeNode::GetCascadedclearLeft()
{
    return !!GetFancyFormat()->_fClearLeft;
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedclearRight
//
// Synopsis:
//
//-----------------------------------------------------------------------------
BOOL CTreeNode::GetCascadedclearRight()
{
    return !!GetFancyFormat()->_fClearRight;
}

//+------------------------------------------------------------------------
//
//  Member:     CTreeNode::IsInlinedElement
//
//  Synopsis:   Determines if the element is rendered inflow or not, If the
//              element is not absolutely positioned and is left or right
//              aligned with hr and legend an exception (they can only be
//              aligned but nothing should wrap around them).
//              For non-sites we return TRUE.
//
//  Returns:    BOOL indicating whether or not the site is inlined
//
//-------------------------------------------------------------------------
BOOL CTreeNode::IsInlinedElement()
{
    if(NeedsLayout())
    {
        const CFancyFormat* pFF = GetFancyFormat();

        return pFF->_bPositionType!=stylePositionabsolute&&!pFF->_fAlignedLayout;
    }

    return TRUE;
}

//+----------------------------------------------------------------
//
//  Member:     CTreeNode::Depth
//
//  Synopsis:   Finds the depth of the node in the html tree
//
//  Returns:    int
//
//-----------------------------------------------------------------
int CTreeNode::Depth() const
{
    CTreeNode* pNode = const_cast<CTreeNode*>(this);
    int nDepth = 0;

    while(pNode)
    {
        nDepth++, pNode = pNode->Parent();
    }
    Assert(nDepth > 0);

    return nDepth;
}

#define MAX_FORMAT_INDEX    0x7FFF
//+---------------------------------------------------------------------
//
//  Member:     CTreeNode::CacheNewFormats
//
//  Synopsis:   This function is called on conclusion on ComputeFormats
//              It caches the XFormat's we have just computed.
//              This exists so we can share more code between
//              CElement::ComputeFormats and CTable::ComputeFormats
//
//  Arguments:  pCFI - Format Info needed for cascading
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------
HRESULT CTreeNode::CacheNewFormats(CFormatInfo* pCFI)
{
    THREADSTATE* pts = GetThreadState();
    LONG lIndex, iExpando = -1;
    HRESULT hr = S_OK;

    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || (_iCF!=-1 && _iPF!=-1));
    Assert(_iFF == -1);

    if(_iCF == -1)
    {
        // CCharFormat
        if(!pCFI->_fPreparedCharFormat)
        {
            _iCF = pCFI->_icfSrc;
            pts->_pCharFormatCache->AddRefData(_iCF);
        }
        else
        {
            pCFI->_cf()._bCrcFont = pCFI->_cf().ComputeFontCrc();
            pCFI->_cf()._fHasDirtyInnerFormats = !!pCFI->_cf().AreInnerFormatsDirty();

            hr = pts->_pCharFormatCache->CacheData(&pCFI->_cf(), &lIndex);
            if(hr)
            {
                goto Error;
            }

            Assert(lIndex<MAX_FORMAT_INDEX && lIndex>=0);

            _iCF = lIndex;
        }

        // ParaFormat
        if(!pCFI->_fPreparedParaFormat)
        {
            _iPF = pCFI->_ipfSrc;
            pts->_pParaFormatCache->AddRefData(_iPF);
        }
        else
        {
            pCFI->_pf()._fHasDirtyInnerFormats = pCFI->_pf().AreInnerFormatsDirty();

            hr = pts->_pParaFormatCache->CacheData(&pCFI->_pf(), &lIndex);
            if(hr)
            {
                goto Error;
            }

            Assert(lIndex<MAX_FORMAT_INDEX && lIndex>=0);

            _iPF = lIndex;
        }
    }

    // CFancyFormat
    if(pCFI->_pAAExpando)
    {
        hr = pts->_pStyleExpandoCache->CacheData(pCFI->_pAAExpando, &iExpando);
        if(hr)
        {
            goto Error;
        }

        if(pCFI->_pff->_iExpandos != iExpando)
        {
            pCFI->PrepareFancyFormat();
            pCFI->_ff()._iExpandos = iExpando;
            pCFI->UnprepareForDebug();
        }

        pCFI->_pAAExpando->Free();
        pCFI->_pAAExpando = NULL;
    }
    else
    {
    }

    if(!pCFI->_fPreparedFancyFormat)
    {
        _iFF = pCFI->_iffSrc;
        pts->_pFancyFormatCache->AddRefData(_iFF);
    }
    else
    {
        hr = pts->_pFancyFormatCache->CacheData(&pCFI->_ff(), &lIndex);
        if(hr)
        {
            goto Error;
        }

        Assert(lIndex<MAX_FORMAT_INDEX && lIndex>=0);

        _iFF = lIndex;

        if(iExpando >= 0)
        {
            pts->_pStyleExpandoCache->ReleaseData(iExpando);
        }
    }

    Assert(_iCF>=0 && _iPF>=0 && _iFF>=0);

    pCFI->UnprepareForDebug();

    Assert(!pCFI->_pAAExpando);
    Assert(!pCFI->_cstrBgImgUrl);
    Assert(!pCFI->_cstrLiImgUrl);

    return(S_OK);

Error:
    if(_iCF >= 0) pts->_pCharFormatCache->ReleaseData(_iCF);
    if(_iPF >= 0) pts->_pParaFormatCache->ReleaseData(_iPF);
    if(_iFF >= 0) pts->_pFancyFormatCache->ReleaseData(_iFF);
    if(iExpando >= 0) pts->_pStyleExpandoCache->ReleaseData(iExpando);

    pCFI->Cleanup();

    _iCF = _iPF = _iFF = -1;

    RRETURN(hr);
}

//+-----------------------------------------------------------------
//
// Member:      EnsureFormats
//
// Synopsis:    Compute the formats if dirty
//
//------------------------------------------------------------------
void CTreeNode::EnsureFormats()
{
    if(_iCF < 0)
    {
        GetCharFormatHelper();
    }
    if(_iPF < 0)
    {
        GetParaFormatHelper();
    }
    if(_iFF < 0)
    {
        GetFancyFormatHelper();
    }

    // BUGBUG (MohanB) Might unnecessarily cause of multiple walks of the
    // slave tree if the master element is overlapping and CMarkup::EnsureFormats
    // is called on the master's markup. Rare case, not important for IE5.
    if(Element()->HasSlaveMarkupPtr())
    {
        Element()->GetSlaveMarkupPtr()->EnsureFormats();
    }
}

CCharFormat     g_cfStock;
CParaFormat     g_pfStock;
CFancyFormat    g_ffStock;
BOOL            g_fStockFormatsInitialized = FALSE;
void DeinitFormatCache(THREADSTATE* pts)
{
    if(pts->_pParaFormatCache)
    {
        if(pts->_ipfDefault >= 0)
        {
            pts->_pParaFormatCache->ReleaseData(pts->_ipfDefault);
        }
        delete pts->_pParaFormatCache;
    }

    if(pts->_pFancyFormatCache)
    {
        if(pts->_iffDefault >= 0)
        {
            pts->_pFancyFormatCache->ReleaseData(pts->_iffDefault);
        }
        delete pts->_pFancyFormatCache;
    }

    delete pts->_pCharFormatCache;
    delete pts->_pStyleExpandoCache;
}

HRESULT InitFormatCache(THREADSTATE* pts)
{
    CParaFormat pf;
    CFancyFormat ff;
    HRESULT hr = S_OK;

    if(!g_fStockFormatsInitialized)
    {
        g_cfStock._ccvTextColor = RGB(0, 0, 0);
        g_ffStock._ccvBackColor = RGB(0xff, 0xff, 0xff);
        g_pfStock._lFontHeightTwips = 240;
        g_fStockFormatsInitialized = TRUE;
    }

    pts->_pCharFormatCache = new CCharFormatCache;
    if(!pts->_pCharFormatCache)
    {
        goto MemoryError;
    }

    pts->_pParaFormatCache = new CParaFormatCache;
    if(!pts->_pParaFormatCache)
    {
        goto MemoryError;
    }

    pts->_ipfDefault = -1;
    pf.InitDefault();
    pf._fHasDirtyInnerFormats = pf.AreInnerFormatsDirty();
    hr = pts->_pParaFormatCache->CacheData(&pf, &pts->_ipfDefault);
    if(hr)
    {
        goto Cleanup;
    }

    pts->_ppfDefault = &(*pts->_pParaFormatCache)[pts->_ipfDefault];

    pts->_pFancyFormatCache = new CFancyFormatCache;
    if(!pts->_pFancyFormatCache)
    {
        goto MemoryError;
    }

    pts->_iffDefault = -1;
    ff.InitDefault();
    hr = pts->_pFancyFormatCache->CacheData(&ff, &pts->_iffDefault);
    if(hr)
    {
        goto Cleanup;
    }

    pts->_pffDefault = &(*pts->_pFancyFormatCache)[pts->_iffDefault];

    pts->_pStyleExpandoCache = new CStyleExpandoCache;
    if(!pts->_pStyleExpandoCache)
    {
        goto MemoryError;
    }

Cleanup:
    RRETURN(hr);

MemoryError:
    DeinitFormatCache(pts);
    hr = E_OUTOFMEMORY;
    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Function:   EnsureUserStyleSheets
//
//  Synopsis:   Ensure the user stylesheets collection exists if specified by
//              user in option setting, creates it if not..
//
//----------------------------------------------------------------------------
HRESULT EnsureUserStyleSheets(LPTSTR pchUserStylesheet)
{
    // bail out if user SS file is not specified, but "Use My Stylesheet" is checked in options
    if(!*pchUserStylesheet)
    {
        RRETURN(E_FAIL);
    }

    RRETURN(E_FAIL);
}

const CCharFormat* CTreeNode::GetCharFormatHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iCF>=0 ? ::GetCharFormatEx(_iCF) : &g_cfStock);
}

const CParaFormat* CTreeNode::GetParaFormatHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iPF>=0 ? ::GetParaFormatEx(_iPF) : &g_pfStock);
}

const CFancyFormat* CTreeNode::GetFancyFormatHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iFF>=0 ? ::GetFancyFormatEx(_iFF) : &g_ffStock);
}

long CTreeNode::GetCharFormatIndexHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iCF);
}

long CTreeNode::GetParaFormatIndexHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iPF);
}

long CTreeNode::GetFancyFormatIndexHelper()
{
    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || _iFF==-1);

    BYTE ab[sizeof(CFormatInfo)];
    ((CFormatInfo*)ab)->_eExtraValues = ComputeFormatsType_Normal;
    Element()->ComputeFormats((CFormatInfo*)ab, this);
    return (_iFF);
}

//+-----------------------------------------------------
//
//  Member  : GetFontHeightInTwips
//
//  Sysnopsis : helper function that returns the base font
//          height in twips. by default this will return 1
//
//          for now only EMs are wired up to the point that
//          it makes sense to pass through fontsize.
//--------------------------------------------------------
long CTreeNode::GetFontHeightInTwips(CUnitValue* pCuv)
{
    long  lFontHeight = 1;

    if(pCuv->GetUnitType()==CUnitValue::UNIT_EM || pCuv->GetUnitType()==CUnitValue::UNIT_EX)
    {
        const CCharFormat* pCF = GetCharFormat();
        if(pCF)
        {
            lFontHeight = pCF->GetHeightInTwips(Doc());
        }
    }

    return lFontHeight;
}

void CTreeNode::GetRelTopLeft(CElement* pElementFL, CParentInfo* ppi,
                              long* pxRelLeft, long* pyRelTop)
{
    CTreeNode*          pNode = this;
    CDocument*          pDoc = pElementFL->Doc();
    long                lFontHeight;
    const CCharFormat*  pCF;
    const CFancyFormat* pFF;

    Assert(pyRelTop && pxRelLeft);

    *pyRelTop = 0;
    *pxRelLeft = 0;

    while(pNode && pNode->Element()!=pElementFL)
    {
        pCF = pNode->GetCharFormat();

        if(!pCF->_fRelative)
        {
            break;
        }

        lFontHeight = pCF->GetHeightInTwips(pDoc);
        pFF = pNode->GetFancyFormat();

        if(pFF->_fRelative)
        {
            *pyRelTop  += pFF->_cuvTop.YGetPixelValue(ppi, ppi->_sizeParent.cy, lFontHeight);
            *pxRelLeft += pFF->_cuvLeft.XGetPixelValue(ppi, ppi->_sizeParent.cx, lFontHeight);
        }

        pNode = pNode->Parent();
    }
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedbackgroundColor
//
// Synopsis:    Return the background color of the element
//
//-----------------------------------------------------------------------------
CColorValue CTreeNode::GetCascadedbackgroundColor()
{
    return (CColorValue)CTreeNode::GetFancyFormat()->_ccvBackColor;
}

CColorValue CTreeNode::GetCascadedcolor()
{
    return (CColorValue)CTreeNode::GetCharFormat()->_ccvTextColor;
}

CUnitValue CTreeNode::GetCascadedletterSpacing()
{
    return (CUnitValue)CTreeNode::GetCharFormat()->_cuvLetterSpacing;
}

styleTextTransform CTreeNode::GetCascadedtextTransform()
{
    return (styleTextTransform)CTreeNode::GetCharFormat()->_bTextTransform;
}

CUnitValue CTreeNode::GetCascadedpaddingTop()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvPaddingTop;
}

CUnitValue CTreeNode::GetCascadedpaddingRight()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvPaddingRight;
}

CUnitValue CTreeNode::GetCascadedpaddingBottom()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvPaddingBottom;
}

CUnitValue CTreeNode::GetCascadedpaddingLeft()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvPaddingLeft;
}

CColorValue CTreeNode::GetCascadedborderTopColor()
{
    return (CColorValue)CTreeNode::GetFancyFormat()->_ccvBorderColors[BORDER_TOP];
}

CColorValue CTreeNode::GetCascadedborderRightColor()
{
    return (CColorValue)CTreeNode::GetFancyFormat()->_ccvBorderColors[BORDER_RIGHT];
}

CColorValue CTreeNode::GetCascadedborderBottomColor()
{
    return (CColorValue)CTreeNode::GetFancyFormat()->_ccvBorderColors[BORDER_BOTTOM];
}

CColorValue CTreeNode::GetCascadedborderLeftColor()
{
    return (CColorValue)CTreeNode::GetFancyFormat()->_ccvBorderColors[BORDER_LEFT];
}

//+----------------------------------------------------------------------------
//
// Member:      ConvertFmToCSSBorderStyle
//
// Synopsis:    Converts the border style from the internal type to the type
//                  used to set it.
//-----------------------------------------------------------------------------
styleBorderStyle ConvertFmToCSSBorderStyle(BYTE bFmBorderStyle)
{
    switch(bFmBorderStyle)
    {
    case fmBorderStyleDotted:
        return styleBorderStyleDotted;
    case fmBorderStyleDashed:
        return styleBorderStyleDashed;
    case fmBorderStyleDouble:
        return styleBorderStyleDouble;
    case fmBorderStyleSingle:
        return styleBorderStyleSolid;
    case fmBorderStyleEtched:
        return styleBorderStyleGroove;
    case fmBorderStyleBump:
        return styleBorderStyleRidge;
    case fmBorderStyleSunken:
        return styleBorderStyleInset;
    case fmBorderStyleRaised:
        return styleBorderStyleOutset;
    case fmBorderStyleNone:
        return styleBorderStyleNone;
    case 0xff:
        return styleBorderStyleNotSet;
    }

    Assert(FALSE && "Unknown Border Style!");
    return styleBorderStyleNotSet;
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedborderTopStyle
//
// Synopsis:    Return the top border style value for the node
//
//-----------------------------------------------------------------------------
styleBorderStyle CTreeNode::GetCascadedborderTopStyle()
{
    const CFancyFormat* pFF = GetFancyFormat();
    Assert(pFF);
    return ConvertFmToCSSBorderStyle(pFF->_bBorderStyles[BORDER_TOP]);
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedborderRightStyle
//
// Synopsis:    Return the right border style value for the node
//
//-----------------------------------------------------------------------------
styleBorderStyle CTreeNode::GetCascadedborderRightStyle()
{
    const CFancyFormat* pFF = GetFancyFormat();
    Assert(pFF);
    return ConvertFmToCSSBorderStyle(pFF->_bBorderStyles[BORDER_RIGHT]);
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedborderBottomStyle
//
// Synopsis:    Return the bottom border style value for the node
//
//-----------------------------------------------------------------------------
styleBorderStyle CTreeNode::GetCascadedborderBottomStyle()
{
    const CFancyFormat* pFF = GetFancyFormat();
    Assert(pFF);
    return ConvertFmToCSSBorderStyle(pFF->_bBorderStyles[BORDER_BOTTOM]);
}

//+----------------------------------------------------------------------------
//
// Member:      GetCascadedborderLeftStyle
//
// Synopsis:    Return the left border style value for the node
//
//-----------------------------------------------------------------------------
styleBorderStyle CTreeNode::GetCascadedborderLeftStyle()
{
    const CFancyFormat* pFF = GetFancyFormat();
    Assert(pFF);
    return ConvertFmToCSSBorderStyle(pFF->_bBorderStyles[BORDER_LEFT]);
}

void CTreeNode::ReplacePtr(CTreeNode** pbrlhs, CTreeNode* brrhs)
{
    if(pbrlhs)
    {
        if(brrhs)
        {
            brrhs->NodeAddRef();
        }
        if(*pbrlhs)
        {
            (*pbrlhs)->NodeRelease();
        }
        *pbrlhs = brrhs;
    }
}

void CTreeNode::SetPtr(CTreeNode** pbrlhs, CTreeNode* brrhs)
{
    if(pbrlhs)
    {
        if(brrhs)
        {
            brrhs->NodeAddRef();
        }
        *pbrlhs = brrhs;
    }
}

void CTreeNode::ClearPtr(CTreeNode** pbrlhs)
{
    if(pbrlhs && *pbrlhs)
    {
        CTreeNode* pNode = *pbrlhs;
        *pbrlhs = NULL;
        pNode->NodeRelease();
    }
}

void CTreeNode::ReleasePtr(CTreeNode* pNode)
{
    if(pNode)
    {
        pNode->NodeRelease();
    }
}

void CTreeNode::StealPtrSet(CTreeNode** pbrlhs, CTreeNode* brrhs)
{
    SetPtr(pbrlhs, brrhs);

    if(pbrlhs && *pbrlhs)
    {
        (*pbrlhs)->NodeRelease();
    }
}

void CTreeNode::StealPtrReplace(CTreeNode** pbrlhs, CTreeNode* brrhs)
{
    ReplacePtr(pbrlhs, brrhs);

    if(pbrlhs && *pbrlhs)
    {
        (*pbrlhs)->NodeRelease();
    }
}

void CTreeNode::VoidCachedNodeInfo()
{
    THREADSTATE* pts = GetThreadState();

    // Only CharFormat and ParaFormat are in sync.
    Assert((_iCF==-1 && _iPF==-1) || (_iCF>=0  && _iPF>=0));

    if(_iCF != -1)
    {
        (pts->_pCharFormatCache)->ReleaseData(_iCF);
        _iCF = -1;

        (pts->_pParaFormatCache)->ReleaseData(_iPF);
        _iPF = -1;
    }

    if(_iFF != -1)
    {
        (pts->_pFancyFormatCache)->ReleaseData(_iFF);
        _iFF = -1;
    }

    Assert(_iCF==-1 && _iPF==-1 && _iFF==-1);
}

void CTreeNode::VoidCachedInfo()
{
    Assert(Element());
    Element()->_fDefinitelyNoBorders = FALSE;

    VoidCachedNodeInfo();
}

void CTreeNode::VoidFancyFormat()
{
    THREADSTATE* pts = GetThreadState();

    if(_iFF != -1)
    {
        (pts->_pFancyFormatCache)->ReleaseData(_iFF);
        _iFF = -1;
    }
}

CTreeNode::CLock::CLock(CTreeNode* pNode)
{
    Assert(pNode);
    _pNode = pNode;
    pNode->NodeAddRef();
}

CTreeNode::CLock::~CLock()
{
    _pNode->NodeRelease();
}

void* CTreeNode::GetLookasidePtr(int iPtr)
{
    return (HasLookasidePtr(iPtr) ? Doc()->GetLookasidePtr((DWORD*)this+iPtr) : NULL);
}

HRESULT CTreeNode::SetLookasidePtr(int iPtr, void* pvVal)
{
    Assert(Doc());
    Assert(!HasLookasidePtr(iPtr) && "Can't set lookaside ptr when the previous ptr is not cleared");

    HRESULT hr = Doc()->SetLookasidePtr((DWORD*)this+iPtr, pvVal);

    if(hr == S_OK)
    {
        _fHasLookasidePtr |= 1 << iPtr;
    }

    RRETURN(hr);
}

void* CTreeNode::DelLookasidePtr(int iPtr)
{
    Assert(Doc());
    if(HasLookasidePtr(iPtr))
    {
        void* pvVal = Doc()->DelLookasidePtr((DWORD*)this+iPtr);
        _fHasLookasidePtr &= ~(1 << iPtr);
        return(pvVal);
    }

    return NULL;
}