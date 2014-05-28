
#include "stdafx.h"
#include "TreePos.h"

BOOL CTreePos::IsBeginElementScope(CElement* pElem)
{
    return IsBeginElementScope()&&Branch()->Element()==pElem;
}

BOOL CTreePos::IsEndElementScope(CElement* pElem)
{
    return IsEndElementScope()&&Branch()->Element()==pElem;
}

CMarkup* CTreePos::GetMarkup()
{
    CTreePos *ptp=this, *ptpParent=Parent();

    while(ptpParent)
    {
        ptp = ptpParent;
        ptpParent = ptp->Parent();
    }

    AssertSz(!ptp->IsLeftChild(), "GetList called when not in a CMarkup");
    return ptp->IsLeftChild()?NULL:CONTAINING_RECORD(ptp, CMarkup, _tpRoot);
}

// Returns (not logically)
//
//   -1: this <  that
//    0: this == that
//   +1: this >  that
int CTreePos::InternalCompare(CTreePos* ptpThat)
{
    Assert(GetMarkup() && ptpThat->GetMarkup()==GetMarkup());

    if(this == ptpThat)
    {
        return 0;
    }

    static long cSplayThis = 0;

    if(cSplayThis++ & 1)
    {
        CTreePos* ptpThis = this;

        ptpThis->Splay();

        for(; ;)
        {
            CTreePos* ptpChild = ptpThat;

            ptpThat = ptpThat->Parent();

            if(ptpThat == ptpThis)
            {
                return ptpChild->IsLeftChild()?+1:-1;
            }
        }
    }
    else
    {
        CTreePos* ptpThis = this;

        ptpThat->Splay();

        for(; ;)
        {
            CTreePos* ptpChild = ptpThis;

            ptpThis = ptpThis->Parent();

            if(ptpThis == ptpThat)
            {
                return ptpChild->IsLeftChild()?-1:+1;
            }
        }
    }
}

CTreeNode* CTreePos::Branch() const
{
    AssertSz(!IsDataPos()&&IsNode(), "CTreePos::Branch called on non-node pos");
    return (CONTAINING_RECORD((this+!!TestFlag(NodeBeg)), CTreeNode, _tpEnd));
}

CTreeNode* CTreePos::GetBranch() const
{
    CTreePos* ptp = const_cast<CTreePos*>(this);

    if(ptp->IsNode())
    {
        return ptp->Branch();
    }

    while(!ptp->IsNode())
    {
        ptp = ptp->PreviousTreePos();
    }

    Assert(ptp && (ptp->IsBeginNode()||ptp->IsEndNode()));

    return ptp->IsBeginNode()?ptp->Branch():ptp->Branch()->Parent();
}

CTreeNode* CTreePos::GetInterNode() const
{
    CTreePos* ptp = const_cast<CTreePos*>(this);

    while(!ptp->IsNode())
    {
        ptp = ptp->PreviousTreePos();
    }

    Assert(ptp && (ptp->IsBeginNode()||ptp->IsEndNode()));

    return ptp->IsBeginNode()?ptp->Branch():ptp->Branch()->Parent();
}

long CTreePos::SourceIndex()
{
    CTreePos* ptp;
    long cSourceIndex = GetElemLeft();
    BOOL fLeftChild = IsLeftChild();
    long cDepth = -1;
    CMarkup* pMarkup;
    CTreePos* pRoot = NULL;

    for(ptp=Parent(); ptp; ptp=ptp->Parent())
    {
        if(!fLeftChild)
        {
            cSourceIndex += ptp->GetElemLeft() + (ptp->IsBeginElementScope()?1:0);
        }
        fLeftChild = ptp->IsLeftChild();

        ++cDepth;
        pRoot = ptp;
    }

    pMarkup = CONTAINING_RECORD(pRoot, CMarkup, _tpRoot);

    if(pMarkup->ShouldSplay(cDepth))
    {
        Splay();
    }

    return cSourceIndex;
}

long CTreePos::GetCp()
{
    CTreePos* ptp;
    long cch = _cchLeft;
    BOOL fLeftChild = IsLeftChild();
    long cDepth = -1;
    CMarkup* pMarkup;
    CTreePos* pRoot = NULL;

    for(ptp=Parent(); ptp; ptp = ptp->Parent())
    {
        if(!fLeftChild)
        {
            cch += ptp->_cchLeft + ptp->GetCch();
        }

        fLeftChild = ptp->IsLeftChild();

        ++cDepth;
        pRoot = ptp;
    }

    pMarkup = CONTAINING_RECORD(pRoot, CMarkup, _tpRoot);


    if(pMarkup->ShouldSplay(cDepth))
    {
        Splay();
    }

    return cch;
}

int CTreePos::Gravity() const
{
    Assert(IsPointer());
    return DataThis()->p._dwPointerAndGravityAndCling&0x1;
}

void CTreePos::SetGravity ( BOOL fRight )
{
    Assert(IsPointer());

    DataThis()->p._dwPointerAndGravityAndCling = (DataThis()->p._dwPointerAndGravityAndCling&~1)|(!!fRight);
}

int CTreePos::Cling() const
{
    Assert(IsPointer());
    return !!(DataThis()->p._dwPointerAndGravityAndCling&0x2);
}

void CTreePos::SetCling(BOOL fStick)
{
    Assert(IsPointer());

    DataThis()->p._dwPointerAndGravityAndCling = (DataThis()->p._dwPointerAndGravityAndCling&~2)|(fStick?2:0);
}

void CTreePos::SetScopeFlags(BOOL fEdge)
{
    long cElemDelta = IsBeginElementScope()? -1 : 0;

    SetFlags((GetFlags()&~TPF_EDGE)|BOOLFLAG(fEdge, TPF_EDGE));

    if(IsBeginElementScope())
    {
        cElemDelta += 1;
    }

    if(cElemDelta != 0)
    {
        CTreePos* ptp;
        BOOL fLeftChild = IsLeftChild();

        for(ptp=Parent(); ptp; ptp=ptp->Parent())
        {
            if(fLeftChild)
            {
                ptp->AdjElemLeft(cElemDelta);
            }
            fLeftChild = ptp->IsLeftChild();
        }
    }
}

void CTreePos::ChangeCch(long cchDelta)
{
    BOOL fLeftChild = IsLeftChild();
    CTreePos* ptp;

    Assert(IsText());

    DataThis()->t._cch += cchDelta;

    for(ptp=Parent(); ptp; ptp=ptp->Parent())
    {
        if(fLeftChild)
        {
            ptp->_cchLeft += cchDelta;
        }

        fLeftChild = ptp->IsLeftChild();
    }
}

CTreePos* CTreePos::NextTreePos()
{
    CTreePos* ptp;

    if(!FirstChild())
    {
        goto Up;
    }

    if(!FirstChild()->IsLeftChild())
    {
        ptp = FirstChild();
        goto Right;
    }

    if(FirstChild()->IsLastChild())
    {
        goto Up;
    }

    ptp = FirstChild()->Next();

Right:
    // Leftmost descendent
    while(ptp->FirstChild() && ptp->FirstChild()->IsLeftChild())
    {
        ptp = ptp->FirstChild();
    }
    return ptp;

Up:
    for(ptp=this; !ptp->IsLeftChild(); ptp=ptp->Next());

    ptp = ptp->Parent();

    if(ptp->Next() == NULL)
    {
        return NULL;
    }

    return ptp;
}

CTreePos* CTreePos::PreviousTreePos()
{
    CTreePos* ptp = this;

    // Do we have a left child?
    if(ptp->FirstChild() && ptp->FirstChild()->IsLeftChild())
    {
        // Rightmost descendent for ptp->LeftChild()
Loop:
        ptp = ptp->FirstChild();

        while(ptp->FirstChild())
        {
            if(!ptp->FirstChild()->IsLeftChild())
            {
                goto Loop;
            }

            if(ptp->FirstChild()->IsLastChild())
            {
                return ptp;
            }

            ptp = ptp->FirstChild()->Next();
        }

        return ptp;
    }

    // No left child
    while(ptp->IsLeftChild())
    {
        // ptp = ptp->Parent() -- inlined here
        if(ptp->IsLastChild())
        {
            ptp = ptp->Next();
        }
        else
        {
            ptp = ptp->Next()->Next();
        }

        // Root pos (marked as right child) protects us from NULL
    }

    // ptp = ptp->Parent() (we know ptp is a right (and last) child)
    ptp = ptp->Next();

    return ptp;
}

CTreePos* CTreePos::NextNonPtrTreePos()
{
    CTreePos* ptp = this;

    do
    {
        ptp = ptp->NextTreePos();
    }
    while(ptp && ptp->IsPointer());

    return ptp;
}

CTreePos* CTreePos::PreviousNonPtrTreePos()
{
    CTreePos* ptp = this;

    do
    {
        ptp = ptp->PreviousTreePos();
    }
    while(ptp && ptp->IsPointer());

    return ptp;
}

BOOL CTreePos::IsLegalPosition(CTreePos* ptpLeft, CTreePos* ptpRight)
{
    // use the marks to determine if content is allowed between the inputs
    return (!ptpLeft->IsNode() || !ptpRight->IsNode() ||
        !((ptpLeft->IsEndNode() && !ptpLeft->IsEdgeScope()) ||
        (ptpRight->IsBeginNode() && !ptpRight->IsEdgeScope())));
}

void CTreePos::InitSublist()
{
    SetFirstChild(NULL);
    SetNext(NULL);
    MarkLast();
    MarkRight(); // this distinguishes a sublist from a splay tree root
    ClearCounts();
}

CTreePos* CTreePos::LeftmostDescendant() const
{
    const CTreePos* ptp = this;
    const CTreePos* ptpLeft = FirstChild();

    while(ptpLeft && ptpLeft->IsLeftChild())
    {
        ptp = ptpLeft;
        ptpLeft = ptp->FirstChild();
    }

    return const_cast<CTreePos*>(ptp);
}

CTreePos* CTreePos::RightmostDescendant() const
{
    CTreePos* ptp = const_cast<CTreePos*>(this);
    CTreePos* ptpRight = RightChild();

    while(ptpRight)
    {
        ptp = ptpRight;
        ptpRight = ptp->RightChild();
    }

    return ptp;
}

void CTreePos::GetChildren(CTreePos** ppLeft, CTreePos** ppRight) const
{
    if(FirstChild())
    {
        if(FirstChild()->IsLeftChild())
        {
            *ppLeft = FirstChild();
            *ppRight = (FirstChild()->IsLastChild()) ? NULL : FirstChild()->Next();
        }
        else
        {
            *ppLeft = NULL;
            *ppRight = FirstChild();
        }
    }
    else
    {
        *ppLeft = *ppRight = NULL;
    }
}

HRESULT CTreePos::Remove()
{
    Assert(!HasNonzeroCounts(TP_DIRECT)); // otherwise we need to adjust counts to the root
    CTreePos *pLeft, *pRight, *pParent=Parent();

    GetChildren(&pLeft, &pRight);

    if(pLeft == NULL)
    {
        pParent->ReplaceOrRemoveChild(this, pRight);
    }
    else
    {
        while(pRight)
        {
            pRight->RotateUp(this, pParent);
            pRight = RightChild();
        }
        pParent->ReplaceChild(this, pLeft);
    }

    return S_OK;
}

void CTreePos::Splay()
{
    CTreePos *p=Parent(), *g=p->Parent(), *gg;      // parent, grandparent, great-grandparent

    for(; g; p=Parent(),g=p->Parent())
    {
        gg = g->Parent();

        if(gg)
        {
            if (IsLeftChild() == p->IsLeftChild())  // zig-zig
            {
                p->RotateUp(g, gg);
                RotateUp(p, gg);
            }
            else                                    // zig-zag
            {
                RotateUp(p, g);
                RotateUp(g, gg);
            }
        }
        else                                        // zig
        {
            RotateUp(p, g);
        }
    }
}

void CTreePos::RotateUp(CTreePos* p, CTreePos* g)
{
    CTreePos *ptp1, *ptp2, *ptp3;

    if(IsLeftChild()) // rotate right
    {
        GetChildren(&ptp1, &ptp2);
        ptp3 = p->RightChild();
        g->ReplaceChild(p, this);

        // recreate my family
        if(ptp1)
        {
            ptp1->MarkFirst();
            ptp1->SetNext(p);
        }
        else
        {
            SetFirstChild(p);
        }

        // recreate p's family
        if(ptp2)
        {
            p->SetFirstChild(ptp2);
            ptp2->MarkLeft();
            if(ptp3)
            {
                ptp2->MarkFirst();
                ptp2->SetNext(ptp3);
            }
            else
            {
                ptp2->MarkLast();
                ptp2->SetNext(p);
            }
        }
        else
        {
            p->SetFirstChild(ptp3);
        }
        p->MarkRight();
        p->MarkLast();
        p->SetNext(this);

        // adjust cumulative counts
        p->DecreaseCounts(this, TP_BOTH);
    }
    else // rotate left
    {
        ptp1 = p->LeftChild();
        GetChildren(&ptp2, &ptp3);
        g->ReplaceChild(p, this);

        // recreate my family
        SetFirstChild(p);
        p->MarkLeft();
        if(ptp3)
        {
            p->MarkFirst();
            p->SetNext(ptp3);
        }
        else
        {
            p->MarkLast();
            p->SetNext(this);
        }

        // recreate p's family
        if(ptp1)
        {
            if(ptp2)
            {
                ptp1->MarkFirst(); // BUGBUG not needed?
                ptp1->SetNext(ptp2);
            }
            else
            {
                ptp1->MarkLast();
                ptp1->SetNext(p);
            }
        }
        else
        {
            p->SetFirstChild(ptp2);
        }
        if(ptp2)
        {
            ptp2->MarkRight();
            ptp2->MarkLast();
            ptp2->SetNext(p);
        }

        // adjust cumulative counts
        IncreaseCounts(p, TP_BOTH);
    }
}

void CTreePos::ReplaceChild(CTreePos* pOld, CTreePos* pNew)
{
    pNew->MarkLeft(pOld->IsLeftChild());
    pNew->MarkLast(pOld->IsLastChild());
    pNew->SetNext(pOld->Next());

    if(FirstChild() == pOld)
    {
        SetFirstChild(pNew);
    }
    else
    {
        FirstChild()->SetNext(pNew);
    }
}

void CTreePos::RemoveChild(CTreePos* pOld)
{
    if(FirstChild() == pOld)
    {
        SetFirstChild(pOld->IsLastChild()?NULL:pOld->Next());
    }
    else
    {
        FirstChild()->MarkLast();
        FirstChild()->SetNext(this);
    }
}

CTreeNode* CTreePos::SearchBranchForElement(CElement* pElement, BOOL fLeft)
{
    Assert(pElement);
    CTreePos *ptpLeft, *ptpRight;
    CTreeNode *pNode;

    // start at the requested gap
    if(fLeft)
    {
        ptpLeft = PreviousTreePos(), ptpRight = this;
    }
    else
    {
        ptpLeft = this,  ptpRight = NextTreePos();
    }

    // move right to a valid gap (invalid gaps may give too high a scope)
    while(!IsLegalPosition(ptpLeft, ptpRight))
    {
        // but stop if we stumble across the element we're looking for
        if(ptpRight->IsNode() && ptpRight->Branch()->Element()==pElement)
        {
            return ptpRight->Branch();
        }

        ptpLeft = ptpRight;
        ptpRight = ptpLeft->NextTreePos();
    }

    // now search for a nearby node
    if(ptpLeft->IsNode())
    {
        pNode = ptpLeft->IsBeginNode() ? ptpLeft->Branch() : ptpLeft->Branch()->Parent();
    }
    else
    {
        while(!ptpRight->IsNode())
        {
            ptpRight = ptpRight->NextTreePos();
        }
        pNode = ptpRight->IsBeginNode() ? ptpRight->Branch()->Parent() : ptpRight->Branch();
    }

    // finally, look for the element along this node's branch
    for(; pNode; pNode=pNode->Parent())
    {
        if(pNode->Element() == pElement)
        {
            return pNode;
        }
    }

    return NULL;
}

BOOL CTreePos::HasNonzeroCounts(unsigned fFlags)
{
    BOOL result = FALSE;

    if(fFlags & TP_DIRECT)
    {
        Assert(!IsUninit());
        result = result || !IsPointer();
    }

    if(fFlags & TP_LEFT)
    {
        result = result || GetElemLeft()>0 || _cchLeft>0;
    }

    return result;
}

// Retunrs TRUE if this and ptpRight are separated only by
// pointers or empty text positions.  ptpRight must already
// be to the right of this.
BOOL CTreePos::LogicallyEqual(CTreePos* ptpRight)
{
    Assert(InternalCompare(ptpRight) <= 0);

    for(CTreePos* ptp=this; ptp->IsPointer()||(ptp->IsText()&&ptp->Cch()==0); ptp=ptp->NextTreePos())
    {
        if(ptp == ptpRight)
        {
            return TRUE;
        }
    }

    return FALSE;
}