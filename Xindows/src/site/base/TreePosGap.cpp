
#include "stdafx.h"
#include "TreePosGap.h"

CTreePos* CTreePosGap::AdjacentTreePos(TPG_DIRECTION eDir) const
{
    Assert(eDir==TPG_LEFT || eDir==TPG_RIGHT);
    AssertSz(_ptpAttach, "Gap is not positioned");
    BOOL fLeft = (eDir == TPG_LEFT);

    return (fLeft==!!_fLeft) ? _ptpAttach :
        _fLeft?_ptpAttach->NextTreePos():_ptpAttach->PreviousTreePos();
}

CTreeNode* CTreePosGap::Branch() const
{
    AssertSz(_ptpAttach, "Gap is not positioned");
    BOOL fLeft = _fLeft;
    CTreePos* ptp= _ptpAttach;

    if(!_ptpAttach->IsNode()) 
    {
        fLeft = FALSE;
        do
        {
            ptp = ptp->NextTreePos();
        } while(!ptp->IsNode());  
    }

    return (fLeft==ptp->IsEndNode())?ptp->Branch()->Parent():ptp->Branch();
}

BOOL CTreePosGap::IsValid() const
{
    AssertSz(_ptpAttach, "Gap is not positioned");

    BOOL result = !_ptpAttach->IsNode();

    if(!result)
    {
        CTreePos* ptpLeft = _fLeft ? _ptpAttach : _ptpAttach->PreviousTreePos();
        CTreePos* ptpRight = _fLeft ? _ptpAttach->NextTreePos() : _ptpAttach;

        result = !((ptpLeft->IsEndNode() && !ptpLeft->IsEdgeScope())
            || (ptpRight->IsBeginNode() && !ptpRight->IsEdgeScope()));
    }

    return result;
}

void CTreePosGap::SetAttachPreference(TPG_DIRECTION eDir)
{
    _eAttach = eDir;

    if(eDir!=TPG_EITHER && (eDir==TPG_LEFT)!=!!_fLeft && _ptpAttach)
    {
        if(_fLeft)
        {
            MoveTo(_ptpAttach->NextTreePos(), TPG_LEFT);
        }
        else
        {
            MoveTo(_ptpAttach->PreviousTreePos(), TPG_RIGHT);
        }
    }
}

HRESULT CTreePosGap::MoveTo(CTreePos* ptp, TPG_DIRECTION eDir)
{
    Assert(ptp && (eDir==TPG_LEFT || eDir==TPG_RIGHT));
    BOOL fLeft = (eDir != TPG_LEFT); // now from the gap's point of view
    HRESULT hr;

    // adjust for my attach preference
    if(_eAttach!=TPG_EITHER && fLeft!=(_eAttach==TPG_LEFT))
    {
        ptp = fLeft ? ptp->NextTreePos() : ptp->PreviousTreePos();
        fLeft = !fLeft;
    }

    // make sure the requested gap is in the scope of the restricting element
    if(_pElemRestrict)
    {
        if(ptp->SearchBranchForElement(_pElemRestrict, !fLeft) == NULL)
        {
            hr = E_UNEXPECTED;
            goto Cleanup;
        }
    }

    UnPosition();
    _ptpAttach = ptp;
    _fLeft = fLeft;
    hr = S_OK;

Cleanup:
    return hr;
}

HRESULT CTreePosGap::MoveToCp(CMarkup* pMarkup, long cp)
{
    HRESULT     hr;
    long        ich;
    CTreePos*   ptp;

    ptp = pMarkup->TreePosAtCp(cp, &ich);
    Assert(ptp);

    if(ich)
    {
        Assert(ptp->IsText());

        hr = pMarkup->Split(ptp, ich);
        if(hr)
        {
            goto Cleanup;
        }

        hr = MoveTo(ptp, TPG_RIGHT);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else
    {
        hr = MoveTo(ptp, TPG_LEFT);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

/////////////////////////////////////////////////////////
// Method:      PartitionPointers
//
// Synopsis:    Permute any Pointer treepos around this gap so that the
//              ones with left-gravity come before the ones with right-
//              gravity, and the gap is positioned between the two groups.
//              Within each group, the pointers should be stable - i.e. they
//              maintain their relative positions.
//
//              Normal operation treats empty text poses as pointers
//              with right gravity.  However, if fDOMPartition is TRUE,
//              then empty text poses are treated as content.

// BUGBUG: this might be faster if we didn't use gaps internally
void CTreePosGap::PartitionPointers(CMarkup* pMarkup, BOOL fDOMPartition)
{
    TPG_DIRECTION eSaveAttach = (TPG_DIRECTION)_eAttach;
    CTreePosGap tpgMiddle(TPG_RIGHT);

    Assert(!AttachedTreePos()->GetMarkup()->HasUnembeddedPointers());
    Assert(pMarkup);

    // move left to the beginning of the block of pointers
    SetAttachPreference(TPG_LEFT);
    while(AttachedTreePos()->IsPointer()
        || AttachedTreePos()->IsText() && AttachedTreePos()->Cch()==0 
        && (!fDOMPartition||AttachedTreePos()->TextID()==0))
    {
        MoveLeft();
    }

    // now move right, looking for left-gravity pointers
    SetAttachPreference(TPG_RIGHT);
    tpgMiddle.MoveTo(this);
    while(AttachedTreePos()->IsPointer()
        || AttachedTreePos()->IsText() && AttachedTreePos()->Cch()==0 
        && (!fDOMPartition||AttachedTreePos()->TextID()==0))
    {
        if(AttachedTreePos()->IsPointer() && AttachedTreePos()->Gravity()==POINTER_GRAVITY_Left)
        {
            // If tpg is in the same place, no need to move the pointer.
            if(AttachedTreePos() != tpgMiddle.AttachedTreePos())
            {
                CTreePos* ptpToMove = AttachedTreePos();

                // Don't move us along with the pointer.
                SetAttachPreference(TPG_LEFT);

                Verify(!pMarkup->Move(ptpToMove, &tpgMiddle));

                // So now we're attached to the next pointer.
                SetAttachPreference(TPG_RIGHT);
            }
            else
            {
                // Nothing changed, so move on ahead.
                MoveRight();
                tpgMiddle.MoveTo(this);
            }
        }
        else
        {
            MoveRight();
        }
    }

    // Position myself in the middle
    MoveTo(&tpgMiddle);

    // restore my state
    SetAttachPreference(eSaveAttach);
}

/////////////////////////////////////////////////////////
// Method:      CleanCling
//
// Synopsis:    Assuming that the current gap is in 
//              the middle of a group of pointers that
//              has been positioned
void CTreePosGap::CleanCling(CMarkup* pMarkup, TPG_DIRECTION eDir, BOOL fDOMPartition)
{
    Assert(eDir==TPG_LEFT || eDir==TPG_RIGHT);
    BOOL fLeft = (eDir == TPG_LEFT);
    CTreePos* ptp = AdjacentTreePos(eDir);
    CTreePos* ptpNext;

    while(ptp->IsPointer() || ptp->IsText() && ptp->Cch()==0 && (!fDOMPartition||ptp->TextID()==0))
    {
        ptpNext = fLeft ? ptp->PreviousTreePos() : ptp->NextTreePos();

        if(ptp->IsPointer() && ptp->Cling() || ptp->IsText() && ptp->Cch()==0 && (!fDOMPartition||ptp->TextID()==0))
        {
            if(AttachedTreePos() == ptp)
            {
                Assert(AttachDirection() == eDir);
                Move(eDir);
            }

            Verify(!pMarkup->Remove(ptp));
        }
        ptp = ptpNext;
    }
}

HRESULT CTreePosGap::MoveImpl(BOOL fLeft, DWORD dwMoveFlags, CTreePos** pptpEdgeCrossed)
{
    HRESULT hr;
    BOOL fFirstTime = TRUE;
    const BOOL fAttachToCurr = (fLeft == !_fLeft);
    CTreePos* ptpCurr = fAttachToCurr ? _ptpAttach :
        _fLeft?_ptpAttach->NextTreePos():_ptpAttach->PreviousTreePos();
    CTreePos* ptpAdv;
    CTreePos* ptpAttach = NULL;

    Assert(_eAttach!=TPG_EITHER ||
        0==(dwMoveFlags&(TPG_SKIPPOINTERS|TPG_FINDTEXT|TPG_FINDEDGESCOPE)));

    if(pptpEdgeCrossed)
    {
        *pptpEdgeCrossed = NULL;
    }

    for(ptpAdv=fLeft?ptpCurr->PreviousTreePos():ptpCurr->NextTreePos();
        ptpAdv;
        ptpCurr=ptpAdv,ptpAdv=fLeft?ptpCurr->PreviousTreePos():ptpCurr->NextTreePos())
    {
        ptpAttach = fAttachToCurr ? ptpCurr : ptpAdv;

        if(!fFirstTime)
        {
            // if we've left the scope of the restricting element, fail
            if(_pElemRestrict)
            {
                if(ptpCurr->IsNode() && ptpCurr->IsEdgeScope() &&
                    ptpCurr->Branch()->Element()==_pElemRestrict)
                {
                    ptpAdv = NULL;
                    break;
                }
            }

            // track any begin-scope or end-scope nodes that we cross over
            if(pptpEdgeCrossed && ptpCurr->IsNode() && ptpCurr->IsEdgeScope())
            {
                AssertSz(*pptpEdgeCrossed == NULL, "Crossed multiple begin or endscope TreePos");
                *pptpEdgeCrossed = ptpCurr;
            }
        }
        else
        {
            // if it's not OK to stay still, force the first advance
            fFirstTime = FALSE;
            if(!(dwMoveFlags & TPG_OKNOTTOMOVE))
            {
                continue;
            }
        }

        // if caller wants a valid gap, advance if we're not in one
        if((dwMoveFlags & TPG_VALIDGAP))
        {
            CTreePos* ptpLeft = fLeft ? ptpAdv : ptpCurr;
            CTreePos* ptpRight = fLeft ? ptpCurr : ptpAdv;
            if(!CTreePos::IsLegalPosition(ptpLeft, ptpRight))
            {
                continue;
            }
        }

        // if caller wants to skip pointers, advance if we're at one
        if(ptpAttach->IsPointer() && (dwMoveFlags&TPG_SKIPPOINTERS))
        {
            continue;
        }

        // if caller wants a text position, advance if we're not at one
        if(!ptpAttach->IsText() && (dwMoveFlags&TPG_FINDTEXT))
        {
            continue;
        }

        // if caller wants an edge-scope position, advance if we're not at one
        if((dwMoveFlags&TPG_FINDEDGESCOPE) &&
            !(ptpAttach->IsNode() && ptpAttach->IsEdgeScope()))
        {
            continue;
        }

        // if we've survived the gauntlet above, we're in a good gap
        break;
    }

    // if we found a good gap, move there
    if(ptpAdv)
    {
        UnPosition();
        _ptpAttach = ptpAttach;
        hr = S_OK;
    }
    else
    {
        hr = E_UNEXPECTED;
    }

    return hr;
}



///////////////////////////////////////////////////////////////////
//
//  Class:      CChildIterator
//
///////////////////////////////////////////////////////////////////
CChildIterator::CChildIterator(
       CElement*    pElementParent,
       CElement*    pElementChild,
       DWORD        dwFlags,
       ELEMENT_TAG* pStopTags,
       long         cStopTags,
       ELEMENT_TAG* pChildTags,
       long         cChildTags)
{
    AssertSz(!(dwFlags&~CHILDITERATOR_PUBLICMASK), "Invalid flags passed to CChildIterator constructor");
    _dwFlags = dwFlags;

    Assert(pElementParent);

    if(pElementChild)
    {
        _pNodeChild = pElementChild->GetFirstBranch();
        Assert(_pNodeChild);

        _pNodeParent = _pNodeChild->SearchBranchToRootForScope(pElementParent);
        Assert(_pNodeParent);
    }
    else
    {
        SetBeforeBeginBit();
        _pNodeChild = NULL;
        _pNodeParent = pElementParent->GetFirstBranch();
    }

    AssertSz(_pNodeParent, "CChildIterator used with parent not in tree -- not a tree bug");

    if(UseTags())
    {
        _pStopTags = pStopTags;
        _cStopTags = cStopTags;

        _pChildTags = pChildTags;
        _cChildTags = cChildTags;
    }
}

CTreeNode* CChildIterator::NextChild()
{
    // If we are after the end of the child list,
    // we can't go any further
    if(IsAfterEnd())
    {
        return NULL;
    }

    CTreePos*   ptpCurr;
    CTreeNode*  pNodeParentCurr = _pNodeParent;
    BOOL        fFirst = TRUE;

    // Decide where to start.  
    // * If we already have a child either start after that 
    //   child or just inside of it, dependent on the virtual 
    //   IsRecursionStopChildNode call.
    // * If we are before begin, start just inside of the
    //   parent node.
    if(_pNodeChild)
    {
        ptpCurr = IsRecursionStopChildNode(_pNodeChild) ?
            _pNodeChild->GetEndPos() : _pNodeChild->GetBeginPos();
    }
    else
    {
        Assert(IsBeforeBeginBit());
        ClearBeforeBeginBit();

        ptpCurr = _pNodeParent->GetBeginPos();
    }

    Assert(ptpCurr && !ptpCurr->IsUninit());

    // Do this loop while we have a parent that we are interested in
    while(pNodeParentCurr)
    {
        CTreePos* ptpParentEnd = pNodeParentCurr->GetEndPos();

        fFirst = FALSE;

        // Since this is a good parent, set the member variable
        _pNodeParent = pNodeParentCurr;

        // Loop through everything under this parent.
        for(ptpCurr=ptpCurr->NextTreePos();
            ptpCurr!=ptpParentEnd;
            ptpCurr=ptpCurr->NextTreePos())
        {
            // If this is a begin node, investigate further
            if(ptpCurr->IsBeginNode())
            {
                CTreeNode* pNodeChild = ptpCurr->Branch();

                // return this node if we have an edge scope and
                // the child is interesting.
                if(ptpCurr->IsEdgeScope() && IsInterestingChildNode(pNodeChild))
                {
                    _pNodeChild = pNodeChild;
                    return _pNodeChild;
                }

                // Decide if we want to skip over this node's
                // subtree
                if(IsRecursionStopChildNode(pNodeChild))
                {
                    ptpCurr = pNodeChild->GetEndPos();
                }
            }
        }

        // Move on to the next parent node
        pNodeParentCurr = pNodeParentCurr->NextBranch();
        if(pNodeParentCurr)
        {
            ptpCurr = pNodeParentCurr->GetBeginPos();
        }
    }

    // If we got here, we ran out of parent nodes, so we must be
    // after the end.
    SetAfterEndBit();
    _pNodeChild = NULL;
    return NULL;
}

CTreeNode* CChildIterator::PreviousChild()
{
    // If we are after the end of the child list,
    // we can't go any further
    if(IsBeforeBegin())
    {
        return NULL;
    }

    CTreePos* ptpCurr;
    CTreeNode* pNodeParentCurr = _pNodeParent;

    // Decide where to start.  
    // * If we already have a child either start before that 
    //   child or just inside of it, dependent on the virtual 
    //   IsRecursionStopChildNode call.
    // * If we are after end, start just inside of the
    //   parent node.
    if(_pNodeChild)
    {
        ptpCurr = IsRecursionStopChildNode(_pNodeChild) ?
            _pNodeChild->GetBeginPos() : _pNodeChild->GetEndPos();
    }
    else
    {
        Assert(IsAfterEndBit());
        ClearAfterEndBit();

        ptpCurr = _pNodeParent->GetEndPos();
    }

    while(pNodeParentCurr)
    {
        CTreePos* ptpStop = pNodeParentCurr->GetBeginPos();

        // Since this is a good parent, set the member variable
        _pNodeParent = pNodeParentCurr;

        // Loop through everything under this parent.
        for(ptpCurr=ptpCurr->PreviousTreePos();
            ptpCurr!=ptpStop;
            ptpCurr=ptpCurr->PreviousTreePos())
        {
            Assert(! ptpCurr->IsNode() || ptpCurr->Branch()->SearchBranchToRootForNode(pNodeParentCurr));

            // If this is a begin node, investigate further
            if(ptpCurr->IsEndNode())
            {
                CTreeNode* pNodeChild = ptpCurr->Branch();
                CTreePos* ptpBegin = pNodeChild->GetBeginPos();

                // return this node if we have an edge scope and
                // the child is interesting.
                if(ptpBegin->IsEdgeScope() && IsInterestingChildNode(pNodeChild))
                {
                    _pNodeChild = pNodeChild;
                    return _pNodeChild;
                }

                // Decide if we want to skip over this nodes
                // subtree
                if(IsRecursionStopChildNode(pNodeChild))
                {
                    ptpCurr = ptpBegin;
                }
            }
        }

        // Find the previous node in the context chain
        if(pNodeParentCurr == pNodeParentCurr->Element()->GetFirstBranch())
        {
            pNodeParentCurr = NULL;
        }
        else
        {
            CTreeNode* pNodeTemp = pNodeParentCurr->Element()->GetFirstBranch();
            while(pNodeTemp->NextBranch() != pNodeParentCurr)
            {
                pNodeTemp = pNodeTemp->NextBranch();
            }
            Assert(pNodeTemp);
            pNodeParentCurr = pNodeTemp;
        }

        if(pNodeParentCurr)
        {
            ptpCurr = pNodeParentCurr->GetEndPos();
        }
    }

    // If we got here, we ran out of parent nodes, so we must be
    // before begin.
    SetBeforeBeginBit();
    _pNodeChild = NULL;
    return NULL;
}

void CChildIterator::SetChild(CElement* pElementChild)
{
    Assert(pElementChild && pElementChild->GetFirstBranch());

    _pNodeChild = pElementChild->GetFirstBranch();
    _pNodeParent = _pNodeChild->SearchBranchToRootForScope(_pNodeParent->Element());

    Assert(_pNodeParent);
}

void CChildIterator::SetBeforeBegin( )
{
    _pNodeChild = NULL;
    _pNodeParent = _pNodeParent->Element()->GetFirstBranch();

    Assert(_pNodeParent);

    SetBeforeBeginBit();
    ClearAfterEndBit();
}

void CChildIterator::SetAfterEnd( )
{
    _pNodeChild = NULL;
    _pNodeParent = _pNodeParent->Element()->GetFirstBranch();

    while(_pNodeParent->NextBranch())
    {
        _pNodeParent = _pNodeParent->NextBranch();
    }

    Assert(_pNodeParent);

    ClearBeforeBeginBit();
    SetAfterEndBit();
}

BOOL CChildIterator::IsInterestingNode(ELEMENT_TAG* pary, long c, CTreeNode* pNode)
{
    if(UseLayout())
    {
        return pNode->NeedsLayout();
    }

    if(UseTags() && c>0)
    {
        long i;

        for(i=0; i<c; i++,pary++)
        {
            if(*pary == pNode->Tag())
            {
                return TRUE;
            }
        }

        return FALSE;
    }

    return TRUE;
}



void EnsureTotalOrder(CTreePosGap*& ptpgStart, CTreePosGap*& ptpgFinish)
{
    CTreePos* ptpStart = ptpgStart->AdjacentTreePos(TPG_RIGHT);
    CTreePos* ptpFinish = ptpgFinish->AdjacentTreePos(TPG_RIGHT);

    Assert(ptpStart->GetCp() <= ptpFinish->GetCp());

    if(ptpStart == ptpFinish)
    {
        return;
    }

    // Move the finish as far to the right as possible without going over any content,
    // looking for the start.  If we find the start, then they are not ordered properly.
    while(ptpFinish->IsPointer() || ptpFinish->IsText() && ptpFinish->Cch()==0)
    {
        ptpFinish = ptpFinish->NextTreePos();

        if(ptpFinish == ptpStart)
        {
            CTreePosGap* ptpgTemp = ptpgStart;
            ptpgStart = ptpgFinish;
            ptpgFinish = ptpgTemp;

            return;
        }
    }
}