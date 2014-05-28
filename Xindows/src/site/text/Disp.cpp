
#include "stdafx.h"
#include "Disp.h"

#include "lsrender.h"
#include "rclclptr.h"

void CLed::SetNoMatch()
{
    _cpMatchNew = _cpMatchOld = _iliMatchNew = _iliMatchOld =
        _yMatchNew = _yMatchOld = MAXLONG;
}



//+----------------------------------------------------------------------------
//
// Function:    CRelDispNodeCache::GetOwner
//
// Synopsis:    Get the owning element of the given disp node.
//
//-----------------------------------------------------------------------------
void CRelDispNodeCache::GetOwner(CDispNode* pDispNode, void** ppv)
{
    CRelDispNode* prdn;
    long lEntry;

    Assert(pDispNode);
    Assert(pDispNode->GetDispClient() == this);
    Assert(ppv);
    Assert(Size());

    // we could be passed in a dispNode corresponding to the
    // content node of a container, in which case return the
    // element owner of the container.
    if(!pDispNode->IsOwned())
    {
        CDispNode* pDispNodeParent = pDispNode->GetParentNode();

        Assert(pDispNodeParent->GetDispClient() == this);

        lEntry = (LONG)(LONG_PTR)pDispNodeParent->GetExtraCookie();
    }
    else
    {
        lEntry = (LONG)(LONG_PTR)pDispNode->GetExtraCookie();
    }

    Assert(lEntry>=0 && lEntry<Size());

    prdn = (*this)[lEntry];

    Assert(prdn->_pDispNode==pDispNode ||
        prdn->_pDispNode==pDispNode->GetParentNode());

    *ppv = prdn->GetElement();
}

long GetRelCacheEntry(CDispNode* pDispNode)
{
    long lEntry;

    // Find the index of the the cache entry that corresponds to the given node.
    // Extra cookie on the disp node stores the index into the cache.
    //
    // For container nodes, their flow node is passed in as a display node
    // to be rendered, so we need to get the parent's cookie instead.
    if(!pDispNode->IsOwned())
    {
        CDispNode* pDispNodeParent = pDispNode->GetParentNode();

        lEntry = (LONG)(LONG_PTR)pDispNodeParent->GetExtraCookie();
    }
    else
    {
        lEntry = (LONG)(LONG_PTR)pDispNode->GetExtraCookie();
    }

    return lEntry;
}

BOOL CRelDispNodeCache::HitTestContent(const POINT* pptHit, CDispNode* pDispNode, void* pClientData)
{
    Assert(pptHit);
    Assert(pDispNode);
    Assert(pClientData);

    CFlowLayout*    pFL = _pdp->GetFlowLayout();
    CElement*       pElementDispNode;
    CRelDispNode*   prdn;
    CHitTestInfo*   phti        = (CHitTestInfo*)pClientData;
    BOOL            fBrowseMode = !pFL->ElementContent()->IsEditable();
    long            lEntry      = GetRelCacheEntry(pDispNode);
    long            lSize       = Size();
    long            ili;
    long            iliLast;
    long            yli;
    long            cp;
    CPoint          pt;

    Assert(lSize && lEntry>=0 && lEntry<lSize);

    prdn    = (*this)[lEntry];
    cp      = prdn->GetElement()->GetFirstCp() - 1;
    ili     = prdn->_ili;
    iliLast = ili + prdn->_cLines;
    yli     = prdn->_yli;

    // convert the point relative to the layout parent
    pt.x = pptHit->x + prdn->_ptOffset.x;
    pt.y = pptHit->y + prdn->_ptOffset.y + prdn->_yli;

    pElementDispNode = prdn->GetElement();

    // HitTest lines that are owned by the given relative node.
    while(ili < iliLast)
    {
        CLine* pli = _pdp->Elem(ili);

        // Hit test the current line here.
        if(!pli->_fHidden && (!pli->_fDummyLine || fBrowseMode))
        {
            // if the point is in the vertical bounds of the line
            if(pt.y>=yli+pli->GetYTop() && pt.y<yli+pli->GetYLineBottom())
            {
                // check if the point lies in the horzontal bounds of the
                // line
                if(pt.x>=pli->_xLeftMargin &&
                    pt.x<(pli->_fForceNewLine
                    ? pli->_xLeftMargin+pli->_xLineWidth
                    : pli->GetTextRight(ili==_pdp->LineCount()-1)))
                {
                    break;
                }
            }
        }

        if(pli->_fForceNewLine)
        {
            yli += pli->_yHeight;
        }
        cp += pli->_cch;
        ili++;
    }

    // If a line is hit, find the branch hit.
    if(ili < iliLast)
    {
        HITTESTRESULTS* presultsHitTest = phti->_phtr;
        HTC             htc = HTC_YES;
        BOOL            fPseudoHit;

        // If the line hit belongs to a nested relative line then it is
        // a pseudo hit.
        fPseudoHit = FALSE;

        for(++lEntry; lEntry<lSize; lEntry++)
        {
            prdn++;

            if(ili>=prdn->_ili && ili<prdn->_ili+prdn->_cLines)
            {
                fPseudoHit = TRUE;
                break;
            }
        }

        if(!fPseudoHit)
        {
            CLinePtr    rp(_pdp);
            CTreePos*   ptp = NULL;
            DWORD       dwFlags = 0;

            dwFlags |= !!(phti->_grfFlags&HT_ALLOWEOL) ? CDisplay::CFP_ALLOWEOL : 0;
            dwFlags |= !(phti->_grfFlags&HT_DONTIGNOREBEFOREAFTER) ? CDisplay::CFP_IGNOREBEFOREAFTERSPACE : 0;
            dwFlags |= !(phti->_grfFlags&HT_NOEXACTFIT) ? CDisplay::CFP_EXACTFIT : 0;

            cp = _pdp->CpFromPointEx(
                ili, yli,
                cp, pt,
                &rp, &ptp, NULL, dwFlags,
                &phti->_phtr->_fRightOfCp,
                &fPseudoHit, &presultsHitTest->_cchPreChars, NULL);

            if(cp<pElementDispNode->GetFirstCp()-1
                || cp>pElementDispNode->GetLastCp()+1)
            {
                return FALSE;
            }

            presultsHitTest->_iliHit = rp;
            presultsHitTest->_ichHit = rp.RpGetIch();


            htc = pFL->BranchFromPointEx(pt, rp, ptp,
                pElementDispNode->GetFirstBranch(),
                &phti->_pNodeElement,
                fPseudoHit, &presultsHitTest->_fWantArrow,
                dwFlags&CDisplay::CFP_IGNOREBEFOREAFTERSPACE);
        }
        else
        {
            phti->_pNodeElement = pElementDispNode->GetFirstBranch();
            presultsHitTest->_iliHit = ili;
            presultsHitTest->_ichHit = 0;
        }

        presultsHitTest->_cpHit = cp;

        if(htc != HTC_YES)
        {
            presultsHitTest->_fWantArrow = TRUE;
        }
        else
        {
            phti->_ptContent = pt;
            phti->_pDispNode = pDispNode;
        }

        phti->_htc = htc;

        return htc == HTC_YES;
    }
    else
    {
        return FALSE;
    }
}

BOOL CRelDispNodeCache::HitTestFuzzy(const POINT* pptHitInParentCoords, CDispNode* pDispNode, void* pClientData)
{
    // no fuzzy hit test
    return FALSE;
}

LONG CRelDispNodeCache::CompareZOrder(CDispNode* pDispNode1, CDispNode* pDispNode2)
{
    CElement* pElement1 = ::GetDispNodeElement(pDispNode1);
    CElement* pElement2 = ::GetDispNodeElement(pDispNode2);

    return (pElement1!=pElement2 ? pElement1->CompareZOrder(pElement2) : 0);
}

//+----------------------------------------------------------------------------
//
// Function:    CRelDispNodeCache::DrawClient
//
// Synopsis:    Draw the background of the given disp node
//
//-----------------------------------------------------------------------------
void CRelDispNodeCache::DrawClientBackground(
                                        const RECT* prcBounds,
                                        const RECT* prcRedraw,
                                        CDispSurface* pSurface,
                                        CDispNode* pDispNode,
                                        void* pClientData,
                                        DWORD dwFlags)
{
    long lEntry = GetRelCacheEntry(pDispNode);

    Assert(Size() && lEntry>=0 && lEntry<Size());

    if(lEntry >= 0)
    {
        Assert(pClientData);

        CFormDrawInfo*  pDI = (CFormDrawInfo*)pClientData;
        CSetDrawSurface sds(pDI, prcBounds, prcRedraw, pSurface);
        CRelDispNode*   prdn;

        prdn = (*this)[lEntry];

        Assert(!prdn->GetElement()->NeedsLayout());

        CTreePos*   ptp;
        CDisplay*   pdp = GetDisplay();
        long        cp;

        prdn = (*this)[lEntry];

        ((CRect&)(pDI->_rcClip)).IntersectRect(*prcBounds);

        prdn->GetElement()->GetTreeExtent(&ptp, NULL);

        cp = ptp->GetCp();

        // Draw background and border of the current relative
        // element or any of it's descendents.
        pdp->DrawRelElemBgAndBorder(cp, ptp, prdn, prcBounds, prcRedraw, pDI);
    }
}

//+----------------------------------------------------------------------------
//
// Function:    CRelDispNodeCache::DrawClient
//
// Synopsis:    Draw the given disp node
//
//-----------------------------------------------------------------------------
void CRelDispNodeCache::DrawClient(
                              const RECT* prcBounds,
                              const RECT* prcRedraw,
                              CDispSurface* pSurface,
                              CDispNode* pDispNode,
                              void* cookie,
                              void* pClientData,
                              DWORD dwFlags)
{
    CRelDispNode* prdn;
    long lEntry = GetRelCacheEntry(pDispNode);

    Assert(Size() && lEntry>=0 && lEntry<Size());

    if(lEntry >= 0)
    {
        Assert(pClientData);

        // draw the lines here
        CFormDrawInfo* pDI = (CFormDrawInfo*)pClientData;
        CSetDrawSurface sds(pDI, prcBounds, prcRedraw, pSurface);

        prdn = (*this)[lEntry];

        Assert(!prdn->GetElement()->NeedsLayout());

        // Draw the lines that correspond to the given disp node
        HDC         hdc = pDI->GetDC();
        long        ili;
        CTreePos*   ptp;
        CDisplay*   pdp = GetDisplay();
        CElement*   pElementFL = pdp->GetFlowLayoutElement();
        CLine*      pli;
        CLSRenderer lsre(pdp, pDI);
        long        cp;

        if(!lsre.GetLS())
        {
            return;
        }

        prdn = (*this)[lEntry];

        ((CRect&)(pDI->_rcClip)).IntersectRect(*prcBounds);

        lsre.StartRender(*prcBounds, pDI->_rcClip);

        prdn->GetElement()->GetTreeExtent(&ptp, NULL);

        cp = ptp->GetCp();

        if(!prdn->_pDispNode->IsContainer())
        {
            // If the current relative element is not a container,
            // draw its background and border (or those of its
            // descendents).  Containers have their background
            // drawn via DrawClientBackground().
            pdp->DrawRelElemBgAndBorder(cp, ptp, prdn, prcBounds, prcRedraw, pDI);
        }

        lsre.SetCurPoint(CPoint(0, prcBounds->top - prdn->_ptOffset.y));
        lsre.SetCp(cp, ptp);

        for(ili=prdn->_ili; ili<prdn->_ili+prdn->_cLines; ili++)
        {
            // Find the current relative node that owns the current line
            CTreeNode* pNodeRelative = ptp->GetBranch()->GetCurrentRelativeNode(pElementFL);

            pli = pdp->Elem(ili);

            // Skip the current line if the owner is not the same element
            // that owns the current display node.
            if(pNodeRelative && pNodeRelative->Element()!=prdn->GetElement())
            {
                lsre.SkipCurLine(pli);
            }
            else
            {
                lsre.RenderLine(*pli, prcBounds->left-prdn->_ptOffset.x);
            }

            Assert(pli == pdp->Elem(ili));

            ptp = lsre.GetPtp();
            // It's possible for the renderer to return a pointer treepos; this is no good
            // to us, since we need the next content related node.  Keep walking till we
            // find one. (Bug #72264).
            while(ptp->Type() == CTreePos::Pointer)
            {
                ptp = ptp->NextTreePos();
            }
        }

        // restore the original text align, the renderer might have modified the
        // text align. This caused some pretty bad rendering problems(specially with
        // radio buttons).
        if(lsre._lastTextOutBy != CLSRenderer::DB_LINESERV)
        {
            lsre._lastTextOutBy = CLSRenderer::DB_NONE;
            SetTextAlign(hdc, TA_TOP|TA_LEFT);
        }
    }
}

CDispNode* CRelDispNodeCache::FindElementDispNode(CElement* pElement)
{
    int lSize = _aryRelDispNodes.Size();

    if (lSize)
    {
        CRelDispNode* prdn = _aryRelDispNodes;
        for(; lSize ; lSize--,prdn++)
        {
            if(prdn->GetElement() == pElement)
            {
                return prdn->_pDispNode;
            }
        }
    }
    return NULL;
}

void CRelDispNodeCache::Delete(long iPosFrom, long iPosTo)
{
    // we need to clearContents on each item
    if(iPosTo >= iPosFrom)
    {
        int i = iPosFrom;
        CRelDispNode* prdn = (*this)[i];

        for(; i<=iPosTo; i++,prdn++)
        {
            prdn->SetElement(NULL);
        }
    }

    _aryRelDispNodes.DeleteMultiple(iPosFrom, iPosTo);
}

void CRelDispNodeCache::SetElementDispNode(CElement* pElement, CDispNode* pDispNode)
{
    Assert(pElement);
    Assert(pDispNode);

    int lSize = _aryRelDispNodes.Size();

    if(lSize)
    {
        CRelDispNode* prdn = _aryRelDispNodes;
        for(; lSize ; lSize--,prdn++)
        {
            if(prdn->GetElement() == pElement)
            {
                prdn->_pDispNode = pDispNode;
            }
        }
    }
}

void CRelDispNodeCache::EnsureDispNodeVisibility(CElement* pElement)
{
    CLayout*        pLayout = _pdp->GetFlowLayout();
    CRelDispNode*   prdn    = _aryRelDispNodes;
    long            cSize   = _aryRelDispNodes.Size();

    for(; cSize; prdn++,cSize--)
    {
        if(!pElement || prdn->GetElement()==pElement)
        {
            pLayout->EnsureDispNodeVisibility(prdn->GetElement(), prdn->_pDispNode);
        }
    }
}

void CRelDispNodeCache::HandleDisplayChange()
{
    CRelDispNode*   prdn    = _aryRelDispNodes;
    long            cSize   = _aryRelDispNodes.Size();
    BOOL            fDisplayNone = _pdp->GetFlowLayoutElement()->IsDisplayNone();

    for(; cSize; prdn++,cSize--)
    {
        if(fDisplayNone || prdn->GetElement()->IsDisplayNone())
        {
            prdn->_pDispNode->ExtractFromTree();
        }
        else
        {
            CPoint ptAuto(prdn->_ptOffset.x, prdn->_ptOffset.y+prdn->_yli);

            prdn->GetElement()->ZChangeElement(NULL, &ptAuto);
        }
    }
}

void CRelDispNodeCache::DestroyDispNodes()
{
    long lSize = _aryRelDispNodes.Size();

    if(lSize)
    {
        CRelDispNode* prdn;
        for(prdn=(*this)[0]; lSize; lSize--,prdn++)
        {
            prdn->ClearContents();
        }
    }
}

void CRelDispNodeCache::Invalidate(
                                   CElement* pElement,          // Relative element to invalidate
                                   const RECT* prc /*=NULL*/,   // array of rects describing region to inval
                                   int nRects /*=1*/)           // count of rects in array
{
    // The element must itself actually be relative; it is insufficient for
    // it to have "inherited" relativeness (see element.hxx discussion of
    // IsRelative vs. IsInheritingRelativeness), since the RDNC tracks only
    // only relative elements and not their children.
    // We could use GetCurrentRelativeNode() instead of relying on the caller
    // to do that, but it would require an additional param (pElementFL)
    // that might obfuscate.
    Assert(pElement->IsRelative());

    CDispNode* pDispNode = FindElementDispNode(pElement);

    // If callers pass a relative element that we aren't responsible for, then
    // bail!  This shouldn't happen.
    if(!pDispNode)
    {
        // BUGBUG: this assert is valid.  See CFlowLayout::Notify() comments
        // on handling of invalidation.
        //Assert( FALSE && "Caller passed a relative element that doesn't belong to this RDNC!" );
        return;
    }

    {
        if(_pdp->GetFlowLayout()->OpenView())
        {
            // Incoming rects are in the coordinate system of the
            // flow layout.  We need to convert them into the system
            // of the dispnode, so we need the flow offset of the
            // dispnode..
            CPoint pt;
            _pdp->GetRelNodeFlowOffset(pDispNode, &pt);

            // Conversion involves subtracting flow offset of dispnode
            // from each rect:
            pt.x = -pt.x;
            pt.y = -pt.y;

            Assert(prc);
            for(int i=0; i<nRects; i++,prc++)
            {
                ::OffsetRect((RECT*)prc, pt.x, pt.y);
                pDispNode->Invalidate((CRect&)*prc, COORDSYS_CONTENT);
            }
        }
    }
}



// Timer tick counts for background task
#define cmsecBgndInterval   300
#define cmsecBgndBusy       100

// Lines ahead
const LONG g_cExtraBeforeLazy = 10;

// This function does exactly what IntersectRect does, except that
// if one of the rects is empty, it still returns TRUE if the rect
// is located inside the other rect. [ IntersectRect rect in such
// case returns FALSE. ]
BOOL IntersectRectE(RECT* prRes, const RECT* pr1, const RECT* pr2)
{
    // nAdjust is used to control what statement do we use in conditional
    // expressions: a < b or a <= b. nAdjust can be 0 or 1;
    // when (nAdjust == 0): (a - nAdjust < b) <==> (a <  b)  (*)
    // when (nAdjust == 1): (a - nAdjust < b) <==> (a <= b)  (**)
    // When at least one of rects to intersect is empty, and the empty
    // rect lies on boundary of the other, then we consider that the
    // rects DO intersect - in this case nAdjust == 0 and we use (*).
    // If both rects are not empty, and rects touch, then we should
    // consider that they DO NOT intersect and in that case nAdjust is
    // 1 and we use (**).
    int nAdjust;

    Assert(prRes && pr1 && pr2);
    Assert(pr1->left<=pr1->right && pr1->top<=pr1->bottom &&
        pr2->left<=pr2->right && pr2->top<=pr2->bottom);

    prRes->left  = max(pr1->left,  pr2->left);
    prRes->right = min(pr1->right, pr2->right);
    nAdjust = (int)((pr1->left!=pr1->right) && (pr2->left!=pr2->right));
    if(prRes->right-nAdjust < prRes->left)
    {
        goto NoIntersect;
    }

    prRes->top    = max(pr1->top, pr2->top);
    prRes->bottom = min(pr1->bottom, pr2->bottom);
    nAdjust = (int)((pr1->top!=pr1->bottom) && (pr2->top!=pr2->bottom));
    if(prRes->bottom-nAdjust < prRes->top)
    {
        goto NoIntersect;
    }

    return TRUE;

NoIntersect:
    SetRect(prRes, 0, 0, 0, 0);
    return FALSE;
}

// Start: Code to implement background recalc in lightwt tasks
class CRecalcTask : public CTask
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CRecalcTask(CDisplay* pdp, DWORD grfLayout)
    {
        _pdp = pdp ;
        _grfLayout = grfLayout;
    }

    virtual void OnRun(DWORD dwTimeOut)
    {
        _pdp->StepBackgroundRecalc(dwTimeOut, _grfLayout);
    }

    virtual void OnTerminate() {}

private:
    CDisplay* _pdp ;
    DWORD _grfLayout;
};
// End: Code to implement background recalc in lightwt tasks



// ===========================  CDisplay  =====================================================
CDisplay::~CDisplay()
{
    // The recalc task should have disappeared during the detach!
    Assert(!HasBgRecalcInfo() && !RecalcTask());
}

CElement* CDisplay::GetFlowLayoutElement() const
{
    return GetFlowLayout()->ElementContent();
}

CMarkup* CDisplay::GetMarkup() const
{
    return GetFlowLayout()->GetContentMarkup();
}

CDisplay::CDisplay()
{
    _fRecalcDone = TRUE;
}

//+----------------------------------------------------------------------------
//  Member:     CDisplay::Init
//
//  Synopsis:   Initializes CDisplay
//
//  Returns:    TRUE - initialization succeeded
//              FALSE - initalization failed
//
//+----------------------------------------------------------------------------
BOOL CDisplay::Init()
{
    CFlowLayout* pFL = GetFlowLayout();

    Assert(_yCalcMax     == 0 ); // Verify allocation zeroed memory out
    Assert(_xWidth       == 0 );
    Assert(_yHeight      == 0 );
    Assert(RecalcTask()  == NULL);

    SetWordWrap(pFL->GetWordWrap());

    _xWidthView = 0;

    return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     Detach
//
//  Synopsis:   Do stuff before dying
//
//-------------------------------------------------------------------------
void CDisplay::Detach()
{
    // If there's a timer on, get rid of it before we detach the
    // object. This prevents us from trying to recalc the lines
    // after the CTxtSite has gone away.
    FlushRecalc();
}

/*
*  CDisplay::GetFirstCp
*
*  @mfunc
*      Return the first cp
*/
LONG CDisplay::GetFirstCp() const
{
    return GetFlowLayout()->GetContentFirstCp();
}

/*
*  CDisplay::GetLastCp
*
*  @mfunc
*      Return the last cp
*/
LONG CDisplay::GetLastCp() const
{
    return GetFlowLayout()->GetContentLastCp();
}

/*
*  CDisplay::GetMaxCpCalced
*
*  @mfunc
*      Return the last cp calc'ed. Note that this is
*      relative to the start of the site, not the story.
*/
LONG CDisplay::GetMaxCpCalced() const
{
    return GetFlowLayout()->GetContentFirstCp()+_dcpCalcMax;
}

BOOL CDisplay::AllowBackgroundRecalc(CCalcInfo* pci, BOOL fBackground)
{
    CFlowLayout* pFL = GetFlowLayout();

    // Allow background recalc when:
    //  a) Not currently calcing in the background
    //  b) It is a SIZEMODE_NATURAL request
    //  c) The CTxtSite does not size to its contents
    //  d) The site is not part of a print document
    //  e) The site allows background recalc
    return (!fBackground &&
        (pci->_smMode==SIZEMODE_NATURAL) &&
        !(pci->_grfLayout&LAYOUT_NOBACKGROUND) &&
        !pFL->_fContentsAffectSize &&
        !pFL->GetAutoSize() &&
        !pFL->TestClassFlag(CElement::ELEMENTDESC_NOBKGRDRECALC));
}

/*
*  CDisplay::FlushRecalc()
*
*  @mfunc
*      Destroys the line array, therefore ensure a full recalc next time
*      RecalcView or UpdateView is called.
*
*/
void CDisplay::FlushRecalc()
{
    CFlowLayout* pFL = GetFlowLayout();

    StopBackgroundRecalc();

    if(LineCount())
    {
        Remove(0, -1, AF_KEEPMEM); // Remove all old lines from *this
    }

    pFL->_fContainsRelative   = FALSE;
    pFL->_fChildWidthPercent  = FALSE;
    pFL->_fChildHeightPercent = FALSE;
    pFL->CancelChanges();

    VoidRelDispNodeCache();
    DestroyFlowDispNodes();

    _fRecalcDone = FALSE;
    _yCalcMax   = 0;    // Set both maxes to start of text
    _dcpCalcMax = 0;    // Relative to the start cp of site.
    _xWidth     = 0;
    _yHeight    = 0;
    _yHeightMax = 0;

    _fContainsHorzPercentAttr = _fContainsVertPercentAttr =
        _fNavHackPossible = _fHasEmbeddedLayouts = _fHasMultipleTextNodes = FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     NoteMost
//
//  Purpose:    Notes if the line has anything which will need us to compute
//              information for negative lines/absolute or relative divs
//
//----------------------------------------------------------------------------
void CDisplay::NoteMost(CLine* pli)
{
    Assert (pli);

    if(!_fRecalcMost && (pli->GetYMostTop()<0 || pli->_fHasAbsoluteElt))
    {
        _fRecalcMost = TRUE;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RecalcMost
//
//  Purpose:    Calculates the most negative line and most positive/negative
//              positioned site from scratch.
//
//  BUGBUG(sujalp): We initially had an incremental way of computing both the
//  negative line hts info AND +/- positioned site info. However, this logic was
//  incorrect for consecutive lines with -ve line height. So we changed it so
//  that we compute this AND +/- info always. If this becomes a performance issue
//  we could go back for incremental computation for div's easily -- but would
//  have to maintain extra state information. For negative line heights we could
//  also do some incremental stuff, but it would be much much more complicated
//  than what we have now.
//
//----------------------------------------------------------------------------
void CDisplay::RecalcMost()
{
    if(_fRecalcMost)
    {
        CFlowLayout* pFlowLayout = GetFlowLayout();
        LONG ili;

        long yNegOffset = 0;        // offset at which the current line is drawn
                                    // as a result of a series of lines with negative
                                    // height
        long yPosOffset = 0;
        long yBottomOffset = 0;     // offset by which the current lines contents
                                    // extend beyond the yHeight of the line.
        long yTopOffset = 0;        // offset by which the current lines contents
                                    // extend before the current y position

        pFlowLayout->_fContainsAbsolute = FALSE;
        _yMostNeg = 0;
        _yMostPos = 0;

        for(ili=0; ili<LineCount(); ili++)
        {
            CLine* pli = Elem (ili);
            LONG yLineBottomOffset = pli->GetYHeightBottomOff();

            // Update most positive/negative positions for relative positioned sites
            if(pli->_fHasAbsoluteElt)
            {
                pFlowLayout->_fContainsAbsolute = FALSE;
            }

            // top offset of the current line
            yTopOffset = pli->GetYMostTop() + yNegOffset;

            yBottomOffset = yLineBottomOffset + yPosOffset;

            // update the most negative value if the line has negative before space
            // or line height < actual extent
            if(yTopOffset<0 && _yMostNeg>yTopOffset)
            {
                _yMostNeg = yTopOffset;
            }

            if(yBottomOffset>0 && _yMostPos<yBottomOffset)
            {
                _yMostPos = yBottomOffset;
            }

            // if the current line forces a new line and has negative height
            // update the negative offset at which the next line is drawn.
            if(pli->_fForceNewLine)
            {
                if(pli->_yHeight < 0)
                {
                    yNegOffset += pli->_yHeight;
                }
                else
                {
                    yNegOffset = 0;
                }

                if (yLineBottomOffset > 0)
                {
                    yPosOffset += yLineBottomOffset;
                }
                else
                {
                    yPosOffset = 0;
                }
            }
        }

        _fRecalcMost = FALSE;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RecalcPlainTextSingleLine
//
//  Purpose:    A high-performance substitute for RecalcView. Does not go
//              through Line Services. Can only be used to measure a single
//              line of plain text (i.e. no HTML).
//
//----------------------------------------------------------------------------
BOOL CDisplay::RecalcPlainTextSingleLine(CCalcInfo* pci)
{
    CFlowLayout*        pFlowLayout = GetFlowLayout();
    CTreeNode*          pNode       = pFlowLayout->GetFirstBranch();
    TCHAR               chPassword  = pFlowLayout->GetPasswordCh();
    long                cch         = pFlowLayout->GetContentTextLength();
    const CCharFormat*  pCF         = pNode->GetCharFormat();
    const CParaFormat*  pPF         = pNode->GetParaFormat();
    CCcs*               pccs = NULL;
    CBaseCcs*           pBaseCcs = NULL;
    long                lWidth;
    long                lCharWidth;
    long                xShift;
    long                xWidth, yHeight;
    long                lPadding[PADDING_MAX];
    BOOL                fViewChanged = FALSE;
    CLine*              pli;
    UINT                uJustified;
    long                xDummy, yBottomMarginOld = _yBottomMargin;


    Assert(pPF);
    Assert(pCF);
    Assert(pci);
    Assert(cch >= 0);

    if(!pPF || !pCF || !pci || cch<0)
    {
        return FALSE;
    }

    // Bail out if there is anything special in the format
    // BUGBUG (MohanB) Should be able to handle line height. For now though,
    // do this to keep RoboVision happy (with lheight.htm).
    // (paulnel) Right to left text cannot be done the 'quick' way
    if(pCF->IsTextTransformNeeded()
        || !pCF->_cuvLineHeight.IsNullOrEnum()
        || !pCF->_cuvLetterSpacing.IsNullOrEnum()
        || pCF->_fRTL)
    {
        goto HardCase;
    }


    pccs = fc().GetCcs(pci->_hdc, pci, pCF);
    if(!pccs)
    {
        return FALSE;
    }
    pBaseCcs = pccs->GetBaseCcs();

    lWidth = 0;

    if(cch)
    {
        if(chPassword)
        {
            if(!pccs->Include(chPassword, lCharWidth) )
            {
                Assert(0 && "Char not in font!");
            }
            lWidth = cch * lCharWidth;
        }
        else
        {
            CTxtPtr tp(GetMarkup(), pFlowLayout->GetContentFirstCp());
            LONG    cchValid;
            LONG    cchRemaining = cch;

            for(;;)
            {
                const TCHAR* pchText = tp.GetPch(cchValid);
                LONG i = min(cchRemaining, cchValid);

                while(i--)
                {
                    const TCHAR ch = *pchText++;

                    // Bail out if not a simple ASCII char
                    if(!InRange(ch, 32, 127))
                    {
                        goto HardCase;
                    }

                    if(!pccs->Include(ch, lCharWidth))
                    {
                        Assert(0 && "Char not in font!");
                    }
                    lWidth += lCharWidth;
                }

                if(cchRemaining <= cchValid)
                {
                    break;
                }
                else
                {
                    cchRemaining -= cchValid;
                    tp.AdvanceCp(cchValid);
                }
            }
        }
    }

    GetPadding(pci, lPadding, pci->_smMode==SIZEMODE_MMWIDTH);
    FlushRecalc();

    pli = Add(1, NULL);
    if(!pli)
    {
        return FALSE;
    }

    pli->_cch               = cch;
    pli->_xWidth            = lWidth;
    pli->_yTxtDescent       = pBaseCcs->_yDescent;
    pli->_yDescent          = pBaseCcs->_yDescent;
    pli->_xLineOverhang     = pBaseCcs->_xOverhang;
    pli->_yExtent           = pBaseCcs->_yHeight;
    pli->_yBeforeSpace      = lPadding[PADDING_TOP];
    pli->_yHeight           = pBaseCcs->_yHeight + pli->_yBeforeSpace;
    pli->_xLeft             = lPadding[PADDING_LEFT];
    pli->_xRight            = lPadding[PADDING_RIGHT];
    pli->_xLeftMargin       = 0;
    pli->_xRightMargin      = 0;
    pli->_fForceNewLine     = TRUE;
    pli->_fFirstInPara      = TRUE;
    pli->_fFirstFragInLine  = TRUE;
    pli->_fCanBlastToScreen = !chPassword && !pCF->_fDisabled;

    _yBottomMargin  = lPadding[PADDING_BOTTOM];
    _dcpCalcMax     = cch;

    yHeight          = pli->_yHeight;
    xWidth           = pli->_xLeft + pli->_xWidth + pli->_xLineOverhang + pli->_xRight;

    Assert(pci->_smMode != SIZEMODE_MINWIDTH);

    xShift = ComputeLineShift(
        (htmlAlign)pPF->GetBlockAlign(TRUE),
        _fRTL,
        pPF->HasRTL(TRUE),
        pci->_smMode==SIZEMODE_MMWIDTH,
        _xWidthView,
        xWidth+GetCaret(),
        &uJustified,
        &xDummy);

    pli->_fJustified = uJustified;

    if(!_fRTL)
    {
        pli->_xLeft += xShift;
    }
    else
    {
        pli->_xRight += xShift;
    }

    xWidth += xShift;

    pli->_xLineWidth = max(xWidth, _xWidthView);

    if(yHeight+yBottomMarginOld!=_yHeight+_yBottomMargin || xWidth!=_xWidth)
    {
        fViewChanged = TRUE;
    }

    _yCalcMax       = _yHeightMax = _yHeight = yHeight;
    _xWidth         = xWidth;
    _fRecalcDone    = TRUE;
    _fMinMaxCalced  = TRUE;
    _xMinWidth      = _xMaxWidth  = _xWidth + GetCaret();

    if(pci->_smMode==SIZEMODE_NATURAL && fViewChanged)
    {
        pFlowLayout->SizeContentDispNode(CSize(GetMaxWidth(), GetHeight()));

        // If our contents affects our size, ask our parent to initiate a re-size
        ElementResize(pFlowLayout);

        pFlowLayout->NotifyMeasuredRange(pFlowLayout->GetContentFirstCp(),
            GetMaxCpCalced());
    }

    pccs->Release();

    return TRUE;

HardCase:
    if(pccs)
    {
        pccs->Release();
    }

    // Just do it the hard way
    return RecalcLines(pci);
}

/*
*  CDisplay::RecalcLines()
*
*  @mfunc
*      Recalc all line breaks.
*      This method does a lazy calc after the last visible line
*      except for a bottomless control
*
*  @rdesc
*      TRUE if success
*/

BOOL CDisplay::RecalcLines(CCalcInfo* pci)
{
    BOOL fRet;

    if(GetFlowLayout()->ElementOwner()->IsInMarkup())
    {
        //if(pci->_fTableCalcInfo && ((CTableCalcInfo*)pci)->_pme)
        //{
        //    // Save calcinfo's measurer.
        //    CTableCalcInfo* ptci = (CTableCalcInfo*) pci;
        //    CLSMeasurer* pme = ptci->_pme;

        //    // Reinitialize the measurer.
        //    pme->Reinit(this, ptci);

        //    // Make sure noone else uses this measurer.
        //    ptci->_pme = NULL;

        //    // Do actual RecalcLines work with this measurer.
        //    fRet = RecalcLinesWithMeasurer(ptci, pme);

        //    // Restore TableCalcInfo measurer.
        //    ptci->_pme = pme;
        //}
        //else
        {
            // Cook up measurer on the stack.
            CLSMeasurer me(this, pci);

            fRet = RecalcLinesWithMeasurer(pci, &me);
        }
    }
    else
    {
        fRet = FALSE;
    }

    return fRet;
}

/*
*  CDisplay::RecalcLines()
*
*  @mfunc
*      Recalc all line breaks.
*      This method does a lazy calc after the last visible line
*      except for a bottomless control
*
*  @rdesc
*      TRUE if success
*/
BOOL CDisplay::RecalcLinesWithMeasurer(CCalcInfo* pci, CLSMeasurer* pme)
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CElement*       pElementFL  = pFlowLayout->ElementOwner();
    CElement::CLock Lock(pElementFL, CElement::ELEMENTLOCK_RECALC);

    CRecalcLinePtr  RecalcLinePtr(this, pci);
    CLine*          pliNew    = NULL;
    int             iLine     = -1;
    int             iLinePrev = -1;
    long            yHeightPrev;
    UINT            uiMode;
    UINT            uiFlags;
    BOOL            fDone                   = TRUE;
    BOOL            fFirstInPara            = TRUE;
    BOOL            fWaitForCpToBeCalcedTo  = TRUE;
    LONG            cpToBeCalcedTo          = 0;
    BOOL            fAllowBackgroundRecalc;
    LONG            yHeight         = 0;
    LONG            yAlignDescent   = 0;
    LONG            yHeightView     = GetViewHeight();
    LONG            xMinLineWidth   = 0;
    LONG*           pxMinLineWidth  = NULL;
    LONG            xMaxLineWidth   = 0;
    LONG            yHeightDisplay  = 0;     // to keep track of line with top negative margins
    LONG            yHeightOld      = _yHeight;
    LONG            yHeightMaxOld   = _yHeightMax;
    LONG            xWidthOld       = _xWidth;
    LONG            yBottomMarginOld= _yBottomMargin;
    BOOL            fViewChanged    = FALSE;
    BOOL            fNormalMode     = pci->_smMode==SIZEMODE_NATURAL
        || pci->_smMode==SIZEMODE_SET || pci->_smMode==SIZEMODE_FULLSIZE;
    CDispNode*      pDNBefore;
    long            lPadding[PADDING_MAX];

    // we should not be measuring hidden stuff.
    Assert(!pElementFL->IsDisplayNone() || pElementFL->Tag()==ETAG_BODY);

    if(pElementFL->Tag() == ETAG_ROOT)
    {
        return TRUE;
    }

    if(!pme->_pLS)
    {
        return FALSE;
    }

    if(!pElementFL->IsInMarkup())
    {
        return TRUE;
    }

    GetPadding(pci, lPadding, pci->_smMode==SIZEMODE_MMWIDTH);

    _yBottomMargin = lPadding[PADDING_BOTTOM];

    // Set up the CCalcInfo to the correct mode and parent size
    if(pci->_smMode == SIZEMODE_SET)
    {
        pci->_smMode = SIZEMODE_NATURAL;
    }

    // Determine the mode
    uiMode = MEASURE_BREAKATWORD;

    Assert(pci);

    switch(pci->_smMode)
    {
    case SIZEMODE_MMWIDTH:
        uiMode = MEASURE_BREAKATWORD | MEASURE_MAXWIDTH;
        _xMinWidth = 0;
        pxMinLineWidth = &xMinLineWidth;
        break;
    case SIZEMODE_MINWIDTH:
        uiMode = MEASURE_BREAKATWORD | MEASURE_MINWIDTH;
        pme->AdvanceToNextAlignedObject();
        break;
    }

    uiFlags = uiMode;

    // Determine if background recalc is allowed
    // (And if it is, do not calc past the visible portion)
    fAllowBackgroundRecalc = AllowBackgroundRecalc(pci);

    if(fAllowBackgroundRecalc)
    {
        cpToBeCalcedTo = max(GetFirstVisibleCp(), CPWait());
    }

    // Flush all old lines
    FlushRecalc();

    if(GetWordWrap() && GetWrapLongLines())
    {
        uiFlags |= MEASURE_BREAKLONGLINES;
    }

    pme->_pDispNodePrev = NULL;

    if(fAllowBackgroundRecalc)
    {
        if(!SUCCEEDED(EnsureBgRecalcInfo()))
        {
            fAllowBackgroundRecalc = FALSE;
            AssertSz(FALSE, "CDisplay::RecalcLinesWithMeasurer - Could not create BgRecalcInfo");
        }
    }

    RecalcLinePtr.Init((CLineArray*)this, 0, NULL);

    // recalculate margins
    RecalcLinePtr.RecalcMargins(0, 0, yHeight, 0);

    if(pElementFL->IsDisplayNone())
    {
        AssertSz(pElementFL->Tag()==ETAG_BODY, "Cannot measure any hidden layout other than the body");

        pliNew = RecalcLinePtr.AddLine();
        if(!pliNew)
        {
            Assert(FALSE);
            goto err;
        }
        pme->NewLine(uiFlags&MEASURE_FIRSTINPARA);
        *pliNew = pme->_li;
        pliNew->_cch = pElementFL->GetElementCch();
    }
    else
    {
        // The following loop generates new lines
        do
        {
            // Add one new line
            pliNew = RecalcLinePtr.AddLine();
            if(!pliNew)
            {
                Assert(FALSE);
                goto err;
            }

            uiFlags &= ~MEASURE_FIRSTINPARA;
            uiFlags |= (fFirstInPara ? MEASURE_FIRSTINPARA : 0);

            iLine = LineCount() - 1;

            if(long(pme->GetCp()) == pme->GetLastCp())
            {
                uiFlags |= MEASURE_EMPTYLASTLINE;
            }

            // Cache the previous dispnode to insert a new content node
            // if the next line measured has negative margins(offset) which
            // causes the next line on top of the previous lines.
            pDNBefore = pme->_pDispNodePrev;
            iLinePrev = iLine;
            yHeightPrev = yHeight;

            if(!RecalcLinePtr.MeasureLine(*pme, uiFlags,
                &iLine, &yHeight, &yAlignDescent, pxMinLineWidth, &xMaxLineWidth))
            {
                goto err;
            }

            // fix for bug 16519 (srinib)
            // Keep track of the line that contributes to the max height, ex: if the
            // last lines top margin is negative.
            if(yHeightDisplay < yHeight)
            {
                yHeightDisplay = yHeight;
            }

            // iLine returned is the last text line inserted. There ay be
            // aligned lines and clear lines added before and after the
            // text line
            pliNew = iLine>=0 ? RecalcLinePtr[iLine] : NULL;
            fFirstInPara = (iLine>=0) ? pliNew->IsNextLineFirstInPara() : TRUE;

            if(fNormalMode && iLine>=0)
            {
                HandleNegativelyPositionedContent(pliNew, pme, pDNBefore, iLinePrev, yHeightPrev);
            }

            if(fAllowBackgroundRecalc)
            {
                Assert(HasBgRecalcInfo());

                if(fWaitForCpToBeCalcedTo)
                {
                    if((LONG)pme->GetCp() > cpToBeCalcedTo)
                    {
                        BgRecalcInfo()->_yWait = yHeight + yHeightView;
                        fWaitForCpToBeCalcedTo = FALSE;
                    }
                }
                else if(yHeightDisplay > YWait())
                {
                    fDone = FALSE;
                    break;
                }
            }

            // When doing a full min pass, just calculate the islands around aligned
            // objects. Note that this still calc's the lines around right aligned
            // images, which isn't really necessary because they are willing to
            // overlap text. Fix it next time.
            if((uiMode & MEASURE_MINWIDTH) &&
                !RecalcLinePtr._marginInfo._xLeftMargin &&
                !RecalcLinePtr._marginInfo._xRightMargin)
            {
                pme->AdvanceToNextAlignedObject();
            }
        } while(pme->GetCp() < pme->GetLastCp());
    }

    _fRecalcDone = fDone;

    if(fDone && pliNew)
    {
        if(GetFlowLayout()->IsEditable() && // if we are in design mode
            (pliNew->_fHasEOP||pliNew->_fHasBreak||pliNew->_fSingleSite))
        {
            Assert(pliNew == RecalcLinePtr[iLine]);
            CreateEmptyLine(pme, &RecalcLinePtr, &yHeight, pliNew->_fHasEOP);

            // Creating the line could have reallocated the memory to which pliNew points,
            // so we need to refresh the pointer just in case.
            pliNew = RecalcLinePtr[iLine];
        }

        // fix for bug 16519
        // Keep track of the line that contributes to the max height, ex: if the
        // last lines top margin is negative.
        if(yHeightDisplay < yHeight)
        {
            yHeightDisplay = yHeight;
        }
        // In table cells, Netscape actually adds the interparagraph spacing
        // for any closed tags to the height of the table cell.
        // BUGBUG: This actually applies the spacing to all nested text containers, not just
        //         table cells. Is this correct? (brendand).
        // It's not correct to add the spacing at all, but given that Netscape
        // has done so, and that they will probably do so for floating block
        // elements. Arye.
        else
        {
            int iLineLast = iLine;

            // we need to force scrolling when bottom-margin is set on the last block tag
            // in body. (20400)
            while(iLineLast-->0 && // Lines available
                !pliNew->_fForceNewLine) // Just a dummy line.
            {
                --pliNew;
            }
            if(iLineLast>0 || pliNew->_fForceNewLine)
            {
                _yBottomMargin += RecalcLinePtr.NetscapeBottomMargin(pme);
            }
        }
    }

    if(!(uiMode & MEASURE_MAXWIDTH))
    {
        xMaxLineWidth = CalcDisplayWidth();
    }

    _dcpCalcMax = (LONG)pme->GetCp() - GetFirstCp();
    _yCalcMax   =  _yHeight = yHeightDisplay;
    _yHeightMax = max(yHeightDisplay, yAlignDescent);
    _xWidth     = xMaxLineWidth;

    // Fire a resize notification if the content size changed
    //  
    // fix for 52305, change to margin bottom affects the content size,
    // so we should fire a resize.
    if(_yHeight!=yHeightOld
        || _yHeightMax+_yBottomMargin!=yHeightMaxOld+yBottomMarginOld
        || _xWidth!=xWidthOld)
    {
        fViewChanged = TRUE;
    }

    {
        BOOL fAlignedLayouts = RecalcLinePtr._cLeftAlignedLayouts
            || (RecalcLinePtr._cRightAlignedLayouts>1);
        if(pxMinLineWidth)
        {
            pFlowLayout->MarkHasAlignedLayouts(fAlignedLayouts);
        }
        if(fAlignedLayouts)
        {
            _fNoContent = FALSE;
        }
    }

    // If the view or size changed, re-size or update scrollbars as appropriate
    // BUGBUG: Collapse both RecalcLines into one and remove knowledge of
    //         CSite::SIZEMODE_xxxx (it should be based entirely on the
    //         passed MEASURE_xxxx flags set in CalcSize) (brendand)
    if(pci->_smMode==SIZEMODE_NATURAL && (fViewChanged||_fHasMultipleTextNodes))
    {
        pFlowLayout->SizeContentDispNode(CSize(GetMaxWidth(), GetHeight()));

        // If our contents affects our size, ask our parent to initiate a re-size
        if(fViewChanged)
        {
            ElementResize(pFlowLayout);
        }
    }

    if(fNormalMode)
    {
        if(pFlowLayout->_fContainsRelative)
        {
            UpdateRelDispNodeCache(NULL);
        }
        AdjustDispNodes(pme->_pDispNodePrev, NULL);

        pFlowLayout->NotifyMeasuredRange(GetFlowLayoutElement()->GetFirstCp(), GetMaxCpCalced());
    }

    if(!fDone) // if not done, do rest in background
    {
        StartBackgroundRecalc(pci->_grfLayout);
    }
    else if(fAllowBackgroundRecalc)
    {
        Assert((!CanHaveBgRecalcInfo()||BgRecalcInfo()) && "Should have a BgRecalcInfo");
        if(HasBgRecalcInfo())
        {
            CBgRecalcInfo* pBgRecalcInfo = BgRecalcInfo();
            pBgRecalcInfo->_yWait = -1;
            pBgRecalcInfo->_cpWait = -1;
        }
        _fLineRecalcErr = FALSE;

        RecalcMost();
    }

    // cache min/max only if there are no sites inside !
    if(pxMinLineWidth && !pFlowLayout->ContainsChildLayout())
    {
        _xMaxWidth     = _xWidth + GetCaret();
        _fMinMaxCalced = TRUE;
    }

    // adjust for caret only when are calculating for MMWIDTH or MINWIDTH
    if(pci->_smMode==SIZEMODE_MMWIDTH || pci->_smMode==SIZEMODE_MINWIDTH)
    {
        if(pci->_smMode == SIZEMODE_MMWIDTH)
        {
            _xMinWidth = max(_xMinWidth, RecalcLinePtr._xMaxRightAlign);
        }
        _xMinWidth += GetCaret(); // adjust for caret only when are calculating for MMWIDTH
    }

    // NETSCAPE: If there is no text or special characters, treat the site as
    //           empty. For example, with an empty BLOCKQUOTE tag, _xWidth will
    //           not be zero while the site should be considered empty.
    if(_fNoContent)
    {
        _xWidth = _xMinWidth = lPadding[PADDING_LEFT] + lPadding[PADDING_RIGHT];
    }

    return TRUE;

err:
    if(!_fLineRecalcErr)
    {
        _dcpCalcMax = pme->GetCp() - GetFirstCp();
        _yCalcMax = yHeightDisplay;
    }

    return FALSE;
}

/*
*  CDisplay::RecalcLines(cpFirst, cchOld, cchNew, fBackground, pled)
*
*  @mfunc
*      Recompute line breaks after text modification
*
*  @rdesc
*      TRUE if success
*/
BOOL CDisplay::RecalcLines (
                            CCalcInfo* pci,
                            LONG cpFirst,               //@parm Where change happened
                            LONG cchOld,                //@parm Count of chars deleted
                            LONG cchNew,                //@parm Count of chars added
                            BOOL fBackground,           //@parm This method called as background process
                            CLed* pled,                 //@parm Records & returns impact of edit on lines (can be NULL)
                            BOOL fHack)                 //@parm This comes from WaitForRecalc ... don't mess with BgRecalcInfo
{
    CSaveCalcInfo   sci(pci);
    CElement::CLock Lock(GetFlowLayoutElement(), CElement::ELEMENTLOCK_RECALC);

    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CElement*       pElementContent  = pFlowLayout->ElementContent();

    LONG            cchEdit;
    LONG            cchSkip = 0;
    INT             cliBackedUp = 0;
    LONG            cliWait = g_cExtraBeforeLazy;
    BOOL            fDone = TRUE;
    BOOL            fFirstInPara = TRUE;
    BOOL            fAllowBackgroundRecalc;
    CLed            led;
    LONG            lT = 0;         // long Temporary
    CLine*          pliNew;
    CLinePtr        rpOld(this);
    CLine*          pliCur;
    LONG            xWidth;
    LONG            yHeight, yAlignDescent=0;
    LONG            cpLast = GetLastCp();
    long            cpLayoutFirst = GetFirstCp();
    UINT            uiFlags = MEASURE_BREAKATWORD;
    BOOL            fReplaceResult = TRUE;
    BOOL            fTryForMatch = TRUE;
    BOOL            fNeedToBackUp = TRUE;
    int             diNonBlankLine = -1;
    int             iOldLine;
    int             iNewLine;
    int             iLinePrev = -1;
    int             iMinimumLinesToRecalc = 4;   // Guarantee some progress
    LONG            yHeightDisplay = 0;
    LONG            yHeightMax;
    CLineArray      aryNewLines;
    CDispNode*      pDNBefore;
    long            yHeightPrev;
    long            yBottomMarginOld = _yBottomMargin;

    if(!pFlowLayout->ElementOwner()->IsInMarkup())
    {
        return TRUE;
    }

    // we should not be measuring hidden stuff.
    Assert(!pFlowLayout->ElementOwner()->IsDisplayNone() || pFlowLayout->ElementOwner()->Tag()==ETAG_BODY);

    // If no lines, this routine should not be called
    // Call the other RecalcLines above instead !
    Assert(LineCount() > 0);

    // Set up the CCalcInfo to the correct mode and parent size
    if(pci->_smMode == SIZEMODE_SET)
    {
        pci->_smMode = SIZEMODE_NATURAL;
    }

    // Init measurer at cpFirst
    CLSMeasurer me(this, cpFirst, pci);
    CRecalcLinePtr RecalcLinePtr(this, pci);

    if(!me._pLS)
    {
        goto err;
    }

    if(!pled)
    {
        pled = &led;
    }

    if(GetWordWrap() && GetWrapLongLines())
    {
        uiFlags |= MEASURE_BREAKLONGLINES;
    }

    // Determine if background recalc is allowed
    fAllowBackgroundRecalc = AllowBackgroundRecalc(pci, fBackground);

    Assert(!fBackground || HasBgRecalcInfo());
    if(!fBackground && !fHack && fAllowBackgroundRecalc)
    {
        if(SUCCEEDED(EnsureBgRecalcInfo()))
        {
            BgRecalcInfo()->_yWait  = pFlowLayout->GetClientBottom();
            BgRecalcInfo()->_cpWait = -1;
        }
        else
        {
            fAllowBackgroundRecalc = FALSE;
            AssertSz(FALSE, "CDisplay::RecalcLines - Could not create BgRecalcInfo");
        }
    }

    // Init line pointer on old CLineArray
    rpOld.RpSetCp(cpFirst, FALSE, FALSE);

    pliCur = rpOld.CurLine();

    while(pliCur->IsClear() ||
        (pliCur->IsFrame()&&!pliCur->_fFrameBeforeText))
    {
        if(!rpOld.AdjustBackward())
        {
            break;
        }

        pliCur = rpOld.CurLine();
    }

    // If the current line has single site
    if(pliCur->_fSingleSite)
    {
        // If we are in the middle of the current line (some thing changed
        // in a table cell or 1d-div, or marquee)
        if(rpOld.RpGetIch() && rpOld.GetCchRemaining()!=0)
        {
            // we dont need to back up
            if(rpOld>0 && rpOld[-1]._fForceNewLine)
            {
                fNeedToBackUp = FALSE;
                cchSkip = rpOld.RpGetIch();
            }
        }
    }

    iOldLine = rpOld;

    if(fNeedToBackUp)
    {
        if(!rpOld->IsBlankLine())
        {
            cchSkip = rpOld.RpGetIch();
            if(cchSkip)
            {
                rpOld.RpAdvanceCp(-cchSkip);
            }
        }

        // find the first previous non blank line, if the first non blank line
        // is a single site line or a line with EOP, we don't need to skip.
        while(rpOld+diNonBlankLine>=0 && rpOld[diNonBlankLine].IsBlankLine())
        {
            diNonBlankLine--;
        }

        // (srinib) - if the non text line does not have a line break or EOP, or
        // a force newline or is single site line and the cp at which the characters changed is
        // ambiguous, we need to back up.

        // if the single site line does not force a new line, we do need to
        // back up. (bug #44346)
        if(rpOld+diNonBlankLine >= 0)
        {
            pliCur = &rpOld[diNonBlankLine];
            if(!pliCur->_fSingleSite || !pliCur->_fForceNewLine || cchSkip==0)
            {
                // Skip past all the non text lines
                while(diNonBlankLine++ < 0)
                {
                    rpOld.AdjustBackward();
                }

                // Skip to the begining of the line
                cchSkip += rpOld.RpGetIch();
                rpOld.RpAdvanceCp(-rpOld.RpGetIch());

                // we want to skip all the lines that do not force a
                // newline, so find the last line that does not force a new line.
                long diTmpLine = -1;
                long cchSkipTemp = 0;
                while(rpOld+diTmpLine>=0 &&
                    ((pliCur=&rpOld[diTmpLine])!=0) &&
                    !pliCur->_fForceNewLine)
                {
                    if(!pliCur->IsBlankLine())
                    {
                        cchSkipTemp += pliCur->_cch;
                    }
                    diTmpLine--;
                }
                if(cchSkipTemp)
                {
                    cchSkip += cchSkipTemp;
                    rpOld.RpAdvanceCp(-cchSkipTemp);
                }
            }
        }

        // back up further if the previous lines are either frames inserted
        // by aligned sites at the beginning of the line or auto clear lines
        while(rpOld && ((pliCur=&rpOld[-1])!=0) &&
            // frame line before the actual text or
            (pliCur->_fClearBefore ||
            (pliCur->IsFrame() && pliCur->_fFrameBeforeText)) &&
            rpOld.AdjustBackward());
    }

    cliBackedUp = iOldLine - rpOld;

    // Need to get a running start.
    me.Advance(-cchSkip);
    RecalcLinePtr.InitPrevAfter(&me._fLastWasBreak, rpOld);

    cchEdit = cchNew + cchSkip; // Number of chars affected by edit

    Assert(cchEdit <= GetLastCp()-long(me.GetCp()) );
    if(cchEdit > GetLastCp()-long(me.GetCp()))
    {
        // BUGBUG (istvanc) this of course shouldn't happen !!!!!!!!
        cchEdit = GetLastCp() - me.GetCp();
    }

    // Determine whether we're on first line of paragraph
    if(rpOld > 0)
    {
        int iLastNonFrame = -1;

        // frames do not hold any info, so go back past all frames.
        while(rpOld+iLastNonFrame && (rpOld[iLastNonFrame].IsFrame() || rpOld[iLastNonFrame].IsClear()))
        {
            iLastNonFrame--;
        }
        if(rpOld+iLastNonFrame >= 0)
        {
            CLine* pliNew = &rpOld[iLastNonFrame];

            fFirstInPara = pliNew->IsNextLineFirstInPara();
        }
        else
        {
            // we are at the Beginning of a document
            fFirstInPara = TRUE;
        }
    }

    yHeight = YposFromLine(rpOld, pci);
    yHeightDisplay = yHeight;
    yAlignDescent = 0;

    // Update first-affected and pre-edit-match lines in pled
    pled->_iliFirst = rpOld;
    pled->_cpFirst  = pled->_cpMatchOld = me.GetCp();
    pled->_yFirst   = pled->_yMatchOld  = yHeight;

    // In case of error, set both maxes to where we are now
    _yCalcMax   = yHeight;
    _dcpCalcMax = me.GetCp() - cpLayoutFirst;

    // find the display node the corresponds to cpStart
    me._pDispNodePrev = pled->_iliFirst
        ? GetPreviousDispNode(pled->_cpFirst, pled->_iliFirst) : 0;


    // If we are past the requested area to recalc and background recalc is
    // allowed, then just go directly to background recalc.
    if(fAllowBackgroundRecalc && yHeight>YWait() && (LONG)me.GetCp()>CPWait())
    {
        // Remove all old lines from here on
        rpOld.RemoveRel(-1, AF_KEEPMEM);

        // Start up the background recalc
        StartBackgroundRecalc(pci->_grfLayout);
        pled->SetNoMatch();

        // Update the relative line cache.
        if(pFlowLayout->_fContainsRelative)
        {
            UpdateRelDispNodeCache(NULL);
        }

        AdjustDispNodes(me._pDispNodePrev, pled);

        goto cleanup;
    }

    aryNewLines.Clear(AF_KEEPMEM);
    pliNew = NULL;

    iOldLine = rpOld;
    RecalcLinePtr.Init((CLineArray*)this, iOldLine, &aryNewLines);

    // recalculate margins
    RecalcLinePtr.RecalcMargins(iOldLine, iOldLine, yHeight, 0);

    Assert(cchEdit <= GetLastCp()-long(me.GetCp()));

    if(iOldLine)
    {
        RecalcLinePtr.SetupMeasurerForBeforeSpace(&me);
    }

    // The following loop generates new lines for each line we backed
    // up over and for lines directly affected by edit
    while(cchEdit > 0)
    {
        LONG iTLine;
        LONG cpTemp;
        pliNew = RecalcLinePtr.AddLine(); // Add one new line
        if(!pliNew)
        {
            Assert(FALSE);
            goto err;
        }

        // Stuff text into new line
        uiFlags &= ~MEASURE_FIRSTINPARA;
        uiFlags |= (fFirstInPara ? MEASURE_FIRSTINPARA : 0);

        // measure can add lines for aligned sites
        iNewLine = iOldLine + aryNewLines.Count() - 1;

        // save the index of the new line to added
        iTLine = iNewLine;

#ifdef _DEBUG
        {
            // Just so we can see the damn text.
            const TCHAR* pch;
            long cchInStory = (long)GetFlowLayout()->GetContentTextLength();
            long cp = (long)me.GetCp();
            long cchRemaining =  cchInStory - cp;
            CTxtPtr tp(GetMarkup(), cp);
            pch = tp.GetPch(cchRemaining);
#endif
            cpTemp = me.GetCp();

            if(cpTemp == pFlowLayout->GetContentLastCp()-1)
            {
                uiFlags |= MEASURE_EMPTYLASTLINE;
            }

            // Cache the previous dispnode to insert a new content node
            // if the next line measured has negative margins(offset) which
            // causes the next line on top of the previous lines.
            pDNBefore = me._pDispNodePrev;
            iLinePrev = iNewLine;
            yHeightPrev = yHeight;

            if(!RecalcLinePtr.MeasureLine(me, uiFlags,
                &iNewLine, &yHeight, &yAlignDescent, NULL, NULL))
            {
                goto err;
            }

            // iNewLine returned is the last text line inserted. There ay be
            // aligned lines and clear lines added before and after the
            // text line
            pliNew = iNewLine>=0 ? RecalcLinePtr[iNewLine] : NULL;

            fFirstInPara = (iNewLine>=0)
                ? pliNew->IsNextLineFirstInPara() : TRUE;
#ifdef _DEBUG
        }
#endif
        // fix for bug 16519 (srinib)
        // Keep track of the line that contributes to the max height, ex: if the
        // last lines top margin is negative.
        if(yHeightDisplay < yHeight)
        {
            yHeightDisplay = yHeight;
        }

        if(iNewLine >= 0)
        {
            HandleNegativelyPositionedContent(pliNew, &me, pDNBefore, iLinePrev, yHeightPrev);
        }

        // If we have added any clear lines, iNewLine points to a textline
        // which could be < iTLine
        if(iNewLine >= iTLine)
        {
            cchEdit -= me.GetCp() - cpTemp;
        }

        if(cchEdit <= 0)
        {
            break;
        }

        --iMinimumLinesToRecalc;
        if(iMinimumLinesToRecalc<0 &&
            fBackground && (GetTickCount()>=BgndTickMax())) // GetQueueStatus(QS_ALLEVENTS))
        {
            fDone = FALSE; // took too long, stop for now
            goto no_match;
        }

        if(fAllowBackgroundRecalc
            && yHeight>YWait()
            && (LONG) me.GetCp()>CPWait()
            && cliWait--<=0)
        {
            // Not really done, just past region we're waiting for
            // so let background recalc take it from here
            fDone = FALSE;
            goto no_match;
        }
    }

    // Edit lines have been exhausted.  Continue breaking lines,
    // but try to match new & old breaks
    while((LONG)me.GetCp() < cpLast)
    {
        // Assume there are no matches to try for
        BOOL frpOldValid = FALSE;
        BOOL fChangedLineSpacing = FALSE;

        // If we run out of runs, then there is not match possible. Therefore,
        // we only try for a match as long as we have runs.
        if(fTryForMatch)
        {
            // We are trying for a match so assume that there
            // is a match after all
            frpOldValid = TRUE;

            // Look for match in old line break CArray
            lT = me.GetCp() - cchNew + cchOld;
            while(rpOld.IsValid() && pled->_cpMatchOld<lT)
            {
                if(rpOld->_fForceNewLine)
                {
                    pled->_yMatchOld += rpOld->_yHeight;
                }
                pled->_cpMatchOld += rpOld->_cch;

                BOOL fDone=FALSE;
                BOOL fSuccess = TRUE;
                while(!fDone)
                {
                    if(!rpOld.NextLine(FALSE,FALSE)) // NextRun()
                    {
                        // No more line array entries so we can give up on
                        // trying to match for good.
                        fTryForMatch = FALSE;
                        frpOldValid = FALSE;
                        fDone = TRUE;
                        fSuccess = FALSE;
                    }
                    if(!rpOld->IsFrame() ||
                        (rpOld->IsFrame() && rpOld->_fFrameBeforeText))
                    {
                        fDone = TRUE;
                    }
                }
                if (!fSuccess)
                {
                    break;
                }
            }
        }

        // skip frame in old line array
        if(rpOld->IsFrame() && !rpOld->_fFrameBeforeText)
        {
            BOOL fDone=FALSE;
            while(!fDone)
            {
                if(!rpOld.NextLine(FALSE,FALSE))
                {
                    // No more line array entries so we can give up on
                    // trying to match for good.
                    fTryForMatch = FALSE;
                    frpOldValid = FALSE;
                    fDone = TRUE;
                }
                if(!rpOld->IsFrame())
                {
                    fDone = TRUE;
                }
            }
        }

        pliNew = aryNewLines.Count()>0 ? aryNewLines.Elem(aryNewLines.Count()-1) : NULL;

        // If perfect match, stop
        if(frpOldValid
            && rpOld.IsValid()
            && pled->_cpMatchOld==lT
            && rpOld->_cch!=0
            && (!rpOld->_fPartOfRelChunk // lines do not match if they are a part of
            || rpOld->_fFirstFragInLine  // the previous relchunk (bug 48513 SujalP)
            )
            && pliNew
            && pliNew->_fForceNewLine
            && (yHeight-pliNew->_yHeight>RecalcLinePtr._marginInfo._yBottomLeftMargin)
            && (yHeight-pliNew->_yHeight>RecalcLinePtr._marginInfo._yBottomRightMargin)
            )
        {
            BOOL fFoundMatch = TRUE;

            if(rpOld->_xLeftMargin || rpOld->_xRightMargin)
            {
                // we might have found a match based on cp, but if an
                // aligned site is resized to a smaller size. There might
                // be lines that used to be aligned to the aligned site
                // which are not longer aligned to it. If so, recalc those lines.
                RecalcLinePtr.RecalcMargins(iOldLine, iOldLine+aryNewLines.Count(), yHeight,
                    rpOld->_yBeforeSpace);
                if(RecalcLinePtr._marginInfo._xLeftMargin!=rpOld->_xLeftMargin ||
                    (rpOld->_xLeftMargin + rpOld->_xLineWidth+
                    RecalcLinePtr._marginInfo._xRightMargin)<pFlowLayout->GetMaxLineWidth())
                {
                    fFoundMatch = FALSE;
                }
            }

            // There are cases where we've matched characters, but want to continue
            // to recalc anyways. This requires us to instantiate a new measurer.
            if(fFoundMatch && !fChangedLineSpacing && rpOld<LineCount() &&
                (rpOld->_fFirstInPara||pliNew->_fHasEOP))
            {
                BOOL fInner;
                const CParaFormat* pPF;
                CLSMeasurer tme(this, me.GetCp(), pci);

                if(!tme._pLS)
                {
                    goto err;
                }

                if(pliNew->_fHasEOP)
                {
                    rpOld->_fFirstInPara = TRUE;
                }
                else
                {
                    rpOld->_fFirstInPara = FALSE;
                    rpOld->_fHasBulletOrNum = FALSE;
                }

                // If we've got a <DL COMPACT> <DT> line. For now just check to see if
                // we're under the DL.
                pPF = tme.CurrBranch()->GetParaFormat();

                fInner = SameScope(tme.CurrBranch(), pElementContent);

                if(pPF->HasCompactDL(fInner))
                {
                    // If the line is a DT and it's the end of the paragraph, and the COMPACT
                    // attribute is set.
                    fChangedLineSpacing = TRUE;
                }
                // Check to see if the line height is the same. This is necessary
                // because there are a number of block operations that can
                // change the prespacing of the next line.
                else
                {
                    // We'd better not be looking at a frame here.
                    Assert(!rpOld->IsFrame());

                    // Make it possible to check the line space of the
                    // current line.
                    tme.InitLineSpace(&me, rpOld);

                    RecalcLinePtr.CalcInterParaSpace(&tme,
                        pled->_iliFirst+aryNewLines.Count()-1, FALSE);

                    if(rpOld->_yBeforeSpace!=tme._li._yBeforeSpace ||
                        rpOld->_fHasBulletOrNum!=tme._li._fHasBulletOrNum)
                    {
                        rpOld->_fHasBulletOrNum = tme._li._fHasBulletOrNum;

                        fChangedLineSpacing = TRUE;
                    }
                }
            }
            else
            {
                fChangedLineSpacing = FALSE;
            }

            if(fFoundMatch && !fChangedLineSpacing)
            {
                pled->_iliMatchOld = rpOld;

                // Replace old lines by new ones
                lT = rpOld - pled->_iliFirst;
                rpOld = pled->_iliFirst;

                fReplaceResult = rpOld.Replace(lT, &aryNewLines);
                if (!fReplaceResult)
                {
                    Assert(FALSE);
                    goto err;
                }

                frpOldValid =
                    rpOld.SetRun(rpOld.GetIRun()+aryNewLines.Count(), 0);

                aryNewLines.Clear(AF_DELETEMEM); // Clear aux array
                iOldLine = rpOld;

                // Remember information about match after editing
                pled->_yMatchNew = yHeight;
                pled->_iliMatchNew = rpOld;
                pled->_cpMatchNew = me.GetCp();

                _dcpCalcMax = me.GetCp() - cpLayoutFirst;

                // Compute height and cp after all matches
                if(frpOldValid && rpOld.IsValid())
                {
                    do
                    {
                        if(rpOld->_fForceNewLine)
                        {
                            yHeight += rpOld->_yHeight;
                            // fix for bug 16519
                            // Keep track of the line that contributes to the max height, ex: if the
                            // last lines top margin is negative.
                            if(yHeightDisplay < yHeight)
                            {
                                yHeightDisplay = yHeight;
                            }
                        }
                        else if(rpOld->IsFrame())
                        {
                            yAlignDescent = yHeight + rpOld->_yHeight;
                        }

                        _dcpCalcMax += rpOld->_cch;
                    } while(rpOld.NextLine(FALSE,FALSE)); // NextRun()
                }

                // Make sure _dcpCalcMax is sane after the above update
                AssertSz(GetMaxCpCalced()<=cpLast,
                    "CDisplay::RecalcLines match extends beyond EOF");

                // We stop calculating here.Note that if _dcpCalcMax < size
                // of text, this means a background recalc is in progress.
                // We will let that background recalc get the arrays
                // fully in sync.
                AssertSz(GetMaxCpCalced()<=cpLast,
                    "CDisplay::Match less but no background recalc");

                if(GetMaxCpCalced() != cpLast)
                {
                    // This is going to be finished by the background recalc
                    // so set the done flag appropriately.
                    fDone = FALSE;
                }

                goto match;
            }
        }

        // Add a new line
        pliNew = RecalcLinePtr.AddLine();
        if(!pliNew)
        {
            Assert(FALSE);
            goto err;
        }

        // measure can add lines for aligned sites
        iNewLine = iOldLine + aryNewLines.Count() - 1;

        uiFlags = MEASURE_BREAKATWORD |
            (yHeight==0?MEASURE_FIRSTLINE:0) |
            (fFirstInPara?MEASURE_FIRSTINPARA:0);

        if(GetWordWrap() && GetWrapLongLines())
        {
            uiFlags |= MEASURE_BREAKLONGLINES;
        }

        if(long(me.GetCp()) == pFlowLayout->GetContentLastCp()-1)
        {
            uiFlags |= MEASURE_EMPTYLASTLINE;
        }

        // Cache the previous dispnode to insert a new content node
        // if the next line measured has negative margins(offset) which
        // causes the next line on top of the previous lines.
        pDNBefore = me._pDispNodePrev;
        iLinePrev = iNewLine;
        yHeightPrev = yHeight;

        if(!RecalcLinePtr.MeasureLine(me, uiFlags, &iNewLine, &yHeight, &yAlignDescent, NULL, NULL))
        {
            goto err;
        }

        // fix for bug 16519
        // Keep track of the line that contributes to the max height, ex: if the
        // last lines top margin is negative.
        if(yHeightDisplay < yHeight)
        {
            yHeightDisplay = yHeight;
        }

        // iNewLine returned is the last text line inserted. There ay be
        // aligned lines and clear lines added before and after the
        // text line
        pliNew = iNewLine>=0 ? RecalcLinePtr[iNewLine] : NULL;

        fFirstInPara = (iNewLine>=0) ? pliNew->IsNextLineFirstInPara() : TRUE;
        if(iNewLine >= 0)
        {
            HandleNegativelyPositionedContent(pliNew, &me, pDNBefore, iLinePrev, yHeightPrev);
        }

        --iMinimumLinesToRecalc;
        if(iMinimumLinesToRecalc<0 &&
            fBackground && (GetTickCount()>=(DWORD)BgndTickMax())) // GetQueueStatus(QS_ALLEVENTS))
        {
            fDone = FALSE; // Took too long, stop for now
            break;
        }

        if(fAllowBackgroundRecalc && yHeight>YWait()
            && (LONG) me.GetCp()>CPWait() && cliWait--<=0)
        {                           // Not really done, just past region we're
            fDone = FALSE;          //  waiting for so let background recalc
            break;                  //  take it from here
        }
    } // while(me < cpLast ...

no_match:
    // Didn't find match: whole line array from _iliFirst needs to be changed
    // Indicate no match
    pled->SetNoMatch();

    // Update old match position with previous total height of the document so
    // that UpdateView can decide whether to invalidate past the end of the
    // document or not
    pled->_iliMatchOld = LineCount();
    pled->_cpMatchOld  = cpLast + cchOld - cchNew;
    pled->_yMatchOld   = _yHeightMax;

    // Update last recalced cp
    _dcpCalcMax = me.GetCp() - cpLayoutFirst;

    // Replace old lines by new ones
    rpOld = pled->_iliFirst;

    // We store the result from the replace because although it can fail the
    // fields used for first visible must be set to something sensible whether
    // the replace fails or not. Further, the setting up of the first visible
    // fields must happen after the Replace because the lines could have
    // changed in length which in turns means that the first visible position
    // has failed.
    fReplaceResult = rpOld.Replace(-1, &aryNewLines);
    if(!fReplaceResult)
    {
        Assert(FALSE);
        goto err;
    }

    // Adjust first affected line if this line is gone
    // after replacing by new lines
    if(pled->_iliFirst>=LineCount() && LineCount()>0)
    {
        Assert(pled->_iliFirst == LineCount());
        pled->_iliFirst = LineCount() - 1;
        Assert(!Elem(pled->_iliFirst)->IsFrame());
        pled->_yFirst -= Elem(pled->_iliFirst)->_yHeight;

        // See comment before as to why its legal for this to be possible
        //
        //AssertSz(pled->_yFirst >= 0, "CDisplay::RecalcLines _yFirst < 0");
        pled->_cpFirst -= Elem(pled->_iliFirst)->_cch;
    }

match:
    // If there is a background on the paragraph, we have to make sure to redraw the
    // lines to the end of the paragraph.
    for(; pled->_iliMatchNew < LineCount() ;)
    {
        pliNew = Elem(pled->_iliMatchNew);
        if(pliNew)
        {
            if(!pliNew->_fHasBackground)
            {
                break;
            }

            pled->_iliMatchOld++;
            pled->_iliMatchNew++;
            pled->_cpMatchOld += pliNew->_cch;
            pled->_cpMatchNew += pliNew->_cch;
            me.Advance(pliNew->_cch);
            if(pliNew->_fForceNewLine)
            {
                pled->_yMatchOld +=  pliNew->_yHeight;
                pled->_yMatchNew +=  pliNew->_yHeight;
            }
            if(pliNew->_fHasEOP)
            {
                break;
            }
        }
        else
        {
            break;
        }
    }

    _fRecalcDone = fDone;
    _yCalcMax = yHeightDisplay;

    if(HasBgRecalcInfo())
    {
        CBgRecalcInfo* pBgRecalcInfo = BgRecalcInfo();
        // Clear wait fields since we want caller's to set them up.
        pBgRecalcInfo->_yWait = -1;
        pBgRecalcInfo->_cpWait = -1;
    }

    // We've measured the last line in the document. Do we want an empty lne?
    if((LONG)me.GetCp() == cpLast)
    {
        LONG ili = LineCount() - 1;
        long lPadding[PADDING_MAX];

        // Update the padding once we measure the last line
        GetPadding(pci, lPadding);
        _yBottomMargin = lPadding[PADDING_BOTTOM];

        // If we haven't measured any lines (deleted an empty paragraph),
        // we need to get a pointer to the last line rather than using the
        // last line measured.
        while(ili >= 0)
        {
            pliNew = Elem(ili);
            if(pliNew->IsTextLine())
            {
                break;
            }
            else
            {
                ili--;
            }
        }

        // If the last line has a paragraph break or we don't have any
        // line any more, we need to create a empty line only if we are in design mode
        if(LineCount()==0
            || (GetFlowLayout()->IsEditable()
            && pliNew
            && (pliNew->_fHasEOP
            || pliNew->_fHasBreak
            || pliNew->_fSingleSite)))
        {
            RecalcLinePtr.Init((CLineArray*)this, 0, NULL);
            CreateEmptyLine(&me, &RecalcLinePtr, &yHeight,
                pliNew?pliNew->_fHasEOP:TRUE);
            // fix for bug 16519
            // Keep track of the line that contributes to the max height, ex: if the
            // last lines top margin is negative.
            if(yHeightDisplay < yHeight)
            {
                yHeightDisplay = yHeight;
            }
        }
        // In table cells, Netscape actually adds the interparagraph spacing
        // for any closed tags to the height of the table cell.
        // BUGBUG: This actually applies the spacing to all nested text containers, not just
        //         table cells. Is this correct? (brendand)
        // It's not correct to add the spacing at all, but given that Netscape
        // has done so, and that they will probably do so for floating block
        // elements. Arye.
        else
        {
            int iLineLast = ili;

            // we need to force scrolling when bottom-margin is set on the last block tag
            // in body. (bug 20400)

            // Only do this if we're the last line in the text site.
            // This means that the last line is a text line.
            if(iLineLast == LineCount()-1)
            {
                while(iLineLast-->0 && // Lines available
                    !pliNew->_fForceNewLine) // Just a dummy line.
                {
                    --pliNew;
                }
            }
            if(iLineLast>0 || pliNew->_fForceNewLine)
            {
                _yBottomMargin += RecalcLinePtr.NetscapeBottomMargin(&me);
            }
        }
    }

    if(fDone)
    {
        RecalcMost();

        if(fBackground)
        {
            StopBackgroundRecalc();
        }
    }

    xWidth = CalcDisplayWidth();
    yHeightMax = max(yHeightDisplay, yAlignDescent);

    Assert (pled);

    // If the size changed, re-size or update scrollbars as appropriate
    //
    // Fire a resize notification if the content size changed
    //  
    // fix for 52305, change to margin bottom affects the content size,
    // so we should fire a resize.
    if(yHeightMax+yBottomMarginOld!=_yHeightMax+_yBottomMargin
        || yHeightDisplay!=_yHeight
        || xWidth!=_xWidth)
    {
        _xWidth       = xWidth;
        _yHeight      = yHeightDisplay;
        _yHeightMax   = yHeightMax;

        pFlowLayout->SizeContentDispNode(CSize(GetMaxWidth(), GetHeight()));

        // If our contents affects our size, ask our parent to initiate a re-size
        ElementResize(pFlowLayout);
    }
    else if(_fHasMultipleTextNodes)
    {
        pFlowLayout->SizeContentDispNode(CSize(GetMaxWidth(), GetHeight()));
    }

    // Update the relative line cache.
    if(pFlowLayout->_fContainsRelative)
    {
        UpdateRelDispNodeCache(pled);
    }

    AdjustDispNodes(me._pDispNodePrev, pled);

    pFlowLayout->NotifyMeasuredRange(pled->_cpFirst, me.GetCp());

    if(pled->_cpMatchNew!=MAXLONG && (pled->_yMatchNew - pled->_yMatchOld))
    {
        CSize size(0, pled->_yMatchNew - pled->_yMatchOld);

        pFlowLayout->NotifyTranslatedRange(size, pled->_cpMatchNew, cpLayoutFirst+_dcpCalcMax);
    }

    // If not done, do the rest in background
    if(!fDone && !fBackground)
    {
        StartBackgroundRecalc(pci->_grfLayout);
    }

    if(fDone)
    {
        _fLineRecalcErr = FALSE;
    }

cleanup:
    return TRUE;

err:
    if(!_fLineRecalcErr)
    {
        _dcpCalcMax = me.GetCp() - cpLayoutFirst;
        _yCalcMax   = yHeightDisplay;
        _fLineRecalcErr = FALSE; //  fix up CArray & bail
    }

    Trace0("CDisplay::RecalcLines() failed\n");

    pled->SetNoMatch();

    if(!fReplaceResult)
    {
        FlushRecalc();
    }

    // Update the relative line cache.
    if(pFlowLayout->_fContainsRelative)
    {
        UpdateRelDispNodeCache(pled);
    }

    AdjustDispNodes(me._pDispNodePrev, pled);

    return FALSE;
}

/*
*  CDisplay::UpdateView(&tpFirst, cchOld, cchNew, fRecalc)
*
*  @mfunc
*      Recalc lines and update the visible part of the display
*      (the "view") on the screen.
*
*  @devnote
*      --- Use when in-place active only ---
*
*  @rdesc
*      TRUE if success
*/
BOOL CDisplay::UpdateView(
                          CCalcInfo* pci,
                          LONG cpFirst, //@parm Text ptr where change happened
                          LONG cchOld,  //@parm Count of chars deleted
                          LONG cchNew)  //@parm No recalc need (only rendering change) = FALSE
{
    BOOL            fReturn = TRUE;
    BOOL            fNeedViewChange = FALSE;
    RECT            rcView;
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CRect           rc;
    long            xWidthParent;
    long            yHeightParent;

    BOOL fInvalAll = FALSE;
    CStackDataAry<RECT, 10> aryInvalRects;

    Assert(pci);

    if(_fNoUpdateView)
    {
        return fReturn;
    }

    // If we have no width, don't even try to recalc, it causes
    // crashes in the scrollbar code, and is just a waste of
    // time, anyway.
    // However, since we're not sized, request sizing from the parent
    // of the ped (this will leave the necessary bread-crumb chain
    // to ensure we get sized later) - without this, not all containers
    // of text sites (e.g., 2D sites) will know to size the ped.
    // (This is easier than attempting to propagate back out that sizing
    //  did not occur.)
    if(GetViewWidth()<=0 && !pFlowLayout->_fContentsAffectSize)
    {
        FlushRecalc();
        return TRUE;
    }


    pFlowLayout->GetClientRect((CRect*)&rcView);
    GetViewWidthAndHeightForChild(pci, &xWidthParent, &yHeightParent);

    // BUGBUG(SujalP, SriniB and BrendanD): These 2 lines are really needed here
    // and in all places where we instantiate a CI. However, its expensive right
    // now to add these lines at all the places and hence we are removing them
    // from here right now. We have also removed them from CTxtSite::CalcTextSize()
    // for the same reason. (Bug#s: 58809, 62517, 62977)
    pci->SizeToParent(xWidthParent, yHeightParent);

    // If we get here, we are updating some general characteristic of the
    // display and so we want the cursor updated as well as the general
    // change otherwise the cursor will land up in the wrong place.
    _fUpdateCaret = TRUE;

    if(!LineCount())
    {
        // No line yet, recalc everything
        RecalcView(pci, TRUE);

        // Invalidate entire view rect
        fInvalAll = TRUE;
        fNeedViewChange = TRUE;
    }
    else
    {
        CLed led;

        if(cpFirst > GetMaxCpCalced())
        {
            // we haven't even calc'ed this far, so don't bother with updating
            // here.  Background recalc will eventually catch up to us.
            return TRUE;
        }

        AssertSz(cpFirst<=GetMaxCpCalced(),
            "CDisplay::UpdateView(...) - cpFirst > GetMaxCpCalced()");

        if(!RecalcLines(pci, cpFirst, cchOld, cchNew, FALSE, &led))
        {
            // we're in deep crap now, the recalc failed
            // let's try to get out of this with our head still mostly attached
            fReturn = FALSE;
            fInvalAll = TRUE;
            fNeedViewChange = TRUE;
        }

        if(!fInvalAll)
        {
            // Invalidate all the individual rectangles necessary to work around
            // any embedded images. Also, remember if this was a simple or a complex
            // operation so that we can avoid scrolling in difficult places.
            CLine*  pLine;
            int     iLine, iLineLast;
            long    xLineLeft, xLineRight, yLineTop, yLineBottom;
            long    yTop;
            long    dy = led._yMatchNew - led._yMatchOld;

            // Start out with a zero area rectangle.
            // NOTE(SujalP): _yFirst can legally be less than 0. Its OK, because
            // the rect we are constructing here is going to be used for inval purposes
            // only and that is going to clip it to the content rect of the flowlayout.
            yTop = rc.bottom = rc.top = led._yFirst;
            rc.left = MAXLONG;
            rc.right = 0;

            // Need this to adjust for negative line heights.
            // Note that this also initializes the counter for the
            // for loop just below.
            iLine = led._iliFirst;

            // Accumulate rectangles of lines and invalidate them.
            iLineLast = min(led._iliMatchNew, LineCount());

            // Invalidate only the lines that have been touched by the recalc
            for(; iLine < iLineLast; iLine++)
            {
                pLine = Elem(iLine);

                if(pLine == NULL)
                {
                    break;
                }

                // Get the left and right edges of the line.
                if(!_fRTL)
                {
                    xLineLeft  = pLine->_xLeftMargin;
                    xLineRight = pLine->_xLeftMargin + pLine->_xLineWidth;
                }
                else
                {
                    xLineLeft  = -(pLine->_xRightMargin + pLine->_xLineWidth);
                    xLineRight = -(pLine->_xRightMargin);
                }

                // Get the top and bottom edges of the line
                yLineTop    = yTop + pLine->GetYLineTop();
                yLineBottom = yTop + pLine->GetYLineBottom();

                // First line of a new rectangle
                if(rc.right < rc.left)
                {
                    rc.left  = xLineLeft;
                    rc.right = xLineRight;
                }

                // Different margins, output the old one and restart.
                else if(rc.left!=xLineLeft || rc.right!=xLineRight)
                {
                    // if we have multiple chunks in the same line
                    if(rc.right==xLineLeft && rc.top==yLineTop && rc.bottom==yLineBottom)
                    {
                        rc.right = xLineRight;
                    }
                    else
                    {
                        aryInvalRects.AppendIndirect(&rc);

                        fNeedViewChange = TRUE;

                        // Zero height.
                        rc.top = rc.bottom = yTop;

                        // Just the width of the given line.
                        rc.left = xLineLeft;
                        rc.right = xLineRight;
                    }
                }

                // Negative line height.
                rc.top = min(rc.top, yLineTop);

                rc.bottom = max(rc.bottom, yLineBottom);

                // Otherwise just accumulate the height.
                if(pLine->_fForceNewLine)
                {
                    yTop += pLine->_yHeight;

                    // Stop when reaching the bottom of the visible area
                    if(rc.bottom > rcView.bottom)
                    {
                        break;
                    }
                }
            }

            // BUBUG (srinib) - This is a temporary hack to handle the
            // scrolling case until the display tree provides the functionality
            // to scroll an rc in view. If the new content height changed then
            // scroll the content below the change by the dy. For now we are just
            // to the end of the view.
            if(dy)
            {
                rc.left   = rcView.left;
                rc.right  = rcView.right;
                rc.bottom = rcView.bottom;
            }

            // Get the last one.
            if(rc.right>rc.left && rc.bottom>rc.top)
            {
                aryInvalRects.AppendIndirect(&rc);
                fNeedViewChange = TRUE;
            }

            // There might be more stuff which has to be
            // invalidated because of lists (numbers might change etc)
            if(UpdateViewForLists(&rcView, cpFirst, iLineLast, rc.bottom, &rc))
            {
                aryInvalRects.AppendIndirect(&rc);
                fNeedViewChange = TRUE;
            }

            if(led._yFirst>=rcView.top &&
                (led._yMatchNew<=rcView.bottom || led._yMatchOld<=rcView.bottom))
            {
                WaitForRecalcView(pci);
            }
        }
    }

    {
        // Now do the invals
        if(fInvalAll)
        {
            pFlowLayout->Invalidate();
        }
        else
        {
            pFlowLayout->Invalidate(&aryInvalRects[0], aryInvalRects.Size());
        }
    }

    if(fNeedViewChange)
    {
        pFlowLayout->ViewChange(FALSE);
    }

    return fReturn;
}

/*
*  CDisplay::CalcDisplayWidth()
*
*  @mfunc
*      Computes and returns width of this display by walking line
*      CArray and returning the widest line.  Used for
*      horizontal scroll bar routines
*
*  @todo
*      This should be done by a background thread
*/
LONG CDisplay::CalcDisplayWidth()
{
    LONG    ili = LineCount();
    CLine*  pli;
    LONG    xWidth = 0;

    if(ili)
    {
        // Note: pli++ breaks array encapsulation (pli = &(*this)[ili] doesn't)
        pli = Elem(0);
        for(xWidth=0; ili--; pli++)
        {
            // calc line width based on the direction
            if(!_fRTL)
            {
                xWidth = max(xWidth, pli->GetTextRight(ili==0)+pli->_xRight);
            }
            else
            {
                xWidth = max(xWidth, pli->GetRTLTextLeft()+pli->_xLeft);
            }
        }
    }

    return xWidth;
}

/*
*  CDisplay::StartBackgroundRecalc()
*
*  @mfunc
*      Starts background line recalc (at _dcpCalcMax position)
*
*  @todo
*      ??? CF - Should use an idle thread
*/
void CDisplay::StartBackgroundRecalc(DWORD grfLayout)
{
    // We better be in the middle of the job here.
    Assert(LineCount() > 0);

    Assert(CanHaveBgRecalcInfo());

    // StopBack.. kills the recalc task, *if it exists*
    StopBackgroundRecalc();

    EnsureBgRecalcInfo();

    if(!RecalcTask() && GetMaxCpCalced()<GetLastCp())
    {
        BgRecalcInfo()->_pRecalcTask = new CRecalcTask(this, grfLayout) ;
        if(RecalcTask())
        {
            _fRecalcDone = FALSE;
        }
        // BUGBUG (sujalp): what to do if we fail on mem allocation?
        // One solution is to create this task as blocked when we
        // construct CDisplay, and then just unblock it here. In
        // StopBackgroundRecalc just block it. In CDisplay destructor,
        // destruct the CTask too.
    }
}

/*
*  CDisplay::StepBackgroundRecalc()
*
*  @mfunc
*      Steps background line recalc (at _dcpCalcMax position)
*      Called by timer proc
*
*  @todo
*      ??? CF - Should use an idle thread
*/
VOID CDisplay::StepBackgroundRecalc(DWORD dwTimeOut, DWORD grfLayout)
{
    LONG cch = GetLastCp() - (GetMaxCpCalced());

    // don't try recalc when processing OOM or had an error doing recalc or
    // if we are asserting.
    if(_fInBkgndRecalc || _fLineRecalcErr)
    {
        return;
    }

    _fInBkgndRecalc = TRUE;

    // Background recalc is over if no more characters or we are no longer
    // active.
    if(cch <= 0)
    {
        if(RecalcTask())
        {
            StopBackgroundRecalc();
        }

        goto Cleanup;
    }

    {
        CFlowLayout*    pFlowLayout = GetFlowLayout();
        CElement*       pElementFL = pFlowLayout->ElementOwner();
        LONG            cp = GetMaxCpCalced();

        if(!pElementFL->IsDisplayNone() || pElementFL->Tag()==ETAG_BODY)
        {
            CCalcInfo   CI;
            CLed        led;
            long        xParentWidth;
            long        yParentHeight;

            pFlowLayout->OpenView();

            // Setup the amount of time we have this time around
            Assert(BgRecalcInfo() && "Supposed to have a recalc info in stepbackgroundrecalc");
            BgRecalcInfo()->_dwBgndTickMax = dwTimeOut;

            CI.Init(pFlowLayout);
            GetViewWidthAndHeightForChild(
                &CI,
                &xParentWidth,
                &yParentHeight,
                CI._smMode==SIZEMODE_MMWIDTH);
            CI.SizeToParent(xParentWidth, yParentHeight);

            CI._grfLayout = grfLayout;

            RecalcLines(&CI, cp, cch, cch, TRUE, &led);
        }
        else
        {
            CNotification nf;

            // Kill background recalc, if the layout is hidden
            StopBackgroundRecalc();

            // BUGBUG (MohanB) Need to use ElementContent()?
            // calc the rest by accumulating a dirty range.
            nf.CharsResize(GetMaxCpCalced(), cch, pElementFL->GetFirstBranch());
            GetMarkup()->Notify(&nf);
        }
    }

Cleanup:
    _fInBkgndRecalc = FALSE;

    return;
}

/*
*  CDisplay::StopBackgroundRecalc()
*
*  @mfunc
*      Steps background line recalc (at _dcpCalcMax position)
*      Called by timer proc
*
*/
VOID CDisplay::StopBackgroundRecalc()
{
    if(HasBgRecalcInfo())
    {
        if(RecalcTask())
        {
            RecalcTask()->Terminate () ;
            RecalcTask()->Release () ;
            _fRecalcDone = TRUE;
        }
        DeleteBgRecalcInfo();
    }
}

/*
*  CDisplay::WaitForRecalc(cpMax, yMax, pDI)
*
*  @mfunc
*      Ensures that lines are recalced until a specific character
*      position or ypos.
*
*  @rdesc
*      success
*/
BOOL CDisplay::WaitForRecalc(
                             LONG cpMax,    //@parm Position recalc up to (-1 to ignore)
                             LONG yMax,     //@parm ypos to recalc up to (-1 to ignore)
                             CCalcInfo* pci)

{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    BOOL            fReturn = TRUE;
    LONG            cch;
    CCalcInfo       CI;

    Assert(cpMax<0 || (cpMax>=GetFirstCp() && cpMax<=GetLastCp()));

    // Return immediately if hidden, already measured up to the correct point, or currently measuring
    // or if there is no dispnode (ie haven't been CalcSize'd yet)
    if(pFlowLayout->IsDisplayNone() && pFlowLayout->Tag()!=ETAG_BODY)
    {
        return fReturn;
    }

    if(!pFlowLayout->GetElementDispNode())
    {
        return FALSE;
    }

    if(yMax<0
        && cpMax>=0
        && ((!pFlowLayout->IsDirty()
        && cpMax<=GetMaxCpCalced())
        || (pFlowLayout->IsDirty()
        && pFlowLayout->IsRangeBeforeDirty(cpMax-pFlowLayout->GetContentFirstCp(), 0))))
    {
        return fReturn;
    }

    if(pFlowLayout->TestLock(CElement::ELEMENTLOCK_RECALC)
        || pFlowLayout->TestLock(CElement::ELEMENTLOCK_PROCESSREQUESTS))
    {
        return FALSE;
    }

    // Calculate up through the request location
    if(!pci)
    {
        CI.Init(pFlowLayout);
        pci = &CI;
    }

    pFlowLayout->CommitChanges(pci);

    if((yMax<0 || yMax>=_yCalcMax)
        && (cpMax<0 || cpMax>GetMaxCpCalced()))
    {
        cch = GetLastCp() - GetMaxCpCalced();
        if(cch>0 || LineCount()==0)
        {

            HCURSOR hcur = NULL;
            CDocument* pDoc = pFlowLayout->Doc();

            if(EnsureBgRecalcInfo() == S_OK)
            {
                CBgRecalcInfo* pBgRecalcInfo = BgRecalcInfo();
                Assert(pBgRecalcInfo && "Should have a BgRecalcInfo");
                pBgRecalcInfo->_cpWait = cpMax;
                pBgRecalcInfo->_yWait  = yMax;
            }

            if(pDoc/* && pDoc->State()>=OS_INPLACE wlw note*/)
            {
                hcur = SetCursorIDC(IDC_WAIT);
            }
            Trace0("Lazy recalc\n");

            if(GetMaxCpCalced() == GetFirstCp() )
            {
                fReturn = RecalcLines(pci);
            }
            else
            {
                CLed led;

                fReturn = RecalcLines(pci, GetMaxCpCalced(), cch, cch, FALSE, &led, TRUE);
            }

            // Either we were not waiting for a cp or if we were, then we have been calcd to that cp
            Assert(cpMax<0 || GetMaxCpCalced()>=cpMax);

            SetCursor(hcur);
        }
    }

    return fReturn;
}

/*
*  CDisplay::WaitForRecalcIli
*
*  @mfunc
*      Wait until line array is recalculated up to line <p ili>
*
*  @rdesc
*      Returns TRUE if lines were recalc'd up to ili
*/
BOOL CDisplay::WaitForRecalcIli(
                                LONG ili,       //@parm Line index to recalculate line array up to
                                CCalcInfo* pci)
{
    return ili<LineCount();

    LONG cchGuess;

    while(!_fRecalcDone && ili>=LineCount())
    {
        cchGuess = 5 * (ili-LineCount()+1) * Elem(0)->_cch;
        if(!WaitForRecalc(GetMaxCpCalced()+cchGuess, -1, pci))
        {
            return FALSE;
        }
    }
    return ili<LineCount();
}

//+----------------------------------------------------------------------------
//
//  Member:     WaitForRecalcView
//
//  Synopsis:   Calculate up through the bottom of the visible content
//
//  Arguments:  pci - CCalcInfo to use
//
//  Returns:    TRUE if all Ok, FALSE otherwise
//
//-----------------------------------------------------------------------------
BOOL CDisplay::WaitForRecalcView(CCalcInfo* pci)
{
    return WaitForRecalc(-1, GetFlowLayout()->GetClientBottom(), pci);
}

//===================================  View Updating  ===================================
/*
*  CDisplay::SetViewSize(rcView)
*
*  Purpose:
*      Set the view size and return whether a full recalc is needed
*
*  Returns:
*      TRUE if a full recalc is needed
*/
BOOL CDisplay::SetViewSize(const RECT& rcView, BOOL fViewUpdate)
{
    BOOL fRecalc = ((rcView.right-rcView.left) != _xWidthView);

    _xWidthView  = rcView.right - rcView.left;
    _yHeightView = rcView.bottom - rcView.top;

    return fRecalc||!LineCount();
}

/*
*  CDisplay::RecalcView
*
*  Sysnopsis:  Recalc view and update first visible line
*
*  Arguments:
*      fFullRecalc - TRUE if recalc from first line needed, FALSE if enough
*                    to start from _dcpCalcMax
*
*  Returns:
*      TRUE if success
*/
BOOL CDisplay::RecalcView(CCalcInfo* pci, BOOL fFullRecalc)
{
    CFlowLayout* pFlowLayout = GetFlowLayout();
    BOOL fAllowBackgroundRecalc;

    // If we have no width, don't even try to recalc, it causes
    // crashes in the scrollbar code, and is just a waste of time
    if(GetViewWidth()<=0 && !pFlowLayout->_fContentsAffectSize)
    {
        FlushRecalc();
        return TRUE;
    }

    fAllowBackgroundRecalc = AllowBackgroundRecalc(pci);

    // If a full recalc (from first line) is not needed
    // go from current _dcpCalcMax on
    if(!fFullRecalc)
    {
        return (fAllowBackgroundRecalc
            ? WaitForRecalcView(pci)
            : WaitForRecalc(GetLastCp(), -1, pci));
    }

    // Else do full recalc
    BOOL fRet = TRUE;

    // If all that the element is likely to have is a single line of plain text,
    // use a faster mechanism to compute the lines. This is great for perf
    // of <INPUT>
    if(pFlowLayout->Tag() == ETAG_INPUT)
    {
        fRet = RecalcPlainTextSingleLine(pci);
    }
    else
    {
        // full recalc lines
        if(!RecalcLines(pci))
        {
            // we're in deep crap now, the recalc failed
            // let's try to get out of this with our head still mostly attached
            fRet = FALSE;
            goto Done;
        }
    }

Done:
    return fRet;
}

void SetLineWidth(long xWidthMax, CLine* pli)
{
    pli->_xLineWidth = max(xWidthMax,
        pli->_xLeft+pli->_xWidth+pli->_xLineOverhang+pli->_xRight);
}

//+---------------------------------------------------------------
//
//  Member:     CDisplay::RecalcLineShift
//
//  Synopsis:   Run thru line array and adjust line shift
//
//---------------------------------------------------------------
void CDisplay::RecalcLineShift(CCalcInfo* pci, DWORD grfLayout)
{
    CFlowLayout* pFlowLayout = GetFlowLayout();
    LONG        lCount = LineCount();
    LONG        ili;
    LONG        iliFirstChunk = 0;
    BOOL        fChunks = FALSE;
    CLine*      pli;
    long        xShift;
    long        xWidthMax = GetMaxPixelWidth();

    Assert(pFlowLayout->_fSizeToContent ||
        (_fMinMaxCalced && !pFlowLayout->ContainsChildLayout()));

    for(ili=0,pli=Elem(0); ili<lCount; ili++,pli++)
    {
        // if the current line does not force a new line, then
        // find a line that forces the new line and distribute
        // width.
        if(!fChunks && !pli->_fForceNewLine && !pli->IsFrame())
        {
            iliFirstChunk = ili;
            fChunks = TRUE;
        }

        if(pli->_fForceNewLine)
        {
            xShift = 0;

            if(!fChunks)
            {
                iliFirstChunk = ili;
            }
            else
            {
                fChunks = FALSE;
            }

            if(pli->_fJustified && long(pli->_fJustified)!=JUSTIFY_FULL)
            {
                // for pre whitespace is already include in _xWidth
                xShift = xWidthMax - pli->_xLeftMargin - pli->_xLeft - pli->_xWidth - pli->_xLineOverhang -
                    pli->_xRight - pli->_xRightMargin - GetCaret();

                xShift = max(xShift, 0L); // Don't allow alignment to go < 0
                // Can happen with a target device
                if(long(pli->_fJustified) == JUSTIFY_CENTER)
                {
                    xShift /= 2;
                }
            }

            while(iliFirstChunk < ili)
            {
                pli = Elem(iliFirstChunk++);
                if(!_fRTL)
                {
                    pli->_xLeft += xShift;
                }
                else
                {
                    pli->_xRight += xShift;
                }

                pli->_xLineWidth = pli->_xLeft + pli->_xWidth +
                    pli->_xLineOverhang + pli->_xRight;
            }

            pli = Elem(iliFirstChunk++);
            if(!_fRTL)
            {
                pli->_xLeft += xShift;
            }
            else
            {
                pli->_xRight += xShift;
            }

            SetLineWidth(xWidthMax, pli);
        }
    }

    if(pFlowLayout->_fContainsRelative)
    {
        VoidRelDispNodeCache();
        UpdateRelDispNodeCache(NULL);
    }

    // Nested layouts need to be repositioned, to account for lineshift.
    if(pFlowLayout->_fSizeToContent)
    {
        RecalcLineShiftForNestedLayouts();
    }

    // NOTE(SujalP): Ideally I would like to do the NMR in all cases. However, that causes a misc perf
    // regession of about 2% on pages with a lot of small table cells. To avoid that problem I am doing
    // this only for edit mode. If you have other needs then add those cases to the if condition.
    if(pFlowLayout->IsEditable())
    {
        pFlowLayout->NotifyMeasuredRange(GetFlowLayoutElement()->GetFirstCp(), GetMaxCpCalced());
    }
}

void CDisplay::RecalcLineShiftForNestedLayouts()
{
    CLayout*        pLayout;
    CFlowLayout*    pFL = GetFlowLayout();
    CDispNode*      pDispNode = pFL->GetFirstContentDispNode();

    if(pFL->_fAutoBelow)
    {
        DWORD_PTR dw;

        for(pLayout=pFL->GetFirstLayout(&dw); pLayout; pLayout=pFL->GetNextLayout(&dw))
        {
            CElement* pElement = pLayout->ElementOwner();
            CTreeNode* pNode = pElement->GetFirstBranch();
            const CFancyFormat* pFF = pNode->GetFancyFormat();

            if(!pFF->IsAligned()
                && (pFF->IsAutoPositioned()
                || pNode->GetCharFormat()->IsRelative(FALSE)))
            {
                pElement->ZChangeElement();
            }
        }
        pFL->ClearLayoutIterator(dw);
    }

    pDispNode = pDispNode ? pDispNode->GetNextSiblingNode(TRUE) : NULL;
    if(pDispNode)
    {
        CLinePtr rp(this);

        do
        {
            // if the current disp node is not a text node
            if(pDispNode->GetDispClient() != pFL)
            {
                void* pvOwner;

                pDispNode->GetDispClient()->GetOwner(pDispNode, &pvOwner);

                if(pvOwner)
                {
                    CElement* pElement = DYNCAST(CElement, (CElement*)pvOwner);

                    // aligned layouts are not affected by text-align
                    if(!pElement->IsAligned())
                    {
                        htmlAlign atAlign = (htmlAlign)pElement->GetFirstBranch()->GetParaFormat()->_bBlockAlign;

                        if(atAlign==htmlAlignRight || atAlign==htmlAlignCenter)
                        {
                            pLayout = pElement->GetCurLayout();

                            Assert(pLayout);

                            rp.RpSetCp(pElement->GetFirstCp(), FALSE, TRUE);

                            pLayout->SetPosition(
                                rp->GetTextLeft()+pLayout->GetXProposed(),
                                pLayout->GetPositionTop(), TRUE); // notify auto
                        }
                    }
                }
            }
            pDispNode = pDispNode->GetNextSiblingNode(TRUE);
        } while(pDispNode);
    }
}

/*
*  CDisplay::CreateEmptyLine()
*
*  @mfunc
*      Create an empty line
*
*  @rdesc
*      TRUE - worked <nl>
*      FALSE - failed
*
*/
BOOL CDisplay::CreateEmptyLine(CLSMeasurer* pMe, CRecalcLinePtr* pRecalcLinePtr,
                               LONG* pyHeight, BOOL fHasEOP)
{
    UINT uiFlags;
    LONG yAlignDescent;
    INT  iNewLine;

    // Make sure that this is being called appropriately
    AssertSz(!pMe||GetLastCp()==long(pMe->GetCp()),
        "CDisplay::CreateEmptyLine called inappropriately");

    // Assume failure
    BOOL fResult = FALSE;

    // Add one new line
    CLine* pliNew = pRecalcLinePtr->AddLine();

    if(!pliNew)
    {
        Assert(FALSE);
        goto err;
    }

    iNewLine = pRecalcLinePtr->Count() - 1;

    Assert(iNewLine >= 0);

    uiFlags = fHasEOP ? MEASURE_BREAKATWORD|MEASURE_FIRSTINPARA : MEASURE_BREAKATWORD;

    uiFlags |= MEASURE_EMPTYLASTLINE;

    // If this is the first line in the document.
    if(*pyHeight == 0)
    {
        uiFlags |= MEASURE_FIRSTLINE;
    }

    if(!pRecalcLinePtr->MeasureLine(*pMe, uiFlags,
        &iNewLine, pyHeight, &yAlignDescent,
        NULL, NULL))
    {
        goto err;
    }

    // If we made it to here, everything worked.
    fResult = TRUE;

err:
    return fResult;
}

//====================================  Rendering  =======================================
/*
*  CDisplay::Render(rcView, rcRender)
*
*  @mfunc
*      Searches paragraph boundaries around a range
*
*  returns: the lowest yPos
*/
void CDisplay::Render (
                  CFormDrawInfo* pDI,
                  const RECT& rcView,     // View RECT
                  const RECT& rcRender,   // RECT to render (must be container in
                  CDispNode* pDispNode)
{
    CFlowLayout* pFlowLayout = GetFlowLayout();
    CElement* pElementFL = pFlowLayout->ElementOwner();
    LONG    ili;
    LONG    iliBgDrawn = -1;
    POINT   pt;
    LONG    cp;
    LONG    yLine;
    LONG    yLi = 0;
    long    lCount;
    CLine*  pli = NULL;
    RECT    rcClip;
    BOOL    fLineIsPositioned;
    CRect   rcLocal;
    long    iliStart  = -1;
    long    iliFinish = -1;
    CPoint  ptOffset;
    CPoint* pptOffset = NULL;

#ifdef _DEBUG
    LONG    cpLi;
#endif

    // BUGBUG: Assert removed for NT5 B3 (bug #75432)
    // IE6: Reenable for IE6+
    // AssertSz(!pFlowLayout->IsDirty(), "Rendering when layout is dirty -- not a measurer/renderer problem!");
    if(!pFlowLayout->ElementOwner()->IsInMarkup()
        || (!pFlowLayout->IsEditable() && pFlowLayout->IsDisplayNone())
        || pFlowLayout->IsDirty()) // Prevent us from crashing when we do get here in retail mode.
    {
        return;
    }

    // Create renderer
    CLSRenderer lsre(this, pDI);

    if(!lsre._pLS)
    {
        return;
    }

    Assert(pDI->_rc == rcView);
    Assert(((CRect&)(pDI->_rcClip)) == rcRender);

    // Calculate line and cp to start the display at
    rcLocal = rcRender;
    rcLocal.OffsetRect(-((CRect&)rcView).TopLeft().AsSize());

    // if the current layout has multiple text nodes then
    // compute the range of lines the belong to the disp node
    // being rendered.
    if(_fHasMultipleTextNodes && pDispNode)
    {
        GetFlowLayout()->GetTextNodeRange(pDispNode, &iliStart, &iliFinish);

        Assert(iliStart < iliFinish);

        // For backgrounds, RegionFromElement is going to return the
        // rects relative to the layout. So, when we have multiple text
        // nodes pass the (0, -top) as the offset to make the rects
        // text node relative
        pptOffset    = &ptOffset;
        pptOffset->x = 0;
        pptOffset->y = -pDispNode->GetPosition().y;
    }

    // For multiple text node, we want to search for the point only
    // in the lines owned by the text node.
    ili = LineFromPos(rcLocal, &yLine, &cp, 0, iliStart, iliFinish);

    lCount = iliFinish<0 ? LineCount() : iliFinish;

    if(lCount<=0 || ili<0)
    {
        return;
    }

    rcClip = rcRender;

    // Prepare renderer

    if(!lsre.StartRender(rcView, rcRender))
    {
        return;
    }

    // If we're just painting the inset, don't check all the lines.
    if(rcRender.right<=rcView.left ||
        rcRender.left>=rcView.right ||
        rcRender.bottom<=rcView.top ||
        rcRender.top>=rcView.bottom)
    {
        return;
    }

    // Calculate the point where the text will start being displayed
    pt = ((CRect&)rcView).TopLeft();
    pt.y += yLine;

    // Init renderer at the start of the first line to render
    lsre.SetCurPoint(pt);
    lsre.SetCp(cp, NULL);

    DEBUG_ONLY(cpLi = long(lsre.GetCp()));

    yLi = pt.y;

    // Check if this line begins BEFORE the previous line ended ...
    // Would happen with negative line heights:
    //
    //           ----------------------- <-----+---------- Line 1
    //                                         |
    //           ======================= <-----|-----+---- Line 2
    //  yBeforeSpace__________^                |     |
    //                        |                |     |
    //  yLi ---> -------------+--------- <-----+     |
    //                                               |
    //                                               |
    //           ======================= <-----------+
    //
    RecalcMost();

    // Render each line in the update rectangle
    for(; ili<lCount; ili++)
    {
        // current line
        pli = Elem(ili);

        // if the most negative line is out off the view from the current
        // yOffset, don't look any further, we have rendered all the lines
        // in the inval'ed region
        if(yLi+min(long(0), pli->GetYTop())+_yMostNeg>=rcClip.bottom)
        {
            break;
        }

        fLineIsPositioned = FALSE;

        // if the current line is interesting (ignore aligned, clear,
        // hidden and blank lines).
        if(pli->_cch && !pli->_fHidden)
        {
            // if the current line is relative get its y offset and
            // zIndex
            if(pli->_fRelative)
            {
                fLineIsPositioned = TRUE;
            }
        }

        // now check to see if the current line is in view
        if((yLi+min(long(0), pli->GetYTop()))>rcClip.bottom ||
            (yLi+pli->GetYLineBottom()<rcClip.top))
        {
            // line is not in view, so skip it
            lsre.SkipCurLine(pli);
        }
        else
        {
            // current line is in view, so render it
            //
            // Note: we have to render the background on a relative line,
            // the parent of the relative element might have background.(srinib)
            // fix for #51465
            //
            // if the paragraph has background or borders then compute the bounding rect
            // for the paragraph and draw the background and/or border
            if(iliBgDrawn<ili &&
                (pli->_fHasParaBorder || // if we need to draw borders
                (pli->_fHasBackground&&pElementFL->Doc()->PaintBackground())))
            {
                DrawBackgroundAndBorder(lsre.GetDrawInfo(), lsre.GetCp(), ili, lCount,
                    &iliBgDrawn, yLi, &rcView, &rcClip, pptOffset);

                // N.B. (johnv) Lines can be added by as
                // DrawBackgroundAndBorders (more precisely,
                // RegionFromElement, which it calls) waits for a
                // background recalc.  Recompute our cached line pointer.
                pli = Elem(ili);
            }

            // if the current line has is positioned it will be taken care
            // of through the SiteDrawList and we shouldn't draw it here.
            if(fLineIsPositioned)
            {
                lsre.SkipCurLine(pli);
            }
            else
            {
                // Finally, render the current line
                lsre.RenderLine(*pli);
            }
        }

        Assert(pli == Elem(ili));

        // update the yOffset for the next line
        if(pli->_fForceNewLine)
        {
            yLi += pli->_yHeight;
        }

        DEBUG_ONLY(cpLi += pli->_cch);

        AssertSz(long(lsre.GetCp())==cpLi,
            "CDisplay::Render() - cp out of sync. with line table");
        AssertSz(lsre.GetCurPoint().y==yLi,
            "CDisplay::Render() - y out of sync. with line table");
    }

    if(lsre._lastTextOutBy != CLSRenderer::DB_LINESERV)
    {
        lsre._lastTextOutBy = CLSRenderer::DB_NONE;
        SetTextAlign(lsre._hdc, TA_TOP|TA_LEFT);
    }
}

//+---------------------------------------------------------------------------
//
// Member:      DrawBackgroundAndBorders()
//
// Purpose:     Draw background and borders for elements on the current line,
//              and all the consecutive lines that have background.
//
//----------------------------------------------------------------------------
void CDisplay::DrawBackgroundAndBorder(
                                  CFormDrawInfo* pDI,
                                  long          cpIn,
                                  LONG          ili,
                                  LONG          lCount,
                                  LONG*         piliDrawn,
                                  LONG          yLi,
                                  const RECT*   prcView,
                                  const RECT*   prcClip,
                                  const CPoint* pptOffset)
{
    const CCharFormat*  pCF;
    const CFancyFormat* pFF;
    const CParaFormat*  pPF;
    CStackPtrAry<CTreeNode*, 8> aryNodesWithBgOrBorder;
    CDataAry<RECT>      aryRects;
    CFlowLayout*        pFlowLayout = GetFlowLayout();
    CElement*           pElementFL = pFlowLayout->ElementContent();
    CMarkup*            pMarkup = pFlowLayout->GetContentMarkup();
    BOOL                fPaintBackground = pElementFL->Doc()->PaintBackground();
    CTreeNode*          pNodeCurrBranch;
    CTreeNode*          pNode;
    CTreePos *          ptp;
    long                ich;
    long                cpClip = cpIn;
    long                cp;
    long                lSize;

    // find the consecutive set of lines that have background
    while(ili<lCount && yLi+_yMostNeg<prcClip->bottom)
    {
        CLine* pli = Elem(ili++);

        // if the current line has borders or background then
        // continue otherwise return.
        if(!(pli->_fHasBackground && fPaintBackground) &&
            !pli->_fHasParaBorder)
        {
            break;
        }

        if(pli->_fForceNewLine)
        {
            yLi += pli->_yHeight;
        }

        cpClip += pli->_cch;
    }

    if(cpIn != cpClip)
    {
        *piliDrawn = ili - 1;
    }

    // initialize the tree pos that corresponds to the begin cp of the
    // current line
    ptp = pMarkup->TreePosAtCp(cpIn, &ich);

    cp = cpIn - ich;

    // first draw any backgrounds extending into the current region
    // from the previous line.

    pNodeCurrBranch = ptp->GetBranch();

    if(DifferentScope(pNodeCurrBranch, pElementFL))
    {
        pNode = pNodeCurrBranch;

        // run up the current branch and find the ancestors with background
        // or border
        while(pNode)
        {
            if(pNode->HasLayout())
            {
                CLayout* pLayout = pNode->GetCurLayout();

                if(pLayout == pFlowLayout)
                {
                    break;
                }

                Assert(pNode == pNodeCurrBranch);
            }
            else
            {
                // push this element on to the stack
                aryNodesWithBgOrBorder.Append(pNode);
            }
            pNode = pNode->Parent();
        }

        Assert(pNode);

        // now that we have all the elements with background or borders
        // for the current branch render them.
        for(lSize=aryNodesWithBgOrBorder.Size(); lSize>0; lSize--)
        {
            CTreeNode* pNode = aryNodesWithBgOrBorder[lSize-1];

            // In design mode relative elements are drawn in flow (they are treated
            // as if they are not relative). Relative elements draw their own
            // background and their children's background. So, draw only the
            // background of ancestor's if any. (#25583)
            if(pNode->IsRelative() )
            {
                pNodeCurrBranch = pNode;
                break;
            }
            else
            {
                pCF = pNode->GetCharFormat();
                pFF = pNode->GetFancyFormat();
                pPF = pNode->GetParaFormat();

                if(!pCF->IsVisibilityHidden() && !pCF->IsDisplayNone())
                {
                    BOOL fDrawBorder = pPF->_fPadBord  && pFF->_fBlockNess;
                    BOOL fDrawBackground = fPaintBackground &&
                        (pFF->_lImgCtxCookie || pFF->_ccvBackColor.IsDefined());

                    if(fDrawBackground || fDrawBorder)
                    {
                        DrawElemBgAndBorder(
                            pNode->Element(), &aryRects,
                            prcView, prcClip,
                            pDI, pptOffset,
                            fDrawBackground, fDrawBorder,
                            cpIn, -1);
                    }
                }
            }
        }

        // In design mode relative elements are drawn in flow (they are treated
        // as if they are not relative).
        if(pNodeCurrBranch->HasLayout() || pNode->IsRelative())
        {
            CTreePos* ptpBegin;

            pNodeCurrBranch->Element()->GetTreeExtent(&ptpBegin, &ptp);

            cp = ptp->GetCp();
        }

        cp += ptp->GetCch();
        ptp = ptp->NextTreePos();
    }

    // now draw the background of all the elements comming into scope of
    // in the cpRange
    while(ptp && cpClip>=cp)
    {
        if(ptp->IsBeginElementScope())
        {
            pNode = ptp->Branch();
            pCF   = pNode->GetCharFormat();

            // Background and border for a relative element or an element
            // with layout are drawn when the element is hit with a draw.
            if(pNode->HasLayout() || pCF->_fRelative)
            {
                if(DifferentScope(pNode, pElementFL))
                {
                    CTreePos* ptpBegin;

                    pNode->Element()->GetTreeExtent(&ptpBegin, &ptp);
                    cp = ptp->GetCp();
                }
            }
            else
            {
                pCF = pNode->GetCharFormat();
                pFF = pNode->GetFancyFormat();
                pPF = pNode->GetParaFormat();

                if(!pCF->IsVisibilityHidden() && !pCF->IsDisplayNone())
                {
                    BOOL fDrawBorder = pPF->_fPadBord  && pFF->_fBlockNess;
                    BOOL fDrawBackground = fPaintBackground &&
                        (pFF->_lImgCtxCookie || pFF->_ccvBackColor.IsDefined());

                    if(fDrawBackground || fDrawBorder)
                    {
                        DrawElemBgAndBorder(
                            pNode->Element(), &aryRects,
                            prcView, prcClip,
                            pDI, pptOffset,
                            fDrawBackground, fDrawBorder,
                            cp, -1);
                    }
                }
            }
        }

        cp += ptp->GetCch();
        ptp = ptp->NextTreePos();
    }
}

//+----------------------------------------------------------------------------
//
// Function:    BoundingRectForAnArrayOfRectsWithEmptyOnes
//
// Synopsis:    Find the bounding rect that contains a given set of rectangles
//              It does not ignore the rectangles that have left=right, top=bottom
//              or both. It still ignores the rects that have left=right=top=bottom=0
//
//-----------------------------------------------------------------------------
void BoundingRectForAnArrayOfRectsWithEmptyOnes(RECT* prcBound, CDataAry<RECT>* paryRects)
{
    RECT*   prc;
    LONG    iRect;
    LONG    lSize = paryRects->Size();
    BOOL    fFirst = TRUE;

    SetRectEmpty(prcBound);

    for(iRect=0,prc=*paryRects; iRect<lSize; iRect++,prc++)
    {
        if((prc->left<=prc->right && prc->top<=prc->bottom) &&
            (prc->left!=0 || prc->right!=0 || prc->top!=0 || prc->bottom!=0))
        {
            if(fFirst)
            {
                *prcBound = *prc;
                fFirst = FALSE;
            }
            else
            {
                if(prcBound->left > prc->left) prcBound->left = prc->left;
                if(prcBound->top > prc->top) prcBound->top = prc->top;
                if(prcBound->right < prc->right) prcBound->right = prc->right;
                if(prcBound->bottom < prc->bottom) prcBound->bottom = prc->bottom;
            }
        }
    }
}

//+----------------------------------------------------------------------------
//
// Function:    BoundingRectForAnArrayOfRects
//
// Synopsis:    Find the bounding rect that contains a given set of rectangles
//
//-----------------------------------------------------------------------------
void BoundingRectForAnArrayOfRects(RECT* prcBound, CDataAry<RECT>* paryRects)
{
    RECT*   prc;
    LONG    iRect;
    LONG    lSize = paryRects->Size();

    SetRectEmpty(prcBound);

    for(iRect=0,prc=*paryRects; iRect<lSize; iRect++,prc++)
    {
        if(!IsRectEmpty(prc))
        {
            UnionRect(prcBound, prcBound, prc);
        }
    }
}

//+----------------------------------------------------------------------------
//
// Member:      DrawElementBackground
//
// Synopsis:    Draw the background for a an element, given the region it
//              occupies in the display
//
//-----------------------------------------------------------------------------
void CDisplay::DrawElementBackground(
                                     CTreeNode*         pNodeContext,
                                     CDataAry<RECT>*    paryRects,
                                     RECT*              prcBound,
                                     const RECT*        prcView,
                                     const RECT*        prcClip,
                                     CFormDrawInfo*     pDI)
{
    RECT    rcDraw;
    RECT    rcBound = { 0 };
    RECT*   prc;
    LONG    lSize;
    LONG    iRect;
    SIZE    sizeImg;
    RECT    rcBackClip;
    CPoint  ptBackOrg;
    BACKGROUNDINFO bginfo;

    const CFancyFormat* pFF = pNodeContext->GetFancyFormat();
    BOOL    fBlockElement = pNodeContext->Element()->IsBlockElement();

    Assert(pFF->_lImgCtxCookie || pFF->_ccvBackColor.IsDefined());

    CDocument* pDoc = GetFlowLayout()->Doc();
    CImgCtx* pImgCtx = (pFF->_lImgCtxCookie ? pDoc->GetUrlImgCtx(pFF->_lImgCtxCookie) : 0);

    if(pImgCtx && !(pImgCtx->GetState(FALSE, &sizeImg) & IMGLOAD_COMPLETE))
    {
        pImgCtx = NULL;
    }

    // if the background image is not loaded yet and there is no background color
    // return (we dont have anything to draw)
    if(!pImgCtx && !pFF->_ccvBackColor.IsDefined())
    {
        return;
    }

    // now given the rects for a given element
    // draw its background

    // if we have a background image, we need to compute its origin
    lSize = paryRects->Size();

    if(lSize == 0)
    {
        return;
    }

    memset(&bginfo, 0, sizeof(bginfo));

    bginfo.pImgCtx       = pImgCtx;
    bginfo.lImgCtxCookie = pFF->_lImgCtxCookie;
    bginfo.crTrans       = CLR_INVALID;
    bginfo.crBack        = pFF->_ccvBackColor.IsDefined() ? pFF->_ccvBackColor.GetColorRef() : CLR_INVALID;

    if(pImgCtx || fBlockElement)
    {
        if(!prcBound)
        {
            // compute the bounding rect for the element.
            BoundingRectForAnArrayOfRects(&rcBound, paryRects);
        }
        else
        {
            rcBound = *prcBound;
        }
    }

    if(pImgCtx)
    {
        if(!IsRectEmpty(&rcBound))
        {
            SIZE sizeBound;

            sizeBound.cx = rcBound.right - rcBound.left;
            sizeBound.cy = rcBound.bottom - rcBound.top;

            CalcBgImgRect(pNodeContext, pDI, &sizeBound, &sizeImg, &ptBackOrg, &rcBackClip);

            OffsetRect(&rcBackClip, rcBound.left, rcBound.top);

            ptBackOrg.x += rcBound.left - prcView->left;
            ptBackOrg.y += rcBound.top - prcView->top;

            bginfo.ptBackOrg = ptBackOrg;
            pDI->TransformToDeviceCoordinates(&bginfo.ptBackOrg);
        }
    }

    prc = *paryRects;
    rcDraw = *prc++;

    // Background for block element needs to extend for the
    // left to right of rcBound.
    if(fBlockElement)
    {
        rcDraw.left  = rcBound.left;
        rcDraw.right = rcBound.right;
    }

    for(iRect=1; iRect<=lSize; iRect++,prc++)
    {
        if(iRect==lSize || !IsRectEmpty(prc))
        {
            if(iRect != lSize)
            {
                if(fBlockElement)
                {
                    if(prc->top < rcDraw.top)
                    {
                        rcDraw.top = prc->top;
                    }
                    if(prc->bottom > rcDraw.bottom)
                    {
                        rcDraw.bottom = prc->bottom;
                    }
                    continue;
                }
                else if(prc->left==rcDraw.left &&
                    prc->right==rcDraw.right && prc->top==rcDraw.bottom)
                {
                    // add the current rect
                    rcDraw.bottom = prc->bottom;
                    continue;
                }
            }

            {
                IntersectRect(&rcDraw, prcClip, &rcDraw);

                if(!IsRectEmpty(&rcDraw))
                {
                    IntersectRect(&bginfo.rcImg, &rcBackClip, &rcDraw);
                    GetFlowLayout()->DrawBackground(pDI, &bginfo, &rcDraw);
                }

                if(iRect != lSize)
                {
                    rcDraw = *prc;
                }
            }
        }
    }
}

//+----------------------------------------------------------------------------
//
// Function:    DrawElementBorder
//
// Synopsis:    Find the bounding rect that contains a given set of rectangles
//
//-----------------------------------------------------------------------------
void CDisplay::DrawElementBorder(
                                 CTreeNode*     pNodeContext,
                                 CDataAry <RECT>* paryRects,
                                 RECT*          prcBound,
                                 const RECT*    prcView,
                                 const RECT*    prcClip,
                                 CFormDrawInfo* pDI)
{
    CBorderInfo borderInfo;
    CElement* pElement = pNodeContext->Element();

    if(pNodeContext->GetCharFormat()->IsVisibilityHidden())
    {
        return;
    }

    if(!pElement->_fDefinitelyNoBorders &&
        FALSE==(pElement->_fDefinitelyNoBorders=!GetBorderInfoHelper(pNodeContext, pDI, &borderInfo, TRUE)))
    {
        RECT rcBound;

        if(!prcBound)
        {
            if(paryRects->Size() == 0)
            {
                return;
            }

            // compute the bounding rect for the element.
            BoundingRectForAnArrayOfRects(&rcBound, paryRects);
        }
        else
        {
            rcBound = *prcBound;
        }

        DrawBorder(pDI, &rcBound, &borderInfo);
    }
}

// =============================  Misc  ===========================================
//+--------------------------------------------------------------------------------
//
// Synopsis: return true if this is the last text line in the line array
//---------------------------------------------------------------------------------
BOOL CDisplay::IsLastTextLine(LONG ili)
{
    Assert(ili>=0 && ili<LineCount());
    if(LineCount() == 0)
    {
        return TRUE;
    }

    for(LONG iliT=ili+1; iliT<LineCount(); iliT++)
    {
        if(Elem(iliT)->IsTextLine())
        {
            return FALSE;
        }
    }
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     FormattingNodeForLine
//
//  Purpose:    Returns the node which controls the formatting at the BOL. This
//              is needed because the first char on the line may not necessarily
//              be the first char in the paragraph.
//
//----------------------------------------------------------------------------
CTreeNode* CDisplay::FormattingNodeForLine(
                                LONG        cpForLine,                 // IN
                                CTreePos*   ptp,                       // IN
                                LONG        cchLine,                   // IN
                                LONG*       pcch,                      // OUT
                                CTreePos**  pptp,                      // OUT
                                BOOL*       pfMeasureFromStart) const  // OUT
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    BOOL            fIsEditable = pFlowLayout->IsEditable();
    CMarkup*        pMarkup     = pFlowLayout->GetContentMarkup();
    CTreeNode*      pNode       = NULL;
    CElement*       pElement;
    LONG            lNotNeeded;
    LONG            cch = cchLine;
    BOOL            fSawOpenLI  = FALSE;
    BOOL            fSeenAbsolute = FALSE;
    BOOL            fSeenBeginBlockTag = FALSE;

    if(!ptp)
    {
        ptp = pMarkup->TreePosAtCp(cpForLine, &lNotNeeded);
        Assert(ptp);
    }
    else
    {
        Assert(ptp->GetCp() <= cpForLine);
        Assert(ptp->GetCp()+ptp->GetCch() >= cpForLine);
    }
    if(pfMeasureFromStart)
    {
        *pfMeasureFromStart = FALSE;
    }
    while(cch > 0)
    {
        if(ptp->IsPointer())
        {
            ptp = ptp->NextTreePos();
            continue;
        }

        if(ptp->IsText())
        {
            if(ptp->Cch())
            {
                break;
            }
            else
            {
                ptp = ptp->NextTreePos();
                continue;
            }
        }

        Assert(ptp->IsNode());

        pNode = ptp->Branch();
        pElement = pNode->Element();

        if(ptp->IsBeginElementScope())
        {
            if(fSeenAbsolute)
            {
                break;
            }

            const CCharFormat* pCF = pNode->GetCharFormat();
            if(pCF->IsDisplayNone())
            {
                cch -= pFlowLayout->GetNestedElementCch(pElement, &ptp) - 1;
                pNode = pNode->Parent();
            }
            else if(pNode->NeedsLayout())
            {
                if(pNode->IsAbsolute())
                {
                    cch -= pFlowLayout->GetNestedElementCch(pElement, &ptp) - 1;
                    pNode = pNode->Parent();
                    fSeenAbsolute = TRUE;
                }
                else
                {
                    break;
                }
            }
            else if(pElement->Tag() == ETAG_BR)
            {
                break;
            }
            else if(pElement->IsFlagAndBlock(TAGDESC_LIST) && fSawOpenLI)
            {
                CTreePos* ptpT = ptp;

                do
                {
                    ptpT = ptpT->PreviousTreePos();
                } while(ptpT->GetCch() == 0);

                pNode = ptpT->Branch();

                break;
            }                        
            else if(pElement->IsTagAndBlock(ETAG_LI))
            {
                fSawOpenLI = TRUE;
            }
            else if(pFlowLayout->IsElementBlockInContext(pElement))
            {
                fSeenBeginBlockTag = TRUE;
            }
        }
        else if(ptp->IsEndNode())
        {
            if(fSeenAbsolute)
            {
                break;
            }

            // If we encounter a break on empty block end tag, then we should
            // give vertical space otherwise a <X*></X> where X is a block element
            // will not produce any vertical space. (Bug 45291).
            if(fSeenBeginBlockTag && pElement->_fBreakOnEmpty
                && pFlowLayout->IsElementBlockInContext(pElement))
            {
                break;
            }

            if(fSawOpenLI                               // Skip over the close LI, unless we saw an open LI,
                && ptp->IsEdgeScope()                   // which would imply that we have an empty LI.  An
                && pElement->IsTagAndBlock(ETAG_LI))    // empty LI gets a bullet, so we need to break.
            {
                break;
            }

            pNode = pNode->Parent();
        }

        cch--;
        ptp = ptp->NextTreePos();
    }

    if(pcch)
    {
        *pcch = cchLine - cch;
    }

    if(pptp)
    {
        *pptp = ptp;
    }

    if(!pNode)
    {
        pNode = ptp->GetBranch();
        if(ptp->IsEndNode())
        {
            pNode = pNode->Parent();
        }
    }

    return pNode;
}

//+---------------------------------------------------------------------------
//
//  Member:     EndNodeForLine
//
//  Purpose:    Returns the first node which ends this line. If the line ends
//              because of insufficient width, then the node is the node
//              above the last character in the line, else it is the node
//              which causes the line to end (like /p).
//
//----------------------------------------------------------------------------
CTreeNode* CDisplay::EndNodeForLine(
                         LONG       cpEndForLine,               // IN
                         CTreePos*  ptp,                        // IN
                         LONG*      pcch,                       // OUT
                         CTreePos** pptp,                       // OUT
                         CTreeNode** ppNodeForAfterSpace) const // OUT
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    BOOL            fIsEditable = pFlowLayout->IsEditable();
    CTreePos        *ptpStart, *ptpStop;
    CTreePos*       ptpNext = ptp;
    CTreePos*       ptpOriginal = ptp;
    CTreeNode*      pNode;
    CElement*       pElement;
    CTreeNode*      pNodeForAfterSpace = NULL;

    // If we are in the middle of a text run then we do not need to
    // do anything, since this line will not be getting any para spacing
    if(ptpNext->IsText() && ptpNext->GetCp()<cpEndForLine)
    {
        goto Cleanup;
    }

    pFlowLayout->GetContentTreeExtent(&ptpStart, &ptpStop);

    // We should never be here if we start measuring at the beginning
    // of the layout.
    Assert(ptp != ptpStart);

    ptpStart = ptpStart->NextTreePos();
    while(ptp != ptpStart)
    {
        ptpNext = ptp;
        ptp = ptp->PreviousTreePos();

        if(ptp->IsPointer())
        {
            continue;
        }

        if(ptp->IsNode())
        {
            pNode = ptp->Branch();
            pElement = pNode->Element();
            if(ptp->IsEndElementScope())
            {
                const CCharFormat* pCF = pNode->GetCharFormat();
                if(pCF->IsDisplayNone())
                {
                    pElement->GetTreeExtent(&ptp, NULL);
                }
                else if(pNode->NeedsLayout())
                {
                    // We need to collect after space info from the nodes which
                    // have layouts.
                    if(pElement->IsOwnLineElement(pFlowLayout))
                    {
                        pNodeForAfterSpace = pNode;
                    }
                    break;
                }
                else if(pElement->Tag() == ETAG_BR)
                {
                    break;
                }
            }
            else if(ptp->IsBeginElementScope())
            {
                if(pFlowLayout->IsElementBlockInContext(pElement)
                    || !pNode->Element()->IsNoScope())
                {
                    break;
                }
                else if(pElement->_fBreakOnEmpty
                    && pFlowLayout->IsElementBlockInContext(pElement))
                {
                    break;
                }
            }
        }
        else
        {
            Assert(ptp->IsText());
            if(ptp->Cch())
            {
                break;
            }
        }
    }

    Assert(ptpNext);

Cleanup:
    if(pptp)
    {
        *pptp = ptpNext;
    }
    if(pcch)
    {
        if(ptpNext == ptpOriginal)
        {
            *pcch = 0;
        }
        else
        {
            *pcch = cpEndForLine - ptpNext->GetCp();
        }
    }
    if(ppNodeForAfterSpace)
    {
        *ppNodeForAfterSpace = pNodeForAfterSpace;
    }

    return ptpNext->GetBranch();
}

long ComputeLineShift(
                      htmlAlign atAlign,
                      BOOL  fRTLDisplay,
                      BOOL  fRTLLine,
                      BOOL  fMinMax,
                      long  xWidthMax,
                      long  xWidth,
                      UINT* puJustified,
                      long* pdxRemainder)
{
    long xShift = 0;
    long xRemainder = 0;

    switch(atAlign)
    {
    case htmlAlignNotSet:
        if(fRTLDisplay == fRTLLine)
        {
            *puJustified = JUSTIFY_LEAD;
        }
        else
        {
            *puJustified = JUSTIFY_TRAIL;
        }
        break;

    case htmlAlignLeft:
        if(!fRTLDisplay)
        {
            *puJustified = JUSTIFY_LEAD;
        }
        else
        {
            *puJustified = JUSTIFY_TRAIL;
        }
        break;

    case htmlAlignRight:
        if(!fRTLDisplay)
        {
            *puJustified = JUSTIFY_TRAIL;
        }
        else
        {
            *puJustified = JUSTIFY_LEAD;
        }
        break;

    case htmlAlignCenter:
        *puJustified = JUSTIFY_CENTER;
        break;

    case htmlBlockAlignJustify:
        if(!fRTLDisplay)
        {
            if(!fRTLLine)
            {
                *puJustified = JUSTIFY_FULL;
            }
            else
            {
                *puJustified = JUSTIFY_TRAIL;
            }
        }
        else
        {
            *puJustified = JUSTIFY_FULL;
        }
        break;

    default:
        AssertSz(FALSE, "Did we introduce new type of alignment");
        break;
    }
    if(!fMinMax && *puJustified!=JUSTIFY_FULL)
    {
        if(*puJustified != JUSTIFY_LEAD)
        {
            // for pre whitespace is already include in _xWidth
            xShift = xWidthMax - xWidth;

            xShift = max(xShift, 0L); // Don't allow alignment to go < 0 Can happen with a target device
            if(*puJustified == JUSTIFY_CENTER)
            {
                Assert(atAlign == htmlAlignCenter);
                xShift /= 2;
            }
        }
        xRemainder = xWidthMax - xWidth - xShift;
    }

    Assert(pdxRemainder != NULL);
    *pdxRemainder = xRemainder;

    return xShift;
}

extern CDispNode* EnsureContentNode(CDispNode* pDispContainer);

HRESULT CDisplay::InsertNewContentDispNode(
    CDispNode*  pDNBefore,
    CDispNode** ppDispContent,
    long        iLine,
    long        yHeight)
{
    HRESULT         hr  = S_OK;
    CFlowLayout*    pFlowLayout     = GetFlowLayout();
    CDispNode*      pDispContainer  = pFlowLayout->GetElementDispNode(); 
    CDispNode*      pDispNewContent = NULL;
    BOOL            fRightToLeft;

    Assert(pDispContainer);

    fRightToLeft = pDispContainer->IsRightToLeft();

    // if a content node is not created yet, ensure that we have a content node.
    if(!pDNBefore)
    {
        pDispContainer = pFlowLayout->EnsureDispNodeIsContainer();

        EnsureContentNode(pDispContainer);

        pDNBefore = pFlowLayout->GetFirstContentDispNode();

        Assert(pDNBefore);

        if(!pDNBefore)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        *ppDispContent = pDNBefore;
    }

    Assert(pDispContainer->IsContainer());

    // Create a new content dispNode and size the previous dispNode
    pDispNewContent = CDispRoot::CreateDispItemPlus(
        pFlowLayout,
        TRUE,
        FALSE,
        FALSE,
        DISPNODEBORDER_NONE,
        fRightToLeft);

    if(!pDispNewContent)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pDispNewContent->SetPosition(CPoint(0, yHeight));
    pDispNewContent->SetSize(CSize(_xWidthView, 1), FALSE);

    if(fRightToLeft)
    {
        pDispNewContent->FlipBounds();
    }

    pDispNewContent->SetVisible(pDispContainer->IsVisible());
    pDispNewContent->SetExtraCookie((void*)(DWORD_PTR)(iLine));
    pDispNewContent->SetLayerType(DISPNODELAYER_FLOW);

    pDNBefore->InsertNextSiblingNode(pDispNewContent);

    *ppDispContent = pDispNewContent;

Cleanup:
    RRETURN(hr);
}

HRESULT CDisplay::HandleNegativelyPositionedContent(
    CLine*          pliNew,
    CLSMeasurer*    pme,
    CDispNode*      pDNBefore,
    long            iLinePrev,
    long            yHeight)
{
    HRESULT hr = S_OK;
    CDispNode* pDNContent = NULL;

    Assert(pliNew);

    NoteMost(pliNew);

    if(iLinePrev > 0)
    {
        long yLineTop = pliNew->GetYTop();

        // Create and insert a new content disp node, if we have negatively
        // positioned content.

        // NOTE(SujalP): Changed from GetYTop to _yBeforeSpace. The reasons are
        // outlined in IE5 bug 62737.
        if(pliNew->_yBeforeSpace<0 && !pliNew->_fDummyLine)
        {
            hr = InsertNewContentDispNode(pDNBefore, &pDNContent, iLinePrev, yHeight+yLineTop);

            if(hr)
            {
                goto Cleanup;
            }

            _fHasMultipleTextNodes = TRUE;

            if(pDNBefore == pme->_pDispNodePrev)
            {
                pme->_pDispNodePrev = pDNContent;
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------------
//
//  Member : ElementResize
//
//  Synopsis : CDisplay helper function to resize the element when the text content
//      has been remeasured and the container needs to be resized
//
//--------------------------------------------------------------------------------
void CDisplay::ElementResize(CFlowLayout* pFlowLayout)
{
    if(!pFlowLayout)
    {
        return;
    }

    // If our contents affects our size, ask our parent to initiate a re-size
    if(pFlowLayout->GetAutoSize() || pFlowLayout->_fContentsAffectSize)
    {
        pFlowLayout->ElementOwner()->ResizeElement();
    }
}

// ================================  Line info retrieval  ====================================
/*
*  CDisplay::YposFromLine(ili)
*
*  @mfunc
*      Computes top of line position
*
*  @rdesc
*      top position of given line (relative to the first line)
*      Computed by accumulating CLine _yHeight's for each line with
*      _fForceNewLine true.
*/
LONG CDisplay::YposFromLine(
                            LONG        ili,    //@parm Line we're interested in
                            CCalcInfo*  pci)
{
    LONG yPos;
    CLine* pli;

    // if the flowlayout is hidden, all we have is zero height lines.
    if(GetFlowLayout()->IsDisplayNone())
    {
        return 0;
    }

    if(!WaitForRecalcIli(ili, pci)) // out of range, use last valid line
    {
        ili = LineCount() -1;
        ili = (ili>0) ? ili : 0;
    }

    yPos = 0;
    for(long i=0; i<ili; i++)
    {
        pli = Elem(i);
        if(pli->_fForceNewLine)
        {
            yPos += pli->_yHeight;
        }
    }

    return yPos;
}

/*
*  CDisplay::CpFromLine(ili, pyHeight)
*
*  @mfunc
*      Computes cp at start of given line
*      (and top of line position relative to this display)
*
*  @rdesc
*      cp of given line
*/
LONG CDisplay::CpFromLine(
                           LONG ili,               // Line we're interested in (if <lt> 0 means caret line)
                           LONG* pyHeight) const   // Returns top of line relative to display
{
    long cp = GetFirstCp();
    long y  = 0;
    CLine* pli;

    for(long i=0; i<ili; i++)
    {
        pli = Elem(i);
        if(pli->_fForceNewLine)
        {
            y += pli->_yHeight;
        }
        cp += pli->_cch;
    }

    if(pyHeight)
    {
        *pyHeight = y;
    }

    return cp;
}

//+----------------------------------------------------------------------------
//
//  Member:     Notify
//
//  Synopsis:   Adjust all internal caches in response to a text change
//              within a nested CTxtSite
//
//  Arguments:  pnf - Notification that describes the change
//
//  NOTE: Only those caches which the notification manager is unaware of
//        need updating (e.g., _dcpCalcMax). Others, such as outstanding
//        change transactions, are handled by the notification manager.
//
//-----------------------------------------------------------------------------
BOOL LineIsNotInteresting(CLinePtr& rp)
{
    return rp->IsFrame()||rp->IsClear();
}

void CDisplay::Notify(CNotification* pnf)
{
    CFlowLayout*    pFlowLayout    = GetFlowLayout();
    CElement*       pElementFL     = pFlowLayout->ElementOwner();
    BOOL            fIsDirty       = pFlowLayout->IsDirty();
    long            cpFirst        = pFlowLayout->GetContentFirstCp();
    long            cpDirty        = cpFirst + pFlowLayout->Cp();
    long            cchNew         = pFlowLayout->CchNew();
    long            cchOld         = pFlowLayout->CchOld();
    long            cpMax          = _dcpCalcMax;
    long            cchDelta       = pnf->CchChanged();

    // If no lines yet exist, exit
    if(!LineCount())
    {
        goto Cleanup;
    }

    // Determine the end of the line array
    // (This is normally the maximum calculated cp, but that cp must
    //  be adjusted to take into account outstanding changes.
    //  Changes which overlap the maximum calculated cp effectively
    //  move it to the first cp affected by the change. Changes which
    //  come entirely before move it by the delta in characters of the
    //  change.)
    if(fIsDirty && pFlowLayout->Cp()<cpMax)
    {
        if(pFlowLayout->Cp()+cchOld >= cpMax)
        {
            cpMax = pFlowLayout->Cp();
        }
        else
        {
            cpMax += cchNew - cchOld;
        }
    }

    // If the change is past the end of the line array, exit
    if((pnf->Cp(cpFirst)-cpFirst) > cpMax)
    {
        goto Cleanup;
    }

    // If the change is entirely before or after any pending changes,
    // update the cch of the affected line, the maximum calculated cp,
    // and, if necessary, the first visible cp
    // (Changes which occur within the range of a pending change,
    //  need not be tracked since the affected lines will be completely
    //  recalculated. Additionally, the outstanding change which
    //  represents those changes will itself be updated (by other means)
    //  to cover the changes within the range.)
    // NOTE: The "old" cp must be used to search the line array
    if(cchDelta
        && (!fIsDirty || pnf->Cp(cpFirst)<cpDirty || pnf->Cp(cpFirst)>=cpDirty+cchNew))
    {
        CLinePtr    rp(this);
        long        cchCurrent  = pElementFL->GetElementCch();
        long        cchPrevious = cchCurrent - (fIsDirty ? cchNew-cchOld : 0);
        long        cpPrevious  = !fIsDirty||pnf->Cp(cpFirst)<cpDirty
            ? pnf->Cp(cpFirst) : pnf->Cp(cpFirst)+(cchOld-cchNew);

        // Adjust the maximum calculated cp
        // (Sanitize the value as well - invalid values can result when handling notifications from
        //  elements that extend outside the layout)
        _dcpCalcMax += cchDelta;
        if(_dcpCalcMax < 0)
        {
            _dcpCalcMax = 0;
        }
        else if(_dcpCalcMax > cchPrevious)
        {
            _dcpCalcMax = cchPrevious;
        }

        //----------------------------------------------------------------------------------
        //
        // BEGIN HACK ALERT! BEGIN HACK ALERT! BEGIN HACK ALERT! BEGIN HACK ALERT! 
        //
        // All the code here to find out the correct line to add or remove chars from.
        // Its impossible to accurately detect a line to which characters can be added.
        // So this code makes a best attempt!
        //
        //

        //
        // Find and adjust the affected line
        rp.RpSetCp(cpPrevious, FALSE, FALSE);

        if(cchDelta > 0)
        {
            // We adding at the end of the site?
            if(pnf->Handler()->GetLastCp() == pnf->Cp(cpFirst)+cchDelta)
            {
                // We are adding to the end of the site, we need to go back to a line which
                // has characters so that we can tack on these characters to that line.
                // Note that we have to go back _atleast_ once so that its the prev line
                // onto which we tack on these characters.
                //
                // Note: This also handles the case when the site was empty (no chars) and
                // we are adding characters to it. In this case we will add to the first
                // line prior to the site, because there is no line in the line array where
                // we can add these chars (since it was empty in the first place). If we
                // cannot find a line with characters before this line -- could be we are
                // adding chars to an empty table at beg. of doc., then we will find the
                // line _after_ cp to add it. If we cannot find that either then we will
                // bail out.

                // So go back once only if we were ambigously positioned (We will be correctly
                // positioned if this is last line and last cp of handler is the last cp of this
                // flowlayout). Dont do anything if you cannot go back -- in that
                // case we will go forward later on.
                if(rp.RpGetIch() == 0)
                {
                    rp.PrevLine(FALSE, FALSE);
                }

                // OK, so look backwards to find a line with characters.
                while(LineIsNotInteresting(rp))
                {
                    if(!rp.PrevLine(FALSE, FALSE))
                    {
                        break;
                    }
                }

                // If we broke out of the while look above, it would mean that we did not
                // find an interesting line before the one at which we were positioned.
                // This should only happen when we have an empty site (like table) at the
                // beginning of the document. In this case we will go forward (the code
                // outside this if block) to find a line. YUCK! This is not ideal but is
                // the best we can do in this situation.
            }

            // We will fall into the following while loop in 3 cases:
            // 1) When we were positioned at the beginning of the site: In this case
            //    we have to make sure that we add the chars to an interesting line.
            //    (note the 3rd possibility, i.e. adding in the middle of a site is trivial
            //    since the original RpSetCp will position us to an interesting line).
            // 2) We were positioned at the end of the site but were unable to find a prev
            //    line, hence we are now looking forward.
            //
            // If we cannot find _any_ interesting line then we shrug our shoulders
            // and bail out.
            while(LineIsNotInteresting(rp))
            {
                if(!rp.NextLine(FALSE, FALSE))
                {
                    goto Cleanup;
                }
            }

            // BUGBUG: Arye - It shouldn't be necessary to do this here, it should be possible
            // to do it only in the case where we're adding to the end. Since this code
            // is likely to go away with the display tree I'm not going to spend a lot of
            // time making that work right now.
            // Before this, however, in edit mode we might end up on the last (BLANK) line.
            // This is bad, nothing is there to change the number of characters,
            // so just back up.
            if(rp->_cch==0 && pElementFL->IsEditable() &&
                rp.GetLineIndex()==LineCount()-1)
            {
                do  
                {
                    if(!rp.PrevLine(FALSE, FALSE))
                    {
                        goto Cleanup;
                    }
                } while(LineIsNotInteresting(rp));
            }

            // Right, if we are here then we have a line into which we can pour
            // the characters.
        }
        // We are removing chars. Easy problem; just find a line from which we can remove
        // the chars!
        else
        {
            while(!rp->_cch)
            {
                if(!rp.NextLine(FALSE, FALSE))
                {
                    Assert("No Line to goto, doing nothing");
                    goto Cleanup;
                }
            }
        }
        // END HACK ALERT! END HACK ALERT! END HACK ALERT! END HACK ALERT! END HACK ALERT! 
        //
        //----------------------------------------------------------------------------------

        Assert(rp.GetLineIndex() >= 0);
        Assert(cchDelta > 0 || !rp.GetLineIndex() || ( rp->_cch+cchDelta )>=0);

        // Update the number of characters in the line
        // (Sanitize the character count if the delta is too large - this can happen for
        // notifications from elements that extend outside the layout)
        rp->_cch += cchDelta;
        if(!rp.GetLineIndex())
        {
            rp->_cch = max(rp->_cch, 0L);
        }
        else if(rp.IsLastTextLine())
        {
            rp->_cch = min(rp->_cch, rp.RpGetIch()+(cchCurrent-(cpPrevious-cpFirst)));
        }
    }

Cleanup:
    return;
}

//+----------------------------------------------------------------------------
//
//  Member:     LineFromPos
//
//  Synopsis:   Computes the line at a given x/y and returns the appropriate
//              information
//
//  Arguments:  prc      - CRect that describes the area in which to look for the line
//              pyLine   - Returned y-offset of the line (may be NULL)
//              pcpLine  - Returned cp at start of line (may be NULL)
//              grfFlags - LFP_xxxx flags
//
//  Returns:    Index of found line (-1 if no line is found)
//
//-----------------------------------------------------------------------------
LONG CDisplay::LineFromPos(
                           const CRect& rc,
                           LONG*        pyLine,
                           LONG*        pcpLine,
                           DWORD        grfFlags,
                           LONG         iliStart,
                           LONG         iliFinish) const
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CElement*       pElement    = pFlowLayout->ElementOwner();
    CLine*          pli;
    long            ili;
    long            yli;
    long            cpli;
    long            iliCandidate;
    long            yliCandidate;
    long            cpliCandidate;
    BOOL            fCandidateWhiteHit;
    long            yIntersect;
    BOOL            fInBrowse = !pElement->IsEditable();

    CRect myRc(rc);

    if(myRc.top < 0)
    {
        myRc.top = 0;
    }
    if(myRc.bottom < 0)
    {
        myRc.bottom = 0;
    }

    Assert(myRc.top <= myRc.bottom);
    Assert(myRc.left <= myRc.right);

    // If hidden or no lines exist, return immediately
    if((pElement->IsDisplayNone() && pElement->Tag()!=ETAG_BODY)
        || LineCount()==0)
    {
        ili  = -1;
        yli  = 0;
        cpli = GetFirstCp();
        goto FoundIt;
    }

    if(iliStart < 0)
    {
        iliStart = 0;
    }

    if(iliFinish < 0)
    {
        iliFinish = LineCount();
    }

    Assert(iliStart<iliFinish && iliFinish<=LineCount());

    ili  = iliStart;
    pli  = Elem(ili);;

    if(iliStart > 0)
    {
        yli  = -pli->GetYTop();
        cpli = CpFromLine(ili);
    }
    else
    {
        yli  = 0;
        cpli = GetFirstCp();
    }

    iliCandidate  = -1;
    yliCandidate  = -1;
    cpliCandidate = -1;
    fCandidateWhiteHit = TRUE;

    // Set up to intersect top or bottom of the passed rectangle
    yIntersect = (grfFlags&LFP_INTERSECTBOTTOM
        ? myRc.bottom-1 : myRc.top);

    // Examine all lines that intersect the passed offsets
    while(ili<iliFinish && yIntersect>=yli+_yMostNeg)
    {
        pli = Elem(ili);

        //  Skip over lines that should be ignored
        //  These include:
        //      1. Lines that do not intersect
        //      2. Hidden and dummy lines
        //      3. Relative lines (when requested)
        //      4. Line chunks that do not intersect the x-offset
        //      5. Its an aligned line and we have been asked to skip aligned lines
        if(yIntersect>=yli+pli->GetYLineTop()
            && yIntersect<yli+pli->GetYLineBottom()
            && !pli->_fHidden
            && (!pli->_fDummyLine
            || !fInBrowse)
            && (!pli->_fRelative
            || !(grfFlags&LFP_IGNORERELATIVE))
            && (pli->_fForceNewLine
            || (!pli->_fRTL
            ? myRc.left<=pli->GetTextRight(ili==LineCount()-1)
            : myRc.left>=pli->GetRTLTextLeft()))
            && ((!pli->IsFrame()
            || !(grfFlags & LFP_IGNOREALIGNED))))
        {
            //  If searching for the top-most line in z-order,
            //  then save the "hit" line and continue the search
            //  NOTE: Progressing up through the z-order, multiple lines can be hit.
            //        Hits on text always win over hits on whitespace.
            if(grfFlags & LFP_ZORDERSEARCH)
            {
                BOOL fWhiteHit = (!pli->_fRTL
                    ? (myRc.left<pli->GetTextLeft()-(pli->_fHasBulletOrNum?pli->_xLeft:0)
                    || myRc.left>pli->GetTextRight(ili==LineCount()-1))
                    : (myRc.left<pli->GetRTLTextLeft()
                    || myRc.left>pli->GetRTLTextRight()-(pli->_fHasBulletOrNum?pli->_xRight: 0)))
                    || (yIntersect>=yli && yIntersect<(yli+pli->GetYTop()))
                    || (yIntersect>=(yli+pli->GetYBottom()) && yIntersect<(yli+pli->_yHeight));

                if(iliCandidate<0 || !fWhiteHit || fCandidateWhiteHit)
                {
                    iliCandidate       = ili;
                    yliCandidate       = yli;
                    cpliCandidate      = cpli;
                    fCandidateWhiteHit = fWhiteHit;
                }
            }
            // Otherwise, the line is found
            else
            {
                goto FoundIt;
            }

        }

        if(pli->_fForceNewLine)
        {
            yli += pli->_yHeight;
        }

        cpli += pli->_cch;
        ili++;
    }

    // if we are lookig for an exact line hit and
    // do not have a candidate line, it's a miss
    if(iliCandidate<0 && grfFlags&LFP_EXACTLINEHIT)
    {
        return -1;
    }

     Assert(ili <= LineCount());

    // we better have a candidate, if yIntersect < yli + _yMostNeg
    Assert(iliCandidate>=0 || yIntersect>=yli+_yMostNeg || (grfFlags&LFP_IGNORERELATIVE));

    //  No intersecting line was found, take either the candidate or last line
    //

    //
    //  ili == LineCount() - is TRUE only if the point we are looking for is
    //  below all the content or we found a candidate line but are performing
    //  a Z-Order search on a layout with lines with negative margin.
    if(ili==iliFinish
        //  Here we don't really need to check if iliCandidate >= 0. It is added
        //  to make the code more robust to handle cases like a negative yIntersect
        //  passed in.
        || (yIntersect<yli+_yMostNeg  && iliCandidate>=0))
    {
        // If a candidate line was found, use it
        if(iliCandidate >= 0)
        {
            Assert(yliCandidate  >= 0);
            Assert(cpliCandidate >= 0);

            ili  = iliCandidate;
            yli  = yliCandidate;
            cpli = cpliCandidate;
        }
        // Otherwise use the last line
        else
        {
            Assert(pli);
            Assert(ili > 0);
            Assert(LineCount());

            ili--;

            if(pli->_fForceNewLine)
            {
                yli -= pli->_yHeight;
            }

            cpli -= pli->_cch;
        }

        // Ensure that frame lines are not returned if they are to be ignored
        if(grfFlags & LFP_IGNOREALIGNED)
        {
            pli = Elem(ili);

            if(pli->IsFrame())
            {
                while(pli->IsFrame() && ili)
                {
                    ili--;
                    pli = Elem(ili);
                }

                if(pli->_fForceNewLine)
                {
                    yli -= pli->_yHeight;
                }
                cpli -= pli->_cch;
            }
        }
    }

FoundIt:
    Assert(ili < LineCount());

    if(pyLine)
    {
        *pyLine = yli;
    }

    if(pcpLine)
    {
        *pcpLine = cpli;
    }

    return ili;
}

/*
*  CDisplay::LineFromCp(cp, fAtEnd)
*
*  @mfunc
*      Computes line containing given cp.
*
*  @rdesc
*      index of line found, -1 if no line at that cp.
*/
LONG CDisplay::LineFromCp(
                          LONG cp,      //@parm cp to look for
                          BOOL fAtEnd)  //@parm If true, return previous line for ambiguous cp
{
    CLinePtr rp(this);

    if(GetFlowLayout()->IsDisplayNone() || !WaitForRecalc(cp, -1) || !rp.RpSetCp(cp, fAtEnd))
    {
        return -1;
    }

    return (LONG)rp;
}

//==============================  Point <-> cp conversion  ==============================
LONG CDisplay::CpFromPointReally(
                                 POINT      pt,               // Point to compute cp at (client coords)
                                 CLinePtr* const prp,         // Returns line pointer at cp (may be NULL)
                                 CTreePos** pptp,             // pointer to return TreePos corresponding to the cp
                                 DWORD      dwCFPFlags,       // flags to CpFromPoint
                                 BOOL*      pfRightOfCp,
                                 LONG*      pcchPreChars)
{
    CMessage msg;
    HTC htc;
    CTreeNode* pNodeElementTemp;
    DWORD dwFlags = HT_SKIPSITES | HT_VIRTUALHITTEST | HT_IGNORESCROLL;
    LONG cpHit;
    CFlowLayout* pFlowLayout = GetFlowLayout();
    CTreeNode* pNode = pFlowLayout->GetFirstBranch();
    CRect rcClient;

    Assert(pNode);

    CElement* pContainer = pNode->GetContainer();

    pFlowLayout->GetContentRect(&rcClient, COORDSYS_GLOBAL);

    if(pt.x < rcClient.left)
    {
        pt.x = rcClient.left;
    }
    if(pt.x >= rcClient.right)
    {
        pt.x = rcClient.right - 1;
    }
    if(pt.y < rcClient.top)
    {
        pt.y = rcClient.top;
    }
    if(pt.y >= rcClient.bottom)
    {
        pt.y = rcClient.bottom - 1;
    }

    msg.pt = pt;

    if(dwCFPFlags & CFP_ALLOWEOL)
    {
        dwFlags |= HT_ALLOWEOL;
    }
    if(!(dwCFPFlags & CFP_IGNOREBEFOREAFTERSPACE))
    {
        dwFlags |= HT_DONTIGNOREBEFOREAFTER;
    }
    if(!(dwCFPFlags & CFP_EXACTFIT))
    {
        dwFlags |= HT_NOEXACTFIT;
    }

    // Ideally I would have liked to perform the hit test on _pFlowLayout itself,
    // however when we have relatively positioned lines, then this will miss
    // some lines (bug48689). To avoid missing such lines we have to hittest
    // from the ped. However, doing that has other problems when we are
    // autoselecting. For example, if the point is on a table, the table finds
    // the closest table cell and passes that table cell the message to
    // extend a selection. However, this hittest hits the table and
    // msg._cpHit is not initialized (bug53706). So in this case, do the CpFromPoint
    // in the good ole' traditional manner.
    //

    // Note (yinxie): I changed GetPad()->HitTestPoint to GetContainer->HitTestPoint
    // for the flat world, container is the new notion to replace ped
    // this will ix the bug (IE5 17135), the bugs mentioned above are all checked
    // there is no regression.
    htc = pContainer->HitTestPoint(&msg, &pNodeElementTemp, dwFlags);

    if(htc>=HTC_YES && pNodeElementTemp
        && (pNodeElementTemp->IsContainer()
        && pNodeElementTemp->GetContainer()!=pContainer))
    {
        htc= HTC_NO;
    }

    if(htc>=HTC_YES && msg.resultsHitTest._cpHit>=0)
    {
        cpHit = msg.resultsHitTest._cpHit;
        if(prp)
        {
            prp->RpSet(msg.resultsHitTest._iliHit, msg.resultsHitTest._ichHit);
        }
        if(pfRightOfCp)
        {
            *pfRightOfCp = msg.resultsHitTest._fRightOfCp;
        }
        if(pcchPreChars)
        {
            *pcchPreChars = msg.resultsHitTest._cchPreChars;
        }
    }
    else
    {
        CPoint ptLocal(pt);
        // We now need to convert pt to client coordinates from global coordinates
        // before we can call CpFromPoint...
        pFlowLayout->TransformPoint(&ptLocal, COORDSYS_GLOBAL, COORDSYS_CONTENT);
        cpHit = CpFromPoint(ptLocal, prp, pptp, NULL, dwCFPFlags,
            pfRightOfCp, NULL, pcchPreChars, NULL);
    }
    return cpHit;
}

LONG CDisplay::CpFromPoint(
                      POINT         pt,             // Point to compute cp at (site coords)
                      CLinePtr* const prp,          // Returns line pointer at cp (may be NULL)
                      CTreePos**    pptp,           // pointer to return TreePos corresponding to the cp
                      CLayout**     ppLayout,       // can be NULL
                      DWORD         dwFlags,
                      BOOL*         pfRightOfCp,
                      BOOL*         pfPseudoHit,
                      LONG*         pcchPreChars,
                      CCalcInfo*    pci)
{
    CCalcInfo   CI;
    CRect       rc;
    LONG        ili;
    LONG        cp;
    LONG        yLine;

    if(!pci)
    {
        CI.Init(GetFlowLayout());
        pci = &CI;
    }

    // Get line under hit
    GetFlowLayout()->GetClientRect(&rc);

    rc.MoveTo(pt);

    ili = LineFromPos(
        rc, &yLine, &cp,
        LFP_ZORDERSEARCH|LFP_IGNORERELATIVE|
        (!ppLayout ? LFP_IGNOREALIGNED: 0) |
        (dwFlags & CFP_NOPSEUDOHIT?LFP_EXACTLINEHIT:0));
    if(ili < 0)
    {
        return -1;
    }

    return CpFromPointEx(ili, yLine, cp, pt, prp, pptp, ppLayout,
        dwFlags, pfRightOfCp, pfPseudoHit,
        pcchPreChars, pci);

}

LONG CDisplay::CpFromPointEx(
                        LONG        ili,
                        LONG        yLine,
                        LONG        cp,
                        POINT       pt,             // Point to compute cp at (site coords)
                        CLinePtr*const prp,         // Returns line pointer at cp (may be NULL)
                        CTreePos**  pptp,           // pointer to return TreePos corresponding to the cp
                        CLayout**   ppLayout,       // can be NULL
                        DWORD       dwFlags,
                        BOOL*       pfRightOfCp,
                        BOOL*       pfPseudoHit,
                        LONG*       pcchPreChars,
                        CCalcInfo*  pci)
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CElement*       pElementFL  = pFlowLayout->ElementOwner();
    CCalcInfo       CI;
    CLine*          pli         = Elem(ili);
    LONG            cch         = 0;
    LONG            dx          = 0;
    BOOL            fPseudoHit  = FALSE;
    CTreePos*       ptp         = NULL;
    CTreeNode*      pNode;
    LONG            cchPreChars = 0;

    if(!pci)
    {
        CI.Init(pFlowLayout);
        pci = &CI;
    }

    if(dwFlags&CFP_IGNOREBEFOREAFTERSPACE
        && (pli==NULL || (!pli->_fRelative && pli->_fSingleSite)))
    {
        return -1 ;
    }

    if(pli)
    {
        if(pli->IsFrame())
        {
            if(ppLayout)
            {
                *ppLayout = pli->_pNodeLayout->GetUpdatedLayout();
            }
            cch = 0;

            if(pfRightOfCp)
            {
                *pfRightOfCp = TRUE;
            }

            fPseudoHit = TRUE;
        }
        else
        {
            if(!(dwFlags&CFP_IGNOREBEFOREAFTERSPACE)
                && pli->_fHasNestedRunOwner
                && yLine+pli->_yHeight<=pt.y)
            {
                // If the we are not ignoring whitespaces and we have hit a line
                // which has a nested runowner, but are BELOW the line (happens when
                // that line is the last line in the document) then we want
                // to be at the end of that line. The measurer would put us at the
                // beginning or end depending upon the X position.
                cp += pli->_cch;
            }
            else
            {
                // Create measurer
                CLSMeasurer me(this, pci);
                LONG yHeightRubyBase = 0;

                AssertSz((pli!=NULL) || (ili==0),
                    "CDisplay::CpFromPoint invalid line pointer");

                if(!me._pLS)
                {
                    return -1;
                }

                // Get character in the line
                me.SetCp(cp, NULL);

                // The y-coordinate should be relative to the baseline, and positive going up
                cch = pli->CchFromXpos(
                    me, pt.x,
                    yLine+pli->_yHeight-pli->_yDescent-pt.y,
                    &dx, 
                    _fRTL, dwFlags&CFP_EXACTFIT, &yHeightRubyBase);
                cchPreChars = me._cchPreChars;

                if(pfRightOfCp)
                {
                    *pfRightOfCp = dx < 0;
                }

                if(ppLayout)
                {
                    ptp = me.GetPtp();
                    if(ptp->IsBeginElementScope())
                    {
                        pNode = ptp->Branch();
                        if(pNode->NeedsLayout() && pNode->IsInlinedElement())
                        {
                            *ppLayout = pNode->GetUpdatedLayout();
                        }
                        else
                        {
                            *ppLayout = NULL;
                        }
                    }
                }

                // Don't allow click at EOL to select EOL marker and take into account
                // single line edits as well
                if(!(dwFlags&CFP_ALLOWEOL) && cch==pli->_cch)
                {
                    long cpPtp;

                    ptp = me.GetPtp();
                    Assert(ptp);

                    cpPtp = ptp->GetCp();

                    // cch > 0 && we are not in the middle of a text run,
                    // skip past all the previous node/pointer tree pos's.
                    // and position the measurer just after text.
                    if(cp<cpPtp && cpPtp==me.GetCp())
                    {
                        while(cp < cpPtp)
                        {
                            CTreePos* ptpPrev = ptp->PreviousTreePos();

                            if(!ptpPrev->GetBranch()->IsDisplayNone())
                            {
                                if(ptpPrev->IsText())
                                {
                                    break;
                                }
                                if(ptpPrev->IsEndElementScope()
                                    && ptpPrev->Branch()->NeedsLayout())
                                {
                                    break;
                                }
                            }
                            ptp = ptpPrev;
                            Assert(ptp);
                            cch   -= ptp->GetCch();
                            cpPtp -= ptp->GetCch();
                        }
                        
                        while(ptp->GetCch()==0)
                        {
                            ptp = ptp->NextTreePos();
                        }
                    }
                    else if(pElementFL->GetFirstBranch()->GetParaFormat()->HasInclEOLWhite(TRUE))
                    {
                        CTxtPtr tpTemp(GetMarkup(), cp+cch);
                        if(tpTemp.GetPrevChar() == WCH_CR)
                        {
                            cch--;
                            if(cp+cch < cpPtp)
                            {
                                ptp = NULL;
                            }
                        }
                    }

                    me.SetCp(cp+cch, ptp);
                }

                // Check if the pt is within bounds *vertically* too.
                if(dwFlags & CFP_IGNOREBEFOREAFTERSPACE)
                {
                    LONG top, bottom;

                    ptp = me.GetPtp();
                    if(ptp->IsBeginElementScope()
                        && ptp->Branch()->NeedsLayout()
                        && ptp->Branch()->IsInlinedElement())
                    {
                        // Hit a site. Check if we are within the boundaries
                        // of the site.
                        RECT rc;
                        CLayout* pLayout = ptp->Branch()->GetUpdatedLayout();

                        // BUGBUG: This is wrong, fix it (brendand)
                        pLayout->GetRect(&rc);

                        top = rc.top;
                        bottom = rc.bottom;
                    }
                    else
                    {
                        GetTopBottomForCharEx(
                            pci,
                            &top,
                            &bottom,
                            yLine,
                            pli,
                            pt.x,
                            me.CurrBranch()->GetCharFormat());
                        top -= yHeightRubyBase;
                        bottom -= yHeightRubyBase;
                    }

                    // BUGBUG (t-ramar): this is fine for finding 99% of the
                    // pseudo hits, but if someone has a ruby with wider base
                    // text than pronunciation text, or vice versa, hits in the
                    // whitespace that results will not register as pseudo hits.
                    if(pt.y<top
                        || pt.y>=bottom
                        || (!_fRTL
                        && (pt.x<pli->GetTextLeft()
                        || pt.x>=pli->GetTextRight(ili==LineCount()-1)))
                        || (_fRTL
                        && (-pt.x<pli->GetRTLTextRight()
                        || -pt.x>=pli->GetRTLTextLeft())))
                    {
                        fPseudoHit = TRUE;
                    }
                }

                cp = (LONG)me.GetCp();

                ptp = me.GetPtp();
            }
        }
    }

    if(prp)
    {
        prp->RpSet(ili, cch);
    }

    if(pfPseudoHit)
    {
        *pfPseudoHit = fPseudoHit;
    }

    if(pcchPreChars)
    {
        *pcchPreChars = cchPreChars;
    }

    if(pptp)
    {
        LONG ich;

        *pptp = ptp ? ptp : pFlowLayout->GetContentMarkup()->TreePosAtCp(cp, &ich);
    }

    return cp;
}

//+----------------------------------------------------------------------------
//
// Member:      CDisplay::PointFromTp(tp, prcClient, fAtEnd, pt, prp, taMode,
//                                    pci, pfComplexLine, pfRTLFlow)
//
// Synopsis:    return the origin that corresponds to a given text pointer,
//              relative to the view
//
//-----------------------------------------------------------------------------
LONG CDisplay::PointFromTp(
                           LONG cp,                 // point for the cp to be computed
                           CTreePos* ptp,           // tree pos for the cp passed in, can be NULL
                           BOOL fAtEnd,             // Return end of previous line for ambiguous cp
                           BOOL fAfterPrevCp,       // Return the trailing point of the previous cp (for an ambigous bidi cp)
                           POINT& pt,               // Returns point at cp in client coords
                           CLinePtr* const prp,     // Returns line pointer at tp (may be null)
                           UINT taMode,             // Text Align mode: top, baseline, bottom
                           CCalcInfo* pci,
                           BOOL* pfComplexLine,
                           BOOL* pfRTLFlow)
{
    CFlowLayout* pFL = GetFlowLayout();
    CLinePtr    rp(this);
    BOOL        fLastTextLine;
    CCalcInfo   CI;
    BOOL        fRTLLine;
    BOOL        fAtEndOfRubyBase;
    RubyInfo rubyInfo = { -1, 0, 0 };

    // If cp points to a node position somewhere between ruby base and ruby text, 
    // then we have to remember for later use that we are at the end of ruby text
    // (fAtEndOfRubyBase is set to TRUE).
    // If cp points to a node at the and of ruby element, then we have to move
    // cp to the beginning of text which follows ruby or to the beginning of block
    // element, whichever is first.
    fAtEndOfRubyBase = FALSE;

    CTreePos* ptpNode = ptp ? ptp : GetMarkup()->TreePosAtCp(cp, NULL);
    LONG cpOffset = 0;
    while(ptpNode && ptpNode->IsNode())
    {
        ELEMENT_TAG eTag = ptpNode->Branch()->Element()->Tag();
        if(eTag == ETAG_RT)
        {
            fAtEndOfRubyBase = ptpNode->IsBeginNode();
            break;
        }
        else if(eTag == ETAG_RP)
        {
            if(ptpNode->IsBeginNode())
            {
                // At this point we check where the RP tag is located: before or 
                // after RT tag. If this is before RT tag then we set fAtEndOfRubyBase
                // to TRUE, in the other case we do nothing.
                //
                // If RT tag is a parent of RP tag then we are after RT tag. 
                // In the other case we are before RT tag.
                CTreeNode* pParent = ptpNode->Branch()->Parent();
                while(pParent)
                {
                    if(pParent->Tag() == ETAG_RUBY)
                    {
                        fAtEndOfRubyBase = TRUE;
                        break;
                    }
                    else if(pParent->Tag() == ETAG_RT)
                    {
                        break;
                    }
                    pParent = pParent->Parent();
                }
            }
            break;
        }
        else if(ptpNode->IsEndNode() && eTag==ETAG_RUBY)
        {
            for(;;)
            {
                cpOffset++;
                ptpNode = ptpNode->NextTreePos();
                if(!ptpNode || !ptpNode->IsNode() || 
                    ptpNode->Branch()->Element()->IsBlockElement())
                {
                    cp += cpOffset;
                    break;
                }
            }
            break;
        }
        cpOffset++;
        ptpNode = ptpNode->NextNonPtrTreePos();
    }

    if(!pci)
    {
        CI.Init(pFL);
        pci = &CI;
    }

    if(pFL->IsDisplayNone() || !WaitForRecalc(cp, -1, pci))
    {
        return -1;
    }

    if(!rp.RpSetCp(cp, fAtEnd))
    {
        return -1;
    }

    if(!WaitForRecalc(min(cp+rp->_cch, GetLastCp()), -1, pci))
    {
        return -1;
    }

    if(!rp.RpSetCp(cp, fAtEnd))
    {
        return -1;
    }

    if(rp.IsValid())
    {
        fRTLLine = rp->_fRTL;
    }
    else if(ptp)
    {
        fRTLLine = ptp->GetBranch()->GetParaFormat()->HasRTL(TRUE);
    }
    else
    {
        fRTLLine = _fRTL;
    }

    if(pfRTLFlow)
    {
        *pfRTLFlow = fRTLLine;
    }

    pt.y = YposFromLine(rp, pci);

    if(!_fRTL)
    {
        pt.x = rp.IsValid()
            ? rp->_xLeft+rp->_xLeftMargin+(fRTLLine&&!(rp.RpGetIch())?rp->_xWidth:0) : 0;
    }
    else
    {
        pt.x = rp.IsValid()
            ? rp->_xRight+rp->_xRightMargin+(!fRTLLine&&!(rp.RpGetIch())?rp->_xWidth:0) : 0;
    }
    fLastTextLine = rp.IsLastTextLine();

    if(rp.RpGetIch())
    {
        CLSMeasurer me(this, pci);
        if(!me._pLS)
        {
            return -1;
        }

        // Backup to start of line
        if(rp.GetIch())
        {
            me.SetCp(cp-rp.RpGetIch(), NULL);
        }
        else
        {
            me.SetCp(cp, ptp);
        }

        // And measure from there to where we are
        me.NewLine(*rp);

        Assert(rp.IsValid());

        me._li._xLeft = rp->_xLeft;
        me._li._xLeftMargin = rp->_xLeftMargin;
        me._li._xWidth = rp->_xWidth;
        me._li._xRight = rp->_xRight;
        me._li._xRightMargin = rp->_xRightMargin;
        me._li._fRTL = rp->_fRTL;

        // can we also add the _cch and _cchWhite here so we can pass them to 
        // the BidiLine stuff?
        me._li._cch = rp->_cch;
        me._li._cchWhite = rp->_cchWhite;

        LONG xCalc = me.MeasureText(rp.RpGetIch(), rp->_cch, fAfterPrevCp, pfComplexLine, pfRTLFlow, &rubyInfo);

        // Remember we ignore trailing spaces at the end of the line
        // in the width, therefore the x value that MeasureText finds can
        // be greater than the width in the line so we truncate to the
        // previously calculated width which will ignore the spaces.
        // pt.x += min(xCalc, rp->_xWidth);
        //
        // Why anyone would want to ignore the trailing spaces at the end
        // of the line is beyond me. For certain, we DON'T want to ignore
        // them when placing the caret at the end of a line with trailing
        // spaces. If you can figure out a reason to ignore the spaces,
        // please do, just leave the caret placement case intact. - Arye
        if(_fRTL == (unsigned)fRTLLine)
        {
            pt.x += xCalc;
        }
        else
        {
            pt.x += rp->_xWidth - xCalc;
        }
    }

    if(prp)
    {
        *prp = rp;
    }

    if(rp>=0 && taMode!=TA_TOP)
    {
        // Check for vertical calculation request
        if(taMode & TA_BOTTOM)
        {
            const CLine* pli = Elem(rp);

            if(pli)
            {
                pt.y += pli->_yHeight;
                if(rubyInfo.cp!=-1 && !fAtEndOfRubyBase)
                {
                    pt.y -= rubyInfo.yHeightRubyBase + pli->_yDescent - pli->_yTxtDescent;
                }

                if(!pli->IsFrame() && (taMode&TA_BASELINE)==TA_BASELINE)
                {
                    if(rubyInfo.cp!=-1 && !fAtEndOfRubyBase)
                    {
                        pt.y -= rubyInfo.yDescentRubyText;
                    }
                    else
                    {
                        pt.y -= pli->_yDescent;
                    }
                }
            }
        }

        // Do any specical horizontal calculation
        if(taMode & TA_CENTER)
        {
            CLSMeasurer me(this, pci);

            if(!me._pLS)
            {
                return -1;
            }

            me.SetCp(cp, ptp);

            me.NewLine(*rp);

            pt.x += (me.MeasureText(1, rp->_cch) >> 1);
        }
    }

    return rp;
}

//+----------------------------------------------------------------------------
//
// Function:    AppendRectToElemRegion
//
// Synopsis:    Utility function for region from element, which appends the
//              given rect to region if it is within the clip range and
//              updates the bounding rect.
//
//-----------------------------------------------------------------------------
void AppendRectToElemRegion(
                            CDataAry <RECT>* paryRects,
                            RECT* prcBound,
                            RECT* prcLine,
                            CPoint& ptTrans,
                            LONG cp,
                            LONG cpClipStart,
                            LONG cpClipFinish)
{
    if(ptTrans.x || ptTrans.y)
    {
        OffsetRect(prcLine, ptTrans.x, ptTrans.y);
    }

    if(cp>=cpClipStart && cp<=cpClipFinish)
    {
        paryRects->AppendIndirect(prcLine);
    }

    if(prcBound)
    {
        if(!IsRectEmpty(prcLine))
        {
            UnionRect(prcBound, prcBound, prcLine);
        }
        else if(paryRects->Size() == 1)
        {
            *prcBound = *prcLine;
        }
    }
}

//+----------------------------------------------------------------------------
//
// Function:    RcFromAlignedLine
//
// Synopsis:    Utility function for region from element, which computes the
//              rect for a given aligned line
//
//-----------------------------------------------------------------------------
void RcFromAlignedLine(
                       RECT* prcLine,
                       CLine* pli,
                       LONG yPos,
                       BOOL fBlockElement,
                       BOOL fFirstLine, BOOL fRTL,
                       long xParentLeftIndent,
                       long xParentRightIndent)
{
    CSize size;
    CLayout* pLayout = pli->_pNodeLayout->GetUpdatedLayout();

    long xProposed = pLayout->GetXProposed();
    long yProposed = pLayout->GetYProposed();

    pLayout->GetSize(&size);

    // add the curent line to the region
    prcLine->top = yProposed + yPos + pli->_yBeforeSpace;

    if(!fRTL)
    {
        prcLine->left = xProposed + pli->_xLeftMargin + pli->_xLeft;
        prcLine->right = prcLine->left + size.cx;
    }
    else
    {
        prcLine->right = -(xProposed + pli->_xRightMargin + pli->_xRight);
        prcLine->left = prcLine->right - size.cx;
    }
    prcLine->bottom = prcLine->top + size.cy;
}

//+----------------------------------------------------------------------------
//
// Function:    ComputeIndentsFromParentNode
//
// Synopsis:    Compute the indent for a given Node and a left and/or
//              right aligned site that a current line is aligned to.
//
//-----------------------------------------------------------------------------
void CDisplay::ComputeIndentsFromParentNode(
    CCalcInfo*  pci,            // (IN)
    CTreeNode*  pNode,          // (IN) node we want to compute indents for
    DWORD       dwFlags,        // (IN) flags from RFE
    LONG*       pxLeftIndent,   // (OUT) the node is indented this many pixels from left edge of the layout
    LONG*       pxRightIndent)  // (OUT) ...
{
    CElement*   pElement = pNode->Element();
    CElement*   pElementFL = GetFlowLayoutElement();
    LONG        xParentLeftPadding, xParentRightPadding;
    LONG        xParentLeftBorder, xParentRightBorder;

    const CParaFormat* pPF = pNode->GetParaFormat();
    BOOL        fInner = pNode->Element() == pElementFL;

    Assert(pNode);

    // GetLeft/RightIndent returns the cumulative CSS margin (+ some other gunk
    // like list bullet offsets).
    LONG        xLeftMargin  = pPF->GetLeftIndent(pci, fInner);
    LONG        xRightMargin = pPF->GetRightIndent(pci, fInner);

    // We only want to include the area for the bullet for hit-testing;
    // we _don't_ draw backgrounds and borders around the bullet for list items.
    if(dwFlags==RFE_HITTEST
        && pPF->_bListPosition != styleListStylePositionInside
        && (pElement->IsFlagAndBlock(TAGDESC_LISTITEM)
        || pElement->IsFlagAndBlock(TAGDESC_LIST)))
    {
        if(!pPF->HasRTL(fInner))
        {
            xLeftMargin -= pPF->GetBulletOffset(pci, fInner);
        }
        else
        {
            xRightMargin -= pPF->GetBulletOffset(pci, fInner);
        }
    }

    // Compute the padding and border space cause by the current
    // element's ancestors (up to the layout).
    if(pNode->Element() == pElementFL)
    {
        // If the element in question is actually the layout owner,
        //b then we don't want to offset by our own border/padding,
        // so set the values to 0.  
        xParentLeftPadding = xParentLeftBorder =
            xParentRightPadding = xParentRightBorder = 0;
    }
    else
    {
        // We need to get the cumulative sum of our ancestor's borders/padding.
        pNode->Parent()->Element()->ComputeHorzBorderAndPadding(
            pci, pNode->Parent(), pElementFL,
            &xParentLeftBorder, &xParentLeftPadding,
            &xParentRightBorder, &xParentRightPadding);

        // The return results of ComputeHorzBorderAndPadding() DO NOT include
        // the border or padding of the layout itself; this makes sense for
        // borders, because the layout's border is outside the bounds we're
        // interested in.  However, we do want to account for the layout's
        // padding since that's inside the bounds.  We fetch that separately
        // here, and add it to the cumulative padding.
        long lPadding[PADDING_MAX];
        GetPadding(pci, lPadding, pci->_smMode==SIZEMODE_MMWIDTH);

        xParentLeftPadding += lPadding[PADDING_LEFT];
        xParentRightPadding += lPadding[PADDING_RIGHT];
    }

    // The element is indented by the sum of CSS margins and ancestor
    // padding/border.  This indent value specifically ignores aligned/floated
    // elements, per CSS!
    *pxLeftIndent = xLeftMargin + xParentLeftBorder + xParentLeftPadding;
    *pxRightIndent = xRightMargin + xParentRightBorder + xParentRightPadding;
}

//+----------------------------------------------------------------------------
//
// Member:      RegionFromElement
//
// Synopsis:    for a given element, find the set of rects (or lines) that this
//              element occupies in the display. The rects returned are relative
//              the site origin.
//              Certain details about the returned region are determined by the
//              call type parameter:..
//
//-----------------------------------------------------------------------------
void CDisplay::RegionFromElement(
                                 CElement*  pElement,       // (IN)
                                 CDataAry<RECT>* paryRects, // (OUT)
                                 CPoint*    pptOffset,      // == NULL, point to offset the rects by (IN param)
                                 CFormDrawInfo* pDI,        // == NULL
                                 DWORD      dwFlags,        // == 0
                                 LONG       cpClipStart,    // == -1, (IN)
                                 LONG       cpClipFinish,   // == -1, (IN) clip range
                                 RECT*      prcBound)       // == NULL, (OUT param) returns bounding rect that ignores clipping

{
    CFlowLayout*    pFL = GetFlowLayout();
    CElement*       pElementFL = pFL->ElementOwner();
    CTreePos*       ptpStart;
    CTreePos*       ptpFinish;
    CCalcInfo       CI;
    RECT            rcLine;
    CPoint          ptTrans = _afxGlobalData._Zero.pt;
    BOOL            fScreenCoord = dwFlags&RFE_SCREENCOORD ? TRUE : FALSE;

    Assert(pElement->IsInMarkup());

    // clear the array before filling it
    paryRects->SetSize(0);

    if(prcBound)
    {
        memset(prcBound, 0, sizeof(RECT));
    }

    if(pElementFL->IsDisplayNone() || !pElement || pElement->IsDisplayNone() ||
        pElement->Tag()==ETAG_ROOT || pElementFL->Tag()==ETAG_ROOT)
    {
        return;
    }

    if(pElement == pElementFL)
    {
        pFL->GetContentTreeExtent(&ptpStart, &ptpFinish);
    }
    else
    {
        pElement->GetTreeExtent(&ptpStart, &ptpFinish);
    }
    Assert(ptpStart && ptpFinish);

    // now that we have the scope of the element, find its range.
    // and compute the rects (lines) that contain the range.
    {
        CTreePos*   ptpElemStart;
        CTreePos*   ptpElemFinish;
        CTreeNode*  pNode = pElement->GetFirstBranch();
        CLSMeasurer me(this, pDI);
        CMarkup*    pMarkup = pFL->GetContentMarkup();
        BOOL        fBlockElement;
        BOOL        fTightRects = (dwFlags & RFE_TIGHT_RECTS);
        BOOL        fIgnoreBlockness = (dwFlags & RFE_IGNORE_BLOCKNESS);
        BOOL        fIncludeAlignedElements = !(dwFlags & RFE_IGNORE_ALIGNED);
        BOOL        fIgnoreClipForBoundRect = (dwFlags & RFE_IGNORE_CLIP_RC_BOUNDS);
        BOOL        fIgnoreRel = (dwFlags & RFE_IGNORE_RELATIVE);
        BOOL        fNestedRel = (dwFlags & RFE_NESTED_REL_RECTS);
        BOOL        fScrollIntoView = (dwFlags & RFE_SCROLL_INTO_VIEW);
        BOOL        fNeedToMeasureLine;
        LONG        cp, cpStart, cpFinish, cpElementLast;
        LONG        cpElementStart, cpElementFinish;
        LONG        iCurLine, iFirstLine, ich;
        LONG        yPos;
        LONG        xParentRightIndent = 0;
        LONG        xParentLeftIndent = 0;
        LONG        yParentPaddBordTop = 0;
        LONG        yParentPaddBordBottom = 0;
        LONG        yTop;
        BOOL        fFirstLine;
        CLinePtr    rp(this);
        BOOL        fRestorePtTrans = FALSE;
        CStackDataAry<RECT, 8> aryChunks;

        if(!me._pLS)
        {
            goto Cleanup;
        }

        // Do we treat the element we're finding the region for
        // as a block element?  Things that influence this decision:
        // If RFE_SELECTION was specified, it means we're doing
        // finding a region for selection, which is always character
        // based (never block based).  RFE_SELECTION causes fIgnoreBlockness to light up.
        // The only time RFE should get called on an element that
        // needs layout is when the pElement == pElementFL.  When
        // this happens, even though the element _is_ block, we
        // want to treat it as though it isn't, since the caller
        // in this situation wants the rects of some text that's
        // directly under the layout (e.g. "x" in <BODY>x<DIV>...</DIV></BODY>)
        // BUGBUG: (KTam) this would be a lot more obvious if we changed
        // the !pElement->NeedsLayout() condition to pElement != pElementFL
        fBlockElement =  !fIgnoreBlockness &&   
            pElement->IsBlockElement() && !pElement->NeedsLayout();   

        if(pDI)
        {
            CI.Init(pDI, pFL);
        }
        else
        {
            CI.Init(pFL);
        }

        long xParentWidth;
        long yParentHeight;
        // Fetch parent's width and height (i.e. view width & height minus padding)
        GetViewWidthAndHeightForChild(&CI, &xParentWidth, &yParentHeight);
        CI.SizeToParent(xParentWidth, yParentHeight);

        // If caller hasn't specified non-relative behaviour, then account for
        // relativeness on the layout by fetching the relative offset in ptTrans.
        if(!fIgnoreRel)
        {
            pNode->GetRelTopLeft(pElementFL, &CI, &ptTrans.x, &ptTrans.y);
        }

        // Transform the relative offset to global coords if caller specified
        // RFE_SCREENCOORD.        
        if(fScreenCoord)
        {
            pFL->TransformPoint(&ptTrans, COORDSYS_CONTENT, COORDSYS_GLOBAL);
        }

        // Combine caller-specified offset (if any) into relative offset.
        if(pptOffset)
        {
            ptTrans.x += pptOffset->x;
            ptTrans.y += pptOffset->y;
        }

        cpStart  = pFL->GetContentFirstCp();
        cpFinish = pFL->GetContentLastCp();        

        // get the cp range for the element
        cpElementStart  = ptpStart->GetCp();
        cpElementFinish = ptpFinish->GetCp();

        // Establish correct cp's and treepos's, including clipping stuff..
        // We may have elements overlapping multiple layout's, so clip the cp range
        // to the current flow layout's bounds,
        cpStart       = max(cpStart, cpElementStart);
        cpFinish      = min(cpFinish, cpElementFinish);
        cpElementLast = cpFinish;

        // clip cpFinish to max calced cp
        cpFinish      = min(cpFinish, GetMaxCpCalced());

        if(cpStart != cpElementStart)
        {
            ptpStart  = pMarkup->TreePosAtCp(cpStart, &ich);
        }
        ptpElemStart = ptpStart;

        if(cpFinish != cpElementFinish)
        {
            ptpFinish = pMarkup->TreePosAtCp(cpFinish, &ich);
        }
        ptpElemFinish = ptpFinish;

        if(cpClipStart < cpStart)
        {
            cpClipStart = cpStart;
        }
        if(cpClipFinish==-1 || cpClipFinish>cpFinish)
        {
            cpClipFinish = cpFinish;
        }

        if(!fIgnoreClipForBoundRect)
        {
            if(cpStart != cpClipStart)
            {
                cpStart  = cpClipStart;
                ptpStart = pMarkup->TreePosAtCp(cpStart, &ich);
            }
            if(cpFinish != cpClipFinish)
            {
                cpFinish = cpClipFinish;
                ptpFinish = pMarkup->TreePosAtCp(cpFinish, &ich);
            }
        }

        if(!LineCount())
        {
            return;
        }

        // BUGBUG: we pass in absolute cp here so RpSetCp must call
        // pElementFL->GetFirstCp while we have just done this.  We
        // should optimize this out by having a relative version of RpSetCp.

        // skip frames and position it at the beginning of the line
        rp.RpSetCp(cpStart, /* fAtEnd = */FALSE, TRUE);

        // (cpStart - rp.GetIch) == cp of beginning of line (pointed to by rp)
        // If fFirstLine is true, it means that the first line of the element
        // is within the clipping range (accounted for by cpStart).
        fFirstLine = cpElementStart >= cpStart - rp.GetIch();

        // BUGBUG: (KTam) re-enable this assert when Sujal fixes
        // CpFromPointEx().
        // Assert( cpStart != cpFinish );
        // Update 12/09/98: It is valid for cpStart == cpFinish
        // under some circumstances!  When a hyperlink jump is made to
        // a local anchor (#) that's empty (e.g. <A name="foo"></A>, 
        // the clipping range passed in is empty! (HOT bug 60358)
        // If the element has no content then return the null rectangle
        if(cpStart == cpFinish)
        {
            CLine* pli = Elem(rp);

            yPos = YposFromLine(rp, &CI);

            rcLine.top = yPos + rp->GetYTop();
            rcLine.bottom = yPos + rp->GetYBottom();

            rcLine.left = rp->GetTextLeft();

            if(rp.GetIch())
            {
                me.SetCp(cpStart-rp.GetIch(), NULL);
                me._li = *pli;

                rcLine.left +=  me.MeasureText(rp.RpGetIch(), pli->_cch);
            }
            rcLine.right = rcLine.left;

            AppendRectToElemRegion(
                paryRects, prcBound, &rcLine,
                ptTrans,
                cpStart, cpClipStart, cpClipFinish);
            return;
        }

        // Compute the padding and border space for the first line
        // and the last line for ancestors that are entering and leaving
        // scope respectively.  For example, consider an element x such that:

        // < elem 1 >
        //   < elem 2 >
        //     < ..... >
        //       < elem x >
        //         <!-- some content -->
        //       </ elem x >
        //     </ ..... >
        //   </ elem 2 >
        // </ elem 1 >
        // 
        // I.e., elem x is immediately preceded by some number of other
        // elements _entering_ scope (with no content in between them).
        // The opening of such elements may introduce vertical space in
        // terms of border and padding; we want to compute that so we
        // can adjust for it when determining the rect of the first line
        // of elem x.
        // 
        // A similar situation exists with respect to elements _leaving_
        // scope immediately following the closing of element x.
        if(pElement && fBlockElement && pElement!=pElementFL)
        {
            ComputeVertPaddAndBordFromParentNode(
                &CI, ptpStart, ptpFinish,
                &yParentPaddBordTop, &yParentPaddBordBottom);
        }

        iCurLine = iFirstLine = rp;

        // Now that we know what line contains the cp, let's try to compute the
        // bounding rectangle.

        // If the line has aligned images in it and if they are at the begining of the
        // the line, back up all the frame lines that are before text and contained
        // within the element of interest.  This only necessary when doing selection
        // and hit testing, since the region for borders/backgrounds ignores aligned
        // and floated stuff.

        // (recall rp is like a smart ptr that points to a particular cp by
        // holding a line index (ili) AND an offset into the line's chars
        // (ich), and that we called CLinePtr::RpSetCp() some time back.

        // (recall frame lines are those that were created expressly for
        // aligned elems; they are the result of "breaking" a line, which
        // we had to do when we were measuring and saw it had aligned elems)
        if(fIncludeAlignedElements  && rp->HasAligned())
        {
            LONG diLine = -1;
            CLine* pli;

            // rp[x] means "Given that rp is indexing line y, return the
            // line indexed by y - x".

            // A line is "blank" if it's "clear" or "frame".  Here we
            // walk backwards through the line array until all consecutive
            // frame lines before text are passed.
            while((iCurLine+diLine>=0) && (pli=&rp[diLine]) && pli->IsBlankLine())
            {
                // Stop walking back if we've found a frame line that _isn't_ "before text",
                // or one whose ending cp is prior to the beginning of our clipping range
                // (which means it's not contained in the element of interest)
                // Consider: <div><img align=left><b><img align=left>text</b>
                // The region for the bold element includes the 2nd image but
                // not the 1st, but both frame lines will show up in front of the bold line.
                // The logic below notes that the last cp of the 1st img is before the cpStart
                // of the bold element, and breaks so we don't include it.
                if(pli->IsFrame() &&
                    (!pli->_fFrameBeforeText || pli->_pNodeLayout->Element()->GetLastCp()<cpStart))
                {
                    break;
                }

                diLine--;
            }
            iFirstLine = iCurLine + diLine + 1;
        }

        // compute the ypos for the first line
        yTop = yPos = YposFromLine(iFirstLine, &CI);

        // For calls other than backgrounds/borders, add all the frame lines
        // before the current line under the influence of
        // the element to the region.
        if(fIncludeAlignedElements)
        {
            for(; iFirstLine<iCurLine; iFirstLine++)
            {
                CLine* pli = Elem(iFirstLine);

                // If the element of interest is block level, find out
                // how much it's indented (left and right) in from its
                // parent layout.
                if(fBlockElement)
                {
                    CTreeNode* pNodeTemp = pMarkup->SearchBranchForScopeInStory(ptpStart->GetBranch(), pElement);
                    if(pNodeTemp)
                    {
                        //Assert(pNodeTemp);
                        ComputeIndentsFromParentNode(
                            &CI, pNodeTemp, dwFlags,
                            &xParentLeftIndent,
                            &xParentRightIndent);
                    }
                }

                // If it's a frame line, add the line's rc to the region.
                if(pli->IsFrame())
                {
                    long cpLayoutStart  = pli->_pNodeLayout->Element()->GetFirstCp();
                    long cpLayoutFinish = pli->_pNodeLayout->Element()->GetLastCp();

                    if(cpLayoutStart>=cpStart && cpLayoutFinish<=cpFinish)
                    {
                        RcFromAlignedLine(
                            &rcLine, pli, yPos,
                            fBlockElement, fFirstLine, _fRTL,
                            xParentLeftIndent, xParentRightIndent);

                        AppendRectToElemRegion(
                            paryRects, prcBound, &rcLine,
                            ptTrans,
                            cpFinish, cpClipStart, cpClipFinish);
                    }
                }
                // if it's not a frame line, it MUST be a clear line (since we
                // only walked by while seeing blank lines).  All clear lines
                // force a new physical line, so account for change in yPos.
                else
                {
                    Assert(pli->IsClear() && pli->_fForceNewLine);
                    yPos += pli->_yHeight;
                }
            }
        }

        // now add all the lines that are contained by the range
        for(cp=cpStart; cp<cpFinish; iCurLine++)
        {
            BOOL    fRTLLine;
            LONG    xStart = 0;
            LONG    xEnd = 0;
            long    yTop, yBottom;
            LONG    cchAdvance = min(rp.GetCchRemaining(), cpFinish-cp);
            CLine*  pli;
            LONG    i;
            LONG    cChunk;
            CPoint  ptTempForPtTrans(0, 0);

            Assert(!fRestorePtTrans);
            if(iCurLine >= LineCount())
            {
                // (srinib) - please check with one of the text
                // team members if anyone hits this assert. We want
                // to investigate. We shoud never be here unless
                // we run out of memory.
                AssertSz(FALSE, "WaitForRecalc failed to measure lines till cpFinish");
                break;
            }

            pli = Elem(iCurLine);

            // When drawing backgrounds, we skip frame lines contained
            // within the element (since aligned stuff doesn't affect our borders/background)
            // but for hit testing and selection we need to account for them.
            if(!fIncludeAlignedElements && pli->IsFrame())
            {
                goto AdvanceToNextLineInRange;
            }

            // If line is relative then add in the relative offset
            // In the case of Wigglys we ignore the line's relative positioning, but want any nested
            // elements to be handled CLineServices::CalcRectsOfRangeOnLine
            if(fNestedRel && pli->_fRelative)
            {
                CPoint ptTemp;
                CTreeNode* pNodeNested = pMarkup->TreePosAtCp(cp, &ich)->GetBranch();

                // We only want to adjust for nested elements that do not have layouts. The pElement
                // we passed in is handled by the fIgnoreRel flag. Layout elements are handled in
                // CalcRectsOfRegionOnLine
                if(pNodeNested->Element()!=pElement && !pNodeNested->HasLayout())
                {
                    pNodeNested->GetRelTopLeft(pElementFL, &CI, &ptTemp.x, &ptTemp.y);
                    ptTempForPtTrans = ptTrans;
                    if(!fIgnoreRel)
                    {
                        ptTrans.x = ptTemp.x - ptTrans.x;
                        ptTrans.y = ptTemp.y - ptTrans.y;
                    }
                    else
                    {
                        // We were told to ignore the pElement's relative positioning. Therefore
                        // we only want to adjust by the amount of the nested element's relative
                        // offset from pElement
                        long xElemRelLeft = 0, yElemRelTop = 0;
                        pNode->GetRelTopLeft(pElementFL, &CI, &xElemRelLeft, &yElemRelTop);
                        ptTrans.x = ptTemp.x - xElemRelLeft;
                        ptTrans.y = ptTemp.y - yElemRelTop;
                    }
                    fRestorePtTrans = TRUE;
                }
            }

            fRTLLine = pli->_fRTL;

            // If the element of interest is block level, find out
            // how much it's indented (left and right) in from its
            // parent layout.
            if(fBlockElement)
            {
                if(cp != cpStart)
                {
                    ptpStart = pMarkup->TreePosAtCp(cp, &ich);
                }

                CTreeNode* pNodeTemp = pMarkup->SearchBranchForScopeInStory(ptpStart->GetBranch(), pElement);
                // (fix this for IE5:RTM, #46824) srinib - if we are in the inclusion, we
                // wont find the node.
                if(pNodeTemp)
                {
                    //Assert(pNodeTemp);
                    ComputeIndentsFromParentNode(
                        &CI, pNodeTemp, dwFlags,
                        &xParentLeftIndent,
                        &xParentRightIndent);
                }
            }

            // For RTL lines, we will work from the right side. LTR lines work
            // from the left, of course.

            // WARNING!!! Be sure that any changes that are made to the LTR
            // case are reflected in the RTL case.           
            if(!fRTLLine)
            {              
                if(fBlockElement)
                {
                    // Block elems, except when processing selection:

                    // _fFirstFragInLine means it's the first logical line (chunk) in the physical line.
                    // If that is the case, the starting edge is just the indent from the parent
                    // (margin due to floats is ignored).
                    // If it's not the first logical line, then treat element as though it were inline.
                    xStart = pli->_fFirstFragInLine ? xParentLeftIndent : pli->_xLeftMargin+pli->_xLeft;
                    // _fForceNewLine means it's the last logical line (chunk) in a physical line.
                    // If that is the case, the line or view width includes the parent's right indent,
                    // so we need to subtract it out.  We also do this for dummy lines (actually,
                    // we should think about whether it's possible to skip processing of dummy lines altogether!)
                    // Otherwise, xLeftMargin+xLineWidth doesn't already include parent's right indent
                    // (because there's at least 1 other line to the right of this one), so we don't
                    // need to subtract it out, UNLESS the line is right aligned, in which case
                    // it WILL include it so we DO need to subtract it (wheee..).
                    xEnd = (pli->_fForceNewLine || pli->_fDummyLine)
                        ? max(pli->_xLineWidth, GetViewWidth())-xParentRightIndent
                        : pli->_xLeftMargin+pli->_xLineWidth-(pli->IsRightAligned()?xParentRightIndent:0);
                }
                else
                {
                    // Inline elems, and all selections begin at the text of
                    // the element, which is affected by margin.
                    xStart = pli->_xLeftMargin + pli->_xLeft;
                    // GetTextRight() tells us where the text ends, which is
                    // just what we want.
                    xEnd = pli->GetTextRight(iCurLine == LineCount()-1);
                    // Only include whitespace for selection
                    if(fIgnoreBlockness)
                    {
                        xEnd += pli->_xWhite;
                    }
                }
            }
            else
            {
                if(fBlockElement)
                {
                    xStart = pli->_fFirstFragInLine
                        ? xParentRightIndent : pli->_xRightMargin+pli->_xRight;
                    xEnd = (pli->_fForceNewLine || pli->_fDummyLine)
                        ? max(pli->_xLineWidth, GetViewWidth() )-xParentLeftIndent
                        : pli->_xRightMargin+pli->_xLineWidth-(pli->IsLeftAligned()?xParentLeftIndent:0);
                }
                else
                {
                    xStart = pli->_xRightMargin + pli->_xRight;
                    xEnd = pli->GetRTLTextLeft();
                    // Only include whitespace for selection
                    if(fIgnoreBlockness)
                    {
                        xEnd += pli->_xWhite;
                    }
                }
            }

            if(xEnd < xStart)
            {
                // Only clear lines can have a _xLineWidth < _xWidth + _xLeft + _xRight ...
                Assert(pli->IsClear());
                xEnd = xStart;
            }

            // Set the top and bottom
            if(fBlockElement)
            {
                yTop = yPos;
                yBottom = yPos + max(pli->_yHeight, pli->GetYBottom());

                if(fFirstLine)
                {
                    yTop += pli->_yBeforeSpace + 
                        min(0L, pli->GetYHeightTopOff()) + yParentPaddBordTop;

                    Assert(yBottom >= yTop);
                }
                else
                {
                    yTop += min(0L, pli->GetYTop());
                }

                if(pli->_fForceNewLine && cp+cchAdvance >= cpElementLast)
                {
                    yBottom -= yParentPaddBordBottom;
                }
            }
            else
            {
                yTop = yPos + pli->GetYTop();
                yBottom = yPos + pli->GetYBottom();

                // 66677: Let ScrollIntoView scroll to top of yBeforeSpace on first line.
                if(fScrollIntoView && yPos==0)
                {
                    yTop = 0;
                }
            }

            aryChunks.DeleteAll();
            cChunk = 0;

            // At this point we've found the bounds (xStart, xEnd, yTop, yBottom)
            // for a line that is part of the range.  Under certain circumstances,
            // this is insufficient, and we need to actually do measurement.  These
            // circumstances are:
            // 1.) If we're doing selection, we only need to measure partially
            // selected lines (i.e. at most we need to measure 2 lines -- the
            // first and the last).  For completely selected lines, we'll
            // just use the line bounds.  BUGBUG: this MAY introduce selection
            // turds, since LS uses tight-rects to draw the selection, and
            // our line bounds may not be as wide as LS tight-rects measurement
            // (recall that LS measurement catches whitespace that we sometimes
            // omit -- this may be fixable by adjusting our treatment of xWhite
            // and/or cchWhite).
            // Determination of partially selected lines is done as follows:
            // rp.GetIch() != 0 implies it starts mid-line,
            // rp.GetCchRemaining() != 0 implies it ends mid-line (the - _cchWhite
            // makes us ignore whitespace chars at the end of the line)
            // 2.) For all other situations, we're choosing to measure all
            // non-block elements.  This means that situations specifying
            // tight-rects (backgrounds, focus rects) will get them for
            // non-block elements (which is what we want), and hit-testing
            // will get the right rect if the element is inline and doesn't
            // span the entire line.  BUGBUG: there is a possible perf
            // improvement here; for hit-testing we really only need to
            // measure when the inline element doesn't span the entire line
            // (right now we measure when hittesting all inline elements);
            // this condition is the same as that for selecting partial lines,
            // however there may be subtler issues here that need investigation.
            if((dwFlags&RFE_SELECTION) == RFE_SELECTION)
            {
                fNeedToMeasureLine = !rp->IsBlankLine()
                    && (rp.GetIch()
                    || max(0L, (LONG)(rp.GetCchRemaining()-(fBlockElement?rp->_cchWhite:0)))>cchAdvance);
            }
            else
            {
                fNeedToMeasureLine = !fBlockElement;
            }

            if(fNeedToMeasureLine)
            {
                // (KTam) why do we need to set the measurer's cp to the 
                // beginning of the line? (cp - rp.GetIch() is beg. of line)
                Assert(!rp.GetIch() || cp==cpStart);
                ptpStart = pMarkup->TreePosAtCp(cp - rp.GetIch(), &ich);
                me.SetCp(cp-rp.GetIch(), ptpStart);
                me._li = *pli;
                // Get chunk info for the (possibly partial) line.
                // Chunk info is returned in aryChunks.
                cChunk = me.MeasureRangeOnLine(rp.GetIch(), cchAdvance, *pli, yPos, &aryChunks, dwFlags);
                if(cChunk == 0)
                {
                    xEnd = xStart;
                }
            }

            // cChunk == 0 if we didn't need to measure the line, or if we tried
            // to measure it and MeasureRangeOnLine() failed.
            if(cChunk == 0)
            {
                rcLine.left = xStart;
                rcLine.right = xEnd;
                rcLine.top = yTop;
                rcLine.bottom = yBottom; 

                // Adjust xStart and xEnd to the display coordinates.
                if((BOOL)_fRTL != fRTLLine)
                {
                    int xWidth = pli->_xLeftMargin + pli->_xRightMargin + pli->_xLineWidth;
                    rcLine.left = xWidth - rcLine.left;
                    rcLine.right = xWidth - rcLine.right;
                }

                // adjust for origin at top right
                if(_fRTL)
                {
                    rcLine.left = -rcLine.left;
                    rcLine.right = -rcLine.right;
                }

                // make sure we have a positive rect
                if(rcLine.left > rcLine.right)
                {
                    long temp = rcLine.right;
                    rcLine.right = rcLine.left;
                    rcLine.left = temp;
                }

                AppendRectToElemRegion(
                    paryRects, prcBound, &rcLine,
                    ptTrans,
                    cp, cpClipStart, cpClipFinish);
            }
            else
            {
                Assert(aryChunks.Size() > 0);

                i = 0;
                while(i < aryChunks.Size())
                {
                    RECT rcChunk = aryChunks[i++];
                    LONG xStartChunk = xStart + rcChunk.left;
                    LONG xEndChunk;

                    // if it is the first or the last chunk, use xStart & xEnd
                    if(fBlockElement && (i==1 || i==aryChunks.Size()))
                    {
                        if(i == 1)
                        {
                            xStartChunk = xStart;
                        }

                        if(i == aryChunks.Size())
                        {
                            xEndChunk = xEnd;
                        }
                        else
                        {
                            xEndChunk = xStart + rcChunk.right;
                        }
                    }
                    else
                    {
                        xEndChunk = xStart + rcChunk.right;
                    }

                    // Adjust xStart and xEnd to the display coordinates.
                    if((BOOL)_fRTL != fRTLLine)
                    {
                        int xWidth = pli->_xLeftMargin + pli->_xRightMargin + pli->_xLineWidth;
                        xStartChunk = xWidth - xStartChunk;
                        xEndChunk = xWidth - xEndChunk;
                    }

                    // adjust for origin at top right
                    if(_fRTL)
                    {
                        xStartChunk = -xStartChunk;
                        xEndChunk = -xEndChunk;
                    }

                    if(xStartChunk <= xEndChunk)
                    {
                        rcLine.left = xStartChunk;
                        rcLine.right = xEndChunk;
                    }
                    else
                    {
                        rcLine.left = xEndChunk;
                        rcLine.right = xStartChunk;
                    }

                    if(!fTightRects)
                    {
                        rcLine.top = yTop;
                        rcLine.bottom = yBottom; 
                    }
                    else
                    {
                        rcLine.top = rcChunk.top;
                        rcLine.bottom = rcChunk.bottom; 
                    }

                    AppendRectToElemRegion(
                        paryRects, prcBound, &rcLine,
                        ptTrans,
                        cp, cpClipStart, cpClipFinish);
                }
            }

AdvanceToNextLineInRange:
            if(fRestorePtTrans)
            {
                ptTrans = ptTempForPtTrans;
                fRestorePtTrans = FALSE;
            }

            cp += cchAdvance;

            if(cchAdvance)
            {
                rp.RpAdvanceCp(cchAdvance, FALSE);
            }
            else
            {
                rp.AdjustForward(); // frame lines or clear lines
            }

            if(pli->_fForceNewLine)
            {
                yPos += pli->_yHeight;
                fFirstLine = FALSE;
            }
        }

        // For calls for selection or hittesting (but not background/borders),
        // if the last line contains any aligned images, check to see if
        // there are any aligned lines following the current line that come
        // under the scope of the element
        if(fIncludeAlignedElements)
        {
            if(!rp.GetIch())
            {
                rp.AdjustBackward();
            }

            iCurLine = rp;

            if(rp->HasAligned())
            {
                LONG diLine = 1;
                CLine* pli;

                // we dont have to worry about clear lines because, all the aligned lines that
                // follow should be consecutive.
                while((iCurLine+diLine<LineCount()) &&
                    (pli=&rp[diLine]) && pli->IsFrame() && !pli->_fFrameBeforeText)
                {
                    long cpLayoutStart  = pli->_pNodeLayout->Element()->GetFirstCp();
                    long cpLayoutFinish = pli->_pNodeLayout->Element()->GetLastCp();

                    if(fBlockElement)
                    {
                        CTreeNode* pNodeTemp = pMarkup->SearchBranchForScopeInStory(ptpFinish->GetBranch(), pElement);
                        if(pNodeTemp)
                        {
                            // Assert(pNodeTemp);
                            ComputeIndentsFromParentNode(
                                &CI, pNodeTemp, dwFlags,
                                &xParentLeftIndent, 
                                &xParentRightIndent);
                        }
                    }

                    // if the current line is a frame line and if the site contained
                    // in it is contained by the current element include it other wise
                    if(cpStart<=cpLayoutStart && cpFinish>=cpLayoutFinish)
                    {
                        RcFromAlignedLine(
                            &rcLine, pli, yPos,
                            fBlockElement, fFirstLine, _fRTL,
                            xParentLeftIndent, xParentRightIndent);

                        AppendRectToElemRegion(
                            paryRects, prcBound, &rcLine,
                            ptTrans,
                            cpFinish, cpClipStart, cpClipFinish);
                    }
                    diLine++;
                }
            }
        }
    }
Cleanup:
    return;
}

//+----------------------------------------------------------------------------
//
// Member:      RegionFromRange
//
// Synopsis:    Return the set of rectangles that encompass a range of
//              characters
//
//-----------------------------------------------------------------------------
void CDisplay::RegionFromRange(CDataAry<RECT>* paryRects, long cp, long cch)
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    long            cpFirst     = pFlowLayout->GetContentFirstCp();
    CLinePtr        rp(this);
    CLine*          pli;
    CRect           rc;
    long            ili;
    long            yTop, yBottom;

    if(pFlowLayout->IsRangeBeforeDirty(cp-cpFirst, cch)
        && rp.RpSetCp(cp, FALSE, TRUE))
    {
        //  First, determine the starting line
        ili = rp;

        if(rp->HasAligned())
        {
            while(ili > 0)
            {
                pli = Elem(ili);

                if(!pli->IsFrame()
                    || !pli->_fFrameBeforeText
                    || pli->_pNodeLayout->Element()->GetFirstCp()<cp)
                {
                    break;
                }

                Assert(pli->_cch == 0);
                ili--;
            }
        }

        //  Start with an empty rectangle (whose width is that of the client rectangle)
        GetFlowLayout()->GetClientRect(&rc);

        rc.top = rc.bottom = YposFromLine(rp);

        // 1) There is no guarantee that cp passed in will be at the beginning of a line
        // 2) In the loop below we decrement cch by the count of chars in the line
        //
        // This would be correct if the cp passed in was the beginning of the line
        // but since it is not, we need to bump up the char count by the offset
        // of cp within the line. If at BOL then this would be 0. (bug 47687 fix 2)
        cch += rp.RpGetIch();

        // Extend the rectangle over the affected lines
        for(; cch>0&&ili<LineCount(); ili++)
        {
            pli = Elem(ili);

            yTop = rc.top + pli->GetYLineTop();
            yBottom = rc.bottom + pli->GetYLineBottom();

            rc.top = min(rc.top, yTop);
            rc.bottom = max(rc.bottom, yBottom);

            Assert(!pli->IsFrame() || !pli->_cch);
            cch -= pli->_cch;
        }

        // Save the invalid rectangle
        if(rc.top != rc.bottom)
        {
            paryRects->AppendIndirect(&rc);
        }
    }
}

//+----------------------------------------------------------------------------
//
// Member:      RenderedPointFromTp
//
// Synopsis:    Find the rendered position of a given cp. For cp that corresponds
//              normal text return its position in the line. For cp that points
//              to an aligned site find the aligned line rather than its poition
//              in the text flow. This function also takes care of relatively
//              positioned lines. Returns point relative to the display
//
//-----------------------------------------------------------------------------
LONG CDisplay::RenderedPointFromTp(
                              LONG      cp,         // point for the cp to be computed
                              CTreePos* ptp,        // tree pos for the cp passed in, can be NULL
                              BOOL      fAtEnd,     // Return end of previous line for ambiguous cp
                              POINT&    pt,         // Returns point at cp in client coords
                              CLinePtr* const prp,  // Returns line pointer at tp (may be null)
                              UINT      taMode,     // Text Align mode: top, baseline, bottom
                              CCalcInfo* pci)
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CLayout*        pLayout = NULL;
    CElement*       pElementLayout = NULL;
    CLinePtr        rp(this);
    LONG            ili;
    BOOL            fAlignedSite = FALSE;
    CCalcInfo       CI;
    CTreeNode*      pNode;

    if(!pci)
    {
        CI.Init(pFlowLayout);
        pci = &CI;
    }

    if(pFlowLayout->IsDisplayNone() || !WaitForRecalc(cp, -1, pci))
    {
        return -1;
    }

    // now position the line array point to the cp in the rtp.
    // Skip frames, let us worry about them latter.
    if(!rp.RpSetCp(cp, FALSE))
    {
        return -1;
    }

    if(!ptp)
    {
        LONG ich;
        ptp = pFlowLayout->GetContentMarkup()->TreePosAtCp(cp, &ich);
    }

    pNode   = ptp->GetBranch();
    pLayout = pNode->GetUpdatedNearestLayout();

    if(pLayout != pFlowLayout)
    {
        pElementLayout = pLayout ? pLayout->ElementOwner() : NULL;

        // is the current run owner the txtsite, if not get the runowner
        if(rp->_fHasNestedRunOwner)
        {
            CLayout* pRunOwner = pFlowLayout->GetContentMarkup()->GetRunOwner(pNode, pFlowLayout);

            if(pRunOwner != pFlowLayout)
            {
                pLayout = pRunOwner;
                pElementLayout = pLayout->ElementOwner();
            }
        }

        // if the site is left or right aligned and not a hr
        if(!pElementLayout->IsInlinedElement() && !pElementLayout->IsAbsolute())
        {
            long iDLine = -1;
            BOOL fFound = FALSE;

            fAlignedSite = TRUE;

            // run back and forth in the line array and figure out
            // which line belongs to the current site

            // first pass let's go back, look at lines before the current line
            while(rp+iDLine>=0 && !fFound)
            {
                CLine* pLine = &rp[iDLine];

                if(pLine->IsClear() || (pLine->IsFrame() && pLine->_fFrameBeforeText))
                {
                    if(pLine->IsFrame() && pElementLayout==pLine->_pNodeLayout->Element())
                    {
                        fFound = TRUE;

                        // now back up the linePtr to point to this line
                        rp.SetRun(rp+iDLine, 0);
                        break;
                    }
                }
                else
                {
                    break;
                }
                iDLine--;
            }

            // second pass let's go forward, look at lines after the current line.
            if(!fFound)
            {
                iDLine = 1;
                while(rp+iDLine < LineCount())
                {
                    CLine* pLine = &rp[iDLine];

                    // If it is a frame line
                    if(pLine->IsFrame() && !pLine->_fFrameBeforeText)
                    {
                        if(pElementLayout == pLine->_pNodeLayout->Element())
                        {
                            fFound = TRUE;

                            // now adjust the linePtr to point to this line
                            rp.SetRun(rp+iDLine, 0);
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                    iDLine++;
                }
            }

            // if we didn't find an aligned line, we are in deep trouble.
            Assert(fFound);
        }
    }

    if(!fAlignedSite)
    {
        // If it is not an aligned site then use PointFromTp
        ili = PointFromTp(cp, ptp, fAtEnd, FALSE, pt, prp, taMode, pci);
        if(ili > 0)
        {
            rp.SetRun(ili, 0);
        }
    }
    else
    {
        ili = rp;

        pt.y = YposFromLine(rp, pci);

        if(!_fRTL)
        {
            pt.x = rp->_xLeft + rp->_xLeftMargin;
        }
        else
        {
            pt.x = rp->_xRight + rp->_xRightMargin;
        }

        if(prp)
        {
            *prp = rp;
        }
    }

    return rp;
}

/*
*  CDisplay::UpdateViewForLists()
*
*  @mfunc
*      Invalidate the number regions for numbered list items.
*
*  @params
*      prcView:   The rect for this display
*      tpFirst:  Where the change happened -- the place where we check for
*                parentedness by an OL.
*      iliFirst: The first line where we start checking for a line beginning a
*                list item. It may not necessarily be the line containing
*                tpFirst. Could be any line after it.
*      yPos:     The yPos for the line iliFirst.
*      prcInval: The rect which returns the invalidation region
*
*  @rdesc
*      TRUE if updated anything else FALSE
*/
BOOL CDisplay::UpdateViewForLists(
                             RECT*  prcView,
                             LONG   cpFirst,
                             long   iliFirst,
                             long   yPos,
                             RECT*  prcInval)
{
    BOOL fHasListItem = FALSE;
    CLine* pLine = NULL; // Keep compiler happy.
    CMarkup* pMarkup = GetMarkup();
    CTreePos* ptp;
    LONG cchNotUsed;
    Assert(prcView);
    Assert(prcInval);

    ptp = pMarkup->TreePosAtCp(cpFirst, &cchNotUsed);
    // BUGBUG (sujalp): We might want to search for other interesting
    // list elements here.
    CElement* pElement = pMarkup->SearchBranchForTag(ptp->GetBranch(), ETAG_OL)->SafeElement();
    if(pElement && pElement->IsBlockElement())
    {
        while(iliFirst<LineCount() && yPos<prcView->bottom)
        {
            pLine = Elem(iliFirst);

            if(pLine->_fHasBulletOrNum)
            {
                fHasListItem = TRUE;
                break;
            }

            if(pLine->_fForceNewLine)
            {
                yPos += pLine->_yHeight;
            }
            iliFirst++;
        }

        if(fHasListItem)
        {
            // Invalidate the complete strip starting at the current
            // line, right down to the bottom of the view. And only
            // invalidate the strip containing the numbers not the
            // lines themselves.
            prcInval->top    = yPos;
            prcInval->left   = prcView->left;
            prcInval->right  = pLine->_xLeft;
            prcInval->bottom = prcView->bottom;
        }
    }

    return fHasListItem;
}

//==================================  Inversion (selection)  ============================
//+==============================================================================
//
// Method: ShowSelected
//
// Synopsis: The "Selected-ness" between the two CTreePos's has changed.
//           We tell the renderer about it - via Invalidate.
//
//           We also need to set the TextSelectionNess of any "sites"
//           on screen
//
//-------------------------------------------------------------------------------
#define CACHED_INVAL_RECTS  20

VOID CDisplay::ShowSelected(CTreePos* ptpStart, CTreePos* ptpEnd, BOOL fSelected) 
{
    CFlowLayout* pFlowLayout = GetFlowLayout();
    CElement* pElement = pFlowLayout->ElementContent();
    CStackDataAry<RECT, CACHED_INVAL_RECTS> aryInvalRects;
    CTreePos* pPrevPos = NULL;

    int cpClipStart = ptpStart->GetCp();
    int cpClipFinish = ptpEnd->GetCp();
    Assert(cpClipStart <= cpClipFinish);

    if(cpClipFinish > cpClipStart) // don't bother with cpClipFinish==cpClipStart
    {
        // BUGBUG we make the minimum selection size 3 chars to plaster over any off-by-one
        // problems with Region-From-Element, we also don't do this trick if a TreePos is
        // at an element - as RFE may get confused.        
        if(cpClipFinish+1<GetLastCp() && ptpEnd->IsPointer())
        {
            cpClipFinish++; 
        }
        if(cpClipStart-1>GetFirstCp() && ptpStart->IsPointer())
        {
            // Make sure you're not about to set the Cp to that of another valid TreePos.
            pPrevPos = ptpStart->PreviousTreePos();
            Assert(pPrevPos);
            if(pPrevPos && pPrevPos->GetCp()!=cpClipStart-1)
            {
                cpClipStart--;
            }
        }    
        WaitForRecalc(min(GetLastCp(), long(cpClipFinish)), -1);

        RegionFromElement(pElement, &aryInvalRects, NULL, NULL, RFE_SELECTION, cpClipStart, cpClipFinish, NULL); 

        pFlowLayout->Invalidate(&aryInvalRects[0], aryInvalRects.Size());           
    }
}

//+----------------------------------------------------------------------------
//
// Member:      CDisplay::GetViewHeightAndWidthForChild
//
// Synopsis:    returns the view width/height - padding
//
//-----------------------------------------------------------------------------
void CDisplay::GetViewWidthAndHeightForChild(
    CParentInfo* ppri,
    long*   pxWidthParent,
    long*   pyHeightParent,
    BOOL    fMinMax)
{
    long lPadding[PADDING_MAX];

    Assert(pxWidthParent && pyHeightParent);

    GetPadding(ppri, lPadding, fMinMax);

    *pxWidthParent  = GetViewWidth() - lPadding[PADDING_LEFT] - lPadding[PADDING_RIGHT];
    *pyHeightParent = GetViewHeight() - lPadding[PADDING_TOP] - lPadding[PADDING_BOTTOM];
}

//+----------------------------------------------------------------------------
//
// Member:      CDisplay::GetPadding
//
// Synopsis:    returns the top/left/right/bottom padding of the current
//              flowlayout
//
//-----------------------------------------------------------------------------
void CDisplay::GetPadding(CParentInfo* ppri, long lPadding[], BOOL fMinMax)
{
    CElement*   pElementFL   = GetFlowLayoutElement();
    CTreeNode*  pNode        = pElementFL->GetFirstBranch();
    CDocument*  pDoc         = pElementFL->Doc();
    ELEMENT_TAG etag         = pElementFL->Tag();
    LONG        lFontHeight  = pNode->GetCharFormat()->GetHeightInTwips(pDoc);
    long        lParentWidth = fMinMax ? ppri->_sizeParent.cx : GetViewWidth();
    long        lPaddingTop, lPaddingLeft, lPaddingRight, lPaddingBottom;
    const CFancyFormat* pFF  = pNode->GetFancyFormat();


    if(etag==ETAG_MARQUEE && !pElementFL->IsEditable())
    {
        AssertSz(FALSE, "must improve");
    }
    else
    {
        lPaddingTop = lPaddingBottom = 0;
    }

    if(etag != ETAG_TC)
    {
        lPaddingTop +=
            pFF->_cuvPaddingTop.YGetPixelValue(ppri, lParentWidth, lFontHeight);
        // (srinib) parent width is intentional as per css spec

        lPaddingBottom +=
            pFF->_cuvPaddingBottom.YGetPixelValue(ppri, lParentWidth, lFontHeight);
        // (srinib) parent width is intentional as per css spec

        lPaddingLeft =
            pFF->_cuvPaddingLeft.XGetPixelValue(ppri, lParentWidth, lFontHeight);

        lPaddingRight =
            pFF->_cuvPaddingRight.XGetPixelValue(ppri, lParentWidth, lFontHeight);

        if(etag == ETAG_BODY)
        {
            lPaddingLeft +=
                pFF->_cuvMarginLeft.XGetPixelValue(ppri, lParentWidth, lFontHeight);
            lPaddingRight +=
                pFF->_cuvMarginRight.XGetPixelValue(ppri, lParentWidth, lFontHeight);
            lPaddingTop +=
                pFF->_cuvMarginTop.YGetPixelValue(ppri, lParentWidth, lFontHeight);
            lPaddingBottom +=
                pFF->_cuvMarginBottom.YGetPixelValue(ppri, lParentWidth, lFontHeight);
        }
    }
    else
    {
        lPaddingLeft = 0;
        lPaddingRight = 0; 
    }

    // negative padding is not supported. What does it really mean?
    lPadding[PADDING_TOP] = lPaddingTop>SHRT_MAX
        ? SHRT_MAX : lPaddingTop<0?0:lPaddingTop;
    lPadding[PADDING_BOTTOM] = lPaddingBottom > SHRT_MAX
        ? SHRT_MAX : lPaddingBottom<0?0:lPaddingBottom;
    lPadding[PADDING_LEFT] = lPaddingLeft > SHRT_MAX
        ? SHRT_MAX : lPaddingLeft<0?0:lPaddingLeft;
    lPadding[PADDING_RIGHT] = lPaddingRight > SHRT_MAX
        ? SHRT_MAX : lPaddingRight<0?0:lPaddingRight;

    _fContainsVertPercentAttr |= pFF->_fPercentVertPadding;
    _fContainsHorzPercentAttr |= pFF->_fPercentHorzPadding;
}

//+----------------------------------------------------------------------------
//
// Member:      GetRectForChar
//
// Synopsis:    Returns the top, the bottom and the width for a character
//
//  Notes:      If pWidth is not NULL, the witdh of ch is returned in it.
//              Otherwise, ch is ignored.
//-----------------------------------------------------------------------------
void CDisplay::GetRectForChar(
                         CCalcInfo* pci,
                         LONG*      pTop,
                         LONG*      pBottom,
                         LONG*      pWidth,
                         WCHAR      ch,
                         LONG       yTop,
                         CLine*     pli,
                         const CCharFormat* pcf)
{
    CCcs* pccs = fc().GetCcs(pci->_hdc, pci, pcf);
    CBaseCcs* pBaseCcs;

    Assert(pci && pTop && pBottom && pli && pcf);
    Assert(pccs);

    pBaseCcs = pccs->GetBaseCcs();

    if(pWidth)
    {
        pccs->Include(ch, *pWidth);

        // Account for letter spacing (#18676)
        LONG xLetterSpacing;
        CUnitValue cuvLetterSpacing = pcf->_cuvLetterSpacing;

        switch(cuvLetterSpacing.GetUnitType())
        {
        case CUnitValue::UNIT_INTEGER:
            xLetterSpacing = (int)cuvLetterSpacing.GetUnitValue();
            break;

        case CUnitValue::UNIT_ENUM:
            xLetterSpacing = 0; // the only allowable enum value for l-s is normal=0
            break;

        default:
            xLetterSpacing = (int)cuvLetterSpacing.XGetPixelValue(pci, 0,
                GetFlowLayout()->GetFirstBranch()->GetFontHeightInTwips(&cuvLetterSpacing));
        }
        *pWidth += xLetterSpacing;
    }

    *pBottom = yTop + pli->_yHeight - pli->_yDescent + pli->_yTxtDescent;
    *pTop = *pBottom - pBaseCcs->_yHeight - pBaseCcs->_yOffset -
        (pli->_yTxtDescent - pBaseCcs->_yDescent);

    pccs->Release();
}

//+----------------------------------------------------------------------------
//
// Member:      GetTopBottomForCharEx
//
// Synopsis:    Returns the top and the bottom for a character
//
// Params:  [pDI]:     The DI
//          [pTop]:    The top is returned in this
//          [pBottom]: The bottom is returned in this
//          [yTop]:    The top of the line containing the character
//          [pli]:     The line itself
//          [xPos]:    The xPos at which we need the top/bottom
//          [pcf]:     Character format for the char
//
//-----------------------------------------------------------------------------
void CDisplay::GetTopBottomForCharEx(
                                CCalcInfo*  pci,
                                LONG*       pTop,
                                LONG*       pBottom,
                                LONG        yTop,
                                CLine*      pli,
                                LONG        xPos,
                                const CCharFormat* pcf)
{
    // If we are on a line in a list, and we are on the area occupied
    // (horizontally) by the bullet, then we want to find the height
    // of the bullet.
    if(pli->_fHasBulletOrNum
        && ((xPos>=pli->_xLeftMargin && xPos<pli->GetTextLeft())))
    {
        Assert(pci && pTop && pBottom && pli);

        *pBottom = yTop + pli->_yHeight - pli->_yDescent;
        *pTop = *pBottom - pli->_yBulletHeight;
    }
    else
    {
        GetRectForChar(pci, pTop, pBottom, NULL, 0, yTop, pli, pcf);
    }
}

//+----------------------------------------------------------------------------
//
// Member:      GetClipRectForLine
//
// Synopsis:    Returns the clip rect for a given line
//
// Params:  [prcClip]: Returns the rect for the line
//          [pTop]:    The top is returned in this
//          [pBottom]: The bottom is returned in this
//          [yTop]:    The top of the line containing the character
//          [pli]:     The line itself
//          [pcf]:     Character format for the char
//
//-----------------------------------------------------------------------------
void CDisplay::GetClipRectForLine(RECT* prcClip, LONG top, LONG xOrigin, CLine* pli) const
{
    Assert(prcClip && pli);

    if(!_fRTL)
    {
        prcClip->left   = xOrigin + pli->GetTextLeft();
        prcClip->right  = xOrigin + pli->GetTextRight();
    }
    else
    {
        prcClip->right  = xOrigin - pli->GetRTLTextRight();
        prcClip->left   = xOrigin - pli->GetRTLTextLeft();
    }
    if(pli->_fForceNewLine)
    {
        if(!pli->_fRTL)
        {
            prcClip->right += pli->_xWhite;
        }
        else
        {
            prcClip->left -= pli->_xWhite;
        }
    }
    prcClip->top    = top + pli->GetYTop();
    prcClip->bottom = top + pli->GetYBottom();
}

//+----------------------------------------------------------------------------
//
// Member:      ComputeVertPaddAndBordForParentNode
//
// Synopsis:    Computes the vertical padding and border for a given
//              element or range.
//
//-----------------------------------------------------------------------------
void CDisplay::ComputeVertPaddAndBordFromParentNode(
    CCalcInfo*  pci,
    CTreePos*   ptpStart,
    CTreePos*   ptpFinish,
    LONG*       pyPaddBordTop,
    LONG*       pyPaddBordBottom)
{
    CFlowLayout*    pFlowLayout = GetFlowLayout();
    CElement*       pElementFL  = pFlowLayout->ElementOwner();
    CTreePos*       ptpCurr;

    Assert(pci);
    Assert(ptpStart && ptpFinish);

    *pyPaddBordTop = 0;
    *pyPaddBordBottom = 0;

    ptpCurr = ptpStart->PreviousTreePos();

    // Walk up the tree as long as we're seeing "open tags w/o intervening
    // content".  Once we see some content, the vertical accumulation of
    // border and padding stops.
    while(ptpCurr && !ptpCurr->IsText())
    {
        if(ptpCurr->IsBeginNode())
        {
            CTreeNode* pNode = ptpCurr->Branch();
            CElement* pElement = pNode->Element();

            if(! ptpCurr->IsEdgeScope() || pElement==pElementFL)
            {
                break;
            }

            // add up padding for ptpCurr->Branch()
            if(pElement->IsBlockElement())
            {
                CBorderInfo borderinfo;

                if(!pElement->_fDefinitelyNoBorders)
                {
                    pElement->_fDefinitelyNoBorders =
                        !GetBorderInfoHelper(pNode, pci, &borderinfo, FALSE);

                    *pyPaddBordTop += borderinfo.aiWidths[BORDER_TOP];
                }

                // Note: _sizeParent.cx is to be used for percent base padding top
                // and bottom as per spec(seems weird though). Please talk to
                // cwilso about this.
                *pyPaddBordTop += pNode->GetFancyFormat()->_cuvPaddingTop.
                    YGetPixelValue(pci, pci->_sizeParent.cx, 1);
            }
        }
        else if(ptpCurr->IsEndNode())
        {
            CTreeNode* pNode = ptpCurr->Branch();
            CElement* pElement = pNode->Element();

            // The presence of no scape elements does not end a block of
            // "open tags w/o intervening content", because they don't
            // take up space in the rendering.  E.g.
            // <blockquote>
            //  <blockquote>
            //   <script>
            //        this script element doesn't interrupt the stack of
            //        open blockquotes w/o intervening content.
            //   </script>
            //    <blockquote>
            if(!pElement->IsNoScope())
            {
                break;
            }
        }

        ptpCurr = ptpCurr->PreviousTreePos();
    }

    // BUGBUG: there might be some overlapping weirdness to be handled here...
    ptpCurr = ptpFinish;
    if(ptpCurr->GetBranch()->Element() != pElementFL)
    {
        ptpCurr = ptpCurr->NextTreePos();

        while(ptpCurr && !ptpCurr->IsText()
            && !(ptpCurr->IsBeginNode() && ptpCurr->IsEdgeScope()))
        {
            if(ptpCurr->IsEndNode() && ptpCurr->IsEdgeScope())
            {
                CTreeNode* pNode = ptpCurr->Branch();
                CElement* pElement = pNode->Element();

                if(pElement == pElementFL)
                {
                    break;
                }

                // add up padding for ptpCurr->Branch()
                if(pElement->IsBlockElement())
                {
                    CBorderInfo borderinfo;

                    if(!pElement->_fDefinitelyNoBorders)
                    {
                        pElement->_fDefinitelyNoBorders =
                            !GetBorderInfoHelper(pNode, pci, &borderinfo, FALSE);

                        *pyPaddBordBottom += borderinfo.aiWidths[BORDER_BOTTOM];
                    }

                    // (srinib) Note: _sizeParent.cx is to be used for percent base padding top
                    // and bottom as per spec(seems weird though). Please talk to cwilso
                    // about this if you need to change this
                    *pyPaddBordBottom += pNode->GetFancyFormat()->_cuvPaddingBottom.
                        YGetPixelValue(pci, pci->_sizeParent.cx, 1);
                }
            }
            ptpCurr = ptpCurr->NextTreePos();
        }
    }
}

//+----------------------------------------------------------------------------
//
// Member:      GetWigglyFromRange
//
// Synopsis:    Gets rectangles to use for focus rects. This
//              element or range.
//
//-----------------------------------------------------------------------------
HRESULT CDisplay::GetWigglyFromRange(CDocInfo* pdci, long cp, long cch, CShape** ppShape)
{
    CStackDataAry<RECT, 8> aryWigglyRects;
    HRESULT         hr       = S_FALSE;
    CWigglyShape*   pShape   = NULL;
    CMarkup*        pMarkup  = GetMarkup();
    long            ich, cRects;
    CTreePos*       ptp      = pMarkup->TreePosAtCp(cp, &ich);
    CTreeNode*      pNode    = ptp->GetBranch();
    CElement*       pElement = pNode->Element();

    if(!cch)
    {
        goto Cleanup;
    }

    pShape = new CWigglyShape;
    if(!pShape)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    // get all of the rectangles that apply and load them into CWigglyShape's
    // array of CRectShapes.
    // RegionFromElement will give rects for any chunks within the number of
    // lines required.
    RegionFromElement(
        pElement, 
        &aryWigglyRects, 
        NULL, 
        NULL, 
        RFE_ELEMENT_RECT|RFE_IGNORE_RELATIVE|RFE_NESTED_REL_RECTS, 
        cp, // of lines that have different heights
        cp+cch, 
        NULL ); 

    for(cRects=0; cRects<aryWigglyRects.Size(); cRects++)
    {
        CRectShape* pWiggly = NULL;

        pWiggly = new CRectShape;
        if(!pWiggly)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        pWiggly->_rect = aryWigglyRects[cRects];

        // Expand the focus shape by shifting the left edges by 1 pixel.
        // This is a hack to get around the fact that the Windows
        // font rasterizer occasionally uses the leftmost pixel column of its
        // measured text area, and if we don't do this, the focus rect
        // overlaps that column.  IE bug #76378, taken for NT5 RTM (ARP).
        if(pWiggly->_rect.left > 0)
        {
            --(pWiggly->_rect.left);
        }

        pShape->_aryWiggly.Append(pWiggly);
    }

    *ppShape = pShape;
    hr = S_OK;

Cleanup:
    if(hr && pShape)
    {
        delete pShape;
    }

    RRETURN1(hr, S_FALSE);
}