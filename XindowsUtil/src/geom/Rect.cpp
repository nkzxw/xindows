
#include "stdafx.h"
#include "Rect.h"

//+---------------------------------------------------------------------------
//
//  Member:     CRect::Union
//              
//  Synopsis:   Extend rect to contain the given point.  If the rect is
//              initially empty, it will contain only the point afterwards.
//              
//  Arguments:  p       point to extend to
//              
//----------------------------------------------------------------------------
void CRect::Union(const POINT& p)
{
    if(IsRectEmpty())
    {
        SetRect(p.x,p.y,p.x+1,p.y+1);
    }
    else
    {
        if(p.x < left)      left = p.x;
        if(p.y < top)       top = p.y;
        if(p.x >= right)    right = p.x+1;
        if(p.y >= bottom)   bottom = p.y+1;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CRect::Intersects
//              
//  Synopsis:   Determine whether the given rect intersects this rect without
//              taking the time to compute the intersection.
//              
//  Arguments:  rc      other rect
//              
//  Returns:    TRUE if the rects are not empty and overlap    
//              
//----------------------------------------------------------------------------
BOOL CRect::Intersects(const RECT& rc) const
{
    return (left<rc.right && top<rc.bottom &&
        right>rc.left && bottom>rc.top &&
        !IsRectEmpty() && !((CRect&)rc).IsRectEmpty());
}

//+---------------------------------------------------------------------------
//
//  Function:   CalcScrollDelta
//
//  Synopsis:   Calculates the distance needed to scroll to make a given rect
//              visible inside this rect.
//
//  Arguments:  rc          Rectangle which needs to be visible inside this rect
//              psizeScroll Amount to scroll
//              vp, hp      Where to "pin" given RECT inside this RECT
//
//  Returns:    TRUE if scrolling required.
//
//----------------------------------------------------------------------------
BOOL CRect::CalcScrollDelta(
            const CRect&     rc,
            CSize*           psizeScroll,
            CRect::SCROLLPIN spVert,
            CRect::SCROLLPIN spHorz) const
{
    int         i;
    long        cxLeft;
    long        cxRight;
    SCROLLPIN   sp;

    Assert(psizeScroll);

    if(spVert==SP_MINIMAL && spHorz==SP_MINIMAL && Contains(rc))
    {
        psizeScroll->SetSize(0, 0);
        return FALSE;
    }

    sp = spHorz;
    for(i=0; i<2; i++)
    {
        // Calculate amount necessary to "pin" the left edge
        cxLeft = rc[i] - (*this)[i];

        // Examine right edge only if not "pin"ing to the left
        if(sp != SP_TOPLEFT)
        {
            Assert(sp==SP_BOTTOMRIGHT || sp==SP_MINIMAL);

            cxRight = (*this)[i+2] - rc[i+2];

            // "Pin" the inner RECT to the right side of the outer RECT
            if(sp == SP_BOTTOMRIGHT)
            {
                cxLeft = -cxRight;
            }
            // Otherwise, move the minimal amount necessary to make the
            // inner RECT visible within the outer RECT
            // (This code will try to make the entire inner RECT visible
            //  and gives preference to the left edge)
            else if(cxLeft > 0)
            {
                if(cxRight >= 0)
                {
                    cxLeft = 0;
                }
                else if(-cxRight < cxLeft)
                {
                    cxLeft = -cxRight;
                }
            }
        }
        (*psizeScroll)[i] = cxLeft;
        sp = spVert;
    }

    return !psizeScroll->IsZero();
}

//+---------------------------------------------------------------------------
//
//  Member:     CRect::CountContainedCorners
//              
//  Synopsis:   Count how many corners of the given rect are contained by
//              this rect.  This is tricky, because a rect doesn't technically
//              contain any of its corners except the top left.  This method
//              returns a count of 4 for rc.CountContainedCorners(rc).
//              
//  Arguments:  rc      rect to count contained corners for
//              
//  Returns:    -1 if rectangles do not intersect, or 0-4 if they do.  Zero
//              if rc completely contains this rect.
//              
//  Notes:      
//              
//----------------------------------------------------------------------------
int CRect::CountContainedCorners(const RECT& rc) const
{
    if(!Intersects(rc))
    {
        return -1;
    }

    int c = 0;
    if(rc.left >= left)
    {
        if(rc.top >= top) c++;
        if(rc.bottom <= bottom) c++;
    }
    if(rc.right <= right)
    {
        if(rc.top >= top) c++;
        if(rc.bottom <= bottom) c++;
    }
    return c;
}

//+----------------------------------------------------------------------------
//
//  Function:   CombineRectsAggressive
//
//  Synopsis:   Given an array of non-overlapping rectangles sorted top-to-bottom,
//              left-to-right combine them agressively, where there may
//              be some extra area create (but not too much).
//
//  Arguments:  pcRects     - Number of RECTs passed
//              arc         - Array of RECTs
//
//  Returns:    pcRects - Count of combined RECTs
//              arc     - Array of combined RECTs
//
//  NOTE: 1) RECTs must be sorted top-to-bottom, left-to-right and cannot be overlapping
//        2) arc is not reduced in size, the unused entries at the end are ignored
//
//-----------------------------------------------------------------------------
void CombineRectsAggressive(int* pcRects, RECT* arc)
{
    int cRects;
    int iDest, iSrc = 0;
    int cTemp = 0;
    int aryTemp[MAX_INVAL_RECTS];       // touching rects to the right and bottom of rect iDest
    int aryWhichComb[MAX_INVAL_RECTS];  // index matches arc, value is which 
                                        // combined rect the invalid rect belongs
                                        // to. Zero = not set yet.
    int cCombined, marker, index, index2;
    int areaCombined, aryArea[MAX_INVAL_RECTS];
    RECT arcCombined[MAX_INVAL_RECTS];  // accumulates combined rects from arc

    Assert(pcRects);
    Assert(arc);

    // If there are no RECTs to combine, return immediately
    if(*pcRects <= 1)
    {
        return;
    }

    // Combine rects aggressively. Touching rects are combined together,
    // which may make our final regions include areas that wasn't in 
    // the original list of rects. 
    memset(aryWhichComb, 0, sizeof(aryWhichComb)); 
    memset(aryArea, 0, sizeof(aryArea));
    cCombined = 0;
    for(iDest=0,cRects=*pcRects; iDest<cRects; iDest++)
    {
        // Combine abutting rects. Iterate through the array of rects 
        // (arranged from top-to-bottom, left-to-right) and enumerate 
        // those that touch to the right and bottom. Since this misses
        // those touching to the right and upwards we keep a list of 
        // which original rect belongs to which rect that is going to be 
        // passed back. As touching rects are accumulated, we check to 
        // see if any of these already belong to a rect that is going to 
        // be passed back. If it is, we add the ones currently being 
        // accumulated to that one.
        cTemp = 0;
        marker = cCombined + 1;
        aryTemp[cTemp++] = iDest;
        if(aryWhichComb[iDest] > 0)
        {
            marker = aryWhichComb[iDest];
        }

        for(iSrc=iDest+1; 
            arc[iDest].bottom>arc[iSrc].top&&iSrc<cRects;
            iSrc++)
        {
            Assert(arc[iDest].right <= arc[iSrc].left);
            if(arc[iDest].right == arc[iSrc].left)
            {
                // I don't think this ever happens. but just in case it 
                // does this will do the right thing
                Assert(0 && "Horizontal abutting invalid rects");
                if(aryWhichComb[iSrc] > 0)
                {
                    marker = aryWhichComb[iSrc];
                }
                aryTemp[cTemp++] = iSrc;
            }
        }

        // Check rects below for abuttment. 
        for(;
            arc[iDest].bottom==arc[iSrc].top&&
            arc[iDest].right>arc[iSrc].left&&iSrc<cRects;
            iSrc++)
        {
            if(arc[iDest].left < arc[iSrc].right)
            {
                if(aryWhichComb[iSrc] > 0)
                {
                    marker = aryWhichComb[iSrc];
                }
                aryTemp[cTemp++] = iSrc;
            }
        }

        if(cCombined+1 == marker)
        {
            // this group of invalid rects doesn't combine with any 
            // existing rect that is going to be returned. Start a one
            arcCombined[cCombined] = arc[iDest];
            cCombined++;
        }

        // Add all rects accumulated on this pass to the rect that will be passed back.
        index = marker - 1;
        while(--cTemp >= 0)
        {
            index2 = aryTemp[cTemp];
            if(aryWhichComb[index2] != marker)
            {
                aryWhichComb[index2] = marker;
                arcCombined[index].left    = min(arcCombined[index].left,   arc[index2].left);
                arcCombined[index].top     = min(arcCombined[index].top,    arc[index2].top);
                arcCombined[index].right   = max(arcCombined[index].right,  arc[index2].right);
                arcCombined[index].bottom  = max(arcCombined[index].bottom, arc[index2].bottom);
                aryArea[index] += (arc[index2].right-arc[index2].left) * (arc[index2].bottom-arc[index2].top);
            }
        }
    }

    // check to make sure each rect meets the fitness criteria
    // don't want to be creating excessively large non-invalid
    // regions to draw
    cRects = cCombined;
    for(index=0,marker=1; index<cCombined; index++,marker++)
    {
        areaCombined = (arcCombined[index].right - arcCombined[index].left) *
            (arcCombined[index].bottom - arcCombined[index].top);
        if(areaCombined>1000 && areaCombined>3*aryArea[index])
        {
            // scrap combined rect and fall back on just combining
            // areas that will not create any extra space.
            int index3=cRects, count=0;
            for(index2=0; index2<*pcRects; index2++)
            {
                if(marker == aryWhichComb[index2])
                {
                    arcCombined[index3++] = arc[index2];
                    count++;
                }
            }
            CombineRects(&count, &(arcCombined[cRects]));
            cRects += count-1;
            memmove(&arcCombined[index], &arcCombined[index+1], 
                (cRects-index)*sizeof(arcCombined[0]));
            memmove(&aryArea[index], &aryArea[index+1], 
                (cRects-index)*sizeof(aryArea[0]));
            cCombined--;
            index--;
        }
    }

    // Weed out rects that may now be totally enclosed by the extra 
    // region gain in combining rects.
    for(iSrc=cRects-1; iSrc>=0; iSrc--)
    {
        for(iDest=0; iDest<cRects; iDest++)
        {
            if( arcCombined[iSrc].left   >= arcCombined[iDest].left   &&
                arcCombined[iSrc].top    >= arcCombined[iDest].top    &&
                arcCombined[iSrc].right  <= arcCombined[iDest].right  &&
                arcCombined[iSrc].bottom <= arcCombined[iDest].bottom &&
                iDest != iSrc )
            {
                memmove(&(arcCombined[iSrc]), &(arcCombined[iSrc+1]), 
                    (cRects-1-iSrc)*sizeof(arcCombined[0]));
                cRects--;
                break;
            }
        }
    }
    // set output vars
    memmove(arc, arcCombined, cRects*sizeof(arc[0]));
    *pcRects = cRects;

    return;
}

//+----------------------------------------------------------------------------
//
//  Function:   CombineRects
//
//  Synopsis:   Given an array of non-overlapping rectangles sorted top-to-bottom,
//              left-to-right combine them such that they create no extra area.
//
//  Arguments:  pcRects     - Number of RECTs passed
//              arc         - Array of RECTs
//
//  Returns:    pcRects - Count of combined RECTs
//              arc     - Array of combined RECTs
//
//  NOTE: 1) RECTs must be sorted top-to-bottom, left-to-right and cannot be overlapping
//        2) arc is not reduced in size, the unused entries at the end are ignored
//
//-----------------------------------------------------------------------------
void CombineRects(int* pcRects, RECT* arc)
{
    int cRects;
    int iDest, iSrc=0;

    Assert(pcRects);
    Assert(arc);

    // If there are no RECTs to combine, return immediately
    if(*pcRects <= 1)
    {
        return;
    }

    // Combine RECTs of similar shape with adjoining boundaries
    for(iDest=0,cRects=*pcRects-1; iDest<cRects; iDest++)
    {
        // First, combine left-to-right those RECTs with the same top and bottom
        // (Since the array is sorted top-to-bottom, left-to-right, adjoining RECTs
        //  with the same top and bottom will be contiguous in the array. As a result,
        //  the loop only needs to continue looking at elements until one is found whose
        //  top or bottom coordinates do not match, or whose left edge is not adjoining.)
        for(iSrc=iDest+1;
            iSrc<=cRects&&arc[iDest].top==arc[iSrc].top&&
            arc[iDest].bottom==arc[iSrc].bottom&&
            arc[iDest].right>=arc[iSrc].left;
            iSrc++)
        {
            arc[iDest].right = arc[iSrc].right;
        }

        // If RECTs were combined, shift those remaining downward and adjust the total count
        if((iSrc-1) > iDest)
        {
            cRects -= iSrc - iDest - 1;
            memmove(&arc[iDest+1], &arc[iSrc], cRects*sizeof(arc[0]));
        }

        // Next, combine top-to-bottom those RECTs whose bottoms and tops meet
        // (Again, since the array is sorted top-to-bottom, left-to-right, RECTs which share
        //  the left and right coordinates and touch bottom-to-top may not be next to one
        //  another in the array. The loop must scan until it founds an element whose top
        //  or left edge exceeds that of the destination RECT. It will skip elements whose
        //  tops occur above the bottom of the destination or which occur left of the
        //  destination. It will combine elements, one at a time, which touch bottom-to-top
        //  and have matching left/right coordinates.)
        for(iSrc=iDest+1; iSrc <= cRects; )
        {
            if(arc[iDest].bottom < arc[iSrc].top)
            {
                break;
            }
            else if(arc[iDest].bottom == arc[iSrc].top)
            {
                if(arc[iDest].left < arc[iSrc].left)
                {
                    break;
                }
                else if(arc[iDest].left==arc[iSrc].left && arc[iDest].right==arc[iSrc].right)
                {
                    arc[iDest].bottom = arc[iSrc].bottom;
                    memmove(&arc[iSrc], &arc[iSrc+1], (cRects-iSrc)*sizeof(arc[0]));
                    cRects--;
                    continue;
                }
            }

            iSrc++;
        }
    }

    // Adjust the returned number RECTs
    *pcRects = cRects + 1;
    return;
}