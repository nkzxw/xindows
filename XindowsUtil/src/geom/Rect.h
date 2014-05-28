
#ifndef __XINDOWSUTIL_GEOM_RECT_H__
#define __XINDOWSUTIL_GEOM_RECT_H__

//+---------------------------------------------------------------------------
//
//  Class:      CRect
//
//  Synopsis:   Rectangle class (same as struct RECT, but with handy methods).
//
//----------------------------------------------------------------------------

// type of coordinates used in RECT structure (must match RECT definition)
typedef LONG RECTCOORD;

class XINDOWS_PUBLIC CRect : public RECT
{
public:
    CRect() {}
    ~CRect() {}

    // handy constructors
    enum CRECT_INITIALIZER { CRECT_EMPTY };
    CRect(CRECT_INITIALIZER)
    {
        left = top = right = bottom = 0;
    }
    CRect(const POINT& topLeft, const POINT& bottomRight)
    {
        left = topLeft.x;
        top = topLeft.y;
        right = bottomRight.x;
        bottom = bottomRight.y;
    }
    CRect(const SIZE& size)
    {
        left = top = 0;
    right = size.cx; bottom = size.cy;
    }
    CRect(const POINT& topLeft, const SIZE& size)
    {
        left = topLeft.x; top = topLeft.y;
        right = left + size.cx;
        bottom = top + size.cy;
    }
    CRect(RECTCOORD left_, RECTCOORD top_, RECTCOORD right_, RECTCOORD bottom_)
    {
        left = left_;
        top = top_;
        right = right_;
        bottom = bottom_;
    }
    CRect(const RECT& rc)
    {
        left = rc.left;
        top = rc.top;
        right = rc.right;
        bottom = rc.bottom;
    }

    CRect& operator=(const RECT& rc)
    {
        left = rc.left;
        top = rc.top;
        right = rc.right;
        bottom = rc.bottom;
        return *this;
    }

    BOOL IsRectEmpty() const
    {
        return (left>=right || top>=bottom);
    }
    BOOL IsEmpty() const { return IsRectEmpty(); }

    // access to individual edges
    enum EDGE
    {
        EDGE_LEFT   = 0,
        EDGE_TOP    = 1,
        EDGE_RIGHT  = 2,
        EDGE_BOTTOM = 3
    };
    RECTCOORD& operator[](int index) { return ((RECTCOORD*)this)[index]; }
    const RECTCOORD& operator[](int index) const { return ((const RECTCOORD*)this)[index]; }

    // empty rectangles may report intersections
    BOOL FastIntersects(const RECT& rc) const
    {
        return (left<rc.right && top<rc.bottom && right>rc.left && bottom>rc.top);
    }

    // empty rectangles will not produce intersections
    BOOL Intersects(const RECT& rc) const;

    BOOL Contains(const POINT& p) const
    {
        return (p.x>=left && p.x<right && p.y>=top && p.y<bottom);
    }

    BOOL Contains(const RECT& rc) const
    {
        return (rc.left>=left &&
            rc.top>=top &&
            rc.right<=right &&
            rc.bottom<=bottom &&
            rc.left<rc.right &&
            rc.top<rc.bottom);
    }

    // NOTE: empty rect may be contained in non-empty rect
    BOOL FastContains(const RECT& rc) const
    {
        return (rc.left>=left &&
            rc.top>=top &&
            rc.right<=right &&
            rc.bottom<=bottom);
    }

    // synonym for Contains
    BOOL PtInRect(const POINT& p) const { return Contains(p); }

    // return points 0=topleft, 1=topright,
    // 2=bottomleft, 3=bottomright
    void GetPoint(int i, POINT* p) const
    {
        p->x = (i&1) ? right : left;
        p->y = (i<2) ? top : bottom;
    }

    RECTCOORD Width() const { return right-left; }
    RECTCOORD Height() const { return bottom-top; }
    CSize Size() const { return CSize(right-left, bottom-top); }
    RECTCOORD Size(int axis) const { return (axis==0?right-left:bottom-top); }
    void GetSize(SIZE* size) const
    {
        size->cx = right-left;
        size->cy = bottom-top;
    }
    CPoint Center() const { return CPoint((left+right)/2, (top+bottom)/2); }
    void GetCenter(POINT* p) const
    {
        p->x = (left+right)/2;
        p->y = (top+bottom)/2;
    }
    ULONG Area() const
    {
        RECTCOORD w = Width();
        RECTCOORD h = Height();
        return (w>0&&h>0)?(ULONG)(w*h):(ULONG)0;
    }
    LONG FastArea() const { return Width()*Height(); }

    BOOL operator==(const RECT& rc) const
    {
        return (left==rc.left && top==rc.top &&
            right==rc.right && bottom==rc.bottom);
    }
    BOOL operator!=(const RECT& rc) const { return !(*this==rc); }
    BOOL EqualSize(const RECT& rc) const
    {
        return (Width()==rc.right-rc.left && Height()==rc.bottom-rc.top);
    }

    const CPoint& TopLeft() const { return *((const CPoint*)&left); }
    const CPoint& BottomRight() const { return *((const CPoint*)&right); }
    void GetTopLeft(POINT* p) const { p->x = left; p->y = top; }
    void GetBottomRight(POINT* p) const { p->x = right; p->y = bottom; }
    void GetTopRight(POINT* p) const { p->x = right; p->y = top; }
    void GetBottomLeft(POINT* p) const { p->x = left; p->y = bottom; }

    void SetTopLeft(const POINT& p) { left = p.x; top = p.y; }
    void SetBottomRight(const POINT& p) { right = p.x; bottom = p.y; }
    void SetTopRight(const POINT& p) { right = p.x; top = p.y; }
    void SetBottomLeft(const POINT& p) { left = p.x; bottom = p.y; }
    void SetRect(const POINT& topLeft, const POINT& bottomRight)
    {
        left = topLeft.x;
        top = topLeft.y;
        right = bottomRight.x;
        bottom = bottomRight.y;
    }
    void SetRect(const POINT& topLeft, const SIZE& size)
    {
        left = topLeft.x;
        top = topLeft.y;
        right = left + size.cx;
        bottom = top + size.cy;
    }
    void SetRect(const SIZE& size)
    {
        left = top = 0;
        right = size.cx;
        bottom = size.cy;
    }
    void SetRect(RECTCOORD left_, RECTCOORD top_, RECTCOORD right_, RECTCOORD bottom_)
    {
        left = left_;
        top = top_;
        right = right_;
        bottom = bottom_;
    }
    void SetRect(RECTCOORD all) { left = top = right = bottom = all; }
    void SetRectEmpty() { left = top = right = bottom = 0; }
    void SetMaxRect()
    {
        left = top = MINLONG;
        right = bottom = MAXLONG;
    }

    // move so upper right corner is at topRight
    void MoveTopRightTo(const POINT& topRight)
    {
        left += topRight.x - right;
        bottom += topRight.y - top;
        right = topRight.x;
        top = topRight.y;
    }
    // move so upper left corner is at topLeft
    void MoveTo(const POINT& topLeft)
    {
        right += topLeft.x - left;
        bottom += topLeft.y - top;
        left = topLeft.x;
        top = topLeft.y;
    }
    void MoveTo(RECTCOORD left_, RECTCOORD top_)
    {
        right += left_ - left;
        bottom += top_ - top;
        left = left_;
        top = top_;
    }
    void MoveToOrigin()
    {
        right -= left;
        bottom -= top;
        left = top = 0;
    }
    void OffsetRect(const SIZE& offset)
    {
        left += offset.cx;
        top += offset.cy;
        right += offset.cx;
        bottom += offset.cy;
    }
    void OffsetRect(POINTCOORD x, POINTCOORD y)
    {
        left += x;
        top += y;
        right += x;
        bottom += y;
    }
    void OffsetX(POINTCOORD x) { left += x; right += x; }
    void OffsetY(POINTCOORD y) { top += y; bottom += y; }
    void SafeOffsetRect(const SIZE& offset)
    {
        if(left != MINLONG)     left += offset.cx;
        if(top != MINLONG)      top += offset.cy;
        if(right != MAXLONG)    right += offset.cx;
        if(bottom != MAXLONG)   bottom += offset.cy;
    }
    void InflateRect(const SIZE& offset)
    {
        left -= offset.cx;
        top -= offset.cy;
        right += offset.cx;
        bottom += offset.cy;
    }
    void InflateRect(POINTCOORD x, POINTCOORD y)
    {
        left -= x;
        top -= y;
        right += x;
        bottom += y;
    }
    void MirrorX()
    {
        long temp = left;
        left = - right;
        right= - temp;
    }


    // expand rect in given direction
    void ExpandRect(const SIZE& offset)
    {
        if(offset.cx < 0)
        {
            left += offset.cx;
        }
        else
        {
            right += offset.cx;
        }
        if(offset.cy < 0)
        {
            top += offset.cy;
        }
        else
        {
            bottom += offset.cy;
        }
    }
    void ExpandRect(POINTCOORD x, POINTCOORD y)
    {
        if(x < 0)
        {
            left += x;
        }
        else
        {
            right += x;
        }
        if(y < 0)
        {
            top += y;
        }
        else
        {
            bottom += y;
        }
    }

    void SetCenter(const POINT& p) { OffsetRect((const CPoint&)p-Center()); }
    void SetWidth(RECTCOORD width) { right = left + width; }
    void SetHeight(RECTCOORD height) { bottom = top + height; }
    void SetSize(const SIZE& size)
    {
        right = left + size.cx;
        bottom = top + size.cy;
    }
    void SetMinSize(const SIZE& size)
    {
        if(Width() < size.cx) SetWidth(size.cx);
        if(Height() < size.cy) SetHeight(size.cy);
    }
    void SetMaxSize(const SIZE& size)
    {
        if(Width() > size.cx) SetWidth(size.cx);
        if(Height() > size.cy) SetHeight(size.cy);
    }

    void Union(const POINT& p);
    void Union(const RECT& rc)
    {
        ::UnionRect((LPRECT)this,(CONST RECT*)this, &rc);
    }

    BOOL IntersectRect(const RECT& rc)
    {
        return ::IntersectRect((LPRECT)this,(CONST RECT*)this, &rc);
    }

    void AddOffsets(const RECT& offsets)
    {
        left += offsets.left;
        top += offsets.top;
        right -= offsets.right;
        bottom -= offsets.bottom;
    }

    enum SCROLLPIN          
    {    
        SP_TOPLEFT = 1, // Pin RECT to the top or left of this RECT
        SP_BOTTOMRIGHT, // Pin RECT to the bottom or right of this RECT
        SP_MINIMAL      // Calculate minimal scroll to move RECT into this RECT
    };

    BOOL CalcScrollDelta(
        const CRect&    rc,
        CSize*          psizeScroll,
        SCROLLPIN       spVert,
        SCROLLPIN       spHorz) const;

    // count how many corners of the given rect are
    // contained by this one.  This is tricky, because
    // a rect doesn't technically contain any of its
    // corners except the top left.  This method returns
    // a count of four for rc.CountContainedCorners(rc).
    // NOTE: returns -1 if rectangles do not intersect!
    int CountContainedCorners(const RECT& rc) const;
};



inline BOOL SetRectl(LPRECTL prcl, int xLeft, int yTop, int xRight, int yBottom)
{
    prcl->left = xLeft;
    prcl->top = yTop;
    prcl->right = xRight;
    prcl->bottom = yBottom;
    return TRUE;
}

inline BOOL SetRectlEmpty(LPRECTL prcl)
{
    return SetRectEmpty((RECT*)prcl);
}

inline BOOL IntersectRectl(LPRECTL prcDst, CONST RECTL* prcSrc1, CONST RECTL* prcSrc2)
{
    return IntersectRect((LPRECT)prcDst, (RECT*)prcSrc1, (RECT*)prcSrc2);
}

inline BOOL PtlInRectl(RECTL* prc, POINTL ptl)
{
    return PtInRect((RECT*)prc, *((POINT*)&ptl));
}

inline BOOL InflateRectl(RECTL* prcl, long xAmt, long yAmt)
{
    return InflateRect((RECT*)prcl, (int)xAmt, (int)yAmt);
}

inline BOOL OffsetRectl(RECTL* prcl, long xAmt, long yAmt)
{
    return OffsetRect((RECT*)prcl, (int)xAmt, (int)yAmt);
}

inline BOOL IsRectlEmpty(RECTL* prcl)
{
    return IsRectEmpty((RECT*)prcl);
}

inline BOOL UnionRectl(RECTL* prclDst, const RECTL* prclSrc1, const RECTL* prclSrc2)
{
    return UnionRect((RECT*)prclDst, (RECT*)prclSrc1, (RECT*)prclSrc2);
}

inline BOOL EqualRectl(const RECTL* prcl1, const RECTL* prcl2)
{
    return EqualRect((RECT*)prcl1, (RECT*)prcl2);
}

inline int ExcludeClipRect(HDC hdc, RECT* prc)
{
    return ::ExcludeClipRect(hdc, prc->left, prc->top, prc->right, prc->bottom);
}

inline BOOL ContainsRect(const RECT* prcOuter, const RECT* prcInner)
{
    return (prcOuter->left<=prcInner->left &&
        prcOuter->right>=prcInner->right &&
        prcOuter->top<=prcInner->top &&
        prcOuter->bottom>=prcInner->bottom);
}

#define MAX_INVAL_RECTS     50 // max number of rects we will bother combine before just doing one big one.
XINDOWS_PUBLIC void CombineRectsAggressive(int* pcRects, RECT* arc);
XINDOWS_PUBLIC void CombineRects(int* pcRects, RECT* arc);

#endif //__XINDOWSUTIL_GEOM_RECT_H__