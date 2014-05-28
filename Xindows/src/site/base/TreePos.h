
#ifndef __XINDOWS_SITE_BASE_TREEPOS_H__
#define __XINDOWS_SITE_BASE_TREEPOS_H__

class CElement;
class CTreeNod;
class CMarkup;
class CMarkupPointer;
class CTreeDataPos;

#define BOOLFLAG(f, dwFlag)     ((DWORD)(-(LONG)!!(f))&(dwFlag))

//+------------------------------------------------------------
//
//  Class:      CTreePos (Position In Tree)
//
//-------------------------------------------------------------
class CTreePos
{
    friend class CMarkup;
    friend class CTreePosGap;
    friend class CTreeNode;
    friend class CSpliceTreeEngine;
    friend class CMarkupUndoUnit;

public:
    DECLARE_MEMALLOC_NEW_DELETE();

    // TreePos come in various flavors:
    //  Uninit  this node is uninitialized
    //  Node    marks begin or end of a CTreeNode's scope
    //  Text    holds a bunch of text (formerly known as CElementRun)
    //  Pointer implements an IMarkupPointer
    // Be sure the bit field _eType (below) is big enough for all the flavors
    enum EType { Uninit=0x0, NodeBeg=0x1, NodeEnd=0x2, Text=0x4, Pointer=0x8 };

    // cast to CTreeDataPos
    CTreeDataPos* DataThis();
    const CTreeDataPos* DataThis() const;

    // accessors
    EType   Type() const                { return (EType)(GetFlags()&TPF_ETYPE_MASK); }
    void    SetType(EType etype)        { Assert(etype <= Pointer); SetFlags((GetFlags()&~TPF_ETYPE_MASK)|(etype)); }
    BOOL    IsUninit() const            { return !TestFlag(NodeBeg|NodeEnd|Text|Pointer); }
    BOOL    IsNode() const              { return TestFlag(NodeBeg|NodeEnd); }
    BOOL    IsText() const              { return TestFlag(Text); }
    BOOL    IsPointer() const           { return TestFlag(Pointer); }
    BOOL    IsDataPos() const           { return TestFlag(TPF_DATA_POS); }
    BOOL    IsData2Pos() const          { Assert(!IsNode()); return TestFlag(TPF_DATA2_POS); }
    BOOL    IsBeginNode() const         { return TestFlag(NodeBeg); }
    BOOL    IsEndNode() const           { return TestFlag(NodeEnd); }
    BOOL    IsEdgeScope() const         { Assert(IsNode()); return TestFlag(TPF_EDGE); }
    BOOL    IsBeginElementScope() const { return IsBeginNode()&&IsEdgeScope(); }
    BOOL    IsEndElementScope() const   { return IsEndNode()&&IsEdgeScope(); }
    BOOL    IsBeginElementScope(CElement* pElem);
    BOOL    IsEndElementScope(CElement* pElem);

    CMarkup* GetMarkup();
    BOOL    IsInMarkup(CMarkup* pMarkup) { return GetMarkup()==pMarkup; }

    // The following are logical comparison operations (two pos are equal
    // when separated by only pointers or empty text positions).
    int InternalCompare(CTreePos* ptpThat);

    CTreeNode* Branch() const;      // Only valid to call on NodePoses
    CTreeNode* GetBranch() const;   // Can be called on any pos, may be expensive

    CMarkupPointer* MarkupPointer() const;
    void SetMarkupPointer(CMarkupPointer*);

    // GetInterNode finds the node with direct influence
    // over the position directly after this CTreePos.
    CTreeNode* GetInterNode() const;

    long Cch() const;
    long Sid() const;

    BOOL HasTextID() const { return IsText()&&TestFlag(TPF_DATA2_POS); }
    long TextID() const;

    long GetCElements() const { return IsBeginElementScope()?1:0; }
    long SourceIndex();

    long GetCch() const { return IsNode()?1:IsText()?Cch():0; }

    long GetCp();
    
    int  Gravity() const;
    void SetGravity(BOOL fRight);
    int  Cling() const;
    void SetCling(BOOL fStick);

    // modifiers
    void SetScopeFlags(BOOL fEdge);
    void ChangeCch(long cchDelta);

    // navigation
    CTreePos* NextTreePos();
    CTreePos* PreviousTreePos();

    CTreePos* NextNonPtrTreePos();
    CTreePos* PreviousNonPtrTreePos();

    static BOOL IsLegalPosition(CTreePos* ptpLeft, CTreePos* ptpRight);
    BOOL        IsLegalPosition(BOOL fBefore)
    {
        return fBefore?IsLegalPosition(PreviousTreePos(), this):IsLegalPosition(this, NextTreePos());
    }

protected:
    void      InitSublist();
    CTreePos* Parent() const;
    CTreePos* LeftChild() const;
    CTreePos* RightChild() const;
    CTreePos* LeftmostDescendant() const;
    CTreePos* RightmostDescendant() const;
    void      GetChildren(CTreePos** ppLeft, CTreePos** ppRight) const;
    HRESULT   Remove();
    void      Splay();
    void      RotateUp(CTreePos* p, CTreePos* g);
    void      ReplaceChild(CTreePos* pOld, CTreePos* pNew);
    void      RemoveChild(CTreePos* pOld);
    void      ReplaceOrRemoveChild(CTreePos* pOld, CTreePos* pNew);

    // constructors (for use only by CMarkup and CTreeNode)
    CTreePos() {}

private:
    // distributed order-statistic fields
    DWORD       _cElemLeftAndFlags; // number of elements that begin in my left subtree
    DWORD       _cchLeft;           // number of characters in my left subtree
                                    // structure fields (to maintain the splay tree)
    CTreePos*   _pFirstChild;       // pointer to my leftmost child   第一个孩子（有可能是左，也有可能是右）
    CTreePos*   _pNext;             // pointer to right sibling or parent    右兄弟或者父亲

    enum
    {
        TPF_ETYPE_MASK      = 0x0F,
        TPF_LEFT_CHILD      = 0x10,
        TPF_LAST_CHILD      = 0x20,
        TPF_EDGE            = 0x40,
        TPF_DATA2_POS       = 0x40,
        TPF_DATA_POS        = 0x80,
        TPF_FLAGS_MASK      = 0xFF,
        TPF_FLAGS_SHIFT     = 8
    };

    DWORD   GetFlags() const                { return (_cElemLeftAndFlags); }
    void    SetFlags(DWORD dwFlags)         { _cElemLeftAndFlags = dwFlags; }
    BOOL    TestFlag(DWORD dwFlag) const    { return (!!(GetFlags()&dwFlag)); }
    void    SetFlag(DWORD dwFlag)           { SetFlags(GetFlags()|dwFlag); }
    void    ClearFlag(DWORD dwFlag)         { SetFlags(GetFlags()&~(dwFlag)); }

    long    GetElemLeft() const             { return ((long)(_cElemLeftAndFlags>>TPF_FLAGS_SHIFT)); }
    void    SetElemLeft(DWORD cElem)        { _cElemLeftAndFlags = (_cElemLeftAndFlags&TPF_FLAGS_MASK)|(DWORD)(cElem<<TPF_FLAGS_SHIFT); }
    void    AdjElemLeft(long cDelta)        { _cElemLeftAndFlags += cDelta << TPF_FLAGS_SHIFT; }
    BOOL    IsLeftChild() const             { return (TestFlag(TPF_LEFT_CHILD)); } // 确保当前位置左右信息
    BOOL    IsLastChild() const             { return (TestFlag(TPF_LAST_CHILD)); } // 确保是否有下一个兄弟 左/右 Next表示什么   在左边的时候如何判断Next表示什么呢？这时候就需要IsLastChild判断有兄弟没
    void    MarkLeft()                      { SetFlag(TPF_LEFT_CHILD); }
    void    MarkRight()                     { ClearFlag(TPF_LEFT_CHILD); }
    void    MarkLeft(BOOL fLeft)            { SetFlags((GetFlags()&~TPF_LEFT_CHILD)|BOOLFLAG(fLeft, TPF_LEFT_CHILD)); }
    void    MarkFirst()                     { ClearFlag(TPF_LAST_CHILD); }
    void    MarkLast()                      { SetFlag(TPF_LAST_CHILD); }
    void    MarkLast(BOOL fLast)            { SetFlags((GetFlags()&~TPF_LAST_CHILD)|BOOLFLAG(fLast, TPF_LAST_CHILD)); }

    void      SetFirstChild(CTreePos* ptp)  { _pFirstChild = ptp; }
    void      SetNext(CTreePos* ptp)        { _pNext = ptp; }
    CTreePos* FirstChild() const            { return (_pFirstChild); }
    CTreePos* Next() const                  { return (_pNext); }

    // support for CTreePosGap
    CTreeNode* SearchBranchForElement(CElement* pElement, BOOL fLeft);

    // count encapsulation
    enum ECountFlags { TP_LEFT=0x1, TP_DIRECT=0x2, TP_BOTH=0x3 };

    struct SCounts
    {
        DWORD _cch;
        DWORD _cElem;
        void Clear();
        void Increase(const CTreePos* ptp);    // TP_DIRECT is implied
        BOOL IsNonzero();
    };

    void ClearCounts();
    void IncreaseCounts(const CTreePos* ptp, unsigned fFlags);
    void IncreaseCounts(const SCounts* pCounts );
    void DecreaseCounts(const CTreePos* ptp, unsigned fFlags);
    BOOL HasNonzeroCounts(unsigned fFlags);

    BOOL LogicallyEqual(CTreePos* ptpRight);

public:
    NO_COPY(CTreePos);
};



struct DATAPOSTEXT
{
    unsigned long   _cch:26;    // [Text] number of characters I own directly
    unsigned long   _sid:6;     // [Text] the script id for this run
                                // This member will only be valid if the TPF_DATA2_POS flag is turned
                                // on.  Otherwise, assume that _lTextID is 0.
    long            _lTextID;   // [Text] Text ID for DOM text nodes
};

struct DATAPOSPOINTER
{
    DWORD_PTR       _dwPointerAndGravityAndCling; // [Pointer] my CMarkupPointer and Gravity
};

class CTreeDataPos : public CTreePos
{
    friend class CTreePos;
    friend class CMarkup;
    friend class CMarkupUndoUnit;

public:
    DECLARE_MEMALLOC_NEW_DELETE()

protected:
    union
    {
        DATAPOSTEXT t;
        DATAPOSPOINTER p;
    };

private:
    CTreeDataPos() {}
    NO_COPY(CTreeDataPos);
};


inline CTreeDataPos* CTreePos::DataThis()
{
    Assert(IsDataPos());
    return (CTreeDataPos*)this;
}

inline const CTreeDataPos* CTreePos::DataThis() const
{
    Assert(IsDataPos());
    return (const CTreeDataPos*)this;
}

inline CMarkupPointer* CTreePos::MarkupPointer() const
{
    Assert(IsPointer() && IsDataPos());
    return (CMarkupPointer*)(DataThis()->p._dwPointerAndGravityAndCling&~3);
}

inline void CTreePos::SetMarkupPointer(CMarkupPointer* pmp)
{
    Assert(IsPointer() && IsDataPos());
    Assert((DWORD_PTR(pmp)&0x1) == 0);
    DataThis()->p._dwPointerAndGravityAndCling = (DataThis()->p._dwPointerAndGravityAndCling&1)|(DWORD_PTR)(pmp);
}

inline long CTreePos::Cch() const
{
    Assert(IsText() && IsDataPos());
    return DataThis()->t._cch;
}

inline long CTreePos::Sid() const
{
    Assert(IsText() && IsDataPos());
    return DataThis()->t._sid;
}

inline long CTreePos::TextID() const
{
    Assert(IsText() && IsDataPos());
    return HasTextID()?DataThis()->t._lTextID:0;
}

inline CTreePos* CTreePos::Parent() const
{
    return IsLastChild()?Next():Next()->Next();
}

inline CTreePos* CTreePos::LeftChild() const
{
    return (FirstChild()&&FirstChild()->IsLeftChild())?FirstChild():NULL;
}

inline CTreePos* CTreePos::RightChild() const
{
    if(!FirstChild())
    {
        return NULL;
    }
    if(!FirstChild()->IsLeftChild())
    {
        return FirstChild();
    }
    if(!FirstChild()->IsLastChild())
    {
        return FirstChild()->Next();
    }
    return NULL;
}

inline void CTreePos::ReplaceOrRemoveChild(CTreePos* pOld, CTreePos* pNew)
{
    if(pNew)
    {
        ReplaceChild(pOld, pNew);
    }
    else
    {
        RemoveChild(pOld);
    }
}

inline void CTreePos::SCounts::Clear()
{
    _cch = 0;
    _cElem = 0;
}

inline void CTreePos::SCounts::Increase(const CTreePos* ptp)
{
    if(ptp->IsNode())
    {
        if(ptp->IsBeginElementScope())
        {
            _cElem++;
        }
        _cch++;;
    }
    else if(ptp->IsText())
    {
        _cch += ptp->Cch();
    }
}

inline BOOL CTreePos::SCounts::IsNonzero()
{
    return _cElem||_cch;
}

inline void CTreePos::ClearCounts()
{
    SetElemLeft(0);
    _cchLeft = 0;
}

inline void CTreePos::IncreaseCounts(const CTreePos* ptp, unsigned fFlags)
{
    if(fFlags & TP_LEFT)
    {
        AdjElemLeft(ptp->GetElemLeft());
        _cchLeft += ptp->_cchLeft;
    }
    if(fFlags & TP_DIRECT)
    {
        if(ptp->IsNode())
        {
            if(ptp->IsBeginElementScope())
            {
                AdjElemLeft(1);
            }
            _cchLeft += 1;
        }
        else if(ptp->IsText())
        {
            _cchLeft += ptp->Cch();
        }
    }
}

inline void CTreePos::IncreaseCounts(const SCounts* pCounts)
{
    AdjElemLeft(pCounts->_cElem);
    _cchLeft += pCounts->_cch;
}

inline void CTreePos::DecreaseCounts(const CTreePos* ptp, unsigned fFlags)
{
    if(fFlags & TP_LEFT)
    {
        AdjElemLeft(-ptp->GetElemLeft());
        _cchLeft -= ptp->_cchLeft;
    }
    if(fFlags & TP_DIRECT)
    {
        if(ptp->IsNode())
        {
            if(ptp->IsBeginElementScope())
            {
                AdjElemLeft(-1);
            }
            _cchLeft -= 1;
        }
        else if(ptp->IsText())
        {
            _cchLeft -= ptp->Cch();
        }
    }
}

#endif //__XINDOWS_SITE_BASE_TREEPOS_H__