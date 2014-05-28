
#ifndef __XINDOWS_SITE_BASE_TREENODE_H__
#define __XINDOWS_SITE_BASE_TREENODE_H__

class CRootElement;
class CMarkup;
class CFlowLayout;

// {1CFD2488-0385-4a0a-9448-EEC8465F54ED}
static const CLSID CLSID_CTreeNode = 
{ 0x1cfd2488, 0x385, 0x4a0a, { 0x94, 0x48, 0xee, 0xc8, 0x46, 0x5f, 0x54, 0xed } };

// {3D0F394C-D15B-43b9-817E-21310FDDD906}
static const CLSID CLSID_CElement = 
{ 0x3d0f394c, 0xd15b, 0x43b9, { 0x81, 0x7e, 0x21, 0x31, 0xf, 0xdd, 0xd9, 0x6 } };

// {B210B98D-E577-42d4-A6C9-75D29EEC5325}
static const CLSID CLSID_CTextSite = 
{ 0xb210b98d, 0xe577, 0x42d4, { 0xa6, 0xc9, 0x75, 0xd2, 0x9e, 0xec, 0x53, 0x25 } };


class NOVTABLE CTreeNode : public CVoid
{
    friend class CTreePos;

    DECLARE_CLASS_TYPES(CTreeNode, CVoid)

public:
    DECLARE_MEMCLEAR_NEW_DELETE();

    CTreeNode(CTreeNode* pParent, CElement* pElement=NULL);

    // Use this to get an interface to the element/node
    HRESULT GetElementInterface(REFIID riid, void** ppUnk);

    // NOTE: these functions may look like IUnknown functions
    //       but don't make that mistake.  They are here to
    //       manage creation of the tearoff to handle all
    //       external references.
    // These functions should not be called!
    NV_DECLARE_TEAROFF_METHOD(GetInterface, getinterface, (REFIID riid, LPVOID* ppv));
    NV_DECLARE_TEAROFF_METHOD_(ULONG, PutRef, putref, ());
    NV_DECLARE_TEAROFF_METHOD_(ULONG, RemoveRef, removeref, ());

    // These functions are to be used to keep a node/element
    // combination alive while it may leave the tree.
    // BEWARE: this may cause the creation of a tearoff.
    ULONG   NodeAddRef();
    ULONG   NodeRelease();

    // These functions manage the _fInMarkup bit
    void    PrivateEnterTree();
    void    PrivateExitTree();
    void    PrivateMakeDead();
    void    PrivateMarkupRelease();

    void    SetElement(CElement* pElement);
    void    SetParent(CTreeNode* pNodeParent);

    // Element access and structure methods
    CElement*   Element()       { return _pElement; }
    CElement*   SafeElement()   { return this?_pElement:NULL; }

    CTreePos*   GetBeginPos()   { return &_tpBegin; }
    CTreePos*   GetEndPos()     { return &_tpEnd;   }

    // Context chain access
    BOOL        IsFirstBranch() { return _tpBegin.IsEdgeScope(); }
    BOOL        IsLastBranch()  { return _tpEnd.IsEdgeScope(); }
    CTreeNode*  NextBranch();
    CTreeNode*  PreviousBranch();

    CDocument*  Doc();

    BOOL            IsInMarkup()    { return _fInMarkup; }
    BOOL            IsDead()        { return ! _fInMarkup; }
    CRootElement*   IsRoot();
    CMarkup*        GetMarkup();
    CRootElement*   MarkupRoot();

    // Does the element that this node points to have currency?
    BOOL        HasCurrency();

    BOOL        IsContainer();
    CTreeNode*  GetContainerBranch();
    CElement*   GetContainer()      { return GetContainerBranch()->SafeElement(); }

    BOOL        SupportsHtml();

    CTreeNode*  Parent()            { return _pNodeParent; }

    CTreeNode*  Ancestor(ELEMENT_TAG etag);
    CTreeNode*  Ancestor(ELEMENT_TAG* arytag);

    CElement*   ZParent()           { return ZParentBranch()->SafeElement(); }
    CTreeNode*  ZParentBranch();

    CElement*   RenderParent()      { return RenderParentBranch()->SafeElement(); }
    CTreeNode*  RenderParentBranch();

    CElement*   ClipParent()        { return ClipParentBranch()->SafeElement(); }
    CTreeNode*  ClipParentBranch();

    CElement*   ScrollingParent()   { return ScrollingParentBranch()->SafeElement(); }
    CTreeNode*  ScrollingParentBranch();

    inline ELEMENT_TAG Tag()        { return (ELEMENT_TAG)_etag; }

    ELEMENT_TAG TagType()
    {
        switch(_etag)
        {
        case ETAG_GENERIC_LITERAL:
        case ETAG_GENERIC_BUILTIN:
            return ETAG_GENERIC;
        default:
            return (ELEMENT_TAG)_etag;
        }
    }

    CTreeNode* GetFirstCommonAncestor(CTreeNode* pNode, CElement* pEltStop);
    CTreeNode* GetFirstCommonBlockOrLayoutAncestor(CTreeNode* pNodeTwo, CElement* pEltStop);
    CTreeNode* GetFirstCommonAncestorNode(CTreeNode* pNodeTwo, CElement* pEltStop);

    CTreeNode* SearchBranchForPureBlockElement(CFlowLayout*);
    CTreeNode* SearchBranchToFlowLayoutForTag(ELEMENT_TAG etag);
    CTreeNode* SearchBranchToRootForTag(ELEMENT_TAG etag);
    CTreeNode* SearchBranchToRootForScope(CElement* pElementFindMe);
    BOOL       SearchBranchToRootForNode(CTreeNode* pNodeFindMe);

    CTreeNode* GetCurrentRelativeNode(CElement* pElementFL);

    // The layout attached to the current element may not be accurate, when a
    // property changes, current element can gain/lose layoutness. When an
    // element gains/loses layoutness, its layout is created/destroyed lazily.
    //
    // So, for the following functions "cur" means return the layout currently
    // associated with the layout which may not be accurate. "Updated" means
    // compute the state and return the accurate information.
    //
    // Note: Calling "Updated" function may cause the formats to be computed.
    //
    // If there is any confusion please talk to (srinib/lylec/brendand)
    inline  CLayout*    GetCurLayout();
    inline  BOOL        HasLayout();

    CLayout*            GetCurNearestLayout();
    CTreeNode*          GetCurNearestLayoutNode();
    CElement*           GetCurNearestLayoutElement();

    CLayout*            GetCurParentLayout();
    CTreeNode*          GetCurParentLayoutNode();
    CElement*           GetCurParentLayoutElement();

    // the following get functions may create the layout if it is not
    // created yet.
    inline  CLayout*    GetUpdatedLayout();     // checks for NeedsLayout()
    inline  CLayout*    GetUpdatedLayoutPtr();  // Call this if NeedsLayout() is already called
    inline  BOOL        NeedsLayout();

    CLayout*            GetUpdatedNearestLayout();
    CTreeNode*          GetUpdatedNearestLayoutNode();
    CElement*           GetUpdatedNearestLayoutElement();

    CLayout*            GetUpdatedParentLayout();
    CTreeNode*          GetUpdatedParentLayoutNode();
    CElement*           GetUpdatedParentLayoutElement();

    // BUGBUG - these functions should go, we should not need
    // to know if the element has flowlayout.
    CFlowLayout*        GetFlowLayout();
    CTreeNode*          GetFlowLayoutNode();
    CElement*           GetFlowLayoutElement();
    CFlowLayout*        HasFlowLayout();

    // Helper methods
    htmlBlockAlign      GetParagraphAlign(BOOL fOuter);
    htmlControlAlign    GetSiteAlign();

    BOOL                IsInlinedElement();

    BOOL                IsPositionStatic(void);
    BOOL                IsPositioned(void);
    BOOL                IsAbsolute(stylePosition st);
    BOOL                IsAbsolute(void);

    BOOL                IsAligned();

    // IsRelative() tells you if the specific element has had a CSS position
    // property set on it ( by examining _fRelative and _bPositionType on the
    // FF).  It will NOT tell you if something is relative because one of its
    // ancestors is relative; that information is stored in the CF, and can be
    // had via IsInheritingRelativeness()
    BOOL IsRelative(stylePosition st);
    BOOL IsRelative(void);
    BOOL IsInheritingRelativeness(void);

    BOOL IsScrollingParent(void);
    BOOL IsClipParent(void);
    BOOL IsZParent(void);
    BOOL IsDisplayNone(void);
    BOOL IsVisibilityHidden(void);

    // Depth is defined to be 1 plus the count of parents above this element
    int  Depth() const;

    // Format info functions
    HRESULT             CacheNewFormats(CFormatInfo* pCFI);

    void                EnsureFormats();
    BOOL                IsCharFormatValid()     { return _iCF>=0; }
    BOOL                IsParaFormatValid()     { return _iPF>=0; }
    BOOL                IsFancyFormatValid()    { return _iFF>=0; }
    const CCharFormat*  GetCharFormat()         { return (_iCF>=0 ? ::GetCharFormatEx(_iCF) : GetCharFormatHelper()); }
    const CParaFormat*  GetParaFormat()         { return (_iPF>=0 ? ::GetParaFormatEx(_iPF) : GetParaFormatHelper()); }
    const CFancyFormat* GetFancyFormat()        { return (_iFF>=0 ? ::GetFancyFormatEx(_iFF) : GetFancyFormatHelper()); }
    const CCharFormat*  GetCharFormatHelper();
    const CParaFormat*  GetParaFormatHelper();
    const CFancyFormat* GetFancyFormatHelper();
    long                GetCharFormatIndex()    { return (_iCF>=0 ? _iCF : GetCharFormatIndexHelper()); }
    long                GetParaFormatIndex()    { return (_iPF>=0 ? _iPF : GetParaFormatIndexHelper()); }
    long                GetFancyFormatIndex()   { return (_iFF>=0 ? _iFF : GetFancyFormatIndexHelper()); }
    long                GetCharFormatIndexHelper();
    long                GetParaFormatIndexHelper();
    long                GetFancyFormatIndexHelper();

    long                GetFontHeightInTwips(CUnitValue* pCuv);
    void                GetRelTopLeft(CElement* pElementFL, CParentInfo* ppi, long* pxOffset, long* pyOffset);

    // These GetCascaded methods are taken from style.hdl where they were
    // originally generated by the PDL parser.
    CColorValue        GetCascadedbackgroundColor();
    CColorValue        GetCascadedcolor();
    CUnitValue         GetCascadedletterSpacing();
    styleTextTransform GetCascadedtextTransform();
    CUnitValue         GetCascadedpaddingTop();
    CUnitValue         GetCascadedpaddingRight();
    CUnitValue         GetCascadedpaddingBottom();
    CUnitValue         GetCascadedpaddingLeft();
    CColorValue        GetCascadedborderTopColor();
    CColorValue        GetCascadedborderRightColor();
    CColorValue        GetCascadedborderBottomColor();
    CColorValue        GetCascadedborderLeftColor();
    styleBorderStyle   GetCascadedborderTopStyle();
    styleBorderStyle   GetCascadedborderRightStyle();
    styleBorderStyle   GetCascadedborderBottomStyle();
    styleBorderStyle   GetCascadedborderLeftStyle();
    CUnitValue         GetCascadedborderTopWidth();
    CUnitValue         GetCascadedborderRightWidth();
    CUnitValue         GetCascadedborderBottomWidth();
    CUnitValue         GetCascadedborderLeftWidth();
    CUnitValue         GetCascadedwidth();
    CUnitValue         GetCascadedheight();
    CUnitValue         GetCascadedtop();
    CUnitValue         GetCascadedbottom();
    CUnitValue         GetCascadedleft();
    CUnitValue         GetCascadedright();
    styleOverflow      GetCascadedoverflowX();
    styleOverflow      GetCascadedoverflowY();
    styleOverflow      GetCascadedoverflow();
    styleStyleFloat    GetCascadedfloat();
    stylePosition      GetCascadedposition();
    long               GetCascadedzIndex();
    CUnitValue         GetCascadedclipTop();
    CUnitValue         GetCascadedclipLeft();
    CUnitValue         GetCascadedclipRight();
    CUnitValue         GetCascadedclipBottom();
    BOOL               GetCascadedtableLayout();    // fixed - 1, auto - 0
    BOOL               GetCascadedborderCollapse(); // collapse - 1, separate - 0
    BOOL               GetCascadedborderOverride();
    WORD               GetCascadedfontWeight();
    WORD               GetCascadedfontHeight();
    CUnitValue         GetCascadedbackgroundPositionX();
    CUnitValue         GetCascadedbackgroundPositionY();
    BOOL               GetCascadedbackgroundRepeatX();
    BOOL               GetCascadedbackgroundRepeatY();
    htmlBlockAlign     GetCascadedblockAlign();
    styleVisibility    GetCascadedvisibility();
    styleDisplay       GetCascadeddisplay();
    BOOL               GetCascadedunderline();
    styleAccelerator   GetCascadedaccelerator();
    BOOL               GetCascadedoverline();
    BOOL               GetCascadedstrikeOut();
    CUnitValue         GetCascadedlineHeight();
    CUnitValue         GetCascadedtextIndent();
    BOOL               GetCascadedsubscript();
    BOOL               GetCascadedsuperscript();
    BOOL               GetCascadedbackgroundAttachmentFixed();
    styleListStyleType GetCascadedlistStyleType();
    styleListStylePosition GetCascadedlistStylePosition();
    long               GetCascadedlistImageCookie();
    const TCHAR*       GetCascadedfontFaceName();
    const TCHAR*       GetCascadedfontFamilyName();
    BOOL               GetCascadedfontItalic();
    long               GetCascadedbackgroundImageCookie();
    BOOL               GetCascadedclearLeft();
    BOOL               GetCascadedclearRight();
    styleCursor        GetCascadedcursor();
    styleTableLayout   GetCascadedtableLayoutEnum();
    styleBorderCollapse GetCascadedborderCollapseEnum();
    styleDir           GetCascadedBlockDirection();
    styleDir           GetCascadeddirection();
    styleBidi          GetCascadedunicodeBidi();
    styleLayoutGridMode GetCascadedlayoutGridMode();
    styleLayoutGridType GetCascadedlayoutGridType();
    CUnitValue         GetCascadedlayoutGridLine();
    CUnitValue         GetCascadedlayoutGridChar();
    LONG               GetCascadedtextAutospace();
    styleWordBreak     GetCascadedwordBreak();
    styleLineBreak     GetCascadedlineBreak();
    styleTextJustify   GetCascadedtextJustify();
    styleTextJustifyTrim GetCascadedtextJustifyTrim();
    CUnitValue         GetCascadedmarginTop();
    CUnitValue         GetCascadedmarginRight();
    CUnitValue         GetCascadedmarginBottom();
    CUnitValue         GetCascadedmarginLeft();
    CUnitValue         GetCascadedtextKashida();

    // Ref helpers
    // Right now these just drop right through to the element
    static void ReplacePtr      (CTreeNode** ppNodelhs, CTreeNode* pNoderhs);
    static void SetPtr          (CTreeNode** ppNodelhs, CTreeNode* pNoderhs);
    static void ClearPtr        (CTreeNode** ppNodelhs);
    static void StealPtrSet     (CTreeNode** ppNodelhs, CTreeNode* pNoderhs);
    static void StealPtrReplace (CTreeNode** ppNodelhs, CTreeNode* pNoderhs);
    static void ReleasePtr      (CTreeNode*  pNode);

    // Other helpers
    void VoidCachedInfo();
    void VoidCachedNodeInfo();
    void VoidFancyFormat();

    // Helpers for contained CTreePos's
    CTreePos* InitBeginPos(BOOL fEdge)
    {
        _tpBegin.SetFlags(
            (_tpBegin.GetFlags()&~(CTreePos::TPF_ETYPE_MASK|CTreePos::TPF_DATA_POS|CTreePos::TPF_EDGE))
            | CTreePos::NodeBeg
            | BOOLFLAG(fEdge, CTreePos::TPF_EDGE));
        return &_tpBegin;
    }

    CTreePos* InitEndPos(BOOL fEdge)
    {
        _tpEnd.SetFlags(
            (_tpEnd.GetFlags()&~(CTreePos::TPF_ETYPE_MASK|CTreePos::TPF_DATA_POS|CTreePos::TPF_EDGE))
            | CTreePos::NodeEnd
            | BOOLFLAG(fEdge, CTreePos::TPF_EDGE));
        return &_tpEnd;
    }

    //+-----------------------------------------------------------------------
    //
    //  CTreeNode::CLock
    //
    //------------------------------------------------------------------------
    class CLock
    {
    public:
        DECLARE_MEMALLOC_NEW_DELETE()
        CLock(CTreeNode* pNode);
        ~CLock();

    private:
        CTreeNode* _pNode;
    };

    // Lookaside pointers
    enum
    {
        LOOKASIDE_PRIMARYTEAROFF    = 0,
        LOOKASIDE_CURRENTSTYLE      = 1,
        LOOKASIDE_NODE_NUMBER       = 2
        // *** There are only 2 bits reserved in the node.
        // *** if you add more lookasides you have to make sure 
        // *** that you make room for those bits.
    };

    BOOL            HasLookasidePtr(int iPtr)                   { return (_fHasLookasidePtr&(1<<iPtr)); }
    void*           GetLookasidePtr(int iPtr);
    HRESULT         SetLookasidePtr(int iPtr, void* pv);
    void*           DelLookasidePtr(int iPtr);

    // Primary Tearoff pointer management
    BOOL            HasPrimaryTearoff()                         { return (HasLookasidePtr(LOOKASIDE_PRIMARYTEAROFF)); }
    IUnknown*       GetPrimaryTearoff()                         { return ((IUnknown*)GetLookasidePtr(LOOKASIDE_PRIMARYTEAROFF)); }
    HRESULT         SetPrimaryTearoff(IUnknown* pTearoff)       { return (SetLookasidePtr(LOOKASIDE_PRIMARYTEAROFF, pTearoff)); }
    IUnknown*       DelPrimaryTearoff()                         { return ((IUnknown*)DelLookasidePtr(LOOKASIDE_PRIMARYTEAROFF)); }

    // Class Data
    CElement*   _pElement;                          // The element for this node
    CTreeNode*  _pNodeParent;                       // The parent in the CTreeNode tree

    // DWORD 1
    BYTE        _etag;                              // 0-7:     element tag
    BYTE        _fFirstCommonAncestorNode   : 1;    // 8:       for finding common ancestor
    BYTE        _fInMarkup                  : 1;    // 9:       this node is in a markup and shouldn't die
    BYTE        _fInMarkupDestruction       : 1;    // 10:      Used by CMarkup::DestroySplayTree
    BYTE        _fHasLookasidePtr           : 2;    // 11-12    Lookaside flags
    BYTE        _fBlockNess                 : 1;    // 13:      Cached from format -- valid if _iFF != -1
    BYTE        _fHasLayout                 : 1;    // 14:      Cached from format -- valid if _iFF != -1
    BYTE        _fUnused                    : 1;    // 15:      Unused

    SHORT       _iPF;                               // 16-31:   Paragraph Format

    // DWORD 2
    SHORT       _iCF;                               // 0-15:    Char Format
    SHORT       _iFF;                               // 16-31:   Fancy Format

protected:
    // Use GetBeginPos() or GetEndPos() to get at these members
    CTreePos    _tpBegin;                           // The begin CTreePos for this node
    CTreePos    _tpEnd;                             // The end CTreePos for this node

public:
    // STATIC MEMBERS
    DECLARE_TEAROFF_TABLE_NAMED(s_apfnNodeVTable)

private:
    NO_COPY(CTreeNode);
};

inline BOOL CTreeNode::IsPositionStatic(void)
{
    return !GetFancyFormat()->_fPositioned;
}

inline BOOL CTreeNode::IsPositioned(void)
{
    return GetFancyFormat()->_fPositioned;
}

inline BOOL CTreeNode::IsAbsolute(stylePosition st)
{
    return (Tag()==ETAG_ROOT || Tag()==ETAG_BODY || st==stylePositionabsolute);
}

inline BOOL CTreeNode::IsAbsolute(void)
{
    return IsAbsolute(GetCascadedposition());
}

inline BOOL CTreeNode::IsAligned(void)
{
    return GetFancyFormat()->IsAligned();
}

inline BOOL CTreeNode::IsRelative(stylePosition st)
{
    return (st==stylePositionrelative);
}

inline BOOL CTreeNode::IsRelative(void)
{
    return IsRelative(GetCascadedposition());
}

inline BOOL CTreeNode::IsInheritingRelativeness(void)
{
    return (GetCharFormat()->_fRelative);
}

// Returns TRUE if we're the body or a scrolling DIV or SPAN
inline BOOL CTreeNode::IsScrollingParent(void)
{
    return GetFancyFormat()->_fScrollingParent;
}

inline BOOL CTreeNode::IsClipParent(void)
{
    const CFancyFormat* pFF = GetFancyFormat();

    return IsAbsolute(stylePosition(pFF->_bPositionType))||pFF->_fScrollingParent;
}

inline BOOL CTreeNode::IsZParent(void)
{
    return GetFancyFormat()->_fZParent;
}

inline BOOL CTreeNode::IsDisplayNone()
{
    return (BOOL)GetCharFormat()->IsDisplayNone();
}

inline BOOL CTreeNode::IsVisibilityHidden()
{
    return (BOOL)GetCharFormat()->IsVisibilityHidden();
}

inline CUnitValue CTreeNode::GetCascadedborderTopWidth()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvBorderWidths[BORDER_TOP];
}

inline CUnitValue CTreeNode::GetCascadedborderRightWidth()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvBorderWidths[BORDER_RIGHT];
}

inline CUnitValue CTreeNode::GetCascadedborderBottomWidth()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvBorderWidths[BORDER_BOTTOM];
}

inline CUnitValue CTreeNode::GetCascadedborderLeftWidth()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvBorderWidths[BORDER_LEFT];
}

inline CUnitValue CTreeNode::GetCascadedwidth()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvWidth;
}

inline CUnitValue CTreeNode::GetCascadedheight()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvHeight;
}

inline CUnitValue CTreeNode::GetCascadedtop()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvTop;
}

inline CUnitValue CTreeNode::GetCascadedbottom()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvBottom;
}

inline CUnitValue CTreeNode::GetCascadedleft()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvLeft;
}

inline CUnitValue CTreeNode::GetCascadedright()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvRight;
}

inline styleOverflow CTreeNode::GetCascadedoverflowX()
{
    return (styleOverflow)CTreeNode::GetFancyFormat()->_bOverflowX;
}

inline styleOverflow CTreeNode::GetCascadedoverflowY()
{
    return (styleOverflow)CTreeNode::GetFancyFormat()->_bOverflowY;
}

inline styleOverflow CTreeNode::GetCascadedoverflow()
{
    const CFancyFormat* pFF = CTreeNode::GetFancyFormat();
    return (styleOverflow)(pFF->_bOverflowX>pFF->_bOverflowY ? pFF->_bOverflowX : pFF->_bOverflowY);
}

inline styleStyleFloat CTreeNode::GetCascadedfloat()
{
    return (styleStyleFloat)CTreeNode::GetFancyFormat()->_bStyleFloat;
}

inline stylePosition CTreeNode::GetCascadedposition()
{
    return (stylePosition) CTreeNode::GetFancyFormat()->_bPositionType;
}

inline long CTreeNode::GetCascadedzIndex()
{
    return (long) CTreeNode::GetFancyFormat()->_lZIndex;
}

inline CUnitValue CTreeNode::GetCascadedclipTop()
{
    return CTreeNode::GetFancyFormat()->_cuvClipTop;
}

inline CUnitValue CTreeNode::GetCascadedclipLeft()
{
    return CTreeNode::GetFancyFormat()->_cuvClipLeft;
}

inline CUnitValue CTreeNode::GetCascadedclipRight()
{
    return CTreeNode::GetFancyFormat()->_cuvClipRight;
}

inline CUnitValue CTreeNode::GetCascadedclipBottom()
{
    return CTreeNode::GetFancyFormat()->_cuvClipBottom;
}

inline BOOL CTreeNode::GetCascadedtableLayout()
{
    return (BOOL)CTreeNode::GetFancyFormat()->_bTableLayout;
}

inline BOOL CTreeNode::GetCascadedborderCollapse()
{
    return (BOOL) CTreeNode::GetFancyFormat()->_bBorderCollapse;
}

inline BOOL CTreeNode::GetCascadedborderOverride()
{
    return (BOOL) CTreeNode::GetFancyFormat()->_fOverrideTablewideBorderDefault;
}

inline WORD CTreeNode::GetCascadedfontWeight()
{
    return GetCharFormat()->_wWeight;
}

inline WORD CTreeNode::GetCascadedfontHeight()
{
    return GetCharFormat()->_yHeight;
}

inline CUnitValue CTreeNode::GetCascadedbackgroundPositionX()
{
    return CTreeNode::GetFancyFormat()->_cuvBgPosX;
}

inline CUnitValue CTreeNode::GetCascadedbackgroundPositionY()
{
    return CTreeNode::GetFancyFormat()->_cuvBgPosY;
}

inline BOOL CTreeNode::GetCascadedbackgroundRepeatX()
{
    return (BOOL)CTreeNode::GetFancyFormat()->_fBgRepeatX;
}

inline BOOL CTreeNode::GetCascadedbackgroundRepeatY()
{
    return (BOOL)CTreeNode::GetFancyFormat()->_fBgRepeatY;
}

inline htmlBlockAlign CTreeNode::GetCascadedblockAlign()
{
    return (htmlBlockAlign)CTreeNode::GetParaFormat()->_bBlockAlignInner;
}

inline styleVisibility CTreeNode::GetCascadedvisibility()
{
    return (styleVisibility)GetFancyFormat()->_bVisibility;
}

inline styleDisplay CTreeNode::GetCascadeddisplay()
{
    return (styleDisplay)GetFancyFormat()->_bDisplay;
}

inline BOOL CTreeNode::GetCascadedunderline()
{
    return (BOOL)GetCharFormat()->_fUnderline;
}

inline styleAccelerator CTreeNode::GetCascadedaccelerator()
{
    return (styleAccelerator)GetCharFormat()->_fAccelerator;
}

inline BOOL CTreeNode::GetCascadedoverline()
{
    return (BOOL)GetCharFormat()->_fOverline;
}

inline BOOL CTreeNode::GetCascadedstrikeOut()
{
    return (BOOL)GetCharFormat()->_fStrikeOut;
}

inline CUnitValue CTreeNode::GetCascadedlineHeight()
{
    return GetCharFormat()->_cuvLineHeight;
}

inline CUnitValue CTreeNode::GetCascadedtextIndent()
{
    return CTreeNode::GetParaFormat()->_cuvTextIndent;
}

inline BOOL CTreeNode::GetCascadedsubscript()
{
    return (BOOL)GetCharFormat()->_fSubscript;
}

inline BOOL CTreeNode::GetCascadedsuperscript()
{
    return (BOOL)GetCharFormat()->_fSuperscript;
}

inline BOOL CTreeNode::GetCascadedbackgroundAttachmentFixed()
{
    return (BOOL)CTreeNode::GetFancyFormat()->_fBgFixed;
}

inline styleListStyleType CTreeNode::GetCascadedlistStyleType()
{
    styleListStyleType slt;

    slt = (styleListStyleType)CTreeNode::GetFancyFormat()->_ListType;
    if(slt == styleListStyleTypeNotSet)
    {
        slt = GetParaFormat()->_cListing.GetStyle();
    }
    return slt;
}

inline styleListStylePosition CTreeNode::GetCascadedlistStylePosition()
{
    return (styleListStylePosition)GetParaFormat()->_bListPosition;
}

inline long CTreeNode::GetCascadedlistImageCookie()
{
    return GetParaFormat()->_lImgCookie;
}

inline const TCHAR* CTreeNode::GetCascadedfontFaceName()
{
    return GetCharFormat()->GetFaceName();
}

inline const TCHAR* CTreeNode::GetCascadedfontFamilyName()
{
    return GetCharFormat()->GetFamilyName();
}

inline BOOL CTreeNode::GetCascadedfontItalic()
{
    return (BOOL)(GetCharFormat()->_fItalic);
}

inline long CTreeNode::GetCascadedbackgroundImageCookie()
{
    return CTreeNode::GetFancyFormat()->_lImgCtxCookie;
}

inline styleCursor CTreeNode::GetCascadedcursor()
{
    return (styleCursor)(GetCharFormat()->_bCursorIdx);
}

inline styleTableLayout CTreeNode::GetCascadedtableLayoutEnum()
{
    return GetCascadedtableLayout()?styleTableLayoutFixed:styleTableLayoutAuto;
}

inline styleBorderCollapse CTreeNode::GetCascadedborderCollapseEnum()
{
    return GetCascadedborderCollapse()?styleBorderCollapseCollapse:styleBorderCollapseSeparate;
}

// This is used during layout to get the block level direction for the node.
inline styleDir CTreeNode::GetCascadedBlockDirection()
{
    return (styleDir)(GetParaFormat()->_fRTLInner?styleDirRightToLeft:styleDirLeftToRight);
}

// This is used for OM support to get the direction of any node
inline styleDir CTreeNode::GetCascadeddirection()
{
    return (styleDir)(GetCharFormat()->_fRTL?styleDirRightToLeft:styleDirLeftToRight);
}
inline styleBidi CTreeNode::GetCascadedunicodeBidi()
{
    return (styleBidi)(GetCharFormat()->_fBidiEmbed ? (GetCharFormat()->_fBidiOverride?styleBidiOverride:styleBidiEmbed) : styleBidiNormal);
}

inline styleLayoutGridMode CTreeNode::GetCascadedlayoutGridMode()
{
    styleLayoutGridMode mode = GetCharFormat()->GetLayoutGridMode(TRUE);
    return (mode!=styleLayoutGridModeNotSet)?mode:styleLayoutGridModeBoth;
}

inline styleLayoutGridType CTreeNode::GetCascadedlayoutGridType()
{
    styleLayoutGridType type = GetCharFormat()->GetLayoutGridType(TRUE);
    return (type!=styleLayoutGridTypeNotSet)?type:styleLayoutGridTypeLoose;
}

inline CUnitValue CTreeNode::GetCascadedlayoutGridLine()
{
    return GetParaFormat()->GetLineGridSize(TRUE);
}

inline CUnitValue CTreeNode::GetCascadedlayoutGridChar()
{
    return GetParaFormat()->GetCharGridSize(TRUE);
}

inline LONG CTreeNode::GetCascadedtextAutospace()
{
    return GetCharFormat()->_fTextAutospace;
};

inline styleWordBreak CTreeNode::GetCascadedwordBreak()
{
    return (GetParaFormat()->_fWordBreak==styleWordBreakNotSet
        ? styleWordBreakNormal : (styleWordBreak)GetParaFormat()->_fWordBreak);
}

inline styleLineBreak CTreeNode::GetCascadedlineBreak()
{
    return GetCharFormat()->_fLineBreakStrict?styleLineBreakStrict:styleLineBreakNormal;
}

inline styleTextJustify CTreeNode::GetCascadedtextJustify()
{
    return styleTextJustify(CTreeNode::GetParaFormat()->_uTextJustify);
}

inline styleTextJustifyTrim CTreeNode::GetCascadedtextJustifyTrim()
{
    return styleTextJustifyTrim(CTreeNode::GetParaFormat()->_uTextJustifyTrim);
}

inline CUnitValue CTreeNode::GetCascadedmarginTop()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvMarginTop;
}

inline CUnitValue CTreeNode::GetCascadedmarginRight()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvMarginRight;
}

inline CUnitValue CTreeNode::GetCascadedmarginBottom()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvMarginBottom;
}

inline CUnitValue CTreeNode::GetCascadedmarginLeft()
{
    return (CUnitValue)CTreeNode::GetFancyFormat()->_cuvMarginLeft;
}

inline CUnitValue CTreeNode::GetCascadedtextKashida()
{
    return CTreeNode::GetParaFormat()->_cuvTextKashida;
}

#endif //__XINDOWS_SITE_BASE_TREENODE_H__