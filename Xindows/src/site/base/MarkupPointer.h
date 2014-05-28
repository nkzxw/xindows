
#ifndef __XINDOWS_SITE_BASE_MARKUPPOINTER_H__
#define __XINDOWS_SITE_BASE_MARKUPPOINTER_H__

#include "../gen/mshtmext.h"

#define MPTR_SHOWSLAVE  0x00000001

// {5B2B058B-BB9D-48d0-B65C-FDAEEB95BDBA}
static const CLSID CLSID_CMarkupPointer = 
{ 0x5b2b058b, 0xbb9d, 0x48d0, { 0xb6, 0x5c, 0xfd, 0xae, 0xeb, 0x95, 0xbd, 0xba } };

class CMarkupPointer : public CBase, public IMarkupPointer
{
    DECLARE_CLASS_TYPES(CMarkupPointer, CBase);

    friend class CDocument; // for implementation of "OrSlave" versions of Left, Right, MoveToContainer on IHTMLViewServices
    friend class CMarkup;

public:
    DECLARE_MEMALLOC_NEW_DELETE()

    CMarkupPointer(CDocument* pDoc)
        : _pDoc(pDoc), _pMarkup(NULL), _pmpNext(NULL), _pmpPrev(NULL),
        _fRightGravity(FALSE), _fCling(FALSE), _fEmbedded(FALSE), _fKeepMarkupAlive(FALSE),
        _fAlwaysEmbed(FALSE), _ptpRef(NULL), _ichRef(0), _cpCache(-1), _verCp(0)
    {}

    virtual ~CMarkupPointer()
    {
        Unposition();
        Assert(!Markup());
    }

    //////////////////////////////////////////////
    // CBase methods
    DECLARE_PLAIN_IUNKNOWN(CMarkupPointer);
    DECLARE_PRIVATE_QI_FUNCS(CBase);

    virtual const CBase::CLASSDESC* GetClassDesc() const;

    ///////////////////////////////////////////////
    // IMarkupPointer methods
    STDMETHODIMP OwningDoc(IHTMLDocument2** ppDoc);
    STDMETHODIMP Gravity(POINTER_GRAVITY* peGravity);
    STDMETHODIMP SetGravity(POINTER_GRAVITY eGravity);
    STDMETHODIMP Cling(BOOL* pfCling);
    STDMETHODIMP SetCling(BOOL fCling);
    STDMETHODIMP MoveAdjacentToElement(IHTMLElement* pElement, ELEMENT_ADJACENCY eAdj);
    STDMETHODIMP MoveToPointer(IMarkupPointer* pPointer);
    STDMETHODIMP MoveToContainer(IMarkupContainer* pContainer, BOOL fAtStart);
    STDMETHODIMP Unposition();
    STDMETHODIMP IsPositioned(BOOL*);
    STDMETHODIMP GetContainer(IMarkupContainer**);

    STDMETHODIMP Left(
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        IHTMLElement**          ppElement,
        long*                   pcch,
        OLECHAR*                pchText);

    STDMETHODIMP Right(
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        IHTMLElement**          ppElement,
        long*                   pcch,
        OLECHAR*                pchText);

    STDMETHODIMP MoveUnit(MOVEUNIT_ACTION muAction);

    STDMETHODIMP CurrentScope(IHTMLElement** ppElemCurrent);

    STDMETHODIMP FindText( 
        OLECHAR*            pchFindText, 
        DWORD               dwFlags,
        IMarkupPointer*     pIEndMatch=NULL,
        IMarkupPointer*     pIEndSearch=NULL);

    STDMETHODIMP IsLeftOf           (IMarkupPointer* pPointer,     BOOL* pfResult);
    STDMETHODIMP IsLeftOfOrEqualTo  (IMarkupPointer* pPointer,     BOOL* pfResult);
    STDMETHODIMP IsRightOf          (IMarkupPointer* pPointer,     BOOL* pfResult);
    STDMETHODIMP IsRightOfOrEqualTo (IMarkupPointer* pPointer,     BOOL* pfResult);
    STDMETHODIMP IsEqualTo          (IMarkupPointer* pPointerThat, BOOL* pfAreEqual);

    HRESULT FindTextIdentity(long textID, CMarkupPointer* pPointerOtherEnd);
    HRESULT SetTextIdentity(CMarkupPointer* pPointerOtherEnd, long* plNewTextID);

    HRESULT IsInsideURL(IMarkupPointer* pRight, BOOL* pfResult);

    ///////////////////////////////////////////////
    // CMarkupPointer methodse
    HRESULT Left(
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        CTreeNode**             ppNode,
        long*                   pcch,
        OLECHAR*                pchText,
        long*                   plTextID)
    {
        return There(TRUE, fMove, pContext, ppNode, pcch, pchText, plTextID, 0);
    }

    HRESULT Right(
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        CTreeNode**             ppNode,
        long*                   pcch,
        OLECHAR*                pchText,
        long*                   plTextID)
    {
        return There(FALSE, fMove, pContext, ppNode, pcch, pchText, plTextID, 0);
    }

    int  Gravity() const { return _fRightGravity; }
    BOOL Cling() const { return _fCling; }
    BOOL KeepMarkupAlive() const { return _fKeepMarkupAlive; }
    void SetKeepMarkupAlive(BOOL fKeepAlive);
    BOOL AlwaysEmbed() const { return _fAlwaysEmbed; }
    void SetAlwaysEmbed(BOOL fAlwaysEmbed);

    HRESULT MoveAdjacentToElement(CElement* pElement, ELEMENT_ADJACENCY adj);
    HRESULT MoveToPointer(CMarkupPointer* pPointer);
    HRESULT MoveToContainer(CMarkup* pContainer, BOOL fAtStart, DWORD dwFlags=0);

    HRESULT MoveToGap(CTreePosGap* ptpg, CMarkup* pMarkup, BOOL fForceEmbedding=FALSE);

    HRESULT MoveToReference(CTreePos* ptp, long ich, CMarkup* pMarkup, long cpNew);

    HRESULT MoveToOrphan(CTreePos*);

    CTreeNode* CurrentScope(DWORD dwFlags=0);

    BOOL FindText(
        TCHAR*          pstr, 
        DWORD           dwFlags,
        CMarkupPointer* pEndMatch,
        CMarkupPointer* pEndSearch);

    BOOL IsEqualTo          (CMarkupPointer* pPointerThat);
    BOOL IsLeftOf           (CMarkupPointer* pPointerThat);
    BOOL IsLeftOfOrEqualTo  (CMarkupPointer* pPointerThat);
    BOOL IsRightOf          (CMarkupPointer* pPointerThat);
    BOOL IsRightOfOrEqualTo (CMarkupPointer* pPointerThat);

    HRESULT MoveToCp(long cp, CMarkup* pMarkup);

    HRESULT QueryBreaks(DWORD* pdwBreaks);

    // public helpers

    // called from CTreePos when it goes away
    void OnPositionReleased();

    CDocument* Doc() const { return _pDoc; }

    CTreeNode* Branch() { return _pMarkup?_ptp->GetInterNode():NULL; }

    BOOL IsPositioned() const { return _pMarkup!=NULL; }

    CMarkup* Markup() const { return _pMarkup; }

    void SetMarkup(CMarkup* pMarkup);

    // Get the embedded treepos.  Be careful, pointers are not always
    // embedded.
    CTreePos* GetEmbeddedTreePos()
    {
        Assert(_fEmbedded);
        return _fEmbedded?_ptpEmbeddedPointer:NULL;
    }

    // GetNormalizedReference returns a ptp/ich pair which is
    // immediately after REAL content.  Not real content is
    // pointers and empty text runs.
    CTreePos* GetNormalizedReference(long& ich) const;

    // Returns the cp for this pointer.  -1 if unpositioned;
    long GetCp();

    // "There" does the work of both Left and Right
    HRESULT There(
        BOOL                    fLeft,
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        CTreeNode**             ppNode, // Not AddRefed
        long*                   pcch,
        OLECHAR*                pchText,
        long*                   plTextID,
        DWORD*                  dwFlags);

    HRESULT There(
        BOOL                    fLeft,
        BOOL                    fMove,
        MARKUP_CONTEXT_TYPE*    pContext,
        IHTMLElement**          ppElement,
        long*                   pcch,
        OLECHAR*                pchText,
        DWORD*                  dwFlags);

private:
    HRESULT UnEmbed(CTreePos** pptp, long* pich);

    HRESULT Embed(CMarkup* pMarkup, CTreePos* ptp, long ich, long cpNew);

    // representation
    static const CBase::CLASSDESC s_classdesc; // classDesc (for CBase)

    CDocument* _pDoc; // The doc that owns me

    // The _pMarkup member points the markup I'm positioned in,
    // when I'm positioned.  If it is NULL, then I'm not positioned
    // in any markup, and thus the members which indicate where
    // are unused.
    CMarkup* _pMarkup;

    // Each markup has a list of markup pointers which are in the markup, but
    // do not have a pointer pos.
    CMarkupPointer* _pmpNext;
    CMarkupPointer* _pmpPrev;

    void AddMeToList();
    void RemoveMeFromList();

    // The members _fRightGravity and _fCling always indicate the state of
    // gravity and cling, even when we are embedded, when the pointer pos
    // has redundant indicators of this.  They are redundant because we want
    // gravity and cling with out CMarkupPointer.
    //
    // The _fEmbedded member is used when _pMarkup is non-NULL to indicate
    // if our position is indicated by an embedded pointer pos or a reference
    // to a non-pointer pos / offset pair.
    unsigned _fRightGravity    : 1; // When unpositioned, gravity is stored here
    unsigned _fCling           : 1; // When unpositioned, cling is stored here
    unsigned _fEmbedded        : 1; // Do I have have a pointer pos in the splay tree?
    unsigned _fKeepMarkupAlive : 1; // Keep addref on markup
    unsigned _fAlwaysEmbed     : 1; // Always embed this pointer

    // The _cpCache member records the cp this pointer is currently at.  If
    // the contents version of the markup matches that stored here, then
    // _cpCache is up to date.
    long _cpCache;
    long _verCp;

    long GetCpSlow() const;
    BOOL CpIsCached() const;

    void Validate() const {}

    union
    {
        CTreePos*       _ptp;                   // Quick access to ptp if embedded or not

        CTreePos*       _ptpEmbeddedPointer;    // When embeded, this points to Pointer pos

        struct                                  // When not embedded, this says where I am
        {
            CTreePos*   _ptpRef;                // I live just after this (non-pointer) pos
            long       _ichRef;                 // If text pos, this many chars into it
        };
    };

    NO_COPY(CMarkupPointer);
};

void EnsureLogicalOrder(CMarkupPointer*& pStart, CMarkupPointer*& pFinish);

#endif //__XINDOWS_SITE_BASE_MARKUPPOINTER_H__