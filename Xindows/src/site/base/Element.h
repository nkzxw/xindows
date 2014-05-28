
#ifndef __XINDOWS_SITE_BASE_ELEMENT_H__
#define __XINDOWS_SITE_BASE_ELEMENT_H__

#include <IntShCut.h>

#define DEFAULT_ATTR_SIZE   128

#define _hxx_
#include "../gen/element.hdl"

class CTreePos;
class CRootElement;
class CRequest;
class CStyle;
class CImgCtx;
class CStreamWriteBuff;

#define BREAK_NONE          0x0
#define BREAK_BLOCK_BREAK   0x1
#define BREAK_SITE_BREAK    0x2
#define BREAK_SITE_END      0x4


// A success code which implies that the action was successful, but
// somehow incomplete.  This is used in InitAttrBag if some attribute
// was not fully loaded.  This is also the first hresult value after
// the reserved ole values.
#define S_INCOMPLETE        MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x201)

enum SCROLLPIN
{
    SP_TOPLEFT = 1,     // Pin inner RECT to the top or left of the outer RECT
    SP_BOTTOMRIGHT,     // Pin inner RECT to the bottom or right of the outer RECT
    SP_MINIMAL,         // Calculate minimal scroll necessary to move the inner RECT into the outer RECT
    SP_MAX,
    SP_FORCE_LONG = (ULONG)-1 // Force this to be long for win16 world.
};

enum FOCUS_DIRECTION
{
    DIRECTION_BACKWARD,
    DIRECTION_FORWARD,
};

//+------------------------------------------------------------------------
//
//  Literals:   ENTERTREE types
//
//-------------------------------------------------------------------------
enum ENTERTREE_TYPE
{
    ENTERTREE_MOVE  = 0x1 << 0,             // [IN] This is the second half of a move opertaion
    ENTERTREE_PARSE = 0x1 << 1              // [IN] This is happening during parse
};

//+------------------------------------------------------------------------
//
//  Literals:   EXITTREE types
//
//-------------------------------------------------------------------------
enum EXITTREE_TYPE
{
    EXITTREE_MOVE               = 0x1<<0,	// [IN] This is the first half of a move operation
    EXITTREE_DESTROY            = 0x1<<1,	// [IN] The entire markup is being shut down -- 
                                            // element is not allowed to talk to other elements
    EXITTREE_PASSIVATEPENDING   = 0x1<<2,	// [IN] The tree has the last ref on this element
    EXITTREE_DELAYRELEASENEEDED = 0x1<<3	// [OUT]Element talks to outside world on final release
                                            // so release should be delayed
};

enum FOCUSSABILITY
{
    // Used by CElement::GetDefaultFocussability()
    // We may need to add more states, but we will think hard before adding them.
    //
    // All four modes are allowed for browse mode.
    // Only the first and the last (FOCUSSABILTY_NEVER and FOCUSSABILTY_TABBABLE)
    // are allowed for design mode.
    FOCUSSABILITY_NEVER         = 0,    // never focussable (and hence never tabbable)
    FOCUSSABILITY_MAYBE         = 1,    // possible to become focussable and tabbable
    FOCUSSABILITY_FOCUSSABLE    = 2,    // focussable by default, possible to become tabbable
    FOCUSSABILITY_TABBABLE      = 3,    // tabbable (and hence focussable) by default
};

//+---------------------------------------------------------------------------
//
//  Flag values for CDocument::OnPropertyChange
//
//----------------------------------------------------------------------------
enum FORMCHNG_FLAG
{
    FORMCHNG_LAYOUT  = (SERVERCHNG_LAST << 1),
    FORMCHNG_NOINVAL = (FORMCHNG_LAYOUT << 1),
    FORMCHNG_LAST    = FORMCHNG_NOINVAL,
    FORMCHNG_MAX     = LONG_MAX // needed to force enum to be dword on macintosh
};

//+-----------------------------------------------------------------------
//
//  Flag values for CELEMENT::OnPropertyChange
//
//------------------------------------------------------------------------
enum ELEMCHNG_FLAG
{
    ELEMCHNG_INLINESTYLE_PROPERTY	= FORMCHNG_LAST<<1,
    ELEMCHNG_CLEARCACHES			= FORMCHNG_LAST<<2,
    ELEMCHNG_SITEREDRAW				= FORMCHNG_LAST<<3,
    ELEMCHNG_REMEASURECONTENTS		= FORMCHNG_LAST<<4,
    ELEMCHNG_CLEARFF				= FORMCHNG_LAST<<5,
    ELEMCHNG_UPDATECOLLECTION		= FORMCHNG_LAST<<6,
    ELEMCHNG_SITEPOSITION			= FORMCHNG_LAST<<7,
    ELEMCHNG_RESIZENONSITESONLY		= FORMCHNG_LAST<<8,
    ELEMCHNG_SIZECHANGED			= FORMCHNG_LAST<<9,
    ELEMCHNG_REMEASUREINPARENT		= FORMCHNG_LAST<<10,
    ELEMCHNG_ACCESSIBILITY			= FORMCHNG_LAST<<11,
    ELEMCHNG_REMEASUREALLCONTENTS	= FORMCHNG_LAST<<12,
    ELEMCHNG_LAST					= FORMCHNG_LAST<<12
};

#define FEEDBACKRECTSIZE    1
#define GRABSIZE            7

//+------------------------------------------------------------------------
//
//  Function:   TagPreservationType
//
//  Synopsis:   Describes how white space is handled in a tag
//
//-------------------------------------------------------------------------
enum WSPT
{
    WSPT_PRESERVE,
    WSPT_COLLAPSE,
    WSPT_NEITHER
};

WSPT TagPreservationType(ELEMENT_TAG);
FOCUSSABILITY GetDefaultFocussability(ELEMENT_TAG);

// CInit2Context flags
enum
{
    INIT2FLAG_EXECUTE = 0x01
};

struct COnCommandExecParams
{
    inline COnCommandExecParams()
    {
        memset(this, 0, sizeof(COnCommandExecParams));
    }

    GUID*       pguidCmdGroup;
    DWORD       nCmdID;
    DWORD       nCmdexecopt;
    VARIANTARG* pvarargIn;
    VARIANTARG* pvarargOut;
};

// Used by NTYPE_QUERYFOCUSSABLE and NTYPE_QUERYTABBABLE
struct CQueryFocus
{
    long    _lSubDivision;
    BOOL    _fRetVal;
};

// Used by NTYPE_SETTINGFOCUS and NTYPE_SETFOCUS
struct CSetFocus
{
    CMessage*   _pMessage;
    long        _lSubDivision;
    HRESULT     _hr;
};

class CHtmTag;
class CNotification;
class CMessage;

//+----------------------------------------------------------------------------
//
//  Class:      CElement (element)
//
//   Note:      Derivation and virtual overload should be used to
//              distinguish different nodes.
//
//-----------------------------------------------------------------------------
class CElement : public CBase
{
    DECLARE_CLASS_TYPES(CElement, CBase)

    friend class CDBindMethods;
    friend class CLayout;
    friend class CFlowLayout;

private:
    DECLARE_MEMCLEAR_NEW_DELETE();

public:
    CElement(ELEMENT_TAG etag, CDocument* pDoc);
    virtual ~CElement();

    CDocument* Doc() const { return GetDocPtr(); }

    // creating thunks with AddRef and Release set to peer holder, if present
    HRESULT CreateTearOffThunk(
        void*       pvObject1,
        const void* apfn1,
        IUnknown*   pUnkOuter,
        void**      ppvThunk,
        void*       appropdescsInVtblOrder=NULL);

    HRESULT CreateTearOffThunk(
        void*       pvObject1,
        const void* apfn1,
        IUnknown*   pUnkOuter,
        void**      ppvThunk,
        void*       pvObject2,
        void*       apfn2,
        DWORD       dwMask,
        const IID* const* apIID,
        void*       appropdescsInVtblOrder=NULL);

    CAttrArray** GetAttrArray() const
    {
        return CBase::GetAttrArray();
    }
    void SetAttrArray(CAttrArray* pAA)
    {
        CBase::SetAttrArray(pAA);
    }

    // *********************************************
    //
    // ENUMERATIONS, CLASSES, & STRUCTS
    //
    // *********************************************
    enum ELEMENTDESC_FLAG
    {
        ELEMENTDESC_DONTINHERITSTYLE= (1 << 1), // Do not inherit style from parent
        ELEMENTDESC_NOANCESTORCLICK = (1 << 2), // We don't want our ancestors to fire clicks
        ELEMENTDESC_NOTIFYENDPARSE  = (1 << 3), // We want to be notified when we're parsed

        ELEMENTDESC_OLESITE         = (1 << 4), // class derived from COleSite
        ELEMENTDESC_OMREADONLY      = (1 << 5), // element's value can not be accessed through OM (input/file)
        ELEMENTDESC_ALLOWSCROLLING  = (1 << 6), // allow scrolling
        ELEMENTDESC_HASDEFDESCENT   = (1 << 7), // use 4 pixels descent for default vertical alignment
        ELEMENTDESC_BODY            = (1 << 8), // class is used the BODY element
        ELEMENTDESC_TEXTSITE        = (1 << 9), // class derived from CTxtSite
        ELEMENTDESC_ANCHOROUT       = (1 <<10), // draw anchor border outside site/inside
        ELEMENTDESC_SHOWTWS         = (1 <<11), // show trailing whitespaces
        ELEMENTDESC_XTAG            = (1 <<12), // an xtag - derived from CXElement
        ELEMENTDESC_NOOFFSETCTX     = (1 <<13), // Shift-F10 context menu shows on top-left (not offset)
        ELEMENTDESC_CARETINS_SL     = (1 <<14), // Dont select site after inserting site
        ELEMENTDESC_CARETINS_DL     = (1 <<15), // show caret (SameLine or DifferntLine)
        ELEMENTDESC_NOSELECT        = (1 <<16), // do not select site in edit mode
        ELEMENTDESC_TABLECELL       = (1 <<17), // site is a table cell. Also implies do not
                                                // word break before sites.
        ELEMENTDESC_VPADDING        = (1 <<18), // add a pixel vertical padding for this site
        ELEMENTDESC_EXBORDRINMOV    = (1 <<19), // Exclude the border in CSite::Move
        ELEMENTDESC_DEFAULT         = (1 <<20), // acts like a default site to receive return key
        ELEMENTDESC_CANCEL          = (1 <<21), // acts like a cancel site to receive esc key
        ELEMENTDESC_NOBKGRDRECALC   = (1 <<22), // Dis-allow background recalc
        ELEMENTDESC_NOPCTRESIZE     = (1 <<23), // Don't force resize due to percentage height/width
        ELEMENTDESC_NOLAYOUT        = (1 <<24), // element's like script and comment etc. cannot have layout
        ELEMENTDESC_NEVERSCROLL     = (1 <<26), // element can scroll
        ELEMENTDESC_CANSCROLL       = (1 <<27), // element can scroll
        ELEMENTDESC_LAST            = ELEMENTDESC_CANSCROLL,
        ELEMENTDESC_MAX             = LONG_MAX  // Force enum to DWORD on Macintosh.
    };

    NV_DECLARE_TEAROFF_METHOD(IsEqualObject, isequalobject, (IUnknown* ppunk));

    // ******************************************
    //
    // Virtual overrides
    //
    // ******************************************
    virtual void Passivate();

    DECLARE_PLAIN_IUNKNOWN(CElement);

    STDMETHOD(PrivateQueryInterface)(REFIID, void**);
    STDMETHOD_(ULONG, PrivateAddRef)();
    STDMETHOD_(ULONG, PrivateRelease)();

    void PrivateEnterTree();
    void PrivateExitTree(CMarkup* pMarkupOld);

    virtual HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags);

    virtual HRESULT GetNaturalExtent(DWORD dwExtentMode, SIZEL* psizel) { return E_FAIL; }

    typedef enum
    {
        GETINFO_ISCOMPLETED,    // Has loading of an element completed
        GETINFO_HISTORYCODE     // Code used to validate history
    } GETINFO;

    virtual DWORD GetInfo(GETINFO gi);

    DWORD HistoryCode() { return GetInfo(GETINFO_HISTORYCODE); }

    NV_DECLARE_TEAROFF_METHOD(get_tabIndex, GET_tabIndex, (short*));

    // The dir property is declared baseimplementation in the pdl.
    // We need prototype here
    NV_DECLARE_TEAROFF_METHOD(put_dir, PUT_dir, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_dir, GET_dir, (BSTR* p));

    // these delegaters implement redirection to the window object
    NV_DECLARE_TEAROFF_METHOD(get_onload, GET_onload, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onload, PUT_onload, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onunload, GET_onunload, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onunload, PUT_onunload, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onfocus, GET_onfocus, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onfocus, PUT_onfocus, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onblur, GET_onblur, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onblur, PUT_onblur, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onbeforeunload, GET_onbeforeunload, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onbeforeunload, PUT_onbeforeunload, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onhelp, GET_onhelp, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onhelp, PUT_onhelp, (VARIANT));

    NV_DECLARE_TEAROFF_METHOD(get_onscroll, GET_onscroll, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onscroll, PUT_onscroll, (VARIANT));
    NV_DECLARE_TEAROFF_METHOD(get_onresize, GET_onresize, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onresize, PUT_onresize, (VARIANT));

    // non-abstract getters for tagName and scopeName. See element.pdl
    NV_DECLARE_TEAROFF_METHOD(GettagName, gettagname, (BSTR*));

    // IServiceProvider methods
    NV_DECLARE_TEAROFF_METHOD(QueryService, queryservice, (REFGUID guidService, REFIID iid, LPVOID* ppv));

    virtual CAtomTable* GetAtomTable(BOOL* pfExpando=NULL);

    // init / deinit methods
    class CInit2Context
    {
    public:
        CInit2Context(CHtmTag* pht, CMarkup* pTargetMarkup, DWORD dwFlags) :
          _pTargetMarkup(pTargetMarkup), _pht(pht), _dwFlags(dwFlags)
          {
          }

          CInit2Context(CHtmTag* pht, CMarkup* pTargetMarkup) :
          _pTargetMarkup(pTargetMarkup), _pht(pht), _dwFlags(0)
          {
          }

          CHtmTag*  _pht;
          CMarkup*  _pTargetMarkup;
          DWORD     _dwFlags;
    };

    virtual HRESULT Init();
    virtual HRESULT Init2(CInit2Context* pContext);

    HRESULT         InitAttrBag(CHtmTag* pht);
    HRESULT         MergeAttrBag(CHtmTag* pht);

    virtual void    Notify (CNotification* pnf);

    HRESULT         EnterTree();
    void            ExitTree(DWORD dwExitFlags);

    // other
    CBase*          GetOmWindow(void);

    // Get the Base Object that owns the attr array for a given property
    // Allows us to re-direct properties to another objects storage
    CBase*          GetBaseObjectFor(DISPID dispID);

    // Pass the call to the form.
    //-------------------------------------------------------------------------
    //  +override : special process
    //  +call super : first
    //  -call parent : no
    //  -call children : no
    //-------------------------------------------------------------------------
    virtual HRESULT CloseErrorInfo(HRESULT hr);

    // Scrolling methods
    // Scroll this element into view
    virtual HRESULT ScrollIntoView(
        SCROLLPIN spVert=SP_MINIMAL,
        SCROLLPIN spHorz=SP_MINIMAL,
        BOOL fScrollBits=TRUE);

    HRESULT DeferScrollIntoView(
        SCROLLPIN spVert=SP_MINIMAL,
        SCROLLPIN spHorz=SP_MINIMAL);

    NV_DECLARE_ONCALL_METHOD(DeferredScrollIntoView, deferredscrollintoview, (DWORD_PTR dwParam));

    // Relative element (non-site) helper methods
    virtual HTC HitTestPoint(CMessage* pMessage, CTreeNode** ppNodeElement, DWORD dwFlags);

    BOOL CanHandleMessage()
    {
        return (HasLayout())?(IsEnabled()&&IsVisible(TRUE)):(TRUE);
    }

    // STDMETHODCALLTYPE required when passing fn through generic ptr
    virtual HRESULT STDMETHODCALLTYPE HandleMessage(CMessage* pMessage);

    HRESULT STDMETHODCALLTYPE HandleCaptureMessage(CMessage* pMessage)
    {
        return HandleMessage(pMessage);
    }

    // marka these are to be DELETED !!
    BOOL    DisallowSelection();

    // set the state of the IME.
    HRESULT SetImeState();
    HRESULT ComputeExtraFormat(DISPID dispID, BOOL fInherits, CTreeNode* pTreeNode, VARIANT* pVarReturn);

    // DoClick() is called by click(). It is also called internally in response
    // to a mouse click by user.
    // DoClick() fires the click event and then calls ClickAction() if the event
    // is not cancelled.
    // Derived classes can override ClickAction() to provide specific functionality.
    virtual HRESULT DoClick(
        CMessage*   pMessage=NULL,
        CTreeNode*  pNodeContext=NULL,
        BOOL        fFromLabel=FALSE);
    virtual HRESULT ClickAction(CMessage* pMessage);

    virtual HRESULT ShowTooltip(CMessage* pmsg, POINT pt);

    HRESULT SetCursorStyle(LPCTSTR pstrIDC, CTreeNode* pNodeContext=NULL);

    void    Invalidate();

    // Element rect and element invalidate support
    enum GERC_FLAGS
    {
        GERC_ONALINE = 1,
        GERC_CLIPPED = 2
    };

    void    GetElementRegion(CDataAry<RECT>* paryRects, RECT* prcBound=NULL, DWORD dwFlags=0);
    HRESULT GetElementRc(RECT* prc, DWORD dwFlags, POINT* ppt=NULL);

    // these helper functions return in container coordinates
    void    GetBoundingSize(SIZE& sz);
    HRESULT GetBoundingRect(CRect* pRect, DWORD dwFlags=0);
    HRESULT GetElementTopLeft(POINT& pt);

    // helper to return the actual background color
    COLORREF GetInheritedBackgroundColor(CTreeNode* pNodeContext=NULL);
    HRESULT  GetInheritedBackgroundColorValue(CColorValue* pVal, CTreeNode* pNodeContext=NULL);
    virtual HRESULT GetColors(CColorInfo* pCI);
    COLORREF GetBackgroundColor()
    {
        CTreeNode* pNodeContext = GetFirstBranch();
        CTreeNode* pNodeParent  = pNodeContext->Parent() ? pNodeContext->Parent() : pNodeContext;

        return pNodeParent->Element()->GetInheritedBackgroundColor(pNodeParent);
    }

    // Events related stuff
    //--------------------------------------------------------------------------------------
    inline BOOL ShouldFireEvents() { return _fEventListenerPresent; }
    inline void SetEventsShouldFire() { _fEventListenerPresent = TRUE;  }
    BOOL    FireCancelableEvent(DISPID dispidMethod, DISPID dispidProp, LPCTSTR pchEventType, BYTE* pbTypes, ...);
    BOOL    BubbleCancelableEvent(CTreeNode* pNodeContext, long lSubDivision, DISPID dispidMethod, DISPID dispidProp, LPCTSTR pchEventType, BYTE* pbTypes, ...);
    HRESULT FireEventHelper(DISPID dispidMethod, DISPID dispidProp, BYTE* pbTypes,  ...);
    HRESULT FireEvent(DISPID dispidMethod, DISPID dispidProp, LPCTSTR pchEventType, BYTE* pbTypes, ...);
    HRESULT BubbleEvent(CTreeNode* pNodeContext, long lSubDivision, DISPID dispidMethod, DISPID dispidProp, LPCTSTR pchEventType, BYTE* pbTypes, ...);
    HRESULT BubbleEventHelper(CTreeNode* pNodeContext, long lSubDivision, DISPID dispidMethod, DISPID dispidProp, BOOL fRaisedByPeer, VARIANT* pvb, BYTE* pbTypes, ...);
    virtual HRESULT DoSubDivisionEvents(long lSubDivision, DISPID dispidMethod, DISPID dispidProp, VARIANT* pvb, BYTE* pbTypes, ...);
    HRESULT FireStdEventOnMessage(
        CTreeNode* pNodeContext,
        CMessage* pMessage,
        CTreeNode* pNodeBeginBubbleWith=NULL,
        CTreeNode* pNodeEvent=NULL);
    BOOL FireStdEvent_KeyDown(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode, short shift);
    BOOL FireStdEvent_KeyUp(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode, short shift);
    BOOL FireStdEvent_KeyPress(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode);
    BOOL FireStdEvent_MouseHelper(
        CTreeNode*  pNodeContext,
        CMessage *  pMessage,
        DISPID      dispidMethod,
        DISPID      dispidProp,
        short       button,
        short       shift,
        float       x,
        float       y,
        CTreeNode*  pNodeFrom=NULL,
        CTreeNode*  pNodeTo=NULL,
        CTreeNode*  pNodeBeginBubbleWith=NULL,
        CTreeNode* pNodeEvent=NULL);
    void    Fire_onpropertychange(LPCTSTR strPropName);
    void    Fire_onscroll();
    HRESULT Fire_PropertyChangeHelper(DISPID dispid, DWORD dwFlags);

    void STDMETHODCALLTYPE Fire_onfocus(DWORD_PTR dwContext);
    void STDMETHODCALLTYPE Fire_onblur(DWORD_PTR dwContext);

    BOOL DragElement(
        CLayout*                    pFlowLayout,
        DWORD                       dwStateKey,
        IUniformResourceLocator*    pUrlToDrag,
        long                        lSubDivision);

    BOOL Fire_ondragHelper(
        long    lSubDivision,
        DISPID  dispidEvent,
        DISPID  dispidProp,
        LPCTSTR pchType,
        DWORD*  pdwDropEffect);
    void Fire_ondragend(long lSubDivision, DWORD dwDropEffect);

    // Associated image context helpers
    HRESULT GetImageUrlCookie(LPCTSTR lpszURL, long* plCtxCookie);
    HRESULT AddImgCtx(DISPID dispID, LONG lCookie);
    void    ReleaseImageCtxts();
    void    DeleteImageCtx(DISPID dispid);

    // Internal events
    HRESULT EnsureFormatCacheChange(DWORD dwFlags);
    BOOL    IsFormatCacheValid();

    // Element Tree methods
    BOOL IsAligned();
    BOOL IsContainer();
    BOOL IsNoScope();
    BOOL IsBlockElement();
    BOOL IsBlockTag();
    BOOL IsOwnLineElement(CFlowLayout* pFlowLayoutContext);
    BOOL IsRunOwner() { return _fOwnsRuns; }
    BOOL BreaksLine();
    BOOL IsTagAndBlock(ELEMENT_TAG eTag) { return Tag()==eTag&&IsBlockElement(); }
    BOOL IsFlagAndBlock(TAGDESC_FLAGS flag) { return HasFlag(flag)&&IsBlockElement(); }

    HRESULT     ClearRunCaches(DWORD dwFlags);

    // Get the first branch for this element
    CTreeNode*  GetFirstBranch() const { return __pNodeFirstBranch; }
    CTreeNode*  GetLastBranch();
    BOOL        IsOverlapped();

    // Get the CTreePos extent of this element
    void        GetTreeExtent(CTreePos** pptpStart, CTreePos** pptpEnd);

    //-------------------------------------------------------------------------
    //
    // Layout related functions
    //
    //-------------------------------------------------------------------------
private:
    BOOL HasLayoutLazy();   // should only be called by HasLayout
    inline CLayout* GetLayoutLazy();
    // Create the layout object to be associated with the current element
    virtual HRESULT CreateLayout();

public:
    // CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION
    // CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION
    // (please read the comments below)
    //
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
    CLayout*    GetCurLayout() { return GetLayoutPtr(); }
    BOOL        HasLayout() { return !!HasLayoutPtr(); }

    CLayout*    GetCurNearestLayout();
    CTreeNode*  GetCurNearestLayoutNode();
    CElement*   GetCurNearestLayoutElement();

    CLayout*    GetCurParentLayout();
    CTreeNode*  GetCurParentLayoutNode();
    CElement*   GetCurParentLayoutElement();

    // CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION
    // CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION
    // (please read the comments above)
    //
    //
    // the following get functions may create the layout if it is not
    // created yet.
    inline CLayout*     GetUpdatedLayout();     // checks for NeedsLayout()
    inline CLayout*     GetUpdatedLayoutPtr();  // Call this if NeedsLayout() is already called
    inline BOOL         NeedsLayout();

    CLayout*            GetUpdatedNearestLayout();
    CTreeNode*          GetUpdatedNearestLayoutNode();
    CElement*           GetUpdatedNearestLayoutElement();

    CLayout*            GetUpdatedParentLayout();
    CTreeNode*          GetUpdatedParentLayoutNode();
    CElement*           GetUpdatedParentLayoutElement();


    void                DestroyLayout();

    // this functions return GetFirstBranch()->Parent()->Ancestor(etag)->Element(); safely
    // returns NULL if the element is not in the tree, or it doesn't have a parent, or no such ansectors in the tree, etc.
    CElement*           GetParentAncestorSafe(ELEMENT_TAG etag) const;
    CElement*           GetParentAncestorSafe(ELEMENT_TAG* arytag) const;

    // BUGBUG - these functions should go, we should not need
    // to know if the element has flowlayout.
    CFlowLayout*        GetFlowLayout();
    CTreeNode*          GetFlowLayoutNode();
    CElement*           GetFlowLayoutElement();
    CFlowLayout*        HasFlowLayout();

    // Notification helpers
    void InvalidateElement(DWORD grfFlags=0);
    void MinMaxElement(DWORD grfFlags=0);
    void ResizeElement(DWORD grfFlags=0);
    void RemeasureElement(DWORD grfFlags=0);
    void RemeasureInParentContext(DWORD grfFlags=0);
    void RepositionElement(DWORD grfFlags=0);
    void ZChangeElement(DWORD grfFlags=0, CPoint* ppt=NULL);

    void SendNotification(enum NOTIFYTYPE ntype, DWORD grfFlags=0, void* pvData=0);
    void SendNotification(enum NOTIFYTYPE ntype, void* pvData)
    {
        SendNotification(ntype, 0, pvData);
    }
    void SendNotification(CNotification* pNF);

    long GetSourceIndex(void);

    long CompareZOrder(CElement* pElement);

    // Mark an element's layout (if any) dirty
    void DirtyLayout(DWORD grfLayout=0);
    BOOL OpenView();

    BOOL HasPercentBgImg();

    // cp and run related helper functions
    long GetFirstCp();
    long GetLastCp();
    long GetElementCch();
    long GetFirstAndLastCp(long* pcpFirst, long* pcpLast);

    // get the border information related to the element
    virtual DWORD GetBorderInfo(CDocInfo* pdci, CBorderInfo* pbi, BOOL fAll=FALSE);

    HRESULT GetRange(long* pcp, long* pcch);
    HRESULT GetPlainTextInScope(CString* pstrText);

    virtual HRESULT Save(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd);
    HRESULT WriteTag(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd, BOOL fForce=FALSE);
    virtual HRESULT SaveAttributes(CStreamWriteBuff* pStreamWrBuff, BOOL* pfAny=NULL);
    HRESULT SaveAttributes(IPropertyBag* pPropBag, BOOL fSaveBlankAttributes=TRUE);
    HRESULT SaveAttribute(
        CStreamWriteBuff*   pStreamWrBuff,
        LPTSTR              pchName,
        LPTSTR              pchValue,
        const PROPERTYDESC* pPropDesc=NULL,
        CBase*              pBaseObj=NULL,
        BOOL                fEqualSpaces=TRUE,
        BOOL                fAlwaysQuote=FALSE);

    ELEMENT_TAG Tag() const { return (ELEMENT_TAG)_etag; }
    inline ELEMENT_TAG TagType() const
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

    virtual const TCHAR* TagName();
    const TCHAR* NamespaceHtml();

    // Support for sub-objects created through pdl's
    // CStyle & Celement implement this differently
    CElement* GetElementPtr() { return this; }

    BOOL CanShow();

    BOOL HasFlag(TAGDESC_FLAGS) const;

    static void ReplacePtr(CElement** pplhs, CElement* prhs);
    static void ReplacePtrSub(CElement** pplhs, CElement* prhs);
    static void SetPtr(CElement** pplhs, CElement* prhs);
    static void ClearPtr(CElement** pplhs);
    static void StealPtrSet(CElement** pplhs, CElement* prhs);
    static void StealPtrReplace(CElement** pplhs, CElement* prhs);
    static void ReleasePtr(CElement* pElement);

    // Write unknown attr set
    HRESULT SaveUnknown(CStreamWriteBuff* pStreamWrBuff, BOOL* pfAny=NULL);
    HRESULT SaveUnknown(IPropertyBag* pPropBag, BOOL fSaveBlankAttributes=TRUE);

    // Helpers
    BOOL IsNamed() const { return !!_fIsNamed; }

    // comparison
    LPCTSTR NameOrIDOfParentForm();

    // Helper for name or ID
    LPCTSTR GetIdentifier(void);
    HRESULT GetUniqueIdentifier(CString* pcstr, BOOL fSetWhenCreated=FALSE, BOOL* pfDidCreate=NULL);
    LPCTSTR GetAAname() const;
    LPCTSTR GetAAsubmitname() const;

    // Paste helpers
    enum Where
    {
        Inside,
        Outside,
        BeforeBegin,
        AfterBegin,
        BeforeEnd,
        AfterEnd
    };

    HRESULT Inject(Where, BOOL fIsHtml, LPTSTR pStr, long cch);

    virtual HRESULT PasteClipboard() { return S_OK; }

    HRESULT InsertAdjacent(Where where, CElement* pElementInsert);

    HRESULT RemoveOuter();

    // Helper to get the specified text under the element -- invokes saver.
    HRESULT GetText(BSTR* pbstr, DWORD dwStm);

    // Another helper for databinding
    HRESULT GetBstrFromElement(BOOL fHTML, BSTR* pbstr);

    // Collection Management helpers
    NV_DECLARE_PROPERTY_METHOD(GetIDHelper, GETIDHelper, (CString* pf));
    NV_DECLARE_PROPERTY_METHOD(SetIDHelper, SETIDHelper, (CString* pf));
    NV_DECLARE_PROPERTY_METHOD(GetnameHelper, GETNAMEHelper, (CString* pf));
    NV_DECLARE_PROPERTY_METHOD(SetnameHelper, SETNAMEHelper, (CString* pf));
    LPCTSTR GetAAid() const;
    void    InvalidateCollection(long lIndex);
    NV_STDMETHOD(removeAttribute)(BSTR strPropertyName, LONG lFlags, VARIANT_BOOL* pfSuccess);
    HRESULT SetUniqueNameHelper(LPCTSTR szUniqueName);
    HRESULT SetIdentifierHelper(LPCTSTR lpszValue, DISPID dspIDThis, DISPID dspOther1, DISPID dspOther2);
    void    OnEnterExitInvalidateCollections(BOOL);
    void    DoElementNameChangeCollections(void);

    // Clone - make a duplicate new element
    //
    // !! returns an element with no node!
    virtual HRESULT Clone(CElement** ppElementClone, CDocument* pDoc);

    void ComputeHorzBorderAndPadding(
        CCalcInfo*  pci,
        CTreeNode*  pNodeContext,
        CElement*   pTxtSite,
        LONG*       pxBorderLeft,
        LONG*       pxPaddingLeft,
        LONG*       pxBorderRight,
        LONG*       pxPaddingRight);

    HRESULT SetDim(
        DISPID                      dispID,
        float                       fValue,
        CUnitValue::UNITVALUETYPE   uvt,
        long                        lDimOf,
        CAttrArray**                ppAA,
        BOOL                        fInlineStyle,
        BOOL*                       pfChanged);

    virtual HRESULT ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget);
    virtual HRESULT ApplyDefaultFormat(CFormatInfo* pCFI);
    BOOL    ElementNeedsLayout(CFormatInfo* pCFI);
    BOOL    DetermineBlockness(CFormatInfo* pCFI);
    HRESULT ApplyInnerOuterFormats(CFormatInfo* pCFI);
    CImgCtx* GetBgImgCtx();

    // Access Key Handling Helper Functions
    FOCUS_ITEM GetMnemonicTarget();
    HRESULT HandleMnemonic(CMessage* pmsg, BOOL fDoClick, BOOL* pfYieldFailed=NULL);
    HRESULT GotMnemonic(CMessage* pMessage);
    HRESULT LostMnemonic(CMessage* pMessage);  
    BOOL    MatchAccessKey(CMessage* pmsg);
    HRESULT OnTabIndexChange();

    // Styles
    HRESULT      GetStyleObject(CStyle** ppStyle);
    CAttrArray*  GetInLineStyleAttrArray();
    CAttrArray** CreateStyleAttrArray(DISPID);

    BOOL         HasInLineStyles(void);
    BOOL         HasClassOrID(void);

    CStyle*      GetInLineStylePtr();
    CStyle*      GetRuntimeStylePtr();

    // Helpers for abstract name properties implemented on derived elements
    DECLARE_TEAROFF_METHOD(put_name, PUT_name, (BSTR v));
    DECLARE_TEAROFF_METHOD(get_name, GET_name, (BSTR* p));

    htmlBlockAlign      GetParagraphAlign(BOOL fOuter);
    htmlControlAlign    GetSiteAlign();

    BOOL IsInlinedElement();

    BOOL IsPositionStatic();
    BOOL IsPositioned();
    BOOL IsAbsolute();
    BOOL IsRelative();
    BOOL IsInheritingRelativeness();
    BOOL IsScrollingParent();
    BOOL IsClipParent();
    BOOL IsZParent();
    BOOL IsLocked();
    BOOL IsDisplayNone();
    BOOL IsVisibilityHidden();

    // Reset functionallity
    // Returns S_OK if successful and E_NOTIMPL if not applicable
    virtual HRESULT DoReset(void) { return E_NOTIMPL; }

    // Get control's window, does control have window?
    virtual HWND GetHwnd() { return NULL; }

    // Take the capture.
    void TakeCapture(BOOL fTake);
    BOOL HasCapture();

    // IMsoCommandTarget methods
    HRESULT STDMETHODCALLTYPE QueryStatus(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext);
    HRESULT STDMETHODCALLTYPE Exec(
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut);

    BOOL IsEditable(BOOL fCheckContainerOnly=FALSE);

    HRESULT EnsureInMarkup();

    // Currency / UI-Activity

    // Does this element (or its master) have currency?
    BOOL HasCurrency();

    virtual HRESULT RequestYieldCurrency(BOOL fForce);

    // Relinquishes currency
    virtual HRESULT YieldCurrency(CElement* pElemNew);

    // Relinquishes UI
    virtual void YieldUI(CElement* pElemNew);

    // Forces UI upon an element
    virtual HRESULT BecomeUIActive();

    BOOL    NoUIActivate(); // tell us if element can be UIActivate

    BOOL    IsFocussable(long lSubDivision);
    BOOL    IsTabbable(long lSubDivision);

    HRESULT PreBecomeCurrent(long lSubDivision, CMessage* pMessage);
    HRESULT BecomeCurrentFailed(long lSubDivision, CMessage* pMessage);
    HRESULT PostBecomeCurrent(CMessage* pMessage);

    HRESULT BecomeCurrent(
        long        lSubDivision,
        BOOL*       pfYieldFailed=NULL,
        CMessage*   pMessage=NULL,
        BOOL        fTakeFocus=FALSE);

    HRESULT BubbleBecomeCurrent(
        long        lSubDivision,
        BOOL*       pfYieldFailed=NULL,
        CMessage*   pMessage=NULL,
        BOOL        fTakeFocus=FALSE);

    CElement* GetFocusBlurFireTarget(long lSubDiv);

    HRESULT focusHelper(long lSubDivision);

    virtual HRESULT GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape);

    // Forces Currency and uiactivity upon an element
    HRESULT BecomeCurrentAndActive(
        CMessage*   pmsg=NULL,
        long        lSubDivision=0,
        BOOL        fTakeFocus=FALSE, 
        BOOL*       pfYieldFailed=NULL);
    HRESULT BubbleBecomeCurrentAndActive(CMessage* pmsg=NULL, BOOL fTakeFocus=FALSE);

    virtual HRESULT GetSubDivisionCount(long* pc);
    virtual HRESULT GetSubDivisionTabs(long* pTabs, long c);
    virtual HRESULT SubDivisionFromPt(POINT pt, long* plSub);

    // Find an element with the given set of SITE_FLAG's set
    CElement* FindDefaultElem(BOOL fDefault, BOOL fFull=FALSE);

    // set the default element in a form or in the doc
    void SetDefaultElem(BOOL fFindNew = FALSE);

    HRESULT GetNextSubdivision(FOCUS_DIRECTION dir, long lSubDivision, long* plSubNew);

    virtual BOOL IsEnabled();

    virtual BOOL IsValid() { return TRUE; }

    BOOL IsVisible(BOOL fCheckParent);
    BOOL IsParent(CElement* pElement); // Is pElement a parent of this element?

    virtual HRESULT GetControlInfo(CONTROLINFO* pCI) { return E_FAIL; }
    virtual BOOL OnMenuEvent(int id, UINT code) { return FALSE; }
    HRESULT STDMETHODCALLTYPE OnCommand(int id, HWND hwndCtrl, UINT codeNotify) { return S_OK; }
    HRESULT OnContextMenu(int x, int y, int id);

    HRESULT OnInitMenuPopup(HMENU hmenu, int item, BOOL fSystemMenu);
    HRESULT OnMenuSelect(UINT uItem, UINT fuFlags, HMENU hmenu);

    // Helper for translating keystrokes into commands
    HRESULT PerformTA(CMessage* pMessage);

    CImgCtx* GetNearestBgImgCtx();

    DWORD GetCommandID(LPMSG lpmsg);

    //+---------------------------------------------------------------------------
    //
    //  Flag values for CElement::CLock
    //
    //----------------------------------------------------------------------------
    enum ELEMENTLOCK_FLAG
    {
        ELEMENTLOCK_NONE            = 0,
        ELEMENTLOCK_CLICK           = 1 <<  0,
        ELEMENTLOCK_PROCESSPOSITION = 1 <<  1,
        ELEMENTLOCK_PROCESSADORNERS = 1 <<  2,
        ELEMENTLOCK_DELETE          = 1 <<  3,
        ELEMENTLOCK_FOCUS           = 1 <<  4,
        ELEMENTLOCK_CHANGE          = 1 <<  5,
        ELEMENTLOCK_UPDATE          = 1 <<  6,
        ELEMENTLOCK_SIZING          = 1 <<  7,
        ELEMENTLOCK_COMPUTEFORMATS  = 1 <<  8,
        ELEMENTLOCK_QUERYSTATUS     = 1 <<  9,
        ELEMENTLOCK_BLUR            = 1 << 10,
        ELEMENTLOCK_RECALC          = 1 << 11,
        ELEMENTLOCK_BLOCKCALC       = 1 << 12,
        ELEMENTLOCK_ATTACHPEER      = 1 << 13,
        ELEMENTLOCK_PROCESSREQUESTS = 1 << 14,
        ELEMENTLOCK_PROCESSMEASURE  = 1 << 15,
        ELEMENTLOCK_LAST            = 1 << 15,

        // don't add anymore flags, we only have 16 bits
    };

    //+-----------------------------------------------------------------------
    //
    //  CElement::CLock
    //
    //------------------------------------------------------------------------
    class CLock
    {
    public:
        DECLARE_MEMALLOC_NEW_DELETE();
        CLock(CElement* pElement, ELEMENTLOCK_FLAG enumLockFlags=ELEMENTLOCK_NONE);
        ~CLock();

    private:
        CElement* _pElement;
        WORD _wLockFlags;
    };

    BOOL TestLock(ELEMENTLOCK_FLAG enumLockFlags) { return _wLockFlags&((WORD)enumLockFlags); }

    BOOL TestClassFlag(ELEMENTDESC_FLAG dwFlag) const { return ElementDesc()->_classdescBase._dwFlags&dwFlag; }

    inline BOOL WantEndParseNotification() { return TestClassFlag(CElement::ELEMENTDESC_NOTIFYENDPARSE); }

    // BUGBUG: we should have a general notification mechanism to tell what
    // elements are listening to which notifications
    BOOL WantTextChangeNotifications();

    //+-----------------------------------------------------------------------
    //
    //  CLASSDESC (class descriptor)
    //
    //------------------------------------------------------------------------
    class ACCELS
    {
    public:
        ACCELS(ACCELS* pSuper, WORD wAccels);
        ~ACCELS();
        HRESULT EnsureResources();
        HRESULT LoadAccelTable();
        DWORD   GetCommandID(LPMSG pmsg);

        ACCELS* _pSuper;

        BOOL    _fResourcesLoaded;

        WORD    _wAccels;
        LPACCEL _pAccels;
        int     _cAccels;
    };

    struct CLASSDESC
    {
        CBase::CLASSDESC _classdescBase;
        void* _apfnTearOff;

        BOOL TestFlag(ELEMENTDESC_FLAG dw) const { return (_classdescBase._dwFlags&dw)!=0; }

        // move from CSite::CLASSDESC
        ACCELS* _pAccelsDesign;
        ACCELS* _pAccelsRun;
    };

    const CLASSDESC* ElementDesc() const
    {
        return (const CLASSDESC*)BaseDesc();
    }

public:
    // Lookaside pointers
    enum
    {
        LOOKASIDE_FILTER            = 0,
        LOOKASIDE_DATABIND          = 1,
        LOOKASIDE_PEER              = 2,
        LOOKASIDE_PEERMGR           = 3,
        LOOKASIDE_ACCESSIBLE        = 4,
        LOOKASIDE_SLAVEMARKUP       = 5,
        LOOKASIDE_REQUEST           = 6,
        LOOKASIDE_ELEMENT_NUMBER    = 7

        // *** There are only 7 bits reserved in the element.
        // *** if you add more lookasides you have to make sure 
        // *** that you make room for those bits.
    };

private:
    BOOL        HasLookasidePtr(int iPtr)                       { return (_fHasLookasidePtr&(1<<iPtr)); }
    void*       GetLookasidePtr(int iPtr);
    HRESULT     SetLookasidePtr(int iPtr, void* pv);
    void*       DelLookasidePtr(int iPtr);

public:

    BOOL        HasLayoutPtr() const                            { return _fHasLayoutPtr; }
    CLayout*    GetLayoutPtr() const;
    void        SetLayoutPtr(CLayout* pLayout);
    CLayout*    DelLayoutPtr();

    BOOL        IsInMarkup() const                              { return _fHasMarkupPtr; }
    BOOL        HasMarkupPtr() const                            { return _fHasMarkupPtr; }
    CMarkup*    GetMarkupPtr() const;
    void        SetMarkupPtr(CMarkup* pMarkup);
    void        DelMarkupPtr();

    CRootElement* IsRoot()                                      { return Tag()==ETAG_ROOT?(CRootElement*)this:NULL; }
    CMarkup*    GetMarkup() const                               { return GetMarkupPtr(); }
    BOOL        IsInPrimaryMarkup() const;
    BOOL        IsInThisMarkup(CMarkup* pMarkup) const;
    CRootElement* MarkupRoot();

    CElement*   MarkupMaster() const;
    CMarkup*    SlaveMarkup() const                             { return ((CElement*)this)->GetSlaveMarkupPtr(); }
    CElement*   FireEventWith();

    CDocument*  GetDocPtr() const;

    BOOL        HasSlaveMarkupPtr()                             { return (HasLookasidePtr(LOOKASIDE_SLAVEMARKUP)); }
    CMarkup*    GetSlaveMarkupPtr()                             { return ((CMarkup*)GetLookasidePtr(LOOKASIDE_SLAVEMARKUP)); }
    HRESULT     SetSlaveMarkupPtr(CMarkup* pSlaveMarkup)        { return (SetLookasidePtr(LOOKASIDE_SLAVEMARKUP, pSlaveMarkup)); }
    CMarkup*    DelSlaveMarkupPtr()                             { return ((CMarkup*)DelLookasidePtr(LOOKASIDE_SLAVEMARKUP)); }

    BOOL        HasRequestPtr()                                 { return (HasLookasidePtr(LOOKASIDE_REQUEST)); }
    CRequest*   GetRequestPtr()                                 { return ((CRequest*)GetLookasidePtr(LOOKASIDE_REQUEST)); }
    HRESULT     SetRequestPtr(CRequest* pRequest)               { return (SetLookasidePtr(LOOKASIDE_REQUEST, pRequest)); }
    CRequest*   DelRequestPtr()                                 { return ((CRequest*)DelLookasidePtr(LOOKASIDE_REQUEST)); }

    long        GetReadyState();
    virtual void OnReadyStateChange();

    // The element is the head of a linked list of important structures.  If the element has layout,
    // then the __pvChain member points to that layout.  If not and if the element is in a tree then
    // then the __pvChain member points to the ped that it is in.  Otherwise __pvChain points to the
    // document.
private:
    void* __pvChain;

public:
    CTreeNode* __pNodeFirstBranch;

    // First DWORD of bits
    ELEMENT_TAG _etag                    : 8; //  0- 7 element tag
    unsigned _fHasLookasidePtr           : 7; //  8-14 TRUE if lookaside table has pointer
    unsigned _fIsNamed                   : 1; // 15 set if element has a name or ID attribute
    unsigned _wLockFlags                 :16; // 16-31 Element lock flags for preventing recursion

    // Second DWORD of bits
    //
    // Note that the _fMark1 and _fMark2 bits are only safe to use in routines which can *guarantee*
    // their processing will not be interrupted. If there is a chance that a message loop can cause other,
    // unrelated code, to execute, these bits could get reset before the original caller is finished.
    unsigned _fHasMarkupPtr              : 1; //  0 TRUE if element has a Markup pointer
    unsigned _fHasLayoutPtr              : 1; //  1 TRUE if element has layout ptr
    unsigned _fHasPendingFilterTask      : 1; //  2 TRUE if there is a pending filter task (see EnsureView)
    unsigned _fHasPendingRecalcTask      : 1; //  3 TRUE if there is a pending recalc task (see EnsureView)
    unsigned _fLayoutAlwaysValid         : 1; //  4 TRUE if element is a site or never has layout
    unsigned _fOwnsRuns                  : 1; //  5 TRUE if element owns the runs underneath it
    unsigned _fInheritFF                 : 1; //  6 TRUE if element to inherit site and fancy format
    unsigned _fBreakOnEmpty              : 1; //  7 this element breaks a line even is it own no text
    unsigned _fUnused2                   : 4; //  8-11
    unsigned _fDefinitelyNoBorders       : 1; // 12 There are no borders on this element
    unsigned _fHasTabIndex               : 1; // 13 Has a tabindex associated with this element. Look in doc _aryTabIndexInfo
    unsigned _fHasImage                  : 1; // 14 has at least one image context
    unsigned _fResizeOnImageChange       : 1; // 15 need to force resize when image changes
    unsigned _fExplicitEndTag            : 1; // 16 element had a close tag (for P)
    unsigned _fSynthesized               : 1; // 17 FALSE (default) if user created, TRUE if synthesized
    unsigned _fUnused3                   : 1; // 18 Unused bit
    unsigned _fActsLikeButton            : 1; // 19 does element act like a push button?
    unsigned _fEditAtBrowse              : 1; // 20 to TestClassFlag(SITEDESC_EDITATBROWSE) in init
    unsigned _fSite                      : 1; // 21 element with layout by default
    unsigned _fDefault                   : 1; // 22 Is this the default "ok" button/control
    unsigned _fCancel                    : 1; // 23 Is this the default "cancel" button/control
    unsigned _fHasPendingEvent           : 1; // 24 A posted event for element is pending
    unsigned _fEventListenerPresent      : 1; // 25 Someone has asked for a connection point/ or set an event handler
    unsigned _fHasFilterCollectionPtr    : 1; // 26 FilterCollectionPtr has been added to _pAA
    unsigned _fExittreePending           : 1; // 27 There is a pending Exittree notification for this element
    unsigned _fFirstCommonAncestor       : 1; // 28 Used in GetFirstCommonAncestor - don't touch!
    unsigned _fMark1                     : 1; // 29 Random mark
    unsigned _fMark2                     : 1; // 30 Another random mark
    unsigned _fHasStyleExpressions       : 1; // 31 There are style expressions on this element

    // STATIC DATA

    //  Default property page list for elements that don't have their own.
    //  This gives them the allpage by default.
    static ACCELS s_AccelsElementDesign;
    static ACCELS s_AccelsElementRun;

    // Style methods
#include "../gen/style.hdl"
    
    // IHTMLElement methods
#define _CElement_
#include "../gen/element.hdl"

    DECLARE_TEAROFF_TABLE(IServiceProvider)

private:
    NO_COPY(CElement);
};

inline BOOL CElement::HasInLineStyles(void)
{
    CAttrArray* pStyleAA = GetInLineStyleAttrArray();
    if(pStyleAA && pStyleAA->Size())
    {
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

inline BOOL CElement::HasClassOrID(void)
{
    return GetAAclassName()||GetAAid();
}

inline CStyle* CElement::GetInLineStylePtr()
{
    CStyle* pStyleInLine = NULL;
    GetPointerAt(FindAAIndex(DISPID_INTERNAL_CSTYLEPTRCACHE,CAttrValue::AA_Internal),
        (void**)&pStyleInLine);
    return pStyleInLine;
}

inline CStyle* CElement::GetRuntimeStylePtr()
{
    CStyle* pStyle = NULL;
    GetPointerAt(FindAAIndex(DISPID_INTERNAL_CRUNTIMESTYLEPTRCACHE, CAttrValue::AA_Internal), (void**)&pStyle);
    return pStyle;
}

inline LPCTSTR CElement::GetAAid() const 
{
    LPCTSTR v;

    if(!IsNamed())
    {
        return NULL;
    }

    CAttrArray::FindString(*GetAttrArray(), &s_propdescCElementid.a, &v);
    return *(LPCTSTR*)&v;
}



BOOL SameScope(CTreeNode* pNode1, const CElement* pElement2);
BOOL SameScope(const CElement* pElement1, CTreeNode* pNode2);
BOOL SameScope(CTreeNode* pNode1, CTreeNode* pNode2);

inline BOOL DifferentScope(CTreeNode* pNode1, const CElement* pElement2)
{
    return !SameScope(pNode1, pElement2);
}
inline BOOL DifferentScope(const CElement * pElement1, CTreeNode* pNode2)
{
    return !SameScope(pElement1, pNode2);
}
inline BOOL DifferentScope(CTreeNode* pNode1, CTreeNode* pNode2)
{
    return !SameScope(pNode1, pNode2);
}

void TransformSlaveToMaster(CTreeNode** ppNode);

HRESULT StoreLineAndOffsetInfo(CBase* pBaseObj, DISPID dispid, ULONG uLine, ULONG uOffset);
HRESULT GetLineAndOffsetInfo(CAttrArray* pAA, DISPID dispid, ULONG* puLine, ULONG* puOffset);

#endif //__XINDOWS_SITE_BASE_ELEMENT_H__