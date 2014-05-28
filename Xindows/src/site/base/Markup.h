
#ifndef __XINDOWS_SITE_BASE_MARKUP_H__
#define __XINDOWS_SITE_BASE_MARKUP_H__

class CCollectionCache;
class CAutoRange;
class CSpliceRecordList;
class CSelectionRenderingServiceProvider;
struct HighlightSegment;
class CCollectionBuildContext;
class CProgSink;
struct HTMPASTEINFO;

EXTERN_C const GUID CLSID_CMarkup;

#include "../text/_text.h"

#define _hxx_
#include "../gen/markup.hdl"

// The initial size of the the CTreePos pool.  This
// has to be at least 1.
#define INITIAL_TREEPOS_POOL_SIZE	1

#define TREEDATA2SIZE (sizeof(CTreePos)+8)

#ifdef _WIN64
#define TREEDATA1SIZE (sizeof(CTreePos)+8)
#else
#define TREEDATA1SIZE (sizeof(CTreePos)+4)
#endif

#define MUS_DOMOPERATION    (0x1<<0)    // DOM operations -> text id is content

//---------------------------------------------------------------------------
//
//  Class:   CMarkupTextFragContext
//
//---------------------------------------------------------------------------
struct MarkupTextFrag
{
    CTreePos*   _ptpTextFrag;
    TCHAR*      _pchTextFrag;
};

class CMarkupTextFragContext
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    ~CMarkupTextFragContext();

    DECLARE_CDataAry(CAryMarkupTextFrag, MarkupTextFrag)
    CAryMarkupTextFrag _aryMarkupTextFrag;

    long    FindTextFragAtCp(long cp, BOOL* pfFragFound);
    HRESULT AddTextFrag(CTreePos* ptpTextFrag, TCHAR* pchTextFrag, ULONG cchTextFrag, long iTextFrag);
    HRESULT RemoveTextFrag(long iTextFrag, CMarkup* pMarkup);
};


//---------------------------------------------------------------------------
//
//  Class:   CMarkupTopElemCache
//
//---------------------------------------------------------------------------
class CMarkupTopElemCache
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CElement*       __pElementClientCached;
};

class CMarkup : public CBase
{
	friend class CTxtPtr;
	friend class CTreePos;
	friend class CMarkupPointer;

	DECLARE_CLASS_TYPES(CMarkup, CBase)

public:
    DECLARE_TEAROFF_TABLE(ISelectionRenderingServices)
    DECLARE_TEAROFF_TABLE(IMarkupContainer)

    DECLARE_MEMCLEAR_NEW_DELETE()

	CMarkup(CDocument* pDoc, CElement* pElementMaster=NULL);
	~CMarkup();

	HRESULT Init(CRootElement* pElementRoot);

	HRESULT CreateInitialMarkup(CRootElement* pElementRoot);

	HRESULT UnloadContents(BOOL fForPassivate=FALSE);

	HRESULT DestroySplayTree(BOOL fReinit);

	void Passivate();

	HRESULT CreateElement(
		ELEMENT_TAG	etag,
		CElement**	ppElementNew);

#define _CMarkup_
#include "../gen/markup.hdl"

    // Document props\methods:
    NV_DECLARE_TEAROFF_METHOD(get_Script, GET_Script, (IDispatch** p));
    NV_DECLARE_TEAROFF_METHOD(get_all, GET_all, (IHTMLElementCollection** p));
    NV_DECLARE_TEAROFF_METHOD(get_body, GET_body, (IHTMLElement** p));
    NV_DECLARE_TEAROFF_METHOD(get_activeElement, GET_activeElement, (IHTMLElement** p));
   NV_DECLARE_TEAROFF_METHOD(put_title, PUT_title, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_title, GET_title, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(get_readyState, GET_readyState, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(put_expando, PUT_expando, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_expando, GET_expando, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(get_parentWindow, GET_parentWindow, (IHTMLWindow2** p));
    NV_DECLARE_TEAROFF_METHOD(get_nameProp, GET_nameProp, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandSupported, querycommandsupported, (BSTR cmdID,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandEnabled, querycommandenabled, (BSTR cmdID,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandState, querycommandstate, (BSTR cmdID,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandIndeterm, querycommandindeterm, (BSTR cmdID,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandText, querycommandtext, (BSTR cmdID,BSTR* pcmdText));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, queryCommandValue, querycommandvalue, (BSTR cmdID,VARIANT* pcmdValue));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, execCommand, execcommand, (BSTR cmdID,VARIANT_BOOL showUI,VARIANT value,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, execCommandShowHelp, execcommandshowhelp, (BSTR cmdID,VARIANT_BOOL* pfRet));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, createElement, createelement, (BSTR eTag,IHTMLElement** newElem));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, elementFromPoint, elementfrompoint, (long x,long y,IHTMLElement** elementHit));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, toString, tostring, (BSTR* String));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, releaseCapture, releasecapture, ());
    NV_DECLARE_TEAROFF_METHOD(get_documentElement, GET_documentelement, (IHTMLElement**pRootElem));
    NV_DECLARE_TEAROFF_METHOD(get_uniqueID, GET_uniqueID, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, attachEvent, attachevent, (BSTR event,IDispatch* pDisp,VARIANT_BOOL* pfResult));
    NV_DECLARE_TEAROFF_METHOD_(HRESULT, detachEvent, detachevent, (BSTR event,IDispatch* pDisp));
    NV_DECLARE_TEAROFF_METHOD(get_bgColor, GET_bgColor, (VARIANT* p));
    NV_DECLARE_TEAROFF_METHOD(put_bgColor, PUT_bgColor, (VARIANT p));
    NV_DECLARE_TEAROFF_METHOD(get_fgColor, GET_fgColor, (VARIANT* p));
    NV_DECLARE_TEAROFF_METHOD(put_fgColor, PUT_fgColor, (VARIANT p));
    NV_DECLARE_TEAROFF_METHOD(get_linkColor, GET_linkColor, (VARIANT* p));
    NV_DECLARE_TEAROFF_METHOD(put_linkColor, PUT_linkColor, (VARIANT p));
    NV_DECLARE_TEAROFF_METHOD(get_alinkColor, GET_alinkColor, (VARIANT* p));
    NV_DECLARE_TEAROFF_METHOD(put_alinkColor, PUT_alinkColor, (VARIANT p));
    NV_DECLARE_TEAROFF_METHOD(get_vlinkColor, GET_vlinkColor, (VARIANT* p));
    NV_DECLARE_TEAROFF_METHOD(put_vlinkColor, PUT_vlinkColor, (VARIANT p));
    NV_DECLARE_TEAROFF_METHOD(get_parentDocument, GET_parentDocument, (IHTMLDocument2** p));
    NV_DECLARE_TEAROFF_METHOD(put_enableDownload, PUT_enableDownload, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_enableDownload, GET_enableDownload, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(put_baseUrl, PUT_baseUrl, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_baseUrl, GET_baseUrl, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(put_inheritStyleSheets, PUT_inheritStyleSheets, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_inheritStyleSheets, GET_inheritStyleSheets, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(getElementsByName, getelementsbyname, (BSTR v, IHTMLElementCollection** p));
    NV_DECLARE_TEAROFF_METHOD(getElementsByTagName, getelementsbytagname, (BSTR v, IHTMLElementCollection** p));
    NV_DECLARE_TEAROFF_METHOD(getElementById, getelementbyid, (BSTR v, IHTMLElement** p));

	// IServiceProvider
	STDMETHOD(QueryService)(REFGUID guidService, REFIID riid, void** ppv);

	// Data Access
	CDocument*			    Doc() { return _pDoc; }
	CRootElement*			Root() { return _pElementRoot; }

	BOOL					HasCollectionCache() { return HasLookasidePtr(LOOKASIDE_COLLECTIONCACHE); }
	CCollectionCache*		CollectionCache() { return (CCollectionCache*)GetLookasidePtr(LOOKASIDE_COLLECTIONCACHE); }
	CCollectionCache*		DelCollectionCache() { return (CCollectionCache*)DelLookasidePtr(LOOKASIDE_COLLECTIONCACHE); }
	HRESULT					SetCollectionCache(CCollectionCache* pCollectionCache) { return SetLookasidePtr(LOOKASIDE_COLLECTIONCACHE, pCollectionCache); }

	BOOL					HasParentMarkup() { return HasLookasidePtr(LOOKASIDE_PARENTMARKUP); }
	CMarkup*				ParentMarkup() { return (CMarkup*)GetLookasidePtr(LOOKASIDE_PARENTMARKUP); }
	CMarkup*				DelParentMarkup() { return (CMarkup*)DelLookasidePtr(LOOKASIDE_PARENTMARKUP); }
	HRESULT					SetParentMarkup(CMarkup* pMarkup) { return SetLookasidePtr(LOOKASIDE_PARENTMARKUP, pMarkup); }

	CElement*				Master() { return _pElementMaster; }
	void					ClearMaster() { _pElementMaster = NULL; }

    CProgSink*              GetProgSinkC();
    IProgSink*              GetProgSink();

	CElement*				FirstElement();

	// Remove 'Assert(Master())', once the root element is gone and these functions
	// become meaningful for non-slave markups as well.
	long				GetContentFirstCp() { Assert(Master()); return 2; }
	long				GetContentLastCp() { Assert(Master()); return GetTextLength()-2; }
	long				GetContentTextLength() { Assert(Master()); return GetTextLength()-4; }
	void				GetContentTreeExtent(CTreePos** pptpStart, CTreePos** pptpEnd);

	BOOL				GetAutoWordSel() const { return TRUE; }

	BOOL				GetOverstrike() const { return _fOverstrike; }
	void				SetOverstrike(BOOL fSet) { _fOverstrike = (fSet) ? 1 : 0;}

	BOOL				GetLoaded() const { return _fLoaded;     }
	void				SetLoaded(BOOL fLoaded) { _fLoaded = fLoaded;  }

	void				SetStreaming(BOOL flag) { _fStreaming = (flag) ? 1 : 0; }
	BOOL				IsStreaming() const { return _fStreaming; } 

	// BUGBUG: need to figure this function out...
	BOOL				IsEditable(BOOL fCheckContainerOnly=FALSE) const;

	// Top element cache
	void				EnsureTopElems();
	CElement*			GetElementTop();
	CElement*			GetElementClient() { EnsureTopElems(); return HasTopElemCache()?GetTopElemCache()->__pElementClientCached:NULL; }

	// Other helpers
	BOOL			IsPrimaryMarkup() { return this==Doc()->_pPrimaryMarkup; }
	long			GetTextLength() { return _TxtArray._cchText; }

	// GetRunOwner is an abomination that should be eliminated -- or at least moved off of the markup
	CLayout*		GetRunOwner(CTreeNode* pNode, CLayout* pLayoutParent=NULL);
	CTreeNode*		GetRunOwnerBranch(CTreeNode*, CLayout* pLayoutParent=NULL);

	// Markup manipulation functions
	void CompactStory() { _TxtArray.ShrinkBlocks(); }

	HRESULT RemoveElementInternal(
		CElement*	pElementRemove,
		DWORD		dwFlags=NULL);

	HRESULT InsertElementInternal(
		CElement*		pElementInsertThis,
		CTreePosGap*	ptpgBegin,
		CTreePosGap*	ptpgEnd,
		DWORD			dwFlags=NULL);

	HRESULT SpliceTreeInternal(
		CTreePosGap*	ptpgStartSource,
		CTreePosGap*	ptpgEndSource,
		CMarkup*		pMarkupTarget =NULL,
		CTreePosGap*	ptpgTarget=NULL,
		BOOL			fRemove=TRUE,
		DWORD			dwFlags=NULL);

	HRESULT InsertTextInternal(
		CTreePos*		ptpAfterInsert,
		const TCHAR*	pch,
		long			cch,
		DWORD			dwFlags=NULL);

	HRESULT FastElemTextSet(
		CElement*       pElem,
		const TCHAR*    psz,
		int             cch,
		BOOL            fAsciiOnly);

    // Undo only operations
    HRESULT UndoRemoveSplice(
        CMarkupPointer*     pmpBegin,
        CSpliceRecordList*  paryRegion,
        long                cchRemove,
        TCHAR*              pchRemove,
        long                cchNodeReinsert,
        DWORD               dwFlags);

	// Markup operation helpers
	HRESULT ReparentDirectChildren(
		CTreeNode*		pNodeParentNew,
		CTreePosGap*	ptpgStart=NULL,
		CTreePosGap*	ptpgEnd=NULL);

	HRESULT CreateInclusion(
		CTreeNode*		pNodeStop,
		CTreePosGap*	ptpgLocation,
		CTreePosGap*	ptpgInclusion,
		long*			pcchNeeded=NULL,
		CTreeNode*		pNodeAboveLocation=NULL,
		BOOL			fFullReparent=TRUE,
		CTreeNode**		ppNodeLastAdded=NULL);

	HRESULT CloseInclusion(CTreePosGap* ptpgMiddle, long* pcchRemove=NULL);

	HRESULT RangeAffected(CTreePos* ptpStart, CTreePos* ptpFinish);
	HRESULT ClearCaches(CTreePos* ptpStart, CTreePos* ptpFinish);
	HRESULT ClearRunCaches(DWORD dwFlags, CElement* pElement);

	// Markup pointers
	HRESULT EmbedPointers() { return _pmpFirst?DoEmbedPointers():S_OK; }

	HRESULT DoEmbedPointers();

	BOOL HasUnembeddedPointers() { return !!_pmpFirst; }

	// TextID manipulation

	// Give a unique text id to every text chunk in
	// the range given
	HRESULT SetTextID(CTreePosGap* ptpgStart, CTreePosGap* ptpgEnd, long* plNewTextID);

	// Get text ID for text to right
	// -1 --> no text to right
	// 0  --> no assigned TextID
	long GetTextID(CTreePosGap* ptpg);

	// Starting with ptpgStart, look for
	// the extent of lTextID
	HRESULT FindTextID(long lTextID, CTreePosGap* ptpgStart, CTreePosGap* ptpgEnd);

	// If text to the left of ptpLeft has
	// the same ID as text to the right of
	// ptpRight, give the right fragment a
	// new ID
	void SplitTextID(CTreePos* ptpLeft, CTreePos* ptpRight);

	// Splay Tree Primitives
	CTreePos* NewTextPos(long cch, SCRIPT_ID sid=sidAsciiLatin, long lTextID=0);
	CTreePos* NewPointerPos(CMarkupPointer* pPointer, BOOL fRight, BOOL fStick);

	typedef CTreePos SUBLIST;
	HRESULT Append(CTreePos* ptp);
	HRESULT Insert(CTreePos* ptpNew, CTreePosGap* ptpgInsert);
	HRESULT Insert(CTreePos* ptpNew, CTreePos* ptpInsert, BOOL fBefore);
	HRESULT Move(CTreePos* ptpMove, CTreePosGap* ptpgDest);
	HRESULT Move(CTreePos* ptpMove, CTreePos* ptpDest, BOOL fBefore);
	HRESULT Remove(CTreePos* ptpStart, CTreePos* ptpFinish);
	HRESULT Remove(CTreePos* ptp) { return Remove(ptp, ptp); }
	HRESULT Split(CTreePos* ptpSplit, long cchLeft, SCRIPT_ID sidNew=sidNil);
	HRESULT Join(CTreePos* ptpJoin);
	HRESULT ReplaceTreePos(CTreePos* ptpOld, CTreePos* ptpNew); 
	HRESULT MergeText(CTreePos* ptpMerge);
	HRESULT SetTextPosID(CTreePos** pptpText, long lTextID);
	HRESULT RemovePointerPos(CTreePos* ptp, CTreePos** pptpUpdate, long* pichUpdate);

	HRESULT SpliceOut(CTreePos* ptpStart, CTreePos* ptpFinish, SUBLIST* pSublistSplice);
	HRESULT SpliceIn(SUBLIST* pSublistSplice, CTreePos* ptpSplice);
	HRESULT InsertPosChain(CTreePos* ptpChainHead, CTreePos* ptpChainTail, CTreePos* ptpInsertBefore);

	// splay tree search routines
	CTreePos* FirstTreePos() const;
	CTreePos* LastTreePos() const;
	CTreePos* TreePosAtCp(long cp, long* pcchOffset) const;
	CTreePos* TreePosAtSourceIndex(long iSourceIndex);

	// General splay information
	long NumElems() const { return _tpRoot.GetElemLeft(); }
	long Cch() const { return _tpRoot._cchLeft; }
	long CchInRange(CTreePos* ptpFirst, CTreePos* ptpLast);

	// Force a splay
	void FocusAt(CTreePos* ptp) { ptp->Splay(); }

	// splay tree primitive helpers
protected:
	CTreePos*	AllocData1Pos();
	CTreePos*	AllocData2Pos();
	void		FreeTreePos(CTreePos* ptp);
	void		ReleaseTreePos(CTreePos* ptp, BOOL fLastRelease=FALSE);
	BOOL		ShouldSplay(long cDepth) const;
	HRESULT		MergeTextHelper(CTreePos* ptpMerge);

public:
	// Text Story
	LONG		GetTextLength() const { return _TxtArray._cchText; }

	// Notifications
	void		Notify(CNotification* pnf);
	void		Notify(CNotification& rnf) { Notify(&rnf); }

	// notification helpers
protected:
	void SendNotification(CNotification* pnf, CDataAry<CNotification>* paryNotification);
	BOOL ElementWantsNotification(CElement* pElement, CNotification* pnf);

	void NotifyElement(CElement* pElement, CNotification* pnf);
	void NotifyAncestors(CNotification* pnf);
	void NotifyDescendents(CNotification* pnf);
	void NotifyTreeLevel(CNotification* pnf);

public:
	// Branch searching functions - I'm sure some of these aren't needed
	//
	// Note: All of these (except InStory versions) implicitly stop searching at a FlowLayout
	CTreeNode*	FindMyListContainer(CTreeNode* pNodeStartHere);
	CTreeNode*	SearchBranchForChildOfScope(CTreeNode* pNodeStartHere, CElement* pElementFindChildOfMyScope);
	CTreeNode*	SearchBranchForChildOfScopeInStory(CTreeNode* pNodeStartHere, CElement* pElementFindChildOfMyScope);
	CTreeNode*	SearchBranchForScopeInStory(CTreeNode* pNodeStartHere, CElement* pElementFindMyScope);
	CTreeNode*	SearchBranchForScope(CTreeNode* pNodeStartHere, CElement* pElementFindMyScope);
	CTreeNode*	SearchBranchForNode(CTreeNode* pNodeStartHere, CTreeNode* brFindMe);
	CTreeNode*	SearchBranchForNodeInStory(CTreeNode* pNodeStartHere, CTreeNode* brFindMe);
	CTreeNode*	SearchBranchForTag(CTreeNode* pNodeStartHere, ELEMENT_TAG);
	CTreeNode*	SearchBranchForTagInStory(CTreeNode* pNodeStartHere, ELEMENT_TAG);
	CTreeNode*	SearchBranchForBlockElement(CTreeNode* pNodeStartHere, CFlowLayout* pFLContext=NULL);
	CTreeNode*	SearchBranchForNonBlockElement(CTreeNode* pNodeStartHere, CFlowLayout* pFLContext=NULL);
	CTreeNode*	SearchBranchForAnchor(CTreeNode* pNodeStartHere);
	CTreeNode*	SearchBranchForCriteria(CTreeNode* pNodeStartHere, BOOL (*pfnSearchCriteria)(CTreeNode*));
	CTreeNode*	SearchBranchForCriteriaInStory(CTreeNode* pNodeStartHere, BOOL (*pfnSearchCriteria)(CTreeNode*));

    void    EnsureFormats();

    // Markup TextFrag services
    // Markup TextFrags are used to store arbitrary string data in the markup.  Mostly
    // this is used to persist and edit conditional comment tags
    CMarkupTextFragContext*     EnsureTextFragContext();

    BOOL                        HasTextFragContext()    { return HasLookasidePtr(LOOKASIDE_TEXTFRAGCONTEXT); }
    CMarkupTextFragContext*     GetTextFragContext()    { return (CMarkupTextFragContext*)GetLookasidePtr(LOOKASIDE_TEXTFRAGCONTEXT); }
    HRESULT                     SetTextFragContext(CMarkupTextFragContext* ptfc) { return SetLookasidePtr(LOOKASIDE_TEXTFRAGCONTEXT, ptfc); }
    CMarkupTextFragContext*     DelTextFragContext()    { return (CMarkupTextFragContext*)DelLookasidePtr(LOOKASIDE_TEXTFRAGCONTEXT); }

    // Stores the cached values for the client element
	CMarkupTopElemCache*		EnsureTopElemCache();

	BOOL						HasTopElemCache() { return HasLookasidePtr(LOOKASIDE_TOPELEMCACHE); }
	CMarkupTopElemCache*		GetTopElemCache() { return (CMarkupTopElemCache*)GetLookasidePtr(LOOKASIDE_TOPELEMCACHE); }
	HRESULT						SetTopElemCache(CMarkupTopElemCache* ptec) { return SetLookasidePtr(LOOKASIDE_TOPELEMCACHE, ptec); }
	CMarkupTopElemCache*		DelTopElemCache() { return (CMarkupTopElemCache*)DelLookasidePtr(LOOKASIDE_TOPELEMCACHE); }

    // Selection Methods
    VOID GetSelectionChunksForLayout(CFlowLayout* pFlowLayout, CPtrAry<HighlightSegment*>* paryHighlight, int* piCpMin, int* piCpMax);
    HRESULT EnsureSelRenSvc();
    VOID HideSelection();
    VOID ShowSelection();
    VOID InvalidateSelection(BOOL fFireOM);


    // Collections support
    enum
    {
        ELEMENT_COLLECTION = 0,
        FORMS_COLLECTION,
        ANCHORS_COLLECTION,
        LINKS_COLLECTION,
        IMAGES_COLLECTION,
        APPLETS_COLLECTION,
        SCRIPTS_COLLECTION,
        WINDOW_COLLECTION,
        EMBEDS_COLLECTION,
        REGION_COLLECTION,
        LABEL_COLLECTION,
        NAVDOCUMENT_COLLECTION,
        FRAMES_COLLECTION,
        NUM_DOCUMENT_COLLECTIONS
    };
    // DISPID range for FRAMES_COLLECTION
    enum
    {
        FRAME_COLLECTION_MIN_DISPID = ((DISPID_COLLECTION_MIN+DISPID_COLLECTION_MAX)*2)/3+1,
        FRAME_COLLECTION_MAX_DISPID = DISPID_COLLECTION_MAX
    };

    HRESULT EnsureCollectionCache(long lCollectionIndex);
    HRESULT AddToCollections(CElement* pElement, CCollectionBuildContext* pWalk);

    HRESULT InitCollections(void);

    NV_DECLARE_ENSURE_METHOD(EnsureCollections, ensurecollections, (long lIndex, long* plCollectionVersion));
    HRESULT GetCollection(int iIndex, IHTMLElementCollection** ppdisp);
    HRESULT GetElementByNameOrID(LPTSTR szName, CElement** ppElement);
    HRESULT GetDispByNameOrID(LPTSTR szName, IDispatch** ppDisp, BOOL fAlwaysCollection=FALSE);

    CTreeNode* InFormCollection(CTreeNode* pNode);

	// Lookaside pointers
	enum
	{
		LOOKASIDE_COLLECTIONCACHE	= 0,
		LOOKASIDE_PARENTMARKUP		= 1,
		LOOKASIDE_BEHAVIORCONTEXT	= 2,
		LOOKASIDE_TEXTFRAGCONTEXT	= 3,
		LOOKASIDE_TOPELEMCACHE		= 4,
		LOOKASIDE_STYLESHEETS		= 5,
		LOOKASIDE_TEXTRANGE			= 6,
		LOOKASIDE_MARKUP_NUMBER		= 7
	};

	BOOL		HasLookasidePtr(int iPtr) { return (_fHasLookasidePtr&(1<<iPtr)); }
	void*		GetLookasidePtr(int iPtr);
	HRESULT		SetLookasidePtr(int iPtr, void* pv);
	void*		DelLookasidePtr(int iPtr);

	void        ClearLookasidePtrs();

	// IUnknown
	DECLARE_PLAIN_IUNKNOWN(CMarkup);
    STDMETHOD(PrivateQueryInterface)(REFIID, void**);

    // ISegmentList
    NV_DECLARE_TEAROFF_METHOD(MovePointersToSegment, movepointerstosegment, (
        int iSegmentIndex, IMarkupPointer* pStart, IMarkupPointer* pEnd));
    NV_DECLARE_TEAROFF_METHOD(GetSegmentCount, getsegmentcount, (
        int* piSegmentCount, SELECTION_TYPE* peType));

    // ISelectionRenderingServices
    NV_DECLARE_TEAROFF_METHOD(AddSegment, addsegment, (
        IMarkupPointer* pStart, IMarkupPointer* pEnd, HIGHLIGHT_TYPE HighlightType, int* iSegmentIndex));
    NV_DECLARE_TEAROFF_METHOD(AddElementSegment, addelementsegment, (
        IHTMLElement* pElement, int* iSegmentIndex));
    NV_DECLARE_TEAROFF_METHOD(MoveSegmentToPointers, movesegmenttopointers, (
        int iSegmentIndex, IMarkupPointer* pStart, IMarkupPointer* pEnd, HIGHLIGHT_TYPE HighlightType));
    NV_DECLARE_TEAROFF_METHOD(GetElementSegment, getelementsegment, (
        int iSegmentIndex, IHTMLElement** ppElement));
    NV_DECLARE_TEAROFF_METHOD(SetElementSegment, setelementsegment, (
        int iSegmentIndex, IHTMLElement* pElement));
    NV_DECLARE_TEAROFF_METHOD(ClearSegment, clearsegment, (
        int iSegmentIndex, BOOL fInvalidate));
    NV_DECLARE_TEAROFF_METHOD(ClearSegments, clearsegments, (BOOL fInvalidate));
    NV_DECLARE_TEAROFF_METHOD(ClearElementSegments, clearelementsegments, ());

    // IMarkupContainer methods
    NV_DECLARE_TEAROFF_METHOD(OwningDoc, owningdoc, (
        IHTMLDocument2** ppDoc));

	// Ref counting helpers
	static void ReplacePtr(CMarkup** pplhs, CMarkup* prhs);
	static void SetPtr(CMarkup** pplhs, CMarkup* prhs);
	static void ClearPtr(CMarkup** pplhs);
	static void StealPtrSet(CMarkup** pplhs, CMarkup* prhs);
	static void StealPtrReplace(CMarkup** pplhs, CMarkup* prhs);
	static void ReleasePtr(CMarkup* pMarkup);

	// CMarkup::CLock
	class CLock
	{
	public:
		CLock(CMarkup* pMarkup);
		~CLock();

	private:
		CMarkup* _pMarkup;
	};

protected:
    static const CLASSDESC s_classdesc;
    virtual const CBase::CLASSDESC* GetClassDesc() const { return &s_classdesc; }

	// Data
public:
    CProgSink*  _pProgSink;

	CDocument*  _pDoc;

	// The following are similar to the CDoc's equivalent, but pertain to this
	// markup alone.
	long GetMarkupTreeVersion() { return __lMarkupTreeVersion; }
	long GetMarkupContentsVersion() { return __lMarkupContentsVersion; }

	// Do NOT modify these version numbers unless the document structure
	// or content is being modified.
	//
	// In particular, incrementing these to get a cache to rebuild is
	// BAD because it causes all sorts of other stuff to rebuilt.
	long __lMarkupTreeVersion;		// Element structure
	long __lMarkupContentsVersion;	// Any content

	void UpdateMarkupTreeVersion()
	{
		CDocument* pDoc = Doc();

		__lMarkupTreeVersion++;
		__lMarkupContentsVersion++;

		pDoc->__lDocTreeVersion++;
		pDoc->__lDocContentsVersion++;
	}

	void UpdateMarkupContentsVersion()
	{
		__lMarkupContentsVersion++;
		Doc()->UpdateDocContentsVersion();
	}

private:
	CRootElement*	_pElementRoot;
	CElement*		_pElementMaster;

	// Story data
	CTxtArray		_TxtArray;

	long _lTopElemsVersion;

    // Selection state
    CSelectionRenderingServiceProvider* _pSelRenSvcProvider; // Object to Delegate to.

	// Notification data
	DECLARE_CDataAry(CAryANotify, CNotification)
	CAryANotify _aryANotification;

private:
	// This is the list of pointers positioned in this markup
	// which do not have an embedding.
	CMarkupPointer*	_pmpFirst;

	// Splay Tree Data
	CTreePos		_tpRoot;		// dummy root node
	CTreePos*		_ptpFirst;		// cached first (leftmost) node
	void*			_pvPool;		// list of pool blocks (so they can be freed)
	CTreeDataPos*	_ptdpFree;		// head of free list
	BYTE			_abPoolInitial[sizeof(void*)+TREEDATA1SIZE*INITIAL_TREEPOS_POOL_SIZE]; // The initial pool of TreePos objects

public:
	struct
	{
		DWORD   _fOverstrike       : 1; // 1 Overstrike mode vs insert mode
		DWORD   _fLoaded           : 1; // 3 Is the markup completely parsed
		DWORD   _fNoUndoInfo       : 1; // 4 Don't store any undo info for this markup
		DWORD   _fIncrementalAlloc : 1; // 5 The text array should slow start
		DWORD   _fStreaming        : 1; // 6 True during parsing
		DWORD   _fUnstable         : 1; // 7 the tree is unstable because of the tree services/DOM operations 
										// were performed on the tree and nobody call to validate the tree
		DWORD   _fInSendAncestor   : 1; // 8 Notification - We're sending something to ancestors
		DWORD   _fUnused1          : 1; // 9 
		DWORD   _fEnableDownload   : 1; // 10 Allows content to be downloaded in this markup
		DWORD   _fPad              : 5; // 11-15 Padding to align lookaside flags on byte
		DWORD   _fHasLookasidePtr  : 8; // 16-23 Lookaside flags
		DWORD   _fUnused2          : 8; // 24-31
	};

    // Style sheets moved from CDocument
    HRESULT EnsureStyleSheets();
    HRESULT ApplyStyleSheets(
        CStyleInfo*     pStyleInfo,
        ApplyPassType   passType=APPLY_All,
        BOOL*           pfContainsImportant=NULL);

    BOOL HasStyleSheets()
    {   
        return FALSE; 
    }

private:
	NO_COPY(CMarkup);
};

#endif //__XINDOWS_SITE_BASE_MARKUP_H__