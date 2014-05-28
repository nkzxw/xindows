
#ifndef __XINDOWS_SITE_BASE_DOC_H__
#define __XINDOWS_SITE_BASE_DOC_H__

#include "../include/mshtmhst.h"
#include "../gen/internal.h"
#include "../gen/mshtmext.h"

#define _hxx_
#include "../gen/document.hdl"

#include "../view/View.h"
#include <ocmm.h>
#include "EditRouter.h"

#define UNICODE_SIGNATURE_LITTLEENDIAN  0xfeff
#define UNICODE_SIGNATURE_BIGENDIAN     0xfffe

#define NATIVE_UNICODE_SIGNATURE        UNICODE_SIGNATURE_LITTLEENDIAN
#define NONNATIVE_UNICODE_SIGNATURE     UNICODE_SIGNATURE_BIGENDIAN


struct IMGANIMSTATE;
struct DWNLOADINFO;
struct MIMEINFO;
class CImgCtx;
class CXindowsEditor;
class CDwnCtx;
class CDwnBindData;
class CDwnDoc;
class CDragDropSrcInfo;
class CDragDropTargetInfo;
class CDragStartInfo;
class CSelDragDropSrcInfo;
class CProgSink;

class CCaret;
class CCharFormat;
class CNotification;
class CMessage;
class CElement;
class CBodyElement;
class CRootElement;
class CMarkup;
class CMarkupPointer;
class CTreeNode;
class CColorInfo;
class CFlowLayout;
class CObjectElement;

//+---------------------------------------------------------------------------
//
// defines: NEED3DBORDER
//
// Synopsis: See CDocument::CheckDoc3DBorder for description
//
//----------------------------------------------------------------------------
#define NEED3DBORDER_NO     (0x00)
#define NEED3DBORDER_TOP    (0x01)
#define NEED3DBORDER_LEFT   (0x02)
#define NEED3DBORDER_BOTTOM (0x04)
#define NEED3DBORDER_RIGHT  (0x08)

// Timer IDs
#define TIMER_DEFERUPDATEUI 0x1000
#define TIMER_ID_MOUSE_EXIT 0x1001
#define TIMER_IMG_ANIM      0x1002


HRESULT CTQueryStatus(IUnknown* pUnk, const GUID* pguidCmdGroup,
                      ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pcmdtext);
HRESULT CTExec(IUnknown* pUnk, const GUID* pguidCmdGroup, DWORD nCmdID,
               DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut);


//--------------------------------------------------------------
//
//  Global variables
//
//--------------------------------------------------------------
extern SIZE     g_sizeDragMin;
extern int      g_iDragScrollDelay;
extern SIZE     g_sizeDragScrollInset;
extern int      g_iDragDelay;
extern int      g_iDragScrollInterval;


typedef HRESULT (STDMETHODCALLTYPE CElement::*PFN_VISIT)(DWORD_PTR dw);

#define VISIT_METHOD(kls, FN, fn) \
    (PFN_VISIT)&kls::FN

#define NV_DECLARE_VISIT_METHOD(FN, fn, args) \
    HRESULT STDMETHODCALLTYPE FN args;

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_VOID_MOUSECAPTURE)(CMessage* pMessage);

#define MOUSECAPTURE_METHOD(klass, fn, FN) \
    (PFN_VOID_MOUSECAPTURE)&klass::fn

#define NV_DECLARE_MOUSECAPTURE_METHOD(fn, FN, args) \
    HRESULT STDMETHODCALLTYPE fn args

#define DECLARE_MOUSECAPTURE_METHOD(fn, FN, args) \
    virtual HRESULT STDMETHODCALLTYPE fn args

//+--------------------------------------------------------------------
//
//  Enumeration:    INVAL_FLAGS
//
//  Synopsis:       See CDocument::Invalidate for description.
//
//---------------------------------------------------------------------
enum INVAL_FLAGS
{
	INVAL_CHILDWINDOWS	= 1,	// Invalidate all child windows
	INVAL_GRABHANDLES	= 2		// Invalidate grab handles if any
};

struct RADIOGRPNAME
{
	LPCTSTR			lpstrName;
	RADIOGRPNAME*	_pNext;
};


// Private window messages
#define WM_SERVER_FIRST                 (WM_APP         +1)
#define WM_SERVER_LAST                  (WM_SERVER_FIRST+1)

#define WM_FORM_FIRST                   (WM_SERVER_LAST +1)
#define WM_DEFERZORDER                  (WM_FORM_FIRST  +0)
#define WM_MOUSEOVER                    (WM_FORM_FIRST  +1)
#define WM_ACTIVEMOVIE                  (WM_FORM_FIRST  +2)
#define WM_PRINTSTATUS                  (WM_FORM_FIRST  +3)
#define WM_DEFERBLUR                    (WM_FORM_FIRST  +4)
#define WM_DEFERFOCUS                   (WM_FORM_FIRST  +5)
// Add new WM_USER messsages here and update WM_FORM_LAST.
#define WM_FORM_LAST                    (WM_FORM_FIRST  +5)

#define MSGNAME_WM_XINDOWS_GETOBJECT    _T("WM_XINDOWS_GETOBJECT")
#define UNIQUE_NAME_PREFIX              _T("ms__id")


//+-----------------------------------------------------------------
//
//  Flag values for CServer::CLock
//
//------------------------------------------------------------------
enum FORMLOCK_FLAG
{
    FORMLOCK_ARYSITE        = (SERVERLOCK_LAST << 1),   // don't allow any mods to site lists
    FORMLOCK_SITES          = (FORMLOCK_ARYSITE << 1),  // don't delete sites
    FORMLOCK_CURRENT        = (FORMLOCK_SITES << 1),    // don't allow current change
    FORMLOCK_LOADING        = (FORMLOCK_CURRENT << 1),  // don't allow loading
    FORMLOCK_UNLOADING      = (FORMLOCK_LOADING << 1),  // unloading now
    FORMLOCK_QSEXECCMD      = (FORMLOCK_UNLOADING << 1),// In the middle of queryCommand/execCommand
    FORMLOCK_FILTER         = (FORMLOCK_QSEXECCMD << 1),// Doing filter hookup
    FORMLOCK_FLAG_Last_Enum = 0
};

//+---------------------------------------------------------------------------
//
//  Class:      CDefaultElement
//
//  Purpose:    There are assumptions made in CDocument that there
//              always be a site available to point to.  The
//              default site is used for these purposes.
//
//              The root site used to erroneously serve this
//              purpose.
//
//----------------------------------------------------------------------------
class CDefaultElement : public CElement
{
	typedef CElement super;
public:
    DECLARE_MEMCLEAR_NEW_DELETE();

	CDefaultElement(CDocument* pDoc);

    DECLARE_CLASSDESC_MEMBERS;
};



// CMessage helpers
void CMessageToSelectionMessage(const CMessage* pMessage, SelectionMessage* pSelMessage);
void SelectionMessageToCMessage(const SelectionMessage* pSelMessage, CMessage* pMessage);

class CDocument : public CServer, public IHTMLDocument2
{
	DECLARE_CLASS_TYPES(CDocument, CServer)

public:
    DECLARE_MEMCLEAR_NEW_DELETE();

    DECLARE_TEAROFF_TABLE(IServiceProvider)
    DECLARE_TEAROFF_TABLE(IMarkupServices)
    DECLARE_TEAROFF_TABLE(IHTMLViewServices)

    // baseimplementation overrides
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
    NV_DECLARE_TEAROFF_METHOD(put_dir, PUT_dir, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_dir, GET_dir, (BSTR* p));

    // for faulting (JIT) in USP10
    NV_DECLARE_ONCALL_METHOD(FaultInUSP, faultinusp, (DWORD_PTR));
    NV_DECLARE_ONCALL_METHOD(FaultInJG, faultinjg, (DWORD_PTR));

	CDocument();
	virtual ~CDocument();
	
    DECLARE_PLAIN_IUNKNOWN(CDocument)
	DECLARE_PRIVATE_QI_FUNCS(CServer)

	virtual HRESULT Init();
    void SetLoadfFromPrefs();

    // IOleInPlaceObject methods
    HRESULT STDMETHODCALLTYPE SetObjectRects(LPCRECT prcPos, LPCRECT prcClip);

    // IOleInPlaceObjectWindowless methods
    HRESULT STDMETHODCALLTYPE OnWindowMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult);

    // IMsoCommandTarget methods
    struct CTQueryStatusArg
    {
        ULONG       cCmds;
        MSOCMD*     rgCmds;
        MSOCMDTEXT* pcmdtext;
    };

    struct CTExecArg
    {
        DWORD       nCmdID;
        DWORD       nCmdexecopt;
        VARIANTARG* pvarargIn;
        VARIANTARG* pvarargOut;
    };

    struct CTArg
    {
        BOOL        fQueryStatus;
        GUID*       pguidCmdGroup;
        union
        {
            CTQueryStatusArg*   pqsArg;
            CTExecArg*          pexecArg;
        };
    };

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
    HRESULT RouteCTElement(CElement* pElement, CTArg* parg);

    HRESULT STDMETHODCALLTYPE OnAmbientPropertyChange(DISPID dispid);

    // IPersistStream methods
    NV_STDMETHOD(InitNew)(void);

	// IDropTarget methods
	STDMETHOD(DragEnter)(
		IDataObject*	pIDataSource,
		DWORD			grfKeyState,
		POINTL			pt,
		DWORD*			pdwEffect);
	STDMETHOD(DragOver)(
		DWORD			grfKeyState,
		POINTL			pt,
		DWORD*			pdwEffect);
	STDMETHOD(DragLeave)(BOOL fDrop);
	STDMETHOD(Drop)(
		IDataObject*	pIDataSource,
		DWORD			grfKeyState,
		POINTL			pt,
		DWORD*			pdwEffect);

    // IMarkupServices methods
    NV_DECLARE_TEAROFF_METHOD(CreateMarkupPointer, createmarkuppointer, (
        IMarkupPointer** ppPointer));
    NV_DECLARE_TEAROFF_METHOD(CreateMarkupContainer, createmarkupcontainer, (
        IMarkupContainer** ppMarkupContainer));
    NV_DECLARE_TEAROFF_METHOD(CreateElement, createelement, (
        ELEMENT_TAG_ID tagID,
        OLECHAR* pchAttributes,
        IHTMLElement** ppElement));
    NV_DECLARE_TEAROFF_METHOD(InsertElement, insertelement, (
        IHTMLElement* pElementInsert,
        IMarkupPointer* pPointerStart,
        IMarkupPointer* pPointerFinish));
    NV_DECLARE_TEAROFF_METHOD(RemoveElement, removeelement, (
        IHTMLElement* pElementRemove));
    NV_DECLARE_TEAROFF_METHOD(Remove, remove, (
        IMarkupPointer* pPointerStart,
        IMarkupPointer* pPointerFinish));
    NV_DECLARE_TEAROFF_METHOD(Copy, copy, (
        IMarkupPointer* pPointerSourceStart,
        IMarkupPointer* pPointerSourceFinish,
        IMarkupPointer* pPointerTarget));
    NV_DECLARE_TEAROFF_METHOD(Move, move, (
        IMarkupPointer* pPointerSourceStart,
        IMarkupPointer* pPointerSourceFinish,
        IMarkupPointer* pPointerTarget));
    NV_DECLARE_TEAROFF_METHOD(InsertText, inserttext, (
        OLECHAR* pchText,
        long cch,
        IMarkupPointer* pPointerTarget));
    NV_DECLARE_TEAROFF_METHOD(IsScopedElement, isscopedelement, (IHTMLElement* pElement, BOOL* pfScoped));
    NV_DECLARE_TEAROFF_METHOD(GetElementTagId, getelementtagid, (IHTMLElement* pElement, ELEMENT_TAG_ID* ptagId));
    NV_DECLARE_TEAROFF_METHOD(GetTagIDForName, gettagidforname, (BSTR bstrName, ELEMENT_TAG_ID* ptagId));
    NV_DECLARE_TEAROFF_METHOD(GetNameForTagID, getnamefortagid, (ELEMENT_TAG_ID tagId, BSTR* pbstrName));
    NV_DECLARE_TEAROFF_METHOD(BeginUndoUnit, beginundounit, (OLECHAR* pchUnitTitle));
    NV_DECLARE_TEAROFF_METHOD(EndUndoUnit, endundounit, ());

	// Internal version of Markup Services routines which don't take interfaces.
	HRESULT CreateMarkupPointer(CMarkupPointer** ppPointer);
	HRESULT InsertElement(
		CElement*		pElementInsert,
		CMarkupPointer*	pPointerStart,
		CMarkupPointer*	pPointerFinish,
		DWORD			dwFlags=NULL);
	HRESULT RemoveElement(
		CElement*		pElementRemove,
		DWORD			dwFlags=NULL);
	HRESULT Remove(
		CMarkupPointer*	pPointerStart,
		CMarkupPointer*	pPointerFinish,
		DWORD			dwFlags=NULL);
	HRESULT Copy(
		CMarkupPointer*	pPointerSourceStart,
		CMarkupPointer*	pPointerSourceFinish,
		CMarkupPointer*	pPointerTarget,
		DWORD			dwFlags=NULL);
	HRESULT Move(
		CMarkupPointer*	pPointerSourceStart,
		CMarkupPointer*	pPointerSourceFinish,
		CMarkupPointer*	pPointerTarget,
		DWORD			dwFlags=NULL);
	HRESULT InsertText(
		CMarkupPointer*	pPointerTarget,
		const OLECHAR*	pchText,
		long			cch,
		DWORD			dwFlags=NULL);

    // IHTMLViewServices methods
    NV_DECLARE_TEAROFF_METHOD(MoveMarkupPointerToPoint, movemarkuppointertopoint, (
        POINT   pt,
        IMarkupPointer* pPointer,
        BOOL*   pfNotAtBOL,
        BOOL*   pfAtLogicalBOL,
        BOOL*   pfRightOfCp, 
        BOOL    fScrollIntoView));
    NV_DECLARE_TEAROFF_METHOD(MoveMarkupPointerToPointEx, movemarkuppointertopointex, (
        POINT   pt,
        IMarkupPointer* pPointer,
        BOOL    fGlobalCoordinates,
        BOOL*   pfNotAtBOL,
        BOOL*   pfAtLogicalBOL,
        BOOL*   pfRightOfCp, 
        BOOL    fScrollIntoView));
    NV_DECLARE_TEAROFF_METHOD(MoveMarkupPointerToMessage, movemarkuppointertomessage, (
        IMarkupPointer* pPointer,
        SelectionMessage* pMessage,
        BOOL*   pfNotAtBOL,
        BOOL*   pfAtLogicalBOL,
        BOOL*   pfRightOfCp,
        BOOL*   pfValidTree,
        BOOL    fScrollIntoView,
        IHTMLElement* pIContainerElement, 
        BOOL*   pfSameLayout,
        BOOL    fHitTestEndOfLine));
    NV_DECLARE_TEAROFF_METHOD(GetCharFormatInfo, getcharformatinfo, (
        IMarkupPointer* pPointer,
        WORD family,
        HTMLCharFormatData* pInfo));
    NV_DECLARE_TEAROFF_METHOD(GetLineInfo, getlineinfo, (
        IMarkupPointer* pPointer,
        BOOL fAtEndOfLine,
        HTMLPtrDispInfoRec* pInfo));
    NV_DECLARE_TEAROFF_METHOD( IsPointerBetweenLines, ispointerbetweenlines, (
        IMarkupPointer* pPointer,
        BOOL* pfBetweenLines ));
    NV_DECLARE_TEAROFF_METHOD(GetElementsInZOrder, getelementsinzorder, (
        POINT pt,
        IHTMLElement** rgElements,
        DWORD* pCount));
    NV_DECLARE_TEAROFF_METHOD(GetTopElement, gettopelement, (
        POINT pt,
        IHTMLElement** ppElement));
    NV_DECLARE_TEAROFF_METHOD(MoveMarkupPointer, movemarkuppointer, (
        IMarkupPointer* pPointer,
        LAYOUT_MOVE_UNIT eUnit,
        LONG lXCurReally,
        BOOL* pfNotAtBOL,
        BOOL* pfAtLogicalBOL));
    NV_DECLARE_TEAROFF_METHOD(RegionFromMarkupPointers, regionfrommarkuppointers, (
        IMarkupPointer* pPointerStart,
        IMarkupPointer* pPointerEnd,
        HRGN* phrgn));

    NV_DECLARE_TEAROFF_METHOD(GetCurrentSelectionRenderingServices, getcurrentselectionrenderingservices, (
        ISelectionRenderingServices** ppSelRenSvc));

    NV_DECLARE_TEAROFF_METHOD(GetCurrentSelectionSegmentList, getcurrentsegmentlist , (
        ISegmentList** ppSegmentList));


    NV_DECLARE_TEAROFF_METHOD(FireOnSelectStart, fireonselectstart , (
        IHTMLElement* pIElement)); 

    NV_DECLARE_TEAROFF_METHOD(FireCancelableEvent, firecancelableevent, (
        IHTMLElement* pIElement,
        LONG dispidMethod,
        LONG dispidProp,
        BSTR bstrEventType,
        BOOL* pfResult));

    NV_DECLARE_TEAROFF_METHOD(GetCaret, getcaret, (
        IHTMLCaret** ppCaret));

    NV_DECLARE_TEAROFF_METHOD(ConvertVariantFromHTMLToTwips,
        convertvariantfromhtmltotwips,
        (VARIANT* pvar));

    NV_DECLARE_TEAROFF_METHOD(ConvertVariantFromTwipsToHTML,
        convertvariantfromtwipstohtml,
        (VARIANT* pvar));

    NV_DECLARE_TEAROFF_METHOD(IsBlockElement, isblockelement, (
        IHTMLElement* pIElement,
        BOOL* pfResult));

    NV_DECLARE_TEAROFF_METHOD(IsLayoutElement, islayoutelement, (
        IHTMLElement* pIElement,
        BOOL* pfResult));

    NV_DECLARE_TEAROFF_METHOD(IsContainerElement, iscontainerelement, (
        IHTMLElement* pIElement,
        BOOL* pfContainer,
        BOOL* pfHTML));

    NV_DECLARE_TEAROFF_METHOD(GetFlowElement, getflowelement, (
        IMarkupPointer* pIPointer,
        IHTMLElement** ppIElement));

    NV_DECLARE_TEAROFF_METHOD(GetElementFromCookie, getelementfromcookie , (
        void* elementCookie,
        IHTMLElement** ppIElement));


    NV_DECLARE_TEAROFF_METHOD(InflateBlockElement, inflateblockelement, (
        IHTMLElement* pElem));

    NV_DECLARE_TEAROFF_METHOD(IsInflatedBlockElement, isinflatedblockelement, (
        IHTMLElement* pElem,
        BOOL* pfIsInflated));

    NV_DECLARE_TEAROFF_METHOD(IsMultiLineFlowElement, ismultilineflowelement, (
        IHTMLElement* pIElement,
        BOOL* fMultiLine));

    NV_DECLARE_TEAROFF_METHOD(GetElementAttributeCount, getelementattributecount, (
        IHTMLElement* pIElement, 
        UINT* pCount));

    NV_DECLARE_TEAROFF_METHOD(IsEditableElement, iseditableelement, (
        IHTMLElement* pIElement,
        BOOL* pfResult));
    NV_DECLARE_TEAROFF_METHOD(GetOuterContainer, getoutercontainer , (
        IHTMLElement* pIElement,
        IHTMLElement** ppIOuterElement,
        BOOL fIgnoreOutermostContainer,
        BOOL* pfHitContainer));

    NV_DECLARE_TEAROFF_METHOD(IsNoScopeElement, isnoscopeelement, (
        IHTMLElement* pIElement,
        BOOL* pfResult));

    NV_DECLARE_TEAROFF_METHOD(ShouldObjectHaveBorder, shouldobjecthaveborder , (
        IHTMLElement* pIElement,
        BOOL* pfResult));

    NV_DECLARE_TEAROFF_METHOD(DoTheDarnPasteHTML, dothedarnpastehtml , ( 
        IMarkupPointer*, IMarkupPointer*, HGLOBAL));

    NV_DECLARE_TEAROFF_METHOD(ConvertRTFToHTML, convertrtftohtml, (
        LPOLESTR pszRtf, HGLOBAL* phglobalHTML));

    NV_DECLARE_TEAROFF_METHOD(GetViewHWND, getviewhwnd , (HWND* pHWND));

    NV_DECLARE_TEAROFF_METHOD(ScrollElement, scrollelement, ( 
        IHTMLElement* pIElement,
        LONG lPercentToScroll,
        POINT* pScrollDelta));

    NV_DECLARE_TEAROFF_METHOD(GetScrollingElement, getscrollingelement, (
        IMarkupPointer* pPosition,
        IHTMLElement* pBoundary,
        IHTMLElement** ppElement));

    NV_DECLARE_TEAROFF_METHOD(StartHTMLEditorDblClickTimer, starthtmleditordblclicktimer , ());
    NV_DECLARE_TEAROFF_METHOD(StopHTMLEditorDblClickTimer, stophtmleditordblclicktimer , ());
    NV_DECLARE_TEAROFF_METHOD(HTMLEditorTakeCapture, htmleditortakecapture , ());
    NV_DECLARE_TEAROFF_METHOD(HTMLEditorReleaseCapture, htmleditorreleasecapture , ());  
    NV_DECLARE_TEAROFF_METHOD(SetHTMLEditorMouseMoveTimer, sethtmleditormousemovetimer , ()); 

    NV_DECLARE_TEAROFF_METHOD(GetEditContext, geteditcontext, ( 
        IHTMLElement* pIStartElement, 
        IHTMLElement** ppIEditThisElement, 
        IMarkupPointer* pIStart, 
        IMarkupPointer* pIEnd, 
        BOOL fDrillingIn,  
        BOOL* pfEditThisEditable ,
        BOOL* pfEditParentEditable,
        BOOL* pfNoScopeElement));

    NV_DECLARE_TEAROFF_METHOD(EnsureEditContext, ensureeditcontext, (IMarkupPointer* pIPointer));

    NV_DECLARE_TEAROFF_METHOD(ScrollPointerIntoView, scrollpointerintoview, (
        IMarkupPointer* pPointer,
        BOOL fNotAtBOL,
        POINTER_SCROLLPIN eScrollAmount));

    NV_DECLARE_TEAROFF_METHOD(ScrollPointIntoView, scrollpointintoview, (
        IHTMLElement* pIElement,
        POINT* ptGlobal));

    NV_DECLARE_TEAROFF_METHOD(ArePointersInSameMarkup, arepointersinsamemarkup, ( 
        IMarkupPointer* pP1, 
        IMarkupPointer* pP2, 
        BOOL* pfSameMarkup));

    NV_DECLARE_TEAROFF_METHOD(DragElement, dragelement, (
        IHTMLElement* pIElement,
        DWORD dwKeyState));

    NV_DECLARE_TEAROFF_METHOD(BecomeCurrent, becomecurrent , (
        IHTMLElement* pIElement,
        SelectionMessage* pSelMessage));

    NV_DECLARE_TEAROFF_METHOD(TransformPoint, transformpoint, (
        POINT*          pPoint,
        COORD_SYSTEM    eSource,
        COORD_SYSTEM    eDestination,
        IHTMLElement*   pIElement));

    NV_DECLARE_TEAROFF_METHOD(GetElementDragBounds, getelementdragbounds, (
        IHTMLElement* pIElement,
        RECT* pIElementDragBounds));

    NV_DECLARE_TEAROFF_METHOD(IsElementLocked, iselementlocked, (
        IHTMLElement* pIElement,
        BOOL* pfLocked));

    NV_DECLARE_TEAROFF_METHOD(GetLineDirection, getlinedirection, (
        IMarkupPointer* pPointer,
        BOOL fAtEndOfLine,
        long* peHTMLDir));

    NV_DECLARE_TEAROFF_METHOD(MakeParentCurrent, makeparentcurrent, (
        IHTMLElement* pIElement));


    NV_DECLARE_TEAROFF_METHOD(FireOnBeforeEditFocus, fireonbeforeeditfocus, (
        IHTMLElement* pINextActiveElem, BOOL* pfRet));

    NV_DECLARE_TEAROFF_METHOD(IsEqualElement, isequalelement, (
        IHTMLElement* pIElement1, IHTMLElement* pIElement2));

    NV_DECLARE_TEAROFF_METHOD(GetOuterMostEditableElement, getoutermosteditableelement, ( 
        IHTMLElement* pElement, IHTMLElement** ppIOuterElement));

    NV_DECLARE_TEAROFF_METHOD(IsSite, issite, ( 
        IHTMLElement* pElement, 
        BOOL* pfSite, 
        BOOL* pfText, 
        BOOL* pfMultiLine, 
        BOOL* pfScrollable));

    NV_DECLARE_TEAROFF_METHOD(QueryBreaks, querybreaks, (
        IMarkupPointer* pPointer,
        DWORD* pdwBreaks,
        BOOL fWantPendingBreak));

    NV_DECLARE_TEAROFF_METHOD(MergeDeletion, mergedeletion, (
        IMarkupPointer* pPointer));

    NV_DECLARE_TEAROFF_METHOD(GetElementForSelection, getelementforselection, ( 
        IHTMLElement* pElement, IHTMLElement** ppISiteSelectableElement));

    NV_DECLARE_TEAROFF_METHOD(IsContainedBy, iscontainedby, (
        IMarkupContainer* pIOuterContainer, 
        IMarkupContainer* pIInnerContainer));

    NV_DECLARE_TEAROFF_METHOD(CurrentScopeOrSlave, currentscopeorslave, (
        IMarkupPointer* pPointer, 
        IHTMLElement** ppElemCurrent));

    NV_DECLARE_TEAROFF_METHOD(LeftOrSlave, leftorslave, (
        IMarkupPointer* pPointer,
        BOOL fMove,
        MARKUP_CONTEXT_TYPE* pContext,
        IHTMLElement** ppElement,
        long* pcch,
        OLECHAR* pchText));

    NV_DECLARE_TEAROFF_METHOD(RightOrSlave, rightorslave, (
        IMarkupPointer* pPointer,
        BOOL fMove,
        MARKUP_CONTEXT_TYPE* pContext,
        IHTMLElement** ppElement,
        long* pcch,
        OLECHAR* pchText));

    NV_DECLARE_TEAROFF_METHOD(MoveToContainerOrSlave, movetocontainerorslave, (
        IMarkupPointer* pPointer,
        IMarkupContainer* pContainer,
        BOOL fAtStart));

    NV_DECLARE_TEAROFF_METHOD(FindUrl, findurl, (
        IMarkupPointer* pStart, 
        IMarkupPointer* pEnd,   
        IMarkupPointer* pUrlStart, 
        IMarkupPointer* pUrlEnd));

    NV_DECLARE_TEAROFF_METHOD(IsEnabled, isenabled , (
        IHTMLElement* pIHTMLElement, BOOL* pfEnabled));

    NV_DECLARE_TEAROFF_METHOD(GetElementBlockDirection, getelementblockdirection, (
        IHTMLElement* pElement, BSTR* pbstrDir));

    NV_DECLARE_TEAROFF_METHOD(SetElementBlockDirection, setelementblockdirection, (
        IHTMLElement* pElement, LONG eHTMLDir));

    NV_DECLARE_TEAROFF_METHOD(IsBidiEnabled, isbidienabled , (
        BOOL* pfEnabled));

    NV_DECLARE_TEAROFF_METHOD(SetDocDirection, setdocdirection , (
        LONG eHTMLDir));

    NV_DECLARE_TEAROFF_METHOD(AllowSelection , allowselection , (
        IHTMLElement* pIHTMLElement, SelectionMessage* peMessage));

    NV_DECLARE_TEAROFF_METHOD(MoveWord, moveword, (
        IMarkupPointer* pPointerToMove,
        MOVEUNIT_ACTION muAction,
        IMarkupPointer* pLeftBoundary,
        IMarkupPointer* pRightBoundary));

    NV_DECLARE_TEAROFF_METHOD(GetAdjacencyForPoint, getadjacencyforpoint, ( 
        IHTMLElement* pIElement, 
        POINT ptGlobal, 
        ELEMENT_ADJACENCY* peAdjacent));

    NV_DECLARE_TEAROFF_METHOD(SaveSegmentsToClipboard, savesegmentstoclipboard, (
        ISegmentList* pSegmentList));

    NV_DECLARE_TEAROFF_METHOD(InsertMaximumText, insertmaximumtext, (                                    
        OLECHAR* pstrText, 
        LONG cch,
        IMarkupPointer* pMarkupPointer));

    NV_DECLARE_TEAROFF_METHOD(IsInsideURL, isinsideurl, (
        IMarkupPointer*, IMarkupPointer*, BOOL*));
    NV_DECLARE_TEAROFF_METHOD(GetDocHostUIHandler, getdochostuihandler, (
        IDocHostUIHandler**));

    NV_DECLARE_TEAROFF_METHOD(GetClientRect, getclientrect, (
        IHTMLElement* pIElement, COORD_SYSTEM eSource, RECT* pRect));

    NV_DECLARE_TEAROFF_METHOD(GetContentRect, getcontentrect, (
        IHTMLElement* pIElement, COORD_SYSTEM eSource, RECT* pRect));
    NV_DECLARE_TEAROFF_METHOD(IsElementSized, iselementsized, (
        IHTMLElement* pIElement, BOOL* pfSized));

	// focus/default site helper
	BOOL		TakeFocus();
	BOOL		HasFocus();
	HRESULT		InvalidateDefaultSite();

#define _CDoc_
#include "../gen/document.hdl"

	HRESULT CreateElement(ELEMENT_TAG etag, CElement** ppElementNew);

	// Called to traverse the site hierarchy
	//   call Notify on registered elements and on super
	void BroadcastNotify(CNotification* pNF);

    virtual CAtomTable* GetAtomTable(BOOL* pfExpando=NULL)
    {
        if(pfExpando)
        {
            *pfExpando = _fExpando;
        }
        return &_AtomTable;
    }

	// Group helper
	HRESULT DocTraverseGroup(
		LPCTSTR		strGroupName,
		PFN_VISIT	pfn,
		DWORD_PTR	dw,
		BOOL		fForward);

	CElement* FindDefaultElem(BOOL fDefault, BOOL fCurrent=FALSE);
	HRESULT HandleKeyNavigate(CMessage* pmsg, BOOL fAccessKeyCycle);

    // CServer overrides
    const CBase::CLASSDESC* GetClassDesc() const;
    virtual void    Passivate();
    virtual HRESULT OnPropertyChange(DISPID, DWORD);
    virtual HRESULT Draw(CDrawInfo* pDI, RECT* prc);
    HRESULT DoTranslateAccelerator(LPMSG lpmsg);
    HRESULT FireEventHelper(
        DISPID  dispidMethod,
        DISPID  dispidProp,
        BYTE*   pbTypes, ...);
    HRESULT FireEvent(
        DISPID  dispidMethod,
        DISPID  dispidProp,
        LPCTSTR pchEventType,
        BYTE*   pbTypes, ...);

    HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpmsg);

    // Helper function
    HRESULT FireAccesibilityEvents(DISPID dispidEvent, long lElemID);

	// State transition callbacks
	virtual HRESULT	RunningToInPlace(HWND hWndSite, LPMSG lpmsg);
	virtual HRESULT	InPlaceToRunning(void);

	virtual HRESULT	AttachWin(HWND hwndParent, RECT* prc, HWND* phWnd);
	virtual void	DetachWin();

    // Internal
    void    UnloadContents(BOOL fPrecreated, BOOL fRestartLoad);

    void    InitDownloadInfo(DWNLOADINFO* pdli);

    HRESULT SetClipboard(IDataObject* pDO);

	void    UpdateDocTreeVersion()     { __lDocTreeVersion++; UpdateDocContentsVersion(); }
	void    UpdateDocContentsVersion() { __lDocContentsVersion++; }

	// Site management
	HTC     HitTestPoint(CMessage* pMessage, CTreeNode** ppNodeElement, DWORD dwFlags);

    NV_DECLARE_ONCALL_METHOD(OnControlInfoChanged, oncontrolinfochanged, (DWORD_PTR));

    void    AmbientPropertyChange(DISPID dispidProp);

    NV_DECLARE_TEAROFF_METHOD(QueryService, queryservice, (REFGUID guidService, REFIID iid, LPVOID* ppv));
    HRESULT     CreateService(REFGUID guidService, REFIID iid, LPVOID* ppv);

	HRESULT		GetUniqueIdentifier(CString* pStr);

	CElement*	GetPrimaryElementClient();
	CElement*	GetPrimaryElementTop();

	// Coordinates
	void		DocumentFromWindow(POINTL* pptlDocOut, long xWinIn, long yWinIn);
	void		DocumentFromWindow(POINTL* pptlDocOut, POINT pptWinIn);
	void		DocumentFromWindow(SIZEL* psizelDocOut, SIZE sizeWinIn);
	void		DocumentFromWindow(SIZEL* psizelDocOut, long cxWinin, long cyWinIn);
	void		DocumentFromWindow(RECTL* prclDocOut, const RECT* prcWinIn);

	void		DocumentFromScreen(POINTL* pptlDocOut, POINT pptScreenIn);

	void		ScreenFromWindow (POINT* ppt, POINT pt);

	void		HimetricFromDevice(POINTL* pptlDocOut, int xWinIn, int yWinIn);
	void		HimetricFromDevice(POINTL* pptlDocOut, POINT pptWinIn);
	void		HimetricFromDevice(SIZEL* psizelDocOut, SIZE sizeWinIn);
	void		HimetricFromDevice(SIZEL* psizelDocOut, int cxWinin, int cyWinIn);
	void		HimetricFromDevice(RECTL* prclDocOut, const RECT* prcWinIn);

	void		DeviceFromHimetric(POINT* pptWinOut,  int xlDocIn, int ylDocIn);
	void		DeviceFromHimetric(POINT* pptWinOut,  POINTL pptlDocIn);
	void		DeviceFromHimetric(RECT* prcWinOut,   const RECTL* prclDocIn);
	void		DeviceFromHimetric(SIZE* psizeWinOut, SIZEL sizelDocIn);
	void		DeviceFromHimetric(SIZE* psizeWinOut, int cxlDocIn, int cylDocIn);

	// Rendering
	void		Invalidate();
	void		Invalidate(const RECT* prc, const RECT* prcClip, HRGN hrgn, DWORD dwFlags);
	void		UpdateForm();
	void		UpdateInterval(LONG interval);
	LONG		GetUpdateInterval();

	// UI
	void        DeferUpdateUI();
	void        SetUpdateTimer();
	void        OnUpdateUI();

    // Persistence
    HRESULT     NewDwnCtx(UINT dt, LPCTSTR pchSrc, CElement* pel, CDwnCtx** ppDwnCtx, BOOL fSilent=FALSE, DWORD dwProgsinkClass=0);
    HRESULT     SetDownloadNotify(IUnknown* punk);

	// MarkupServices helpers
	BOOL IsOwnerOf(IHTMLElement* pIElement);
	BOOL IsOwnerOf(IMarkupPointer* pIPointer);
	BOOL IsOwnerOf(IMarkupContainer* pContainer);

	HRESULT CutCopyMove(
		IMarkupPointer*	pIPointerStart,
		IMarkupPointer*	pIPointerFinish,
		IMarkupPointer*	pIPointerTarget,
		BOOL			fRemove);

	HRESULT CutCopyMove(
		CMarkupPointer*	pPointerStart,
		CMarkupPointer*	pPointerFinish,
		CMarkupPointer*	pPointerTarget,
		BOOL			fRemove,
		DWORD			dwFlags);

	HRESULT	CreateMarkup(CMarkup** ppMarkup, CElement* pElementMaster=NULL);
	HRESULT	CreateMarkupWithElement(CMarkup** ppMarkup, CElement* pElement, CElement* pElementMaster=NULL);

    // Option settings change
    HRESULT OnSettingsChange(BOOL fForce=FALSE);
    
    // Layout
	void	EnsureFormatCacheChange(DWORD dwFlags);
	HRESULT	ForceRelayout();

	// Window message handling
	void	OnPaint();
	BOOL	OnEraseBkgnd(HDC hdc);
	void	OnMenuSelect(UINT uItem, UINT fuFlags, HMENU hmenu);
	void	OnCommand(int id, HWND hwndCtl, UINT codeNotify);
	HRESULT	OnHelp(HELPINFO* );

	HRESULT	PumpMessage(CMessage* pMessage, CTreeNode* pNodeTarget, BOOL fPerformTA=FALSE);
	HRESULT	PerformTA(CMessage* pMessage);
	BOOL	AreWeSaneAfterEventFiring(CMessage* pMessage, ULONG cDie);
	BOOL	FCaretInPre();

	HRESULT	OnMouseMessage(
		UINT		msg,
		WPARAM		wParam,
		LPARAM		lParam,
		LRESULT*	plResult,
		int			x,
		int			y);
    NV_DECLARE_ONTICK_METHOD(DetectMouseExit, detectmouseexit, (UINT uTimerID));
	void DeferSetCursor();
    NV_DECLARE_ONCALL_METHOD(SendSetCursor, sendSetCursor, (DWORD_PTR));

    // REgister/revoke drag-drop as appropriate
    NV_DECLARE_ONCALL_METHOD(EnableDragDrop, enabledragdrop, (DWORD_PTR));

	void	SetMouseCapture(PFN_VOID_MOUSECAPTURE pfnTo, void* pvObject, BOOL fElement=TRUE);
	void	ClearMouseCapture(void* pvObject);
	void*	GetCaptureObject() { return _pvCaptureObject; }

protected:
    static UINT _g_msgHtmlGetobject; // Registered window message for WM_HTML_GETOBJECT

private:
    DWORD   GetTargetTime(DWORD dwTimeoutLenMs) { return GetTickCount()+dwTimeoutLenMs; }
    HRESULT ExecuteTimeoutScript(TIMEOUTEVENTINFO* pTimeout);

public:
    HRESULT SetUIActiveElement(CElement* pElemNext);
    HRESULT SetCurrentElem(CElement* pElemNext, long lSubNext, BOOL* pfYieldFailed=NULL);
    void STDMETHODCALLTYPE DeferSetCurrency(DWORD_PTR dwContext);

    // Context menus
    HRESULT ShowDragContextMenu(
        POINTL  ptl,
        DWORD   dwAllowed,
        int*    piSelection,
        LPTSTR  lptszFileType);

    // Keyboard interface
    HRESULT	ActivateDefaultButton(LPMSG lpmsg=NULL);
    HRESULT	ActivateCancelButton(LPMSG lpmsg=NULL);

    HRESULT	ActivateFirstObject(LPMSG lpmsg);

    // Z order
    HRESULT	SetZOrder(int zorder);
    void	FixZOrder();

    // Tools
    HWND	CreateOverlayWindow(HWND hwnd);

    // Misc
    void	SetMouseMoveTimer(UINT uTimeOUt);
    HRESULT	GetBodyElement(IHTMLBodyElement** ppBody);
    HRESULT	GetBodyElement(CBodyElement** ppBody);

    virtual void PreSetErrorInfo();

    //  User Option Settings
    HRESULT UpdateFromRegistry(DWORD dwFlags=0, BOOL* pfNeedLayout=NULL);
    BOOL    RtfConverterEnabled();

    // Helper functions to read from the registry
    HRESULT ReadOptionSettingsFromRegistry(DWORD dwFlags=0);
    HRESULT ReadCodepageSettingsFromRegistry(CODEPAGE cp, UINT uiFamilyCodePage, DWORD dwFlags=0);

    // The following functions should only be called from ReadOptionSettingsFromRegistry or
    //  from ReadCodepageSettingsFromRegistry():
    HRESULT EnsureOptionSettings();
    HRESULT EnsureCodepageSettings(UINT uiFamilyCodePage);

    // focus\blur method helpers for the window object
    HRESULT	WindowBlurHelper();
    HRESULT	WindowFocusHelper();

    // in-line function,
    CInPlace* InPlace() { return _pInPlace; }

    HRESULT GetBaseUrl(
        TCHAR**     ppszBase,
        CElement*   pElementContext=NULL,
        BOOL*       pfDefault=NULL,
        TCHAR*      pszAlternateDocUrl=NULL);

    HRESULT ExpandUrl(
        const TCHAR*    pszRel,
        LONG            dwUrlSize,
        TCHAR*          ppszUrl,
        CElement*       pElementContext,
        DWORD           dwFlags=0xFFFFFFFF,
        TCHAR*          pszAlternateDocUrl=NULL);

    BOOL IsVisitedHyperlink(LPCTSTR szURL, CElement* pElementContext);

    // Progress
    CProgSink*  GetProgSinkC();
    IProgSink*  GetProgSink();


    // Saves head section of html file
    HRESULT WriteDocHeader(CStreamWriteBuff* pStreamWriteBuff);


    NV_DECLARE_ONTICK_METHOD(FireTimeOut, firetimeout, (UINT timerID));
    HRESULT ClearTimeout(long lTimerID);


    // Print support
    BOOL    PaintBackground();
    HDC     GetHDC();
    BOOL    TiledPaintDisabled();

    // TAB order and accessKey support
    BOOL FindNextTabOrder(
        FOCUS_DIRECTION	dir,
        CElement*		pElemFirst,
        long			lSubFirst,
        CElement**		ppElemNext,
        long*			plSubNext);
    HRESULT OnElementEnter(CElement* pElement);
    HRESULT OnElementExit(CElement* pElement, DWORD dwExitFlags);
    void    ClearCachedNodeOnElementExit(CTreeNode** ppNodeToTest, CElement* pElement);
    BOOL    SearchFocusArray(
        FOCUS_DIRECTION	dir,
        CElement*		pElemFirst,
        long			lSubFirst,
        CElement**		ppElemNext,
        long*			plSubNext);
    BOOL    SearchFocusTree(
        FOCUS_DIRECTION	dir,
        CElement*		pElemFirst,
        long			lSubFirst,
        CElement**		ppElemNext,
        long*			plSubNext);
    void    SearchAccessKeyArray(
        FOCUS_DIRECTION	dir,
        CElement*		pElemFirst,
        CElement**		ppElemNext,
        CMessage*		pmsg);
    HRESULT InsertAccessKeyItem(CElement* pElement);
    HRESULT InsertFocusArrayItem(CElement* pElement);

    void    EnterStylesheetDownload(DWORD* pdwCookie);
    void    LeaveStylesheetDownload(DWORD* pdwCookie);
    BOOL    IsStylesheetDownload() { return (_cStylesheetDownloading>0); }

    // onblur and onfocus and Event firing helpers
    void    Fire_onfocus();
    void    Fire_onblur(BOOL fOnblurFiredFromWM=FALSE);
    void    Fire_onpropertychange(LPCTSTR strPropName);
    HRESULT Fire_PropertyChangeHelper(DISPID dispidProperty, DWORD dwFlags);
    void STDMETHODCALLTYPE FirePostedOnPropertyChange(DWORD_PTR dwContext)
    {
        // This is only for DISPID_CDoc_activeElement anyway.
        // If the document is not parse done, then there is a chance that we will 
        // not have the scripts ready. Wait for parse done. The parse done handler
        // will fire the event.
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_activeElement, FORMCHNG_NOINVAL);
        }
    }

    void    SetFocusWithoutFiringOnfocus();

    HRESULT	CreateRoot();

    CMarkup* PrimaryMarkup() { return _pPrimaryMarkup; }
    CRootElement* PrimaryRoot();

    BOOL    NeedRegionCollection() { return _fRegionCollection; }

    // Lookaside storage for elements
    void* GetLookasidePtr(void* pvKey) { return(_HtPvPv.Lookup(pvKey)); }
    HRESULT SetLookasidePtr(void* pvKey, void* pvVal) { return(_HtPvPv.Insert(pvKey, pvVal)); }
    void* DelLookasidePtr(void* pvKey) { return(_HtPvPv.Remove(pvKey)); }
    CHtPvPv _HtPvPv;

#ifdef _DEBUG
    BOOL AreLookasidesClear(void* pvKey, int nLookasides);
#endif

    DWORD   _dwTID;

public:
    // The default site is used when there is no other appropriate site
    // available
    CDefaultElement*	_pElementDefault;
    CMarkup*			_pPrimaryMarkup;	// Pointer to the current root site

private:
    // _EditRouter provides the plumbing for appropriate routing of editing commands
    CEditRouter         _EditRouter;
public:
    // _pCaret is the CCaret object. Use GetCaret to access the IHTMLCaret object.
    CCaret*             _pCaret;

    CDocInfo			_dci;				// Document transform
    CElement*			_pElemEditContext;	// element that contains the caret/current selection;
                                            // not necessarily the same as_pElemCurrent
    CElement*			_pElemUIActive;		// ptr to that's showing UI.  Need not
                                            // be the same as the current site.
    CElement*			_pElemCurrent;		// ptr to current element. Owner of focus
                                            // and commands get routed here.
    CElement*			_pElemDefault;		// ptr to default element.
    long				_lSubCurrent;		// Subdivision within element that is current
    CElement*			_pElemNext;			// the next element to be current.
                                            // used by SetCurrentElem during currency transitions.
    int					_cSurface;			// Number of requests to use offscreen surfaces
    int					_c3DSurface;		// Number of requests to use 3D surfaces (this count is included in _cSurface)

    OPTIONSETTINGS*		_pOptionSettings;	// Points to current user-configurable settings, like color, etc.
    CODEPAGESETTINGS*	_pCodepageSettings;	// Settings which have to do
    long				_icfDefault;		// Default CharFormat index based on option/codepage settings
    const CCharFormat*	_pcfDefault;		// Default CharFormat based on option/codepage settings

    void    ClearDefaultCharFormat();
    HRESULT CacheDefaultCharFormat();

    unsigned            _cInval;                //  Number of calls to CDocument::Invalidate()
    unsigned            _cProcessingTimeout;    // blocks clearing timeouts while exec'ing script

    RADIOGRPNAME*		_pRadioGrpName;		    // names of radio groups having checked radio

    LONG				_lRecursionLevel;	    // the recursion level of the measurer

    // View support
    CView				_view;
    CView* GetView()
    {
        return &_view;
    }

    BOOL OpenView(BOOL fBackground=FALSE)
    {
        return _view.OpenView(fBackground);
    }

    // Text ID pool
    long _lLastTextID;

    // The following version numbers are incremented when modifications
    // are made to any markup associated with this doc.
    //
    // The _lDocTreeVersion is incremented when the element structure of markups
    // are altered.  Modifications only to text between markup does not increment
    // this version number.
    //
    // The _lDocContentsVersion is incremented when any content is modified (text
    // or markup).
    long GetDocTreeVersion() { return __lDocTreeVersion; }
    long GetDocContentsVersion() { return __lDocContentsVersion; }

    // Do NOT modify these version numbers unless the document structure
    // or content is being modified.
    //
    // In particular, incrementing these to get a cache to rebuild is
    // BAD because it causes all sorts of other stuff to rebuilt.
    long __lDocTreeVersion;		// Element structure
    long __lDocContentsVersion;	// Any content

    // Contents version = tree version + character version

    // Object model SelectionObject
    //CSelectionObject*	_pCSelectionObject;	// selection object.
    CAtomTable  _AtomTable;		        // Mapping of elements to names

    // Host intergration
    DWORD       _dwFlagsHostInfo;
    DWORD       _dwFrameOptions;

    // This pair manages mouse capture
    CElement* _pElementOMCapture;

private:
    PFN_VOID_MOUSECAPTURE _pfnCapture;		// The fn to call
    void*				_pvCaptureObject;	// The object on which to call it

public:
    CElement*			_pMenuObject;		 // The site on which to call it
    CTreeNode*			_pNodeLastMouseOver; // element last fired mouseOver event
    long				_lSubDivisionLast;	 // Last subdivision over the mouse
    CTreeNode*			_pNodeGotButtonDown; // Site that last got a button down.

    USHORT              _usNumVerbs;         // Number of verbs on context menu.

    ULONG               _cFreeze;            // Count of the freeze factor

    CString             _cstrUrl;            // Default base Url for document (used internally)

#ifndef NO_IME
    // Input method context cache
    HIMC                _himcCache;          // Cached if window's context is temporarily disabled
#endif

    DWORD               _dwLoadf;            // Load flags (DLCTL_ + offline + silent)

    // Progress
    ULONG               _ulProgressPos;
    ULONG               _ulProgressMax;

    ULONG               _cStylesheetDownloading; // Counts stylesheets being downloaded
    DWORD               _dwStylesheetDownloadingCookie;

    DECLARE_CDataAry(CAryFocusItem, FOCUS_ITEM)
    DECLARE_CPtrAry(CAryAccessKey, CElement*)
    CAryFocusItem       _aryFocusItems;
    CAryAccessKey       _aryAccessKeyItems;
    long                _lFocusTreeVersion;

    DECLARE_CPtrAry(CAryElementReleaseNotify, CElement*)
    CAryElementReleaseNotify _aryElementReleaseNotify;

    HRESULT ReleaseNotify();
    HRESULT RequestReleaseNotify(CElement* pElement);
    HRESULT RevokeRequestReleaseNotify(CElement* pElement);

    CDwnDoc*                _pDwnDoc;
    IDownloadNotify*        _pDownloadNotify;

    // Drag-drop
    CDragDropSrcInfo*       _pDragDropSrcInfo;
    CDragDropTargetInfo*    _pDragDropTargetInfo;
    CDragStartInfo*         _pDragStartInfo;

    // Bit fields
    unsigned			_fOKEmbed:1;					// TRUE if can drop as embedding
    unsigned			_fOKLink:1;						// TRUE if can drop as link
    unsigned			_fDragFeedbackVis:1;			// Feedback rect has been drawn
    unsigned			_fIsDragDropSrc:1;				// Originated the current drag-drop operation
    unsigned			_fDisableTiledPaint:1;
    unsigned			_fUpdateUIPending:1;
    unsigned			_fNeedUpdateUI:1;
    unsigned			_fInPlaceActivating:1;
    unsigned			_fFromCtrlPalette:1;
    unsigned			_fRightBtnDrag:1;				// TRUE if right button drag occuring
    unsigned			_fSlowClick:1;					// TRUE if the user started
                                                        // a right drag but didn't move
                                                        // the mouse when ending the drag
    unsigned			_fElementCapture:1;				// TRUE if _pvCaptureObject is a CElement
    unsigned			_fOnLoseCapture:1;				// TRUE if we are in the process of firing onlosecapture

    unsigned			_fFiredOnLoad:1;				// TRUE if OnLoad has been fired
    unsigned			_fShownProgPos:1;				// TRUE if progress pos has been shown
    unsigned			_fShownProgMax:1;				// TRUE if progress max has been shown
    unsigned			_fShownProgText:1;				// TRUE if progress text has been shown
    unsigned			_fGotKeyDown:1;					// TRUE if we got a key down
    unsigned			_fGotKeyUp:1;					// TRUE if we got a key up
    unsigned			_fGotLButtonDown:1;				// TRUE if we got a left button down
    unsigned			_fGotMButtonDown:1;				// TRUE if we got a middle button down
    unsigned			_fGotRButtonDown:1;				// TRUE if we got a right button down
    unsigned			_fGotDblClk:1;					// TRUE if we got a double click message

    unsigned			_fMouseOverTimer:1;				// TRUE if MouseMove timer is set for detecting exit event
    unsigned			_fSuspendTimeout:1;				// TRUE if timeout firing is off
    unsigned			_fForceCurrentElem:1;			// TRUE if SetCurrentElem must succeed
    unsigned			_fCurrencySet:1;				// TRUE if currency has been set at least once (to an element other than the root)
    unsigned			_fExpando:1;
    unsigned			_fInhibitFocusFiring:1;			// TRUE if we shouldn't Fire_onfocus() when
                                                        // when handling the WM_SETFOCUS event
    unsigned			_fFirstTimeTab:1;				// TRUE if this is the first time
                                                        // Trident receives TAB key,
                                                        // should not process this message
                                                        // and move focus to address bar.

    unsigned			_fNeedTabOut:1;					// TRUE if we should not handle
                                                        // this SHIFT VK_TAB WM_KEYDOWN
                                                        // message, this is Raid 61972
    unsigned			_fFiredWindowFocus:1;			// TRUE if Window onfocus event has been fired,
                                                        // FALSE if Window onblur has been fired.

    unsigned			_fEnableInteraction:1;			// FALSE when the browser window is minimized or
                                                        // totally covered by other windows.

    unsigned			_fModalDialogInScript:1;		// Exclusively for use by PumpMessage() to
                                                        // figure out if an event handler put up a
                                                        // a modal dialog
    unsigned			_fInvalInScript:1;				// Exclusively for use by PumpMessage() to
                                                        // figure out if an event handler caused an
                                                        // invalidation.
    unsigned			_fInPumpMessage:1;				// Exclusively for use by PumpMessage() to
                                                        // to handle recursive calls

    unsigned			_fContainerCapture : 1;			// TRUE if elements inside containers should ignore OM capture

    unsigned			_fRegionCollection : 1;			// TRUE if region collection should be built
    unsigned			_fVID : 1;						// Call CDocument::OnControlInfoChanged when TRUE
    unsigned			_fPeersPossible : 1;			// true if there is a chance that peers could be on this page
    unsigned			_fHasOleSite : 1;				// There is an olesite somewhere
    unsigned			_fModalDialogInOnblur : 1;		// TRUE if a modal dialog pops up in script in any onblur event handler.
                                                        // Enables firing of onfocus again in such cases.
    unsigned			_fInEditCapture:1 ;				// TRUE if mshtmled has Capture
    unsigned			_fOnControlInfoChangedPosted:1;	// TRUE if GWPostMethodCall was called for CDocument::OnControlInfoChanged
    unsigned			_fSelectionHidden:1;			// TRUE if We have hidden the seletion from a WM_KILLFOCUS
    unsigned			_fVirtualHitTest:1;				// TRUE if virtual hit testing enabled
    unsigned			_fNotifyBeginSelection:1;		// TRUE if we are broadcasting a WM_BEGINSELECTION Notification
    unsigned			_fInhibitOnMouseOver:1;			// TRUE if onmouseover event should not be fired, like when over a popup window.
    unsigned			_fBroadcastInteraction:1;		// TRUE if broadcast of EnableInteraction is needed
    unsigned			_fBroadcastStop:1;				// TRUE if broadcast of Stop is needed
    unsigned			_fInTraverseGroup:1;			// TRUE if inside TraverseGroup for input radio
    unsigned			_fVisualOrder:1;				// the document is RTL visual (logical is LTR)
                                                        // This is used for ISO-8859-8 and ISO-8859-6
                                                        // visually ordered documents.
    unsigned			_fRTLDocDirection:1;			// TRUE if the document is right to left <HTML DIR=RTL>

    unsigned			_fPendingFilterCallback:1;		// A filter callback is pending
    unsigned			_fPassivateCalled:1;			// Trap multiple passivate calls

    WORD				_wUIState;

public:
    IHTMLEditor* GetHTMLEditor(BOOL fForceCreate=TRUE ); // Get the html editor, if we have to load mshtmled, should we force it?
    HRESULT GetEditingServices(IHTMLEditingServices** ppIServices);
    HRESULT Select(IMarkupPointer* pStart, IMarkupPointer* pEnd, SELECTION_TYPE eType);
    BOOL    IsElementSiteSelectable(CElement* pCurElement);
    BOOL    IsElementUIActivatable(CElement* pCurElement);
    BOOL    IsElementSiteSelected(CElement* pCurElement);
    HRESULT ProcessFollowUpAction(CMessage* pMessage, DWORD dwFollowUpAction);

private:
    IHTMLEditor* _pIHTMLEditor; // Selection Manager.

    BOOL ShouldCreateHTMLEditor(CMessage* pMessage);
    BOOL ShouldCreateHTMLEditor(SELECTION_NOTIFICATION eNotify);
    BOOL ShouldSetEditContext(CMessage* pMessage);
    NV_DECLARE_MOUSECAPTURE_METHOD(HandleEditMessageCapture, handleeditmessagecapture,
        (CMessage* pMessage));
    NV_DECLARE_ONTICK_METHOD(OnEditDblClkTimer, onselectdblclicktimer,
        (UINT idMessage));
    VOID SetClick(CMessage* inMessage);

    HRESULT DragElement(CMessage* pMessage);

public:
    CSelDragDropSrcInfo* GetSelectionDragDropSource();

    HRESULT UpdateCaret(
        BOOL        fScrollIntoView,    //@parm If TRUE, scroll caret into view if we have
                                        // focus or if not and selection isn't hidden
        BOOL        fForceScrollCaret,  //@parm If TRUE, scroll caret into view regardless
        CDocInfo*   pdci=NULL);
    BOOL    IsCaretVisible(BOOL* pfPositioned=NULL);
    HRESULT HandleSelectionMessage(CMessage* inMessage, BOOL fForceCreation);
    HRESULT SetEditContext(CElement* pElement, BOOL fForceCreation, BOOL fSetSelection, BOOL fDrillingIn=FALSE);

    HRESULT NotifySelection(
        SELECTION_NOTIFICATION  eSelectionNotification,
        IUnknown*               pUnknown,
        DWORD                   dword=0);

    SELECTION_TYPE GetSelectionType();
    BOOL    HasTextSelection();
    BOOL    HasSelection();

    BOOL    IsPointInSelection(POINT pt, CTreeNode* pNode=NULL, BOOL fPtIsContent=FALSE);

    HRESULT ScrollPointersIntoView(IMarkupPointer* pStart, IMarkupPointer* pEnd);

    LPCTSTR GetCursorForHTC(HTC inHTC);

    HRESULT AdjustEditContext( 
        CElement*       pStartElement, 
        CElement**      ppEditThisElement, 
        IMarkupPointer* pStart,
        IMarkupPointer* pEnd , 
        BOOL*           pfEditThisEditable,
        BOOL*           pfEditParentEditable,
        BOOL*           pfNoScopeElement,
        BOOL            fSetCurrencyIfInvalid=FALSE,
        BOOL            fSetElemEditContext=TRUE,
        BOOL            fDrillingIn=FALSE);

    HRESULT GetEditContext( 
        CElement*       pStartElement, 
        CElement**      ppEditThisElement, 
        IMarkupPointer* pIStart=NULL,
        IMarkupPointer* pIEnd=NULL,
        BOOL            fDrillingIn=FALSE,
        BOOL*           pfEditThisEditable=NULL,
        BOOL*           pfEditParentEditable=NULL,
        BOOL*           pfNoScopeElement=NULL);

    HRESULT EnsureEditContext(CElement* pElement, BOOL fDrillingIn=TRUE);

    // 3D border setting
    BYTE                _b3DBorder;
    
    // Rendering properties
    LONG				_bufferDepth;           // sets bits-per-pixel for offscreen buffer
    ITimer*             _pTimerDraw;            // NAMEDTIMER_DRAW, for sync'ing control with paint

    // Persistent state
    ULONG               _ID;

public:
    // Load stuff
    ULONG               _cDie;                  // Incremented whenever UnloadContents is called

    // list of objects that failed to initialize
    DECLARE_CPtrAry(CAryDefunctObjects, CObjectElement*)
    CAryDefunctObjects  _aryDefunctObjects;

    CTimeoutEventList   _TimeoutEvents;         // List for active timeouts
    EVENTPARAM*         _pparam;                // Ptr to event params

    // Helper to clean up script timers
    void CleanupScriptTimers(void);

private:
    //-------------------------------------------------------------------------
    // Cache of loaded images (background, list bullets, etc)
    DECLARE_CPtrAry(CAryUrlImgCtxElems, CElement*)

    struct URLIMGCTX
    {
        LONG                lAnimCookie;
        ULONG               ulRefs;
        CImgCtx*            pImgCtx;
        CString             cstrUrl;
        CAryUrlImgCtxElems  aryElems;
    };

    DECLARE_CDataAry(CAryUrlImgCtx, URLIMGCTX)
    CAryUrlImgCtx _aryUrlImgCtx;

public:
    HRESULT         AddRefUrlImgCtx(LPCTSTR lpszUrl, CElement* pElemContext, LONG* plCookie);
    HRESULT         AddRefUrlImgCtx(LONG lCookie, CElement* pElem);
    CImgCtx*        GetUrlImgCtx(LONG lCookie);
    IMGANIMSTATE*   GetImgAnimState(LONG lCookie);
    void            ReleaseUrlImgCtx(LONG lCookie, CElement* pElem);
    void            StopUrlImgCtx();
    void            UnregisterUrlImgCtxCallbacks();
    static void CALLBACK OnUrlImgCtxCallback(void*, void*);
    BOOL            OnUrlImgCtxChange(URLIMGCTX* purlimgctx, ULONG ulState);
    static void     OnAnimSyncCallback(void* pvObj, DWORD dwReason, void* pvArg, void** ppvDataOut, IMGANIMSTATE* pAnimState);

    static const CLASSDESC s_classdesc;

    // Scaling factor for text
    SHORT _sBaselineFont;
    SHORT GetBaselineFont() { return _sBaselineFont; }

    // Internationalization
    CODEPAGE _codepage;
    CODEPAGE _codepageFamily;
    BOOL     _fCodepageOverridden;
    CODEPAGE GetCodePage() { return _codepage; }
    CODEPAGE GetFamilyCodePage() { return _codepageFamily; }
    HRESULT  SwitchCodePage(CODEPAGE codepage);

    HRESULT GetDocDirection(BOOL* pfRTL);

    BOOL    IsCpAutoDetect(void);

    // behaviors support
    void SetPeersPossible();

    // Filter tasks
    // These are elements that need filter hookup
    DECLARE_CPtrAry(CPendingFilterElementArray, CElement*)
    CPendingFilterElementArray _aryPendingFilterElements;

    BOOL ExecuteSingleFilterTask(CElement* pElement);
    BOOL AddFilterTask(CElement* pElement);
    void RemoveFilterTask(CElement* pElement);

    BOOL ExecuteFilterTasks();
    void PostFilterCallback();
    NV_DECLARE_ONCALL_METHOD(FilterCallback, filtercallback, (DWORD_PTR unused));

    // Helper function for view services methods.
    HRESULT RegionFromMarkupPointers(
        CMarkupPointer*	pStart,
        CMarkupPointer*	pEnd,
        CDataAry<RECT>*	paryRects,
        RECT*			pBoundingRect);

    CTreeNode* GetNodeFromPoint(
        const POINT&	pt,
        BOOL			fGlobalCoordinates,
        POINT*			pptGlobalPoint=NULL,
        LONG*			plCpHitMaybe=NULL ,
        BOOL*			pfEmptySpace=NULL);

    CFlowLayout* GetFlowLayoutForSelection(CTreeNode* pNode);

    CLayout* GetLayoutForSelection(CTreeNode* pNode);

    HRESULT MovePointerToPointInternal(
        POINT			tContentPoint,
        CTreeNode*		pNode,
        CMarkupPointer*	pPointer,
        BOOL*			pfNotAtBOL,
        BOOL*			pfAtLogicalBOL,
        BOOL*			pfRightOfCp,
        BOOL			fScrollIntoView,
        CLayout*		pContainingLayout, 
        BOOL*			pfSameLayout=NULL,
        BOOL			fHitTestEOL=TRUE);

    CMarkup* GetCurrentMarkup();
};

inline void CDocument::DocumentFromWindow(POINTL* pptlDocOut, long xWinIn, long yWinIn)
{
    _dci.DocumentFromWindow(pptlDocOut, xWinIn, yWinIn);
}

inline void CDocument::DocumentFromWindow(POINTL* pptlDocOut, POINT ptWinIn)
{
    _dci.DocumentFromWindow(pptlDocOut, ptWinIn.x, ptWinIn.y);
}

inline void CDocument::DocumentFromWindow(SIZEL* pptlDocOut, SIZE ptWinIn)
{
    _dci.DocumentFromWindow(pptlDocOut, ptWinIn.cx, ptWinIn.cy);
}

inline void CDocument::DocumentFromWindow(SIZEL* psizelDocOut, long cxWinin, long cyWinIn)
{
    _dci.DocumentFromWindow(psizelDocOut, cxWinin, cyWinIn);
}

inline void CDocument::HimetricFromDevice(POINTL* pptlDocOut, int xWinIn, int yWinIn)
{
    _dci.HimetricFromDevice(pptlDocOut, xWinIn, yWinIn);
}

inline void CDocument::HimetricFromDevice(POINTL* pptlDocOut, POINT ptWinIn)
{
    _dci.HimetricFromDevice(pptlDocOut, ptWinIn.x, ptWinIn.y);
}

inline void CDocument::HimetricFromDevice(SIZEL* psizelDocOut, int cxWinin, int cyWinIn)
{
    _dci.HimetricFromDevice(psizelDocOut, cxWinin, cyWinIn);
}

inline void CDocument::HimetricFromDevice(SIZEL* pptlDocOut, SIZE ptWinIn)
{
    _dci.HimetricFromDevice(pptlDocOut, ptWinIn.cx, ptWinIn.cy);
}

inline void CDocument::DeviceFromHimetric(POINT* pptWinOut, int xlDocIn, int ylDocIn)
{
    _dci.DeviceFromHimetric(pptWinOut, xlDocIn, ylDocIn);
}

inline void CDocument::DeviceFromHimetric(POINT* pptPhysOut, POINTL ptlLogIn)
{
    _dci.DeviceFromHimetric(pptPhysOut, ptlLogIn.x, ptlLogIn.y);
}

inline void CDocument::DeviceFromHimetric(SIZE* psizeWinOut, int cxlDocIn, int cylDocIn)
{
    _dci.DeviceFromHimetric(psizeWinOut, cxlDocIn, cylDocIn);
}

inline void CDocument::DeviceFromHimetric(SIZE* pptPhysOut, SIZEL ptlLogIn)
{
    _dci.DeviceFromHimetric(pptPhysOut, ptlLogIn.cx, ptlLogIn.cy);
}


HRESULT InitDocClass(void);

HRESULT ExpandUrlWithBaseUrl(LPCTSTR pchBaseUrl, LPCTSTR pchRel, TCHAR** ppchUrl);


BOOL IsTridentHwnd(HWND hwnd);
static BOOL CALLBACK OnChildBeginSelection(HWND hwnd, LPARAM lParam);

HRESULT FaultInIEFeatureHelper(HWND hWnd, uCLSSPEC* pClassSpec, QUERYCONTEXT* pQuery, DWORD dwFlags);

#endif //__XINDOWS_SITE_BASE_DOC_H__