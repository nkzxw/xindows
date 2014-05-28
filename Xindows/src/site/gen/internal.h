#ifndef __internal_h__
#define __internal_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

/* header files for imported files */
#include "document.h"

/* header files for imported files */
#include "mshtmext.h"

#ifndef __IMarkupPointer_FWD_DEFINED__
#define __IMarkupPointer_FWD_DEFINED__
typedef interface IMarkupPointer IMarkupPointer;
#endif     /* __IMarkupPointer_FWD_DEFINED__ */

#ifndef __IHTMLElement_FWD_DEFINED__
#define __IHTMLElement_FWD_DEFINED__
typedef interface IHTMLElement IHTMLElement;
#endif     /* __IHTMLElement_FWD_DEFINED__ */

#ifndef __ISegmentList_FWD_DEFINED__
#define __ISegmentList_FWD_DEFINED__
typedef interface ISegmentList ISegmentList;
#endif     /* __ISegmentList_FWD_DEFINED__ */

#ifndef __IHTMLCaret_FWD_DEFINED__
#define __IHTMLCaret_FWD_DEFINED__
typedef interface IHTMLCaret IHTMLCaret;
#endif     /* __IHTMLCaret_FWD_DEFINED__ */

#ifndef __ISelectionRenderingServices_FWD_DEFINED__
#define __ISelectionRenderingServices_FWD_DEFINED__
typedef interface ISelectionRenderingServices ISelectionRenderingServices;
#endif     /* __ISelectionRenderingServices_FWD_DEFINED__ */

#ifndef __IHTMLViewServices_FWD_DEFINED__
#define __IHTMLViewServices_FWD_DEFINED__
typedef interface IHTMLViewServices IHTMLViewServices;
#endif     /* __IHTMLViewServices_FWD_DEFINED__ */

#ifndef __IHTMLCaret_FWD_DEFINED__
#define __IHTMLCaret_FWD_DEFINED__
typedef interface IHTMLCaret IHTMLCaret;
#endif     /* __IHTMLCaret_FWD_DEFINED__ */

#ifndef __ISegmentList_FWD_DEFINED__
#define __ISegmentList_FWD_DEFINED__
typedef interface ISegmentList ISegmentList;
#endif     /* __ISegmentList_FWD_DEFINED__ */

#ifndef __ISelectionRenderingServices_FWD_DEFINED__
#define __ISelectionRenderingServices_FWD_DEFINED__
typedef interface ISelectionRenderingServices ISelectionRenderingServices;
#endif     /* __ISelectionRenderingServices_FWD_DEFINED__ */

#ifndef __IHTMLEditor_FWD_DEFINED__
#define __IHTMLEditor_FWD_DEFINED__
typedef interface IHTMLEditor IHTMLEditor;
#endif     /* __IHTMLEditor_FWD_DEFINED__ */

#ifndef __IHTMLEditingServices_FWD_DEFINED__
#define __IHTMLEditingServices_FWD_DEFINED__
typedef interface IHTMLEditingServices IHTMLEditingServices;
#endif     /* __IHTMLEditingServices_FWD_DEFINED__ */
typedef enum _CHAR_FORMAT_FAMILY
{
    CHAR_FORMAT_None = 0,
    CHAR_FORMAT_FontStyle = 1,
    CHAR_FORMAT_FontInfo = 2,
    CHAR_FORMAT_FontName = 4,
    CHAR_FORMAT_ColorInfo = 8,
    CHAR_FORMAT_ParaFormat = 16,
    CHAR_FORMAT_FAMILY_Max = 2147483647L
} CHAR_FORMAT_FAMILY;


EXTERN_C const ENUMDESC s_enumdescCHAR_FORMAT_FAMILY;

typedef enum _LAYOUT_MOVE_UNIT
{
    LAYOUT_MOVE_UNIT_PreviousLine = 1,
    LAYOUT_MOVE_UNIT_NextLine = 2,
    LAYOUT_MOVE_UNIT_CurrentLineStart = 3,
    LAYOUT_MOVE_UNIT_CurrentLineEnd = 4,
    LAYOUT_MOVE_UNIT_NextLineStart = 5,
    LAYOUT_MOVE_UNIT_PreviousLineEnd = 6,
    LAYOUT_MOVE_UNIT_TopOfWindow = 7,
    LAYOUT_MOVE_UNIT_BottomOfWindow = 8,
    LAYOUT_MOVE_UNIT_OuterLineStart = 9,
    LAYOUT_MOVE_UNIT_OuterLineEnd = 10,
    LAYOUT_MOVE_UNIT_Max = 2147483647L
} LAYOUT_MOVE_UNIT;


EXTERN_C const ENUMDESC s_enumdescLAYOUT_MOVE_UNIT;

typedef enum _HIGHLIGHT_TYPE
{
    HIGHLIGHT_TYPE_None = 0x00000000,
    HIGHLIGHT_TYPE_Selected = 0x00000001,
    HIGHLIGHT_TYPE_Squiggle = 0x00000002,
    HIGHLIGHT_TYPE_ImeInput = 0x00000010,
    HIGHLIGHT_TYPE_ImeTargetConverted = 0x00000020,
    HIGHLIGHT_TYPE_ImeConverted = 0x00000030,
    HIGHLIGHT_TYPE_ImeTargetNotConverted = 0x00000040,
    HIGHLIGHT_TYPE_ImeInputError = 0x00000050,
    HIGHLIGHT_TYPE_ImeHangul = 0x00000060,
    HIGHLIGHT_TYPE_Max = 2147483647L
} HIGHLIGHT_TYPE;


EXTERN_C const ENUMDESC s_enumdescHIGHLIGHT_TYPE;

typedef enum _CARET_GRAVITY
{
    CARET_GRAVITY_NoChange = 0,
    CARET_GRAVITY_BeginningOfLine = 1,
    CARET_GRAVITY_EndOfLine = 2,
    CARET_GRAVITY_Max = 2147483647L
} CARET_GRAVITY;


EXTERN_C const ENUMDESC s_enumdescCARET_GRAVITY;

typedef enum _CARET_VISIBILITY
{
    CARET_TYPE_Hide = 0,
    CARET_TYPE_Show = 1,
    CARET_VISIBILITY_Max = 2147483647L
} CARET_VISIBILITY;


EXTERN_C const ENUMDESC s_enumdescCARET_VISIBILITY;

typedef enum _FOLLOW_UP_ACTION_FLAGS
{
    FOLLOW_UP_ACTION_None = 0x00000000,
    FOLLOW_UP_ACTION_ReBubbleEvent = 0x00000001,
    FOLLOW_UP_ACTION_OnClick = 0x00000020,
    FOLLOW_UP_ACTION_DragElement = 0x00000100,
    FOLLOW_UP_ACTION_FLAGS_Max = 2147483647L
} FOLLOW_UP_ACTION_FLAGS;


EXTERN_C const ENUMDESC s_enumdescFOLLOW_UP_ACTION_FLAGS;

typedef enum _SELECTION_NOTIFICATION
{
    SELECT_NOTIFY_TIMER_TICK = 0,
    SELECT_NOTIFY_EMPTY_SELECTION = 1,
    SELECT_NOTIFY_DOC_ENDED = 2,
    SELECT_NOTIFY_DOC_CHANGED = 3,
    SELECT_NOTIFY_DESTROY_SELECTION = 4,
    SELECT_NOTIFY_CARET_IN_CONTEXT = 5,
    SELECT_NOTIFY_EXIT_TREE = 6,
    SELECT_NOTIFY_LOSE_FOCUS_FRAME = 7,
    SELECT_NOTIFY_LOSE_FOCUS = 8,
    SELECT_NOTIFY_DESTROY_ALL_SELECTION = 9,
    SELECT_NOTIFY_DISABLE_IME = 10,
    SELECTION_NOTIFICATION_Max = 2147483647L
} SELECTION_NOTIFICATION;


EXTERN_C const ENUMDESC s_enumdescSELECTION_NOTIFICATION;

typedef enum _SELECTION_TYPE
{
    SELECTION_TYPE_None = 0,
    SELECTION_TYPE_Caret = 1,
    SELECTION_TYPE_Selection = 2,
    SELECTION_TYPE_Control = 3,
    SELECTION_TYPE_Auto = 4,
    SELECTION_TYPE_IME = 5,
    SELECTION_TYPE_Max = 2147483647L
} SELECTION_TYPE;


EXTERN_C const ENUMDESC s_enumdescSELECTION_TYPE;

typedef enum _POINTER_SCROLLPIN
{
    POINTER_SCROLLPIN_TopLeft = 0,
    POINTER_SCROLLPIN_BottomRight = 1,
    POINTER_SCROLLPIN_Minimal = 2,
    POINTER_SCROLLPIN_Max = 2147483647L
} POINTER_SCROLLPIN;


EXTERN_C const ENUMDESC s_enumdescPOINTER_SCROLLPIN;

typedef enum _COORD_SYSTEM
{
    COORD_SYSTEM_GLOBAL = 0,
    COORD_SYSTEM_PARENT = 1,
    COORD_SYSTEM_CONTAINER = 2,
    COORD_SYSTEM_CONTENT = 3,
    COORD_SYSTEM_Max = 2147483647L
} COORD_SYSTEM;


EXTERN_C const ENUMDESC s_enumdescCOORD_SYSTEM;

typedef enum _ADORNER_HTI
{
    ADORNER_HTI_NONE = 0,
    ADORNER_HTI_TOPBORDER = 1,
    ADORNER_HTI_LEFTBORDER = 2,
    ADORNER_HTI_BOTTOMBORDER = 3,
    ADORNER_HTI_RIGHTBORDER = 4,
    ADORNER_HTI_TOPLEFTHANDLE = 5,
    ADORNER_HTI_LEFTHANDLE = 6,
    ADORNER_HTI_TOPHANDLE = 7,
    ADORNER_HTI_BOTTOMLEFTHANDLE = 8,
    ADORNER_HTI_TOPRIGHTHANDLE = 9,
    ADORNER_HTI_BOTTOMHANDLE = 10,
    ADORNER_HTI_RIGHTHANDLE = 11,
    ADORNER_HTI_BOTTOMRIGHTHANDLE = 12,
    ADORNER_HTI_Max = 2147483647L
} ADORNER_HTI;


EXTERN_C const ENUMDESC s_enumdescADORNER_HTI;

typedef enum _CARET_DIRECTION
{
    CARET_DIRECTION_INDETERMINATE = 0,
    CARET_DIRECTION_SAME = 1,
    CARET_DIRECTION_BACKWARD = 2,
    CARET_DIRECTION_FORWARD = 3,
    CARET_DIRECTION_Max = 2147483647L
} CARET_DIRECTION;


EXTERN_C const ENUMDESC s_enumdescCARET_DIRECTION;

typedef struct _SelectionMessage
{
    UINT message;
    DWORD time;
    POINT pt;
    POINT ptContent;
    WPARAM wParam;
    LPARAM lParam;
    int characterCookie;
    DWORD_PTR elementCookie;
    BOOL fShift;
    BOOL fCtrl;
    BOOL fAlt;
    BOOL fFromCapture;
    BOOL fStopForward;
    BOOL fEmptySpace;
    HWND hwnd;
    LPARAM lResult;
} SelectionMessage;

typedef struct _HTMLPtrDispInfoRec
{
    DWORD dwSize;
    LONG lBaseline;
    LONG lXPosition;
    LONG lLineHeight;
    LONG lTextHeight;
    LONG lDescent;
    LONG lTextDescent;
    BOOL fRTLLine;
    BOOL fRTLFlow;
    BOOL fAligned;
    BOOL fHasNestedRunOwner;
} HTMLPtrDispInfoRec;

typedef struct _HTMLCharFormatData
{
    DWORD dwSize;
    WORD wFamilyFlags;
    BOOL fBold;
    BOOL fItalic;
    BOOL fUnderline;
    BOOL fOverline;
    BOOL fStrike;
    BOOL fSubScript;
    BOOL fSuperScript;
    BOOL fHasBgColor;
    BOOL fExplicitFace;
    WORD wWeight;
    WORD wFontSize;
    TCHAR szFont[32];
    DWORD dwTextColor;
    DWORD dwBackColor;
    BOOL fPre;
    BOOL fRTL;
} HTMLCharFormatData;


#ifndef __IHTMLViewServices_INTERFACE_DEFINED__

#define __IHTMLViewServices_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLViewServices;


MIDL_INTERFACE("3050f603-98b5-11cf-bb82-00aa00bdce0b")
IHTMLViewServices : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE MoveMarkupPointerToPoint(
            /* [in] */ POINT pt,/* [in] */ IMarkupPointer* pPointer,/* [in, out] */ BOOL* pfNotAtBOL,/* [in, out] */ BOOL* pfAtLogicalBOL,/* [in, out] */ BOOL* pfRightOfCp,/* [in] */ BOOL fScrollIntoView) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveMarkupPointerToPointEx(
            /* [in] */ POINT pt,/* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fGlobalCoordinates,/* [in, out] */ BOOL* pfNotAtBOL,/* [in, out] */ BOOL* pfAtLogicalBOL,/* [in, out] */ BOOL* pfRightOfCp,/* [in] */ BOOL fScrollIntoView) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveMarkupPointerToMessage(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ SelectionMessage* pMessage,/* [in, out] */ BOOL* pfNotAtBOL,/* [in, out] */ BOOL* pfAtLogicalBOL,/* [in, out] */ BOOL* pfRightOfCp,/* [in, out] */ BOOL* fValidTree,/* [in] */ BOOL fScrollIntoView,/* [in] */ IHTMLElement* pIContainerElement,/* [in, out] */ BOOL* pfSameLayout,/* [in] */ BOOL fHitTestEOL) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCharFormatInfo(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ WORD wFamilyFlags,/* [in, out] */ HTMLCharFormatData* pInfo) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetLineInfo(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fAtEndOfLine,/* [in, out] */ HTMLPtrDispInfoRec* pInfo) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsPointerBetweenLines(
            /* [in] */ IMarkupPointer* pPointer,/* [in, out] */ BOOL* pfBetweenLines) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementsInZOrder(
            /* [in] */ POINT pt,/* [in, out] */ IHTMLElement** rgElements,/* [in, out] */ DWORD* pCount) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetTopElement(
            /* [in] */ POINT pt,/* [out] */ IHTMLElement** ppElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveMarkupPointer(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ LAYOUT_MOVE_UNIT eUnit,/* [in] */ LONG lXCurReally,/* [out] */ BOOL* fAtEOL,/* [in, out] */ BOOL* pfAtLogicalBOL) = 0;

    virtual HRESULT STDMETHODCALLTYPE RegionFromMarkupPointers(
            /* [in] */ IMarkupPointer* pPointerStart,/* [in] */ IMarkupPointer* pPointerEnd,/* [out] */ DWORD_PTR* phrgn) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCurrentSelectionRenderingServices(
            /* [in, out] */ ISelectionRenderingServices** ppSelRenSvc) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCurrentSelectionSegmentList(
            /* [in, out] */ ISegmentList** ppSegment) = 0;

    virtual HRESULT STDMETHODCALLTYPE FireOnSelectStart(
            /* [in] */ IHTMLElement* pIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE FireCancelableEvent(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ LONG dispidMethod,/* [in] */ LONG dispidProp,/* [in] */ BSTR bstrEventType,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCaret(
            /* [out] */ IHTMLCaret** ppCaret) = 0;

    virtual HRESULT STDMETHODCALLTYPE ConvertVariantFromHTMLToTwips(
            /* [in, out] */ VARIANT* pvarargInOut) = 0;

    virtual HRESULT STDMETHODCALLTYPE ConvertVariantFromTwipsToHTML(
            /* [in, out] */ VARIANT* pvarargInOut) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsBlockElement(
            /* [in] */ IHTMLElement* pIElement,/* [out] */ BOOL* fResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsLayoutElement(
            /* [in] */ IHTMLElement* pIElement,/* [out] */ BOOL* fResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsContainerElement(
            /* [in] */ IHTMLElement* pIElement,/* [out] */ BOOL* fContainer,/* [out] */ BOOL* fHTML) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetFlowElement(
            /* [in] */ IMarkupPointer* pMarkup,/* [out] */ IHTMLElement** ppElem) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementFromCookie(
            /* [in] */ DWORD_PTR dwCookie,/* [out] */ IHTMLElement** ppElem) = 0;

    virtual HRESULT STDMETHODCALLTYPE InflateBlockElement(
            /* [in] */ IHTMLElement* pElem) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsInflatedBlockElement(
            /* [in] */ IHTMLElement* pElem,/* [in, out] */ BOOL* pfInflated) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsMultiLineFlowElement(
            /* [in] */ IHTMLElement* pElem,/* [out] */ BOOL* pfMulti) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementAttributeCount(
            /* [in] */ IHTMLElement* pElem,/* [out] */ UINT* pCount) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsEditableElement(
            /* [in] */ IHTMLElement* pElement,/* [out] */ BOOL* pfEditable) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetOuterContainer(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ IHTMLElement** ppIOuterElement,/* [in] */ BOOL fIgnoreOutermostContainer,/* [in] */ BOOL* pfHitContainer) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsNoScopeElement(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ BOOL* pfNoScope) = 0;

    virtual HRESULT STDMETHODCALLTYPE ShouldObjectHaveBorder(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ BOOL* pfDrawBorder) = 0;

    virtual HRESULT STDMETHODCALLTYPE DoTheDarnPasteHTML(
            /* [in] */ IMarkupPointer* pIPointerStart,/* [in] */ IMarkupPointer* pIPointerFinish,/* [in] */ HGLOBAL hGlobalHTML) = 0;

    virtual HRESULT STDMETHODCALLTYPE ConvertRTFToHTML(
            /* [in] */ LPOLESTR pszRtf,/* [in] */ HGLOBAL* phGlobalHTML) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetViewHWND(
            /* [in, out] */ HWND* pHWND) = 0;

    virtual HRESULT STDMETHODCALLTYPE ScrollElement(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ LONG lPercentToScroll,/* [out] */ POINT* ScrollDelta) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetScrollingElement(
            /* [in] */ IMarkupPointer* pPosition,/* [in] */ IHTMLElement* pBoundary,/* [out] */ IHTMLElement** ppElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE StartHTMLEditorDblClickTimer(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE StopHTMLEditorDblClickTimer(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE HTMLEditorTakeCapture(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE HTMLEditorReleaseCapture(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetHTMLEditorMouseMoveTimer(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetEditContext(
            /* [in] */ IHTMLElement* pIStartElement,/* [in] */ IHTMLElement** ppIEditThisElement,/* [in] */ IMarkupPointer* pIStart,/* [in] */ IMarkupPointer* pIEnd,/* [in] */ BOOL fDrillingIn,/* [in] */ BOOL* pfEditThisEditable,/* [in] */ BOOL* pfEditParentEditable,/* [in] */ BOOL* pfNoScopeElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE EnsureEditContext(
            /* [in] */ IMarkupPointer* pIMarkupPointer) = 0;

    virtual HRESULT STDMETHODCALLTYPE ScrollPointerIntoView(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fNotAtBOL,/* [in] */ POINTER_SCROLLPIN eScrollAmount) = 0;

    virtual HRESULT STDMETHODCALLTYPE ScrollPointIntoView(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ POINT* ptGlobal) = 0;

    virtual HRESULT STDMETHODCALLTYPE ArePointersInSameMarkup(
            /* [in] */ IMarkupPointer* pP1,/* [in] */ IMarkupPointer* pP2,/* [out] */ BOOL* pfSameMarkup) = 0;

    virtual HRESULT STDMETHODCALLTYPE DragElement(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ DWORD dwKeyState) = 0;

    virtual HRESULT STDMETHODCALLTYPE BecomeCurrent(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ SelectionMessage* pMessage) = 0;

    virtual HRESULT STDMETHODCALLTYPE TransformPoint(
            /* [in, out] */ POINT* pPoint,/* [in] */ COORD_SYSTEM eSource,/* [in] */ COORD_SYSTEM eDestination,/* [in] */ IHTMLElement* pIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementDragBounds(
            /* [in] */ IHTMLElement* pIElement,/* [in, out] */ RECT* pIElementDragBounds) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsElementLocked(
            /* [in] */ IHTMLElement* pIElement,/* [in, out] */ BOOL* pfLocked) = 0;

    virtual HRESULT STDMETHODCALLTYPE MakeParentCurrent(
            /* [in] */ IHTMLElement* pIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE FireOnBeforeEditFocus(
            /* [in] */ IHTMLElement* pINextActiveElement,/* [in, out] */ BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsEqualElement(
            /* [in] */ IHTMLElement* pIElement1,/* [in] */ IHTMLElement* pIElement2) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetOuterMostEditableElement(
            /* [in] */ IHTMLElement* pIEditableElement,/* [in, out] */ IHTMLElement** ppIOuterEditableElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsSite(
            /* [in] */ IHTMLElement* pElement,/* [out] */ BOOL* pfSite,/* [out] */ BOOL* pfText,/* [out] */ BOOL* pfMultiLine,/* [out] */ BOOL* pfScrollable) = 0;

    virtual HRESULT STDMETHODCALLTYPE QueryBreaks(
            /* [in] */ IMarkupPointer* pPointer,/* [out] */ DWORD* pdwBreaks,/* [in] */ BOOL fWantPendingBreak) = 0;

    virtual HRESULT STDMETHODCALLTYPE MergeDeletion(
            /* [in] */ IMarkupPointer* pPointer) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementForSelection(
            /* [in] */ IHTMLElement* pIElement,/* [out] */ IHTMLElement** ppIAdjustedElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsContainedBy(
            /* [in] */ IMarkupContainer* pIOuterContainer,/* [in] */ IMarkupContainer* pIInnerContainer) = 0;

    virtual HRESULT STDMETHODCALLTYPE CurrentScopeOrSlave(
            /* [in] */ IMarkupPointer* pPointer,/* [out] */ IHTMLElement** ppElemCurrent) = 0;

    virtual HRESULT STDMETHODCALLTYPE LeftOrSlave(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fMove,/* [out] */ MARKUP_CONTEXT_TYPE* pContext,/* [out] */ IHTMLElement** ppElement,/* [in, out] */ long* pcch,/* [out] */ OLECHAR* pchText) = 0;

    virtual HRESULT STDMETHODCALLTYPE RightOrSlave(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fMove,/* [out] */ MARKUP_CONTEXT_TYPE* pContext,/* [out] */ IHTMLElement** ppElement,/* [in, out] */ long* pcch,/* [out] */ OLECHAR* pchText) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveToContainerOrSlave(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ IMarkupContainer* pContainer,/* [in] */ BOOL fAtStart) = 0;

    virtual HRESULT STDMETHODCALLTYPE FindUrl(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd,/* [in] */ IMarkupPointer* pUrlStart,/* [in] */ IMarkupPointer* pUrlEnd) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsEnabled(
            /* [in] */ IHTMLElement* pIElement,/* [in, out] */ BOOL* pfEnabled) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementBlockDirection(
            /* [in] */ IHTMLElement* pElement,/* [out] */ BSTR* pbstrDir) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetElementBlockDirection(
            /* [in] */ IHTMLElement* pElement,/* [in] */ LONG eHTMLDir) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsBidiEnabled(
            /* [in, out] */ BOOL* pfEnabled) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetDocDirection(
            /* [in] */ LONG eHTMLDir) = 0;

    virtual HRESULT STDMETHODCALLTYPE AllowSelection(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ SelectionMessage* peMessage) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveWord(
            /* [in] */ IMarkupPointer* pPointerToMove,/* [in] */ MOVEUNIT_ACTION muAction,/* [in] */ IMarkupPointer* pLeftBoundary,/* [in] */ IMarkupPointer* pRightBoundary) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetAdjacencyForPoint(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ POINT ptGlobal,/* [in, out] */ ELEMENT_ADJACENCY *peAdjacent) = 0;

    virtual HRESULT STDMETHODCALLTYPE SaveSegmentsToClipboard(
            /* [in] */ ISegmentList* pSegmentList) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertMaximumText(
            /* [in] */ OLECHAR* pText,/* [in] */ LONG lLen,/* [in] */ IMarkupPointer* pIMarkupPointer) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsInsideURL(
            /* [in] */ IMarkupPointer* Start,/* [in] */ IMarkupPointer* pRight,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetDocHostUIHandler(
            /* [out] */ IDocHostUIHandler** ppIHandler) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetClientRect(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ COORD_SYSTEM eSource,/* [in, out] */ RECT* pRect) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetContentRect(
            /* [in] */ IHTMLElement* pIElement,/* [in] */ COORD_SYSTEM eSource,/* [in, out] */ RECT* pRect) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsElementSized(
            /* [in] */ IHTMLElement* pIElement,/* [in, out] */ BOOL* pfSized) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetLineDirection(
            /* [in] */ IMarkupPointer* pPointer,/* [in] */ BOOL fAtEndOfLine,/* [in, out] */ LONG* peHTMLDir) = 0;

};

#endif     /* __IHTMLViewServices_INTERFACE_DEFINED__ */


#ifndef __IHTMLCaret_INTERFACE_DEFINED__

#define __IHTMLCaret_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLCaret;


MIDL_INTERFACE("3050f604-98b5-11cf-bb82-00aa00bdce0b")
IHTMLCaret : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE MoveCaretToPointer(
            /* [in] */ IMarkupPointer* pIMarkupPointerCaret,/* [in] */ BOOL fNotAtBOL,/* [in] */ BOOL fScrollIntoView,/* [in] */ CARET_DIRECTION eDir) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveCaretToPointerEx(
            /* [in] */ IMarkupPointer* pIMarkupPointerCaret,/* [in] */ BOOL fAtEOL,/* [in] */ BOOL fVisible,/* [in] */ BOOL fScrollIntoView,/* [in] */ CARET_DIRECTION eDir) = 0;

    virtual HRESULT STDMETHODCALLTYPE MovePointerToCaret(
            /* [in] */ IMarkupPointer* pIMarkupPointerCaret) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsVisible(
            /* [in, out] */ BOOL* pIsVisible) = 0;

    virtual HRESULT STDMETHODCALLTYPE Show(
            /* [in] */ BOOL fScrollIntoView) = 0;

    virtual HRESULT STDMETHODCALLTYPE Hide(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertText(
            /* [in] */ OLECHAR* pText,/* [in] */ LONG lLen) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertMarkup(
            /* [in] */ OLECHAR* pMarkup) = 0;

    virtual HRESULT STDMETHODCALLTYPE ScrollIntoView(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementContainer(
            /* [out] */ IHTMLElement** ppElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetLocation(
            /* [in, out] */ POINT* pPoint,/* [in] */ BOOL fTranslate) = 0;

    virtual HRESULT STDMETHODCALLTYPE UpdateCaret(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetOffset(
            /* [in] */ LONG lXDelta,/* [in] */ LONG lYDelta,/* [in] */ LONG lHDelta) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetOffset(
            /* [in, out] */ LONG* plXDelta,/* [in, out] */ LONG* plYDelta,/* [in, out] */ LONG* plHDelta) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetNotAtBOL(
            /* [in, out] */ BOOL* pfNotAtBOL) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetNotAtBOL(
            /* [in] */ BOOL fNotAtBOL) = 0;

    virtual HRESULT STDMETHODCALLTYPE LoseFocus(
            ) = 0;

};

#endif     /* __IHTMLCaret_INTERFACE_DEFINED__ */


#ifndef __ISegmentList_INTERFACE_DEFINED__

#define __ISegmentList_INTERFACE_DEFINED__

EXTERN_C const IID IID_ISegmentList;


MIDL_INTERFACE("3050f605-98b5-11cf-bb82-00aa00bdce0b")
ISegmentList : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE MovePointersToSegment(
            /* [in] */ int iSegmentIndex,/* [in] */ IMarkupPointer* pIPointerStart,/* [in] */ IMarkupPointer* pIPointerEnd) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetSegmentCount(
            /* [in, out] */ int* piSegmentCount,/* [in, out] */ SELECTION_TYPE* eType) = 0;

};

#endif     /* __ISegmentList_INTERFACE_DEFINED__ */


#ifndef __ISelectionRenderingServices_INTERFACE_DEFINED__

#define __ISelectionRenderingServices_INTERFACE_DEFINED__

EXTERN_C const IID IID_ISelectionRenderingServices;


MIDL_INTERFACE("3050f606-98b5-11cf-bb82-00aa00bdce0b")
ISelectionRenderingServices : public ISegmentList
{
public:
    virtual HRESULT STDMETHODCALLTYPE AddSegment(
            /* [in] */ IMarkupPointer* pIMarkupPointerStart,/* [in] */ IMarkupPointer* pIMarkupPointerEnd,/* [in] */ HIGHLIGHT_TYPE HighlightType,/* [in, out] */ int* piSegmentIndex) = 0;

    virtual HRESULT STDMETHODCALLTYPE AddElementSegment(
            /* [in] */ IHTMLElement* pIElement,/* [in, out] */ int* piSegmentIndex) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveSegmentToPointers(
            /* [in] */ int iSegmentIndex,/* [in] */ IMarkupPointer* pIMarkupPointerStart,/* [in] */ IMarkupPointer* pIMarkupPointerEnd,/* [in] */ HIGHLIGHT_TYPE HighlightType) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementSegment(
            /* [in] */ int iSegmentIndex,/* [in, out] */ IHTMLElement** ppIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetElementSegment(
            /* [in] */ int iSegmentIndex,/* [in] */ IHTMLElement* pIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE ClearSegment(
            /* [in] */ int iSegmentIndex,/* [in] */ BOOL fInvalidate) = 0;

    virtual HRESULT STDMETHODCALLTYPE ClearSegments(
            /* [in] */ BOOL fInvalidate) = 0;

    virtual HRESULT STDMETHODCALLTYPE ClearElementSegments(
            ) = 0;

};

#endif     /* __ISelectionRenderingServices_INTERFACE_DEFINED__ */


#ifndef __IHTMLEditor_INTERFACE_DEFINED__

#define __IHTMLEditor_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLEditor;


MIDL_INTERFACE("3050f608-98b5-11cf-bb82-00aa00bdce0b")
IHTMLEditor : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE HandleMessage(
            /* [in, out] */ SelectionMessage* pSelectionMessage,/* [in, out] */ DWORD* pdwFollowUpActionFlag) = 0;

    virtual HRESULT STDMETHODCALLTYPE Initialize(
            /* [in] */ IUnknown* pIDocument,/* [in] */ IUnknown* pIView) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetEditContext(
            /* [in] */ BOOL fEditable,/* [in] */ BOOL fSetSelection,/* [in] */ BOOL fParentEditable,/* [in] */ IMarkupPointer* pStartPointer,/* [in] */ IMarkupPointer* pEndPointer,/* [in] */ BOOL fNoScopeElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetSelectionType(
            /* [in, out] */ SELECTION_TYPE* eSelectionType) = 0;

    virtual HRESULT STDMETHODCALLTYPE Notify(
            /* [in] */ SELECTION_NOTIFICATION eSelectionNotification,/* [in] */ IUnknown* pUnknown,/* [in, out] */ DWORD* pdwFollowUpActionFlag,/* [in] */ DWORD dword) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsPointerInSelection(
            /* [in] */ IMarkupPointer* pIPointer,/* [in, out] */ BOOL* pfPointerInSelection,/* [in] */ POINT* pptGlobal,/* [in] */ IHTMLElement* pIElementOver) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetRangeCommandTarget(
            /* [in] */ IUnknown* pContext,/* [in, out] */ IUnknown** ppUnkTarget) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementToTabFrom(
            /* [in] */ BOOL fForward,/* [in, out] */ IHTMLElement** ppElement,/* [in, out] */ BOOL* pfFindNext) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsElementSiteSelected(
            /* [in] */ IHTMLElement* pIElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsEditContextUIActive(
            ) = 0;

};

#endif     /* __IHTMLEditor_INTERFACE_DEFINED__ */


#ifndef __IHTMLEditingServices_INTERFACE_DEFINED__

#define __IHTMLEditingServices_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLEditingServices;


MIDL_INTERFACE("3050f609-98b5-11cf-bb82-00aa00bdce0b")
IHTMLEditingServices : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Delete(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd,/* [in] */ BOOL fAdjustPointers) = 0;

    virtual HRESULT STDMETHODCALLTYPE Paste(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd,/* [in] */ BSTR bstrText) = 0;

    virtual HRESULT STDMETHODCALLTYPE PasteFromClipboard(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd,/* [in] */ IDataObject* pDO) = 0;

    virtual HRESULT STDMETHODCALLTYPE Select(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd,/* [in] */ SELECTION_TYPE eType,/* [in] */ DWORD* pdwFollowUpActionFlag) = 0;

    virtual HRESULT STDMETHODCALLTYPE LaunderSpaces(
            /* [in] */ IMarkupPointer* pStart,/* [in] */ IMarkupPointer* pEnd) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertSanitizedText(
            /* [in] */ IMarkupPointer* InsertHere,/* [in] */ OLECHAR* pstrText,/* [in] */ BOOL fDataBinding) = 0;

    virtual HRESULT STDMETHODCALLTYPE AdjustPointerForInsert(
            /* [in] */ IMarkupPointer* pWhereIThinkIAm,/* [in] */ BOOL fFurtherInDocument,/* [in] */ BOOL fNotAtBol,/* [in] */ IMarkupPointer* pConstraintStart,/* [in] */ IMarkupPointer* pConstraintEnd) = 0;

    virtual HRESULT STDMETHODCALLTYPE FindSiteSelectableElement(
            /* [in] */ IMarkupPointer* pPointerStart,/* [in] */ IMarkupPointer* pPointerEnd,/* [in] */ IHTMLElement** ppIHTMLElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsElementSiteSelectable(
            /* [in] */ IHTMLElement* pIHTMLElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsElementUIActivatable(
            /* [in] */ IHTMLElement* pIHTMLElement) = 0;

};

#endif     /* __IHTMLEditingServices_INTERFACE_DEFINED__ */


#endif /*__internal_h__*/

