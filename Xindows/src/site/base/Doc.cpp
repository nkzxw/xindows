
#include "stdafx.h"
#include "Doc.h"

#include <IDispIds.h>
#include <shdeprecated.h>

#define _cxx_
#include "../gen/document.hdl"

CCriticalSection    g_csJitting;
BYTE                g_bUSPJitState  = JIT_OK;   // For UniScribe JIT (USP10.DLL)
BYTE                g_bJGJitState   = JIT_OK;   // JG ART library for AOL (JG*.DLL)

UINT CDocument::_g_msgHtmlGetobject = 0;

#define EXPANDOS_DEFAULT    TRUE

#define SCROLLPERCENT       125

//+------------------------------------------------------------------------
//
// Utility function to force layout on all windows in the thread
//
//-------------------------------------------------------------------------
void OnSettingsChangeAllDocs(BOOL fNeedLayout)
{
    int i;

    for(i=0; i<TLS(_paryDoc).Size(); i++)
    {
        TLS(_paryDoc)[i]->OnSettingsChange(fNeedLayout);
    }
}

const CServer::CLASSDESC CDocument::s_classdesc =
{
    {                                            // _classdescBase
        &CLSID_HTMLDocument,                     // _pclsid
        s_acpi,                                  // _pcpi
        SERVERDESC_CREATE_UNDOMGR |              // _dwFlags
        SERVERDESC_ACTIVATEONDRAG |
        SERVERDESC_SUPPORT_DRAG_DROP |
        SERVERDESC_HAS_MENU |
        SERVERDESC_HAS_TOOLBAR,
        &IID_IHTMLDocument2,                     // _piidDispinterface
        &s_apHdlDescs,                           // _apHdlDesc
    },
};

BOOL        g_fDocClassInitialized = FALSE;
extern HRESULT InitFormatCache(THREADSTATE*);

static ATOM s_atomFormOverlayWndClass;

SIZE        g_sizeDragMin;
int         g_iDragScrollDelay;
SIZE        g_sizeDragScrollInset;
int         g_iDragDelay;
int         g_iDragScrollInterval;

static char s_achWindows[] = "windows"; // Localization: Do not localize

typedef CStackPtrAry<CTreeNode*, 32> NodeArray;

HRESULT OldCompare(CMarkupPointer* p1, CMarkupPointer* p2, int* piResult)
{
    HRESULT hr = S_OK;
    BOOL    fResult;

    Assert(piResult);

    fResult = p1->IsEqualTo(p2);

    if(fResult)
    {
        *piResult = 0;
        goto Cleanup;
    }

    fResult = p1->IsLeftOf(p2);

    *piResult = fResult ? -1 : 1;

Cleanup:
    RRETURN(hr);
}

static inline BOOL IsEqualTo(CMarkupPointer* p1, CMarkupPointer* p2)
{
    BOOL fEqual = p1->IsEqualTo(p2);
    return fEqual;
}

static inline int OldCompare(CMarkupPointer* p1, CMarkupPointer* p2)
{
    int result;
    OldCompare(p1, p2, &result);
    return result;
}


//+---------------------------------------------------------------
//
//  Member:     CDefaultElement
//
//---------------------------------------------------------------
const CElement::CLASSDESC CDefaultElement::s_classdesc =
{
    {
        NULL,                   // _pclsid
        NULL,                   // _pcpi
        0,                      // _dwFlags
        NULL,                   // _piidDispinterface
        NULL
    },
    NULL,
    NULL,                       // _paccelsDesign
    NULL                        // _paccelsRun
};

CDefaultElement::CDefaultElement(CDocument* pDoc) : CElement(ETAG_DEFAULT, pDoc)
{
}



void CMessageToSelectionMessage(const CMessage* pMessage, SelectionMessage* pSelMessage)
{
    pSelMessage->message    = pMessage->message;
    pSelMessage->time       = pMessage->time;
    pSelMessage->pt         = pMessage->pt;

    if(pMessage->pNodeHit)
    {
        CLayout* pLayout = pMessage->pNodeHit->Doc()->GetLayoutForSelection(pMessage->pNodeHit);
        if(pLayout)
        {
            // If we have multiple text nodes, pMessage->ptContent can be relative
            // to the text display node. But most of the selection code works with
            // layout relative coordinates. So, translate the pt to layout relative.
            pSelMessage->ptContent = pMessage->pt;
            pLayout->TransformPoint((CPoint*)&pSelMessage->ptContent, COORDSYS_GLOBAL, COORDSYS_CONTENT);
        }
    }
    else
    {
        pSelMessage->ptContent = pMessage->ptContent;
    }
    pSelMessage->wParam             = pMessage->wParam;
    pSelMessage->lParam             = pMessage->lParam;
    pSelMessage->elementCookie      = (DWORD_PTR)pMessage->pNodeHit;
    pSelMessage->characterCookie    = pMessage->resultsHitTest._cpHit;
    pSelMessage->fCtrl              = pMessage->dwKeyState & FCONTROL;
    pSelMessage->fShift             = pMessage->dwKeyState & FSHIFT;
    pSelMessage->fAlt               = pMessage->dwKeyState & FALT;
    pSelMessage->fStopForward       = pMessage->fStopForward;
    pSelMessage->fFromCapture       = FALSE;
    pSelMessage->fEmptySpace        = pMessage->resultsHitTest._fWantArrow;    
    pSelMessage->hwnd               = pMessage->hwnd;
    pSelMessage->lResult            = pMessage->lresult;
}

void SelectionMessageToCMessage(const SelectionMessage* pSelMessage, CMessage* pMessage)
{
    pMessage->message               = pSelMessage->message;
    pMessage->time                  = pSelMessage->time;
    pMessage->pt                    = pSelMessage->pt;
    pMessage->ptContent             = pSelMessage->ptContent;
    pMessage->wParam                = pSelMessage->wParam;
    pMessage->lParam                = pSelMessage->lParam;
    pMessage->resultsHitTest._cpHit = pSelMessage->characterCookie;

    pMessage->fStopForward          = pSelMessage->fStopForward;
    pMessage->fStopForward          = pSelMessage->fStopForward;
    pMessage->hwnd                  = pSelMessage->hwnd;
    pMessage->lresult               = pSelMessage->lResult;
    pMessage->resultsHitTest._fWantArrow = pSelMessage->fEmptySpace;

    pMessage->SetNodeHit((CTreeNode*)pSelMessage->elementCookie);

    if(pSelMessage->fCtrl)
    {
        pMessage->dwKeyState |= FCONTROL;
    }
    if(pSelMessage->fShift)
    {
        pMessage->dwKeyState |= FSHIFT;
    }
    if(pSelMessage->fAlt)
    {
        pMessage->dwKeyState |= FALT;
    }

    // We assume that this is called only by the selection handler, so we
    // always set this flag to TRUE.
    pMessage->fSelectionHMCalled = TRUE;
}



BEGIN_TEAROFF_TABLE(CDocument, IServiceProvider)
    TEAROFF_METHOD(CDocument, &QueryService, queryservice, (REFGUID rsid, REFIID iid, void** ppvObj))
END_TEAROFF_TABLE()

BEGIN_TEAROFF_TABLE(CDocument, IMarkupServices)
    TEAROFF_METHOD(CDocument, &CreateMarkupPointer, createmakruppointer, (IMarkupPointer** ppPointer))
    TEAROFF_METHOD(CDocument, &CreateMarkupContainer, createmarkupcontainer, (IMarkupContainer** ppMarkupContainer))
    TEAROFF_METHOD(CDocument, &CreateElement, createelement, (ELEMENT_TAG_ID, OLECHAR*, IHTMLElement**))
    TEAROFF_METHOD(CDocument, &InsertElement, insertelement, (IHTMLElement* pElementInsert, IMarkupPointer* pPointerStart, IMarkupPointer* pPointerFinish))
    TEAROFF_METHOD(CDocument, &RemoveElement, removeelement, (IHTMLElement* pElementRemove))
    TEAROFF_METHOD(CDocument, &Remove, remove, (IMarkupPointer*, IMarkupPointer*))
    TEAROFF_METHOD(CDocument, &Copy, copy, (IMarkupPointer*, IMarkupPointer*, IMarkupPointer*))
    TEAROFF_METHOD(CDocument, &Move, move, (IMarkupPointer*, IMarkupPointer*, IMarkupPointer*))
    TEAROFF_METHOD(CDocument, &InsertText, inserttext, (OLECHAR*, long, IMarkupPointer*))
    TEAROFF_METHOD(CDocument, &IsScopedElement, isscopedelement, (IHTMLElement*, BOOL*))
    TEAROFF_METHOD(CDocument, &GetElementTagId, getelementtagid, (IHTMLElement*, ELEMENT_TAG_ID*))
    TEAROFF_METHOD(CDocument, &GetTagIDForName, gettagidforname, (BSTR, ELEMENT_TAG_ID*))
    TEAROFF_METHOD(CDocument, &GetNameForTagID, getnamefortagid, (ELEMENT_TAG_ID, BSTR*))
    TEAROFF_METHOD(CDocument, &BeginUndoUnit, beginundounit, (OLECHAR*))
    TEAROFF_METHOD(CDocument, &EndUndoUnit, beginundounit, ())
END_TEAROFF_TABLE()

BEGIN_TEAROFF_TABLE(CDocument, IHTMLViewServices)
    TEAROFF_METHOD(CDocument, &MoveMarkupPointerToPoint, movemarkuppointertopoint, (POINT pt, IMarkupPointer* pPointer, BOOL* pfNotAtBOL, BOOL* pfAtLogicalBOL, BOOL* pfRightOfCp, BOOL fScrollIntoView))
    TEAROFF_METHOD(CDocument, &MoveMarkupPointerToPointEx, movemarkuppointertopointex, (POINT pt, IMarkupPointer* pPointer, BOOL fGlobalCoordinates, BOOL* pfNotAtBOL, BOOL* pfAtLogicalBOL, BOOL* pfRightOfCp, BOOL fScrollIntoView))
    TEAROFF_METHOD(CDocument, &MoveMarkupPointerToMessage, movemarkuppointertomessage, (IMarkupPointer* pPointer, SelectionMessage* pMessage, BOOL* pfNotAtBOL, BOOL* pfAtLogicalBOL, BOOL* pfRightOfCp, BOOL* pfValidTree, BOOL fScrollIntoView, IHTMLElement* pIContainerElement, BOOL* pfSameLayout, BOOL fHitTestEOL))
    TEAROFF_METHOD(CDocument, &GetCharFormatInfo, getcharformatinfo, (IMarkupPointer* pPointer, WORD family, HTMLCharFormatData* pInfo))
    TEAROFF_METHOD(CDocument, &GetLineInfo, getlineinfo, (IMarkupPointer* pPointer, BOOL fAtEndOfLine, HTMLPtrDispInfoRec* pInfo))
    TEAROFF_METHOD(CDocument, &IsPointerBetweenLines, ispointerbetweenlines,  (IMarkupPointer* pPointer, BOOL* pfBetweenLines))
    TEAROFF_METHOD(CDocument, &GetElementsInZOrder, getelementsinzorder, (POINT pt, IHTMLElement** rgElements, DWORD* pCount))
    TEAROFF_METHOD(CDocument, &GetTopElement, gettopelement, (POINT pt, IHTMLElement** ppElement))
    TEAROFF_METHOD(CDocument, &MoveMarkupPointer, movemarkuppointer, (IMarkupPointer* pPointer, LAYOUT_MOVE_UNIT eUnit, LONG lXCurReally, BOOL* fNotAtBOL, BOOL* fAtLogicalBOL))
    TEAROFF_METHOD(CDocument, &RegionFromMarkupPointers, regionfrommarkuppointers, (IMarkupPointer* pPointerStart, IMarkupPointer* pPointerEnd, HRGN* phrgn))
    TEAROFF_METHOD(CDocument, &GetCurrentSelectionRenderingServices, getcurrentselectionrenderingservices, ( ISelectionRenderingServices** ppSelRenSvc))
    TEAROFF_METHOD(CDocument, &GetCurrentSelectionSegmentList, getcurrentsegmentlist , (ISegmentList** ppSegmentList))
    TEAROFF_METHOD(CDocument, &FireOnSelectStart, fireonselectstart , (IHTMLElement* pIElement))
    TEAROFF_METHOD(CDocument, &FireCancelableEvent, firecancelableevent, (IHTMLElement* pIElement, LONG dispidMethod, LONG dispidProp, BSTR bstrEventType, BOOL* pfResult))
    TEAROFF_METHOD(CDocument, &GetCaret, getcaret, (IHTMLCaret** ppCaret))
    TEAROFF_METHOD(CDocument, &ConvertVariantFromHTMLToTwips, convertvariantfromhtmltotwips, (VARIANT* pvar))
    TEAROFF_METHOD(CDocument, &ConvertVariantFromTwipsToHTML, convertvariantfromtwipstohtml, (VARIANT* pvar))
    TEAROFF_METHOD(CDocument, &IsBlockElement, isblockelement, (IHTMLElement* pIElement, BOOL* pfResult))
    TEAROFF_METHOD(CDocument, &IsLayoutElement, islayoutelement, (IHTMLElement* pIElement, BOOL* pfResult))
    TEAROFF_METHOD(CDocument, &IsContainerElement, iscontainerelement, (IHTMLElement* pIElement, BOOL* pfContainer, BOOL* pfHTML))
    TEAROFF_METHOD(CDocument, &GetFlowElement, getflowelement, (IMarkupPointer* pIPointer, IHTMLElement** ppIElement))
    TEAROFF_METHOD(CDocument, &GetElementFromCookie, getelementfromcookie, (void* elementCookie, IHTMLElement** ppIElement))
    TEAROFF_METHOD(CDocument, &InflateBlockElement, inflateblockelement, (IHTMLElement* pElem ))
    TEAROFF_METHOD(CDocument, &IsInflatedBlockElement, isinflatedblockelement, (IHTMLElement* pElem , BOOL* pfIsInflated))
    TEAROFF_METHOD(CDocument, &IsMultiLineFlowElement, ismultilineflowelement, (IHTMLElement* pIElement, BOOL* fMultiLine))
    TEAROFF_METHOD(CDocument, &GetElementAttributeCount, getelementattributecount, (IHTMLElement* pIElement, UINT* pCount))
    TEAROFF_METHOD(CDocument, &IsEditableElement, iseditableelement, (IHTMLElement* pIElement, BOOL* pfResult))
    TEAROFF_METHOD(CDocument, &GetOuterContainer, getoutercontainer, (IHTMLElement* pIElement, IHTMLElement** ppIOuterElement, BOOL fIgnoreOutermostContainer, BOOL* pfHitContainer))
    TEAROFF_METHOD(CDocument, &IsNoScopeElement, isnoscopeelement, (IHTMLElement* pIElement, BOOL* pfResult))  
    TEAROFF_METHOD(CDocument, &ShouldObjectHaveBorder, shouldobjecthaveborder, (IHTMLElement* pIElement, BOOL* pfDrawBorder))
    TEAROFF_METHOD(CDocument, &DoTheDarnPasteHTML, dothedarnpastehtml, (IMarkupPointer*, IMarkupPointer*, HGLOBAL))
    TEAROFF_METHOD(CDocument, &ConvertRTFToHTML, convertrtftohtml, (LPOLESTR pszRtf, HGLOBAL* phglobalHTML))
    TEAROFF_METHOD(CDocument, &GetViewHWND, getviewhwnd, (HWND* pHWND))
    TEAROFF_METHOD(CDocument, &ScrollElement, scrollelement, (IHTMLElement* pIElement, LONG lPercentToScroll, POINT* pScrollDelta))
    TEAROFF_METHOD(CDocument, &GetScrollingElement, getscrollingelement, (IMarkupPointer* pPosition, IHTMLElement* pBoundary, IHTMLElement** ppElement))          
    TEAROFF_METHOD(CDocument, &StartHTMLEditorDblClickTimer, starthtmleditordblclicktimer, ())
    TEAROFF_METHOD(CDocument, &StopHTMLEditorDblClickTimer, stophtmleditordblclicktimer, ())    
    TEAROFF_METHOD(CDocument, &HTMLEditorTakeCapture, htmleditortakecapture, ())
    TEAROFF_METHOD(CDocument, &HTMLEditorReleaseCapture, htmleditorreleasecapture, ())
    TEAROFF_METHOD(CDocument, &SetHTMLEditorMouseMoveTimer, sethtmleditormousemovetimer, ())       

    TEAROFF_METHOD(CDocument, &GetEditContext, geteditcontext, ( 
                   IHTMLElement*    pIStartElement, 
                   IHTMLElement**   ppIEditThisElement, 
                   IMarkupPointer*  pIStart, 
                   IMarkupPointer*  pIEnd, 
                   BOOL             fDrillingIn,  
                   BOOL*            pfEditThisEditable ,
                   BOOL*            pfEditParentEditable ,
                   BOOL*            pfNoScopeElement))         

    TEAROFF_METHOD(CDocument, &EnsureEditContext, ensureeditcontext, (IMarkupPointer* pIPointer))                                                            
    TEAROFF_METHOD(CDocument, &ScrollPointerIntoView, scrollpointerintoview, (IMarkupPointer* pPointer, BOOL fNotAtBOL, POINTER_SCROLLPIN eScrollAmount))
    TEAROFF_METHOD(CDocument, &ScrollPointIntoView, scrollpointintoview, (IHTMLElement* pIElement , POINT* ptGlobal))

    TEAROFF_METHOD(CDocument, &ArePointersInSameMarkup, arepointersinsamemarkup, (IMarkupPointer* pP1, IMarkupPointer* pP2, BOOL* pfSameMarkup))
    TEAROFF_METHOD(CDocument, &DragElement, dragelement, (IHTMLElement* pIElement, DWORD dwKeyState))
    TEAROFF_METHOD(CDocument, &BecomeCurrent, becomecurrent, (
        IHTMLElement* pIElement,
        SelectionMessage* pSelMessage))
    TEAROFF_METHOD(CDocument, &TransformPoint, transformpoint, (
        POINT*          pPoint, 
        COORD_SYSTEM    eSource, 
        COORD_SYSTEM    eDestination, 
        IHTMLElement*   pIElement))
    TEAROFF_METHOD(CDocument, &GetElementDragBounds, getelementdragbounds, (
        IHTMLElement*   pIElement,
        RECT*           pIElementDragBounds))                                                                         
    TEAROFF_METHOD(CDocument, &IsElementLocked, iselementlocked, (
        IHTMLElement*   pIElement,
        BOOL*           pfLocked))
    TEAROFF_METHOD(CDocument, &MakeParentCurrent, makeparentcurrent, (IHTMLElement* pIElement))                                                             
    TEAROFF_METHOD(CDocument, &FireOnBeforeEditFocus, fireonbeforeeditfocus, (IHTMLElement* pINextActiveElem, BOOL* pfRet)) 
    TEAROFF_METHOD(CDocument, &IsEqualElement , isequalelement, (IHTMLElement* pIElement1, IHTMLElement* pIElement2))
    TEAROFF_METHOD(CDocument, &GetOuterMostEditableElement, getoutermosteditableelement , (
        IHTMLElement*   pIEditableElement, 
        IHTMLElement**  ppIOuterEditableElement))       
    TEAROFF_METHOD(CDocument, &IsSite, issite, (IHTMLElement* pElement, BOOL* pfSite, BOOL* pfText, BOOL* pfMultiLine, BOOL* pfScrollable))
    TEAROFF_METHOD(CDocument, &QueryBreaks, querybreaks, (IMarkupPointer* pPointer, DWORD* pdwBreaks, BOOL fWantPendingBreak))
    TEAROFF_METHOD(CDocument, &MergeDeletion, mergedeletion, (IMarkupPointer* pPointer ))
    TEAROFF_METHOD(CDocument, &GetElementForSelection , getelementforselection, (IHTMLElement* pElement, IHTMLElement** ppISiteSelectableElement))
    TEAROFF_METHOD(CDocument, &IsContainedBy, iscontainedby, (IMarkupContainer* pIOuterContainer, IMarkupContainer* pIInnerContainer))
    TEAROFF_METHOD(CDocument, &CurrentScopeOrSlave, currentscopoeorslave, (IMarkupPointer* pPointer, IHTMLElement** ppElemCurrent))
    TEAROFF_METHOD(CDocument, &LeftOrSlave, leftorslave, (IMarkupPointer* pPointer, BOOL fMove, MARKUP_CONTEXT_TYPE* pContext, IHTMLElement** ppElement, long* pcch, OLECHAR* pchText))
    TEAROFF_METHOD(CDocument, &RightOrSlave, rightorslave, (IMarkupPointer* pPointer, BOOL fMove, MARKUP_CONTEXT_TYPE* pContext, IHTMLElement** ppElement, long* pcch, OLECHAR* pchText))
    TEAROFF_METHOD(CDocument, &MoveToContainerOrSlave, movetocontainerorslave, (IMarkupPointer* pPointer, IMarkupContainer* pContainer, BOOL fAtStart))
    TEAROFF_METHOD(CDocument, &FindUrl, findurl, (IMarkupPointer* pStart, IMarkupPointer* pEnd, IMarkupPointer* pUrlStart, IMarkupPointer* pUrlEnd))
    TEAROFF_METHOD(CDocument, &IsEnabled, isenabled, (IHTMLElement* pIElement, BOOL* pfEnabled))
    TEAROFF_METHOD(CDocument, &GetElementBlockDirection, getelementblockdirection, (IHTMLElement* pElement, BSTR* pbstrDir))
    TEAROFF_METHOD(CDocument, &SetElementBlockDirection, setelementblockdirection, (IHTMLElement* pElement, LONG eHTMLDir))
    TEAROFF_METHOD(CDocument, &IsBidiEnabled, isbidienabled, (BOOL* pfEnabled))
    TEAROFF_METHOD(CDocument, &SetDocDirection, setdocdirection, (LONG eHTMLDir))
    TEAROFF_METHOD(CDocument, &AllowSelection, allowselection, (IHTMLElement* pIElement, SelectionMessage* peMessage)) 
    TEAROFF_METHOD(CDocument, &MoveWord, moveword, (IMarkupPointer* pPointerToMove, MOVEUNIT_ACTION  muAction, IMarkupPointer* pLeftBoundary, IMarkupPointer* pRightBoundary))
    TEAROFF_METHOD(CDocument, &GetAdjacencyForPoint, getadjacencyforpoint, (IHTMLElement* pIElement, POINT ptGlobal, ELEMENT_ADJACENCY *peAdjacent  ))    
    TEAROFF_METHOD(CDocument, &SaveSegmentsToClipboard, savesegmentstoclipboard, (ISegmentList* pSegmentList))
    TEAROFF_METHOD(CDocument, &InsertMaximumText, insertmaximumtext, (OLECHAR* pstrText, LONG cch, IMarkupPointer* pMarkupPointer))
    TEAROFF_METHOD(CDocument, &IsInsideURL , isinsideurl, (IMarkupPointer*, IMarkupPointer*, BOOL*))
    TEAROFF_METHOD(CDocument, &GetDocHostUIHandler , getdochostuihandler , (IDocHostUIHandler**))
    TEAROFF_METHOD(CDocument, &GetClientRect , getclientrect, (IHTMLElement* pIElement, COORD_SYSTEM eSource, RECT* pRect))    
    TEAROFF_METHOD(CDocument, &GetContentRect , getcontentrect, (IHTMLElement* pIElement, COORD_SYSTEM eSource, RECT* pRect))        
    TEAROFF_METHOD(CDocument, &IsElementSized, iselementsized, (IHTMLElement* pIElement, BOOL* pfSized))    
    TEAROFF_METHOD(CDocument, &GetLineDirection, getlinedirection, (IMarkupPointer* pPointer, BOOL fAtEndOfLine, long* peHTMLDir))
END_TEAROFF_TABLE()



//+---------------------------------------------------------------
//
//  Member:     InitDocClass
//
//  Synopsis:   Initializes the CDocument class
//
//  Returns:    TRUE iff the class could be initialized successfully
//
//  Notes:      This method initializes the verb tables in the
//              class descriptor.  Called by the LibMain
//              of the DLL.
//
//---------------------------------------------------------------
HRESULT InitDocClass()
{
    if(!g_fDocClassInitialized)
    {
        LOCK_GLOBALS;

        // If another thread completed initialization while this thread waited
        // for the global lock, immediately return
        if(g_fDocClassInitialized)
        {
            return S_OK;
        }

        int i;

        i = GetProfileIntA(s_achWindows, "DragScrollInset", DD_DEFSCROLLINSET);
        g_sizeDragScrollInset.cx = i;
        g_sizeDragScrollInset.cy = i;
        g_iDragScrollDelay = GetProfileIntA(s_achWindows, "DragScrollDelay", DD_DEFSCROLLDELAY);
        g_iDragDelay = GetProfileIntA(s_achWindows, "DragDelay", DD_DEFDRAGDELAY);
        g_iDragScrollInterval = GetProfileIntA(s_achWindows, "DragScrollInterval", DD_DEFSCROLLINTERVAL);

        g_fDocClassInitialized = TRUE;
    }

    return S_OK;
}

//+----------------------------------------------------------------------------
//
//  Member:     get_bgColor, IOmDocument
//
//  Synopsis: defers to body get_bgcolor
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_bgColor(VARIANT* p)
{
    CBodyElement*   pBody;
    HRESULT         hr;
    CColorValue     Val;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = GetBodyElement(&pBody);
    if(FAILED(hr))
    {
        goto Cleanup;
    }
    Assert(hr==S_FALSE || pBody!=NULL);

    if(hr != S_OK)
    {
        Val = GetAAbgColor();
    }
    else
    {
        Val = pBody->GetFirstBranch()->GetCascadedbackgroundColor();
    }

    // Allocates and returns BSTR that represents the color as #RRGGBB
    V_VT(p) = VT_BSTR;
    hr = CColorValue::FormatAsPound6Str(&(V_BSTR(p)), Val.IsDefined()?Val.GetColorRef():_pOptionSettings->crBack());

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_bgColor, IOmDocument
//
//  Synopsis: defers to body put_bgcolor
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_bgColor(VARIANT p)
{
    IHTMLBodyElement* pBody = NULL;
    HRESULT            hr;

    GetBodyElement(&pBody);

    if(!pBody)
    {
        hr = s_propdescCDocbgColor.b.SetColorProperty(p, this, (CVoid*)(void*)(GetAttrArray()));
        if(hr)
        {
            goto Cleanup;
        }

        // Force a repaint and transition to a load-state where
        // we are allowed to redraw.
        GetView()->Invalidate((CRect*)NULL, TRUE);
    }
    else
    {
        hr = pBody->put_bgColor(p);
        ReleaseInterface(pBody);

        GetView()->EnsureView(LAYOUT_DEFEREVENTS|LAYOUT_SYNCHRONOUSPAINT);

        if(hr ==S_OK)
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_bgColor, 0);
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_fgColor, IOmDocument
//
//  Synopsis: defers to body get_text
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_fgColor(VARIANT* p)
{
    CBodyElement*   pBody;
    HRESULT         hr;
    CColorValue     Val;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = GetBodyElement(&pBody);
    if(FAILED(hr))
    {
        goto Cleanup;
    }
    Assert(hr==S_FALSE || pBody!=NULL);

    if(hr != S_OK)
    {
        // 没有获取到Body说明该文档属于FrameSet
        Val = GetAAfgColor();
    }
    else
    {
        Val = pBody->GetFirstBranch()->GetCascadedcolor();
    }

    // Allocates and returns BSTR that represents the color as #RRGGBB
    V_VT(p) = VT_BSTR;
    hr = CColorValue::FormatAsPound6Str(&(V_BSTR(p)), Val.IsDefined()?Val.GetColorRef():_pOptionSettings->crText());

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_fgColor, IOmDocument
//
//  Synopsis: defers to body put_text
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_fgColor(VARIANT p)
{
    IHTMLBodyElement*   pBody = NULL;
    HRESULT             hr;

    GetBodyElement(&pBody);
    if(!pBody)
    {
        hr = s_propdescCDocfgColor.b.SetColorProperty(p, this, (CVoid*)(void*)(GetAttrArray()));
    }
    else
    {
        hr = pBody->put_text(p);
        ReleaseInterface(pBody);
        if(hr == S_OK)
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_fgColor, 0);
        }
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_linkColor, IOmDocument
//
//  Synopsis: defers to body get_link
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_linkColor(VARIANT* p)
{
    CBodyElement*   pBody;
    HRESULT         hr;
    CColorValue     Val;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = GetBodyElement(&pBody);
    if(FAILED(hr))
    {
        goto Cleanup;
    }
    Assert(hr==S_FALSE || pBody!=NULL);

    if(hr != S_OK)
    {
        // 没有获取到Body说明该文档属于FrameSet
        Val = GetAAlinkColor();
    }
    else
    {
        Val = pBody->GetAAlink();
    }

    // Allocates and returns BSTR that represents the color as #RRGGBB
    V_VT(p) = VT_BSTR;
    hr = CColorValue::FormatAsPound6Str(&(V_BSTR(p)), Val.IsDefined()?Val.GetColorRef():_pOptionSettings->crAnchor());

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_linkColor, IOmDocument
//
//  Synopsis: defers to body put_link
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_linkColor(VARIANT p)
{
    IHTMLBodyElement* pBody = NULL;
    HRESULT hr;

    GetBodyElement(&pBody);
    if(!pBody)
    {
        hr = s_propdescCDoclinkColor.b.SetColorProperty(p, this, (CVoid*)(void*)(GetAttrArray()));
    }
    else
    {
        hr = pBody->put_link(p);
        ReleaseInterface(pBody);
        if(hr == S_OK)
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_linkColor, 0);
        }
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_alinkColor, IOmDocument
//
//  Synopsis: defers to body get_aLink
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_alinkColor(VARIANT* p)
{
    CBodyElement*   pBody;
    HRESULT         hr;
    CColorValue     Val;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = GetBodyElement(&pBody);
    if(FAILED(hr))
    {
        goto Cleanup;
    }
    Assert(hr==S_FALSE || pBody!=NULL);

    if(hr != S_OK)
    {
        // 没有获取到Body说明该文档属于FrameSet
        Val = GetAAalinkColor();
    }
    else
    {
        Val = pBody->GetAAaLink();
    }

    // Allocates and returns BSTR that represents the color as #RRGGBB
    V_VT(p) = VT_BSTR;
    hr = CColorValue::FormatAsPound6Str(&(V_BSTR(p)), Val.IsDefined()?Val.GetColorRef():_pOptionSettings->crAnchor());

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_alinkColor, IOmDocument
//
//  Synopsis: defers to body put_aLink
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_alinkColor(VARIANT p)
{
    IHTMLBodyElement* pBody = NULL;
    HRESULT hr;

    GetBodyElement(&pBody);
    if(!pBody)
    {
        hr = s_propdescCDocalinkColor.b.SetColorProperty(p, this, (CVoid*)(void*)(GetAttrArray()));
    }
    else
    {
        hr = pBody->put_aLink(p);
        ReleaseInterface(pBody);
        if(hr == S_OK)
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_alinkColor, 0);
        }
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_vlinkColor, IOmDocument
//
//  Synopsis: defers to body get_vLink
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_vlinkColor(VARIANT* p)
{
    CBodyElement*   pBody;
    HRESULT         hr;
    CColorValue     Val;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = GetBodyElement(&pBody);
    if(FAILED(hr))
    {
        goto Cleanup;
    }

    Assert(hr==S_FALSE || pBody!=NULL);

    if(hr != S_OK)
    {
        // 没有获取到Body说明该文档属于FrameSet
        Val = GetAAvlinkColor();
    }
    else
    {
        Val = pBody->GetAAvLink();
    }

    // Allocates and returns BSTR that represents the color as #RRGGBB
    V_VT(p) = VT_BSTR;
    hr = CColorValue::FormatAsPound6Str(&(V_BSTR(p)), Val.IsDefined()?Val.GetColorRef():_pOptionSettings->crAnchorVisited());

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     put_vlinkColor, IOmDocument
//
//  Synopsis: defers to body put_vLink
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_vlinkColor(VARIANT p)
{
    IHTMLBodyElement* pBody = NULL;
    HRESULT hr;

    GetBodyElement(&pBody);
    if(!pBody)
    {
        // 没有获取到Body说明该文档属于FrameSet
        hr = s_propdescCDocvlinkColor.b.SetColorProperty(p, this, (CVoid*)(void*)(GetAttrArray()));
    }
    else
    {
        hr = pBody->put_vLink(p);
        ReleaseInterface(pBody);
        if(hr == S_OK)
        {
            Fire_PropertyChangeHelper(DISPID_CDoc_vlinkColor, 0);
        }
    }

    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Members:    Get/SetDir
//
//  Synopsis:   Functions to get at the document's direction from the object
//              model.
//
//----------------------------------------------------------------------------
HRESULT CDocument::get_dir(BSTR* p)
{
    HRESULT hr = S_OK;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = s_propdescCDocdir.b.GetEnumStringProperty(p, this, (CVoid*)(void*)(GetAttrArray()));

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CDocument::put_dir(BSTR v)
{
    HRESULT hr = S_OK;

    hr = s_propdescCDocdir.b.SetEnumStringProperty(v, this, (CVoid*)(void*)(GetAttrArray()));

    // send the property change message to the body. These depend upon
    // being in edit mode or not.
    CBodyElement* pBody;

    hr = GetBodyElement(&pBody);
    if(!hr)
    {
        pBody->OnPropertyChange(DISPID_A_DIR, ELEMCHNG_CLEARCACHES|ELEMCHNG_REMEASUREALLCONTENTS);
    }

    RRETURN(SetErrorInfo(hr));
}

//-----------------------------------------------------------------------------
//
//  Function:   CDocument::FaultInUSP
//
//  Synopsis:   Async callback to JIT install UniScribe (USP10.DLL)
//
//  Arguments:  DWORD (CDocument *)  The current doc from which the hWnd can be gotten
//
//  Returns:    none
//
//-----------------------------------------------------------------------------
void CDocument::FaultInUSP(DWORD_PTR dwContext)
{
    HRESULT hr;
    uCLSSPEC classpec;
    CString cstrGUID;
    ULONG cDie = _cDie;
    BOOL fRefresh = FALSE;

    PrivateAddRef();
    g_csJitting.Enter();

    Assert(g_bUSPJitState == JIT_PENDING);

    // Close the door. We only want one of these running.
    g_bUSPJitState = JIT_IN_PROGRESS;

    // Set the GUID for USP10 so JIT can lookup the feature
    cstrGUID.Set(TEXT("{b1ad7c1e-c217-11d1-b367-00c04fb9fbed}"));

    // setup the classpec
    classpec.tyspec = TYSPEC_CLSID;
    hr = CLSIDFromString((BSTR)cstrGUID, &classpec.tagged_union.clsid);

    if(hr == S_OK)
    {
        hr = FaultInIEFeatureHelper(GetHWND(), &classpec, NULL, 0);
    }

    // if we succeeded or the document navigated away (process was killed in
    // CDocument::UnloadContents) set state to JIT_OK so IOD can be attempted
    // again without having to restart the host.
    if(hr == S_OK)
    {
        g_bUSPJitState = JIT_OK;
        if(cDie == _cDie)
        {
            fRefresh = TRUE;
        }
    }
    else
    {
        // The user cancelled or aborted. Don't ask for this again
        // during this session.
        g_bUSPJitState = JIT_DONT_ASK;
    }

    g_csJitting.Leave();

    // refresh the view if we have just installed.
    if(fRefresh)
    {
        _view.EnsureView(LAYOUT_SYNCHRONOUS|LAYOUT_FORCE);
    }

    PrivateRelease();
}

//-----------------------------------------------------------------------------
//
//  Function:   CDocument::FaultInJG
//
//  Synopsis:   Async callback to JIT install JG ART library for AOL (JG*.DLL)
//
//-----------------------------------------------------------------------------
void CDocument::FaultInJG(DWORD_PTR dwContext)
{
    HRESULT hr;
    uCLSSPEC classpec;
    CString cstrGUID;

    if(g_bJGJitState != JIT_PENDING)
    {
        return;
    }

    // Close the door. We only want one of these running.
    g_bJGJitState = JIT_IN_PROGRESS;

    HWND hWnd = GetHWND();
    // Set the GUID for JG*.dll so JIT can lookup the feature
    cstrGUID.Set(_T("{47f67d00-9e55-11d1-baef-00c04fc2d130}"));

    // setup the classpec
    classpec.tyspec = TYSPEC_CLSID;
    hr = CLSIDFromString((BSTR)cstrGUID, &classpec.tagged_union.clsid);

    if(hr == S_OK)
    {
        hr = FaultInIEFeatureHelper(hWnd, &classpec, NULL, 0);
    }

    // if we succeeded or the document navigated away (process was killed in
    // CDocument::UnloadContents) set state to JIT_OK so IOD can be attempted
    // again without having to restart the host.
    if(hr == S_OK)
    {
        g_bJGJitState = JIT_OK;
    }
    else
    {
        // The user cancelled or aborted. Don't ask for this again
        // during this session.
        g_bJGJitState = JIT_DONT_ASK;
    }
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::CDocument, protected
//
//  Synopsis:   Constructor for the CDocument class
//
//  Arguments:  [pUnkOuter] -- the controlling unknown or NULL if we are
//                             not being created as part of an aggregate
//
//  Notes:      This is the first part of a two-stage construction process.
//              The second part is in the Init method.  Use the static
//              Create method to properly instantiate an CDocument object
//
//---------------------------------------------------------------
CDocument::CDocument() : CServer()
{
    Assert(g_fDocClassInitialized);

    _dwTID = GetCurrentThreadId();

    // Start things off with the cache invalid.
    // The tree version is initialized to 1 so that other versions can be
    // initialized to 0 to automatically sugest invalidness.
    __lDocTreeVersion = 1;
    __lDocContentsVersion = 1;

    // Initialize the document to a default size

    // NB (cthrash) Start at an absolute pixel size.  Compute the himetric
    // based on that, rather than the reverse.  This reduces rounding errors.
    {
        SIZEL sizelDefault;

        sizelDefault.cx = MulDivQuick(100, 2540, _afxGlobalData._sizePixelsPerInch.cx);
        sizelDefault.cy = MulDivQuick(100, 2540, _afxGlobalData._sizePixelsPerInch.cy);

        _dci.CTransform::Init(sizelDefault);
        _dci._pDoc = this;
    }

    Assert(!_pElemCurrent);
    Assert(!_pElemUIActive);

    _view.Initialize(this);

    _fExpando = EXPANDOS_DEFAULT;

    _fEnableInteraction = TRUE;

    _sBaselineFont = BASELINEFONTDEFAULT;

    // Internationalization
    _codepage = _codepageFamily = _afxGlobalData._cpDefault;

    _cstrUrl.Set(_afxGlobalData._szModuleFolder);

    // Append to thread doc array
    TLS(_paryDoc).Append(this);

    MemSetName((this, "CDocument SSN=%d", _ulSSN));

    // Register the window message (if not registered)
    if(_g_msgHtmlGetobject == 0)
    {
        _g_msgHtmlGetobject = RegisterWindowMessage(MSGNAME_WM_XINDOWS_GETOBJECT);
        Assert(_g_msgHtmlGetobject != 0);
    }

    _fNeedTabOut = FALSE;

    _fRegionCollection = FALSE; // default no need to build region collection

    _pCaret = NULL;

    // reset the accessibility object, we don't need it until we're asked
    _fPassivateCalled = FALSE;

    // wlw note: I want to make it dialog behavior
    _dwFlagsHostInfo |= /*DOCHOSTUIFLAG_DIALOG|*/DOCHOSTUIFLAG_SCROLL_NO|DOCHOSTUIFLAG_NO3DBORDER;
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::~CDocument
//
//  Synopsis:   Destructor for the CDocument class
//
//---------------------------------------------------------------
CDocument::~CDocument()
{
    Assert(!_pNodeGotButtonDown);

    // Remove Doc from Thread state array
    TLS(_paryDoc).DeleteByValue(this);

    // In case any extra expandos were added after CDocument::Passivate (see bug 55425)
    _AtomTable.Free();

#ifdef _DEBUG
    // Make sure there is nothing in the image context cache
    {
        URLIMGCTX*  purlimgctx = _aryUrlImgCtx;
        LONG        curlimgctx = _aryUrlImgCtx.Size();
        LONG        iurlimgctx;

        for(iurlimgctx=0; iurlimgctx<curlimgctx; ++iurlimgctx,++purlimgctx)
        {
            if(purlimgctx->ulRefs > 0)
            {
                break;
            }
        }

        AssertSz(iurlimgctx==curlimgctx, "Image context cache leak");
    }
#endif
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::PrivateQueryInterface
//
//  Synopsis:   QueryInterface on our private unknown
//
//---------------------------------------------------------------
HRESULT CDocument::PrivateQueryInterface(REFIID iid, void** ppv)
{
    *ppv = NULL;

    switch(iid.Data1)
    {
        QI_TEAROFF(this, IServiceProvider, NULL)
        QI_INHERITS(this, IHTMLDocument)
        QI_INHERITS(this, IHTMLDocument2)
        QI_TEAROFF(this, IObjectIdentity, NULL)
        QI_TEAROFF(this, IMarkupServices, NULL)
        QI_TEAROFF(this, IHTMLViewServices, NULL)
    default:
        {
            void*       pvTearoff = NULL;
            const void* apfnTearoff = NULL;

            if(DispNonDualDIID(iid))
            {
                *ppv = (IHTMLDocument2*)this;
                break;
            }
            else if(IsEqualIID(iid, CLSID_HTMLDocument))
            {
                *ppv = this;
                return S_OK;
            }
            else if(IsEqualIID(iid, CLSID_CMarkup) && _pPrimaryMarkup)
            {
                *ppv = _pPrimaryMarkup;
                return S_OK;
            }
            else if(IsEqualIID(iid, IID_IMarkupContainer) && _pPrimaryMarkup)
            {
                pvTearoff = _pPrimaryMarkup;
                apfnTearoff = (const void*)CMarkup::s_apfnIMarkupContainer;
            }
            else if(IsEqualIID(iid, IID_ISelectionRenderingServices) && _pPrimaryMarkup)
            {
                pvTearoff = _pPrimaryMarkup;
                apfnTearoff = (const void*)CMarkup::s_apfnISelectionRenderingServices;
            }
            else if(IsEqualIID(iid, IID_ISegmentList) && _pPrimaryMarkup)
            {
                pvTearoff = _pPrimaryMarkup;
                apfnTearoff = (const void*)CMarkup::s_apfnISelectionRenderingServices;
            }

            // Create the tearoff if we need to
            if(pvTearoff)
            {
                HRESULT hr;

                Assert(apfnTearoff);

                hr = CreateTearOffThunk(
                    pvTearoff,
                    apfnTearoff, 
                    NULL, 
                    ppv, 
                    (IUnknown*)(IPrivateUnknown*)this, 
                    *(void**)(IUnknown*)(IPrivateUnknown*)this,
                    QI_MASK,
                    NULL,
                    NULL);
                if (hr)
                    RRETURN(hr);
            }
            else
            {
                RRETURN(CServer::PrivateQueryInterface(iid, ppv));
            }
        }
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    ((IUnknown*)*ppv)->AddRef();

    return S_OK;
}

//+-------------------------------------------------------------------------
//
//  Method:     CDocument::Init
//
//  Synopsis:    Second phase of construction
//
//--------------------------------------------------------------------------
HRESULT CDocument::Init()
{
    HRESULT hr;
    THREADSTATE* pts = GetThreadState();

    hr = super::Init();

    if(hr)
    {
        goto Cleanup;
    }

    // Create the default site (not to be confused with a root site)
    Assert(!_pElementDefault);

    _pElementDefault = new CDefaultElement(this);

    if(!_pElementDefault)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    _pElemCurrent = _pElementDefault;

    _icfDefault = -1;

    // Make a root site to keep everybody happy until we're loaded
    hr = CreateRoot();

    if(hr)
    {
        goto Cleanup;
    }

    // Initialize format caches
    if(!TLS(_pCharFormatCache))
    {
        InitFormatCache(pts);
    }

    _dwStylesheetDownloadingCookie = 1;

Cleanup:
    RRETURN(hr);
}

void CDocument::SetLoadfFromPrefs()
{
    // Read in the preferences, if we don't already have them
    if(_pOptionSettings == NULL)
    {
        UpdateFromRegistry();
        Assert(_pOptionSettings);
    }

    _dwLoadf = DLCTL_DLIMAGES | (_pOptionSettings->fShowVideos?DLCTL_VIDEOS:0) | (DLCTL_BGSOUNDS);

    /*if(_pHostPeerFactory)
    {
        SetPeersPossible();
    } wlw note*/

    if(_dwFlagsHostInfo & DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8)
    {
        _dwLoadf |= DLCTL_URL_ENCODING_DISABLE_UTF8;
    }
    else if(_dwFlagsHostInfo & DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8)
    {
        _dwLoadf |= DLCTL_URL_ENCODING_ENABLE_UTF8;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDoc::InitNew, IPersistStream
//
//  Synopsis:   Initialize ole state of control. Overriden to allow this to
//              occur after already initialized or loaded (yuck!). From an
//              OLE point of view this is totally illegal. Required for
//              MSHTML classic compat.
//
//----------------------------------------------------------------------------
HRESULT CDocument::InitNew()
{
    BYTE*               pbRequestHeaders  = NULL;
    ULONG               cbRequestHeaders  = 0;
    DWORD               dwRefresh         = 0;
    CODEPAGE            codepage          = 0;
    CODEPAGE            codepageURL       = 0;
    DWORD               dwBindf           = 0;
    DWORD               dwDocBindf        = 0;
    DWORD               dwDownf           = 0;
    BOOL                fPrecreated       = FALSE;
    HRESULT             hr;

    // Don't allow re-entrant loading of the document
    if(TestLock(FORMLOCK_LOADING))
    {
        return E_PENDING;
    }

    CLock Lock(this, FORMLOCK_LOADING);

    // Can't be sure who might want the services of MLANG; Ensure it now
    // before a non-OleInitialize-ing thread (the download thread, for
    // example) decides it needs it.
    EnsureMultiLanguage();

    // Load preferences from the registry if we haven't already
    SetLoadfFromPrefs();

    // At this point, we should be defoliated.  Start up a new tree.
    Assert(PrimaryRoot());

    _ulProgressPos  = 0;
    _ulProgressMax  = 0;
    _fShownProgPos  = FALSE;
    _fShownProgMax  = FALSE;
    _fShownProgText = FALSE;

    Assert(!!_cstrUrl);
    MemSetName((this, "CDoc SSN=%d URL=%ls", _ulSSN, (LPTSTR)_cstrUrl));

    if(_dwLoadf & DLCTL_SILENT)
    {
        dwBindf |= BINDF_SILENTOPERATION | BINDF_NO_UI;
    }

    if(_dwLoadf & DLCTL_RESYNCHRONIZE)
    {
        dwBindf |= BINDF_RESYNCHRONIZE;
    }

    if(_dwLoadf & DLCTL_PRAGMA_NO_CACHE)
    {
        dwBindf |= BINDF_PRAGMA_NO_CACHE;
    }

    dwRefresh = IncrementLcl();

    dwDownf = GetDefaultColorMode();

    if(!_pOptionSettings->fSmartDithering)
    {
        dwDownf |= DWNF_FORCEDITHER;
    }

    if(_dwLoadf & DLCTL_DOWNLOADONLY)
    {
        dwDownf |= DWNF_DOWNLOADONLY;
    }

    // use CP_AUTO when so requested
    if(IsCpAutoDetect())
    {
        codepage = CP_AUTO;
    }
    else
    {
        // Get it from the option settings default
        codepage = NavigatableCodePage(_pOptionSettings->codepageDefault);
    }

    codepageURL = _codepage;

    SwitchCodePage(codepage);

    // The default document direction is LTR. Only set this if the document is RTL
    _fRTLDocDirection = FALSE;

    _pDwnDoc = new CDwnDoc;

    if(_pDwnDoc == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    dwDocBindf = dwBindf;

    _pDwnDoc->SetDoc(this);
    _pDwnDoc->SetBindf(dwBindf);
    _pDwnDoc->SetDocBindf(dwDocBindf);
    _pDwnDoc->SetDownf(dwDownf);
    _pDwnDoc->SetLoadf(_dwLoadf);
    _pDwnDoc->SetRefresh(dwRefresh);
    _pDwnDoc->SetDocCodePage(NavigatableCodePage(codepage));
    _pDwnDoc->SetURLCodePage(NavigatableCodePage(codepageURL));
    _pDwnDoc->TakeRequestHeaders(&pbRequestHeaders, &cbRequestHeaders);
    _pDwnDoc->SetDownloadNotify(_pDownloadNotify);

    if(_pOptionSettings->fHaveAcceptLanguage)
    {
        hr = _pDwnDoc->SetAcceptLanguage(_pOptionSettings->cstrLang);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    MemFree(pbRequestHeaders);

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::DragEnter, IDropTarget
//
//  Synopsis:   Setup for possible drop
//
//----------------------------------------------------------------------------
STDMETHODIMP CDocument::DragEnter(
                IDataObject*    pDataObj,
                DWORD           grfKeyState,
                POINTL          ptlScreen,
                DWORD*          pdwEffect)
{
    _fRightBtnDrag = (grfKeyState&MK_RBUTTON) ? 1 : 0;

    HRESULT hr = CServer::DragEnter(pDataObj, grfKeyState, ptlScreen, pdwEffect);
    if(hr)
    {
        RRETURN(hr);
    }

    BeginUndoUnit(_T("Drag & Drop"));

    Assert(!_pDragDropTargetInfo);
    _pDragDropTargetInfo = new CDragDropTargetInfo(this);
    if(!_pDragDropTargetInfo)
    {
        return E_OUTOFMEMORY;
    }

    RRETURN(DragOver(grfKeyState, ptlScreen, pdwEffect));
}

CLayout* GetLayoutForDragDrop(CElement* pElement)
{
    CMarkup* pMarkup;
    CLayout* pLayout = pElement->GetUpdatedLayout();

    // Handle slave markups
    if(!pLayout)
    {
        pMarkup = pElement->GetMarkup();
        if(pMarkup->Master())
        {
            pLayout = pMarkup->Master()->GetUpdatedLayout();
        }
    }
    return pLayout;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::DragOver, IDropTarget
//
//  Synopsis:   Handle scrolling and dispatch to site.
//
//----------------------------------------------------------------------------
STDMETHODIMP CDocument::DragOver(DWORD grfKeyState, POINTL ptScreen, DWORD* pdwEffect)
{
    HRESULT     hr              = S_OK;
    POINT       pt;
    CTreeNode*  pNodeElement    = NULL;
    DWORD       dwLoopEffect;
    CMessage    msg;
    DWORD       dwScrollDir     = 0;
    BOOL        fRedrawFeedback;
    BOOL        fRet;
    BOOL        fDragEnter;

    Assert(_pInPlace->_pDataObj);

    if(!_pDragDropTargetInfo)
    {
        AssertSz(0, "DragDropTargetInfo is NULL. Possible cause - DragOver called without calling DragEnter");
        *pdwEffect = DROPEFFECT_NONE;
        return S_OK;
    }

    pt = *(POINT*)&ptScreen;
    ScreenToClient(_pInPlace->_hwnd, &pt);

    msg.pt = pt;
    HitTestPoint(&msg, &pNodeElement, HT_IGNOREINDRAG);
    if(!pNodeElement)
    {
        goto Cleanup;
    }

    // Point to the content
    if(pNodeElement->Element()->HasSlaveMarkupPtr())
    {
        pNodeElement = pNodeElement->Element()->GetSlaveMarkupPtr()->FirstElement()->GetFirstBranch();
    }

    // We should not change pdwEffect unless DragEnter succeeded, and always call
    // DragEnter with the original effect
    dwLoopEffect = *pdwEffect;

    if(pNodeElement->Element() != _pDragDropTargetInfo->_pElementHit)
    {
        fDragEnter = TRUE;
        _pDragDropTargetInfo->_pElementHit = pNodeElement->Element();
    }
    else
    {
        fDragEnter = FALSE;
    }

    while(pNodeElement)
    {
        CLayout* pLayout = GetLayoutForDragDrop(pNodeElement->Element());

        pNodeElement->NodeAddRef();

        if(pNodeElement->Element() == _pDragDropTargetInfo->_pElementTarget)
        {
            fRet = pNodeElement->Element()->Fire_ondragHelper(
                0,
                DISPID_EVMETH_ONDRAGOVER,
                DISPID_EVPROP_ONDRAGOVER,
                _T("dragover"),
                pdwEffect);

            if(pNodeElement->IsDead())
            {
                pNodeElement->NodeRelease();
                goto Cleanup;
            }
            if(fRet)
            {
                if(pLayout)
                {
                    hr = pLayout->DragOver(grfKeyState, ptScreen, pdwEffect);
                }
                else
                {
                    *pdwEffect = DROPEFFECT_NONE;
                }
            }
            break;
        }

        if(fDragEnter)
        {
            fRet = pNodeElement->Element()->Fire_ondragHelper(
                0,
                DISPID_EVMETH_ONDRAGENTER,
                DISPID_EVPROP_ONDRAGENTER,
                _T("dragenter"),
                &dwLoopEffect);

            if(pNodeElement->IsDead())
            {
                pNodeElement->NodeRelease();
                goto Cleanup;
            }
            if(fRet)
            {
                if(pLayout)
                {
                    hr = pLayout->DragEnter(_pInPlace->_pDataObj, grfKeyState, ptScreen, &dwLoopEffect);
                    if(hr != S_FALSE)
                    {
                        break;
                    }
                }
            }
            else
            {
                break;
            }
        }

        // BUGBUG: what if pNodeElement goes away on this first release?!?
        pNodeElement->NodeRelease();
        pNodeElement = pNodeElement->GetUpdatedParentLayoutNode();
        dwLoopEffect = *pdwEffect;
    }
    if(pNodeElement && DifferentScope(pNodeElement, _pDragDropTargetInfo->_pElementTarget))
    {
        if(_pDragDropTargetInfo->_pElementTarget && _pDragDropTargetInfo->_pElementTarget->GetFirstBranch())
        {
            CLayout* pLayout;

            _pDragDropTargetInfo->_pElementTarget->Fire_ondragleave();

            pLayout = GetLayoutForDragDrop(_pDragDropTargetInfo->_pElementTarget);
            if(pLayout)
            {
                pLayout->DragLeave();
            }
        }

        _pDragDropTargetInfo->_pElementTarget = pNodeElement->Element();
        *pdwEffect = dwLoopEffect;
    }

    if(NULL == _pDragDropTargetInfo->_pElementTarget)
    {
        *pdwEffect = DROPEFFECT_NONE;
    }


    // Find the site to scroll and the direction, if any

    // BUGBUG (MohanB) GetSite/GetLayout may not appropriate in case of <AREA>.
    // Here, we really need to use the corresponding image, not the site that
    // contains the area/map in the source. We should try to fix this in the new
    // layout code.        
    if(pNodeElement)
    {
        // BUGBUG: what if pNodeElement goes away on this first release!?!?
        pNodeElement->NodeRelease();
        if(pNodeElement->IsDead())
        {
            goto Cleanup;
        }
    }

    {
        Assert(_view.IsActive());

        CDispScroller* pDispScroller = _view.HitScrollInset((CPoint*)&pt, &dwScrollDir);

        if(pDispScroller)
        {
            Assert(dwScrollDir);
            *pdwEffect |= DROPEFFECT_SCROLL;

            if(_pDragDropTargetInfo->_pDispScroller == pDispScroller)
            {
                if(EdUtil::IsTimePassed(_pDragDropTargetInfo->_uTimeScroll))
                {
                    CSize sizeOffset;
                    CSize sizePercent(0, 0);
                    CSize sizeDelta;
                    CRect rc;

                    // Hide drag feedback while scrolling
                    fRedrawFeedback = _fDragFeedbackVis;
                    if(_fDragFeedbackVis)
                    {
                        CLayout* pLayout;

                        Assert(_pDragDropTargetInfo->_pElementTarget);

                        pLayout = GetLayoutForDragDrop(_pDragDropTargetInfo->_pElementTarget);
                        if(pLayout)
                        {
                            pLayout->DrawDragFeedback();
                        }
                    }

                    // open display tree for scrolling
                    Verify(_view.OpenView());

                    pDispScroller->GetClientRect(&rc, CLIENTRECT_CONTENT);
                    pDispScroller->GetScrollOffset(&sizeOffset);

                    if(dwScrollDir & SCROLL_LEFT)
                    {
                        sizePercent.cx = -SCROLLPERCENT;
                    }
                    else if(dwScrollDir & SCROLL_RIGHT)
                    {
                        sizePercent.cx = SCROLLPERCENT;
                    }
                    if(dwScrollDir & SCROLL_UP)
                    {
                        sizePercent.cy = -SCROLLPERCENT;
                    }
                    else if(dwScrollDir & SCROLL_DOWN)
                    {
                        sizePercent.cy = SCROLLPERCENT;
                    }

                    sizeDelta.cx = (sizePercent.cx ? (rc.Width()*sizePercent.cx)/1000L : 0);
                    sizeDelta.cy = (sizePercent.cy ? (rc.Height()*sizePercent.cy)/1000L : 0);

                    sizeOffset += sizeDelta;
                    sizeOffset.Max(_afxGlobalData._Zero.size);

                    if(_pCaret)
                    {
                        _pCaret->BeginPaint();
                    }

                    pDispScroller->SetScrollOffset(sizeOffset, TRUE);

                    if(_pCaret)
                    {
                        _pCaret->EndPaint();
                    }

                    //  Ensure all deferred calls are executed
                    _view.EndDeferred();

                    // Show drag feedback after scrolling is finished
                    if(fRedrawFeedback)
                    {
                        CLayout* pLayout;

                        Assert(_pDragDropTargetInfo->_pElementTarget);

                        pLayout = GetLayoutForDragDrop(_pDragDropTargetInfo->_pElementTarget);
                        if(pLayout)
                        {
                            pLayout->DrawDragFeedback();
                        }
                    }

                    // Wait a while before scrolling again
                    _pDragDropTargetInfo->_uTimeScroll = EdUtil::NextEventTime(g_iDragScrollInterval/2);
                }
            }
            else
            {
                _pDragDropTargetInfo->_pDispScroller = pDispScroller;
                _pDragDropTargetInfo->_uTimeScroll = EdUtil::NextEventTime(g_iDragScrollDelay);
            }
        }
        else
        {
            _pDragDropTargetInfo->_pDispScroller = NULL;
        }
    }

Cleanup:
    // S_FALSE from DragOver() does not make sense for OLE. Also, since we
    // call DragOver() from within DragEnter(), we DO NOT want to return
    // S_FALSE because that would result in OLE passing on DragOver() to
    // the parent drop target.
    if(hr == S_FALSE)
    {
        hr = S_OK;
    }

    RRETURN(hr);
}

STDMETHODIMP CDocument::DragLeave(BOOL fDrop)
{
    if(_pDragDropTargetInfo)
    {
        if(_pDragDropTargetInfo->_pElementTarget)
        {
            CLayout* pLayout;

            //BUGBUG (t-jeffg) This makes sure that when dragging out of the source window
            //the dragged item will be deleted.  Will be removed when _pdraginfo is dealt
            //with properly on leave.  (Anandra)
            _fSlowClick = FALSE;

            if(!fDrop)
            {
                _pDragDropTargetInfo->_pElementTarget->Fire_ondragleave();
            }

            pLayout = GetLayoutForDragDrop(_pDragDropTargetInfo->_pElementTarget);
            if(pLayout)
            {
                pLayout->DragLeave();
            }
        }

        delete _pDragDropTargetInfo;
        _pDragDropTargetInfo = NULL;

        if(_pCaret)
        {
            _pCaret->LoseFocus();
        }        
    }

    EndUndoUnit();

    RRETURN(CServer::DragLeave(fDrop));
}

STDMETHODIMP CDocument::Drop(
         IDataObject*   pDataObj,
         DWORD          grfKeyState,
         POINTL         ptScreen,
         DWORD*         pdwEffect)
{
    HRESULT     hr = S_OK;
    CLayout*    pLayout;
    CCurs       curs(IDC_WAIT);
    DWORD       dwEffect = *pdwEffect;

    hr = DragOver(grfKeyState, ptScreen, &dwEffect);

    // Continue only if drop can happen (i.e. at least one of DROPEFFECT_COPY,
    // DROPEFFECT_MOVE, etc. is set in dwEffect).
    if(hr || DROPEFFECT_NONE==dwEffect || DROPEFFECT_SCROLL==dwEffect)
    {
        goto Cleanup;
    }

    Assert(_pDragDropTargetInfo);
    if(!_pDragDropTargetInfo->_pElementTarget)
    {
        *pdwEffect = DROPEFFECT_NONE;
        goto Cleanup;
    }

    if(_pDragDropTargetInfo->_pElementTarget->Fire_ondragHelper(
        0,
        DISPID_EVMETH_ONDROP,
        DISPID_EVPROP_ONDROP,
        _T("drop"),
        pdwEffect))
    {
        if(!_pDragDropTargetInfo->_pElementTarget)
        {
            *pdwEffect = DROPEFFECT_NONE;
            goto Cleanup;
        }

        pLayout = GetLayoutForDragDrop(_pDragDropTargetInfo->_pElementTarget);
        if(pLayout)
        {
            hr = pLayout->Drop(pDataObj, grfKeyState, ptScreen, pdwEffect);
        }
        else
        {
            *pdwEffect = DROPEFFECT_NONE;
        }
    }

    // Drop can change the current site.  Set the UI active
    // site to correspond to the current site.
    if(_pElemUIActive != _pElemCurrent)
    {
        _pElemCurrent->BecomeUIActive();
    }

Cleanup:
    DragLeave(hr == S_OK);

    RRETURN1(hr,S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::SetObjectRects, IOleInplaceObject
//
//  Synopsis:   Set position and clip rectangles.
//
//-------------------------------------------------------------------------
HRESULT CDocument::SetObjectRects(LPCRECT prcPos, LPCRECT prcClip)
{
    RECT    rcPosPrev, rcClipPrev;
    RECT    rcPos, rcClip;
    long    cx, cy;
    HRESULT hr;

    Trace((_T("CDocument::SetObjectRects - rcPos(%d, %d, %d, %d) rcClip(%d, %d, %d, %d)\n"),
        (!prcPos?0:prcPos->left),
        (!prcPos?0:prcPos->top),
        (!prcPos?0:prcPos->right),
        (!prcPos?0:prcPos->bottom),
        (!prcClip?0:prcClip->left),
        (!prcClip?0:prcClip->top),
        (!prcClip?0:prcClip->right),
        (!prcClip?0:prcClip->bottom)));

    RECT    rcPosNew;       // non-scaled size of destination rectangle
    RECT    rcPosNewScaled; // scaled size of destination rectangle

    Assert(_pInPlace);

    rcPosPrev  = _pInPlace->_rcPos;
    rcClipPrev = _pInPlace->_rcClip;

    CLock Lock(this, SERVERLOCK_IGNOREERASEBKGND);

    hr = CServer::SetObjectRects(prcPos, prcClip);
    if(hr)
    {
        goto Cleanup;
    }

    rcPosNew = _pInPlace->_rcPos;
    _dci.RectzFromRectr(&rcPosNewScaled, &rcPosNew);
    rcPos  = rcPosNewScaled;
    rcClip = _pInPlace->_rcClip;

    // Update transform to contain the correct destination RECT
    cx = rcPos.right  - rcPos.left;
    cy = rcPos.bottom - rcPos.top;
    if(cx || cy)
    {
        RECT rc = rcPos;
        SIZEL sizel;

        if(!cx)
        {
            cx = 1;
            rc.right = rc.left + 1;
        }

        if(!cy)
        {
            cy = 1;
            rc.bottom = rc.top + 1;
        }

        sizel.cx = max(_sizel.cx, (LONG)_dci.HimFromDxt(_dci.DxtFromDxz(1)));
        sizel.cy = max(_sizel.cy, (LONG)_dci.HimFromDyt(_dci.DytFromDyz(1)));

        _dci.CTransform::Init(&rc, sizel);
    }

    // If necessary, initiate a re-layout
    if(rcPosNew.right-rcPosNew.left!=rcPosPrev.right-rcPosPrev.left ||
        rcPosNew.bottom-rcPosNew.top!=rcPosPrev.bottom-rcPosPrev.top)
    {
        if(_dci.IsZoomed())
        {
            SIZE sizeNew;

            sizeNew.cx = _pInPlace->_rcPos.right - _pInPlace->_rcPos.left;
            sizeNew.cy = _pInPlace->_rcPos.bottom - _pInPlace->_rcPos.top;

            _view.SetViewSize(sizeNew);
        }
        else
        {
            _view.SetViewSize(_dci._sizeDst);
        }
    }
    // Otherwise, just invalidate the view
    else if(rcClip.right-rcClip.left!=rcClipPrev.right-rcClipPrev.left ||
        rcClip.bottom-rcClip.top!=rcClipPrev.bottom-rcClipPrev.top )
    {
        // BUGBUG: We need an invalidation...will a partial work? (brendand)
        Invalidate();
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::OnWindowMessage
//
//  Synopsis:   Handle window messages dispatched from the wndproc.
//
//  Returns:    S_FALSE if message should be passed on to the
//              default window procedure.
//
//---------------------------------------------------------------
HRESULT CDocument::OnWindowMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult)
{
    HRESULT     hr;
    HWND        hwnd = _pInPlace->_hwnd;
    HWND        hwndControl = NULL;
    CLock       Lock(this);
    BOOL        fWeJustHidSelection = FALSE;
    HWND        hwndParent = NULL;
    HWND        hwndLoseFocus = NULL;
    HWND        hwndTemp = NULL;
    BOOL        fTridentHwnd = FALSE;
    BOOL        fLosingFocusToTridentChild;
    CMarkup*    pCurMarkup;

    Assert(_pElemCurrent);

    hr = S_OK;
    *plResult = 0;

    switch(msg)
    {
    case WM_APPCOMMAND:
    case WM_INPUTLANGCHANGE:
        {
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);
            PumpMessage(&Message, _pElemCurrent->GetFirstBranch());
        }
        break;

    case WM_MENUSELECT:
    case WM_INITMENUPOPUP:
        {
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);

            if(_pMenuObject)
            {
                PumpMessage(&Message, _pMenuObject->GetFirstBranch());
            }
            else
            {
                PumpMessage(&Message, _pElemCurrent->GetFirstBranch());
            }
        }
        break;

    case WM_TIMER:
        if(wParam == TIMER_DEFERUPDATEUI)
        {
            OnUpdateUI();
        }
        if(_fInEditCapture)
        {
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);    
            Assert(Message.pt.x==0 && Message.pt.y==0);            
            hr = HandleSelectionMessage(&Message, FALSE);
            *plResult = Message.lresult;
        }
        break;

        // Somewhere inside us a Selection is begin made. We clear any selection we have
        // and post the message on to all our child windows.
    case WM_BEGINSELECTION:
        {
            if(!_fNotifyBeginSelection)
            {
                NotifySelection(SELECT_NOTIFY_LOSE_FOCUS_FRAME, NULL, wParam);
            }

            ::EnumChildWindows(_pInPlace->_hwnd, (WNDENUMPROC)OnChildBeginSelection, wParam);
        }
        break;

    case WM_KILLFOCUS:
        {
            _pInPlace->_fFocus = FALSE;
            // Release any mouse capture when we loose focus
            SetMouseCapture(NULL, NULL, FALSE);

            // We explicitly check _pElementOMCapture here even though releaseCapture does the same.  This
            // is to prevent SetErrorInfo() from being called and calling DeferUpdateUI when the document
            // loses focus (dinartem).
            if(_pElementOMCapture)
            {
                releaseCapture();
            }

            _fGotKeyUp = TRUE;

            // See if we're losing focus to another Window in the same window as us.
            hwndTemp  = _pInPlace->_hwnd;
            while(hwndTemp)
            {
                hwndParent = hwndTemp;
                hwndTemp = GetParent(hwndTemp);
            }

            fLosingFocusToTridentChild = FALSE;
            hwndTemp = (HWND)wParam;
            while(hwndTemp)
            {      
                hwndLoseFocus = hwndTemp;

                if(hwndLoseFocus == _pInPlace->_hwnd)
                {
                    fLosingFocusToTridentChild = TRUE;

                    // shortcut the serach, because we already
                    // know the top-level parent in this case
                    hwndLoseFocus = hwndParent;
                    break;
                }

                hwndTemp = GetParent(hwndTemp);
            }

            fTridentHwnd = wParam && IsTridentHwnd((HWND)wParam);

            if(hwndLoseFocus!=hwndParent && HasTextSelection()) 
            {
                CMarkup* pMarkup = GetCurrentMarkup();

                if(pMarkup)
                {
                    GetCurrentMarkup()->HideSelection();
                    _fSelectionHidden = TRUE;
                    fWeJustHidSelection = TRUE;
                }
            }
            else if(hwndLoseFocus==hwndParent && fTridentHwnd)
            {
            }
            else
            {
                hr = NotifySelection(SELECT_NOTIFY_LOSE_FOCUS, NULL);
            }

            // If losing focus to a window that is not a Trident child, but is a
            // child of Trident's top-level parent (for example, the address box
            // of IE), clear the first-time-tab flag (37950)
            if(!fLosingFocusToTridentChild && hwndParent==hwndLoseFocus)
            {
                _fFirstTimeTab = FALSE;
            }

        } // Fall through

    case WM_SETFOCUS:
        {
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);
            CElement* pElemFireTarget;

            pElemFireTarget = _pElemCurrent->GetFocusBlurFireTarget(_lSubCurrent);
            Assert(pElemFireTarget);

            // Whenever we get focus, we set the flag to false to indicate
            // that we have not received a key down, so donot fire the keyups
            if(WM_SETFOCUS == msg)
            {
                _fGotKeyDown = FALSE;
                // Do not fire window onfocus if we are here as a result of onblur\onfocus
                // bringing up a modal dialog. However, we do want to fire it if window onblur
                // brought up a modal dialog that was dismissed.
                if(!TestLock(FORMLOCK_CURRENT) ||
                    (!pElemFireTarget->TestLock(CElement::ELEMENTLOCK_FOCUS)
                    && !pElemFireTarget->TestLock(CElement::ELEMENTLOCK_BLUR)
                    && _fModalDialogInOnblur))
                {
                    Fire_onfocus();
                }
            }

            // Don't fire any events for the SELECT object.  It is a special windowed
            // site that handles its own focus and blur events.
            if(_pInPlace && !_pInPlace->_fDeactivating
                && !TestLock(FORMLOCK_CURRENT) && _pElemCurrent->Tag()!=ETAG_SELECT)
            {
                // if the doc is not locked, the elem can't be locked either
                Assert(!pElemFireTarget->TestLock(CElement::ELEMENTLOCK_FOCUS));
                Assert(!pElemFireTarget->TestLock(CElement::ELEMENTLOCK_BLUR));

                // Do not fire site onfocus\onblur if we come here either
                // as a result of calling blur() or focus(), because it would
                // have been already fired in BecomeCurrent()
                if(!_fInhibitFocusFiring)
                {
                    if(WM_KILLFOCUS == msg)
                    {
                        GWPostMethodCall(pElemFireTarget, ONCALL_METHOD(CElement, Fire_onblur, fire_onblur), 1, TRUE, "CElement::Fire_onblur");
                    }

                    if(WM_SETFOCUS == msg)
                    {
                        // Fire window onblur if onfocus previously fired and the
                        // current site is not the body
                        if(_pElemCurrent != GetPrimaryElementClient())
                        {
                            Fire_onblur();
                        }

                        GWPostMethodCall(pElemFireTarget, ONCALL_METHOD(CElement, Fire_onfocus, fire_onfocus), 0, TRUE, "CElement::Fire_onfocus");
                    }
                }
            }
            // we get here if a modal dialog from current element's onblur is dismissed.
            // In this case we wan't to fire its onfocus again.
            else if(pElemFireTarget->TestLock(CElement::ELEMENTLOCK_BLUR) &&
                WM_SETFOCUS==msg && _fModalDialogInOnblur)
            {
                GWPostMethodCall(pElemFireTarget, ONCALL_METHOD(CElement, Fire_onfocus, fire_onfocus), 0, TRUE, "CElement::Fire_onfocus");
            }

            if(_view.IsActive())
            {
                _view.InvalidateFocus();

                // Display the caret
                if(_pCaret && WM_SETFOCUS==msg)
                {
                    _pCaret->UpdateCaret();
                }
            }

            pCurMarkup = GetCurrentMarkup();
            if(!fWeJustHidSelection && pCurMarkup) 
            {     
                if(_fSelectionHidden)
                {
                    pCurMarkup->ShowSelection();
                    _fSelectionHidden = FALSE;
                }
                else if(HasTextSelection())
                {
                    pCurMarkup->InvalidateSelection(FALSE); // Always inval the selection here. Fixes problems with inval from Alerts fired OnSelectStart
                }                
            }

            InvalidateDefaultSite();

            PumpMessage(&Message, _pElemCurrent->GetFirstBranch());

            // Forward the WM_KILLFOCUS, WM_SETFOCUS messages to CServer which
            // will notify the control site of the focus change.  Forward the
            // messages when no site has focus or when the site with focus is
            // a dataframe (non ole site).
            if(!_pElemCurrent->TestClassFlag(CElement::ELEMENTDESC_OLESITE))
            {
                hr = CServer::OnWindowMessage(msg, wParam, lParam, plResult);
            }

            // Do not fire window onblur if we are here as a result of onblur\onfocus
            // bringing up a modal dialog.
            if(WM_KILLFOCUS==msg && !TestLock(FORMLOCK_CURRENT))
            {
                Fire_onblur(TRUE);
            }

            break;
        }

    case WM_CAPTURECHANGED:
        SetMouseCapture(NULL, NULL);
        releaseCapture();
        break;

        // Messages sent to site under mouse.
    case WM_SETCURSOR:
        if(LOWORD(lParam) == HTCLIENT)
        {
            POINT pt;

            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);

            hr = OnMouseMessage(
                msg,
                wParam,
                lParam,
                plResult,
                pt.x, pt.y);
        }
        else
        {
            hr = CServer::OnWindowMessage(msg, wParam, lParam, plResult);
        }

        break;

    case WM_CONTEXTMENU:
        {
            POINT pt;
            pt.x = MAKEPOINTS(lParam).x;
            pt.y = MAKEPOINTS(lParam).y;
            if((pt.x <0 || pt.y<0) && _pElemCurrent)
            {
                // pt.x & pt.y are supposed to be in screen coordinates;
                // the only case when they are -1,-1 is if WM_CONTEXTMENU
                // originated from Shift-F10 or Windows keyboard
                // key 'Menu'; in this case we send the message to the
                // current site.
                CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);
                hr = PumpMessage(&Message, _pElemCurrent->GetFirstBranch());
            }
            else
            {
                Assert(pt.x>=0 && pt.y>=0);
                ScreenToClient(_pInPlace->_hwnd, &pt);

                // CONSIDER: should this set the focus before
                // sending the OnMouseMessage.  Does anything actually
                // use this code path? (jbeda)
                hr = OnMouseMessage(
                    msg,
                    wParam,
                    lParam,
                    plResult,
                    pt.x, pt.y);
            }
        }
        break;

    case WM_MOUSEWHEEL:
        {
            POINT ptCursor;
            RECT  rcCurrent;

            // Check where the wheel rotates.
            // If the wheel is rotated inside the current CDoc window, let
            // us handle it. Otherwise, let DefWindowProc and handle it and
            // bubble to the parent window.
            ptCursor.x = MAKEPOINTS(lParam).x;
            ptCursor.y = MAKEPOINTS(lParam).y;
            ::GetWindowRect(InPlace()->_hwnd, &rcCurrent);

            if(PtInRect(&rcCurrent, ptCursor))
            {
                ScreenToClient(hwnd, &ptCursor);
                Assert(msg == WM_MOUSEWHEEL);
                hr = OnMouseMessage(
                    WM_MOUSEWHEEL,
                    wParam,
                    lParam,
                    plResult,
                    ptCursor.x,
                    ptCursor.y);
            }
            else
            {
                CServer::OnWindowMessage(msg, wParam, lParam, plResult);
            }
        }
        break;

        // Messages sent to site under mouse
    case WM_LBUTTONDOWN:
    case WM_MBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
    case WM_RBUTTONDBLCLK:
    case WM_MBUTTONDBLCLK:
        switch(msg)
        {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
            _fGotLButtonDown = TRUE;
            break;
        case WM_MBUTTONDOWN:
        case WM_MBUTTONDBLCLK:
            _fGotMButtonDown = TRUE;
            break;
        case WM_RBUTTONDOWN:
        case WM_RBUTTONDBLCLK:
            _fGotRButtonDown = TRUE;
            break;
        }
        _fGotKeyUp = FALSE;
        hr = OnMouseMessage(
            msg,
            wParam,
            lParam,
            plResult,
            MAKEPOINTS(lParam).x, MAKEPOINTS(lParam).y);
        break;

    case WM_LBUTTONUP:
        if(!_fGotLButtonDown)
        {
            hr = S_OK;
            goto Cleanup;
        }
        _fGotLButtonDown = FALSE;
        hr = OnMouseMessage(
            msg,
            wParam,
            lParam,
            plResult,
            MAKEPOINTS(lParam).x, MAKEPOINTS(lParam).y);
        break;

    case WM_MBUTTONUP:
        if(!_fGotMButtonDown)
        {
            hr = S_OK;
            goto Cleanup;
        }
        _fGotMButtonDown = FALSE;
        hr = OnMouseMessage(
            msg,
            wParam,
            lParam,
            plResult,
            MAKEPOINTS(lParam).x, MAKEPOINTS(lParam).y);
        break;

    case WM_RBUTTONUP:
        if(!_fGotRButtonDown)
        {
            hr = S_OK;
            goto Cleanup;
        }
        _fGotRButtonDown = FALSE;
        hr = OnMouseMessage(
            msg,
            wParam,
            lParam,
            plResult,
            MAKEPOINTS(lParam).x, MAKEPOINTS(lParam).y);
        break;

    case WM_MOUSEMOVE:
        hr = OnMouseMessage(
            msg,
            wParam,
            lParam,
            plResult,
            MAKEPOINTS(lParam).x, MAKEPOINTS(lParam).y);
        break;

    case WM_MOUSELEAVE:
        {
            // Default messages get forwarded to the base class's
            // window procedure.
            hr = CServer::OnWindowMessage(msg, wParam, lParam, plResult);
        }
        break;

    case WM_HELP:
        hr = OnHelp((HELPINFO*)lParam);
        break;

        // Keyboard messages: (Sent to _pElemCurrent)
    case WM_KEYDOWN:
        _fGotKeyDown = TRUE;
        // fall thru
    case WM_KEYUP:
        if(!_fGotKeyDown)
        {
            // NOTE (sujalp): Do not reset the _fGotKeyDown to
            // FALSE here. It is set to FALSE only when we get
            // the focus and that too to eat up the spurious key
            // up which we might get after we GOT_FOCUS.
            hr = S_OK;
            goto Cleanup;
        }
        // fall thru
    case WM_CHAR:
    case WM_DEADCHAR:
    case WM_SYSKEYDOWN:
    case WM_SYSKEYUP:
    case WM_SYSCHAR:
    case WM_SYSDEADCHAR:

#ifndef NO_IME
    case WM_IME_SETCONTEXT:
    case WM_IME_NOTIFY:
    case WM_IME_CONTROL:
    case WM_IME_COMPOSITIONFULL:
    case WM_IME_SELECT:
    case WM_IME_CHAR:
    case WM_IME_KEYDOWN:
    case WM_IME_KEYUP:
    case WM_IME_STARTCOMPOSITION:
    case WM_IME_ENDCOMPOSITION:
    case WM_IME_COMPOSITION:
#endif // !NO_IME
        {
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);
            hr = S_FALSE;

            // If the captured site didn't handle it, pass message to current
            // site.
            if(S_FALSE == hr)
            {
                hr = PumpMessage(&Message, _pElemCurrent->GetFirstBranch());
                *plResult = Message.lresult;
            }
            DeferUpdateUI();
            break;
        }

        // Messages that are either handled or reflected.
    case WM_COMMAND:
        if(!_pMenuObject && lParam &&
            GetParent(GET_WM_COMMAND_HWND(wParam, lParam))==_pInPlace->_hwnd)
        {
            // Command is bubbling up from a control. Reflect it back.
            hwndControl = GET_WM_COMMAND_HWND(wParam, lParam);
            goto ReflectMessage;
        }
        else
        {
            // It's our command.
            OnCommand(GET_WM_COMMAND_ID(wParam, lParam), GET_WM_COMMAND_HWND(wParam, lParam), GET_WM_COMMAND_CMD(wParam, lParam));
        }
        break;

    case WM_DEFERZORDER:
        {
            FixZOrder();
        }
        break;

    case WM_ACTIVEMOVIE:
        {
            CNotification nf;

            nf.ActiveMovie(PrimaryRoot(), (void*)lParam);
            BroadcastNotify(&nf);
        }
        break;

        //  OLE Control v1.0 reflected messages
    case WM_DRAWITEM:
        hwndControl = ((DRAWITEMSTRUCT*)lParam)->hwndItem;
        goto ReflectMessage;

    case WM_MEASUREITEM:
        //  BUGBUG how did the control ID ever get set?
        hwndControl = GetDlgItem(hwnd, (UINT) wParam);
        goto ReflectMessage;

    case WM_DELETEITEM:
        hwndControl = ((DELETEITEMSTRUCT*)lParam)->hwndItem;
        goto ReflectMessage;

    case WM_COMPAREITEM:
        hwndControl = ((COMPAREITEMSTRUCT*)lParam)->hwndItem;
        goto ReflectMessage;

    case WM_NOTIFY:
        hwndControl = ((LPNMHDR)lParam)->hwndFrom;
        goto ReflectMessage;

    case WM_VSCROLL:
    case WM_HSCROLL:
        if(!lParam)
        {
            // If lParam is NULL, should try to scroll ourselves
            CMessage Message(_pInPlace->_hwnd, msg, wParam, lParam);
            hr = PumpMessage(&Message, _pElemCurrent->GetFirstBranch());
            break;
        }
        // if lParam is defined, should fall through to ReflectMessage

    case WM_CTLCOLORBTN:
    case WM_CTLCOLORDLG:
    case WM_CTLCOLORLISTBOX:
    case WM_CTLCOLORMSGBOX:
    case WM_CTLCOLORSCROLLBAR:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
    case WM_VKEYTOITEM:
    case WM_CHARTOITEM:
        hwndControl = (HWND)lParam;
        goto ReflectMessage;

    case WM_PARENTNOTIFY:
        if(LOWORD(wParam)==WM_CREATE || LOWORD(wParam)==WM_DESTROY)
        {
            hwndControl = (HWND)lParam;
            goto ReflectMessage;
        }
        break;

    case WM_GETDLGCODE:
        *plResult = DLGC_WANTCHARS|DLGC_WANTARROWS;
        hr = S_OK;
        break;

    case WM_ERASEBKGND:
        *plResult = OnEraseBkgnd((HDC)wParam);
        hr = S_OK;
        break;

    // wlw note
    //case WM_GETOBJECT:
    //    // ActiveX "accessibility", creation of in-context proxy
    //    OnAccGetObjectInContext(msg, wParam, lParam, plResult);
    //    hr = S_OK;  // no need to continue 
    //    break;

    default:
        //if(msg == _g_msgHtmlGetobject)
        //{
        //    // ActiveX "accessibility"
        //    OnAccGetObject(msg, wParam, lParam, plResult);
        //    hr = S_OK;  // Stop bubbling.
        //    break;
        //}

        // All other messages get forwarded to the base class's
        // window procedure.
        hr = CServer::OnWindowMessage(msg, wParam, lParam, plResult);
        break;
    }

Cleanup:
    // BUGBUG should we come here with S_FALSE, or does that mean
    // defProc was not called when should?
    RRETURN1(hr, S_FALSE);

ReflectMessage:
    Assert(hwndControl);
    *plResult= SendMessage(hwndControl, msg+OCM__BASE, wParam, lParam);

    hr = S_OK;
    goto Cleanup;
}

//+-------------------------------------------------------------------------
//
//  Method:     CDoc::QueryStatus
//
//  Synopsis:   Called to discover if a given command is supported
//              and if it is, what's its state.  (disabled, up or down)
//
//--------------------------------------------------------------------------
HRESULT CDocument::QueryStatus(
          GUID*         pguidCmdGroup,
          ULONG         cCmds,
          MSOCMD        rgCmds[],
          MSOCMDTEXT*   pcmdtext)
{
    Trace0("CDocument::QueryStatus\n");

    // Check to see if the command is in our command set.
    if(!IsCmdGroupSupported(pguidCmdGroup))
    {
        RRETURN(OLECMDERR_E_UNKNOWNGROUP);
    }

    MSOCMD* pCmd;
    INT     c;
    UINT    idm;
    HRESULT hr = S_OK;
    CTArg   ctarg;
    CTQueryStatusArg qsarg;

    // Loop through each command in the ary, setting the status of each.
    for(pCmd=rgCmds,c=cCmds; --c>=0; pCmd++)
    {
        // By default command status is NOT SUPPORTED.
        pCmd->cmdf = 0;

        idm = IDMFromCmdID(pguidCmdGroup, pCmd->cmdID);
        if(pcmdtext && pcmdtext->cmdtextf==MSOCMDTEXTF_STATUS)
        {
            /*pcmdtext[c].cwActual = LoadString(
                GetResourceHInst(),
                IDS_MENUHELP(idm),
                pcmdtext[c].rgwz,
                pcmdtext[c].cwBuf); wlw note*/
        }

        switch(idm)
        {
        case IDM_REPLACE:
        case IDM_FONT:
        case IDM_GOTO:
        case IDM_HYPERLINK:
        case IDM_BOOKMARK:
        case IDM_IMAGE:
            break;

        case IDM_FIND:
            if(_dwFlagsHostInfo & DOCHOSTUIFLAG_DIALOG)
            {
                pCmd->cmdf = MSOCMDSTATE_DISABLED;
            }
            else
            {
                pCmd->cmdf = MSOCMDSTATE_UP;
            }
            break;


        case IDM_REDO:
        case IDM_UNDO:
            /*QueryStatusUndoRedo((IDM_UNDO==idm), pCmd, pcmdtext); wlw note*/
            break;

        case IDM_BASELINEFONT1:
        case IDM_BASELINEFONT2:
        case IDM_BASELINEFONT3:
        case IDM_BASELINEFONT4:
        case IDM_BASELINEFONT5:
            // depend on that IDM_BASELINEFONT1, IDM_BASELINEFONT2,
            // IDM_BASELINEFONT3, IDM_BASELINEFONT4, IDM_BASELINEFONT5 to be
            // consecutive integers.
            {
                if(GetBaselineFont() == (short)(idm-IDM_BASELINEFONT1+BASELINEFONTMIN))
                {
                    pCmd->cmdf = MSOCMDSTATE_DOWN;
                }
                else
                {
                    pCmd->cmdf = MSOCMDSTATE_UP;
                }
            }
            break;

        case IDM_LANGUAGE:
            pCmd->cmdf = MSOCMDSTATE_UP;
            break;

        case IDM_DIRLTR:
        case IDM_DIRRTL:
            {
                BOOL fDocRTL;

                hr = GetDocDirection(&fDocRTL);
                if(hr==S_OK && ((!fDocRTL)^(idm==IDM_DIRRTL)))
                {
                    pCmd->cmdf = MSOCMDSTATE_DOWN;
                }
                else
                {
                    pCmd->cmdf = MSOCMDSTATE_UP;
                }
            }
            break;
        }

        // If still not handled then try menu object.
        ctarg.pguidCmdGroup = pguidCmdGroup;
        ctarg.fQueryStatus = TRUE;
        ctarg.pqsArg = &qsarg;
        qsarg.cCmds = 1;
        qsarg.rgCmds = pCmd;
        qsarg.pcmdtext = pcmdtext;

        if(!pCmd->cmdf && _pMenuObject)
        {
            hr = RouteCTElement(_pMenuObject, &ctarg);
        }

        // Next try the current element;
        if(!pCmd->cmdf && _pElemCurrent)
        {
            hr = RouteCTElement(_pElemCurrent, &ctarg);
        }

        // Finally try edit router
        if(!pCmd->cmdf)
        {
            hr = _EditRouter.QueryStatusEditCommand(
                pguidCmdGroup,
                1,
                pCmd,
                pcmdtext,
                (IUnknown*)(IPrivateUnknown*)this,
                this);
            if(hr == S_OK
                && (pCmd->cmdf&MSOCMDF_ENABLED)
                && _pElemEditContext
                && _pElemEditContext!=_pElemCurrent)
            {
                DWORD cmdfSave = pCmd->cmdf;

                // Check if the edit context wants to disallow the command (fix for 46807)
                pCmd->cmdf = 0;
                hr = _pElemEditContext->QueryStatus(
                    pguidCmdGroup,
                    1,
                    pCmd,
                    pcmdtext);
                if(pCmd->cmdf != MSOCMDSTATE_DISABLED)
                {
                    pCmd->cmdf = cmdfSave;
                }
            }
        }

        // Prevent any command but the first from setting this.
        pcmdtext = NULL;
    }

    SRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CDocument::Exec
//
//  Synopsis:   Called to execute a given command.  If the command is not
//              consumed, it may be routed to other objects on the routing
//              chain.
//
//--------------------------------------------------------------------------
HRESULT CDocument::Exec(
       GUID*        pguidCmdGroup,
       DWORD        nCmdID,
       DWORD        nCmdexecopt,
       VARIANTARG*  pvarargIn,
       VARIANTARG*  pvarargOut)
{
    Trace0("CDocument::Exec\n");

    if(!IsCmdGroupSupported(pguidCmdGroup))
    {
        RRETURN(OLECMDERR_E_UNKNOWNGROUP);
    }

    CDocument::CLock    Lock(this);
    UINT                idm;
    HRESULT             hr = OLECMDERR_E_NOTSUPPORTED;
    CTArg               ctarg;
    CTExecArg           execarg;

    idm = IDMFromCmdID(pguidCmdGroup, nCmdID);

    ctarg.pguidCmdGroup = pguidCmdGroup;
    ctarg.fQueryStatus = FALSE;
    ctarg.pexecArg = &execarg;
    execarg.nCmdID = nCmdID;
    execarg.nCmdexecopt = nCmdexecopt;
    execarg.pvarargIn = pvarargIn;
    execarg.pvarargOut = pvarargOut;

    if(hr==OLECMDERR_E_NOTSUPPORTED && _pElemCurrent)
    {
        hr = RouteCTElement(_pElemCurrent, &ctarg);
    }

    if(hr == OLECMDERR_E_NOTSUPPORTED)
    {
        hr = _EditRouter.ExecEditCommand(pguidCmdGroup,
            nCmdID, nCmdexecopt,
            pvarargIn, pvarargOut,
            (IUnknown*)(IPrivateUnknown*)this,
            this);
    }

    SRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CDocument::RouteCTElement
//
//  Synopsis:   Route a command target call, either QueryStatus or Exec
//              to an element
//
//--------------------------------------------------------------------------
HRESULT CDocument::RouteCTElement(CElement* pElement, CTArg* parg)
{
    HRESULT     hr = OLECMDERR_E_NOTSUPPORTED;
    CTreeNode*  pNodeParent;
    AAINDEX     aaindex;
    IUnknown*   pUnk = NULL;
    CDocument*  pDocSec = NULL;

    if(TestLock(FORMLOCK_QSEXECCMD))
    {
        pDocSec = this;
    }

    if(pDocSec)
    {
        aaindex = pDocSec->FindAAIndex(DISPID_INTERNAL_INVOKECONTEXT, CAttrValue::AA_Internal);
        if(aaindex != AA_IDX_UNKNOWN)
        {
            hr = pDocSec->GetUnknownObjectAt(aaindex, &pUnk);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    while(pElement)
    {
        Assert(pElement->Tag()!=ETAG_ROOT || pElement==_pPrimaryMarkup->Root());

        if(pElement == _pPrimaryMarkup->Root())
        {
            break;
        }

        if(pUnk)
        {
            pElement->AddUnknownObject(DISPID_INTERNAL_INVOKECONTEXT, pUnk, CAttrValue::AA_Internal);
        }

        if(parg->fQueryStatus)
        {
            Assert(parg->pqsArg->cCmds == 1);

            hr = pElement->QueryStatus(
                parg->pguidCmdGroup,
                parg->pqsArg->cCmds,
                parg->pqsArg->rgCmds,
                parg->pqsArg->pcmdtext);
            if(parg->pqsArg->rgCmds[0].cmdf)
            {
                break; // Element handled it.
            }
        }
        else
        {
            hr = pElement->Exec(
                parg->pguidCmdGroup,
                parg->pexecArg->nCmdID,
                parg->pexecArg->nCmdexecopt,
                parg->pexecArg->pvarargIn,
                parg->pexecArg->pvarargOut);
            if(hr != OLECMDERR_E_NOTSUPPORTED)
            {
                break;
            }
        }

        if(pUnk)
        {
            pElement->FindAAIndexAndDelete(DISPID_INTERNAL_INVOKECONTEXT, CAttrValue::AA_Internal);
        }

        if(pElement == GetPrimaryElementClient())
        {
            break;
        }

        if(!pElement->IsInMarkup())
        {
            break;
        }

        pNodeParent = pElement->GetFirstBranch()->Parent();
        pElement = pNodeParent ? pNodeParent->Element() : NULL;

        Assert(pElement->Tag() != ETAG_ROOT);
    }

Cleanup:
    if(pUnk && pElement)
    {
        pElement->FindAAIndexAndDelete(DISPID_INTERNAL_INVOKECONTEXT, CAttrValue::AA_Internal);
    }
    ReleaseInterface(pUnk);
    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnAmbientPropertyChange, public
//
//  Synopsis:   Captures ambient property changes and takes appropriate action.
//
//  Arguments:  [dispidProp] -- Property which changed.
//
//----------------------------------------------------------------------------
STDMETHODIMP CDocument::OnAmbientPropertyChange(DISPID dispidProp)
{
    BOOL            fAll = (dispidProp == DISPID_UNKNOWN);
    CNotification   nf;

    // Forward to all the sites.
    if(GetPrimaryElementClient())
    {
        nf.AmbientPropChange(PrimaryRoot(), (void*)(DWORD_PTR)dispidProp);
        BroadcastNotify(&nf);
    }
    return S_OK;
}

HRESULT CDocument::CutCopyMove(
           IMarkupPointer*  pIPointerStart,
           IMarkupPointer*  pIPointerFinish,
           IMarkupPointer*  pIPointerTarget,
           BOOL             fRemove)
{
    HRESULT         hr = S_OK;
    CMarkupPointer* pPointerStart;
    CMarkupPointer* pPointerFinish;
    CMarkupPointer* pPointerTarget = NULL;

    // Check argument sanity
    if(pIPointerStart==NULL || !IsOwnerOf(pIPointerStart) ||
        pIPointerFinish==NULL || !IsOwnerOf(pIPointerFinish) ||
        (pIPointerTarget!=NULL && !IsOwnerOf(pIPointerTarget)))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // Get the internal objects
    hr = pIPointerStart->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerStart);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = pIPointerFinish->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerFinish);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    if(pIPointerTarget)
    {
        hr = pIPointerTarget->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerTarget);
        if(hr)
        {
            hr = E_INVALIDARG;
            goto Cleanup;
        }
    }

    // More sanity checks
    if(!pPointerStart->IsPositioned() || !pPointerFinish->IsPositioned())
    {
        hr = CTL_E_UNPOSITIONEDPOINTER;
        goto Cleanup;
    }

    if(pPointerStart->Markup() != pPointerFinish->Markup())
    {
        hr = CTL_E_INCOMPATIBLEPOINTERS;
        goto Cleanup;
    }

    // Make sure the start if before the finish
    EnsureLogicalOrder(pPointerStart, pPointerFinish);

    // More checks
    if(pPointerTarget && !pPointerTarget->IsPositioned())
    {
        hr = CTL_E_UNPOSITIONEDPOINTER;
        goto Cleanup;
    }

    // Do it
    hr = CutCopyMove(pPointerStart, pPointerFinish, pPointerTarget, fRemove, NULL);

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::CutCopyMove(
            CMarkupPointer* pPointerStart,
            CMarkupPointer* pPointerFinish,
            CMarkupPointer* pPointerTarget,
            BOOL            fRemove,
            DWORD           dwFlags)
{
    HRESULT     hr = S_OK;
    CTreePosGap tpgStart;
    CTreePosGap tpgFinish;
    CTreePosGap tpgTarget;
    CMarkup*    pMarkupSource = NULL;
    CMarkup*    pMarkupTarget = NULL;

    // Sanity check the args
    Assert(pPointerStart);
    Assert(pPointerFinish);
    Assert(OldCompare(pPointerStart, pPointerFinish) <= 0);
    Assert(pPointerStart->IsPositioned());
    Assert(pPointerFinish->IsPositioned());
    Assert(pPointerStart->Markup() == pPointerFinish->Markup());
    Assert(!pPointerTarget || pPointerTarget->IsPositioned());

    // Make sure unembedded pointers get in before the modification
    hr = pPointerStart->Markup()->EmbedPointers();
    if(hr)
    {
        goto Cleanup;
    }

    if(pPointerTarget)
    {
        hr = pPointerTarget->Markup()->EmbedPointers();
        if(hr)
        {
            goto Cleanup;
        }
    }

    // Set up the gaps
    tpgStart.MoveTo(pPointerStart->GetEmbeddedTreePos(), TPG_LEFT);
    tpgFinish.MoveTo(pPointerFinish->GetEmbeddedTreePos(), TPG_RIGHT);

    if(pPointerTarget)
    {
        tpgTarget.MoveTo(pPointerTarget->GetEmbeddedTreePos(), TPG_LEFT);
    }

    pMarkupSource = pPointerStart->Markup();

    if(pPointerTarget)
    {
        pMarkupTarget = pPointerTarget->Markup();
    }

    // Do it.
    if(pPointerTarget)
    {
        hr = pMarkupSource->SpliceTreeInternal(&tpgStart, &tpgFinish, pPointerTarget->Markup(),
            &tpgTarget, fRemove, dwFlags);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else
    {
        Assert(fRemove);
        hr = pMarkupSource->SpliceTreeInternal(&tpgStart, &tpgFinish, NULL, NULL, fRemove, dwFlags);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

static ELEMENT_TAG_ID TagIdFromETag(ELEMENT_TAG etag)
{
    ELEMENT_TAG_ID tagID = TAGID_NULL;

    switch(etag)
    {
#define X(Y) case ETAG_##Y:tagID=TAGID_##Y;break;
        X(UNKNOWN) X(A) X(ACRONYM) X(ADDRESS) X(APPLET) X(AREA) X(B) X(BASE)
            X(BDO) X(BIG) X(BLINK) X(BLOCKQUOTE) X(BODY) X(BR) X(BUTTON) X(CAPTION)
            X(CENTER) X(CITE) X(CODE) X(COL) X(COLGROUP) X(COMMENT) X(DD) X(DEL) X(DFN) X(DIR)
            X(DIV) X(DL) X(DT) X(EM) X(EMBED) X(FIELDSET) X(FONT) X(FORM) X(FRAME)
            X(FRAMESET) X(H1) X(H2) X(H3) X(H4) X(H5) X(H6) X(HEAD) X(HR) X(HTML) X(I) X(IFRAME)
            X(IMG) X(INPUT) X(INS) X(KBD) X(LABEL) X(LEGEND) X(LI) X(LINK) X(LISTING)
            X(MARQUEE) X(MENU) X(NEXTID) X(NOBR) X(NOEMBED) X(NOFRAMES)
            X(NOSCRIPT) X(OBJECT) X(OL) X(OPTION) X(P) X(PARAM) X(PLAINTEXT) X(PRE) X(Q)
            X(RP) X(RT) X(RUBY) X(S) X(SAMP) X(SCRIPT) X(SELECT) X(SMALL) X(SPAN) 
            X(STRIKE) X(STRONG) X(STYLE) X(SUB) X(SUP) X(TABLE) X(TBODY) X(TC) X(TD) X(TEXTAREA)
            X(TFOOT) X(TH) X(THEAD) X(TR) X(TT) X(U) X(UL) X(VAR) X(WBR) X(XMP)
#undef X

    case ETAG_TITLE_ELEMENT:
    case ETAG_TITLE_TAG:
        tagID = TAGID_TITLE; break;

    case ETAG_GENERIC:
    case ETAG_GENERIC_BUILTIN:
    case ETAG_GENERIC_LITERAL:
        tagID = TAGID_GENERIC; break;

    case ETAG_RAW_COMMENT:
        tagID = TAGID_COMMENT_RAW; break;

        // BUGBUG (MohanB) For now, use INPUT's tag id for TXTSLAVE. Needs to be generified.
    case ETAG_TXTSLAVE:
        tagID = TAGID_INPUT; break;
    }

    AssertSz(tagID!=TAGID_NULL, "Invalid ELEMENT_TAG");

    return tagID;
}

ELEMENT_TAG ETagFromTagId(ELEMENT_TAG_ID tagID)
{
    ELEMENT_TAG etag = ETAG_NULL;

    switch(tagID)
    {
#define X(Y) case TAGID_##Y: etag = ETAG_##Y; break;
        X(UNKNOWN) X(A) X(ACRONYM) X(ADDRESS) X(APPLET) X(AREA) X(B) X(BASE)
            X(BDO) X(BIG) X(BLINK) X(BLOCKQUOTE) X(BODY) X(BR) X(BUTTON) X(CAPTION)
            X(CENTER) X(CITE) X(CODE) X(COL) X(COLGROUP) X(COMMENT) X(DD) X(DEL) X(DFN) X(DIR)
            X(DIV) X(DL) X(DT) X(EM) X(EMBED) X(FIELDSET) X(FONT) X(FORM) X(FRAME)
            X(FRAMESET) X(GENERIC) X(H1) X(H2) X(H3) X(H4) X(H5) X(H6) X(HEAD) X(HR) X(HTML) X(I) X(IFRAME)
            X(IMG) X(INPUT) X(INS) X(KBD) X(LABEL) X(LEGEND) X(LI) X(LINK) X(LISTING)
            X(MARQUEE) X(MENU) X(NEXTID) X(NOBR) X(NOEMBED) X(NOFRAMES)
            X(NOSCRIPT) X(OBJECT) X(OL) X(OPTION) X(P) X(PARAM) X(PLAINTEXT) X(PRE) X(Q)
            X(RP) X(RT) X(RUBY) X(S) X(SAMP) X(SCRIPT) X(SELECT) X(SMALL) X(SPAN) 
            X(STRIKE) X(STRONG) X(STYLE) X(SUB) X(SUP) X(TABLE) X(TBODY) X(TC) X(TD) X(TEXTAREA) 
            X(TFOOT) X(TH) X(THEAD) X(TR) X(TT) X(U) X(UL) X(VAR) X(WBR) X(XMP)
#undef X

    case TAGID_TITLE: etag = ETAG_TITLE_ELEMENT; break;
    case TAGID_COMMENT_RAW: etag = ETAG_RAW_COMMENT; break;
    }

    AssertSz(etag!=ETAG_NULL, "Invalid ELEMENT_TAG_ID");

    return etag;
}

HRESULT CDocument::CreateMarkupPointer(IMarkupPointer** ppIPointer)
{
    HRESULT hr;
    CMarkupPointer* pPointer = NULL;

    hr = CreateMarkupPointer(&pPointer);

    if(hr)
    {
        goto Cleanup;
    }

    hr = pPointer->QueryInterface(IID_IMarkupPointer, (void**)ppIPointer);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    ReleaseInterface(pPointer);
    RRETURN(hr);
}

HRESULT CDocument::CreateMarkupContainer(IMarkupContainer** ppIMarkupContainer)
{
    HRESULT hr;
    CMarkup* pMarkup = NULL;

    if(!ppIMarkupContainer)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = CreateMarkup(&pMarkup);

    if(hr)
    {
        goto Cleanup;
    }

    hr = pMarkup->QueryInterface(IID_IMarkupContainer, (void**)ppIMarkupContainer);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    if(pMarkup)
    {
        pMarkup->Release();
    }

    RRETURN(hr);
}

HRESULT CDocument::CreateElement(ELEMENT_TAG_ID tagID, OLECHAR* pchAttributes, IHTMLElement** ppIHTMLElement)
{
    HRESULT     hr = S_OK;
    ELEMENT_TAG etag;
    CElement*   pElement = NULL;

    if(!ppIHTMLElement)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    etag = ETagFromTagId(tagID);

    if(etag == ETAG_NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = CreateElement(etag, &pElement);

    if(hr)
    {
        goto Cleanup;
    }

    hr = pElement->QueryInterface(IID_IHTMLElement, (void**)ppIHTMLElement);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    CElement::ReleasePtr(pElement);
    RRETURN(hr);
}

HRESULT CDocument::InsertElement(IHTMLElement* pIElementInsert, IMarkupPointer* pIPointerStart, IMarkupPointer* pIPointerFinish)
{
    HRESULT         hr = S_OK;
    CElement*       pElementInsert;
    CMarkupPointer* pPointerStart;
    CMarkupPointer* pPointerFinish;

    // If the the finish is not specified, set it to the start to make the element span
    // nothing at the start.
    if(!pIPointerFinish)
    {
        pIPointerFinish = pIPointerStart;
    }

    // Make sure all the arguments are specified and belong to this document
    if (!pIElementInsert || !IsOwnerOf(pIElementInsert) ||
        !pIPointerStart  || !IsOwnerOf(pIPointerStart)  ||
        !pIPointerFinish || !IsOwnerOf(pIPointerFinish))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // get the internal objects corresponding to the arguments
    hr = pIElementInsert->QueryInterface(CLSID_CElement, (void**)&pElementInsert);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    Assert(pElementInsert);

    // The element must not already be in a tree
    if(pElementInsert->GetFirstBranch())
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // Get the "real" objects associated with these pointer interfaces
    hr = pIPointerStart->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerStart);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = pIPointerFinish->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerFinish);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // Make sure both pointers are positioned, and in the same tree
    if(!pPointerStart->IsPositioned() || !pPointerFinish->IsPositioned())
    {
        hr = CTL_E_UNPOSITIONEDPOINTER;
        goto Cleanup;
    }

    if(pPointerStart->Markup() != pPointerFinish->Markup())
    {
        hr = CTL_E_INCOMPATIBLEPOINTERS;
        goto Cleanup;
    }

    // Make sure the start if before the finish
    EnsureLogicalOrder(pPointerStart, pPointerFinish);

    hr = InsertElement(pElementInsert, pPointerStart, pPointerFinish);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::RemoveElement(IHTMLElement* pIElementRemove)
{
    HRESULT hr = S_OK;
    CElement* pElementRemove = NULL;

    // Element to be removed must be specified and it must be associated
    // with this document.
    if(!pIElementRemove || !IsOwnerOf(pIElementRemove))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // get the interneal objects corresponding to the arguments
    hr = pIElementRemove->QueryInterface(CLSID_CElement, (void**)&pElementRemove);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // Make sure element is in the tree
    if(!pElementRemove->IsInMarkup())
    {
        hr = CTL_E_UNPOSITIONEDELEMENT;
        goto Cleanup;
    }

    // Do the remove
    hr = RemoveElement(pElementRemove);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::Remove(IMarkupPointer* pIPointerStart, IMarkupPointer* pIPointerFinish)
{
    return CutCopyMove(pIPointerStart, pIPointerFinish, NULL, TRUE);
}

HRESULT CDocument::Copy(
            IMarkupPointer* pIPointerStart,
            IMarkupPointer* pIPointerFinish,
            IMarkupPointer* pIPointerTarget)
{
    return CutCopyMove(pIPointerStart, pIPointerFinish, pIPointerTarget, FALSE);
}

HRESULT CDocument::Move(
            IMarkupPointer* pIPointerStart,
            IMarkupPointer* pIPointerFinish,
            IMarkupPointer* pIPointerTarget)
{
    return CutCopyMove(pIPointerStart, pIPointerFinish, pIPointerTarget, TRUE);
}

HRESULT CDocument::InsertText(OLECHAR* pchText, long cch, IMarkupPointer* pIPointerTarget)
{
    HRESULT hr = S_OK;
    CMarkupPointer* pPointerTarget;

    if(!pIPointerTarget || !IsOwnerOf(pIPointerTarget))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // Get the internal objects corresponding to the arguments
    hr = pIPointerTarget->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerTarget);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // more sanity checks
    if(!pPointerTarget->IsPositioned())
    {
        hr = CTL_E_UNPOSITIONEDPOINTER;
        goto Cleanup;
    }

    // Do it
    hr = InsertText(pPointerTarget, pchText, cch);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::IsScopedElement(IHTMLElement* pIHTMLElement, BOOL* pfScoped)
{
    HRESULT hr = S_OK;
    CElement* pElement = NULL;

    if(!pIHTMLElement || !pfScoped || !IsOwnerOf(pIHTMLElement))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = pIHTMLElement->QueryInterface(CLSID_CElement, (void**)&pElement);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pfScoped = ! pElement->IsNoScope();

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::GetElementTagId(IHTMLElement* pIHTMLElement, ELEMENT_TAG_ID* ptagId)
{
    HRESULT hr;
    CElement* pElement = NULL;

    if(!pIHTMLElement || !ptagId || !IsOwnerOf(pIHTMLElement))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = pIHTMLElement->QueryInterface(CLSID_CElement, (void**)&pElement);
    if(hr)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ptagId = TagIdFromETag(pElement->Tag());

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::GetTagIDForName(BSTR bstrName, ELEMENT_TAG_ID* ptagId)
{
    HRESULT hr = S_OK;

    if(!bstrName || !ptagId)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ptagId = TagIdFromETag(EtagFromName(bstrName, SysStringLen(bstrName)));

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::GetNameForTagID(ELEMENT_TAG_ID tagId, BSTR* pbstrName)
{
    HRESULT hr = S_OK;

    if(!pbstrName)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = FormsAllocString(NameFromEtag(ETagFromTagId(tagId)), pbstrName);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::BeginUndoUnit(OLECHAR* pchDescription)
{
    HRESULT hr = S_OK;

//    if(!pchDescription)
//    {
//        hr = E_INVALIDARG;
//        goto Cleanup;
//    }
//
//    if(!_uOpenUnitsCounter)
//    {
//        _pMarkupServicesParentUndo = new CParentUndo(this);
//
//        if(!_pMarkupServicesParentUndo)
//        {
//            hr = E_OUTOFMEMORY;
//            goto Cleanup;
//        }
//
//        hr = _pMarkupServicesParentUndo->Start(pchDescription);
//    }        
//
//    {
//        CSelectionUndo Undo(_pElemCurrent, GetCurrentMarkup());    
//    }
//
//    _uOpenUnitsCounter++;
//
//Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::EndUndoUnit()
{
    HRESULT hr = S_OK;

//    if(!_uOpenUnitsCounter)
//    {
//        goto Cleanup;
//    }
//
//    _uOpenUnitsCounter--;
//
//    {
//        CDeferredSelectionUndo DeferredUndo(_pPrimaryMarkup);
//    }
//
//    if(_uOpenUnitsCounter == 0)
//    {
//        Assert(_pMarkupServicesParentUndo);
//
//        hr = _pMarkupServicesParentUndo->Finish(S_OK);
//
//        delete _pMarkupServicesParentUndo;
//    }
//
//Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::CreateMarkupPointer(CMarkupPointer** ppPointer)
{
    Assert(ppPointer);

    *ppPointer = new CMarkupPointer(this);
    if(!*ppPointer)
    {
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

HRESULT CDocument::InsertElement(
             CElement*          pElementInsert,
             CMarkupPointer*    pPointerStart,
             CMarkupPointer*    pPointerFinish,
             DWORD              dwFlags)
{
    HRESULT hr = S_OK;
    CTreePosGap tpgStart, tpgFinish;

    Assert(pElementInsert);
    Assert(pPointerStart);

    // If the the finish is not specified, set it to the start to make the element span
    // nothing at the start.
    if(!pPointerFinish)
    {
        pPointerFinish = pPointerStart;
    }

    Assert(!pElementInsert->GetFirstBranch());

    // If the element is no scope, then we must ignore the finish
    if(pElementInsert->IsNoScope())
    {
        pPointerFinish = pPointerStart;
    }

    // Make sure the start if before the finish
    Assert(pPointerStart->IsLeftOfOrEqualTo(pPointerFinish));

    // Make sure both pointers are positioned, and in the same tree
    Assert(pPointerStart->IsPositioned());
    Assert(pPointerFinish->IsPositioned());
    Assert(pPointerStart->Markup() == pPointerFinish->Markup());

    // Make sure unembedded markup pointers go in for the modification
    hr = pPointerStart->Markup()->EmbedPointers();
    if(hr)
    {
        goto Cleanup;
    }

    // Position the gaps and do the insert

    // Note: We embed to make sure the pointers get updated, but we
    // also take advantage of it to get pointer pos's for the input
    // args.  It would be nice to treat the inputs specially in the
    // operation and not have to embed them......
    tpgStart.MoveTo(pPointerStart->GetEmbeddedTreePos(), TPG_LEFT);
    tpgFinish.MoveTo(pPointerFinish->GetEmbeddedTreePos(), TPG_LEFT);

    hr = pPointerStart->Markup()->InsertElementInternal(pElementInsert, &tpgStart,
        &tpgFinish, dwFlags);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::RemoveElement(CElement* pElementRemove, DWORD dwFlags)
{
    HRESULT hr = S_OK;

    // Element to be removed must be specified and it must be associated
    // with this document.
    Assert(pElementRemove);

    // Assert that the element is in the markup
    Assert(pElementRemove->IsInMarkup());

    // Now, remove the element
    hr = pElementRemove->GetMarkup()->RemoveElementInternal(pElementRemove, dwFlags);
    if(hr)
    {
        goto Cleanup;
    }
Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::Remove(CMarkupPointer* pPointerStart, CMarkupPointer* pPointerFinish, DWORD dwFlags)
{
    return CutCopyMove(pPointerStart, pPointerFinish, NULL, TRUE, dwFlags);
}

HRESULT CDocument::Copy(
            CMarkupPointer* pPointerSourceStart,
            CMarkupPointer* pPointerSourceFinish,
            CMarkupPointer* pPointerTarget,
            DWORD           dwFlags)
{
    return CutCopyMove(pPointerSourceStart, pPointerSourceFinish,
        pPointerTarget, FALSE, dwFlags);
}

HRESULT CDocument::Move(
            CMarkupPointer* pPointerSourceStart,
            CMarkupPointer* pPointerSourceFinish,
            CMarkupPointer* pPointerTarget,
            DWORD           dwFlags)
{
    return CutCopyMove(pPointerSourceStart, pPointerSourceFinish,
        pPointerTarget, TRUE, dwFlags);
}

HRESULT CDocument::InsertText(
              CMarkupPointer*   pPointerTarget,
              const OLECHAR*    pchText,
              long              cch,
              DWORD             dwFlags)
{
    HRESULT hr = S_OK;

    Assert(pPointerTarget);
    Assert(pPointerTarget->IsPositioned());

    hr = pPointerTarget->Markup()->EmbedPointers();
    if(hr)
    {
        goto Cleanup;
    }

    if(cch < 0)
    {
        cch = pchText ? _tcslen(pchText) : 0;
    }

    hr = pPointerTarget->Markup()->InsertTextInternal(pPointerTarget->GetEmbeddedTreePos(),
        pchText, cch, dwFlags);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

static HRESULT PrepareStream(
           HGLOBAL          hGlobal,
           IStream**        ppIStream,
           BOOL             fInsertFrags,
           HTMPASTEINFO*    phtmpasteinfo)
{
    CStreamReadBuff* pstreamReader = NULL;
    LPSTR   pszGlobal = NULL;
    long    lSize;
    BOOL    fIsEmpty, fHasSignature;
    TCHAR   szVersion[24];
    TCHAR   szSourceUrl[pdlUrlLen];
    long    iStartHTML, iEndHTML;
    long    iStartFragment, iEndFragment;
    long    iStartSelection, iEndSelection;
    ULARGE_INTEGER ul = { 0, 0 };
    HRESULT hr = S_OK;

    Assert(hGlobal);
    Assert(ppIStream);

    // Get access to the bytes of the global
    pszGlobal = LPSTR(GlobalLock(hGlobal));

    if(!pszGlobal)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    // First, compute the size of the global
    lSize = 0;

    // BUGBUG: Support nonnative signature and sizeof(wchar) == 4/2 (davidd)
    fHasSignature = *(WCHAR*)pszGlobal == NATIVE_UNICODE_SIGNATURE;

    if(fHasSignature)
    {
        lSize = _tcslen((TCHAR*)pszGlobal) * sizeof(TCHAR);
        fIsEmpty = (lSize-sizeof(TCHAR)) == 0;
    }
    else
    {
        lSize = lstrlenA(pszGlobal);
        fIsEmpty = lSize == 0;
    }

    // If the HGLOBAL is effectively empty, do nothing, and return this
    // fact.
    if(fIsEmpty)
    {
        *ppIStream = NULL;
        goto Cleanup;
    }

    // Create if the stream got the load context
    hr = CreateStreamOnHGlobal(hGlobal, FALSE, ppIStream);

    if(hr)
    {
        goto Cleanup;
    }

    // N.B. (johnv) This is necessary for Win95 support.  Apparently
    // GlobalSize() may return different values at different times for the same
    // hGlobal.  This makes IStream's behavior unpredictable.  To get around
    // this, we set the size of the stream explicitly.

    // Make sure we don't have unicode in the hGlobal
    ul.QuadPart = lSize;
    Assert(ul.QuadPart <= GlobalSize(hGlobal));

    hr = (*ppIStream)->SetSize(ul);

    if(hr)
    {
        goto Cleanup;
    }

    iStartHTML      = -1;
    iEndHTML        = -1;
    iStartFragment  = -1;
    iEndFragment    = -1;
    iStartSelection = -1;
    iEndSelection   = -1;

    // Locate the required contextual information in the stream
    pstreamReader = new CStreamReadBuff(*ppIStream, CP_UTF_8);
    if(pstreamReader == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = pstreamReader->GetStringValue(_T("Version"), szVersion, ARRAYSIZE(szVersion));

    if(hr == S_FALSE)
    {
        goto PlainStream;
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = pstreamReader->GetLongValue(_T("StartHTML"), &iStartHTML);

    if(hr == S_FALSE)
    {
        goto PlainStream;
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = pstreamReader->GetLongValue(_T("EndHTML"), &iEndHTML);

    if(hr == S_FALSE)
    {
        goto PlainStream;
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = pstreamReader->GetLongValue(_T("StartFragment"), &iStartFragment);

    if(hr == S_FALSE)
    {
        goto PlainStream;
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = pstreamReader->GetLongValue(_T("EndFragment"), &iEndFragment);

    // Locate optional contextual information
    hr = pstreamReader->GetLongValue(_T("StartSelection"), &iStartSelection);

    if(hr && hr!=S_FALSE)
    {
        goto Cleanup;
    }

    if(hr != S_FALSE)
    {
        hr = pstreamReader->GetLongValue(_T("EndSelection"), &iEndSelection);

        if(hr && hr!=S_FALSE)
        {
            goto Cleanup;
        }

        if(hr == S_FALSE)
        {
            hr = E_FAIL;
            goto Cleanup;
        }
    }
    else
    {
        iStartSelection = -1;
    }

    // Get the source URL info
    hr = pstreamReader->GetStringValue(_T("SourceURL"), szSourceUrl, ARRAYSIZE(szSourceUrl));

    if(hr && hr!=S_FALSE)
    {
        goto Cleanup;
    }

    if(phtmpasteinfo && hr!=S_FALSE)
    {
        hr = phtmpasteinfo->cstrSourceUrl.Set(szSourceUrl);

        if(hr)
        {
            goto Cleanup;
        }
    }

    // Make sure contextual info is sane
    if(iStartHTML<0 && iEndHTML<0)
    {
        // per cfhtml spec, start and end html can be -1 if there is no
        // context.  there must always be a fragment, however
        iStartHTML = iStartFragment;
        iEndHTML   = iEndFragment;
    }

    if (iStartHTML     < 0 || iEndHTML     < 0 ||
        iStartFragment < 0 || iEndFragment < 0 ||
        iStartHTML     > iEndHTML     ||
        iStartFragment > iEndFragment ||
        iStartHTML > iStartFragment ||
        iEndHTML   < iEndFragment)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    if(iStartSelection != -1)
    {
        if (iEndSelection < 0 ||
            iStartSelection > iEndSelection ||
            iStartHTML > iStartSelection ||
            iEndHTML < iEndSelection ||
            iStartFragment > iStartSelection ||
            iEndFragment < iEndSelection)
        {
            hr = E_FAIL;
            goto Cleanup;
        }
    }

    // Rebase the fragment and selection off of the start html
    iStartFragment -= iStartHTML;
    iEndFragment -= iStartHTML;

    if(iStartSelection != -1)
    {
        iStartSelection -= iStartHTML;
        iEndSelection -= iStartHTML;
    }
    else
    {
        iStartSelection = iStartFragment;
        iEndSelection = iEndFragment;
    }

    phtmpasteinfo->cbSelBegin  = iStartSelection;
    phtmpasteinfo->cbSelEnd    = iEndSelection;

    pstreamReader->SetPosition(iStartHTML);

    hr = S_OK;

Cleanup:
    if(pstreamReader)
    {
        phtmpasteinfo->cp = pstreamReader->GetCodePage();
    }

    delete pstreamReader;

    if(pszGlobal)
    {
        GlobalUnlock(hGlobal);
    }

    RRETURN(hr);

PlainStream:
    pstreamReader->SetPosition(0);
    pstreamReader->SwitchCodePage(_afxGlobalData._cpDefault);

    if(fInsertFrags)
    {
        phtmpasteinfo->cbSelBegin  = fHasSignature ? sizeof(TCHAR) : 0;
        phtmpasteinfo->cbSelEnd    = lSize;
    }
    else
    {
        phtmpasteinfo->cbSelBegin  = -1;
        phtmpasteinfo->cbSelEnd    = -1;
    }

    hr = S_OK;

    goto Cleanup;
}

static HRESULT MoveToPointer(CDocument* pDoc, CMarkupPointer* pMarkupPointer, CTreePos* ptp)
{
    HRESULT hr = S_OK;

    if(!pMarkupPointer)
    {
        if(ptp)
        {
            hr = ptp->GetMarkup()->RemovePointerPos(ptp, NULL, NULL);

            if(hr)
            {
                goto Cleanup;
            }
        }

        goto Cleanup;
    }

    if(!ptp)
    {
        hr = pMarkupPointer->Unposition();

        if(hr)
        {
            goto Cleanup;
        }

        goto Cleanup;
    }

    hr = pMarkupPointer->MoveToOrphan(ptp);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:   CDocument::TakeFocus
//
//  Synopsis: To have trident window take focus if it does not already have it.
//
//--------------------------------------------------------------------------
BOOL CDocument::TakeFocus()
{
    BOOL fRet = FALSE;

    if(_pInPlace && !_pInPlace->_fDeactivating)
    {
        if(::GetFocus() != _pInPlace->_hwnd)
        {
            SetFocusWithoutFiringOnfocus();
            fRet = TRUE;
        }
    }

    return fRet;
}

static BOOL CALLBACK CallBackEnumChild(HWND hwnd, LPARAM lparam)
{
    *(BOOL*)lparam = (::GetFocus() == hwnd);

    return !(*(BOOL*)lparam);
}

BOOL CDocument::HasFocus()
{
    BOOL fHasFocus = FALSE;

    if(_pInPlace && _pInPlace->_hwnd)
    {
        Assert(IsWindow(_pInPlace->_hwnd));
        fHasFocus = (::GetFocus() == _pInPlace->_hwnd);
        if(!fHasFocus)
        {
            EnumChildWindows(_pInPlace->_hwnd, CallBackEnumChild, (LPARAM)&fHasFocus);
        }
    }

    return fHasFocus;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::InvalidateDefaultSite
//
//  Synopsis:   Invalidate the current site
//
//----------------------------------------------------------------------------
HRESULT CDocument::InvalidateDefaultSite()
{
    CElement* pElemDefault;
    if(!_pElemCurrent->_fDefault)
    {
        pElemDefault = _pElemCurrent->FindDefaultElem(TRUE);
        if(pElemDefault && pElemDefault!=_pElemCurrent)
        {
            pElemDefault->Invalidate();
        }
    }

    RRETURN(S_OK);
}

HRESULT CDocument::CreateElement(ELEMENT_TAG etag, CElement** ppElementNew)
{
    return PrimaryMarkup()->CreateElement(etag, ppElementNew);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::BroadcastNotify
//
//  Synopsis:   Broadcast this notification through the tree
//
//----------------------------------------------------------------------------
void CDocument::BroadcastNotify(CNotification* pNF)
{
    Assert(pNF);

    CMarkup* pMarkup = pNF->Element()->GetMarkup();

    if(pMarkup)
    {
        pMarkup->Notify(pNF);
    }
}

//+--------------------------------------------------------------
//
//  Member:     CDocument::DocTraverseGroup
//
//  Synopsis:   Called by (e.g.)a radioButton, this function
//      takes the groupname and queries the document's collection for the
//      rest of the group and calls the provided CLEARGROUP function on that
//      element. The traversal stops if the visit function returns S_OK or
//      and an error.
//
//---------------------------------------------------------------
HRESULT CDocument::DocTraverseGroup(LPCTSTR strGroupName, PFN_VISIT pfn, DWORD_PTR dw, BOOL fForward)
{
    HRESULT             hr;
    long                i, c;
    CElement*           pElement;
    CCollectionCache*   pCollectionCache;
    LPCTSTR             lpName;

    _fInTraverseGroup = TRUE;

    // The control is at the document level. Clear all other controls,
    // also at document level, which are in the same group as this control.
    Assert(_pPrimaryMarkup);
    hr = _pPrimaryMarkup->EnsureCollectionCache(CMarkup::ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    pCollectionCache = _pPrimaryMarkup->CollectionCache();
    Assert(pCollectionCache);

    // get size of collection
    c = pCollectionCache->SizeAry(CMarkup::ELEMENT_COLLECTION);

    if(fForward)
    {
        i = 0;
    }
    else
    {
        i = c - 1;
    }

    // if nothing is in the collection, default answer is S_FALSE.
    hr = S_FALSE;

    while(c--)
    {
        hr = pCollectionCache->GetIntoAry(CMarkup::ELEMENT_COLLECTION, i, &pElement);
        if(fForward)
        {
            i++;
        }
        else
        {
            i--;
        }
        if(hr)
        {
            goto Cleanup;
        }

        hr = S_FALSE; // default answer again.

        // ignore the element if it is not a site
        if(!pElement->NeedsLayout())
        {
            continue;
        }

        lpName = pElement->GetAAname();

        // is this item in the target group?
        if(lpName && FormsStringICmp(strGroupName, lpName)==0)
        {
            // Call the function and break out of the
            // loop if it doesn't return S_FALSE.
            hr = CALL_METHOD(pElement, pfn, (dw));
            if(hr != S_FALSE)
            {
                break;
            }
        }
    }

Cleanup:
    _fInTraverseGroup = FALSE;
    RRETURN1(hr, S_FALSE);
}

//+--------------------------------------------------------------
//
//  Member:     CDocument::FindDefaultElem
//
//  Synopsis:   find the default/Cancel button in the Doc
//
//              fCurrent means looking for the current default layout
//
//---------------------------------------------------------------
CElement* CDocument::FindDefaultElem(BOOL fDefault, BOOL fCurrent/*=FALSE*/)
{
    HRESULT             hr      = S_FALSE;
    long                c       = 0;
    long                i       = 0;
    CElement*           pElem   = NULL;
    CCollectionCache*   pCollectionCache;

    Assert(_pPrimaryMarkup);
    hr = _pPrimaryMarkup->EnsureCollectionCache(CMarkup::ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    pCollectionCache = _pPrimaryMarkup->CollectionCache();
    Assert(pCollectionCache);

    // Collection walker cached the value on the doc, go get it

    // get size of collection
    c = pCollectionCache->SizeAry(CMarkup::ELEMENT_COLLECTION);

    while(c--)
    {
        hr = pCollectionCache->GetIntoAry(CMarkup::ELEMENT_COLLECTION, i++, &pElem);

        if(hr)
        {
            pElem = NULL;
            goto Cleanup;
        }

        if(!pElem || pElem->_fExittreePending)
        {
            continue;
        }

        if(fCurrent)
        {
            if(pElem->_fDefault)
            {
                goto Cleanup;
            }
            continue;
        }

        if(pElem->TestClassFlag(fDefault?CElement::ELEMENTDESC_DEFAULT:CElement::ELEMENTDESC_CANCEL)
            && pElem->IsVisible(TRUE) && pElem->IsEnabled())
        {
            goto Cleanup;
        }
    }
    pElem = NULL;

Cleanup:
    return pElem;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::GetClassDesc, CBase
//
//  Synopsis:   Return the class descriptor.
//
//-------------------------------------------------------------------------
const CBase::CLASSDESC* CDocument::GetClassDesc() const
{
    return (CBase::CLASSDESC*)&s_classdesc;
}

extern void ClearSurfaceCache();
//+-------------------------------------------------------------------------
//
//  Method:     CDocument::Passivate
//
//  Synopsis:   Shutdown main object by releasing references to
//              other objects and generally cleaning up.  This
//              function is called when the main reference count
//              goes to zero.  The destructor is called when
//              the reference count for the main object and all
//              embedded sub-objects goes to zero.
//
//              Release any event connections held by the form.
//
//--------------------------------------------------------------------------
void CDocument::Passivate()
{
    if(_fPassivateCalled)
    {
        Assert("Passivate called too many times " && _fPassivateCalled);
        return;
    }

    _fPassivateCalled = TRUE;

    // behaviors support
    /*ClearInterface(&_pDefaultPeerFactory);
    ClearInterface(&_pHostPeerFactory); wlw note*/

    // When my last reference is released, don't accept any excuses
    // while shutting down
    _fForceCurrentElem = TRUE;

    ClearInterface(&_pTimerDraw);

    // Unload the contents of the document
    UnloadContents(FALSE, FALSE);

    // Tear down the editor
    _EditRouter.Passivate();
    ReleaseInterface(_pIHTMLEditor);
    _pIHTMLEditor = NULL;

    // Release caret
    ReleaseInterface(_pCaret);
    _pCaret = NULL;

    if(_pPrimaryMarkup)
    {
        Assert(_pElemCurrent == _pPrimaryMarkup->Root());
        _pElemCurrent = _pElementDefault;
        _pPrimaryMarkup->Release();
        _pPrimaryMarkup = NULL;
    }

    Assert(_pElementDefault);

    CElement::ClearPtr((CElement**)&_pElementDefault);

    ClearInterface(&_pDownloadNotify);

    // Now, we can safely shut down the form.
    /*if(_pOmWindow)
    {
        _pOmWindow->_fTrusted = FALSE;

        _pOmWindow->PrivateRelease();   // because doc always holds private reference on OmWindow,
                                        // the last release should be PrivateRelease
        _pOmWindow = NULL;
    } wlw note*/

    GWKillMethodCall(this, NULL, 0);

    // release caches
    ClearSurfaceCache();

    ClearDefaultCharFormat();

    NotifySelection(SELECT_NOTIFY_DOC_ENDED, NULL);

    CServer::Passivate();
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnPropertyChange
//
//  Synopsis:   Invalidate, fire property change, and so on.
//
//  Arguments:  [dispidProperty] -- PROPID of property that changed
//              [dwFlags]        -- Flags to inhibit behavior
//
//----------------------------------------------------------------------------
HRESULT CDocument::OnPropertyChange(DISPID dispidProperty, DWORD dwFlags)
{
    HRESULT hr = S_OK;

    switch(dispidProperty)
    {
    case DISPID_BACKCOLOR:
        OnAmbientPropertyChange(DISPID_AMBIENT_BACKCOLOR);
        break;

    case DISPID_FORECOLOR:
        OnAmbientPropertyChange(DISPID_AMBIENT_FORECOLOR);
        break;

    case DISPID_A_DIR:
        OnAmbientPropertyChange(DISPID_AMBIENT_RIGHTTOLEFT);
        break;
    }

    if((dwFlags&(SERVERCHNG_NOVIEWCHANGE|FORMCHNG_NOINVAL)) == 0)
    {
        Invalidate();
    }

    if(dwFlags & ELEMCHNG_CLEARCACHES )
    {
        PrimaryRoot()->GetFirstBranch()->VoidCachedInfo();
    }

    if(dwFlags & FORMCHNG_LAYOUT)
    {
        PrimaryRoot()->ResizeElement(NFLAGS_FORCE);
    }

    Verify(!CServer::OnPropertyChange(dispidProperty, dwFlags));

    // see bugbug in CElement:OnPropertyChange
    // if ( fSomeoneIsListening )

    // Post the call to fire onpropertychange if dispid is that of activeElement, this
    // is so that the order of event firing is maintained in SetCurrentElem.
    if(dispidProperty == DISPID_CDoc_activeElement)
    {
        hr = GWPostMethodCall(this, ONCALL_METHOD(CDocument, FirePostedOnPropertyChange, firepostedonpropertychange), 0, TRUE, "CDoc::FirePostedOnPropertyChange");
    }
    else
    {
        hr = Fire_PropertyChangeHelper(dispidProperty, dwFlags);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::Draw, CServer
//
//  Synopsis:   Draw the form.
//              Called from CServer implementation if IViewObject::Draw
//
//----------------------------------------------------------------------------
HRESULT CDocument::Draw(CDrawInfo* pDI, RECT* prc)
{
    CServer::CLock  Lock(this, SERVERLOCK_IGNOREERASEBKGND);
    CSaveTransform  st(_dci);
    CFormDrawInfo   DI;
    int             r;
    CSize           sizeOld;
    CSize           sizeNew;
    CPoint          ptOrgOld;
    CPoint          ptOrg;
    HRESULT         hr = S_OK;
    BOOL            fHadView = !!_view.IsActive();

    if(DVASPECT_CONTENT!=pDI->_dwDrawAspect && DVASPECT_DOCPRINT!=pDI->_dwDrawAspect)
    {
        RRETURN(DV_E_DVASPECT);
    }

    // Do not display if we haven't received all the stylesheets
    if(IsStylesheetDownload())
    {
        return hr;
    }

    // Setup drawing info.
    DI.Init(GetPrimaryElementTop());

    // Copy the CDrawInfo information into CFormDrawInfo
    *(CDrawInfo*)&DI = *pDI;

    ::SetViewportOrgEx(DI._hdc, 0, 0, &ptOrg);

    ((CRect*)prc)->OffsetRect(ptOrg.AsSize());

    DI._pDoc     = this;
    DI._hrgnClip = CreateRectRgnIndirect(prc);

    r = GetClipRgn(DI._hdc, DI._hrgnClip);
    if(r == -1)
    {
        DeleteObject(DI._hrgnClip);
        DI._hrgnClip = NULL;
    }

    r = GetClipBox(DI._hdc, &DI._rcClip);
    if(r == 0)
    {
        // No clip box, assume very large clip box to start.
        DI._rcClip.left   = DI._rcClip.top    = SHRT_MIN;
        DI._rcClip.right  = DI._rcClip.bottom = SHRT_MAX;
    }
    DI._rcClipSet = DI._rcClip;

    DI.CTransform::Init(prc, _sizel);

    // We are in print preview mode if DI._hic is
    // null and DI._ptd is not null.  Setup a second draw info
    // with information about the printer.
    if(!DI._hic)
    {
        if(!DI.Printing())
        {
            DI._hic = DI.IsMetafile() ? TLS(hdcDesktop) : DI._hdc;
        }
        else
        {
            DI._hic =  CreateIC(
                (LPCTSTR)((char*)DI._ptd+DI._ptd->tdDriverNameOffset),
                (LPCTSTR)((char*)DI._ptd+DI._ptd->tdDeviceNameOffset),
                (LPCTSTR)((char*)DI._ptd+DI._ptd->tdPortNameOffset),
                NULL);

            if(NULL == DI._hic)
            {
                // Couldn't create it
                hr = E_FAIL;
                goto Cleanup;
            }

            DI._sizeInch.cx = GetDeviceCaps(DI._hic, LOGPIXELSX);
            DI._sizeInch.cy = GetDeviceCaps(DI._hic, LOGPIXELSY);

            // BUGBUG (garybu) Need to set these sizes up to match the printer.
            // BUGBUG (cthrash) this isn't correct.
            DI._sizeSrc.cx = 2540;
            DI._sizeSrc.cy = 2540;
            DI._sizeDst.cx = DI._sizeInch.cx;
            DI._sizeDst.cy = DI._sizeInch.cy;
        }
    }

    DI._sizeInch.cx = GetDeviceCaps(DI._hic, LOGPIXELSX);
    DI._sizeInch.cy = GetDeviceCaps(DI._hic, LOGPIXELSY);

    // REVIEW sidda:    hack to force scaling from display pixels to printer pixels when printing
    DI.SetScaleFraction(1, 1, 1);
    DI.SetResolutions(&_afxGlobalData._sizePixelsPerInch, &DI._sizeInch);

    _dci = *((CDocInfo*)&DI);

    {
        CRegion rgnInvalid;
        // Update layout if needed
        // NOTE: Normally, this code should allow the paint to take place (in fact, it should force one), but doing so
        //       can cause problems for clients that use IVO::Draw at odd times.
        //       So, instead of pushing the paint through, the simple collects the invalid region and holds on to it
        //       until it's safe (see below). (brendand)
        if(!fHadView)
        {
            _view.Activate();
        }
        else
        {
            _view.EnsureView(LAYOUT_NOBACKGROUND|LAYOUT_SYNCHRONOUS
                |LAYOUT_DEFEREVENTS|LAYOUT_DEFERINVAL|LAYOUT_DEFERPAINT);

            rgnInvalid = _view._rgnInvalid;
            _view.ClearInvalid();
        }

        _view.GetViewSize(&sizeOld);
        sizeNew = ((CRect*)prc)->Size();

        _view.SetViewSize(sizeNew);

        _view.GetViewPosition(&ptOrgOld);
        _view.SetViewPosition(ptOrg);

        // In all cases, ensure the view is up-to-date with the passed dimensions
        // (Do not invalidate the in-place HWND (if any) or do anything else significant as result (e.g., fire events)
        // since the information backing it all is transient and only relevant to this request.)
        _view.EnsureView(LAYOUT_FORCE|LAYOUT_NOBACKGROUND|LAYOUT_SYNCHRONOUS
            |LAYOUT_DEFEREVENTS|LAYOUT_DEFERENDDEFER|LAYOUT_DEFERINVAL|LAYOUT_DEFERPAINT);
        _view.ClearInvalid();

        // Render the sites.
        _view.RenderView(&DI, DI._hrgnClip, &DI._rcClip);

        // Restore layout if inplace.
        // (Now, bring the view et. al. back in-sync with the document. Again, do not force a paint or any other
        // significant work during this layout pass. Once it completes, however, re-establish any held invalid
        // region and post a call such that layout/paint does eventually occur.)
        if(fHadView)
        {
            _view.SetViewSize(sizeOld);
            _view.SetViewPosition(ptOrgOld);
            _view.EnsureView(LAYOUT_FORCE|LAYOUT_NOBACKGROUND|LAYOUT_SYNCHRONOUS
                |LAYOUT_DEFEREVENTS|LAYOUT_DEFERENDDEFER|LAYOUT_DEFERINVAL|LAYOUT_DEFERPAINT);
            _view.ClearInvalid();

            if(!rgnInvalid.IsEmpty())
            {
                _view.OpenView();
                _view._rgnInvalid = rgnInvalid;
            }
        }
        else
        {
            _view.Deactivate();
        }
    }

Cleanup:
    if(DI._hrgnClip)
    {
        DeleteObject(DI._hrgnClip);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     DoTranslateAccelerator, public
//
//  Synopsis:   Overridden method of IOleInPlaceActiveObject
//
//  Arguments:  [lpmsg] -- Message to translate
//
//  Returns:    S_OK if translated, S_FALSE if not.  Error otherwise
//
//  History:    07-Feb-94     LyleC    Created
//              08-Feb-95     AndrewL  Turned off noisy trace.
//
//----------------------------------------------------------------------------
HRESULT CDocument::DoTranslateAccelerator(LPMSG lpmsg)
{
    HRESULT  hr = S_FALSE;
    CMessage Message(lpmsg);

    Assert(_pElemCurrent && _pElemCurrent->GetFirstBranch());

    CTreeNode::CLock Lock(_pElemCurrent->GetFirstBranch());

    switch(lpmsg->message)
    {
    case WM_KEYDOWN:
    case WM_KEYUP:
    case WM_SYSKEYDOWN:
    case WM_SYSKEYUP:
        break;
    default:
        goto Cleanup;
    }

    // Do not handle any messages if our window or one of our children
    // does not have focus.

    // BUGBUG (sujalp): We should not be dispatching this message from here.
    // It should be dispatched only from CDoc::OnWindowMessage. When this
    // problem is fixed up, remove the following switch.
    switch(lpmsg->message)
    {
    case WM_KEYDOWN:
        _fGotKeyDown = TRUE;
        // if we are Tabbing into Trident the first time directly from address bar
        // do not tab into address bar again on next TAB. (sramani: see bug#28426)
        if(lpmsg->wParam==VK_TAB && ::GetFocus()!=_pInPlace->_hwnd)
        {
            _fFirstTimeTab = FALSE;
        }
        break;

    case WM_KEYUP:
        if(!_fGotKeyDown)
        {
            return S_FALSE;
        }
        break;
    }

    // If there was no captured object, or if the capture object did not
    // handle the message, then lets pass it to the current site
    hr = PumpMessage(&Message, _pElemCurrent->GetFirstBranch(), TRUE);

Cleanup:
    RRETURN1_NOTRACE(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDoc::FireEvent
//
//  Synopsis:   Lock the document while firing an event
//
//  return value :   S_OK  -- normal return value.  and event did not return
//                           false
//                  S_FALSE - event explictly returned false
//----------------------------------------------------------------------------
HRESULT CDocument::FireEventHelper(DISPID dispidEvent, DISPID dispidProp, BYTE* pbTypes, ...)
{
    va_list         valParms;
    HRESULT         hr = S_OK;
    CVariant        Var;

    CLock Lock(this);

    va_start(valParms, pbTypes);

    hr = FireEventV(
        dispidEvent,
        dispidProp,
        NULL,
        &Var,
        pbTypes,
        valParms);
    va_end(valParms);

    if(V_VT(&Var) == VT_BOOL)
    {
        hr = (V_BOOL(&Var)!=VB_TRUE) ? S_FALSE : S_OK;
    }

    RRETURN1(hr, S_FALSE);
}

HRESULT CDocument::FireEvent(DISPID dispidEvent, DISPID dispidProp, LPCTSTR pchEventType, BYTE* pbTypes, ...)
{
    EVENTPARAM param(this, TRUE);
    HRESULT hr;

    param.SetType(pchEventType);

    hr = FireEventHelper(dispidEvent, dispidProp, pbTypes);

    // mask off S_FALSE since callers of this fx aren't concerned with
    // canceled return values.
    RRETURN((hr==S_FALSE)?S_OK:hr);
}

//+--------------------------------------------------------------
//
//  Member:     CDocument::HandleKeyNavigate
//
//  Synopsis:   Navigate to the next CSite/CElement that can take focus
//              based on the message.
//
//---------------------------------------------------------------
HRESULT CDocument::HandleKeyNavigate(CMessage* pmsg, BOOL fAccessKeyNeedCycle)
{
    HRESULT         hr = S_FALSE;
    FOCUS_DIRECTION dir = (pmsg->dwKeyState&FSHIFT) ? DIRECTION_BACKWARD : DIRECTION_FORWARD;
    CElement*       pCurrent = NULL;
    CElement*       pNext = NULL;
    CElement*       pStart = NULL;
    long            lSubNew = _lSubCurrent;
    BOOL            fFindNext = FALSE;
    CElement*       pElementClient = GetPrimaryElementClient();

    // It is possible to have site selection even at browse time, if the user
    // is tabbing in an editable container such as HTMLAREA.
    BOOL            fSiteSelected = (GetSelectionType() == SELECTION_TYPE_Control);
    BOOL            fSiteSelectMode = fSiteSelected;
    BOOL            fYieldFailed = FALSE;

    if(!pElementClient)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    if(IsFrameTabKey(pmsg))
    {
        if(pmsg->lParam)
        {
            hr = S_FALSE;
        }
        else
        {
            hr = pElementClient->BecomeCurrentAndActive(pmsg, pmsg->lSubDivision, TRUE);
        }
        goto Cleanup;
    }

    // Detect time when this is the first time tabbing into the doc
    // and bail out in that case.  Let the shell take focus.
    if(_fFirstTimeTab)
    {
        hr = S_FALSE;
        _fFirstTimeTab = FALSE;
        goto Cleanup;
    }

    // First find the element from which to start
    if((!fSiteSelectMode && _pElemCurrent==PrimaryRoot()) ||
        (fAccessKeyNeedCycle && pmsg->message==WM_SYSKEYDOWN) ||
        (IsTabKey(pmsg) && !pmsg->lParam))
    {
        if((IsTabKey(pmsg)) && (!pmsg->lParam) && (DIRECTION_FORWARD==dir))
        {
            // Raid 61972
            // we just tab down to a frame CBodyElement, need to set the flag
            // so we know we need to tab out when SHIFT+_TAB
            _fNeedTabOut = TRUE;
        }
        pStart = NULL;
        fFindNext = TRUE;
    }
    else if(!fSiteSelectMode && _pElemCurrent->IsTabbable(_lSubCurrent))
    {
        pStart = _pElemCurrent;
        fFindNext = TRUE;
    }
    else if(_pElemCurrent->Tag() != ETAG_ROOT)
    {
        // Tab to the element next to the caret/selection (unless the root element is
        // current - bug #65023)
        IHTMLEditor*    pEditor     = GetHTMLEditor(TRUE);
        IHTMLElement*   pIElement   = NULL;
        CElement*       pElemStart  = NULL;
        BOOL            fNext       = TRUE;

        Assert(pEditor);
        if(pEditor && S_OK==pEditor->GetElementToTabFrom(dir==DIRECTION_FORWARD, &pIElement, &fNext) && pIElement)
        {
            Verify(S_OK == pIElement->QueryInterface(CLSID_CElement, ( void**)&pElemStart));
            Assert(pElemStart);

            if(pElemStart->Tag() == ETAG_TXTSLAVE)
            {
                pElemStart = pElemStart->MarkupMaster();
            }
            if(pElemStart)
            {
                pStart = pElemStart;
                fFindNext = fNext || !pStart->IsTabbable(_lSubCurrent);
                if(!fFindNext)
                {
                    pCurrent = pStart;
                }
            }
            pIElement->Release();
        }
        if(!pElemStart)
        {
            pStart = _pElemCurrent;
            fFindNext = TRUE;
        }
    }
    else
    {
        pStart = NULL;
        fFindNext = TRUE;
    }

    if(pStart && pStart==pElementClient && !pStart->IsTabbable(_lSubCurrent))
    {
        pStart = NULL;
    }

    if(fFindNext)
    {
        Assert(!pCurrent);
        if(IsTabKey(pmsg) || IsFrameTabKey(pmsg))
        {
            FindNextTabOrder(dir, pStart, _lSubCurrent, &pCurrent, &lSubNew);
        }
        else
        {
            SearchAccessKeyArray(dir, pStart, &pCurrent, pmsg);
        }
    }

    hr = S_FALSE;
    if(!pCurrent)
    {
        goto Cleanup;
    }

    Assert(!IsFrameTabKey(pmsg));

    if(IsTabKey(pmsg) || IsFrameTabKey(pmsg))
    {
        do
        {
            if(pNext)
            {
                pCurrent = pNext;
            }

            pmsg->lSubDivision = lSubNew;

            // Raid 61972
            // Set _fNeedTabOut if pCurrent is a CBodyElement so that we will
            // not try to SHIFT+TAB to CBodyElement again.
            _fNeedTabOut = (pCurrent->Tag()==ETAG_BODY) ? TRUE : FALSE;

            hr = pCurrent->HandleMnemonic(pmsg, FALSE, &fYieldFailed);
            Assert(!(hr==S_OK && fYieldFailed));
            if(hr && fYieldFailed)
            {
                hr = S_OK;
            }
        } while(hr==S_FALSE && FindNextTabOrder(dir, pCurrent, pmsg->lSubDivision, &pNext, &lSubNew) && pNext);
    }
    else
    {
        // accessKey case here
        _fNeedTabOut = FALSE;
        do
        {
            if(pNext)
            {
                pCurrent = pNext;
            }

            FOCUS_ITEM fi = pCurrent->GetMnemonicTarget();

            if(fi.pElement)
            {
                if(fi.pElement == pCurrent)
                {
                    pmsg->lSubDivision = fi.lSubDivision;
                    hr = fi.pElement->HandleMnemonic(pmsg, TRUE);
                }
                else
                {
                    CMessage msg;

                    msg.message = WM_KEYDOWN;
                    msg.wParam = VK_TAB;
                    msg.lSubDivision = fi.lSubDivision;
                    hr = fi.pElement->HandleMnemonic(&msg, TRUE);
                }
            }

            if(hr == S_FALSE)
            {
                SearchAccessKeyArray(dir, pCurrent, &pNext, pmsg);
            }
        } while(hr==S_FALSE && pNext);
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

HRESULT CDocument::TranslateAccelerator(LPMSG lpmsg)
{
    HRESULT hr = S_FALSE;    

    //  Give tooltips a chance to dismiss.
    {
        // Ignore spurious WM_ERASEBACKGROUNDs generated by tooltips
        CServer::CLock Lock(this, SERVERLOCK_IGNOREERASEBKGND);

        FormsTooltipMessage(lpmsg->message, lpmsg->wParam, lpmsg->lParam);
    }

    if(lpmsg->message<WM_KEYFIRST || lpmsg->message>WM_KEYLAST)
    {
        return S_FALSE;
    }

    hr = super::TranslateAccelerator(lpmsg);

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::RunningToInPlace
//
//  Synopsis:   Effects the running to inplace-active state transition
//
//  Returns:    SUCCESS if the object results in the in-place state
//
//  Notes:      This method in-place activates all inside-out embeddings
//              in addition to the normal CServer base processing.
//
//---------------------------------------------------------------
HRESULT CDocument::RunningToInPlace(HWND hWndSite, LPMSG lpmsg)
{
    HRESULT         hr;
    CNotification   nf;

    _fFirstTimeTab = FALSE; //IsInIEBrowser(this);
    _fInPlaceActivating = TRUE;
    _fEnableInteraction = TRUE;

    // Do the normal transition, creating our main window
    hr = CServer::RunningToInPlace(hWndSite, lpmsg);
    if(hr)
    {
        goto Cleanup; // Do not goto error because CServer has already performed all necessary cleanup.
    }

    // Make sure that we have a current element before we set focus to one.
    // If we are parse done, and just becoming inplace, then we don't have an
    // active element yet.
    {
        DeferSetCurrency(0);
    }

    // Prepare the view
    _view.Activate();
    _view.SetViewSize(_dci._sizeDst);
    _view.SetFocus(_pElemCurrent, _lSubCurrent);

    // not having a DocHostShowUI is not a failing condition for this function
    hr = S_OK;

    // Make sure that everything is laid out correctly
    // *before* doing the BroadcastNotify because with olesites,
    // we want to baseline them upon the broadcast.  When doing this,
    // we better know the olesite's position.
    _view.EnsureView(LAYOUT_SYNCHRONOUS);

    nf.DocStateChange1(PrimaryRoot(), (void*)OS_RUNNING);
    BroadcastNotify(&nf);

    // Show document window *after* all child windows have been
    // created, to reduce clipping region recomputations.  To avoid
    // WM_ERASEBKGND sent outside of a WM_PAINT, we show without
    // redraw and then invalidate.
    SetWindowPos(_pInPlace->_hwnd, NULL, 0, 0, 0, 0,
        SWP_SHOWWINDOW|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOREDRAW|SWP_NOACTIVATE);

    // schedule flushing the AttachPeerQueue. Flushing should be posted, not synchronous,
    // so to allow shdocvw to connect to the doc fully before we start attaching peers
    // and possible firing events to shdocvw (e.g. errors)
    /*GWPostMethodCall(
        this,
        ONCALL_METHOD(CDoc, PeerDequeueTasks, peerdequeue),
        0, FALSE, "CDocument::PeerDequeueTasks"); wlw note*/

    Invalidate(NULL, NULL, NULL, INVAL_CHILDWINDOWS);

    DeferUpdateUI();

Cleanup:
    _fInPlaceActivating = FALSE;
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::InPlaceToRunning
//
//  Synopsis:   Effects the inplace-active to running state transition
//
//  Returns:    SUCCESS in all but catastrophic circumstances
//
//  Notes:      This method in-place deactivates all inside-out embeddings
//              in addition to normal CServer base processing.
//
//---------------------------------------------------------------
HRESULT CDocument::InPlaceToRunning(void)
{
    HRESULT         hr = S_OK;
    CNotification   nf;

    {
        CElement* pElemFireTarget;
        pElemFireTarget = _pElemCurrent->GetFocusBlurFireTarget(_lSubCurrent);
        Assert(pElemFireTarget);

        CElement::CLock LockUpdate(pElemFireTarget, CElement::ELEMENTLOCK_UPDATE);
        _fInhibitFocusFiring = TRUE;
        pElemFireTarget->RequestYieldCurrency(TRUE);

        if(pElemFireTarget->IsInMarkup())
        {
            hr = pElemFireTarget->YieldCurrency(_pElementDefault);
            if(hr)
            {
                goto Cleanup;
            }

            if(!TestLock(FORMLOCK_CURRENT))
            {
                Assert(pElemFireTarget==_pElemCurrent || pElemFireTarget->Tag()==ETAG_AREA);
                pElemFireTarget->Fire_onblur(0);
            }
        }
        _fInhibitFocusFiring = FALSE;
    }

    if(_pElemUIActive)
    {
        _pElemUIActive->YieldUI(PrimaryRoot());
    }
    _pElemCurrent = _pElemUIActive = PrimaryRoot();

    SetMouseCapture(NULL, NULL);
    releaseCapture();

    // stop scripts if not in design mode
    CleanupScriptTimers();

    if(_fFiredOnLoad)
    {
        _fFiredOnLoad = FALSE;
        /*if(_pOmWindow)
        {
            CDocument::CLock Lock(this);
            _pOmWindow->Fire_onunload();
        } wlw note*/
    }

    // Shutdown the view
    _view.Deactivate();

    hr = CServer::InPlaceToRunning();

    // Clear these variables so we don't stumble over stale
    // values should we be inplace activated again.
    FormsKillTimer(this, TIMER_ID_MOUSE_EXIT);
    _fMouseOverTimer = FALSE;
    if(_pNodeLastMouseOver)
    {
        CTreeNode* pNodeRelease = _pNodeLastMouseOver;
        _pNodeLastMouseOver = NULL;
        pNodeRelease->NodeRelease();
    }

    nf.DocStateChange1(PrimaryRoot(), (void*)OS_INPLACE);
    BroadcastNotify(&nf);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:    CDocument::AttachWin
//
//  Synopsis:  Create our InPlace window
//
//  Arguments: [hwndParent] -- our container's hwnd
//
//  Returns:   hwnd of InPlace window, or NULL
//
//---------------------------------------------------------------
HRESULT CDocument::AttachWin(HWND hwndParent, RECT* prc, HWND* phwnd)
{
    HRESULT     hr = S_OK;
    HWND        hwnd = NULL;
    const DWORD STYLE_ENABLED = WS_CHILD|WS_CLIPCHILDREN|WS_CLIPSIBLINGS;
    const DWORD STYLE_DISABLED = STYLE_ENABLED|WS_DISABLED;

    // Note: this code is duplicated in CServer::AttachWin.
    if(!s_atomWndClass)
    {
        hr = RegisterWindowClass(
            _T("Server"),
            CServer::WndProc,
            CS_DBLCLKS,
            NULL, NULL,
            &CServer::s_atomWndClass);
        if(hr)
        {
            goto Cleanup;
        }
    }

    Assert(phwnd);

    if(!hwnd)
    {
        TCHAR* pszBuf;
        TCHAR achClassName[256];

        pszBuf = (TCHAR*)(DWORD_PTR)CServer::s_atomWndClass;

        hwnd = CreateWindowEx(
            0,
            pszBuf,
            NULL,
            STYLE_ENABLED,
            prc->left,
            prc->top,
            prc->right-prc->left,
            prc->bottom-prc->top,
            hwndParent,
            0,              // no child identifier - shouldn't send WM_COMMAND
            _afxGlobalData._hInstThisModule,
            (CServer*)this);

        ::GetClassName(hwndParent, achClassName, ARRAYSIZE(achClassName));
    }

    if(!hwnd)
    {
        goto Win32Error;
    }

    GWPostMethodCall(this, ONCALL_METHOD(CDocument, EnableDragDrop, enabledragdrop), 0, FALSE, "CDocument::EnableDragDrop");

    // Initialize our _wUIState from the window.
    _wUIState = SendMessage(hwnd, WM_QUERYUISTATE, 0, 0);

Cleanup:
    *phwnd = hwnd;
    RRETURN(hr);

Win32Error:
    hr = GetLastWin32Error();
    DetachWin();
    goto Cleanup;
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::DetachWin
//
//  Synopsis:   Our window is going down. Cleanup our UI.
//
//--------------------------------------------------------------------
void CDocument::DetachWin()
{
    THREADSTATE* pts;

    pts = GetThreadState();

    if(_pInPlace->_hwnd)
    {
        LRESULT lResult;

        ShowWindow(_pInPlace->_hwnd, SW_HIDE);
        OnWindowMessage(WM_KILLFOCUS, 0, 0, &lResult);
        if(SetParent(_pInPlace->_hwnd, pts->gwnd.hwndGlobalWindow))
        {
            // BUGBUG: can killing of all timers be done better?
            KillTimer(_pInPlace->_hwnd, 1);                     // mouse move timer
            KillTimer(_pInPlace->_hwnd, TIMER_DEFERUPDATEUI);   // defer update UI timer
            _fUpdateUIPending = FALSE;                          // clear update UI pending flag
            _fNeedUpdateUI = FALSE;
            SetWindowLongPtr(_pInPlace->_hwnd, GWLP_USERDATA, 0);  // disconnect the window from this CDoc
            OnDestroy();    // this call is necessary to balance all ref-counting and init-/deinitialization;
                            // normally in CServer this happens upon WM_DESTROY message; however, as we reuse
                            // the window, we don't get WM_DESTROY in this codepath so we explicitly call OnDestroy
            PrivateRelease();
            _pInPlace->_hwnd = NULL;
        }
        else // SetParent failed
        {
            super::DetachWin();
        }
    }
}

//----------------------------------------------------------
//
//  Member   : CDocument::UnloadContents
//
//  Synopsis : Frees resources
//
//----------------------------------------------------------
void CDocument::UnloadContents(BOOL fPrecreated, BOOL fRestartLoad)
{
    // Don't allow WM_PAINT or WM_ERASEBKGND to get processed while
    // the tree is being deleted.  Some controls when deleting their
    // HWNDs will cause WM_ERASEBKGND to get sent to our window.  That
    // starts the paint cycle which is bad news when the site tree is
    // being destroyed.
    CLock Lock(this, SERVERLOCK_BLOCKPAINT|FORMLOCK_UNLOADING);

    // Indicate to anybody who checks that the document has been unloaded
    _cDie++;

    _cStylesheetDownloading         = 0;
    _dwStylesheetDownloadingCookie += 1;

    UnregisterUrlImgCtxCallbacks();

    /*if(_dwAsyncCookie && GetProgSink())
    {
        GetProgSink()->DelProgress(_dwAsyncCookie);
        _dwAsyncCookie = 0;
    } wlw note*/

    UpdateInterval(0);

    _aryFocusItems.DeleteAll();
    _aryAccessKeyItems.DeleteAll();

    /*CommitDeferredScripts(0); wlw note*/

    GWKillMethodCall(this, ONCALL_METHOD(CDocument, FirePostedOnPropertyChange, firepostedonpropertychange), 0);
    GWKillMethodCall(this, ONCALL_METHOD(CDocument, FaultInUSP, faultinusp), 0);
    GWKillMethodCall(this, ONCALL_METHOD(CDocument, FaultInJG, faultinjg), 0);

    FormsKillTimer(this, TIMER_ID_MOUSE_EXIT);
    _fMouseOverTimer = FALSE;

    CTreeNode::ClearPtr(&_pNodeLastMouseOver);
    CTreeNode::ClearPtr(&_pNodeGotButtonDown);

    releaseCapture();

    // nothing depends on the tree now; release the tree
    //
    // Detach all sites still not detached
    if(_pPrimaryMarkup)
    {
        if(_pInPlace)
        {
            _pInPlace->_fDeactivating = TRUE;
        }

        _view.Unload();

        _pElemCurrent = _pElementDefault;
        _pElemUIActive = NULL;

        _pPrimaryMarkup->UnloadContents();
    }

    Assert(_pElemCurrent == _pElementDefault);
    Assert(!_pElemUIActive);

    if(_pPrimaryMarkup)
    {
        _pElemCurrent = _pPrimaryMarkup->Root();
    }

    ReleaseNotify(); // NOTE (alexz) this has to happen before doing PeerDequeueTasks.

    // behaviors UnloadContents
    /*PeerDequeueTasks(TRUE);

    Assert(0 == _aryPeerQueue.Size()); wlw note*/
    Assert(0 == _aryElementReleaseNotify.Size());

    /*for(i=_aryPeerFactoriesUrl.Size(); i>0; i--)
    {
        _aryPeerFactoriesUrl[i-1]->Release();
    }
    _aryPeerFactoriesUrl.DeleteAll();

    // reset _fPeersPossible, unless it was set because host supplies peer factory. In that case after refresh
    // we won't be requerying again for any css, namespace, and other information provided by host so the bit
    // can't be turned back on
    if(!_pHostPeerFactory)
    {
        _fPeersPossible = FALSE;
    } wlw note*/

    // There might be some filter element tasks pending but we don't care

    // If a filter instantiate caused a navigate and unloaded the doc
    // we will be in a bit of trouble.  This doesn't happen today but just in case.
    Assert(!TestLock(FORMLOCK_FILTER));

    _fPendingFilterCallback = FALSE;
    GWKillMethodCall(this, ONCALL_METHOD(CDocument, FilterCallback, filtercallback), 0);
    _aryPendingFilterElements.DeleteAll();

    if(_pInPlace)
    {
        _pInPlace->_fDeactivating = FALSE;
    }

    if(_pPrimaryMarkup)
    {
        _pPrimaryMarkup->UpdateMarkupTreeVersion();
    }

    _lFocusTreeVersion = 0;

    _fFiredOnLoad = FALSE;
    _fRegionCollection = FALSE;
    _fExpando = EXPANDOS_DEFAULT;
    _fHasOleSite = FALSE;
    _fCurrencySet = FALSE;

    GWKillMethodCall(this, ONCALL_METHOD(CDocument, SendSetCursor, sendsetcursor), 0);

    /*ClearInterface(&_pTypInfo);
    ClearInterface(&_pTypInfoCoClass); wlw note*/

    // Only free expandos if not precreated from shdocvw (window.open).
    if(!fPrecreated)
    {
        _AtomTable.Free();
    }

    _cstrUrl.Set(_afxGlobalData._szModuleFolder);

    _codepage = _afxGlobalData._cpDefault;
    _codepageFamily = _afxGlobalData._cpDefault;

    _fVisualOrder = 0;

    // Kill all the timers
    CleanupScriptTimers();

    // Free cached radio groups
    while(_pRadioGrpName)
    {
        RADIOGRPNAME* pRadioGroup = _pRadioGrpName->_pNext;

        SysFreeString((BSTR)_pRadioGrpName->lpstrName);
        delete _pRadioGrpName;
        _pRadioGrpName = pRadioGroup;
    }

    // Cleanup window attrarray.  Don't just release the window object
    // here because there are other people outside who are holding
    // refs onto the real window object like shdocvw.  If we did release
    // they don't know that we've just tossed the window object away.
    /*if(_pOmWindow)
    {
        GWKillMethodCall(_pOmWindow, NULL, 0);

        if(!fPrecreated)
        {
            COmWindow2* pWindow = _pOmWindow->Window();

            Assert(pWindow);
            if(*(pWindow->GetAttrArray()))
            {
                (*(pWindow->GetAttrArray()))->FreeSpecial();
            }

            // clear up the window proxy also, since that is where the
            // scipt holder is
            delete *(_pOmWindow->GetAttrArray());
            _pOmWindow->SetAttrArray(NULL);
        }
    } wlw note*/

    // Don't delete our attr array outright since we've stored lotsa things
    // in there like prop notify sinks.  We're going to call FreeSpecial to
    // free everything else except these things.
    if(*GetAttrArray())
    {
        (*GetAttrArray())->FreeSpecial();
    }

    _bufferDepth = 0; // reset the buffer depth

    if(_pDwnDoc)
    {
        _pDwnDoc->Disconnect();
        _pDwnDoc->Release();
        _pDwnDoc = NULL;
    }

    _aryDefunctObjects.DeleteAll();

    if(_fHasOleSite)
    {
        CoFreeUnusedLibraries();
    }

    // NOTE(SujalP): Our current usage pattern dictates that at this point there
    // should be no used entries in the cache. However, when we start caching a
    // plsline inside the cache at that point we will have used entries here and
    // then VerifyNonUsed() cannot be called.
    TLS(_pLSCache)->Dispose(TRUE);

    NotifySelection(SELECT_NOTIFY_DOC_ENDED, NULL);

    Assert(_lRecursionLevel == 0);
}

void CDocument::InitDownloadInfo(DWNLOADINFO* pdli)
{
    memset(pdli, 0, sizeof(DWNLOADINFO));
    pdli->pInetSess = TlsGetInternetSession();
    pdli->pDwnDoc   = _pDwnDoc;
}

//+-------------------------------------------------------------------------
//
// Method:   CTQueryStatus
//
// Synopsis: Call IOleCommandTarget::QueryStatus on an object.
//
//--------------------------------------------------------------------------
HRESULT CTQueryStatus(IUnknown* pUnk, const GUID* pguidCmdGroup,
                      ULONG cCmds, MSOCMD rgCmds[], MSOCMDTEXT* pcmdtext)
{
    IOleCommandTarget* pCommandTarget;
    HRESULT hr;

    if(!pUnk)
    {
        Assert(0);
        RRETURN(E_FAIL);
    }

    hr = pUnk->QueryInterface(IID_IOleCommandTarget, (void**)&pCommandTarget);

    if(!hr)
    {
        hr = pCommandTarget->QueryStatus(
            pguidCmdGroup,
            cCmds,
            rgCmds,
            pcmdtext);
        pCommandTarget->Release();
    }

    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
// Method:   CTExec
//
// Synopsis: Call IOleCommandTarget::Exec on an object.
//
//--------------------------------------------------------------------------
HRESULT CTExec(IUnknown* pUnk, const GUID* pguidCmdGroup, DWORD nCmdID,
               DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut)
{
    IOleCommandTarget* pCommandTarget;
    HRESULT hr;

    if(!pUnk)
    {
        Assert(0);
        RRETURN(E_FAIL);
    }

    hr = pUnk->QueryInterface(IID_IOleCommandTarget, (void**)&pCommandTarget);

    if(!hr)
    {
        hr = pCommandTarget->Exec(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut);
        pCommandTarget->Release();
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
// Member:   CDocument::SetClipboard
//
// Synopsis: Set Security Domain before actual FormSetClipboard
//
//----------------------------------------------------------------------------
HRESULT CDocument::SetClipboard(IDataObject* pDO)
{
    RRETURN(FormSetClipboard(pDO));
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::HitTestPoint
//
//  Synopsis:   Find site at given position
//
//  Arguments   pt              The position.
//              ppSite          The site, can be null on return.
//              dwFlags         HT_xxx flags
//
//  Returns:    HTC
//
//----------------------------------------------------------------------------
HTC CDocument::HitTestPoint(CMessage* pMessage, CTreeNode** ppNodeElement, DWORD dwFlags)
{
    HTC htc;
    CTreeNode* pNodeElement;

    Assert(pMessage);

    // Ensure that pointers are set for simple code down the line.
    if(ppNodeElement == NULL)
    {
        ppNodeElement = &pNodeElement;
    }

    htc = _view.HitTestPoint(pMessage, ppNodeElement, dwFlags);

    return htc;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnControlInfoChanged, public
//
//  Synopsis:   To be called when our control info has changed.  This normally
//              happens when a site gets an OnControlInfoChanged call from
//              its control.  Calls OnControlInfoChanged on our site.
//
//  Arguments:  (none)
//
//----------------------------------------------------------------------------
void CDocument::OnControlInfoChanged(DWORD_PTR dwContext)
{
    _fOnControlInfoChangedPosted = FALSE;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::QueryService
//
//  Synopsis:   QueryService for the form.  Delegates to the form's
//              own site.
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CDocument::QueryService(REFGUID guidService, REFIID iid, void** ppv)
{
    HRESULT hr;

    // Certain services should never be bubbled up through 
    // the client site
    // SID_SContainerDispatch   -   Always provide container's IDispatch
    // SID_SLocalRegistry       -   Never bubble due to security concerns.
    //                              License manager is per document.
    // SID_SBindHost            -   We provide the service but might also
    //                              forward some calls to an outer bindhost,
    //                              if any.
    if(SID_SContainerDispatch==guidService || SID_SBindHost==guidService ||
        CLSID_HTMLDocument==guidService ||
        IID_IInternetHostSecurityManager==guidService || IID_IAccessible==guidService)
    {
        hr = CreateService(guidService, iid, ppv);
    }
    else
    {
        hr = CServer::QueryService(guidService, iid, ppv);
        if(hr)
        {
            hr = CreateService(guidService, iid, ppv);
        }
    }

    RRETURN_NOTRACE(hr);
}

HRESULT CDocument::CreateService(REFGUID guidService, REFIID iid, void** ppv)
{
    // wlw note
    return E_NOTIMPL;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::GetUniqueIdentifier
//
//  Synopsis:   Gets a unique ID for the control.  If possible, the form
//              coordinates with its container through the IGetUniqueID
//              service to get an ID unique within the container.
//
//  Arguments:  [pstr]      The string to set into
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CDocument::GetUniqueIdentifier(CString* pstr)
{
    TCHAR ach[64];
    HRESULT hr;

    memset(ach, 0, sizeof(ach));

    // Prefix with id_ because scriptlet code
    // doesn't currently like ID's that are all digits
    hr = Format(_afxGlobalData._hInstResource, 0, ach, ARRAYSIZE(ach), UNIQUE_NAME_PREFIX _T("<0d>"), (long)++_ID);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pstr->Set(ach);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------
//
//  Member:     CDocument::GetElementClient
//
//  Synopsis:   returns CFrameSetSite in document, if the doc contains one
//
//-----------------------------------------------------------------------
CElement* CDocument::GetPrimaryElementClient()
{
    return _pPrimaryMarkup?_pPrimaryMarkup->GetElementClient():NULL;
}

CElement* CDocument::GetPrimaryElementTop()
{
    return _pPrimaryMarkup?_pPrimaryMarkup->GetElementTop():NULL;
}

void CDocument::DocumentFromWindow(RECTL* prcl, const RECT* prc)
{
    DocumentFromWindow((POINTL*)&prcl->left, *(POINT*)&prc->left);
    DocumentFromWindow((POINTL*)&prcl->right, *(POINT*)&prc->right);
}

void CDocument::DocumentFromScreen(POINTL* pptl, POINT pt)
{
    *(POINT*)pptl = pt;
    ScreenToClient(GetHWND(), (POINT*)pptl);
    DocumentFromWindow(pptl, pptl->x, pptl->y);
}

void CDocument::ScreenFromWindow(POINT* ppt, POINT pt)
{
    *ppt = pt;
    ClientToScreen(GetHWND(), ppt);
}

void CDocument::HimetricFromDevice(RECTL* prcl, const RECT* prc)
{
    HimetricFromDevice((POINTL*)&prcl->left, *(POINT*)&prc->left);
    HimetricFromDevice((POINTL*)&prcl->right, *(POINT*)&prc->right);
}

void CDocument::DeviceFromHimetric(RECT* prc, const RECTL* prcl)
{
    DeviceFromHimetric((POINT*)&prc->left, *(POINTL*)&prcl->left);
    DeviceFromHimetric((POINT*)&prc->right, *(POINTL*)&prcl->right);
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::Invalidate
//
//  Synopsis:   Invalidate an area in the form.
//
//              We keep track of whether the background should be
//              erased on paint privately.  If we let Windows keep
//              track of this, then we can get flashing because
//              the WM_ERASEBKGND message can be delivered far in
//              advance of the WM_PAINT message.
//
//              Invalidation flags:
//
//              INVAL_CHILDWINDOWS
//                  Invaildate child windows. Causes the RDW_ALLCHILDREN
//                  flag to be passed to RedrawWindow.
//
//  Arguments:  prc         The physical rectangle to invalidate.
//              prcClip     Clip invalidation against this rectangle.
//              hrgn        ...or, the region to invalidate.
//              dwFlags     See description above.
//
//-------------------------------------------------------------------------
void CDocument::Invalidate(const RECT* prc, const RECT* prcClip, HRGN hrgn, DWORD dwFlags)
{
    UINT uFlags;
    RECT rc;

    if(prcClip)
    {
        Assert(prc);
        if(!IntersectRect(&rc, prc, prcClip))
        {
            return;
        }
        prc = &rc;
    }

    uFlags = RDW_INVALIDATE | RDW_NOERASE;

    if(dwFlags & INVAL_CHILDWINDOWS)
    {
        uFlags |= RDW_ALLCHILDREN;
    }

    Assert(_pInPlace && " about to crash");
    // CHROME
    _cInval++;

    //if(!_pUpdateIntSink)
    //{
        RedrawWindow(_pInPlace->_hwnd, prc, hrgn, uFlags);
    //}
    //else if(_pUpdateIntSink->_state < UPDATEINTERVAL_ALL)
    //{
    //    // accumulate invalid rgns until the updateInterval timer fires
    //    HRGN hrgnSrc;
    //    if(prc)
    //    {
    //        hrgnSrc = CreateRectRgnIndirect(prc);
    //        if(!hrgnSrc)
    //        {
    //            RedrawWindow(_pInPlace->_hwnd, prc, hrgn, uFlags);
    //        }
    //        else if(UPDATEINTERVAL_EMPTY == _pUpdateIntSink->_state)
    //        {
    //            _pUpdateIntSink->_hrgn = hrgnSrc;
    //            _pUpdateIntSink->_state = UPDATEINTERVAL_REGION;
    //            _pUpdateIntSink->_dwFlags |= uFlags;
    //        }
    //        else
    //        {
    //            Assert(UPDATEINTERVAL_REGION == _pUpdateIntSink->_state);
    //            if(ERROR == CombineRgn(_pUpdateIntSink->_hrgn,
    //                _pUpdateIntSink->_hrgn,
    //                hrgnSrc, RGN_OR))
    //            {
    //                RedrawWindow(_pInPlace->_hwnd, prc, hrgn, uFlags);
    //            }
    //            else
    //            {
    //                _pUpdateIntSink->_dwFlags |= uFlags;
    //            }
    //            DeleteObject(hrgnSrc);
    //        }
    //    }
    //    else if(hrgn)
    //    {
    //        if(UPDATEINTERVAL_EMPTY == _pUpdateIntSink->_state)
    //        {
    //            _pUpdateIntSink->_hrgn = hrgn;
    //            _pUpdateIntSink->_state = UPDATEINTERVAL_REGION;
    //            _pUpdateIntSink->_dwFlags |= uFlags;
    //        }
    //        else
    //        {
    //            Assert(UPDATEINTERVAL_REGION == _pUpdateIntSink->_state);
    //            if(ERROR == CombineRgn(_pUpdateIntSink->_hrgn,
    //                _pUpdateIntSink->_hrgn,
    //                hrgn, RGN_OR))
    //            {
    //                RedrawWindow(_pInPlace->_hwnd, prc, hrgn, uFlags);
    //            }
    //            else
    //            {
    //                _pUpdateIntSink->_dwFlags |= uFlags;
    //            }
    //        }
    //    }
    //    else
    //    {
    //        // update entire client area, no need to accumulate anymore
    //        _pUpdateIntSink->_state = UPDATEINTERVAL_ALL;
    //        DeleteObject(_pUpdateIntSink->_hrgn);
    //        _pUpdateIntSink->_hrgn = NULL;
    //        _pUpdateIntSink->_dwFlags |= uFlags;
    //    }
    //}
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::Invalidate
//
//  Synopsis:   Invalidate the form's entire area.
//
//-------------------------------------------------------------------------
void CDocument::Invalidate()
{
    Invalidate(NULL, NULL, NULL, 0);
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::UpdateForm
//
//  Synopsis:   Update the form's window
//
//-------------------------------------------------------------------------
void CDocument::UpdateForm()
{
    if(_pInPlace)
    {
        UpdateChildTree(_pInPlace->_hwnd);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::updateInterval
//
//  Synopsis:   Set the paint update interval. Throttles multiple controls
//              that are randomly invalidating into a periodic painting
//              interval.
//
//----------------------------------------------------------------------------
void CDocument::UpdateInterval(long interval)
{
//    ITimerService*  pTM = NULL;
//    VARIANT         vtimeMin, vtimeMax, vtimeInt;
//
//    interval = interval < 0 ? 0 : interval; // no negative values
//    VariantInit(&vtimeMin); V_VT(&vtimeMin) = VT_UI4;
//    VariantInit(&vtimeMax); V_VT(&vtimeMax) = VT_UI4;
//    VariantInit(&vtimeInt); V_VT(&vtimeInt) = VT_UI4;
//
//    if(!_pUpdateIntSink)
//    {
//        if(0 == interval)
//        {
//            return;
//        }
//
//        // allocate timer and sink
//        Assert(!_pUpdateIntSink);
//        _pUpdateIntSink = new CDocUpdateIntSink(this);
//        if(!_pUpdateIntSink)
//        {
//            goto error;
//        }
//
//        if(FAILED(QueryService( SID_STimerService, IID_ITimerService, (void**)&pTM)))
//        {
//            goto error;
//        }
//
//        if(FAILED(pTM->GetNamedTimer(NAMEDTIMER_DRAW, &_pUpdateIntSink->_pTimer)))
//        {
//            goto error;
//        }
//
//        pTM->Release();
//        pTM = NULL;
//    }
//
//    if(0 == interval)
//    {
//        // disabling updateInterval, invalidate what we have
//        HRGN hrgn = _pUpdateIntSink->_hrgn;
//        DWORD hrgnFlags = _pUpdateIntSink->_dwFlags;
//
//        _pUpdateIntSink->_pDoc = NULL; // let sink know not to respond.
//        _pUpdateIntSink->_pTimer->Unadvise(_pUpdateIntSink->_cookie);
//        _pUpdateIntSink->_pTimer->Release();
//        _pUpdateIntSink->Release();
//        _pUpdateIntSink = NULL;
//        Invalidate(NULL, NULL, hrgn, hrgnFlags);
//        DeleteObject(hrgn);
//    }
//    else if(interval != _pUpdateIntSink->_interval)
//    {
//        // reset timer interval
//        _pUpdateIntSink->_pTimer->GetTime(&vtimeMin);
//        V_UI4(&vtimeMax) = 0;
//        V_UI4(&vtimeInt) = interval;
//        _pUpdateIntSink->_pTimer->Unadvise(_pUpdateIntSink->_cookie); // ok if 0
//        if(FAILED(_pUpdateIntSink->_pTimer->Advise(vtimeMin, vtimeMax,
//            vtimeInt, 0, (ITimerSink*)_pUpdateIntSink, &_pUpdateIntSink->_cookie)))
//        {
//            goto error;
//        }
//
//        _pUpdateIntSink->_interval = interval;
//    }
//
//cleanup:
//    return;
//
//error:
//    if(_pUpdateIntSink)
//    {
//        ReleaseInterface(_pUpdateIntSink->_pTimer);
//        _pUpdateIntSink->Release();
//        _pUpdateIntSink = NULL;
//    }
//    ReleaseInterface(pTM);
//    goto cleanup;
}

LONG CDocument::GetUpdateInterval()
{
    return /*_pUpdateIntSink?_pUpdateIntSink->_interval:*/0;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::DeferUpdateUI, SetUpdateTimer
//
//  Synopsis:   Post a request to ourselves to update the UI.
//
//-------------------------------------------------------------------------
void CDocument::DeferUpdateUI()
{
    _fNeedUpdateUI = TRUE;

    SetUpdateTimer();
}

void CDocument::SetUpdateTimer()
{
    // If called before we're inplace or have a window, just return.
    if(!_pInPlace || !_pInPlace->_hwnd)
    {
        return;
    }

    if(!_fUpdateUIPending)
    {
        _fUpdateUIPending = TRUE;
        SetTimer(_pInPlace->_hwnd, TIMER_DEFERUPDATEUI, 100, NULL);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::OnUpdateUI
//
//  Synopsis:   Process UpdateUI message.
//
//-------------------------------------------------------------------------
void CDocument::OnUpdateUI()
{
    IOleCommandTarget* pCommandTarget = NULL;

    Trace0("CDocument::OnUpdateUI\n");

    Assert(InPlace());

    KillTimer(_pInPlace->_hwnd, TIMER_DEFERUPDATEUI);
    _fUpdateUIPending = FALSE;

    if(_fNeedUpdateUI)
    {
        // update container UI.
        if(pCommandTarget)
        {
            // update menu/toolbar
            pCommandTarget->Exec(NULL, OLECMDID_UPDATECOMMANDS, MSOCMDEXECOPT_DONTPROMPTUSER, NULL, NULL);
            pCommandTarget->Release();
        }

        _fNeedUpdateUI = FALSE;
    }
}

HRESULT CDocument::NewDwnCtx(UINT dt, LPCTSTR pchSrc, CElement* pel, CDwnCtx** ppDwnCtx, BOOL fSilent, DWORD dwProgsinkClass)
{
    DWNLOADINFO dli     = { 0 };
    BOOL        fLoad   = TRUE;
    HRESULT     hr;
    TCHAR*      pchExpUrl = new TCHAR[pdlUrlLen];

    if(pchExpUrl == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    *ppDwnCtx = NULL;

    if(pchSrc == NULL)
    {
        hr = S_OK;
        goto Cleanup;
    }

    InitDownloadInfo(&dli);

    hr = ExpandUrl(pchSrc, pdlUrlLen, pchExpUrl, pel);
    if(hr)
    {
        goto Cleanup;
    }
    dli.pchUrl = pchExpUrl;
    dli.dwProgClass = dwProgsinkClass;

    if(PrimaryMarkup())
    {
        // This flag tells the download mechanism not use the CDwnInfo cache until it has at
        // least verified that the modification time of the underlying bits is the same as the
        // cached version.  Normally we allow connection to existing images if they are on the
        // same page and have the same URL.  Once the page is finished loading and script makes
        // changes to SRC properties, we perform this extra check.
        dli.fResynchronize = TRUE;
    }

    hr = ::NewDwnCtx(dt, fLoad, &dli, ppDwnCtx);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    if(pchExpUrl != NULL)
    {
        delete pchExpUrl;
    }
    RRETURN(hr);
}

//+--------------------------------------------------------------
//
//  Member:     CDocument::SetDownloadNotify
//
//  Synopsis:   Sets the Download Notify callback object
//              to be used next time the document is loaded
//
//---------------------------------------------------------------
HRESULT CDocument::SetDownloadNotify(IUnknown* punk)
{
    IDownloadNotify* pDownloadNotify = NULL;
    HRESULT hr = S_OK;

    if(punk)
    {
        hr = punk->QueryInterface(IID_IDownloadNotify, (void**)&pDownloadNotify);
        if(hr)
        {
            goto Cleanup;
        }
    }

    ReleaseInterface(_pDownloadNotify);
    _pDownloadNotify = pDownloadNotify;

Cleanup:
    RRETURN(hr);
}

BOOL CDocument::IsOwnerOf(IHTMLElement* pIElement)
{
    HRESULT     hr;
    BOOL        result = FALSE;
    CElement*   pElement;

    hr = pIElement->QueryInterface(CLSID_CElement, (void**)&pElement);

    if(hr)
    {
        goto Cleanup;
    }

    result = this == pElement->Doc();

Cleanup:
    return result;
}

BOOL CDocument::IsOwnerOf(IMarkupPointer* pIPointer)
{
    HRESULT         hr;
    BOOL            result = FALSE;
    CMarkupPointer* pPointer;

    hr = pIPointer->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointer);
    if(hr)
    {
        goto Cleanup;
    }

    result = (this == pPointer->Doc());

Cleanup:        
    return result;
}

BOOL CDocument::IsOwnerOf(IMarkupContainer* pContainer)
{
    HRESULT     hr;
    BOOL        result = FALSE;
    CMarkup*    pMarkup;

    hr = pContainer->QueryInterface(CLSID_CMarkup, (void**)&pMarkup);
    if(hr)
    {
        goto Cleanup;
    }

    result = (this == pMarkup->Doc());

Cleanup:
    return result;
}

HRESULT CDocument::CreateMarkup(CMarkup** ppMarkup, CElement* pElementMaster/*=NULL*/)
{
    HRESULT         hr;
    CRootElement*   pRootElement;
    CMarkup*        pMarkup = NULL;

    Assert(ppMarkup);

    pRootElement = new CRootElement(this);

    if(!pRootElement)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pMarkup = new CMarkup(this, pElementMaster);

    if(!pMarkup)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = pRootElement->Init();

    if(hr)
    {
        goto Cleanup;
    }

    {
        CElement::CInit2Context context(NULL, pMarkup);

        hr = pRootElement->Init2(&context);
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = pMarkup->Init(pRootElement);

    if(hr)
    {
        goto Cleanup;
    }

    *ppMarkup = pMarkup;
    pMarkup = NULL;

Cleanup:
    if(pMarkup)
    {
        pMarkup->Release();
    }

    CElement::ReleasePtr(pRootElement);

    RRETURN(hr);
}

HRESULT CDocument::CreateMarkupWithElement( 
        CMarkup** ppMarkup, 
        CElement* pElement, 
        CElement* pElementMaster/*=NULL*/)
{
    HRESULT     hr = S_OK;
    CMarkup*    pMarkup = NULL;

    Assert(pElement && !pElement->IsInMarkup());

    hr = CreateMarkup(&pMarkup, pElementMaster);

    if(hr)
    {
        goto Cleanup;
    }

    // Insert the element into the empty tree
    {
        CTreePos*   ptpRootBegin = pMarkup->FirstTreePos();
        CTreePos*   ptpNew;
        CTreeNode*  pNodeNew;

        Assert(ptpRootBegin);

        // Assert that the only thing in this tree is two WCH_NODE characters
        Assert(pMarkup->Cch() == 2);

        pNodeNew = new CTreeNode(pMarkup->Root()->GetFirstBranch(), pElement);
        if(!pNodeNew)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        Assert(pNodeNew->GetEndPos()->IsUninit());
        ptpNew = pNodeNew->InitEndPos(TRUE);
        hr = pMarkup->Insert(ptpNew, ptpRootBegin, FALSE);
        if(hr)
        {
            // The node never made it into the tree
            // so delete it
            delete pNodeNew;

            goto Cleanup;
        }

        Assert(pNodeNew->GetBeginPos()->IsUninit());
        ptpNew = pNodeNew->InitBeginPos(TRUE);
        hr = pMarkup->Insert( ptpNew, ptpRootBegin, FALSE);
        if(hr)
        {
            goto Cleanup;
        }

        pNodeNew->PrivateEnterTree();

        pElement->SetMarkupPtr(pMarkup);
        pElement->__pNodeFirstBranch = pNodeNew;
        pElement->PrivateEnterTree();

        {
            CNotification nf;
            nf.ElementEntertree(pElement);
            pElement->Notify(&nf);
        }

        // Insert the WCH_NODE characters for the element
        // The 2 is hardcoded since we know that there are
        // only 2 WCH_NODE characters for the root
        Verify(ULONG(CTxtPtr(pMarkup, 2).InsertRepeatingChar(2, WCH_NODE)) == 2);

        pMarkup->SetLoaded(TRUE);
    }

    if(ppMarkup)
    {
        *ppMarkup = pMarkup;
        pMarkup->AddRef();
    }

Cleanup:
    if(pMarkup)
    {
        pMarkup->Release();
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
// OnOptionSettingsChange()
//
//----------------------------------------------------------------------------
HRESULT CDocument::OnSettingsChange(BOOL fForce/*=FALSE*/)
{
    BOOL fNeedLayout = FALSE;

    {
        UpdateFromRegistry(REGUPDATE_REFRESH, &fNeedLayout);

        if(_pOptionSettings)
        {
            THREADSTATE* pts = GetThreadState();

            if(fNeedLayout || fForce)
            {
                ClearDefaultCharFormat();
                ForceRelayout();
            }
        }
    }
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::EnsureFormatCacheChange
//
//  Synopsis:   Clears the format caches for the client site and then
//              invalidates it.
//
//----------------------------------------------------------------------------
void CDocument::EnsureFormatCacheChange(DWORD dwFlags)
{
    if(!PrimaryRoot())
    {
        return;
    }

    PrimaryRoot()->EnsureFormatCacheChange(dwFlags);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ForceRelayout
//
//  Synopsis:   Whack the entire display
//
//----------------------------------------------------------------------------
HRESULT CDocument::ForceRelayout()
{
    HRESULT hr = S_OK;
    CNotification nf;

    if(_pPrimaryMarkup)
    {
        PrimaryRoot()->EnsureFormatCacheChange(ELEMCHNG_CLEARCACHES);
        nf.DisplayChange(PrimaryRoot());
        PrimaryRoot()->SendNotification(&nf);     
    }

    _view.ForceRelayout();

    RRETURN(hr);
}

typedef int WINAPI GETRANDOMRGN(HDC, HRGN, int);
static GETRANDOMRGN* s_pfnGetRandomRgn;
static BOOL s_fGetRandomRgnFetched = FALSE;

void CDocument::OnPaint()
{
    PAINTSTRUCT         ps;
    CFormDrawInfo       DI;
    BOOL                fViewIsReady;
    BOOL                fHtPalette;
    HRGN                hrgn = NULL;
    POINT               ptBefore;
    POINT               ptAfter;

    //_fInvalInScript = FALSE;

    DI._hdc = NULL;

    // Don't allow OnPaint to recurse.  This can occur as a result of the
    // call to RedrawWindow in a random control.

    // BUGBUG: StylesheetDownload should be linked to the LOADSTATUS interactive?
    if(TestLock(SERVERLOCK_BLOCKPAINT) || IsStylesheetDownload())
    {
        // We get endless paint messages when we try to popup a messagebox...
        // and prevented the messagebox from getting paint !
        BeginPaint(_pInPlace->_hwnd, &ps);
        EndPaint(_pInPlace->_hwnd, &ps);

        // Post a delayed paint to make up for what was missed
        // Post a delayed paint to make up for what was missed
        _view.Invalidate(&ps.rcPaint, FALSE, FALSE, FALSE);

        return;
    }

    fViewIsReady = _view.EnsureView(LAYOUT_SYNCHRONOUS|LAYOUT_INPAINT|LAYOUT_DEFERPAINT);
    Assert(!fViewIsReady || !_view.IsDisplayTreeOpen());

    CLock Lock(this, SERVERLOCK_BLOCKPAINT);


    // since we're painting, paint any accumulated regions if any
    //if(_pUpdateIntSink && _pUpdateIntSink->_state!=UPDATEINTERVAL_EMPTY)
    //{
    //    ::InvalidateRgn(_pInPlace->_hwnd, _pUpdateIntSink->_hrgn, FALSE);
    //    if(_pUpdateIntSink->_hrgn)
    //    {
    //        DeleteObject(_pUpdateIntSink->_hrgn);
    //        _pUpdateIntSink->_hrgn = NULL;
    //        _pUpdateIntSink->_state = UPDATEINTERVAL_EMPTY;
    //        _pUpdateIntSink->_dwFlags = 0;
    //    }
    //}

    ptBefore.x = ptBefore.y = 0;
    MapWindowPoints(_pInPlace->_hwnd, NULL, &ptBefore, 1);

    // Setup DC for painting.
    BeginPaint(_pInPlace->_hwnd, &ps);

    if(IsRectEmpty(&ps.rcPaint))
    {
        // It appears there are cases when our window is obscured yet
        // internal invalidations still trigger WM_PAINT with a valid
        // update region but the PAINTSTRUCT RECT is empty.
        goto Cleanup;
    }

    // If the view could not be properly prepared, accumulate the invalid rectangle and return
    // (The view will issue the invalidation/paint once it is safe to do so)
    if(!fViewIsReady)
    {
        _view.Invalidate(&ps.rcPaint);
        _view.SetFlag(CView::VF_FORCEPAINT);

        goto Cleanup;
    }

    if(!s_fGetRandomRgnFetched)
    {
        s_pfnGetRandomRgn = (GETRANDOMRGN*)GetProcAddress(GetModuleHandleA("GDI32.DLL"), "GetRandomRgn");
        s_fGetRandomRgnFetched = TRUE;
    }

    if(s_pfnGetRandomRgn)
    {
        Verify((hrgn = CreateRectRgnIndirect(&_afxGlobalData._Zero.rc)) != NULL);
        Verify(s_pfnGetRandomRgn(ps.hdc, hrgn, 4) != ERROR);
        Verify(OffsetRgn(hrgn, -ptBefore.x, -ptBefore.y) != ERROR);

        // Don't trust the region if the window moved in the meantime
        ptAfter.x = ptAfter.y = 0;
        MapWindowPoints(_pInPlace->_hwnd, NULL, &ptAfter, 1);

        if(ptBefore.x!=ptAfter.x || ptBefore.y!=ptAfter.y)
        {
            Verify(DeleteObject(hrgn));
            hrgn = NULL;
            goto Cleanup;
        }
    }

    GetPalette(ps.hdc, &fHtPalette);

    if(_pTimerDraw)
    {
        _pTimerDraw->Freeze(TRUE); // time stops so controls can synchronize
    }
    if(!TiledPaintDisabled())
    {
        // Invalidation was not on behalf of an ActiveX control.  Paint the
        // document in one pass here and tile if required in CSite::DrawOffscreen.
        DI.Init(GetPrimaryElementTop(), ps.hdc);
        DI._rcClip = ps.rcPaint;
        DI._rcClipSet = ps.rcPaint;
        DI._hrgnPaint = hrgn;
        DI._fHtPalette = fHtPalette;

        Assert(!_view.IsDisplayTreeOpen());

        _view.RenderView(&DI, DI._hrgnPaint);
    }
    else
    {
        RECT*   prc;
        int     c;
        struct REGION_DATA
        {
            RGNDATAHEADER rdh;
            RECT arc[MAX_INVAL_RECTS];
        } rd;

        // Invalidation was on behalf of an ActiveX Control.  We chunk things
        // up here based on the inval region with the hope that the ActiveX
        // Control will be painted in a single pass.  We do this because
        // some ActiveX Controls (controls using Direct Animation are an example)
        // have very bad performance when painted in tiles.

        // if we have more than one invalid rectangle, see if we can combine
        // them so drawing will be more efficient. Windows chops up invalid
        // regions in a funny, but predicable, way to maintain their ordered
        // listing of rectangles. Also, some times it is more efficient to
        // paint a little extra area than to traverse the hierarchy multiple times.
        if(hrgn && GetRegionData(hrgn, sizeof(rd), (RGNDATA*)&rd) &&
            rd.rdh.iType==RDH_RECTANGLES && rd.rdh.nCount<=MAX_INVAL_RECTS)
        {
            c = rd.rdh.nCount;
            prc = rd.arc;

            CombineRectsAggressive(&c, prc);
        }
        else
        {
            c = 1;
            prc = &ps.rcPaint;
        }

        // Paint each rectangle.
        for(; --c>=0; prc++)
        {
            DI.Init(GetPrimaryElementTop());
            DI._hdc = ps.hdc;
            DI._hic = ps.hdc;
            DI._rcClip = *prc;
            DI._rcClipSet = *prc;
            DI._fHtPalette = fHtPalette;

            if(prc != &ps.rcPaint)
            {
                // If painting the update region in more than one
                // pass and painting directly to the screen, then
                // we explicitly set the clip rectangle to insure correct
                // painting.  If we don't do this, the FillRect for the
                // background of a later pass will clobber the foreground
                // for an earlier pass.
                Assert(DI._hdc == ps.hdc);
                IntersectClipRect(DI._hdc, DI._rcClip.left, DI._rcClip.top,
                    DI._rcClip.right, DI._rcClip.bottom);
            }

            _view.RenderView(&DI, &DI._rcClip);

            if(c != 0)
            {
                // Restore the clip region set above.
                SelectClipRgn(DI._hdc, NULL);
            }
        }
    }

    if(_pTimerDraw)
    {
        _pTimerDraw->Freeze(FALSE);
    }

Cleanup:
    if(DI._hdc)
    {
        SelectPalette(DI._hdc, (HPALETTE)GetStockObject(DEFAULT_PALETTE), TRUE);
    }
    EndPaint(_pInPlace->_hwnd, &ps);
    _fDisableTiledPaint = FALSE;

    // Find out if the window moved during paint
    ptAfter.x = ptAfter.y = 0;
    MapWindowPoints(_pInPlace->_hwnd, NULL, &ptAfter, 1);

    if(ptBefore.x!=ptAfter.x || ptBefore.y!=ptAfter.y)
    {
        Invalidate(hrgn?NULL:&ps.rcPaint, NULL, hrgn, 0);
    }

    if(hrgn)
    {
        Verify(DeleteObject(hrgn));
    }

    Trace0("View : CDocument::OnPaint - Exit\n");
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnEraseBkgnd
//
//  Synopsis:   Handle WM_ERASEBKGND
//
//----------------------------------------------------------------------------
BOOL CDocument::OnEraseBkgnd(HDC hdc)
{
    CFormDrawInfo   DI;
    BOOL            fHtPalette;

    if(TestLock(SERVERLOCK_BLOCKPAINT))
    {
        return FALSE;
    }

    // Ignore unnecessary erase requests
    // (e.g., those generated by SetMenu, EndDeferWindowPos)
    if(TestLock(SERVERLOCK_IGNOREERASEBKGND) && WindowFromDC(hdc)==InPlace()->_hwnd)
    {
        return TRUE;
    }

    CLock Lock(this, SERVERLOCK_BLOCKPAINT);

    GetPalette(hdc, &fHtPalette);
    DI.Init(GetPrimaryElementTop());
    GetClipBox(hdc, &DI._rcClip);
    DI._rcClipSet = DI._rcClip;
    DI._hdc = hdc;
    DI._hic = hdc;
    DI._fInplacePaint = TRUE;
    DI._fHtPalette = fHtPalette;

    _view.EraseBackground(&DI, &DI._rcClip);

    return TRUE;
}

void CDocument::OnMenuSelect(UINT uItem, UINT fuFlags, HMENU hmenu)
{
}

void CDocument::OnCommand(int idm, HWND hwndCtl, UINT codeNotify)
{
    Exec((GUID*)&CGID_MSHTML, idm, 0, NULL, NULL);
}

HRESULT CDocument::OnHelp(HELPINFO*)
{
    return E_NOTIMPL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::PumpMessage
//
//  Synopsis:   Fires the event and then hands over the message to pTarget. If
//              ptarget is NULL, the message is handed over to the capture
//              object.
//
//----------------------------------------------------------------------------
HRESULT CDocument::PumpMessage(CMessage* pMessage, CTreeNode* pNodeTarget, BOOL fPerformTA)
{
    HRESULT     hr              = S_OK;
    BOOL        fCallCapture    = !pNodeTarget;
    BOOL        fInRecursion    = _fInPumpMessage;
    BOOL        fSaveModalState = _fModalDialogInScript; // to handle recursion
    BOOL        fEventCancelled = FALSE;
    CTreeNode*  pNodeHit        = pMessage->pNodeHit;
    CTreeNode*  pCurNode        = NULL;
    CTreeNode*  pNodeFireEvent  = NULL;
    CTreeNode*  pNodeFireEventSrcElement = NULL;
    CElement*   pElemCurrentOld = _pElemCurrent;
    ULONG       cDie            = _cDie;
    BOOL        fMouseMessage   = ((pMessage->message>=WM_MOUSEFIRST)
        && (pMessage->message<=WM_MOUSELAST))
        || (pMessage->message==WM_SETCURSOR)
        || (pMessage->message==WM_CAPTURECHANGED)
        || (pMessage->message==WM_MOUSELEAVE)
        || (pMessage->message==WM_MOUSEOVER);
    BOOL        fSiteSelected   = FALSE;

    // BUGBUG(laszlog): This should be changed into a real fix!
    if(!_pElemCurrent || _pElemCurrent->Tag()==ETAG_DEFAULT)
    {
        return S_OK;
    }

    Assert(pMessage);
    AssertSz((pNodeTarget && !pNodeTarget->IsDead() && pNodeTarget->Element() && pNodeTarget->IsInMarkup()) ||
        (!_fElementCapture || ((CElement*)_pvCaptureObject)->IsInMarkup()) ||
        (pMessage->message==WM_MOUSEOVER || pMessage->message==WM_MOUSELEAVE),
        "Trying to send a windows message to an element not in the tree");

    _fInPumpMessage = TRUE;

    // In browse mode, if the mouse is clicked on a child of BUTTON,
    // we want to send that click to the BUTTON instead. IE5 #34796, #26572
    switch(pMessage->message)
    {
    case WM_LBUTTONDOWN:
    case WM_MBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_MBUTTONUP:
    case WM_RBUTTONUP:
        if(pNodeHit && !pNodeHit->Element()->IsEditable(TRUE))
        {
            CTreeNode* pNodeLoop = pNodeHit;

            while(pNodeLoop && pNodeLoop->Tag()!=ETAG_BUTTON)
            {
                pNodeLoop = pNodeLoop->Parent();
            }
            if(pNodeLoop)
            {
                if(pNodeLoop != pNodeHit)
                {
                    if(pNodeTarget == pNodeHit)
                    {
                        pNodeTarget = pNodeLoop;
                    }

                    // pNodeHit is parented by a BUTTON.
                    pNodeHit = pNodeLoop;
                    pMessage->SetNodeHit(pNodeHit);
                }
            }
        }
    }

    switch(pMessage->message)
    {
    case WM_LBUTTONDOWN:
    case WM_MBUTTONDOWN:
    case WM_RBUTTONDOWN:
        Assert(pNodeHit || fCallCapture);
        CTreeNode::ReplacePtr(&_pNodeGotButtonDown, pNodeHit);
        _fFirstTimeTab = FALSE;
        // fall-through
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if(fPerformTA)
        {
            switch(pMessage->wParam)
            {
            case VK_MENU:
                if(_wUIState & UISF_HIDEACCEL)
                {
                    SendMessage(_pInPlace->_hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEACCEL|UISF_HIDEFOCUS), 0);
                }
                break;
            case VK_RIGHT:
            case VK_LEFT:
            case VK_UP:
            case VK_DOWN:
            case VK_TAB:
                if(_wUIState & UISF_HIDEFOCUS)
                {
                    SendMessage(_pInPlace->_hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEFOCUS), 0);
                }
                break;                
            }
        }
        break;
    }

    // Determine who should fire the event for this message
    if(!pNodeHit)
    {
        pNodeFireEvent = _pElemCurrent ? _pElemCurrent->GetFirstBranch() : NULL;
    }
    else
    {
        pNodeFireEvent = pNodeHit;
    }
    Assert(pNodeFireEvent);

    _fModalDialogInScript = FALSE;

    if(_pElementOMCapture)
    {
        if(!_pElementOMCapture->GetFirstBranch())
        {
            releaseCapture();
        }
        else if(fMouseMessage
            && (_fContainerCapture
            || !pNodeFireEvent->SearchBranchToRootForScope(_pElementOMCapture)))
        {
            pNodeFireEventSrcElement = pNodeFireEvent;
            TransformSlaveToMaster(&pNodeFireEventSrcElement);
            pNodeFireEvent = _pElementOMCapture->GetFirstBranch();
        }
    }

    if(!fPerformTA)
    {
        // We have already fired the event for certain messages in the TA pass
        switch(pMessage->message)
        {
        case WM_KEYDOWN:
        case WM_KEYUP:
        case WM_SYSKEYDOWN:
        case WM_SYSKEYUP:
            pMessage->fEventsFired = TRUE;
            break;
        }
    }

    // Fire the event unless the element is an OleSite or the rootsite
    // Also fire the event if in design mode and pMessage->htc is a grab handle or border
    if(!pMessage->fEventsFired
        && (
        !(pNodeFireEvent->Element()->TestClassFlag(CElement::ELEMENTDESC_OLESITE)
        || pNodeFireEvent->Tag()==ETAG_ROOT)))
    {
        CTreeNode::CLock lock(pNodeFireEvent);

        fEventCancelled = (S_FALSE != pNodeFireEvent->Element()->
            FireStdEventOnMessage(pNodeFireEvent, pMessage, NULL, pNodeFireEventSrcElement));

        if(!AreWeSaneAfterEventFiring(pMessage, cDie))
        {
            goto Cleanup;
        }

        // script may have removed pNodeFireEvent from the tree
        if(pNodeFireEvent->IsDead())
        {
            pNodeFireEvent = pNodeFireEvent->Element()->GetFirstBranch();
        }

        if(!pNodeFireEvent || !pNodeFireEvent->Element()->IsInMarkup())
        {
            goto Cleanup;
        }
    }

    // Do special stuff for MouseDown/Up messages
    if(!_pElementOMCapture)
    {
        switch(pMessage->message)
        {
        case WM_LBUTTONDOWN:
        case WM_MBUTTONDOWN:
        case WM_RBUTTONDOWN:
            {
                if(!fCallCapture)
                {
                    // Give currency to the site under the mouse, unless the event
                    // handler changed currency or the mouse is on a selection.
                    // Leave OLEsites alone; let them set their own currency in HandleMessage (#60183)
                    Assert(pNodeHit);

                    if(_fFirstTimeTab)
                    {
                        CLayout* pLayoutClient = GetPrimaryElementClient()->GetUpdatedLayout();
                        RECT rc;

                        pLayoutClient->GetRect(&rc);
                        if(PtInRect(&rc, pMessage->pt))
                        {
                            _fFirstTimeTab = FALSE;
                        }
                    }

                    // marka - in design mode, we change the currency behavior to make the parent
                    // of a 'site selectable' object active.
                    //
                    // BUGBUG - this has to be changed in the presence of element editability
                    // change the _fDesignMode test below when we have a "cheap" 
                    // way of knowing if your parent is editable
                    pCurNode = pMessage->pNodeHit;

                    if(pCurNode && !fSiteSelected && _pElemCurrent==pElemCurrentOld
                        && (DifferentScope(_pElemCurrent, pCurNode)
                        || _lSubCurrent!=pMessage->lSubDivision || !_pInPlace->_fFocus)
                        && !pCurNode->Element()->TestClassFlag(CElement::ELEMENTDESC_OLESITE))
                    {
                        // Only if we have hit client areas then change the currency
                        if(pMessage->htc<HTC_GRPTOPBORDER
                            && (pCurNode->Element()->Tag()!=ETAG_BODY
                            || (pMessage->htc!=HTC_HSCROLLBAR
                            &&  pMessage->htc!=HTC_VSCROLLBAR)))
                        {
                            BOOL fYieldFailed = FALSE;
                            CElement* pElemCurrentNew = pCurNode->Element();

                            if(S_OK!=pElemCurrentNew->BubbleBecomeCurrent(
                                pMessage->lSubDivision, &fYieldFailed, pMessage, TRUE))
                            {
                                if(fYieldFailed)
                                {
                                    fEventCancelled = TRUE;
                                }
                            }
                        }
                    }

                    // if a modal dialog was brought by the event handler, treat it
                    // as a cancel
                    fEventCancelled = (fEventCancelled || _fModalDialogInScript);
                }
            }
            break;

        case WM_LBUTTONUP:
        case WM_MBUTTONUP:
        case WM_RBUTTONUP:
            // If a MouseUp is cancelled, release capture. Any further cleanup
            // required should be performed by the site/element by handling
            // WM_CAPTURECHANGED
            if(fEventCancelled)
            {
                SetMouseCapture(NULL, NULL, FALSE);
            }
            break;
        }
    }

    if(_pElementOMCapture && !_pElementOMCapture->GetFirstBranch())
    {
        releaseCapture();
    }

    if(!fEventCancelled)
    {
        hr = S_FALSE;

        if(fPerformTA)
        {
            // don't bubbel through element handlemessage
            hr = PerformTA(pMessage);
        }
        else
        {
            // if we lost capture in the event handler or while setting focus,
            // send the message to the element that fired the event
            if(fCallCapture && !_pvCaptureObject)
            {
                fCallCapture = FALSE;
                pNodeTarget = pNodeFireEvent;
                Assert(pNodeTarget->IsInMarkup());
            }

            if(fCallCapture && fMouseMessage)
            {
                hr = CALL_METHOD((CVoid*)_pvCaptureObject, _pfnCapture, (pMessage));
            }
            else if(_pElementOMCapture && fMouseMessage && (_fContainerCapture
                || !pNodeTarget->SearchBranchToRootForScope(_pElementOMCapture)))
            {
                hr = _pElementOMCapture->HandleMessage(pMessage);
            }
            else
            {
                // give first chance to the Editor in certain cases
                if((pMessage->message>=WM_MOUSEFIRST
                    && pMessage->message<=WM_MOUSELAST
                    && pMessage->htc!=HTC_VSCROLLBAR
                    && pMessage->htc!=HTC_HSCROLLBAR
                    && IsPointInSelection(pMessage->pt))
                    || 
                    (((pMessage->message>=WM_KEYFIRST
                    && pMessage->message<=WM_KEYLAST))
                    && _pElemCurrent->IsEditable())
                    ) 
                {
                    hr = HandleSelectionMessage(pMessage, FALSE);
                }

                if(hr == S_FALSE)
                {
                    // BUGBUG (MohanB) We need to fix something here for IE6. Rigth now,
                    // if we click the mouse inside an <INPUT type=text> inside an <A>,
                    // we navigate instead of setting the caret. This is because we call
                    // HandleSelectionMessage() after we finish bubbling the message.
                    // Need to figure out how to fix this. Call HandleSelectionMessage()
                    // for every element? Evry container? Ugh.
                    while(pNodeTarget && pNodeTarget->IsInMarkup() && pNodeTarget->Tag()!=ETAG_ROOT)
                    {
                        Assert(pNodeTarget && pNodeTarget->IsInMarkup());
                        // BUGBUG (jbeda) I don't feel confident enough in this
                        // code to let the assert above handle this problem.  Break
                        // out here if we aren't in the markup because otherwise
                        // we *will* crash
                        if(!pNodeTarget || !pNodeTarget->IsInMarkup())
                        {
                            break;
                        }

                        hr = pNodeTarget->Element()->HandleMessage(pMessage);

                        if(hr != S_FALSE)
                        {
                            break;
                        }

                        // BUGBUG (MohanB) We may not want to bubble all messages.
                        // For example, WM_L(R)BUTTONDOWN, 
                        pNodeTarget = pNodeTarget->Parent();
                        if(pNodeTarget && pNodeTarget->Tag()==ETAG_ROOT)
                        {
                            CElement* pElemMaster = pNodeTarget->Element()->MarkupMaster();

                            if(pElemMaster)
                            {
                                pNodeTarget = pElemMaster->GetFirstBranch();
                            }
                        }
                    }
                }
                // give a second and final chance to the Editor
                if(hr == S_FALSE)
                {
                    hr = HandleSelectionMessage(pMessage, FALSE);
                }
            }

            if(hr == S_FALSE)
            {
                LRESULT lr;

                hr = OnDefWindowMessage(pMessage->message, pMessage->wParam, pMessage->lParam, &lr);
            }

        }

        if(!SUCCEEDED(hr))
        {
            goto Cleanup;
        }

        if(S_OK==hr && !fPerformTA)
        {
            CElement*   pelFireTarget = NULL;
            CTreeNode*  pNode = NULL;
            BOOL        fFireDblClick = FALSE;

            // Do click or dblclick as appropriate
            if(_fGotDblClk && pMessage->message==WM_LBUTTONUP)
            {
                fFireDblClick = TRUE;
                if(_pElementOMCapture && (_fContainerCapture || !pNodeHit
                    || !pNodeHit->SearchBranchToRootForScope(_pElementOMCapture)))
                {
                    pelFireTarget = _pElementOMCapture;
                }
                else if(pNodeHit)
                {
                    pelFireTarget = pNodeHit->Element();
                    pNode = pNodeHit;
                }
            }
            else if(pMessage->pNodeClk)
            {
                if(_pElementOMCapture && (_fContainerCapture  
                    || !pMessage->pNodeClk->SearchBranchToRootForScope(_pElementOMCapture)))
                {
                    pelFireTarget = _pElementOMCapture;
                }
                else
                {
                    pelFireTarget = pMessage->pNodeClk->Element();
                    pNode = pMessage->pNodeClk;
                }
            }
            else if(_pElementOMCapture && pMessage->message==WM_LBUTTONUP)
            {
                if(pNodeHit && !_fContainerCapture)
                {
                    pelFireTarget = pNodeHit->Element();
                }
                else
                {
                    pelFireTarget = _pElementOMCapture;
                }
            }

            if(pelFireTarget && !pelFireTarget->TestClassFlag(CElement::ELEMENTDESC_OLESITE))
            {
                if(fFireDblClick)
                {
                    hr = pelFireTarget->Fire_ondblclick(pNode, pMessage?pMessage->lSubDivision:0);
                    _fGotDblClk = FALSE;
                }
                else
                {
                    hr = pelFireTarget->DoClick(pMessage, pNode);
                }
            }
        }
    }

    if(!AreWeSaneAfterEventFiring(pMessage, cDie))
    {
        goto Cleanup;
    }

    if(_fInvalInScript && !fInRecursion)
    {
        Assert(_pInPlace);
        Assert(_pInPlace->_hwnd);
        UpdateWindow(_pInPlace->_hwnd);
    }

Cleanup:
    switch(pMessage->message)
    {
    case WM_LBUTTONUP:
    case WM_MBUTTONUP:
    case WM_RBUTTONUP:
    case WM_KILLFOCUS:
        CTreeNode::ReplacePtr(&_pNodeGotButtonDown, NULL);
        break;
    }
    _fModalDialogInScript = fSaveModalState;
    _fInPumpMessage = fInRecursion;

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
// Helper Function: IsValidAccessKey
//
//----------------------------------------------------------------------------
BOOL IsValidAccessKey(CDocument* pDoc, CMessage* pmsg)
{
    BOOL fResult = pmsg->message == WM_SYSKEYDOWN;
    if(fResult)
    {
        fResult = (pmsg->wParam!=VK_MENU) && (pDoc->_aryAccessKeyItems.Size()>0);
    }
    return fResult;
}

//+---------------------------------------------------------------------------
//
//  Member:     PerformTA
//
//  Synopsis:   Handle any accelerators
//
//  Arguments:  [pMessage]  -- message
//
//  Returns:    Returns S_OK if keystroke processed, S_FALSE if not.
//----------------------------------------------------------------------------
HRESULT CDocument::PerformTA(CMessage* pMessage)
{
    HRESULT hr = S_FALSE;

    // WinUser.h better not change! We are going to assume the order of the
    // navigation keys (Left/Right/Up/Down/Home/End/PageUp/PageDn), so let's
    // assert about it.
    Assert(VK_PRIOR + 1 == VK_NEXT);
    Assert(VK_NEXT  + 1 == VK_END);
    Assert(VK_END   + 1 == VK_HOME);
    Assert(VK_HOME  + 1 == VK_LEFT);
    Assert(VK_LEFT  + 1 == VK_UP);
    Assert(VK_UP    + 1 == VK_RIGHT);
    Assert(VK_RIGHT + 1 == VK_DOWN);

    if(WM_KEYDOWN==pMessage->message || WM_SYSKEYDOWN==pMessage->message)
    {
        // Handle accelerator here.
        //  1. Pass any accelerators that the editor requires
        //  2. Bubble up the element chain starting from _pElemCurrent (or capture elem)
        //     and call PerformTA on them
        //  3. Perform key navigation for TAB and accesskey

        //  Editor requires VK_TAB in <PRE> (61302), VK_BACK(58719, 58774),
        //  and the navigation keys 
        if(pMessage->message==WM_KEYDOWN
            && (pMessage->wParam==VK_BACK
            || (pMessage->wParam>=VK_PRIOR && pMessage->wParam<=VK_DOWN)
            || (pMessage->wParam==VK_TAB && FCaretInPre())))
        {
            hr = HandleSelectionMessage(pMessage, FALSE);
            if(hr != S_FALSE)
            {
                goto Cleanup;
            }
        }

        // BGBUG (MohanB) This will be wrong if an element can take keyboard capture without being current
        CElement*   pElemTarget         = _pElemCurrent;
        CTreeNode*  pNodeTarget         = pElemTarget->GetFirstBranch();
        BOOL        fGotEnterKey        = (pMessage->message==WM_KEYDOWN && pMessage->wParam==VK_RETURN);
        BOOL        fTranslateEnterKey  = FALSE;

        while(pNodeTarget && pElemTarget && pElemTarget->Tag()!=ETAG_ROOT)
        {
            hr = pElemTarget->PerformTA(pMessage);
            if(hr != S_FALSE)
            {
                goto Cleanup;
            }

            // Navigation keys are dealt with in HandleMessage, but we need to
            // treat them as accelerators (because many hosts such as HomePublisher
            // and KatieSoft Scroll expect us to - IE5 66735, 63774).
            if(WM_KEYDOWN==pMessage->message && VK_PRIOR<=pMessage->wParam && VK_DOWN>=pMessage->wParam)
            {
                // On the other hand, VID6.0 wants the fist shot at some of them if
                // there is a site-selection (they should fix this in VID6.1)
                if(_fVID && pMessage->wParam>=VK_LEFT && GetSelectionType()==SELECTION_TYPE_Control)
                {
                    // let go to the host
                }
                else
                {
                    hr = pElemTarget->HandleMessage(pMessage);
                    if(hr != S_FALSE)
                    {
                        goto Cleanup;
                    }
                }
            }

            // Raid 44891
            // Some hosts like AOL, CompuServe and MSN eat up the Enter Key in their
            // TranslateAccelerator, so we never get it in our WindowProc. We work around
            // by explicitly translating WM_KEYDOWN+VK_RETURN to WM_CHAR+VK_RETURN
            if(fGotEnterKey && !pElemTarget->IsEditable(TRUE))
            {
                if(pElemTarget->_fActsLikeButton)
                {
                    fTranslateEnterKey = TRUE;
                }
                else
                {
                    switch(pElemTarget->Tag())
                    {
                    case ETAG_A:
                    case ETAG_IMG:
                    case ETAG_TEXTAREA:
                        fTranslateEnterKey = TRUE;
                        break;
                    }
                }
                if(fTranslateEnterKey)
                {
                    break;
                }
            }

            // Find the next target
            pNodeTarget = pNodeTarget->Parent();
            if(pNodeTarget)
            {
                pElemTarget = pNodeTarget->Element();
            }
            else
            {
                pElemTarget = pElemTarget->MarkupMaster();
                pNodeTarget = pElemTarget->GetFirstBranch();
            }
        }
        // Pressing 'Enter' should activate the default button
        // (unless the focus is on a SELECT - IE5 #64133)
        if(fGotEnterKey && !fTranslateEnterKey && _pElemCurrent->Tag()!=ETAG_SELECT)
        {
            fTranslateEnterKey = !!_pElemCurrent->FindDefaultElem(TRUE);
        }
        if(fTranslateEnterKey)
        {
            ::TranslateMessage(pMessage);
            hr = S_OK;
            goto Cleanup;
        }

        Assert(hr == S_FALSE);

        if(IsFrameTabKey(pMessage) || IsTabKey(pMessage) || IsValidAccessKey(this, pMessage))
        {
            hr = HandleKeyNavigate(pMessage, FALSE);

            if(hr != S_FALSE)
            {
                goto Cleanup;
            }

            // Comment (jenlc). Say that the document has two frames, the
            // first frame has two controls with access keys ALT+A and ALT+B
            // respectively while the second frame has a control with access
            // key ALT+A. Suppose currently the focus is on the control with
            // access key ALT+B (the second control of the first frame) and
            // ALT+A is pressed, which control should get the focus? Currently
            // Trident let the control in the second frame get the focus.
            if(IsTabKey(pMessage) || IsFrameTabKey(pMessage))
            {
                // Clear any selection
                NotifySelection( SELECT_NOTIFY_DESTROY_ALL_SELECTION, NULL );        

                _pPrimaryMarkup->Root()->BecomeCurrentAndActive();
            }
        }
    }

    if(IsFrameTabKey(pMessage) || IsTabKey(pMessage) || IsValidAccessKey(this, pMessage))
    {
        if(hr == S_OK)
        {
            _pElemUIActive = NULL;
        }
        else if(hr == S_FALSE)
        {
            hr = HandleKeyNavigate(pMessage, TRUE);

            if(hr==S_FALSE && pMessage->message!=WM_SYSKEYDOWN)
            {
                CElement* pElement = GetPrimaryElementClient();
                if(pElement)
                {
                    pElement->BecomeCurrentAndActive(NULL, pMessage->lSubDivision, TRUE);
                    hr = S_OK;  
                }
            }
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::AreWeSaneAfterEventFiring
//
//  Synopsis:   Determines if anything during the event firing has changed
//              things adversly.
//
//----------------------------------------------------------------------------
BOOL CDocument::AreWeSaneAfterEventFiring(CMessage* pMessage, ULONG cDie)
{
    BOOL fAreWeSane = FALSE;

    if((pMessage->message>=WM_MOUSEFIRST && pMessage->message<=WM_MOUSELAST)
        || pMessage->message==WM_SETCURSOR)
    {
        if(pMessage->pNodeHit && !pMessage->pNodeHit->Element()->IsInMarkup())
        {
            goto Cleanup;
        }
    }

    fAreWeSane = TRUE;

Cleanup:
    return fAreWeSane;
}

BOOL CDocument::FCaretInPre()
{
    if(_pCaret)
    {
        CTreeNode* pNode = _pCaret->GetNodeContainer(MPTR_SHOWSLAVE);

        if(pNode)
        {
            const CParaFormat* pPF = pNode->GetParaFormat();
            if(pPF)
            {
                return pPF->_fPre;
            }
        }
    }

    return FALSE;
}

//+---------------------------------------------------------------
//
// Local Helper: ShowTooltipHelper
//
//----------------------------------------------------------------
void ShowTooltipHelper(CDocument* pDoc, CTreeNode* pNodeContext, CMessage* pMsg)
{
    {
        BOOL fDismissed;

        // Ignore spurious WM_ERASEBACKGROUNDs generated by tooltips
        CServer::CLock Lock(pDoc, SERVERLOCK_IGNOREERASEBKGND);

        //  Give tooltips a chance to dismiss
        fDismissed = FormsTooltipMessage(pMsg->message, pMsg->wParam, pMsg->lParam);

        // If we're not captured and the tooltip is being dismissed....
        if(fDismissed == FALSE)
        {
            //  If hitted element has tooltip, put up its tooltip, otherwise,
            //  Walk through the element hierarchy.  If any element
            //  has tooltip text, put up the tooltip.
            CTreeNode* pNode = pNodeContext;
            for(; pNode&&pNode->Tag()!=ETAG_ROOT; pNode=pNode->Parent())
            {
                if(pNode->Element()->ShowTooltip(pMsg, pMsg->pt) != S_FALSE)
                {
                    break;
                }
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnMouseMessage
//
//  Synopsis:   Handle WM_MOUSEMOVE, WM_LBUTTONDOWN and so on.
//
//----------------------------------------------------------------------------
HRESULT CDocument::OnMouseMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult, int x, int y)
{
    HRESULT     hr  = S_OK;
    CMessage    Message(_pInPlace->_hwnd, msg, wParam, lParam);
    CTreeNode*  pNodeHit = NULL;
    CTreeNode*  pNodeNewMouse = NULL;
    BOOL        fCapture;
    DWORD       dwHitTestFlags;
    ULONG       cDie = _cDie;

    Message.pt.x = x;
    Message.pt.y = y;

    //--------------------------------------------------------------------
    //
    // NOTE(SujalP and GaryBu):
    //
    // Normally, a button down implies a BecomeCurrent() which implies
    // a transition to the UIActive state. The BecomeCurrent() happens
    // in PumpMessage(). However, we do not call BecomeCurrent() when
    // the mouse goes down on a scrollbar. This is needed to fix bugs
    // like bug48127, where clicking on the scrollbar changes the current
    // site.
    //
    // However, if we do not perform BecomeCurrent() on the button down
    // in PumpMessage(), then we have to atleast transition of UIActive
    // otherwise we will break the frame case. Consider a doc with 2 frames.
    // Lets say the focus is on the left frame. The user now clicks on
    // the right frame scrollbar. The right frame will not call BecomeCurrent
    // because the hit was on the scrollbar and hence not even transition to
    // UIActive state. Hence, we do the transition here. Note that the
    // transition has to be on a button down (not on a button up -- as it
    // was done originally) because if the object takes capture then we
    // may not get the button up message at all (scrollbar are guilty of
    // this).
    //
    // If non-client area like scrollbars are hit, then under certain hosts
    // we do not want to become UIActive either. Consider the case of Athena
    // (bug33562) where clicking on the scrollbar, takes focus away from
    // the "To:" input box of Athena. To prevent us from taking focus, we
    // ask hosts such as Athena to turn on the following doc host flag:
    // DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY. Note that we will still take
    // focus if the hit were on the body (done by PumpMessage() as explained
    // earlier).
    //
    //--------------------------------------------------------------------

    // We need to check to see that we really have the capture
    // because WM_CAPTURECHANGED is not sent on all platforms.
    fCapture = FALSE;
    dwHitTestFlags = 0;
    if(_pvCaptureObject && (::GetCapture()==_pInPlace->_hwnd))
    {
        fCapture = TRUE;

        if(!_fVirtualHitTest)
        {
            dwHitTestFlags = HT_IGNORESCROLL | HT_NOGRABTEST;
        }
        else
        {
            dwHitTestFlags = HT_IGNORESCROLL | HT_NOGRABTEST | HT_VIRTUALHITTEST;
        }
    }

    // Locate the "hit" element
    Message.htc = HitTestPoint(&Message, &pNodeHit, dwHitTestFlags);

    if(_pvCaptureObject==NULL && Message.htc==HTC_NO)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    Assert(!pNodeHit || pNodeHit->IsInMarkup());

    Message.SetNodeHit(pNodeHit);

    if(Message.pNodeHit)
    {
        pNodeHit->Element()->SubDivisionFromPt(Message.ptContent, &Message.lSubDivision);
    }

    pNodeNewMouse = Message.pNodeHit;

    {
        // don't let the hit-tested element go away til we're ready
        if(pNodeNewMouse)
        {
            CTreeNode::CLock Lock(pNodeNewMouse);
        }

        // If the mouse is captured, mouse events go to capture fn,
        // otherwise mouse events go to site/element under the mouse
        if(fCapture)
        {
            hr = PumpMessage(&Message, NULL);
        }
        else
        {
            CTreeNode* pNodeOldOver = _pNodeLastMouseOver;

            // Someone stole the capture
            if(_pvCaptureObject)
            {
                ClearMouseCapture(_pvCaptureObject);
            }

            // Hand the message to the site.  However, if the htc is >= HTC_GRPTOPBORDER
            // we can just hand it off to the parent.  htc >= HTC_GRPTOPBORDER implies
            // a hit on the grab handles or border.
            if((Message.htc>=HTC_GRPTOPBORDER) && Message.pNodeHit->Element()->IsEditable(TRUE))
            {
                pNodeNewMouse = Message.pNodeHit->GetUpdatedParentLayoutNode();
                pNodeNewMouse = (pNodeNewMouse && Message.pNodeHit->Element()->IsInMarkup())
                    ? pNodeNewMouse : Message.pNodeHit;
            }

            // Do not send MouseOver/MouseLeave messages in response to a
            // WM_SETCURSOR! WM_SETCURSOR is not a true mouse message and Windows
            // can send us this message any time it thinks the cursor needs to be
            // redrawn. This can cause nasty situations if the script handler for
            // MouseOver/MouseOut generates another WM_SETCURSOR message! See bug
            // 13590. (MohanB)

            // deal with the last element under mouse
            // BUGBUG (MohanB) SELECT needs to fixed to pass MOUSEMOVE to CDOc and and then
            // this hack should be removed.
            if(Message.message!=WM_SETCURSOR
                && cDie==_cDie
                && _pNodeLastMouseOver
                && pNodeNewMouse
                && (_pNodeLastMouseOver!=pNodeNewMouse
                || _lSubDivisionLast!=Message.lSubDivision)

                // Ignore transitions from master to slave and vice versa (39135)
                && !(_pNodeLastMouseOver->Tag()==ETAG_TXTSLAVE
                && _pNodeLastMouseOver->Element()->MarkupMaster()==pNodeNewMouse->Element()
                || pNodeNewMouse->Tag()==ETAG_TXTSLAVE
                && pNodeNewMouse->Element()->MarkupMaster()==_pNodeLastMouseOver->Element()
                )
                )
            {
                CTreeNode* pNodeTo = pNodeNewMouse->Ancestor(ETAG_A);

                CMessage MessageOut(_pInPlace->_hwnd, WM_MOUSELEAVE, NULL, lParam);
                MessageOut.SetNodeHit(pNodeOldOver);
                MessageOut.pt.x = x;
                MessageOut.pt.y = y;
                MessageOut.lSubDivision = _lSubDivisionLast;

                // set up for the 'from' & 'to' event object parameters. This part
                // is tricky. we are firing and WM_MOUSEEXIT, the this pointer is
                // the 'from' element, the 'to' element is accessed by TEMPORARILY
                // putting it into the _pNodeLastMouseOver pointer (since it is
                // redundant with the this pointer anyhow). after this call to HM
                // (which assumes the _pelemlast.. is the 'to' pointer (see CElement::
                // FireStdEventOnMessage)) the _pNodeLastMouseOver is restored for
                // the remainder of the HM calls.
                _pNodeLastMouseOver = pNodeNewMouse;
                _pNodeLastMouseOver->NodeAddRef();
                // if the timer is set, we want to turn it off. otherwise it is possible
                //  that the _pNodeLastMouseOver will be Nulled out from under use (root
                //  cause of bug 21062)
                if(_fMouseOverTimer)
                {
                    FormsKillTimer(this, TIMER_ID_MOUSE_EXIT);
                    _fMouseOverTimer = FALSE;
                }

                // Cancel mouse hover tracking on mouse leave
                TRACKMOUSEEVENT tme;
                tme.cbSize = sizeof(TRACKMOUSEEVENT);
                tme.dwFlags = TME_HOVER|TME_CANCEL;
                tme.hwndTrack = _pInPlace->_hwnd;
                _TrackMouseEvent(&tme);

                hr = PumpMessage(&MessageOut, pNodeOldOver);
                //restore to fire the message itself
                if(_pNodeLastMouseOver)
                {
                    _pNodeLastMouseOver->NodeRelease();
                }

                _pNodeLastMouseOver = pNodeOldOver;

                // Is this node no longer connected to the primary markup? (#30760)
                if(!_pNodeLastMouseOver->Element()->IsInPrimaryMarkup())
                {
                    CElement* pElemMaster = _pNodeLastMouseOver->Element()->MarkupMaster();

                    while(pElemMaster && !pElemMaster->IsInPrimaryMarkup())
                    {
                        pElemMaster = pElemMaster->MarkupMaster();
                    }
                    if(!pElemMaster)
                    {
                        _pNodeLastMouseOver->NodeRelease();
                        _pNodeLastMouseOver = NULL;
                    }
                }
            }

            // Watch for the mouse going off the window.
            if(!_fMouseOverTimer)
            {
                hr = FormsSetTimer(this,
                    ONTICK_METHOD(CDocument, DetectMouseExit, detectmouseexit),
                    TIMER_ID_MOUSE_EXIT, 300);
                if(!hr)
                {
                    _fMouseOverTimer = TRUE;
                }
            }

            // Handle tooltips
            if(Message.pNodeHit)
            {
                ShowTooltipHelper(this, Message.pNodeHit, &Message);
            }

            hr = PumpMessage(&Message, pNodeNewMouse);

            // Fire the event unless the element got nuked in event handler!
            if(pNodeNewMouse && pNodeNewMouse->IsInMarkup() && !pNodeNewMouse->IsDead())
            {
                // Now fire MouseOver if applicable
                // BUGBUG (MohanB) SELECT needs to fixed to pass MOUSEMOVE to CDOc and and then
                // this hack should be removed.
                if(Message.message!=WM_SETCURSOR
                    && cDie==_cDie
                    && (_pNodeLastMouseOver!=pNodeNewMouse
                    || _lSubDivisionLast!=Message.lSubDivision)
                    // Ignore transitions from master to slave and vice versa (39135)
                    && !(_pNodeLastMouseOver
                    && (_pNodeLastMouseOver->Tag()==ETAG_TXTSLAVE
                    && _pNodeLastMouseOver->Element()->MarkupMaster()==pNodeNewMouse->Element()
                    || pNodeNewMouse->Tag()==ETAG_TXTSLAVE
                    && pNodeNewMouse->Element()->MarkupMaster()==_pNodeLastMouseOver->Element())
                    ))
                {
                    Assert(_pInPlace && (_pInPlace->_hwnd));
                    CMessage MessageOut(_pInPlace->_hwnd, WM_MOUSEOVER, NULL, lParam);
                    MessageOut.SetNodeHit(pNodeNewMouse);
                    MessageOut.pt.x = x;
                    MessageOut.pt.y = y;
                    MessageOut.lSubDivision = Message.lSubDivision;

                    hr = PumpMessage(&MessageOut, pNodeNewMouse);

                    // Start mouse hover tracking after sending mouse over message.
                    TRACKMOUSEEVENT tme;
                    tme.cbSize = sizeof(TRACKMOUSEEVENT);
                    tme.dwFlags = TME_HOVER;
                    tme.hwndTrack = _pInPlace->_hwnd;
                    tme.dwHoverTime = HOVER_DEFAULT;
                    _TrackMouseEvent(&tme);

                    // don't assign if this is the rootsite
                    if(pNodeNewMouse->Tag() != ETAG_ROOT)
                    {
                        CTreeNode::ReplacePtr(&_pNodeLastMouseOver, pNodeNewMouse);
                        _lSubDivisionLast = MessageOut.lSubDivision;
                    }
                }
            }
        }

    }

    // Ensure that UI Active site matches the current site on mouse button up.
    if((msg==WM_LBUTTONUP||msg==WM_RBUTTONUP||msg==WM_MBUTTONUP) &&
        (_pElemUIActive!=_pElemCurrent) && (cDie==_cDie))
    {
        _pElemCurrent->BecomeUIActive();
    }

    if(msg!=WM_MOUSEMOVE && msg!=WM_SETCURSOR)
    {
        DeferUpdateUI();
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//
// Member: DetectMouseExit
//
// Sysnopsis:  Timer Callback. this function's job is to detect when a mouse move
//      has taken us outside the client rectl. when that is detected, we fire
//      an exit message.
//              in any case we clear the doc flag, and we kill the timer.
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::DetectMouseExit(UINT uTimerID)
{
    if(_pNodeLastMouseOver)
    {
        POINT   ptMouse;
        HWND    hwndMouse;

        GetCursorPos(&ptMouse);

        // don't test for child windows in this if statement.  mouseout is 
        // NOT about containership (e.g. moving a mouse over a select on a body
        // SHOULD fire mouseout on the body).  if the mouse is over a HWND that 
        // is not the inplace hwnd then we've moved over a different object
        // and should fire the out.  This is especially poignent for moving
        // over and Iframe (bug 47607) or a frame
        if((hwndMouse = WindowFromPoint(ptMouse))!=NULL &&
            (hwndMouse!=_pInPlace->_hwnd) && !_fInhibitOnMouseOver)
        {
            CMessage    Message(_pInPlace->_hwnd, WM_MOUSELEAVE, NULL, NULL);
            CTreeNode*  pNode = _pNodeLastMouseOver;

            Message.SetNodeHit(pNode);
            Message.pt.x = -1; // we know we are outside client rect
            Message.pt.y = -1;

            // Don't need timer any more.
            FormsKillTimer(this, uTimerID);
            _fMouseOverTimer = FALSE;

            // Set this first in order to properly fill the EVENTPARAM from,to
            _pNodeLastMouseOver = NULL;
            _lSubDivisionLast = 0;

            PumpMessage(&Message, pNode);

            // Release our addref on the lastmouseoverelem.
            pNode->NodeRelease();
        }
    }

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::DeferSetCursor
//
//  Synopsis:   After scrolling we want to post a setcursor message to our
//              window so that the cursor shape will get updated. However,
//              because of nested scrolls, we might end up with multiple
//              setcursor calls. To avoid this, we will post a method call
//              to the function which actually does the postmessage. During
//              the setup of the method call we will delete any existing
//              callbacks, and hence this will delete any existing callbaks.
//
//----------------------------------------------------------------------------
void CDocument::DeferSetCursor()
{
    GWKillMethodCall(this, ONCALL_METHOD(CDocument, SendSetCursor, SendSetCursor), 0);
    GWPostMethodCall(this, ONCALL_METHOD(CDocument, SendSetCursor, SendSetCursor),
        0, FALSE, "CDocument::SendSetCursor");
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::SendMouseMessage
//
//  Synopsis:   This function actually posts the message to our window.
//
//----------------------------------------------------------------------------
void CDocument::SendSetCursor(DWORD_PTR dwContext)
{
    // First be sure that we are all OK
    // Do this only if we have focus. We don't want to generate mouse events
    // when we do not have focus (bug 9144)
    if(_pInPlace && _pInPlace->_hwnd && HasFocus())
    {
        CPoint pt;
        CRect rc;

        ::GetCursorPos(&pt);
        ::GetClientRect(_pInPlace->_hwnd, &rc);

        // Next be sure that the mouse is in our client rect and only then
        // post ourselves the message.
        if(rc.Contains(pt))
        {
            ::PostMessage(_pInPlace->_hwnd, WM_SETCURSOR, (WORD)_pInPlace->_hwnd, HTCLIENT);
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::EnableDragDrop
//
//  Synopsis:   Register or revoke drag-drop as appropriate.
//
//-------------------------------------------------------------------------
void CDocument::EnableDragDrop(DWORD_PTR dwContext)
{
    IDropTarget* pDT;

    {
        if(!GetDropTarget(&pDT))
        {
            BOOL fRegHostDT = FALSE;

            if(!fRegHostDT)
            {
                RegisterDragDrop(_pInPlace->_hwnd, pDT);
            }

            pDT->Release();
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::SetMouseCapture
//
//  Synopsis:   Mouse capture utilities.
//
//----------------------------------------------------------------------------
void CDocument::SetMouseCapture(PFN_VOID_MOUSECAPTURE pfnTo, void* pvObject, BOOL fElement)
{
    HWND hwndCapture;
    BOOL fAddRef = TRUE;

    // If we've got a capture object which is different from the argument
    // object, tell it that we're losing capture
    if(_pvCaptureObject)
    {
        if(_pvCaptureObject != pvObject)
        {
            CMessage Message((MSG*)NULL);

            Message.message = WM_CAPTURECHANGED;

            // Tell the captured object that it's losing the capture
            PumpMessage(&Message, NULL);

            // If capture object was an element, release it
            if(_fElementCapture && _pvCaptureObject)
            {
                ((CElement*)_pvCaptureObject)->Release();
            }
        }
        else
        {
            // Since we're not changing the capture object, we don't want
            // to AddRef() it below.
            fAddRef = FALSE;
        }
    }

    _pvCaptureObject = pvObject;
    _pfnCapture = pfnTo;
    _fElementCapture = fElement;

    // If we're setting up a new capture object and it's an element,
    // AddRef() it now.
    if(pvObject && fElement && fAddRef)
    {
        ((CElement*)pvObject)->AddRef();
    }

    hwndCapture = ::GetCapture();
    {
        // If we're currently captured, but we're releasing capture
        if(hwndCapture == _pInPlace->_hwnd)
        {
            if(!_pvCaptureObject && !_pElementOMCapture)
            {
                ::ReleaseCapture();
            }
        }
        else
        {
            if( _pvCaptureObject )
            {
                ::SetCapture(_pInPlace->_hwnd);
            }
        }
    }
}

//+--------------------------------------------------------------------------
//
//  Member:     CDocument::ClearMouseCapture
//
//  Synopsis:   Releases the capture object without notification. Used by
//              objects to revoke their own capture.
//
//---------------------------------------------------------------------------
void CDocument::ClearMouseCapture(void* pvObject)
{
    if(_pvCaptureObject)
    {
        if(_fElementCapture)
        {
            ((CElement*)_pvCaptureObject)->Release();
        }
        _pvCaptureObject = NULL;
        _pfnCapture = NULL;
        _fElementCapture = FALSE;
    }
}

//+====================================================================================
//
// Method:  GetSelectionType
//
// Synopsis:Check the current selection type of the selection manager
//
//------------------------------------------------------------------------------------
SELECTION_TYPE CDocument::GetSelectionType()
{
    SELECTION_TYPE theType = SELECTION_TYPE_None;

    if(_pIHTMLEditor)
    {
        _pIHTMLEditor->GetSelectionType(&theType);
    }

    return theType;
}

//+====================================================================================
//
// Method: HasTextSelection
//
// Synopsis: Is there a "Text Selection"
//
//------------------------------------------------------------------------------------
BOOL CDocument::HasTextSelection()
{
    return (GetSelectionType() == SELECTION_TYPE_Selection);
}

//+====================================================================================
//
// Method: HasSelection
//
// Synopsis: Is there any form of Selection ?
//
//------------------------------------------------------------------------------------
BOOL CDocument::HasSelection()
{
    return (GetSelectionType() != SELECTION_TYPE_None);
}

//+====================================================================================
//
// Method: PointInSelection
//
// Synopsis: Is the given point in a Selection ? Returns false if there is no selection
//
//------------------------------------------------------------------------------------
BOOL CDocument::IsPointInSelection(POINT pt, CTreeNode* pNode, BOOL fPtIsContent)
{
    // BUGBUG (MohanB) This function does not check for clipping, because
    // MovePointerToPoint always does virtual hit-testing. We should pass
    // in an argument fDoVirtualHitTest to MovePointerToPoint() and set that
    // argument to FALSE when calling from this function.
    HRESULT hr = S_OK;
    IMarkupPointer* pPointer = NULL;
    BOOL fBOL = FALSE;
    BOOL fAtLogicalBOL = FALSE;
    IHTMLEditor* ped = NULL;
    BOOL fPointInSelection = FALSE;
    BOOL fRightOfCp = FALSE;
    CMarkupPointer* pPointerInternal;
    IHTMLElement* pIElementOver = NULL;
    CElement* pElement = NULL;
    SELECTION_TYPE eType = GetSelectionType();

    if(eType==SELECTION_TYPE_Control || eType==SELECTION_TYPE_Selection)
    {
        // marka BUGBUG - consider making this take the node - and work out the pointer more directly
        hr = CreateMarkupPointer(&pPointer);
        if(hr)
        {
            goto Cleanup;
        }

        if(pNode && fPtIsContent)
        {
            hr = pPointer->QueryInterface(CLSID_CMarkupPointer, (void**)&pPointerInternal);
            if(hr)
            {
                goto Cleanup;
            }

            hr = MovePointerToPointInternal(
                pt,
                pNode,
                pPointerInternal,
                &fBOL,
                &fAtLogicalBOL,
                &fRightOfCp,
                FALSE,
                GetFlowLayoutForSelection(pNode));
            if(hr)
            {
                goto Cleanup;
            }
        }
        else
        {
            hr = MoveMarkupPointerToPoint(pt, pPointer, &fBOL, &fAtLogicalBOL, &fRightOfCp, FALSE);
            if(hr)
            {
                goto Cleanup;
            }
        }
        ped = GetHTMLEditor(FALSE);
        Assert(ped);

        if(pNode)
        {
            pElement = pNode->Element();
        }            
        if(pElement)
        {
            hr = pElement->QueryInterface(IID_IHTMLElement, (void**)&pIElementOver);
            if(hr)
            {
                goto Cleanup;
            }
        }

        hr = ped->IsPointerInSelection(pPointer, &fPointInSelection , &pt, pIElementOver);
    }

Cleanup:
    ReleaseInterface(pPointer);
    ReleaseInterface(pIElementOver);
    return fPointInSelection;
}

//+====================================================================================
//
// Method: GetEditingServices
//
// Synopsis: Get a ref-counted IHTMLEditingServices, 
//           forcing creation of the editor if there isn't one
//
//------------------------------------------------------------------------------------
HRESULT CDocument::GetEditingServices(IHTMLEditingServices** ppIServices)
{
    HRESULT hr = S_OK;

    IHTMLEditor* ped = GetHTMLEditor(TRUE);
    if(ped)
    {
        hr = ped->QueryInterface(IID_IHTMLEditingServices, (void**)ppIServices);
    }
    else
    {
        hr = E_FAIL;
    }

    RRETURN(hr);
}

//+====================================================================================
//
// Method: Select
//
// Synopsis: 'Select from here to here' a wrapper to the selection manager.
//
//------------------------------------------------------------------------------------
HRESULT CDocument::Select(IMarkupPointer* pStart, IMarkupPointer* pEnd, SELECTION_TYPE eType)
{
    HRESULT hr = S_OK;
    IHTMLEditingServices* pEdServices = NULL;
    DWORD followUpAction = FOLLOW_UP_ACTION_None;
    IHTMLEditor* pEditor = GetHTMLEditor();
    
    if(!pEditor)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    hr = EnsureEditContext(pStart);
    if(hr)
    {
        AssertSz(0, "Ensure Edit Context failed");
        goto Cleanup;
    }

    hr = pEditor->QueryInterface(IID_IHTMLEditingServices, (void**)&pEdServices);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pEdServices->Select(pStart, pEnd, eType, &followUpAction);  

    if(followUpAction != FOLLOW_UP_ACTION_None)
    {
        ProcessFollowUpAction(NULL, followUpAction);
    }

Cleanup:
    ReleaseInterface(pEdServices);

    RRETURN(hr);
}

//+====================================================================================
//
// Method: IsElementSiteselectable
//
// Synopsis: Determine if a given elemnet is site selectable by asking mshtmled.dll.
//
//------------------------------------------------------------------------------------
BOOL CDocument::IsElementSiteSelectable(CElement* pCurElement)
{
    HRESULT hr = S_OK;
    HRESULT hrSiteSelectable = S_FALSE;
    IHTMLEditor* ped = NULL;
    IHTMLEditingServices* pIEditingServices = NULL;
    IHTMLElement* pICurElement = NULL;    

    ped = GetHTMLEditor(TRUE);        
    Assert(ped);

    if(ped)
    {
        hr = ped->QueryInterface(IID_IHTMLEditingServices, (void**)&pIEditingServices);
        if(hr)
        {
            goto Cleanup;
        }
        hr = pCurElement->QueryInterface(IID_IHTMLElement, (void**)&pICurElement);
        if(hr)
        {
            goto Cleanup;
        }

        hrSiteSelectable = pIEditingServices->IsElementSiteSelectable(pICurElement);
    }

Cleanup:
    ReleaseInterface(pICurElement);
    ReleaseInterface(pIEditingServices);

    return (hrSiteSelectable == S_OK);
}

//+====================================================================================
//
// Method: IsElementUIActivatable
//
// Synopsis: Determine if a given elemnet is site selectable by asking mshtmled.dll.
//
//------------------------------------------------------------------------------------
BOOL CDocument::IsElementUIActivatable(CElement* pCurElement)
{
    HRESULT hr = S_OK;
    HRESULT hrActivatable = S_FALSE;
    IHTMLEditor* ped = NULL;
    IHTMLEditingServices* pIEditingServices = NULL;
    IHTMLElement* pICurElement = NULL;    

    ped = GetHTMLEditor(TRUE);        
    Assert(ped);

    if(ped)
    {
        hr = ped->QueryInterface(IID_IHTMLEditingServices, (void**)&pIEditingServices);
        if(hr)
        {
            goto Cleanup;
        }
        hr = pCurElement->QueryInterface(IID_IHTMLElement, (void**)&pICurElement);
        if(hr)
        {
            goto Cleanup;
        }

        hrActivatable = pIEditingServices->IsElementUIActivatable(pICurElement);
    }

Cleanup:
    ReleaseInterface(pICurElement);
    ReleaseInterface(pIEditingServices);

    return (hrActivatable == S_OK);
}

//+====================================================================================
//
// Method: IsElementSiteselected
//
// Synopsis: Determine if a given elemnet is currently site selected 
//
//------------------------------------------------------------------------------------
BOOL CDocument::IsElementSiteSelected(CElement* pCurElement)
{
    HRESULT hr = S_OK;
    HRESULT hrSiteSelected = S_FALSE;
    IHTMLEditor* ped = NULL;
    IHTMLElement* pICurElement = NULL;    

    ped = GetHTMLEditor(TRUE);        
    Assert(ped);

    if(ped && pCurElement)
    {
        hr = pCurElement->QueryInterface(IID_IHTMLElement, (void**)&pICurElement);
        if(hr)
        {
            goto Cleanup;
        }

        hrSiteSelected = ped->IsElementSiteSelected(pICurElement);
    }

Cleanup:
    ReleaseInterface(pICurElement);
    return (hrSiteSelected == S_OK);
}

//+====================================================================================
//
// Method: ProcessFollowUpAction
//
// Synopsis: Interpret the FOLLOW_UP_ACTION code given to use after processing a
//           message from the HTMLEditor
//
// Note:     Returning S_FALSE says we didn't do anything.  We only return S_OK from
//           a return to a call to DragElement
//
//------------------------------------------------------------------------------------
HRESULT CDocument::ProcessFollowUpAction(CMessage* pMessage, DWORD dwFollowUpAction)
{
    HRESULT hr = S_FALSE;

    Assert(dwFollowUpAction != FOLLOW_UP_ACTION_None);

    // BUGBUG - consider shifting bits and testing ?
    //********************************************************
    //           !!! WARNING WARNING WARNING !!!
    //
    //  Do not change the order of bits it is significant
    //
    //********************************************************
    if((dwFollowUpAction&FOLLOW_UP_ACTION_OnClick) != 0)
    {
        SetClick(pMessage);
    }
    if((dwFollowUpAction&FOLLOW_UP_ACTION_DragElement) != 0)
    {
        hr = DragElement(pMessage);
        if(hr == S_FALSE)
        {
            hr = S_OK;
        }
    }

    RRETURN1(hr, S_FALSE);
}

//+========================================================================
//
// Method: ShouldCreateHTMLEditor
//
// Synopsis: Certain messages require the creation of a selection ( like Mousedown)
//           For these messages return TRUE.
//
//          Or if the Host will host selection Manager - return TRUE.
//-------------------------------------------------------------------------
BOOL CDocument::ShouldCreateHTMLEditor(CMessage* pMessage)
{
    // If this is a MouseDown message, we should force a TSR to be created
    // Per Bug 18568 we should also force a TSR for down/up/right/left arrows
    switch(pMessage->message)
    {
    case WM_LBUTTONDOWN:
        return TRUE;

        // marka BUGBUG - this may not be required anymore.
    case WM_KEYDOWN:
        switch(pMessage->wParam)
        {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
            return TRUE;
        }
    default:
        return FALSE;
    }
}

//+====================================================================================
//
// Method:ShouldCreateHTMLEditor
//
// Synopsis: Should we force the creation of a selection manager for this type of notify ?
//
//------------------------------------------------------------------------------------
BOOL CDocument::ShouldCreateHTMLEditor(SELECTION_NOTIFICATION eSelectionNotification)
{
    return FALSE;
}

//+========================================================================
//
// Method: ShouldSetEditContext.
//
// Synopsis: On Mouse Down we should always set a new EditContext
//
//-------------------------------------------------------------------------
BOOL CDocument::ShouldSetEditContext(CMessage* pMessage)
{
    switch(pMessage->message)
    {
    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
        return TRUE;
    default:
        return FALSE;
    }
}

//+================================================================================
//
// Method: HandleSelMgrMessageCapture
//
// Synopsis: Handle a Selection Manager Message - when the Selection Manager
//           has capture.
//
//          This method is given as a Pointer to a F'n when SetHTMLEditorCapture
//--------------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::HandleEditMessageCapture(CMessage* pMessage)
{
    HRESULT             hr = S_OK;    
    DWORD               theAction = FOLLOW_UP_ACTION_None;
    SelectionMessage    theMessage;

    if(pMessage->pNodeHit==NULL || !pMessage->pNodeHit->IsDead())
    {    
        CMessageToSelectionMessage(pMessage, &theMessage);
        theMessage.fFromCapture = TRUE; // Let them know that the actual point is from me...
        HRESULT hrProcessFollowUp; 

        // Pass on to the editor
        if(_pIHTMLEditor && _fInEditCapture)
        {
            hr = _pIHTMLEditor->HandleMessage(&theMessage, &theAction);

            // BUGBUG (MohanB) Need to copy any other info?
            // Perhaps call SelectionMessageToCMessage?
            pMessage->fStopForward = theMessage.fStopForward;

            // Process follow up action, if any. BTW: This is the only reason we need a 
            // CMessage above. If we kill this, populate the SelectionMessage directly.
            if(pMessage->pNodeHit!=NULL && theAction!=FOLLOW_UP_ACTION_None)
            {
                hrProcessFollowUp = ProcessFollowUpAction(pMessage, theAction);
                if(hrProcessFollowUp != S_OK)
                {
                    hr = hrProcessFollowUp;
                }
            }
        }
    }

    // Dont return an Error code for failure - it makes SetEditCapture not work.
    return hr;
}

//+====================================================================================
//
// Method: OnSelectDblClickTimer
//
// Synopsis: Callback that for WM_TIMER messages inside Trident. Relays the Timer
//           messages back to the HTMLEditor's OnTimerTick method.
//
//------------------------------------------------------------------------------------
HRESULT CDocument::OnEditDblClkTimer(UINT idTimer)
{
    HRESULT hr;

    Assert(_pIHTMLEditor);

    hr = NotifySelection(SELECT_NOTIFY_TIMER_TICK, NULL);

    RRETURN(hr);
}

//+====================================================================================
//
// Method: SetClick
//
// Synopsis: Enable the Setting and passing of click messages
//
//------------------------------------------------------------------------------------
VOID CDocument::SetClick(CMessage* pMessage)
{
    BOOL fEditable = FALSE;
    CFlowLayout* pFlowUp = pMessage->pNodeHit->GetFlowLayout();

    if(pFlowUp)
    {
        fEditable = pFlowUp->ElementOwner()->IsEditable(TRUE);
    }
    else
    {
        fEditable = pMessage->pNodeHit->Element()->IsEditable(TRUE);

    }
    if(!fEditable)
    {
        pMessage->SetNodeClk(pMessage->pNodeHit);
    }
}

//+============================================================================
//
// Method: DragElement
//
// Synopsis: Manage a drag from this Message
//
// BUGBUG - this could be handle via ViewServ DragElement. Do it that way some time
//
//-----------------------------------------------------------------------------
HRESULT CDocument::DragElement(CMessage* pMessage)
{
    Assert(_pElemEditContext);
    /*CParentUndo pu(this); wlw note*/
    HRESULT hr;
    CFlowLayout* pFlowLayout = NULL;

    /*pu.Start(IDS_UNDOMOVE);
    {
        CSelectionUndo Undo(_pElemCurrent, GetCurrentMarkup());
    } wlw note*/

    if(pMessage->pNodeHit && pMessage->pNodeHit->Element() &&
        _pElemEditContext && _pElemEditContext->GetFirstBranch())
    {    
        pFlowLayout = GetFlowLayoutForSelection(_pElemEditContext->GetFirstBranch());
        if(!pFlowLayout)
        {
            hr = E_FAIL;
            goto Cleanup;
        }

        pMessage->pNodeHit->Element()->DragElement(pFlowLayout, pMessage->dwKeyState, 
            NULL, pMessage->lSubDivision);

        hr = S_OK;
    }
    else 
    {
        hr = E_FAIL;
    }
    // BUGBUG marka - make Drag/Drop return a failure code here - so we give
    // the right value to the failure code.
    /*{
        CDeferredSelectionUndo DeferredUndo(GetCurrentMarkup());
    } wlw note*/

Cleanup:
    /*pu.Finish(hr); wlw note*/
    RRETURN(hr);
}

// This function executes given timeout script and kills the associated timer
HRESULT CDocument::ExecuteTimeoutScript(TIMEOUTEVENTINFO* pTimeout)
{
    HRESULT hr = S_OK;

    Assert(pTimeout != NULL);

    Verify(FormsKillTimer(this, pTimeout->_uTimerID) == S_OK);

    if(pTimeout->_pCode)
    {
        DISPPARAMS  dp = _afxGlobalData._Zero.dispparams;
        CVariant    varResult;
        EXCEPINFO   excepinfo;

        // Can't disconnect script engine while we're executing script.
        CDocument::CLock Lock(this);

        hr = pTimeout->_pCode->Invoke(
            DISPID_VALUE,
            IID_NULL,
            0,
            DISPATCH_METHOD,
            &dp,
            &varResult,
            &excepinfo,
            NULL);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::SetUIActiveElement
//
//  Synopsis:   UIActivate a given site, often as part of
//              IOleInPlaceSite::OnUIActivate
//
//---------------------------------------------------------------
HRESULT CDocument::SetUIActiveElement(CElement* pElemNext)
{
    HRESULT hr = S_OK;
    CElement* pElemPrev = _pElemUIActive;

    // Bail out if we are deactivating from Inplace or UI Active.
    if(_pInPlace->_fDeactivating)
    {
        goto Cleanup;
    }

    Assert(!pElemNext || pElemNext->HasLayout() || pElemNext->Tag()==ETAG_ROOT || pElemNext->Tag()==ETAG_DEFAULT);
    Assert(!pElemPrev || pElemPrev->HasLayout() || pElemPrev->Tag()==ETAG_ROOT || pElemPrev->Tag()==ETAG_DEFAULT);

    if(pElemNext != pElemPrev)
    {
        _pElemUIActive = pElemNext;

        // Tell the old ui-active guy to remove it's ui.
        if(pElemPrev)
        {
            pElemPrev->YieldUI(pElemNext);

            if(pElemPrev->_fActsLikeButton)
            {
                CNotification nf;
                nf.AmbientPropChange(pElemPrev, (void*)DISPID_AMBIENT_DISPLAYASDEFAULT);
                pElemPrev->_fDefault = FALSE;
                pElemPrev->Notify(&nf);
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::SetCurrentElem
//
//  Synopsis:   Sets the current element - the element that is or will shortly
//              become UI Active. All keyboard messages and commands will
//              be routed to this element.
//
//  Notes:      Note that this function could be called AFTER _pElemCurrent
//              has been removed from the tree.
//
//  Callee:     If SetCurrentElem succeeds, then the callee should do anything
//              appropriate with gaining currency.  The callee must remember
//              that any action performed here must be cleaned up in
//              YieldCurrency.
//
//----------------------------------------------------------------------------
HRESULT CDocument::SetCurrentElem(CElement* pElemNext, long lSubNext, BOOL* pfYieldFailed)
{
    HRESULT             hr              = S_OK;
    CElement*           pElemPrev       = _pElemCurrent;
    BOOL                fFireEvent;
    BOOL                fFireOnblur     = TRUE;
    BOOL                fPrevDetached   = !(pElemPrev && pElemPrev->GetFirstBranch());
    CElement::CLock*    pLockPrev       = NULL;
    CTreeNode::CLock*   pNodeLockPrev   = NULL;
    CElement*           pElemNewDefault = NULL;
    CElement*           pElemOldDefault = NULL;
    CElement*           pElemFireTarget = NULL;
    long                lSubPrev        = _lSubCurrent;
    BOOL                fSameElem       = (pElemPrev == pElemNext);
    BOOL                fDirty          = FALSE;

    Assert(pElemNext);

    if(pfYieldFailed)
    {
        *pfYieldFailed = FALSE;
    }

    if(fSameElem && lSubNext==_lSubCurrent)
    {
        return S_OK;
    }

    if(!pElemNext || !pElemNext->IsInMarkup())
    {
        return S_FALSE;
    }

    // Someone is trying to set currency to pElemNext from its own onfocus handler (#43161)!
    // Break this loop. Note that it is possible to be in pElemNext's onfocus handler
    // even though pElemNext is the current element.
    if(pElemNext->TestLock(CElement::ELEMENTLOCK_FOCUS))
    {
        return S_OK;
    }

    Assert(this == pElemNext->Doc());

    // We would simply assert here and leave it for the caller to ensure that
    // the element is enabled. Most often, the processing needs to stops way
    // before getting here if the element is disabled. Returning quietly here
    // instead of asserting would hide those bugs.
    Assert(pElemNext->IsEnabled());

    // Prevent attempts to delete the sites.
    Assert(pElemNext->GetFirstBranch());

    CLock LockForm(this, FORMLOCK_CURRENT);

    if(!fPrevDetached)
    {
        pLockPrev = new CElement::CLock(pElemPrev, CElement::ELEMENTLOCK_DELETE);
        pNodeLockPrev = new CTreeNode::CLock(pElemPrev->GetFirstBranch());
    }
    CElement::CLock LockNext(pElemNext, CElement::ELEMENTLOCK_DELETE);
    CTreeNode::CLock NodeLockNext(pElemNext->GetFirstBranch());

    // window onblur will be fired only if body was the current site
    // and we are not refreshing
    if(_pElemCurrent==GetPrimaryElementClient() && !_fForceCurrentElem)
    {
        Fire_onblur();
    }

    _pElemNext = pElemNext;

    // inhibit onblur if the element losing focus is not a select or if a frame in
    // a frameset is not the current one and a previously current site in it is
    // going to be current again.
    fFireOnblur = !fPrevDetached &&
        (_pInPlace && (::GetFocus()==_pInPlace->_hwnd) || pElemPrev->Tag()==ETAG_SELECT);

    if(!fPrevDetached && !fSameElem)
    {
        fFireEvent = !pElemPrev->TestLock(CElement::ELEMENTLOCK_UPDATE);
        CElement::CLock LockUpdate(pElemPrev, CElement::ELEMENTLOCK_UPDATE);

        if(fFireEvent)
        {
            hr = pElemPrev->RequestYieldCurrency(_fForceCurrentElem);
            // yield if currency changed to a different or the same element that is
            // going to become current.
            if(FAILED(hr) || _pElemNext!=pElemNext || _pElemCurrent==pElemNext)
            {
                if(pfYieldFailed)
                {
                    *pfYieldFailed = TRUE;
                }

                goto CanNotYield;
            }
        }

        fFireEvent = !pElemPrev->TestLock(CElement::ELEMENTLOCK_CHANGE);
        CElement::CLock LockChange(pElemPrev, CElement::ELEMENTLOCK_CHANGE);

        if(fFireEvent) // BUGBUG: Why check for fFireEvent here?
        {
            hr = pElemPrev->YieldCurrency(pElemNext);
            if(hr)
            {
                if(pfYieldFailed)
                {
                    *pfYieldFailed = TRUE;
                }

                goto Error;
            }

            // bail out if currency changed
            if(_pElemNext != pElemNext)
            {
                goto Error;
            }
        }
    }

    // bail out if the elem to become current is no longer in the tree, due to some event code
    if(!pElemNext->IsInMarkup())
    {
        goto Error;
    }

    _pElemCurrent = pElemNext;
    _lSubCurrent = lSubNext;

    // Set focus to the current element
    _view.SetFocus(_pElemCurrent, _lSubCurrent);

    // Has currency been set in a non-trivial sense?
    if(!_fCurrencySet && _pElemCurrent->Tag()!=ETAG_ROOT && _pElemCurrent->Tag()!=ETAG_DEFAULT)
    {
        _fCurrencySet = TRUE;
        GWKillMethodCall(this, ONCALL_METHOD(CDocument, DeferSetCurrency, DeferSetCurrency), 0);
    }

    // marka BUGBUG. OnPropertyChange is dirtying the documnet
    // which is bad for editing clients (bugs 10161)
    // this will go away for beta2.
    OnPropertyChange(DISPID_CDoc_activeElement, FORMCHNG_NOINVAL);

    // We fire the blur event AFTER we change the current site. This is
    // because if the onBlur event handler throws up a dialog box then
    // focus will go to the current site (which, if we donot change the
    // current site to be the new one, will still be the previous
    // site which has just yielded currency!).
    if(fFireOnblur)
    {
        Assert(pElemPrev);
        pElemFireTarget = pElemPrev->GetFocusBlurFireTarget(lSubPrev);

        Assert(pElemPrev!=_pElemCurrent || lSubPrev!=_lSubCurrent);

        hr = GWPostMethodCall(pElemFireTarget, ONCALL_METHOD(CElement, Fire_onblur, fire_onblur), 0, TRUE, "CElement::Fire_onblur");
        if(hr)
        {
            goto Error;
        }
    }

    if(_pElemCurrent)
    {
        pElemFireTarget = _pElemCurrent->GetFocusBlurFireTarget(_lSubCurrent);

        Assert(pElemPrev!=_pElemCurrent || lSubPrev!=_lSubCurrent);
        hr = GWPostMethodCall(pElemFireTarget, ONCALL_METHOD(CElement, Fire_onfocus, fire_onfocus), 0, TRUE, "CElement::Fire_onfocus");
        if(hr)
        {
            goto Error;
        }
    }

Cleanup:
    // if forcing, always change the current site as asked
    if(_fForceCurrentElem && _pElemCurrent!=pElemNext)
    {
        _pElemCurrent = pElemNext;
        OnPropertyChange(DISPID_CDoc_activeElement, 0);
        hr = S_OK;
    }

    if(pElemNext == _pElemCurrent)
    {
        if(!fSameElem)
        {
            // if the button is already the default or a button
            pElemNewDefault = _pElemCurrent->_fActsLikeButton
                ? _pElemCurrent : _pElemCurrent->FindDefaultElem(TRUE);

            if(pElemNewDefault && !pElemNewDefault->_fDefault)
            {
                pElemNewDefault->SendNotification(NTYPE_AMBIENT_PROP_CHANGE, (void*)DISPID_AMBIENT_DISPLAYASDEFAULT);
                pElemNewDefault->_fDefault = TRUE;
                pElemNewDefault->Invalidate();
            }
        }
    }

    if(!fPrevDetached && (pElemPrev!=_pElemCurrent))
    {
        // if the button is already the default or a button
        pElemOldDefault = pElemPrev->_fActsLikeButton ? pElemPrev : pElemPrev->FindDefaultElem(TRUE);

        if(pElemOldDefault && pElemOldDefault!=pElemNewDefault)
        {
            pElemOldDefault->SendNotification(NTYPE_AMBIENT_PROP_CHANGE, (void*)DISPID_AMBIENT_DISPLAYASDEFAULT);
            pElemOldDefault->_fDefault = FALSE;
            pElemOldDefault->Invalidate();
        }
    }

CanNotYield:
    if(pLockPrev)
    {
        delete pLockPrev;
    }
    if(pNodeLockPrev)
    {
        delete pNodeLockPrev;
    }

    RRETURN1(hr, S_FALSE);

Error:
    hr = E_FAIL;
    goto Cleanup;
}

void STDMETHODCALLTYPE CDocument::DeferSetCurrency(DWORD_PTR dwContext)
{
    BOOL fWaitParseDone = FALSE;

    // If the currency is already set, or we are not yet inplace active 
    // there is nothing to do... 
    if(_fCurrencySet)
    {
        return;
    }

    // If we are in a dialog, or webview hosting scenario,
    // If parsing is done, then we can activate the first tabbable object.
    if(_dwFlagsHostInfo & DOCHOSTUIFLAG_DIALOG)
    {
        // Parsing is complete, we know which element is the current element, 
        // we can set that element to be the active element
        CElement*   pElement    = NULL;
        long        lSubNext    = 0;

        FindNextTabOrder(DIRECTION_FORWARD, NULL, 0, &pElement, &lSubNext);
        if(pElement)
        {
            Assert(pElement->IsTabbable(lSubNext));

            {
                if(S_OK == pElement->BecomeCurrentAndActive(NULL, lSubNext, FALSE))
                {
                    pElement->ScrollIntoView();
                    _fFirstTimeTab = FALSE;
                }
            }
        }
    }

    // if the currency is not yet set, then make the element client the 
    // current element.
    if(!_fCurrencySet)
    {
        CElement* pel = GetPrimaryElementTop();
        pel->BecomeCurrent(0, NULL, NULL, TRUE);

        // If we are waiting for the parsing to be completed, we have to make sure 
        // that we think the currency is not set when we receive the parse done notification.
        // We will activate the first available object when we receive the parse done
        // notification.
        if(fWaitParseDone)
        {
            _fCurrencySet = FALSE;
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::ShowDragContextMenu
//
//  Synopsis:   Shows the drag context menu upon a right button drop.
//
//  Arguments:  ptl           Where to pop the menu up
//              dwAllowed     Allowed actions for drag
//              piSelection   The choice selected by the menu.
//
//  Returns:    If the user selected a choice on the menu, S_OK.
//              If the menu was brought down, S_FALSE, errors otherwise.
//
//-------------------------------------------------------------------------
HRESULT CDocument::ShowDragContextMenu(POINTL ptl, DWORD dwAllowed, int* piSelection, LPTSTR lptszFileType)
{
    HRESULT hr;
    HMENU   hMenu;
    HMENU   hCtxMenu;
    HCURSOR hOldCursor;
    HCURSOR hArrow = LoadCursor(NULL, IDC_ARROW);

    if(hArrow == NULL)
    {
        RRETURN(GetLastError());
    }

    hMenu = NULL/*LoadMenu(GetResourceHInst(), MAKEINTRESOURCE(IDR_DRAG_CONTEXT_MENU)) wlw note*/;
    if(!hMenu)
    {
        RRETURN(GetLastWin32Error());
    }

    hCtxMenu = GetSubMenu(hMenu, 0);

    // Modify the sub menu based on whether particular choices are
    // available.
    if(!(dwAllowed&DROPEFFECT_MOVE) || (_fFromCtrlPalette))
    {
        EnableMenuItem(hCtxMenu, DROPEFFECT_MOVE, MF_BYCOMMAND|MF_GRAYED);
    }
    if(!(dwAllowed & DROPEFFECT_COPY))
    {
        EnableMenuItem(hCtxMenu, DROPEFFECT_COPY, MF_BYCOMMAND|MF_GRAYED);
    }

    // Get the old cursor, and make new cursor into arrow
    hOldCursor = ::SetCursor(hArrow);

    /*hr = FormsTrackPopupMenu(
        hCtxMenu,
        TPM_LEFTALIGN|TPM_RIGHTBUTTON,
        ptl.x+CX_CONTEXTMENUOFFSET,
        ptl.y+CY_CONTEXTMENUOFFSET,
        NULL,
        piSelection); wlw note*/

    // Now set back the old cursor
    ::SetCursor(hOldCursor);

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ActivateDefaultButton, protected
//
//  Synopsis:   Activate the default button on the form, if any
//
//----------------------------------------------------------------------------
HRESULT CDocument::ActivateDefaultButton(LPMSG lpmsg)
{
    HRESULT hr = S_FALSE;
    CElement* pElem;

    //  The container of this form is always given precedence,
    //    unless the currently active control is a "Push" button
    if(_pElemCurrent && _pElemCurrent->_fActsLikeButton)
    {
        pElem = _pElemCurrent;
        hr = S_OK;
    }
    else
    {
        pElem = _pElemCurrent->FindDefaultElem(TRUE);
    }

    if(pElem)
    {
        Assert(pElem->_fDefault);

        // BUGBUG error dialog?
        _fFirstTimeTab = FALSE;
        hr = pElem->BecomeCurrentAndActive(NULL, 0, TRUE);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pElem->DoClick();
    }
    else
    {
        hr = S_FALSE;
    }
Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ActivateCancelButton, protected
//
//  Synopsis:   Activate the cancel button on the form, if any
//
//----------------------------------------------------------------------------
HRESULT CDocument::ActivateCancelButton(LPMSG lpmsg)
{
    HRESULT hr = S_FALSE;
    CElement* pElem = _pElemCurrent;

    if(!pElem)
    {
        goto Cleanup;
    }

    if(!pElem->TestClassFlag(CElement::ELEMENTDESC_CANCEL))
    {
        // find cancel site
        pElem = _pElemCurrent->FindDefaultElem(FALSE, TRUE);
    }
    if(!pElem)
    {
        goto Cleanup;
    }

    _fFirstTimeTab = FALSE;
    hr = pElem->BecomeCurrentAndActive(NULL, 0, TRUE);
    if(!hr)
    {
        hr = pElem->DoClick();
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     Form::ActivateFirstObject
//
//  Synopsis:   Activate the first object in the tab order.
//
//  Arguments:  lpmsg       Message which prompted this rotation, or
//                          NULL if no message is available
//              pSiteStart  Site to start the search at.  If null, start
//                          at beginning of tab order.
//
//----------------------------------------------------------------------------
HRESULT CDocument::ActivateFirstObject(LPMSG lpmsg)
{
    HRESULT     hr              = S_OK;
    CElement*   pElementClient  = GetPrimaryElementTop();
    BOOL        fDeferActivate  = FALSE;

    if(pElementClient == NULL)
    {
        goto Cleanup;
    }

    if((_dwFlagsHostInfo&DOCHOSTUIFLAG_DIALOG) != 0)
    {
        {
            CElement*       pElement    = NULL;
            long            lSubNext    = 0;
            FOCUS_DIRECTION dir         = DIRECTION_FORWARD;

            if(lpmsg && (lpmsg->message==WM_KEYDOWN || lpmsg->message==WM_SYSKEYDOWN))
            {
                if(GetKeyState(VK_SHIFT) & 0x8000)
                {
                    dir = DIRECTION_BACKWARD;
                }
            }

            FindNextTabOrder(dir, NULL, 0, &pElement, &lSubNext);
            if(pElement)
            {
                Assert(pElement->IsTabbable(lSubNext));
                hr = pElement->BecomeCurrentAndActive(NULL, lSubNext, TRUE);
                if(hr)
                {
                    goto Cleanup;
                }

                pElement->ScrollIntoView();
                _fFirstTimeTab = FALSE;                
                goto Cleanup;
            }
        }
    }

    hr = GetPrimaryElementTop()->BecomeCurrentAndActive(NULL, 0, TRUE);
    // Although we have set the currency, if we have not completed the 
    // parsing yet, we were only able to set it to the element client. Once
    // the parsing is completed, it will be set on the proper tab-stop element.
    // We ignore the _fCurrencySet value set by the BecomeCurrentAndActive call.
    if(fDeferActivate)
    {
        _fCurrencySet = FALSE;
    }

    // Fire window onfocus here the first time if not fired before in
    // onloadstatus when it goes to LOADSTATUS_DONE
    Fire_onfocus();

Cleanup:
    RRETURN1(hr, S_FALSE) ;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::SetZOrder
//
//  Synopsis:   Moves the current selection in the logical Z order.
//
//  Arguments:  [zorder]    action
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CDocument::SetZOrder(int zorder)
{
    CLayout* pLayoutCurrent = _pElemCurrent->GetUpdatedLayout();


    if(!pLayoutCurrent)
    {
        goto Cleanup;
    }

    FixZOrder();
    Invalidate();

Cleanup:
    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::FixZOrder
//
//  Synopsis:   Inserts the given site's window at the proper place in
//              physical Z order, given the site's position in the logical
//              Z order as well as its current OLE state
//
//-------------------------------------------------------------------------
void CDocument::FixZOrder()
{
    _view.SetFlag(CView::VF_DIRTYZORDER);
}

//+------------------------------------------------------------------------
//
//  Member:     FormOverlayWndProc
//
//  Synopsis:   Transparent window for use during drag-drop.
//
//-------------------------------------------------------------------------
static LRESULT CALLBACK FormOverlayWndProc(HWND hWnd, UINT wm, WPARAM wParam, LPARAM lParam)
{
    switch(wm)
    {
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            BeginPaint(hWnd, &ps);
            EndPaint(hWnd, &ps);
        }
        return 0;

    case WM_ERASEBKGND:
        return TRUE;

    default:
        return DefWindowProc(hWnd, wm, wParam, lParam);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::CreateOverlayWindow
//
//  Synopsis:   Creates a transparent window for use during drag-drop.
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HWND CDocument::CreateOverlayWindow(HWND hwndCover)
{
    Assert(_pInPlace->_hwnd);
    HWND hwnd;
    TCHAR* pszBuf;

    if(!s_atomFormOverlayWndClass)
    {
        HRESULT     hr;

        hr = RegisterWindowClass(
            _T("Overlay"),
            FormOverlayWndProc,
            CS_DBLCLKS,
            NULL, NULL,
            &s_atomFormOverlayWndClass);
        if(hr)
        {
            return NULL;
        }
    }

    pszBuf = (TCHAR*)(DWORD_PTR)s_atomFormOverlayWndClass;


    // BUGBUG (garybu) Overlay window should be same size as hwndCover and
    // just above it in the zorder.
    hwnd = CreateWindowEx(
        WS_EX_TRANSPARENT,
        pszBuf,
        NULL,
        WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN|WS_CLIPSIBLINGS,
        0, 0,
        SHRT_MAX/2, SHRT_MAX/2,
        _pInPlace->_hwnd,
        0,
        _afxGlobalData._hInstThisModule,
        this);
    if(!hwnd)
    {
        return NULL;
    }

    SetWindowPos(
        hwnd,
        HWND_TOP,
        0, 0, 0, 0,
        SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);

    return hwnd;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::SetMouseMoveTimer
//
//  Synopsis:   Setup timer to simulate mouse move events.
//
//  Arguments:  uTimeOut    Timer frequency in milliseconds.
//                          If zero, kill the existing timer.
//
//----------------------------------------------------------------------------
void CDocument::SetMouseMoveTimer(UINT uTimeOut)
{
    KillTimer(_pInPlace->_hwnd, 1);
    if(uTimeOut)
    {
        SetTimer(_pInPlace->_hwnd, 1, uTimeOut, 0);
    }
}

//+---------------------------------------------------------------------------
//
//  member: GetBodyElement
//
//  synopsis : helper for the get_/put_ functions that need the body
//
//-----------------------------------------------------------------------------
HRESULT CDocument::GetBodyElement(IHTMLBodyElement** ppBody)
{
    HRESULT hr = S_OK;
    CElement* pElement;

    if(!ppBody)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *ppBody = NULL;

    pElement = GetPrimaryElementClient();
    if(!pElement)
    {
        hr = S_OK;
        goto Cleanup;
    }

    hr = pElement->QueryInterface(IID_IHTMLBodyElement, (void**)ppBody);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  member: GetBodyElement
//
//  synopsis : helper for the get_/put_ functions that need the CBodyElement
//             returns S_FALSE if body element is not found
//
//-----------------------------------------------------------------------------
HRESULT CDocument::GetBodyElement(CBodyElement** ppBody)
{
    HRESULT hr = S_OK;

    if(ppBody == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *ppBody = NULL;

    if(!GetPrimaryElementClient())
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    if(GetPrimaryElementClient()->Tag() == ETAG_BODY)
    {
        *ppBody = DYNCAST(CBodyElement, GetPrimaryElementClient());
    }
    else
    {
        hr = S_FALSE;
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::PreSetErrorInfo
//
//  Synopsis:   Update the UI whenever the form is returned from.
//
//----------------------------------------------------------------------------
void CDocument::PreSetErrorInfo()
{
    super::PreSetErrorInfo();

    DeferUpdateUI();
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::UpdateFromRegistry, public
//
//  Synopsis:   Load configuration information from the registry. Needs to be
//              called after we have our client site so we can do a
//              QueryService.
//
//  Arguments:  flags - REGUPDATE_REFRESH - read even if we find a cache entry
//                    - REGUPDATE_KEEPLOCALSTATE - true if we do not want to
//                                                 override local state
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
struct SAVEDSETTINGS
{
    OLE_COLOR colorBack;
    OLE_COLOR colorText;
    OLE_COLOR colorAnchor;
    OLE_COLOR colorAnchorVisited;
    OLE_COLOR colorAnchorHovered;
    LONG      latmFixedFontFace;
    LONG      latmPropFontFace;
    LONG      nAnchorUnderline;
    SHORT     sBaselineFontDefault;
    BYTE      fAlwaysUseMyColors;
    BYTE      fAlwaysUseMyFontSize;
    BYTE      fAlwaysUseMyFontFace;
    BYTE      fUseMyStylesheet;
    BYTE      bCharSet;
};

static void SaveSettings(OPTIONSETTINGS* pOptionSettings,
                         CODEPAGESETTINGS* pCodepageSettings, SAVEDSETTINGS* pSavedSettings)
{
    memset(pSavedSettings, 0, sizeof(SAVEDSETTINGS));
    pSavedSettings->colorBack             = pOptionSettings->colorBack;
    pSavedSettings->colorText             = pOptionSettings->colorText;
    pSavedSettings->colorAnchor           = pOptionSettings->colorAnchor;
    pSavedSettings->colorAnchorVisited    = pOptionSettings->colorAnchorVisited;
    pSavedSettings->colorAnchorHovered    = pOptionSettings->colorAnchorHovered;
    pSavedSettings->nAnchorUnderline      = pOptionSettings->nAnchorUnderline;
    pSavedSettings->fAlwaysUseMyColors    = pOptionSettings->fAlwaysUseMyColors;
    pSavedSettings->fAlwaysUseMyFontSize  = pOptionSettings->fAlwaysUseMyFontSize;
    pSavedSettings->fAlwaysUseMyFontFace  = pOptionSettings->fAlwaysUseMyFontFace;
    pSavedSettings->fUseMyStylesheet      = pOptionSettings->fUseMyStylesheet;
    pSavedSettings->bCharSet              = pCodepageSettings->bCharSet;
    pSavedSettings->latmFixedFontFace     = pCodepageSettings->latmFixedFontFace;
    pSavedSettings->latmPropFontFace      = pCodepageSettings->latmPropFontFace;
    pSavedSettings->sBaselineFontDefault  = pCodepageSettings->sBaselineFontDefault;
}

HRESULT CDocument::UpdateFromRegistry(DWORD dwFlags, BOOL* pfNeedLayout)
{
    CODEPAGE        codepage;
    BOOL            fFirstTime = !_pOptionSettings;
    SAVEDSETTINGS   savedSettings1;
    SAVEDSETTINGS   savedSettings2;

    // Use the cached values unless we are forced to re-read
    if(!fFirstTime && !(dwFlags&REGUPDATE_REFRESH))
    {
        return S_OK;
    }

    if(pfNeedLayout)
    {
        if(!_pOptionSettings)
        {
            // We assume that in this context, the caller is interested
            // only in whether or not we want to relayout, and not to
            // force us to read in registry settings.
            *pfNeedLayout = FALSE;
            return S_OK;
        }
        SaveSettings(_pOptionSettings, _pCodepageSettings, &savedSettings1);
    }

    // First read in the standard option settings
    ReadOptionSettingsFromRegistry(dwFlags);

    {
        _pOptionSettings->fShowImages      = TRUE;
        _pOptionSettings->fShowVideos      = TRUE;
        _pOptionSettings->fPlaySounds      = TRUE;
        _pOptionSettings->fPlayAnimations  = TRUE;
        _pOptionSettings->fSmartDithering  = TRUE;
    }

    // If we are getting settings for the first time, use the default
    //  codepage.  Otherwise, read based on our current setting.
    codepage = _pCodepageSettings ? _codepage : _pOptionSettings->codepageDefault;
    ReadCodepageSettingsFromRegistry(codepage, WindowsCodePageFromCodePage(codepage), dwFlags);

    if(_pCodepageSettings)
    {
        // Set the baseline font only the first time we read from the registry
        _sBaselineFont = _pCodepageSettings->sBaselineFontDefault;
    }

    // Request the accept language header info from shlwapi.
    {
        TCHAR achLang[256];
        DWORD cchLang = ARRAYSIZE(achLang);
        _pOptionSettings->fHaveAcceptLanguage = (GetAcceptLanguages(achLang, &cchLang)==S_OK)
            && (_pOptionSettings->cstrLang.Set(achLang, cchLang)==S_OK);
    }

    if(pfNeedLayout)
    {
        SaveSettings(_pOptionSettings, _pCodepageSettings, &savedSettings2);
        *pfNeedLayout = !!memcmp(&savedSettings1, &savedSettings2, sizeof(SAVEDSETTINGS));
    }

    _view.RefreshSettings();

    return S_OK;
}

//+-------------------------------------------------------------------------
//
//  Method:     CDocument::RtfConverterEnabled
//
//  Synopsis:   TRUE if this rtf conversions are enabled, FALSE otherwise.
//
//--------------------------------------------------------------------------
BOOL CDocument::RtfConverterEnabled()
{
    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ReadCodepageSettingsFromRegistry, public
//
//  Synopsis:   Read settings for a particular codepage from the registry.
//
//  Arguments:  cp - the codepage to read
//              dwFlags - See UpdateFromRegistry
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CDocument::ReadCodepageSettingsFromRegistry(CODEPAGE cp, UINT uiFamilyCodePage, DWORD dwFlags)
{
    HRESULT hr = S_OK;
    SCRIPT_ID sid = RegistryAppropriateSidFromSid(DefaultSidForCodePage(uiFamilyCodePage));

    Assert(uiFamilyCodePage!=CP_UNDEFINED && uiFamilyCodePage != CP_ACP);

    hr = EnsureCodepageSettings(uiFamilyCodePage);
    if(hr)
    {
        goto Cleanup;
    }

    _pOptionSettings->ReadCodepageSettingsFromRegistry(_pCodepageSettings, dwFlags, sid);

    // Set the codepage on the doc to the actual codepage requested
    _codepage = cp;
    _codepageFamily = uiFamilyCodePage;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::EnsureOptionSettings, public
//
//  Synopsis:   Ensures that this CDoc is pointing to a valid OPTIONSETTINGS
//              object. Should only be called by ReadOptionSettingsFromRegistry.
//
//  Arguments:  (none)
//
//----------------------------------------------------------------------------
HRESULT CDocument::EnsureOptionSettings()
{
    HRESULT             hr = S_OK;
    int                 c;
    OPTIONSETTINGS*     pOS;
    OPTIONSETTINGS**    ppOS;
    TCHAR*              pstr=NULL;
    BOOL                fUseCodePageBasedFontLinking;

    static TCHAR pszDefaultKey[] = _T("Software\\Microsoft\\Internet Explorer");

    if(_pOptionSettings)
    {
        return S_OK;
    }

    if(!pstr)
    {
        pstr = pszDefaultKey;
    }

    // On Unix, we create new option setting for each new CDoc, because all IEwindows
    // run on the same thread.
    for(c=TLS(optionSettingsInfo.pcache).Size(),ppOS=TLS(optionSettingsInfo.pcache);
        c>0; c--,ppOS++)
    {
        if(!StrCmpC((*ppOS)->achKeyPath, pstr))
        {
            _pOptionSettings = *ppOS;
            goto Cleanup;
        }
    }
    // We only make it here if we didn't find an existing entry.

    // OPTIONSETTINGS has one character already in it, which accounts for a
    // NULL terminator.
    pOS = new (_tcslen(pstr)*sizeof(TCHAR)) OPTIONSETTINGS;
    if(!pOS)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    // New clients, such as OE5, will set the UIHost flag.  For compatibility with OE4, however,
    // we need to employ some trickery to get the old font regkey values.
    fUseCodePageBasedFontLinking = 0 != (_dwFlagsHostInfo & DOCHOSTUIFLAG_CODEPAGELINKEDFONTS);
    if(!fUseCodePageBasedFontLinking)
    {
        fUseCodePageBasedFontLinking = FALSE;
    }

    MemSetName((pOS, "OPTIONSETTINGS object, index %d", TLS(optionSettingsInfo.pcache).Size()));

    pOS->Init(pstr, fUseCodePageBasedFontLinking);
    pOS->SetDefaults( );

    hr = TLS(optionSettingsInfo.pcache).Append(pOS);
    if(hr)
    {
        goto Cleanup;
    }

    _pOptionSettings = pOS;

    ClearDefaultCharFormat();

Cleanup:
    if(pstr != pszDefaultKey)
    {
        CoTaskMemFree(pstr);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::EnsureCodepageSettings, public
//
//  Synopsis:   Ensures that this CDoc is pointing to a valid CODEPAGESETTINGS
//              object. Should only be called by ReadCodepageSettingsFromRegistry.
//
//  Arguments:  uiFamilyCodePage - the family to check for
//
//----------------------------------------------------------------------------
HRESULT CDocument::EnsureCodepageSettings(UINT uiFamilyCodePage)
{
    HRESULT hr = S_OK;
    int n;
    CODEPAGESETTINGS **ppCS, *pCS;

    // Make sure we have a valid _pOptionSettings object
    Assert(_pOptionSettings);

    // The first step is to look up the entry in the codepage cache
    for(n=_pOptionSettings->aryCodepageSettingsCache.Size(),
        ppCS=_pOptionSettings->aryCodepageSettingsCache;
        n>0; n--,ppCS++)
    {
        if((*ppCS)->uiFamilyCodePage == uiFamilyCodePage)
        {
            _pCodepageSettings = *ppCS;
            goto Cleanup;
        }
    }

    // We're out of luck, need to read in the codepage setting from the registry
    pCS = (CODEPAGESETTINGS*)MemAlloc(sizeof(CODEPAGESETTINGS));
    if(!pCS)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    MemSetName((pCS, "CODEPAGESETTINGS"));

    pCS->Init( );
    pCS->SetDefaults(uiFamilyCodePage, _pOptionSettings->sBaselineFontDefault);

    hr = _pOptionSettings->aryCodepageSettingsCache.Append(pCS);
    if(hr)
    {
        goto Cleanup;
    }

    _pCodepageSettings = pCS;

    ClearDefaultCharFormat();

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::WindowBlurHelper()
{
    HRESULT hr = S_OK;
    HWND hwndParent = NULL, hwnd;
    BOOL fOldInhibitFocusFiring;

    if(!_pInPlace)
    {
        goto Cleanup;
    }

    hwnd = _pInPlace->_hwnd;

    while(hwnd)
    {
        hwndParent = hwnd;
        hwnd = GetParent(hwnd);
    }

    if(hwndParent)
    {
        ::SetWindowPos(hwndParent, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
    }

    hr = PrimaryRoot()->BecomeCurrent(0);
    if(hr)
    {
        goto Cleanup;
    }

    // remove focus from current frame window with the focus
    fOldInhibitFocusFiring = _fInhibitFocusFiring;
    _fInhibitFocusFiring = TRUE;
    ::SetFocus(NULL);
    _fInhibitFocusFiring = fOldInhibitFocusFiring;

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::WindowFocusHelper()
{
    HRESULT hr = S_OK;
    HWND hwndParent = NULL, hwnd;

    if(!_pInPlace)
    {
        goto Cleanup;
    }

    hwnd = _pInPlace->_hwnd;

    while(hwnd)
    {
        hwndParent = hwnd;
        hwnd = GetParent(hwnd);
    }

    if(hwndParent)
    {
        // restore window first if minimized.
        if(IsIconic(hwndParent))
        {
            ShowWindow(hwndParent, SW_RESTORE);
        }
        else
        {
            ::SetWindowPos(hwndParent, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE);

            // SetWindowPos does not seem to bring the window to top of the z-order
            // when window.focus() is called if the browser app is not in the foreground,
            // unless the z-order was changed with the window.blur() method in a prior call.
            ::SetForegroundWindow(hwndParent);
        }
    }

    hr = GetPrimaryElementTop()->BecomeCurrentAndActive(NULL, 0, TRUE);
    if(hr)
    {
        goto Cleanup;
    }

    // Fire the window onfocus event here if it wasn't previously fired in
    // BecomeCurrent() due to our inplace window not previously having the focus
    Fire_onfocus();

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetBaseUrl
//
//              Get the base URL for the document.
//
//              If supplied with an element, gets the base URL in effect
//              at that element, based on the position of <BASE> tags.
//              Note that this is a pointer to an internal string, it can't
//              be modified. If you need to modify it, make a copy first.
//
//----------------------------------------------------------------------------
HRESULT CDocument::GetBaseUrl(
         TCHAR**    ppchHref,
         CElement*  pElementContext,
         BOOL*      pfDefault,
         TCHAR*     pszAlternateDocUrl)
{
    TCHAR*      pchHref = NULL;
    BOOL        fDefault;
    HRESULT     hr=S_OK;

    Assert(!!_cstrUrl);

    if(pchHref == NULL)
    {
        *ppchHref = pszAlternateDocUrl ? pszAlternateDocUrl : _cstrUrl;
        fDefault = TRUE;
    }
    else
    {
        *ppchHref = pchHref;
        fDefault = FALSE;
    }

    if(pfDefault)
    {
        *pfDefault = fDefault;
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ExpandUrl
//
//----------------------------------------------------------------------------
HRESULT CDocument::ExpandUrl(
        LPCTSTR     pchRel, 
        LONG        dwUrlSize,
        TCHAR*      pchUrlOut,
        CElement*   pElementContext,
        DWORD       dwFlags,
        TCHAR*      pszAlternateDocUrl)
{
    HRESULT hr = S_OK;
    DWORD cchBuf;
    BOOL fDefault;
    BOOL fCombine;
    DWORD dwSize;
    TCHAR* pchBaseUrl = NULL;

    AssertSz(dwUrlSize==pdlUrlLen, "Wrong size URL buffer!!");

    if(dwFlags == 0xFFFFFFFF)
    {
        dwFlags = URL_ESCAPE_SPACES_ONLY | URL_BROWSER_MODE;
    }

    if(pchRel == NULL) // note: NULL is different from "", which is more similar to "."
    {
        *pchUrlOut = _T('\0');
        goto Cleanup;
    }

    hr = GetBaseUrl(&pchBaseUrl, pElementContext, &fDefault, pszAlternateDocUrl);
    if(hr)
    {
        goto Cleanup;
    }

    hr = CoInternetCombineUrl(pchBaseUrl, pchRel, dwFlags, pchUrlOut, pdlUrlLen, &cchBuf, 0);

    if(hr)
    {
        goto Cleanup;
    }

    if(!fDefault
        && S_OK==CoInternetQueryInfo(
        pszAlternateDocUrl?pszAlternateDocUrl:_cstrUrl,
        QUERY_RECOMBINE, 0, &fCombine, sizeof(BOOL), &dwSize, 0)
        && fCombine)
    {
        TCHAR achBuf2[pdlUrlLen];
        DWORD cchBuf2;

        hr = CoInternetCombineUrl(pszAlternateDocUrl?pszAlternateDocUrl:_cstrUrl,
            pchUrlOut, dwFlags, achBuf2, pdlUrlLen, &cchBuf2, 0);

        if(hr)
        {
            goto Cleanup;
        }

        _tcsncpy (pchUrlOut, achBuf2, dwUrlSize);
        pchUrlOut[dwUrlSize-1] = _T('\0'); // If there wasn't room for the URL.
    }

Cleanup:
    RRETURN(hr);
}

BOOL IsSpecialUrl(TCHAR* pchURL)
{
    UINT uProt;
    uProt = GetUrlScheme(pchURL);
    return (URL_SCHEME_JAVASCRIPT==uProt || URL_SCHEME_VBSCRIPT==uProt || URL_SCHEME_ABOUT==uProt);
}

HRESULT WrapSpecialUrl(TCHAR* pchURL, CString* pcstrExpandedUrl, CString& cstrDocUrl, BOOL fNonPrivate, BOOL fIgnoreUrlScheme)
{
    HRESULT hr = S_OK;
    CString cstrSafeUrl;
    TCHAR   *pch, *pchPrev;
    TCHAR   achUrl[pdlUrlLen];
    DWORD   dwSize;

    if(IsSpecialUrl(pchURL) || fIgnoreUrlScheme)
    {
        // If this is javascript:, vbscript: or about:, append the
        // url of this document so that on the other side we can
        // decide whether or not to allow script execution.
        //

        // QFE 2735 (Georgi XDomain): [alanau]
        //
        // If the special URL contains an %00 sequence, then it will be converted to a Null char when
        // encoded.  This will effectively truncate the Security ID.  For now, simply disallow this
        // sequence, and display a "Permission Denied" script error.
        if(_tcsstr(pchURL, _T("%00")))
        {
            hr = E_ACCESSDENIED;
            goto Cleanup;
        }

        // Copy the URL so we can munge it.
        cstrSafeUrl.Set(pchURL);

        // someone could put in a string like this:
        //     %2501 OR %252501 OR %25252501
        // which, depending on the number of decoding steps, will bypass security
        // so, just keep decoding while there are %s and the string is getting shorter
        UINT uPreviousLen = 0;
        while((uPreviousLen!=cstrSafeUrl.Length()) && _tcschr(cstrSafeUrl, _T('%')))
        {
            uPreviousLen = cstrSafeUrl.Length();
            int nNumPercents;
            int nNumPrevPercents = 0;

            // Reduce the URL
            for(; ;)
            {
                // Count the % signs.
                nNumPercents = 0;

                pch = cstrSafeUrl;
                while(pch = _tcschr(pch, _T('%')))
                {
                    pch++;
                    nNumPercents++;
                }

                // If the number of % signs has changed, we've reduced the URL one iteration.
                if(nNumPercents != nNumPrevPercents)
                {
                    // Encode the URL 
                    hr = CoInternetParseUrl(
                        cstrSafeUrl, 
                        PARSE_ENCODE, 
                        0, 
                        achUrl, 
                        ARRAYSIZE(achUrl), 
                        &dwSize,
                        0);

                    cstrSafeUrl.Set(achUrl);

                    nNumPrevPercents = nNumPercents;
                }
                else
                {
                    // The URL is fully reduced.  Break out of loop.
                    break;
                }
            }
        }

        // Now scan for '\1' characters.
        if(_tcschr(cstrSafeUrl, _T('\1')))
        {
            // If there are '\1' characters, we can't guarantee the safety.  Put up "Permission Denied".
            //
            hr = E_ACCESSDENIED;
            goto Cleanup;
        }

        hr = pcstrExpandedUrl->Set(cstrSafeUrl);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pcstrExpandedUrl->Append(fNonPrivate ? _T("\1\1") : _T("\1"));
        if(hr)
        {
            goto Cleanup;
        }


        // Now copy the cstrDocUrl
        cstrSafeUrl.Set(cstrDocUrl);

        // Scan the URL to ensure it appears un-spoofed.
        //
        // There may legitimately be multiple '\1' characters in the URL.  However, each one, except the last one
        // should be followed by a "special" URL (javascript:, vbscript: or about:).
        pchPrev = cstrSafeUrl;
        pch = _tcschr(cstrSafeUrl, _T('\1'));
        while(pch)
        {
            pch++;                              // Bump past security marker
            if(*pch == _T('\1'))                // (Posibly two security markers)
            {
                pch++;
            }

            if(!IsSpecialUrl(pchPrev))          // If URL is not special
            {
                hr = E_ACCESSDENIED;            // then it's spoofed.
                goto Cleanup;
            }
            pchPrev = pch;
            pch = _tcschr(pch, _T('\1'));
        }

        // Look for escaped %01 strings in the Security Context.
        pch = cstrSafeUrl;
        while(pch = _tcsstr(pch, _T("%01")))
        {
            pch[2] = _T('2');  // Just change the %01 to %02.
            pch += 3;          // and skip over
        }

        hr = pcstrExpandedUrl->Append(cstrSafeUrl);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else
    {
        hr = pcstrExpandedUrl->Set(pchURL);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::IsVisitedHyperlink
//
//  Synopsis:   returns TRUE if the given url is in the Hisitory
//
//              Currently ignores #location information
//
//----------------------------------------------------------------------------
BOOL CDocument::IsVisitedHyperlink(LPCTSTR pchURL, CElement* pElementContext)
{
    HRESULT     hr              = S_OK;
    TCHAR       cBuf[pdlUrlLen];
    TCHAR*      pchExpandedUrl  = cBuf;
    BOOL        result          = FALSE;
    TCHAR*      pch;

    // fully resolve URL
    hr = ExpandUrl(pchURL, ARRAYSIZE(cBuf), pchExpandedUrl, pElementContext);
    if(hr)
    {
        goto Cleanup;
    }

    pch = const_cast<TCHAR*>(UrlGetLocation(pchExpandedUrl));

    Assert(!pchExpandedUrl[0] || pchExpandedUrl[0] > _T(' '));

Cleanup:
    return result;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetProgSink
//
//  Synopsis:   Returns progsink interface of the current primary markup
//
//----------------------------------------------------------------------------
IProgSink* CDocument::GetProgSink()
{
    return (PrimaryMarkup() ? PrimaryMarkup()->GetProgSink() : NULL);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetProgSinkC
//
//  Synopsis:   Returns progsink of the current primary markup
//
//----------------------------------------------------------------------------
CProgSink* CDocument::GetProgSinkC()
{
    return (PrimaryMarkup() ? PrimaryMarkup()->GetProgSinkC() : NULL);
}

HRESULT CDocument::WriteDocHeader(CStreamWriteBuff* pStm)
{
    // Write "<!DOCTYPE... >" stuff
    HRESULT hr = S_OK;
    DWORD dwFlagSave = pStm->ClearFlags(WBF_ENTITYREF);

    // Do not write out the header in plaintext mode.
    if(pStm->TestFlag(WBF_SAVE_PLAINTEXT))
    {
        goto Cleanup;
    }

    pStm->SetFlags(WBF_NO_WRAP);

    hr = pStm->Write(_T("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">"));
    if(hr)
    {
        goto Cleanup;
    }

    hr = pStm->NewLine();
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    pStm->RestoreFlags(dwFlagSave);

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::FireTimeoutCode
//
//  Synopsis:   save the code associated with a TimerID
//
//----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::FireTimeOut(UINT uTimerID)
{
    TIMEOUTEVENTINFO*   pTimeout = NULL;
    LONG                id;
    HRESULT             hr = S_OK;

    if(_fSuspendTimeout)
    {
        goto Cleanup;
    }

    // Check the list and if ther are timers with target time less then current
    //      execute them and remove from the list. Only events occured earlier then
    //      timer uTimerID will be retrieved
    // BUGBUG: check with Nav compat to see if Nav wraps to prevent clears during script exec.
    _cProcessingTimeout++;
    Trace1("Got a timeout. Looking for a match to id %d\n", uTimerID);

    while(_TimeoutEvents.GetFirstTimeoutEvent(uTimerID, &pTimeout) == S_OK)
    {
        // Now execute the given script
        hr = ExecuteTimeoutScript(pTimeout);
        if(0==pTimeout->_dwInterval || hr)
        {
            // setTimeout (or something wrong with script): delete the timer
            delete pTimeout;
        }
        else
        {
            // setInterval: put timeout back in queue with next time to fire
            _TimeoutEvents.AddPendingTimeout(pTimeout);
        }
    }
    _cProcessingTimeout--;

    // deal with any clearTimeouts (clearIntervals) that may have occurred as
    // a result of processing the scripts.
    while(_TimeoutEvents.GetPendingClear(&id))
    {
        if(!_TimeoutEvents.ClearPendingTimeout((UINT)id))
        {
            ClearTimeout(id);
        }
    }

    // we cleanup here because clearTimeout might have been called from setTimeout code
    // before an error occurred which we want to get rid of above (nav compat)
    if(hr)
    {
        goto Cleanup;
    }

    // Requeue pending timeouts (from setInterval)
    while(_TimeoutEvents.GetPendingTimeout(&pTimeout))
    {
        pTimeout->_dwTargetTime = (DWORD)GetTargetTime(pTimeout->_dwInterval);
        hr = _TimeoutEvents.InsertIntoTimeoutList(pTimeout, NULL, FALSE);
        if(hr)
        {
            ClearTimeout(pTimeout->_uTimerID);
            goto Cleanup;
        }

        hr = FormsSetTimer(this, ONTICK_METHOD(CDocument, FireTimeOut, firetimeout),
            pTimeout->_uTimerID, pTimeout->_dwInterval);
        if(hr)
        {
            ClearTimeout(pTimeout->_uTimerID);
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ClearTimeout
//
//  Synopsis:   Clears a previously setTimeout and setInterval
//
//----------------------------------------------------------------------------
HRESULT CDocument::ClearTimeout(LONG lTimerID)
{
    HRESULT hr = S_OK;
    TIMEOUTEVENTINFO* pCurTimeout;

    if(_cProcessingTimeout)
    {
        _TimeoutEvents.AddPendingClear(lTimerID);
    }
    else
    {
        // Get the timeout struct with given ID and remove it from the list
        hr = _TimeoutEvents.GetTimeout((DWORD)lTimerID, &pCurTimeout);
        if(hr == S_FALSE)
        {
            // Netscape just ignores the invalid arg silently - so do we.
            hr = S_OK;
            goto Cleanup;
        }

        if(pCurTimeout != NULL)
        {
            Verify(FormsKillTimer(this, pCurTimeout->_uTimerID) == S_OK);
            delete pCurTimeout;
        }
    }
Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::PrintBackgroundColorOrBitmap
//
//  Synopsis:   Returns cached flag indicating whether printing of backgrounds
//              is on or off.
//
//----------------------------------------------------------------------------
BOOL CDocument::PaintBackground()
{
    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetHDC
//
//  Synopsis:   Returns cached flag indicating whether printing of backgrounds
//              is on or off.
//
//----------------------------------------------------------------------------
HDC CDocument::GetHDC()
{
    return TLS(hdcDesktop);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::DisabledTilePaint()
//
//  Synopsis:   return true if tiled painting disabled
//
//----------------------------------------------------------------------------
BOOL CDocument::TiledPaintDisabled()
{
    return _fDisableTiledPaint;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::FindNextTabOrder
//
//  Synopsis:   Given an element and a subdivision within it, find the
//              next tabbable item.
//
//----------------------------------------------------------------------------
BOOL CDocument::FindNextTabOrder(
             FOCUS_DIRECTION dir,
             CElement* pCurrent,
             long lSubCurrent,
             CElement** ppNext,
             long* plSubNext)
{
    BOOL        fFound      = FALSE;
    CElement*   pElemTemp   = pCurrent;
    long        lSubTemp    = lSubCurrent;

    for(;;)
    {
        *ppNext    = NULL;
        *plSubNext = 0;

        // When Shift+Tabbing, start the search with those not in the Focus array (because
        // they are at the bottom of the tab order)
        if(!pCurrent && DIRECTION_BACKWARD==dir)
        {
            break;
        }

        fFound = SearchFocusArray(dir, pElemTemp, lSubTemp, ppNext, plSubNext);

        if(!fFound || !*ppNext || (*ppNext)->IsTabbable(*plSubNext))
        {
            break;
        }

        Assert(pElemTemp!=*ppNext || lSubTemp!=*plSubNext);

        pElemTemp = *ppNext;
        lSubTemp = *plSubNext;
    }

    if(!*ppNext)
    {
        if(fFound)
        {
            if(DIRECTION_BACKWARD == dir)
            {
                return TRUE;
            }
            else
            {
                // Means pCurrent was in focus array, but next tabbable item is not
                // Just retrieve the first tabbable item from the tree.
                pCurrent = NULL;
                lSubCurrent = 0;
            }
        }

        fFound = SearchFocusTree(dir, pCurrent, lSubCurrent, ppNext, plSubNext);

        // If element was found, but next item is not in
        // the tree and we're going backwards, return the last element in
        // the focus array.
        if(!*ppNext && (fFound||!pCurrent) &&
            dir==DIRECTION_BACKWARD && _aryFocusItems.Size()>0)
        {
            long lLast = _aryFocusItems.Size() - 1;

            while(lLast >= 0)
            {
                pElemTemp = _aryFocusItems[lLast].pElement;
                lSubTemp = _aryFocusItems[lLast].lSubDivision;
                if(pElemTemp->IsTabbable(lSubTemp))
                {
                    *ppNext = pElemTemp;
                    *plSubNext = lSubTemp;
                    fFound = TRUE;
                    break;
                }
                lLast--;
            }
        }
    }

    return fFound;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::OnElementEnter
//
//  Synopsis:   Update the focus collection upon an element entering the tree
//
//-------------------------------------------------------------------------
HRESULT CDocument::OnElementEnter(CElement* pElement)
{
    HRESULT hr = S_OK;

    // Insert pElement in _aryAccessKeyItems if accessKey is defined.
    LPCTSTR lpstrAccKey = pElement->GetAAaccessKey();
    if(lpstrAccKey && lpstrAccKey[0])
    {
        hr = InsertAccessKeyItem(pElement);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::OnElementExit
//
//  Synopsis:   Remove all references the doc may have to an element
//              that is exiting the tree
//
//-------------------------------------------------------------------------
HRESULT CDocument::OnElementExit(CElement* pElement, DWORD dwExitFlags )
{
    long i;

    if(_fVID && !_fOnControlInfoChangedPosted)
    {
        _fOnControlInfoChangedPosted = TRUE;
        GWPostMethodCall(this, ONCALL_METHOD(CDocument, OnControlInfoChanged, oncontrolinfochanged), 0, FALSE, "CDocument::OnControlInfoChanged");
    }

    if(_pMenuObject == pElement)
    {
        _pMenuObject = NULL;
    }

    // Clear the edit context, if it is going away
    if(_pElemEditContext && (_pElemEditContext==pElement
        || (pElement->HasSlaveMarkupPtr()
        && _pElemEditContext==pElement->GetSlaveMarkupPtr()->FirstElement())))
    {
        _pElemEditContext = NULL;
    }    

    // marka - find the Element that will be current next - via adjusting the edit context
    // fire on Before Active Elemnet Change. If it's ok - do a SetEditContext on that elemnet,
    // and set _pElemCurrent to it. 
    //
    // Otherwise if FireOnBeforeActiveElement fails - 
    // we make the Body the current element, and call SetEditCOntext on that.
    if(_pElemCurrent == pElement)
    {
        if(dwExitFlags & EXITTREE_DESTROY)
        {
            // BUGBUG: (jbeda) is this right?  What else do we have to do on markup destroy?
            _pElemCurrent = PrimaryRoot();
        }
        else
        {
            CElement* pElementEditContext = pElement;
            CElement* pCurElement = NULL;

            CTreeNode* pNodeSiteParent = pElement->GetFirstBranch()->GetUpdatedParentLayoutNode();
            if(pNodeSiteParent)
            {
                pCurElement = pNodeSiteParent->Element();
            }

            if(pCurElement)
            {
                GetEditContext(pCurElement, &pElementEditContext, NULL, NULL);
            }
            else
            {
                pElementEditContext = PrimaryRoot();
            }

            // if it's the top element, don't defer to itself
            _pElemCurrent = pElementEditContext;

            if(_pElemCurrent == pElement)
            {
                _pElemCurrent = PrimaryRoot();
            }

            if(_pElemCurrent->IsEditable(FALSE) && _pElemCurrent->_etag!=ETAG_ROOT)
            {
                // An editable element has just become current.
                // We tell the mshtmled.dll about this change in editing "context"
                if(_pCaret)
                {
                    _pCaret->Show(FALSE);
                }
                SetEditContext((_pElemCurrent->HasSlaveMarkupPtr()?
                    _pElemCurrent->GetSlaveMarkupPtr()->FirstElement():
                _pElemCurrent), TRUE, FALSE, TRUE);
            }

            OnPropertyChange(DISPID_CDoc_activeElement, 0);
        }
    }

    // Make sure pElement site pointer is not cached by the document, which
    // can happen if we are in the middle of a drag-drop operation
    if(_pDragDropTargetInfo)
    {
        if(pElement == _pDragDropTargetInfo->_pElementTarget)
        {
            _pDragDropTargetInfo->_pElementTarget = NULL;
            _pDragDropTargetInfo->_pElementHit = NULL;
            _pDragDropTargetInfo->_pDispScroller = NULL;
        }
    }

    // Release capture if it owns it
    if(GetCaptureObject() == (void*)pElement)
    {
        pElement->TakeCapture(FALSE);
    }

    if(_pNodeLastMouseOver)
    {
        ClearCachedNodeOnElementExit(&_pNodeLastMouseOver, pElement);
    }
    if(_pNodeGotButtonDown)
    {
        ClearCachedNodeOnElementExit(&_pNodeGotButtonDown, pElement);
    }

    // Reset _pElemUIActive if necessary
    if(_pElemUIActive == pElement)
    {
        // BUGBUG (MohanB) Need to call _pElemUIActive->YieldUI() here ?
        _pElemUIActive = NULL;
        if(_pInPlace && !_pInPlace->_fDeactivating && _pElemUIActive!=PrimaryRoot())
        {
            PrimaryRoot()->BecomeUIActive();
        }
    }

    // Remove all traces of pElement from the focus item array
    // and accessKey array.

    // BUGBUG: this could be N^2 on shutdown
    for(i=0; i<_aryFocusItems.Size();)
    {
        if(_aryFocusItems[i].pElement == pElement)
        {
            _aryFocusItems.Delete(i);
        }
        else
        {
            i++;
        }
    }

    for(i=0; i<_aryAccessKeyItems.Size(); i++)
    {
        if(_aryAccessKeyItems[i] == pElement)
        {
            _aryAccessKeyItems.Delete(i);
            break;
        }
    }

    return S_OK;
}

void CDocument::ClearCachedNodeOnElementExit(CTreeNode** ppNodeToTest, CElement* pElement)
{
    Assert(pElement && ppNodeToTest && *ppNodeToTest);

    CTreeNode* pNode = *ppNodeToTest;

    if(SameScope(pNode, pElement))
    {
        *ppNodeToTest = NULL;
        pNode->NodeRelease();
        return;
    }

    if(pElement->HasSlaveMarkupPtr())
    {
        // BUGBUG (MohanB) This will soon be generalized by DBau
        // through a new content markup notification (like
        // (NTYPE_MASTER_EXITVIEW or something)
        CMarkup* pMarkup = pNode->GetMarkup();

        while(pMarkup && pMarkup->Master())
        {
            if(pMarkup->Master() == pElement)
            {
                *ppNodeToTest = NULL;
                pNode->NodeRelease();
                return;
            }
            pMarkup = pMarkup->Master()->GetMarkup();
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::SearchFocusArray
//
//  Synopsis:   Search the focus array for the next focussable item
//
//  Returns:    FALSE if pElemFirst was not found in the focus arrya
//              TRUE and ppElemNext/plSubNext set to NULL if pElemFirst
//                  was found, but the next item was not present.
//              TRUE and ppElemNext/plSubNext set to valid values if
//                  pElemFirst was found and the next item is also present.
//
//-------------------------------------------------------------------------
BOOL CDocument::SearchFocusArray(
             FOCUS_DIRECTION dir,
             CElement* pElemFirst,
             long lSubFirst,
             CElement** ppElemNext,
             long* plSubNext)
{
    int         i;
    BOOL        fFound = FALSE;
    CMarkup*    pMarkup = PrimaryMarkup();

    *ppElemNext = NULL;
    *plSubNext = 0;

    if(_lFocusTreeVersion != pMarkup->GetMarkupTreeVersion())
    {
        CCollectionCache* pCollCache;
        CElement* pElement;

        _lFocusTreeVersion = pMarkup->GetMarkupTreeVersion();
        _aryFocusItems.DeleteAll();

        pMarkup->EnsureCollectionCache(CMarkup::ELEMENT_COLLECTION);
        pCollCache = pMarkup->CollectionCache();
        if(!pCollCache)
        {
            return FALSE;
        }

        for(i=0; i<pCollCache->SizeAry(CMarkup::ELEMENT_COLLECTION); i++)
        {
            pCollCache->GetIntoAry(CMarkup::ELEMENT_COLLECTION, i, &pElement);
            if(pElement)
            {
                InsertFocusArrayItem(pElement);
            }
        }
    }

    if(!pElemFirst)
    {
        if(DIRECTION_FORWARD==dir && _aryFocusItems.Size()>0)
        {
            i = 0;
        }
        else if(DIRECTION_BACKWARD==dir && _aryFocusItems.Size()>0)
        {
            i = _aryFocusItems.Size() - 1;
        }
        else
        {
            return FALSE;
        }
    }
    else
    {
        // Search for pElemFirst and lSubFirst in the array.
        for(i=0; i<_aryFocusItems.Size(); i++)
        {
            if(pElemFirst==_aryFocusItems[i].pElement &&
                lSubFirst==_aryFocusItems[i].lSubDivision)
            {
                fFound = TRUE;
                break;
            }
        }

        // If pElemFirst is not in the array, just return FALSE.  This
        // will cause SearchFocusTree to get called.
        if(!fFound)
        {
            return FALSE;
        }

        // If pElemFirst is the first/last element in the array
        // return TRUE.  Otherwise return the next tabbable element.
        if(DIRECTION_FORWARD == dir)
        {
            if(i == _aryFocusItems.Size()-1)
            {
                return TRUE;
            }
            i++; // Make i point to the item to return
        }
        else
        {
            if(i == 0)
            {
                return TRUE;
            }
            i--;
        }
    }

    *ppElemNext = _aryFocusItems[i].pElement;
    *plSubNext = _aryFocusItems[i].lSubDivision;
    return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::SearchFocusTree
//
//  Synopsis:   Update the focus collection upon an element exiting the tree
//
//  Returns:    See SearchFocusArray
//
//-------------------------------------------------------------------------
BOOL CDocument::SearchFocusTree(
            FOCUS_DIRECTION dir,
            CElement* pElemFirst,
            long lSubFirst,
            CElement** ppElemNext,
            long* plSubNext)
{
    // Use the all collection for now.  BUGBUG: (anandra) Fix ASAP.
    HRESULT             hr = S_OK;
    long                i;
    long                cElems;
    CElement*           pElement;
    BOOL                fStopOnNextTab = FALSE;
    long                lSubNext;
    int                 iStep;
    BOOL                fFound = FALSE;
    CCollectionCache*   pCollectionCache;

    *ppElemNext = NULL;
    *plSubNext = 0;

    hr = PrimaryMarkup()->EnsureCollectionCache(CMarkup::ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    pCollectionCache = PrimaryMarkup()->CollectionCache();

    cElems = pCollectionCache->SizeAry(CMarkup::ELEMENT_COLLECTION);
    if(!cElems)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    // Search for pElemFirst
    iStep = (DIRECTION_FORWARD==dir) ? 1 : -1;

    for(i=(DIRECTION_FORWARD==dir)?0:cElems-1;
        (DIRECTION_FORWARD==dir)?(i<cElems):(i>=0); i+=iStep)
    {
        hr = pCollectionCache->GetIntoAry(CMarkup::ELEMENT_COLLECTION, i, &pElement);
        if(hr)
        {
            goto Cleanup;
        }

        Assert(pElement);

        // If the _fHasTabIndex bit is set, this element has already
        // been looked at in SearchFocusArray.
        if(pElement->_fHasTabIndex)
        {
            continue;
        }

        if(pElemFirst)
        {
            if(pElemFirst == pElement)
            {
                // Found the element.  Now check if there are any further
                // subdivisions.  If so, then we need to return the next
                // subdivision.  If we're on the last subdivision already
                // or if there are no subdivisions, then we need to return
                // the next tabbable object.
                fFound = TRUE;
                Assert(lSubFirst != -1);
                lSubNext = lSubFirst;

                for(; ;)
                {
                    hr = pElement->GetNextSubdivision(dir, lSubNext, &lSubNext);
                    if(hr)
                    {
                        goto Cleanup;
                    }

                    if(lSubNext == -1)
                    {
                        break;
                    }

                    if(pElement->IsTabbable(lSubNext))
                    {
                        *ppElemNext = pElement;
                        *plSubNext = lSubNext;
                        goto Cleanup;
                    }
                }

                fStopOnNextTab = TRUE;
                continue;
            }
        }
        else
        {
            fStopOnNextTab = TRUE;
        }

        if(fStopOnNextTab)
        {
            hr = pElement->GetNextSubdivision(dir, -1, &lSubNext);
            if(hr)
            {
                goto Cleanup;
            }
            while(lSubNext != -1)
            {
                if(pElement->IsTabbable(lSubNext))
                {
                    fFound = TRUE;
                    *ppElemNext = pElement;
                    *plSubNext = lSubNext;
                    goto Cleanup;
                }
                hr = pElement->GetNextSubdivision(dir, lSubNext, &lSubNext);
                if(hr)
                {
                    goto Cleanup;
                }
            }
        }
    }

Cleanup:
    return fFound;
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::SearchAccessKeyArray
//
//  Synopsis:   Search accessKey array for matched access key
//
//-------------------------------------------------------------------------
void CDocument::SearchAccessKeyArray(
           FOCUS_DIRECTION  dir,
           CElement*        pElemFirst,
           CElement**       ppElemNext,
           CMessage*        pmsg)
{
    long        lIndex;
    long        lSize = _aryAccessKeyItems.Size();
    CElement*   pElemIndex;

    Assert(ppElemNext);
    *ppElemNext = NULL;

    if(pElemFirst)
    {
        long lsrcIndex = pElemFirst->GetSourceIndex();

        for(lIndex=0; lIndex<lSize; lIndex++)
        {
            pElemIndex = _aryAccessKeyItems[lIndex];

            if(pElemIndex->GetSourceIndex() >= lsrcIndex)
            {
                if(pElemIndex->GetSourceIndex() == lsrcIndex)
                {
                    (dir==DIRECTION_FORWARD) ? (lIndex++) : (lIndex--);
                }
                else if(dir == DIRECTION_BACKWARD)
                {
                    lIndex--;
                }
                break;
            }
        }
        if(lIndex==lSize && dir==DIRECTION_BACKWARD)
        {
            lIndex = -1;
        }
    }
    else
    {
        lIndex = (dir==DIRECTION_FORWARD) ? 0 : lSize - 1;
    }

    for(; (dir==DIRECTION_FORWARD)?(lIndex<lSize):(lIndex>=0);
        (dir==DIRECTION_FORWARD)?(lIndex++):(lIndex--))
    {
        pElemIndex = _aryAccessKeyItems[lIndex];
        Assert(pElemIndex);

        if(pElemIndex->MatchAccessKey(pmsg) && pElemIndex!=pElemFirst)
        {
            FOCUS_ITEM fi = pElemIndex->GetMnemonicTarget();

            if(fi.pElement && fi.pElement->IsFocussable(fi.lSubDivision))
            {
                *ppElemNext = pElemIndex;
                break;
            }
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::InsertAccessKeyItems
//
//  Synopsis:   Insert CElement with accessKey defined in _aryAccessKeyItems
//
//-------------------------------------------------------------------------
HRESULT CDocument::InsertAccessKeyItem(CElement* pElement)
{
    HRESULT hr = S_OK;
    long    lSrcIndex = pElement->GetSourceIndex();
    long    lIndex;

    for(lIndex=_aryAccessKeyItems.Size()-1; lIndex>=0; lIndex--)
    {
        Assert(_aryAccessKeyItems[lIndex]->GetSourceIndex() != lSrcIndex);
        if(_aryAccessKeyItems[lIndex]->GetSourceIndex() < lSrcIndex)
        {
            break;
        }
    }

    hr = _aryAccessKeyItems.EnsureSize(_aryAccessKeyItems.Size()+1);
    if(hr)
    {
        goto Cleanup;
    }

    Verify(!_aryAccessKeyItems.Insert(lIndex+1, pElement));

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::InsertFocusArrayItem
//
//  Synopsis:   Insert CElement with tabindex defined in _aryFocusItems
//
//-------------------------------------------------------------------------
HRESULT CDocument::InsertFocusArrayItem(CElement* pElement)
{
    HRESULT     hr = S_OK;
    long        lTabIndex  = pElement->GetAAtabIndex();
    long        c;
    long*       pTabs = NULL;
    long*       pTabsSet;
    FOCUS_ITEM  focusitem;
    long        i;
    long        j;

    // The tabIndex attribute can have three values:
    //
    //      < 0     Does not participate in focus
    //     == 0     Tabindex as per source order
    //      > 0     Tabindex is the set tabindex
    //
    // Precedence is that given tab indices go first, then the elements
    // which don't have one assigned.  If not given a tab index, treat the
    // element same as tabIndex < 0 for those elements which don't have a
    // layout associated with them (e.g. <P>, <SPAN>).  Those that do
    // have a layout (e.g. <INPUT>, <BUTTON>) will be treated as tabindex == 0

    // Figure out if this element has subdivisions.  If so, then we need
    // multiple entries for this guy.

    // Areas don't belong in here.
    if(pElement->Tag() == ETAG_AREA)
    {
        goto Cleanup;
    }

    hr = pElement->GetSubDivisionCount(&c);
    if(hr)
    {
        goto Cleanup;
    }

    // Subdivisions can have their own tabIndices. Therefore, if an element
    // has any subdivisions, the element's own tabIndex is ignored.
    if(!c)
    {
        if(lTabIndex <= 0)
        {
            goto Cleanup;
        }
        pElement->_fHasTabIndex = TRUE;
    }

    // Find the location in the focus item array to insert this element.
    focusitem.pElement = pElement;
    if(c)
    {
        pTabs = new long[c];
        if(!pTabs)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = pElement->GetSubDivisionTabs(pTabs, c);
        if(hr)
        {
            goto Cleanup;
        }

        pTabsSet = pTabs;
    }
    else
    {
        Assert(lTabIndex > 0);
        pTabsSet = &lTabIndex;
        c = 1;
    }

    hr = _aryFocusItems.EnsureSize(_aryFocusItems.Size()+c);
    if(hr)
    {
        goto Cleanup;
    }

    for(i=0; i<c; i++)
    {
        // This is here because subdivisions can also be set to have
        // either a negative tabindex or zero, which means they're in source
        // order.
        if(pTabsSet[i] <= 0)
        {
            continue;
        }

        focusitem.lSubDivision = i;
        focusitem.lTabIndex = pTabsSet[i];

        // Find correct location in _aryFocusItems to insert.
        for(j=0; j<_aryFocusItems.Size(); j++)
        {
            if(_aryFocusItems[j].lTabIndex>focusitem.lTabIndex ||
                _aryFocusItems[j].pElement->GetSourceIndex()>focusitem.pElement->GetSourceIndex())
            {
                break;
            }
        }

        Verify(!_aryFocusItems.InsertIndirect(j, &focusitem));
    }

Cleanup:
    delete[] pTabs;
    RRETURN(hr);
}

void CDocument::SetFocusWithoutFiringOnfocus()
{
    BOOL fOldFiredWindowFocus = _fFiredWindowFocus;
    BOOL fInhibitFocusFiring = _fInhibitFocusFiring;

    // Do not fire window onfocus if the inplace window didn't
    // previously have the focus (NS compat)
    if(fInhibitFocusFiring)
    {
        _fFiredWindowFocus = TRUE;
    }
    ::SetFocus(_pInPlace->_hwnd);
    // restore the old FiredFocus state
    if(fInhibitFocusFiring)
    {
        _fFiredWindowFocus = fOldFiredWindowFocus;
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     CDocument::CreateRoot
//
//  Synopsis:   Create a main root for this doc
//
//--------------------------------------------------------------------------
HRESULT CDocument::CreateRoot()
{
    HRESULT hr = S_OK;

    Assert(!_pPrimaryMarkup);

    hr = CreateMarkup(&_pPrimaryMarkup);
    if(hr)
    {
        goto Cleanup;
    }

    _pElemCurrent = _pPrimaryMarkup->Root();

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::EnterStylesheetDownload
//
//  Synopsis:   Note that a stylesheet is being downloaded
//
//--------------------------------------------------------------------
void CDocument::EnterStylesheetDownload(DWORD* pdwCookie)
{
    if(*pdwCookie != _dwStylesheetDownloadingCookie)
    {
        *pdwCookie = _dwStylesheetDownloadingCookie;
        _cStylesheetDownloading++;
    }
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::LeaveStylesheetDownload
//
//  Synopsis:   Note that stylesheet is finished downloading
//
//--------------------------------------------------------------------
void CDocument::LeaveStylesheetDownload(DWORD* pdwCookie)
{
    if(*pdwCookie == _dwStylesheetDownloadingCookie)
    {
        *pdwCookie = 0;
        _cStylesheetDownloading--;
    }
}

CRootElement* CDocument::PrimaryRoot()
{
    Assert(_pPrimaryMarkup);
    return _pPrimaryMarkup->Root();
}

#ifdef _DEBUG
BOOL CDocument::AreLookasidesClear(void* pvKey, int nLookasides)
{
    DWORD* pdwKey = (DWORD*)pvKey;

    for(; nLookasides>0; nLookasides--,pdwKey++)
    {
        if(_HtPvPv.IsPresent(pdwKey))
        {
            return FALSE;
        }
    }

    return TRUE;
}
#endif

void CDocument::ClearDefaultCharFormat()
{
    if(_icfDefault >= 0)
    {
        TLS(_pCharFormatCache)->ReleaseData(_icfDefault);
        _icfDefault = -1;
        _pcfDefault = NULL;
    }
}

HRESULT CDocument::CacheDefaultCharFormat()
{
    Assert(_pOptionSettings);
    Assert(_pCodepageSettings);
    Assert(_icfDefault == -1);
    Assert(_pcfDefault == NULL);

    THREADSTATE* pts = GetThreadState();
    CCharFormat cf;
    HRESULT hr;

    cf.InitDefault(_pOptionSettings, _pCodepageSettings);
    cf._bCrcFont = cf.ComputeFontCrc();
    cf._fHasDirtyInnerFormats = !!cf.AreInnerFormatsDirty();

    hr = pts->_pCharFormatCache->CacheData(&cf, &_icfDefault);

    if(hr == S_OK)
    {
        _pcfDefault = &(*pts->_pCharFormatCache)[_icfDefault];
    }
    else
    {
        _icfDefault = -1;
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::RequestReleaseNotify
//
//---------------------------------------------------------------
HRESULT CDocument::RequestReleaseNotify(CElement* pElement)
{
    HRESULT hr = S_OK;

    _aryElementReleaseNotify.Append(pElement);
    pElement->SubAddRef();

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::RevokeRequestReleaseNotify
//
//---------------------------------------------------------------
HRESULT CDocument::RevokeRequestReleaseNotify(CElement* pElement)
{
    HRESULT hr = S_OK;
    LONG    idx = _aryElementReleaseNotify.Find(pElement);

    if(0 <= idx)
    {
        _aryElementReleaseNotify.Delete(idx);
        pElement->SubRelease();
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CDocument::ReleaseNotify
//
//  Synopsis:   Notifies registered elements to release contained objects.
//
//  Notes:      elements must register to get this notification using RequestReleaseNotify
//
//---------------------------------------------------------------
HRESULT CDocument::ReleaseNotify()
{
    CElement*       pElement;
    CNotification   nf;
    int             c;

    while(0 < (c=_aryElementReleaseNotify.Size()))
    {
        pElement = _aryElementReleaseNotify[c-1];

        _aryElementReleaseNotify.Delete(c-1);

        if(0 < pElement->GetObjectRefs())
        {
            nf.ReleaseExternalObjects(pElement);
            pElement->Notify(&nf);
        }

        pElement->SubRelease();
    }

    Assert(0 == _aryElementReleaseNotify.Size());

    return S_OK;
}

//+==========================================================
//
// Method: GetHTMLEditor
//
// Synopsis: This is the real GetHTMLEditor
//                QueryService on the Host for the HTMLEditor Service
//              if it's there - we QI the host for it.
//              if it's not there - we cocreate the HTMLEditor in Mshtmled
//
//-----------------------------------------------------------
IHTMLEditor* CDocument::GetHTMLEditor(BOOL fForceCreate/*=TRUE*/)
{
    HRESULT hr = S_OK;

    // If we have an editor already, just return it
    if(_pIHTMLEditor)
    {
        goto Cleanup;
    }

    if(fForceCreate)
    {
        // If the host doesn't want to be the editor, mshtmled sure does!
        _pIHTMLEditor = new CXindowsEditor();
        _pIHTMLEditor->AddRef();

        if(FAILED(hr) || _pIHTMLEditor==NULL)
        {
            goto Error;
        }

        // Use a weak-ref to all doc interfaces. Use _OmDocument for CView while it
        // doesn't exist...
        hr = _pIHTMLEditor->Initialize(
            (IUnknown*)(IPrivateUnknown*)_pPrimaryMarkup,
            (IUnknown*)(IPrivateUnknown*)_pPrimaryMarkup);
        if(FAILED(hr))
        {
            goto Error;
        }
    } // fForceCreate

    goto Cleanup;

Error:
    ClearInterface(&_pIHTMLEditor);

Cleanup:
    AssertSz(!(fForceCreate && _pIHTMLEditor==NULL), "IHTMLEditor Not Found or Allocated on Get!");

    return _pIHTMLEditor;
}

//+====================================================================================
//
// Method: GetSelectionDragDropSource
//
// Synopsis: If this doc is the source of a drag/drop, and we are drag/dropping a selection
//           return the CSelDragDropSrcInfo. Otherwise return NULL.
//
//------------------------------------------------------------------------------------
CSelDragDropSrcInfo* CDocument::GetSelectionDragDropSource()
{
    if(_fIsDragDropSrc && _pDragDropSrcInfo &&
        _pDragDropSrcInfo->_srcType==DRAGDROPSRCTYPE_SELECTION)
    {
        return DYNCAST(CSelDragDropSrcInfo, _pDragDropSrcInfo);
    }
    else
    {
        return NULL;
    }
}

//+==========================================================
//
// Method: UpdateCaret
//
// Synopsis: Informs the caret that it's position has been
//           changed externally.
//
//-----------------------------------------------------------
HRESULT CDocument::UpdateCaret(BOOL fScrollIntoView, BOOL fForceScroll, CDocInfo* pdci)
{
    HRESULT hr = S_OK;

    if(_pCaret)
    {
        hr = _pCaret->UpdateCaret(fScrollIntoView, fForceScroll, pdci);
    }
    RRETURN(hr);
}

//+==========================================================
//
// Method: HandleSelectionMessage
//
// Synopsis: Dispatch a message to the Selection Manager if it exists.
//
//
//-----------------------------------------------------------
HRESULT CDocument::HandleSelectionMessage(CMessage* pMessage, BOOL fForceCreate)
{
    BOOL fNeedToSetEditContext = FALSE;
    BOOL fAllowSelection = TRUE;

    if(pMessage->fSelectionHMCalled)
    {
        return S_FALSE;
    }
    else if(_pElemEditContext
        && pMessage->message>=WM_KEYFIRST
        && pMessage->message<=WM_KEYLAST
        && _pElemEditContext!=_pElemCurrent
        && !(_pElemCurrent->HasSlaveMarkupPtr()
        && _pElemCurrent->GetSlaveMarkupPtr()->FirstElement()==_pElemEditContext))
    {
        // We get here if the selection is in an alement that does not have focus.
        // In such a state, we want commands to be executable on the selection but
        // we do not want keystroke messages to go to the selection.
        return S_FALSE;
    }    
    else
    {
        HRESULT hr = S_FALSE;
        IHTMLEditor* ped = NULL;
        pMessage->fSelectionHMCalled = TRUE;

        CElement* pEditElement = NULL;

        if((pMessage->pNodeHit) && (pMessage->pNodeHit->Element()->_etag==ETAG_ROOT))
        {
            goto Cleanup;
        }

        if(pMessage->pNodeHit)
        {
            pEditElement = pMessage->pNodeHit->Element();
        }
        else
        {
            pEditElement = NULL;
        }

        AssertSz(!_pDragStartInfo||(_pDragStartInfo&&pMessage->message!=WM_MOUSEMOVE),
            "Sending a Mouse Message to the tracker during a drag !");

        // We just got a mouse down. If the element is not editable,
        // we need to set the Edit Context ( as SetEditContext may not have already happened)
        // If we don't do this the manager may not have a tracker for the event !
        fNeedToSetEditContext = (pEditElement && 
            (ShouldSetEditContext(pMessage) && 
            (!pEditElement->IsEditable(FALSE) || 
            !pEditElement->IsEnabled() || 
            !_pElemEditContext)));

        if(fNeedToSetEditContext) // we only check if we think we need to set the ed. context
        {
            fAllowSelection = pEditElement && !pEditElement->DisallowSelection();
        }

        ped = GetHTMLEditor(fForceCreate || (fNeedToSetEditContext && fAllowSelection));

        // Set the ON-Click handler to make sure On-click event still works
        // even if selection is disallowed
        if(!ped && pMessage->pNodeHit && pMessage->message==WM_LBUTTONUP)
        {
            SetClick(pMessage);
        }

        if(ped && fNeedToSetEditContext)
        {
            hr = SetEditContext(pEditElement, FALSE, FALSE);
            if(hr)
            {
                goto Cleanup;
            }
        }

        if(ped && (pMessage->pNodeHit==NULL || !pMessage->pNodeHit->IsDead()))
        {   
            SelectionMessage theMessage;
            DWORD takeAction = FOLLOW_UP_ACTION_None;

            CMessageToSelectionMessage(pMessage, &theMessage);

            hr = ped->HandleMessage(&theMessage, &takeAction);

            // BUGBUG (MohanB) Need to copy any other info?
            // Perhaps call SelectionMessageToCMessage?
            pMessage->fStopForward = theMessage.fStopForward;

            if(takeAction != FOLLOW_UP_ACTION_None)
            {
                // BUGBUG - we go through this rigmarole to not lose the HR passed back from ed.
                // We want to eventually remove ProcessFollowUp entirely.
                //
                // (jbeda) We want to let S_OK override S_FALSE from either call
                // (HandleMessage or ProcessFollowUpAction) but we want to let
                // any error code through.
                HRESULT hrProcessFollowUp;
                hrProcessFollowUp = ProcessFollowUpAction(pMessage, takeAction);
                if(hrProcessFollowUp!=S_FALSE && SUCCEEDED(hr))
                {
                    hr = hrProcessFollowUp;
                }
            }

            pMessage->lresult = theMessage.lResult;
        }

Cleanup:
        return hr;
    }
}

//+====================================================================================
//
// Method: SetEditContext
//
// Synopsis: Sets the 'Edit Context' - by cocreating a Selection Mgr if necessary.
//
//------------------------------------------------------------------------------------
HRESULT CDocument::SetEditContext(CElement* pElement, BOOL fForceCreate, BOOL fSetSelection, BOOL fDrillingIn)
{
    HRESULT hr = S_OK;

    IHTMLEditor* ped = NULL;
    IMarkupPointer* pStart = NULL;
    IMarkupPointer* pEnd = NULL;
    CElement* pEditableElement = NULL;
    BOOL fThisEditable = FALSE;
    BOOL fParentEditable = FALSE;
    BOOL fNoScope = FALSE;

    Assert(pElement);

    ped = GetHTMLEditor(fForceCreate);

    if(ped)
    {
        hr = CreateMarkupPointer(&pStart);
        if(hr)
        {
            goto Cleanup;
        }

        hr = CreateMarkupPointer(&pEnd);
        if(hr)
        {
            goto Cleanup;
        }

        hr = AdjustEditContext( 
            pElement,
            &pEditableElement,
            pStart,
            pEnd,
            &fThisEditable ,
            &fParentEditable,
            &fNoScope ,
            FALSE,
            TRUE,
            fDrillingIn);

        hr = ped->SetEditContext( 
            fThisEditable,
            fSetSelection,
            fParentEditable,
            pStart,
            pEnd , 
            fNoScope);
    }

Cleanup:
    ReleaseInterface(pStart);
    ReleaseInterface(pEnd);

    return hr;
}

//+====================================================================================
//
// Method:NotifySelection
//
// Synopsis: Notify the HTMLEditor that "something" happened.
//
//------------------------------------------------------------------------------------
HRESULT CDocument::NotifySelection(SELECTION_NOTIFICATION eSelectionNotification, IUnknown* pUnknown, DWORD dword)
{
    HRESULT hr = S_OK;
    DWORD theAction = FOLLOW_UP_ACTION_None;
    IHTMLEditor* ped = NULL;

    if((eSelectionNotification==SELECT_NOTIFY_DOC_ENDED) && _pIHTMLEditor)
    {
        StopHTMLEditorDblClickTimer();
    }
    // If we get a timer tick - when the doc is shut down this is very bad
    Assert(!((eSelectionNotification==SELECT_NOTIFY_TIMER_TICK)
        && (_ulRefs==ULREF_IN_DESTRUCTOR)));

    if(_pIHTMLEditor==NULL && !ShouldCreateHTMLEditor(eSelectionNotification))
    {
        goto Cleanup; // Nothing to do
    }

    ped = GetHTMLEditor(TRUE);

    if(ped)
    {
        hr = ped->Notify(eSelectionNotification, pUnknown, &theAction, dword);
        if(theAction != FOLLOW_UP_ACTION_None)
        {
            ProcessFollowUpAction(NULL, theAction);
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::CleanupScriptTimers
//
//----------------------------------------------------------------------------
void CDocument::CleanupScriptTimers(void)
{
    _TimeoutEvents.KillAllTimers(this);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::AddRefUrlImgCtx
//
//  Synopsis:   Returns a cookie to the background image specified by the
//              given Url and element context.
//
//  Arguments:  lpszUrl         Relative URL of the background image
//              pElemContext    Element to begin search for nearest <BASE> tag
//
//  Returns:    Cookie which refers to the background image
//
//----------------------------------------------------------------------------
HRESULT CDocument::AddRefUrlImgCtx(LPCTSTR lpszUrl, CElement* pElemContext, LONG* plCookie)
{
    CDwnCtx*    pDwnCtx;
    CImgCtx*    pImgCtx;
    URLIMGCTX*  purlimgctx;
    LONG        iurlimgctx;
    LONG        iurlimgctxFree = -1;
    LONG        curlimgctx;
    HRESULT     hr;
    TCHAR       cBuf[pdlUrlLen];
    TCHAR*      pszExpUrl = cBuf;

    *plCookie = 0;

    hr = ExpandUrl(lpszUrl, ARRAYSIZE(cBuf), pszExpUrl, pElemContext);
    if(hr)
    {
        goto Cleanup;
    }

    // See if we've already got a background image with this Url
    purlimgctx = _aryUrlImgCtx;
    curlimgctx = _aryUrlImgCtx.Size();

    for(iurlimgctx=0; iurlimgctx<curlimgctx; ++iurlimgctx,++purlimgctx)
    {
        if(purlimgctx->ulRefs == 0)
        {
            iurlimgctxFree = iurlimgctx;
        }
        else
        {
            const TCHAR* pchSlotUrl;

            pchSlotUrl = purlimgctx->pImgCtx ? purlimgctx->pImgCtx->GetUrl() : purlimgctx->cstrUrl;
            Assert(pchSlotUrl);

            if(pchSlotUrl && StrCmpC(pchSlotUrl, pszExpUrl)==0)
            {
                // Found it.  Increment the reference count on this entry and
                // hand out a cookie to it.
                purlimgctx->ulRefs += 1;
                *plCookie = iurlimgctx + 1;

                hr = purlimgctx->aryElems.Append(pElemContext);

                goto Cleanup;
            }
        }
    }

    // No luck finding an existing image.  Get a new one and add it to array.
    if(iurlimgctxFree == -1)
    {
        hr = _aryUrlImgCtx.EnsureSize(iurlimgctx+1);
        if(hr)
        {
            goto Cleanup;
        }

        iurlimgctxFree = iurlimgctx;

        _aryUrlImgCtx.SetSize(iurlimgctx+1);

        // N.B. (johnv) We need this so that the CPtrAry inside URLIMGCTX
        // gets initialized.
        memset(&_aryUrlImgCtx[iurlimgctxFree], 0, sizeof(URLIMGCTX));
    }

    purlimgctx = &_aryUrlImgCtx[iurlimgctxFree];

    {
        // BUGBUG Pass lpszUrl, not pszExpUrl, due to ExpandUrl bug.
        hr = NewDwnCtx(DWNCTX_IMG, lpszUrl, pElemContext, &pDwnCtx, TRUE);
        if(hr)
        {
            goto Cleanup;
        }

        pImgCtx = (CImgCtx*)pDwnCtx;

    }

    if(pImgCtx)
    {
        pImgCtx->SetProgSink(GetProgSink());
        pImgCtx->SetCallback(OnUrlImgCtxCallback, this);
        pImgCtx->SelectChanges((_pOptionSettings->fPlayAnimations) ?
            IMGCHG_COMPLETE|IMGCHG_ANIMATE:IMGCHG_COMPLETE, 0, TRUE); // TRUE: disallow prompting
    }

    hr = purlimgctx->aryElems.Append(pElemContext);
    if(hr)
    {
        goto Cleanup;
    }

    purlimgctx->ulRefs   = 1;
    purlimgctx->pImgCtx  = pImgCtx;
    *plCookie            = iurlimgctxFree + 1;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::AddRefUrlImgCtx
//
//  Synopsis:   Adds a reference to the background image specified by the
//              given cookie.
//
//  Arguments:  lCookie         Cookie given out by AddRefUrlImgCtx
//
//----------------------------------------------------------------------------
HRESULT CDocument::AddRefUrlImgCtx(LONG lCookie, CElement* pElem)
{
    HRESULT hr;

    if(!lCookie)
    {
        return S_OK;
    }

    Assert(lCookie>0 && lCookie<=_aryUrlImgCtx.Size());
    Assert(_aryUrlImgCtx[lCookie-1].ulRefs > 0);

    hr = _aryUrlImgCtx[lCookie-1].aryElems.Append(pElem);
    if(hr)
    {
        goto Cleanup;
    }

    _aryUrlImgCtx[lCookie-1].ulRefs += 1;

    Trace((_T("AddRefUrlImgCtx (#%ld,url=%ls,cRefs=%ld)\n"), lCookie,
        _aryUrlImgCtx[lCookie-1].pImgCtx?_aryUrlImgCtx[lCookie-1].pImgCtx->GetUrl():_aryUrlImgCtx[lCookie-1].cstrUrl,
        _aryUrlImgCtx[lCookie-1].ulRefs));

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetUrlImgCtx
//
//  Synopsis:   Returns the CImgCtx at the specified cookie
//
//  Arguments:  lCookie         Cookie given out by AddRefUrlImgCtx
//
//----------------------------------------------------------------------------
CImgCtx* CDocument::GetUrlImgCtx(LONG lCookie)
{
    if(!lCookie)
    {
        return(NULL);
    }

    Assert(lCookie>0 && lCookie<=_aryUrlImgCtx.Size());
    Assert(_aryUrlImgCtx[lCookie-1].ulRefs > 0);

    return _aryUrlImgCtx[lCookie-1].pImgCtx;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::GetImgAnimState
//
//  Synopsis:   Returns an IMGANIMSTATE for the specified cookie, null if
//              there is none
//
//  Arguments:  lCookie         Cookie given out by AddRefUrlImgCtx
//
//----------------------------------------------------------------------------
IMGANIMSTATE* CDocument::GetImgAnimState(LONG lCookie)
{
    CImgAnim* pImgAnim = GetImgAnim();
    LONG lAnimCookie;

    if(!lCookie || !pImgAnim)
    {
        return NULL;
    }

    Assert(lCookie>0 && lCookie<=_aryUrlImgCtx.Size());
    Assert(_aryUrlImgCtx[lCookie-1].ulRefs > 0);

    lAnimCookie = _aryUrlImgCtx[lCookie-1].lAnimCookie;

    if(lAnimCookie)
    {
        return(pImgAnim->GetImgAnimState(lAnimCookie));
    }

    return NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ReleaseUrlImgCtx
//
//  Synopsis:   Releases a reference to the background image specified by
//              the given cookie.
//
//  Arguments:  lCookie         Cookie given out by AddRefUrlImgCtx
//              pElem           An element associated with this cookie
//
//----------------------------------------------------------------------------
void CDocument::ReleaseUrlImgCtx(LONG lCookie, CElement* pElem)
{
    if(!lCookie)
    {
        return;
    }

    Assert(lCookie>0 && lCookie <= _aryUrlImgCtx.Size());
    URLIMGCTX* purlimgctx = &_aryUrlImgCtx[lCookie-1];
    Assert(purlimgctx->ulRefs > 0);

    Trace((_T("ReleaseUrlImgCtx (#%ld,url=%ls,cRefs=%ld)\n"),
        lCookie,
        _aryUrlImgCtx[lCookie-1].pImgCtx?_aryUrlImgCtx[lCookie-1].pImgCtx->GetUrl():NULL,
        _aryUrlImgCtx[lCookie-1].ulRefs-1));

    Verify(purlimgctx->aryElems.DeleteByValue(pElem));

    if(--purlimgctx->ulRefs == 0)
    {
        Assert(purlimgctx->aryElems.Size() == 0);

        // Release our animation cookie if we have one
        if(purlimgctx->lAnimCookie)
        {
            CImgAnim* pImgAnim = GetImgAnim();

            if(pImgAnim)
            {
                pImgAnim->UnregisterForAnim(this, purlimgctx->lAnimCookie);
            }
        }

        purlimgctx->aryElems.DeleteAll();
        if(purlimgctx->pImgCtx)
        {
            purlimgctx->pImgCtx->SetProgSink(NULL); // detach download from document's load progress
            purlimgctx->pImgCtx->Disconnect();
            purlimgctx->pImgCtx->Release();
        }
        purlimgctx->cstrUrl.Free();
        memset(purlimgctx, 0, sizeof(*purlimgctx));
    }
}

//----------------------------------------------------------
//
//  Member   : CDocument::StopUrlImgCtx
//
//  Synopsis : Stops downloading of all background images
//
//----------------------------------------------------------
void CDocument::StopUrlImgCtx()
{
    URLIMGCTX*  purlimgctx;
    LONG        curlimgctx;

    purlimgctx = _aryUrlImgCtx;
    curlimgctx = _aryUrlImgCtx.Size();

    for(; curlimgctx>0; --curlimgctx,++purlimgctx)
    {
        if(purlimgctx->pImgCtx)
        {
            purlimgctx->pImgCtx->SetLoad(FALSE, NULL, FALSE);
        }
    }
}

//----------------------------------------------------------
//
//  Member   : CDocument::UnregisterUrlImageCtxCallbacks
//
//  Synopsis : Cancels any image callbacks for the doc.  Does
//             not release the image context itself.
//
//----------------------------------------------------------
void CDocument::UnregisterUrlImgCtxCallbacks()
{
    CImgAnim*   pImgAnim = GetImgAnim();
    URLIMGCTX*  purlimgctx;
    LONG        iurlimgctx;
    LONG        curlimgctx;

    purlimgctx = _aryUrlImgCtx;
    curlimgctx = _aryUrlImgCtx.Size();

    for(iurlimgctx=0; iurlimgctx<curlimgctx; ++iurlimgctx,++purlimgctx)
    {
        if(purlimgctx->ulRefs)
        {
            // Unregister callbacks from the animation object, if any
            if(pImgAnim && purlimgctx->lAnimCookie)
            {
                pImgAnim->UnregisterForAnim(this, purlimgctx->lAnimCookie);
                purlimgctx->lAnimCookie = 0;
            }

            if(purlimgctx->pImgCtx)
            {
                // Unregister callbacks from the image context
                purlimgctx->pImgCtx->SetProgSink(NULL); // detach download from document's load progress
                purlimgctx->pImgCtx->Disconnect();
            }
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnUrlImgCtxCallback
//
//  Synopsis:   Callback from background image reporting that it is finished
//              loading.
//
//  Arguments:  pvObj         The pImgCtx that is calling back
//              pbArg         The CDoc pointer
//
//----------------------------------------------------------------------------
void CALLBACK CDocument::OnUrlImgCtxCallback(void* pvObj, void* pvArg)
{
    CDocument*  pDoc       = (CDocument*)pvArg;
    CImgCtx*    pImgCtx    = (CImgCtx*)pvObj;
    LONG        iurlimgctx;
    LONG        curlimgctx = pDoc->_aryUrlImgCtx.Size();
    URLIMGCTX*  purlimgctx = pDoc->_aryUrlImgCtx;
    SIZE        size;
    ULONG       ulState    = pImgCtx->GetState(TRUE, &size);

    pImgCtx->AddRef();

    for(iurlimgctx=0; iurlimgctx<curlimgctx; ++iurlimgctx,++purlimgctx)
    {
        if(pImgCtx && purlimgctx->pImgCtx==pImgCtx)
        {
            Trace3("OnUrlImgCtxCallback (#%ld,url=%ls,cRefs=%ld)\n",
                iurlimgctx+1,
                purlimgctx->pImgCtx?purlimgctx->pImgCtx->GetUrl():purlimgctx->cstrUrl,
                purlimgctx->ulRefs);

            if(ulState & IMGCHG_ANIMATE)
            {
                // Register for animation callbacks
                CImgAnim* pImgAnim = CreateImgAnim();

                if(!purlimgctx->lAnimCookie)
                {
                    pImgAnim->RegisterForAnim(
                        pDoc,
                        (DWORD_PTR)pDoc,
                        pImgCtx->GetImgId(),
                        OnAnimSyncCallback,
                        (void*)(DWORD_PTR)iurlimgctx,
                        &purlimgctx->lAnimCookie);
                }

                if(purlimgctx->lAnimCookie)
                {
                    pImgAnim->ProgAnim(purlimgctx->lAnimCookie);
                }
            }
            if(ulState & (IMGLOAD_COMPLETE|IMGLOAD_STOPPED|IMGLOAD_ERROR))
            {
                pImgCtx->SetProgSink(NULL); // detach download from document's load progress
            }
            if(ulState & IMGLOAD_COMPLETE)
            {
                pDoc->OnUrlImgCtxChange(purlimgctx, IMGCHG_COMPLETE);
            }

            break;
        }
    }
    pImgCtx->Release();
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnUrlImgCtxChange
//
//  Synopsis:   Called when this image context has changed, does the necessary
//              invalidating.
//
//  Arguments:  purlimgctx    The url image ctx that changed
//              ulState       The reason for the callback
//
//----------------------------------------------------------------------------
BOOL CDocument::OnUrlImgCtxChange(URLIMGCTX* purlimgctx, ULONG ulState)
{
    if(!_pInPlace)
    {
        return FALSE; // no window yet, nothing to do
    }

    BOOL        fSynchronousPaint = TRUE;
    int         n;
    CElement**  ppElem;
    CLayout*    pLayout;

    Trace((_T("OnChange for doc %ls, img %ls, %ld elements"),
        (TCHAR*)_cstrUrl,
        purlimgctx->pImgCtx?purlimgctx->pImgCtx->GetUrl():purlimgctx->cstrUrl,
        purlimgctx->aryElems.Size()));

    Assert(purlimgctx->pImgCtx);

    for(n=purlimgctx->aryElems.Size(),ppElem=purlimgctx->aryElems;
        n>0; n--,ppElem++)
    {
        CElement* pElement = *ppElem;
        // marka - check that element still in tree - bug # 15481.
        if(pElement->IsInMarkup())
        {
            pLayout = pElement->GetFirstBranch()->GetUpdatedNearestLayout();

            // if the background is on HTML element or any other ancestor of
            // body/top element, the inval the top element.
            if(!pLayout)
            {
                CElement* pElemClient = pElement->GetMarkup()->GetElementClient();

                if(pElemClient)
                {
                    pLayout = pElemClient->GetUpdatedLayout();
                }
            }

            if(pLayout)
            {
                if(pLayout->ElementOwner() == pElement)
                {
                    if(OpenView())
                    {
                        CDispNodeInfo dni;

                        pLayout->GetDispNodeInfo(&dni);
                        pLayout->EnsureDispNodeBackground(dni);
                    }
                }

                // Some elements require a resize, others a simple invalidation
                if(pElement->_fResizeOnImageChange && (ulState&IMGCHG_COMPLETE))
                {
                    pElement->ResizeElement();
                    fSynchronousPaint = FALSE;
                }
                else
                {
                    // We can get away with just an invalidate
                    pElement->Invalidate();
                }
            }
        }
    }

    return fSynchronousPaint;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::OnAnimSyncCallback
//
//  Synopsis:   Called back on an animation event
//
//----------------------------------------------------------------------------
void CDocument::OnAnimSyncCallback(
           void*    pvObj,
           DWORD    dwReason,
           void*    pvArg,
           void**   ppvDataOut,
           IMGANIMSTATE* pImgAnimState)
{
    CDocument* pDoc  = (CDocument*)pvObj;
    URLIMGCTX* purlimgctx = &pDoc->_aryUrlImgCtx[(LONG)(LONG_PTR)pvArg];

    switch(dwReason)
    {
    case ANIMSYNC_GETIMGCTX:
        *(CImgCtx**)ppvDataOut = purlimgctx->pImgCtx;
        break;

    case ANIMSYNC_GETHWND:
        *(HWND*)ppvDataOut = pDoc->_pInPlace ? pDoc->_pInPlace->_hwnd : NULL;
        break;

    case ANIMSYNC_TIMER:
    case ANIMSYNC_INVALIDATE:
        *(BOOL*)ppvDataOut = pDoc->OnUrlImgCtxChange(purlimgctx, IMGCHG_ANIMATE);
        break;

    default:
        Assert(FALSE);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     SwitchCodePage
//
//  Synopsis:   Change the codepage of the document.  Should only be used
//              when we are already at a valid codepage setting.
//
//-------------------------------------------------------------------------
HRESULT CDocument::SwitchCodePage(CODEPAGE codepage)
{
    HRESULT hr = S_OK;
    BOOL fReadCodePageSettings;
    UINT uiFamilyCodePage = CP_UNDEFINED;

    // If codepage settings don't exist or the information is stale,
    // reset the information, as well as the charformat cache.
    if(_pCodepageSettings)
    {
        if(_pCodepageSettings->uiFamilyCodePage != codepage)
        {
            uiFamilyCodePage = WindowsCodePageFromCodePage(codepage);

            fReadCodePageSettings = _pCodepageSettings->uiFamilyCodePage != uiFamilyCodePage;
        }
        else
        {
            fReadCodePageSettings = FALSE;
        }
    }
    else
    {
        uiFamilyCodePage = WindowsCodePageFromCodePage(codepage);
        fReadCodePageSettings = TRUE;
    }

    if(fReadCodePageSettings)
    {
        // Read the settings from the registry
        // Note that this call modifies CDocument::_codepage.
        hr = ReadCodepageSettingsFromRegistry(codepage, uiFamilyCodePage, FALSE);

        if(_pCodepageSettings)
        {
            _sBaselineFont = _pCodepageSettings->sBaselineFontDefault;
        }

        ClearDefaultCharFormat();
        ForceRelayout();
    }

    // Set the codepage
    _codepage = codepage;
    _codepageFamily = _pCodepageSettings->uiFamilyCodePage;

    RRETURN(hr);
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::GetDocDirection(pfRTL)
//
//  Synopsis:   Gets the document reading order. This is just a
//              reflection of the direction of the HTML element.
//
//  Returns:    S_OK if the direction was successfully set/retrieved.
//
//--------------------------------------------------------------------
HRESULT CDocument::GetDocDirection(BOOL* pfRTL)
{
    long eHTMLDir = htmlDirNotSet;
    BSTR v = NULL;
    HRESULT hr;

    if (!pfRTL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }
    *pfRTL = FALSE;

    hr = get_dir(&v);
    if(hr)
    {
        goto Cleanup;
    }

    hr = s_enumdeschtmlDir.EnumFromString(v, &eHTMLDir);
    if(hr)
    {
        goto Cleanup;
    }

    *pfRTL = (eHTMLDir == htmlDirRightToLeft);

Cleanup:
    FormsFreeString(v);
    RRETURN(hr);
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::IsCpAutoDetect 
//
//  Synopsis:   Get the flag that indicates cp is to be auto-detected
//
//  Returns:    BOOL true if auto mode is set
//
//--------------------------------------------------------------------
BOOL CDocument::IsCpAutoDetect(void)
{
    BOOL bret;

    if(_pOptionSettings)
    {
        bret = (BOOL)_pOptionSettings->fCpAutoDetect;
    }
    else
    {
        bret = FALSE;
    }

    return bret;
}

//+-------------------------------------------------------------------
//
//  Member:     CDocument::SetPeersPossible
//
//--------------------------------------------------------------------
void CDocument::SetPeersPossible()
{
    if(_fPeersPossible) // if already set
    {
        return;
    }

    if(_dwLoadf & DLCTL_NO_BEHAVIORS)
    {
        return;
    }

    _fPeersPossible = TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     ExecuteSingleFilterTask
//
//  Synopsis:   Execute pending filter hookup for a single element
//
//              Can be called at anytime without worrying about trashing
//              the queue array
//
//----------------------------------------------------------------------------
BOOL CDocument::ExecuteSingleFilterTask(CElement* pElement)
{
    if(!pElement->_fHasPendingFilterTask)
    {
        return FALSE;
    }

    int i = _aryPendingFilterElements.Find(pElement);

    Assert(i >= 0);

    if(i < 0)
    {
        return FALSE;
    }

    Assert(_aryPendingFilterElements[i] == pElement);

    // This function doesn't actually delete the entry, it simply nulls it out
    // The array is cleaned up when ExecuteFilterTasks is complete
    _aryPendingFilterElements[i] = NULL;

    pElement->_fHasPendingFilterTask = FALSE;

    Trace1("%08x demand executing filter task\n", pElement);

    return TRUE;
}

BOOL CDocument::AddFilterTask(CElement* pElement)
{
    if(!pElement->_fHasPendingFilterTask)
    {
        Trace1("Adding filter task for element %08x\n", pElement);

        Assert(_aryPendingFilterElements.Find(pElement) == -1);

        pElement->_fHasPendingFilterTask = SUCCEEDED(_aryPendingFilterElements.Append(pElement));

        PostFilterCallback();
    }

    return pElement->_fHasPendingFilterTask;
}

void CDocument::RemoveFilterTask(CElement* pElement)
{
    if(pElement->_fHasPendingFilterTask)
    {
        int i = _aryPendingFilterElements.Find(pElement);

        Assert(i >= 0);
        if(i >= 0)
        {
            _aryPendingFilterElements[i] = NULL;
            pElement->_fHasPendingFilterTask = FALSE;

            Trace1("%08x Removing filter task\n", pElement);

            // We don't delete anything from the array, that will happen when the
            // ExecuteFilterTasks is called.
            Assert(_fPendingFilterCallback);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     ExecuteFilterTasks
//
//  Synopsis:   Execute pending filter hookup
//
//  Notes;      This code is 100% re-entrant safe.  Call it whenever, however
//              you want and it must do the right thing.  Re-entrancy into
//              other functions (like ComputeFormats) is the responsibility
//              of the caller :-)
//
//  Arguments:  grfLayout - Collections of LAYOUT_xxxx flags
//
//  Returns:    TRUE if all tasks were processed, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CDocument::ExecuteFilterTasks()
{
    // We don't want to do this while we're painting
    if(TestLock(SERVERLOCK_BLOCKPAINT))
    {
        if(_aryPendingFilterElements.Size())
        {
            PostFilterCallback();
            return FALSE;
        }

        return TRUE; // No tasks means that they all got done, right?
    }

    // Sometimes this gets called on demand.  We may have a posted callback, get rid of it now.
    // REVIEW (michaelw) Should this be inside the FORMLOCK_FILTER block?
    if(_fPendingFilterCallback)
    {
        _fPendingFilterCallback = FALSE;
        GWKillMethodCall(this, ONCALL_METHOD(CDocument, FilterCallback, filtercallback), 0);
    }

    // If we're in the middle of doing this (for whatever reason), don't start again
    // The primary reason for not starting again has to do with trying to cleanup the
    // array, otherwise this code is re-entrant safe
    Assert(!TestLock(FORMLOCK_FILTER));
    if(!TestLock(FORMLOCK_FILTER))
    {
        CLock lock(this, FORMLOCK_FILTER);

        // We're only going to do as many elements as are there when we start
        int c = _aryPendingFilterElements.Size();

        if(c > 0)
        {
            for(int i=0; i<c; i++)
            {
                CElement* pElement = _aryPendingFilterElements[i];

                if(pElement)
                {
                    _aryPendingFilterElements[i] = NULL;

                    Assert(pElement->_fHasPendingFilterTask);
                    pElement->_fHasPendingFilterTask = FALSE;

                    Trace1("%08x Executing filter task from filter task list (also removes)\n", pElement);

                    // This calls out to external code, anything could happen?
                    if(_aryPendingFilterElements.Size() < c)
                    {
                        return FALSE;
                    }
                }
            }

            // Adding filters occasionally causes more filter
            // work.  Rather than doing it immediately, we'll
            // just wait until next time.
            Assert(_aryPendingFilterElements.Size() >= c);

            if(_aryPendingFilterElements.Size() > c)
            {
                PostFilterCallback();
            }
            _aryPendingFilterElements.DeleteMultiple(0, c-1);
        }
    }

    return (_aryPendingFilterElements.Size() == 0);
}

void CDocument::PostFilterCallback()
{
    if(!_fPendingFilterCallback)
    {
        Trace0("PostFilterCallback\n");
        _fPendingFilterCallback = SUCCEEDED(GWPostMethodCall(this, ONCALL_METHOD(CDocument, FilterCallback, filtercallback), 0, FALSE, "CDocument::FilterCallback"));
        Assert(_fPendingFilterCallback);
    }
}

void CDocument::FilterCallback(DWORD_PTR)
{
    Assert(_fPendingFilterCallback);
    if(_fPendingFilterCallback)
    {
        Trace0("FilterCallback\n");
        _fPendingFilterCallback = FALSE;
        ExecuteFilterTasks();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::elementFromPoint, public
//
//  Synopsis:   Returns the element under the given point.
//
//  Arguments:  [x]         -- x position
//              [y]         -- y position
//              [ppElement] -- Place to return element.
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CDocument::elementFromPoint(long x, long y, IHTMLElement** ppElement)
{
    HRESULT     hr = S_OK;
    POINT       pt = { x, y };
    CTreeNode*  pNode;
    HTC         htc;
    CMessage    msg;

    msg.pt = pt;

    if(!ppElement)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }
    *ppElement = NULL;

    htc = HitTestPoint(&msg, &pNode, HT_VIRTUALHITTEST);

    if(htc != HTC_NO)
    {
        Assert(pNode->Tag() != ETAG_ROOT);

        if(pNode->Tag() == ETAG_TXTSLAVE)
        {
            Assert(pNode->Element()->MarkupMaster());
            pNode = pNode->Element()->MarkupMaster()->GetFirstBranch();
            Assert(pNode);
        }

        hr = pNode->GetElementInterface(IID_IHTMLElement, (void**)ppElement);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     CDocument::get_Script
//
//  Synopsis:   returns OmWindow.  This routine returns the browser's
//   implementation of IOmWindow object, not our own _pOmWindow object.  This
//   is because the browser's object is the one with the longest lifetime and
//   which is handed to external entities.  See window.cxx COmWindow2::XXXX
//   for crazy delegation code.
//
//-----------------------------------------------------------------------------
HRESULT CDocument::get_Script(IDispatch** ppWindow)
{
    HRESULT hr;

    if(!ppWindow)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *ppWindow = NULL;

    //hr = EnsureOmWindow();
    //if(hr)
    //{
    //    goto Cleanup;
    //}

    //hr = _pOmWindow->QueryInterface(IID_IDispatch, (void**)ppWindow);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CDocument::get_expando(VARIANT_BOOL* pfExpando)
{
    HRESULT hr = S_OK;

    if(pfExpando)
    {
        *pfExpando = _fExpando ? VB_TRUE : VB_FALSE;
    }
    else
    {
        hr = E_INVALIDARG;
    }

    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CDocument::put_expando(VARIANT_BOOL fExpando)
{
    if(((_fExpando)?VB_TRUE:VB_FALSE) != fExpando)
    {
        _fExpando = fExpando;
        Fire_PropertyChangeHelper(DISPID_CDoc_expando, 0);
    }

    RRETURN(SetErrorInfo(S_OK));
}

//+----------------------------------------------------------------------------
//
//  Member:     curWindow, IOmDocument
//
//  Synopsis: returns a pointer to the top window
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_parentWindow(IHTMLWindow2** ppWindow)
{
    HRESULT hr;

    if(!ppWindow)
    {
        RRETURN(E_POINTER);
    }

    *ppWindow = NULL;

    /*hr = EnsureOmWindow();
    if(hr)
    {
        goto Cleanup;
    }

    hr = _pOmWindow->QueryInterface(IID_IHTMLWindow2, (void**)ppWindow); wlw note*/

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+------------------------------------------------------------------------
//
//  Member:     CDocument::get_uniqueID
//
//  Synopsis:   Returns a unique identifier that is guaranteed to not be
//              used by any element.
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CDocument::get_uniqueID(BSTR* pID)
{
    HRESULT hr;
    CString cstrUniqueID;

    hr = GetUniqueIdentifier(&cstrUniqueID);
    if(hr)
    {
        goto Cleanup;
    }

    hr = cstrUniqueID.AllocBSTR(pID);

Cleanup:
    RRETURN(hr);
}

HRESULT CDocument::releaseCapture()
{
    if(_pElementOMCapture)
    {
        CElement* pElement = _pElementOMCapture;

        _pElementOMCapture = NULL;
        _fOnLoseCapture = TRUE;
        pElement->Fire_onlosecapture();
        _fOnLoseCapture = FALSE;
        pElement->SubRelease();
        SetMouseCapture(NULL, NULL, TRUE);
    }

    RRETURN(SetErrorInfo(S_OK));
}

void CDocument::Fire_onblur(BOOL fOnblurFiredFromWM)
{
    // if onfocus was last fired then fire onblur
    /*EnsureOmWindow();
    if(_fFiredWindowFocus && _pOmWindow)
    {
        // Enable window onfocus firing next
        _fFiredWindowFocus = FALSE;
        GWPostMethodCall(_pOmWindow, ONCALL_METHOD(COmWindowProxy, Fire_onblur, fire_onblur), fOnblurFiredFromWM, TRUE, "COmWindowProxy::Fire_onblur");
    } wlw note*/
}

void CDocument::Fire_onfocus()
{
    // if onblur was last fired then fire onfocus
    /*EnsureOmWindow();
    if(!_fFiredWindowFocus && _pOmWindow)
    {
        // Enable window onblur firing next
        _fFiredWindowFocus = TRUE;
        GWPostMethodCall(_pOmWindow, ONCALL_METHOD(COmWindowProxy, Fire_onfocus, fire_onfocus), 0, TRUE, "COmWindowProxy::Fire_onfocus");
    } wlw note*/
}

//+----------------------------------------------------------------------------
//
//  Member:     CDocument::Fire_onpropertychange
//
//  Synopsis:   Fires the onpropertychange event, sets up the event param
//
//+----------------------------------------------------------------------------
void CDocument::Fire_onpropertychange(LPCTSTR strPropName)
{
    EVENTPARAM param(this, TRUE);

    param.SetType(_T("propertychange"));
    param.SetPropName(strPropName);

    FireEventHelper(DISPID_EVMETH_ONPROPERTYCHANGE, DISPID_EVPROP_ONPROPERTYCHANGE, (BYTE*)VTS_NONE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::Fire_PropertyChangeHelper
//
//
//+---------------------------------------------------------------------------
HRESULT CDocument::Fire_PropertyChangeHelper(DISPID dispidProperty, DWORD dwFlags)
{
    LPCTSTR pszName;
    PROPERTYDESC* ppropdesc;

    if(S_OK == FindPropDescFromDispID(dispidProperty, &ppropdesc, NULL, NULL))
    {
        Assert(ppropdesc);

        pszName = ppropdesc->pstrExposedName ? ppropdesc->pstrExposedName : ppropdesc->pstrName;

        if(pszName != NULL)
        {
            Fire_onpropertychange(pszName);
        }
    }

    return S_OK;
}

HRESULT CDocument::get_all( IHTMLElementCollection** ppDisp )
{
    return _pPrimaryMarkup->get_all(ppDisp);
}

//+-------------------------------------------------------------------------
//
//  Method:     CDoc::Getbody
//
//  Synopsis:   Get the body interface for this form
//
//--------------------------------------------------------------------------
HRESULT CDocument::get_body(IHTMLElement** ppDisp)
{
    return _pPrimaryMarkup->get_body(ppDisp);
}

//+----------------------------------------------------------------------------
//
//  Member:     activeElement, IOmDocument
//
//  Synopsis: returns a pointer to the active element (the element with the focus)
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_activeElement(IHTMLElement** ppElement)
{
    HRESULT hr=S_OK;

    if(!ppElement)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ppElement = NULL;

    if(_pElemCurrent && _pElemCurrent!=PrimaryRoot())
    {
        CElement* pTarget = _pElemCurrent;

        // all other cases fall through
        pTarget->QueryInterface(IID_IHTMLElement, (void**)ppElement);
    }

    if(*ppElement==NULL && _fVID)
    {
        hr = E_FAIL;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_title, IOmDocument
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::put_title(BSTR v)
{
    Assert(_pPrimaryMarkup);
    return _pPrimaryMarkup->put_title(v);
}

//+----------------------------------------------------------------------------
//
//  Member:     get_title, IOmDocument
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CDocument::get_title(BSTR* p)
{
    return _pPrimaryMarkup->get_title(p);
}

//+----------------------------------------------------------------------------
//
//  Member:     CDocument:get_strReadyState
//
//  Synopsis:  first implementation, this is for the OM and uses the long _readyState
//      to determine the string returned.
//
//+------------------------------------------------------------------------------
HRESULT CDocument::get_readyState(BSTR* p)
{
    return E_NOTIMPL;
}

//+-------------------------------------------------------------------------
//
// Members:     Get/SetCharset
//
// Synopsis:    Functions to get at the document's charset from the object
//              model.
//
//--------------------------------------------------------------------------
STDMETHODIMP CDocument::get_charset(BSTR* retval)
{
    TCHAR   achCharset[MAX_MIMECSET_NAME];
    HRESULT hr;

    hr = GetMlangStringFromCodePage(GetCodePage(), achCharset, ARRAYSIZE(achCharset));
    if(hr)
    {
        goto Cleanup;
    }

    hr = FormsAllocString(achCharset, retval);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CDocument::put_charset(BSTR mlangIdStr)
{
    HRESULT hr;
    CODEPAGE cp;

    hr = GetCodePageFromMlangString(mlangIdStr, &cp);
    if(hr)
    {
        goto Cleanup;
    }

    Assert(cp != CP_UNDEFINED);

    hr = SwitchCodePage(cp);
    if(hr)
    {
        goto Cleanup;
    }

    {
        OnPropertyChange(DISPID_CDoc_charset, 0);
    }

    // Clear our caches and force a replaint since codepages can have
    // distinct fonts.
    EnsureFormatCacheChange(ELEMCHNG_CLEARCACHES|ELEMCHNG_REMEASUREALLCONTENTS);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
// Protocol Identifier and Protocol Friendly Name
// Adapted from wc_html.h and wc_html.c of classic MSHTML
//
//-----------------------------------------------------------------------------
typedef struct
{
    TCHAR* szName;
    TCHAR* szRegKey;
} ProtocolRec;

static ProtocolRec ProtocolTable[] =
{
    _T("file:"),     _T("file"),
    _T("mailto:"),   _T("mailto"),
    _T("gopher://"), _T("gopher"),
    _T("ftp://"),    _T("ftp"),
    _T("http://"),   _T("http"),
    _T("https://"),  _T("https"),
    _T("news:"),     _T("news"),
    NULL, NULL
};

TCHAR* ProtocolFriendlyName(TCHAR* szURL)
{
    TCHAR szBuf[MAX_PATH];
    int i;

    if(!szURL)
    {
        return NULL;
    }

    /*LoadString(GetResourceHInst(), IDS_UNKNOWNPROTOCOL, szBuf, ARRAYSIZE(szBuf)); wlw note*/
    for(i=0; ProtocolTable[i].szName; i++)
    {
        if(_tcsnipre(ProtocolTable[i].szName, -1, szURL, -1))
        {
            break;
        }
    }
    if(ProtocolTable[i].szName)
    {
        DWORD dwLen = sizeof(szBuf);
        //DWORD dwValueType;
        HKEY hkeyProtocol;

        LONG lResult = RegOpenKeyEx(
            HKEY_CLASSES_ROOT,
            ProtocolTable[i].szRegKey,
            0,
            KEY_QUERY_VALUE,
            &hkeyProtocol);
        if(lResult != ERROR_SUCCESS)
        {
            goto Cleanup;
        }

        lResult = RegQueryValue(
            hkeyProtocol,
            NULL,
            szBuf,
            (long*)&dwLen);
        RegCloseKey(hkeyProtocol);
    }

Cleanup:
    return SysAllocString(szBuf);
}

//+----------------------------------------------------------------------------
//
// Member: get_nameProp
//
//-----------------------------------------------------------------------------
STDMETHODIMP CDocument::get_nameProp(BSTR* pName)
{
    RRETURN(SetErrorInfo(get_title(pName)));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandSupported
//
//  Synopsis:
//
//  Returns: returns true if given command (like bold) is supported
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandSupported(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    RRETURN(super::queryCommandSupported(bstrCmdId, pfRet));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandEnabled
//
//  Synopsis:
//
//  Returns: returns true if given command is currently enabled. For toolbar
//          buttons not being enabled means being grayed.
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandEnabled(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    RRETURN(super::queryCommandEnabled(bstrCmdId, pfRet));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandState
//
//  Synopsis:
//
//  Returns: returns true if given command is on. For toolbar buttons this
//          means being down. Note that a command button can be disabled
//          and also be down.
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandState(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    RRETURN(super::queryCommandState(bstrCmdId, pfRet));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandIndeterm
//
//  Synopsis:
//
//  Returns: returns true if given command is in indetermined state.
//          If this value is TRUE the value returnd by queryCommandState
//          should be ignored.
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandIndeterm(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    RRETURN(super::queryCommandIndeterm(bstrCmdId, pfRet));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandText
//
//  Synopsis:
//
//  Returns: Returns the text that describes the command (eg bold)
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandText(BSTR bstrCmdId, BSTR* pcmdText)
{
    RRETURN(super::queryCommandText(bstrCmdId, pcmdText));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandValue
//
//  Synopsis:
//
//  Returns: Returns the  command value like font name or size.
//----------------------------------------------------------------------------
HRESULT CDocument::queryCommandValue(BSTR bstrCmdId, VARIANT* pvarRet)
{
    RRETURN(super::queryCommandValue(bstrCmdId, pvarRet));
}

//+---------------------------------------------------------------------------
//
//  Member:     execCommand
//
//  Synopsis:   Executes given command
//
//  Returns:
//----------------------------------------------------------------------------
HRESULT CDocument::execCommand(BSTR bstrCmdId, VARIANT_BOOL showUI, VARIANT value, VARIANT_BOOL* pfRet)
{
    HRESULT hr;
    CDocument::CLock Lock(this, FORMLOCK_QSEXECCMD);

    hr = super::execCommand(bstrCmdId, showUI, value);

    if(pfRet)
    {
        // We return false when any error occures
        *pfRet = hr ? VB_FALSE : VB_TRUE;
        hr = S_OK;
    }

    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     execCommandShowHelp
//
//  Synopsis:
//
//  Returns:
//----------------------------------------------------------------------------
HRESULT CDocument::execCommandShowHelp(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    HRESULT hr;

    hr = super::execCommandShowHelp(bstrCmdId);

    if(pfRet != NULL)
    {
        // We return false when any error occures
        *pfRet = hr ? VB_FALSE : VB_TRUE;
        hr = S_OK;
    }

    RRETURN(SetErrorInfo(hr));
}

HRESULT CDocument::createElement(BSTR bstrTag, IHTMLElement** ppnewElem)
{
    HRESULT hr = S_OK;

    hr = PrimaryMarkup()->createElement(bstrTag, ppnewElem);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CDocument::toString(BSTR* String)
{
    RRETURN(super::toString(String));
};



HRESULT ExpandUrlWithBaseUrl(LPCTSTR pchBaseUrl, LPCTSTR pchRel, TCHAR** ppchUrl)
{
    HRESULT hr;
    TCHAR achBuf[pdlUrlLen];
    DWORD cchBuf;
    BOOL fCombine;
    DWORD dwSize;

    *ppchUrl = NULL;

    if(pchRel == NULL) // note: NULL is different from "", which is more similar to "."
    {
        hr = MemAllocString(_T(""), ppchUrl);
        goto Cleanup;
    }

    if (!pchRel)
    {
        hr = MemAllocString(pchBaseUrl, ppchUrl);
    }
    else
    {
        hr = CoInternetCombineUrl(pchBaseUrl, pchRel, URL_ESCAPE_SPACES_ONLY|URL_BROWSER_MODE, achBuf, ARRAYSIZE(achBuf), &cchBuf, 0);
        if(hr)
        {
            goto Cleanup;
        }

        if(S_OK==CoInternetQueryInfo(pchBaseUrl, QUERY_RECOMBINE, 0, &fCombine, sizeof(BOOL), &dwSize, 0)  &&
            fCombine)
        {
            TCHAR achBuf2[pdlUrlLen];
            DWORD cchBuf2;

            hr = CoInternetCombineUrl(pchBaseUrl, achBuf, URL_ESCAPE_SPACES_ONLY|URL_BROWSER_MODE, achBuf2, ARRAYSIZE(achBuf2), &cchBuf2, 0);
            if(hr)
            {
                goto Cleanup;
            }

            hr = MemAllocString(achBuf2, ppchUrl);
        }
        else
        {
            hr = MemAllocString(achBuf, ppchUrl);
        }
    }

Cleanup:
    RRETURN(hr);
}

//+====================================================================================
//
// Method: IsTridentHWND
//
// Synopsis: Helper to check to see if a given HWND belongs to a trident window
//
//------------------------------------------------------------------------------------
BOOL IsTridentHwnd(HWND hwnd)
{
    TCHAR strClassName[100];

    ::GetClassName(hwnd, strClassName, 100);

    if(StrCmpIW(strClassName, _T("Xindows_Server")) == 0)
    {
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

//+====================================================================================
//
// Method: OnChildBeginSelection
//
// Synopsis: Called by enumeration of all child windows 
//           If you are a TridentHwnd - post a WM_BEGINSELECTION message to yourself.
//           Otherwise do nothing
//
//------------------------------------------------------------------------------------
BOOL CALLBACK OnChildBeginSelection(HWND hwnd, LPARAM lParam)
{
    if(IsTridentHwnd(hwnd))
    {
        ::SendMessage(
            hwnd,
            WM_BEGINSELECTION,
            lParam,  // Cascade the Selection Type.
            0);
    }
    return TRUE;
}

HRESULT FaultInIEFeatureHelper(HWND hWnd, uCLSSPEC* pClassSpec, QUERYCONTEXT* pQuery, DWORD dwFlags)
{
    HRESULT hr = FaultInIEFeature(hWnd, pClassSpec, pQuery, dwFlags);

    RRETURN1((hr==E_ACCESSDENIED)?S_OK:hr, S_FALSE);
}


#include <ShellAPI.h>
BSTR GetFileTypeInfo(TCHAR* pchFileName)
{
    SHFILEINFO sfi;

    if(pchFileName &&
        pchFileName[0] &&
        SHGetFileInfo(
        pchFileName,
        FILE_ATTRIBUTE_NORMAL,
        &sfi,
        sizeof(sfi),
        SHGFI_TYPENAME|SHGFI_USEFILEATTRIBUTES))
    {
        return SysAllocString(sfi.szTypeName);
    }
    else
    {
        return NULL;
    }
}