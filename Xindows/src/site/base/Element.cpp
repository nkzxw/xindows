
#include "stdafx.h"
#include "Element.h"

#define _cxx_
#include "../gen/types.hdl"

#define _cxx_
#include "../gen/element.hdl"

static const GUID SID_ELEMENT_SCOPE_OBJECT = { 0x3050f408,0x98b5,0x11cf,{ 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };

#define TOTLOWER(ch)        (TCHAR((DWORD)(DWORD_PTR)CharLower((LPTSTR)(DWORD_PTR)ch)))
#define ISVALIDKEY(x)       (((x>>16)&0x00FF)<0x3a)
#define VIRTKEY_TO_SCAN     0

#define URLCOMPAT_NODEFAULT 0x00000001
#define URLCOMPAT_KEYDOWN   0x00000004

//+----------------------------------------------------------------
//
//  Function:   TagPreservationType
//
//  Synopsis:   Describes how white space is handled in this tag
//
//-----------------------------------------------------------------
WSPT TagPreservationType(ELEMENT_TAG etag)
{
    switch(etag)
    {
    case ETAG_PRE :
    case ETAG_PLAINTEXT :
    case ETAG_LISTING :
    case ETAG_XMP :
    case ETAG_TEXTAREA:
    case ETAG_INPUT:
    case ETAG_TXTSLAVE:
        return WSPT_PRESERVE;

    case ETAG_TD :
    case ETAG_TH :
    case ETAG_TC :
    case ETAG_CAPTION :
    case ETAG_BODY :
    case ETAG_ROOT :
    case ETAG_BUTTON :
        return WSPT_COLLAPSE;

    default:
        return WSPT_NEITHER;
    }
}

FOCUSSABILITY GetDefaultFocussability(ELEMENT_TAG etag)
{
    switch(etag)
    {
        // FOCUSSABILITY_TABBABLE
        // These need to be in the tab order by default
    case ETAG_A:
    case ETAG_BODY:
    case ETAG_BUTTON:
    case ETAG_IMG:
    case ETAG_INPUT:
    case ETAG_ISINDEX:
    case ETAG_OBJECT:
    case ETAG_SELECT:
    case ETAG_TEXTAREA:
        return FOCUSSABILITY_TABBABLE;

        // FOCUSSABILITY_FOCUSSABLE
        // Not recommended. Better have a good reason for each tag why
        // it it is focussable but not tabbable

        // Don't tab to applet.  This is to fix ie4 bug 41206 where the
        // VM cannot call using IOleControlSite correctly.  If we allow
        // tabbing into the VM, we can never ever tab out due to this.
        // (AnandRa 8/21/97)
    case ETAG_APPLET:

        // Should be MAYBE, but for IE4 compat (IE5 #63134)
    case ETAG_CAPTION:

        // special element
    case ETAG_DEFAULT:

        // Should be MAYBE, but for IE4 compat
    case ETAG_DIV:

        // Should be MAYBE, but for IE4 compat (IE5 63626)
    case ETAG_FIELDSET:

        // Should be MAYBE, but for IE4 compat (IE5 #62701)
    case ETAG_MARQUEE:

        // special element
    case ETAG_ROOT:

        // Should be MAYBE, but for IE4 compat
    case ETAG_SPAN:

        // Should be MAYBE, but for IE4 compat
    case ETAG_TABLE:    
    case ETAG_TD:
        return FOCUSSABILITY_FOCUSSABLE;

        // Any tag that can render/have renderable content (and does not appear in the above lists)
    case ETAG_ACRONYM:
    case ETAG_ADDRESS:
    case ETAG_B:
    case ETAG_BDO:
    case ETAG_BIG:
    case ETAG_BLINK:
    case ETAG_BLOCKQUOTE:
    case ETAG_CENTER:
    case ETAG_CITE:
    case ETAG_DD:
    case ETAG_DEL:
    case ETAG_DFN:
    case ETAG_DIR:
    case ETAG_DL:
    case ETAG_DT:
    case ETAG_EM:
    case ETAG_FONT:
    case ETAG_FORM:
    case ETAG_GENERIC:
    case ETAG_GENERIC_BUILTIN:
    case ETAG_GENERIC_LITERAL:
    case ETAG_H1:
    case ETAG_H2:
    case ETAG_H3:
    case ETAG_H4:
    case ETAG_H5:
    case ETAG_H6:
    case ETAG_HR:
    case ETAG_I:
    case ETAG_INS:
    case ETAG_KBD:

        // target is always focussable, label itself is not (unless forced by tabIndex) 
    case ETAG_LABEL:
    case ETAG_LEGEND:

    case ETAG_LI:
    case ETAG_LISTING:
    case ETAG_MENU:
    case ETAG_OL:
    case ETAG_P:
    case ETAG_PLAINTEXT:
    case ETAG_PRE:
    case ETAG_Q:
    case ETAG_RT:
    case ETAG_RUBY:
    case ETAG_S:
    case ETAG_SAMP:
    case ETAG_SMALL:
    case ETAG_STRIKE:
    case ETAG_STRONG:
    case ETAG_SUB:
    case ETAG_SUP:
    case ETAG_TBODY:
    case ETAG_TC:
    case ETAG_TFOOT:
    case ETAG_TH:
    case ETAG_THEAD:
    case ETAG_TR:
    case ETAG_TT:
    case ETAG_U:
    case ETAG_UL:
    case ETAG_VAR:
    case ETAG_XMP:
        return FOCUSSABILITY_MAYBE;

        // All the others - tags that do not ever render

        // this is a subdivision, never takes focus direcxtly
    case ETAG_AREA:

    case ETAG_BASE:
    case ETAG_BASEFONT:
    case ETAG_BR:
    case ETAG_CODE:
    case ETAG_COL:
    case ETAG_COLGROUP:
    case ETAG_COMMENT:
    case ETAG_HEAD:
    case ETAG_HTML:
    case ETAG_LINK:
    case ETAG_NEXTID:
    case ETAG_NOBR:
    case ETAG_NOEMBED:
    case ETAG_NOFRAMES:
    case ETAG_NOSCRIPT:

        // May change in future when we re-implement SELECT
    case ETAG_OPTION:

    case ETAG_PARAM:
    case ETAG_RAW_BEGINFRAG:
    case ETAG_RAW_BEGINSEL:
    case ETAG_RAW_CODEPAGE:
    case ETAG_RAW_COMMENT:
    case ETAG_RAW_DOCSIZE:
    case ETAG_RAW_ENDFRAG:
    case ETAG_RAW_ENDSEL:
    case ETAG_RAW_EOF:
    case ETAG_RAW_SOURCE:
    case ETAG_RAW_TEXT:
    case ETAG_RAW_TEXTFRAG:
    case ETAG_SCRIPT:
    case ETAG_STYLE:
    case ETAG_TITLE_ELEMENT:
    case ETAG_TITLE_TAG:
    case ETAG_WBR:
    case ETAG_UNKNOWN:
        return FOCUSSABILITY_NEVER;

    default:
        AssertSz(FALSE, "Focussability undefined for this tag");
        return FOCUSSABILITY_NEVER;
    }
}


BEGIN_TEAROFF_TABLE_(CElement, IServiceProvider)
    TEAROFF_METHOD(CElement, &QueryService, queryservice, (REFGUID guidService, REFIID riid, void** ppvObject))
END_TEAROFF_TABLE()

// Each property which has a url image has an associated internal property
// which holds the cookie for the image stored in the cache.
static const struct
{
    DISPID propID;          // url image property
    DISPID cacheID;         // internal cookie property
}
// make sure that DeleteImageCtx is modified, if any dispid's
// are added to this array
s_aryImgDispID[] =
{
    { DISPID_A_BACKGROUNDIMAGE, DISPID_A_BGURLIMGCTXCACHEINDEX },
    { DISPID_A_LISTSTYLEIMAGE,  DISPID_A_LIURLIMGCTXCACHEINDEX }
};

//+----------------------------------------------------------------------------
//
//  Member:     Inject
//
//  Synopsis:   Stuff text or HTML is various places relative to an element
//
//-----------------------------------------------------------------------------
static BOOL IsInTableThingy(CTreeNode* pNode)
{
    // See if we are between a table and its cells
    for(; pNode; pNode=pNode->Parent())
    {
        switch(pNode->Tag())
        {
        case ETAG_TABLE:
            return TRUE;

        case ETAG_TH:
        case ETAG_TC:
        case ETAG_CAPTION:
        case ETAG_TD:
            return FALSE;
        }
    }

    return FALSE;
}

//+------------------------------------------------------------------------
//
//  Member:     IElement, Get_document
//
//  Synopsis:   Returns the Idocument of this
//
//-------------------------------------------------------------------------
HRESULT CElement::get_document(IDispatch** ppIDoc)
{
    HRESULT hr = CTL_E_METHODNOTAPPLICABLE;
    CDocument* pDoc = Doc();

    if(!ppIDoc)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *ppIDoc = NULL;

    if(pDoc)
    {
        CMarkup* pMarkup;
        hr = EnsureInMarkup();
        if(hr)
        {
            goto Cleanup;
        }

        pMarkup = GetMarkup();
        Assert(pMarkup);

        hr = pMarkup->QueryInterface(IID_IHTMLDocument2, (void**)ppIDoc);
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     get_parentElement
//
//  Synopsis:   Exposes the parent element of this element.
//
//  Note:       This pays close attention to whether or not this interface is
//              based on a proxy element.  If so, use the parent of the proxy.
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::get_parentElement(IHTMLElement** ppDispParent, CTreeNode* pNode)
{
    HRESULT hr = S_OK;

    *ppDispParent = NULL;

    if(!pNode || pNode->IsDead())
    {
        pNode = GetFirstBranch();
        // Assert that either the node is not in the tree or that if it is, it is not dead
        Assert(!pNode || !pNode->IsDead());
    }

    // if still no node, we are not in the tree, return NULL
    if(!pNode)
    {
        goto Cleanup;
    }

    Assert(pNode->Element() == this);

    pNode = pNode->Parent();

    // We should never hand out the root node with a tearoff
    // (root node is only one which has null parent)
    Assert(pNode);

    // don't hand out root node
    if(pNode->Tag() == ETAG_ROOT)
    {
        goto Cleanup;
    }

    hr = pNode->GetElementInterface(IID_IHTMLElement, (void**)ppDispParent);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::CElement
//
//-------------------------------------------------------------------------
CElement::CElement(ELEMENT_TAG etag, CDocument* pDoc)
{
    __pvChain = pDoc;
    pDoc->SubAddRef();

    Assert(pDoc && pDoc->AreLookasidesClear(this, LOOKASIDE_ELEMENT_NUMBER));

    IncrementObjectCount();

    _etag = etag;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::~CElement
//
//-------------------------------------------------------------------------
CElement::~CElement()
{
    // NOTE:  Please cleanup in Passivate() if at all possible.
    //        Thread local storage can be deleted by the time this runs.
    Assert(!IsInMarkup());
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::CreateTearOffThunk
//
//----------------------------------------------------------------------------
HRESULT CElement::CreateTearOffThunk(
         void*          pvObject1,
         const void*    apfn1,
         IUnknown*      pUnkOuter,
         void**         ppvThunk,
         void*          appropdescsInVtblOrder)
{
    return ::CreateTearOffThunk(
        pvObject1,
        apfn1,
        pUnkOuter,
        ppvThunk,
        appropdescsInVtblOrder);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::CreateTearOffThunk
//
//----------------------------------------------------------------------------
HRESULT CElement::CreateTearOffThunk(
         void*          pvObject1,
         const void*    apfn1,
         IUnknown*      pUnkOuter,
         void**         ppvThunk,
         void*          pvObject2,
         void*          apfn2,
         DWORD          dwMask,
         const IID* const* apIID,
         void*          appropdescsInVtblOrder)
{
    return ::CreateTearOffThunk(
        pvObject1,
        apfn1,
        pUnkOuter,
        ppvThunk,
        pvObject2,
        apfn2,
        dwMask,
        apIID,
        appropdescsInVtblOrder);
}

//+---------------------------------------------------------------------------------
//
//  Member :    CElement::IsEqualObject()
//
//  Synopsis :  IObjectIdentity method implementation. it direct comparison of two
//              pUnks fails, then the script engines will QI for IObjectIdentity and
//              call IsEqualObject one one, passing in the other pointer.
//
//   Returns : S_OK if the Objects are the same
//             E_FAIL if the objects are different
//
//----------------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CElement::IsEqualObject(IUnknown* pUnk)
{
    HRESULT             hr = E_POINTER;
    IServiceProvider*   pISP = NULL;
    IUnknown*           pUnkTarget = NULL;
    IUnknown*           pUnkThis = NULL;

    if(!pUnk)
    {
        goto Cleanup;
    }

    hr = QueryInterface(IID_IUnknown, (void**)&pUnkThis);
    if(hr)
    {
        goto Cleanup;
    }

    if(pUnk == pUnkThis)
    {
        hr = S_OK;
        goto Cleanup;
    }

    hr = pUnk->QueryInterface(IID_IServiceProvider, (void**)&pISP);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pISP->QueryService(SID_ELEMENT_SCOPE_OBJECT, IID_IUnknown, (void**)&pUnkTarget);
    if(hr)
    {
        goto Cleanup;
    }

    hr = (pUnkThis==pUnkTarget) ? S_OK : S_FALSE;

Cleanup:
    ReleaseInterface(pUnkThis);
    ReleaseInterface(pUnkTarget);
    ReleaseInterface(pISP);
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     Passivate
//
//-------------------------------------------------------------------------
void CElement::Passivate()
{
    CDocument* pDoc = Doc();

    // If we are in a tree, the tree will keep us alive
    Assert(!IsInMarkup());

    // Make sure we don't have any pending event tasks on the view
    if(_fHasPendingEvent)
    {
        pDoc->GetView()->RemoveEventTasks(this);
    }

    // Destroy slave markup, if any
    if(HasSlaveMarkupPtr())
    {
        CMarkup* pMarkupSlave = DelSlaveMarkupPtr();

        pMarkupSlave->ClearMaster();
        pMarkupSlave->Release();
    }

    if(pDoc->_pElementOMCapture == this)
    {
        releaseCapture();
    }

    GWKillMethodCall(this, NULL, 0);

    if(_fHasImage)
    {
        ReleaseImageCtxts();
    }

    // release layout engine if any.
    if(HasLayoutPtr())
    {
        Assert(_fSite);

        CLayout* pLayout = DelLayoutPtr();

        if(!pLayout->_fDetached)
        {
            pLayout->Detach();
        }

        pLayout->Release();
    }

    if(_fHasPendingFilterTask)
    {
        Doc()->RemoveFilterTask(this);
    }

    if(_pAA)
    {
        // clear up the FiltersCollection
        if(_fHasFilterCollectionPtr)
        {
            AssertSz(FALSE, "must improve");
        }

        // Kill the cached style pointer if present.  super::passivate
        // will delete the attribute array holding it.
        if(_pAA->IsStylePtrCachePossible())
        {
            delete GetInLineStylePtr();
            delete GetRuntimeStylePtr();
        }
    }

    super::Passivate();

    // Ensure Lookaside cleanup.  Go directly to document to avoid bogus flags
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_FILTER) == NULL);
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_DATABIND) == NULL);
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_PEER) == NULL);
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_PEERMGR) == NULL);
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_ACCESSIBLE) == NULL);
    Assert(Doc()->GetLookasidePtr((DWORD*)this+LOOKASIDE_SLAVEMARKUP) == NULL);

    pDoc->SubRelease();

    DecrementObjectCount();
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CElement::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr;

    *ppv = NULL;

    switch(iid.Data1)
    {
        QI_INHERITS((IPrivateUnknown*)this, IUnknown)
        QI_HTML_TEAROFF(this, IHTMLElement, NULL)
        QI_HTML_TEAROFF(this, IHTMLElement2, NULL)
        QI_TEAROFF(this, IObjectIdentity, NULL);
        QI_TEAROFF(this, IServiceProvider, NULL);
        QI_CASE(IConnectionPointContainer)
        {
            if(iid == IID_IConnectionPointContainer)
            {
                *((IConnectionPointContainer**)ppv) = new CConnectionPointContainer(this, NULL);

                if(!*ppv)
                {
                    RRETURN(E_OUTOFMEMORY);
                }

                SetEventsShouldFire();
            }
            break;
        }
    default:
        {
            const CLASSDESC* pclassdesc = ElementDesc();

            if(iid == CLSID_CElement)
            {
                *ppv = this; // Weak ref
                return S_OK;
            }

            // Primary default interface, or the non dual
            // dispinterface return the same object -- the primary interface
            // tearoff.
            if(pclassdesc && pclassdesc->_apfnTearOff &&
                pclassdesc->_classdescBase._piidDispinterface &&
                (iid==*pclassdesc->_classdescBase._piidDispinterface ||
                DispNonDualDIID(iid)))
            {
                hr = CreateTearOffThunk(this, pclassdesc->_apfnTearOff, NULL, ppv, (void*)pclassdesc->_classdescBase._apHdlDesc->ppVtblPropDescs);
                if(hr)
                {
                    RRETURN(hr);
                }
                break;
            }
            hr = E_NOINTERFACE;
            RRETURN(hr);
            break;
        }
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    (*(IUnknown**)ppv)->AddRef();

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::PrivateAddRef, IUnknown
//
//  Synopsis:   Private unknown AddRef.
//
//-------------------------------------------------------------------------
ULONG CElement::PrivateAddRef()
{
    if(_ulRefs==1 && IsInMarkup())
    {
        Assert(GetMarkupPtr());
        GetMarkupPtr()->AddRef();
    }

    return super::PrivateAddRef();
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::PrivateRelease, IUnknown
//
//  Synopsis:   Private unknown Release.
//
//-------------------------------------------------------------------------
ULONG CElement::PrivateRelease()
{
    CMarkup* pMarkup = NULL;

    if(_ulRefs==2 && IsInMarkup())
    {
        Assert(GetMarkupPtr());
        pMarkup = GetMarkupPtr();
    }

    ULONG ret =  super::PrivateRelease();

    if(pMarkup)
    {
        pMarkup->Release();
    }

    return ret;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::PrivateEnterTree
//
//  Synopsis:   Ref counting fixup as tree entered.
//
//-------------------------------------------------------------------------
void CElement::PrivateEnterTree()
{
    Assert(IsInMarkup());
    super::PrivateAddRef();
    GetMarkupPtr()->AddRef();
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::PrivateExitTree
//
//  Synopsis:   Ref counting fixup as tree exited.
//
//-------------------------------------------------------------------------
void CElement::PrivateExitTree(CMarkup* pMarkupOld)
{
    BOOL fReleaseMarkup = _ulRefs > 1;

    Assert(!IsInMarkup());
    Assert(pMarkupOld);

    super::PrivateRelease();

    if(fReleaseMarkup)
    {
        pMarkupOld->Release();
    }
}

HRESULT CElement::OnPropertyChange(DISPID dispid, DWORD dwFlags)
{
    HRESULT             hr = S_OK;
    BOOL                fHasLayout;
    BOOL                fYieldCurrency = FALSE;
    CTreeNode*          pNode = GetFirstBranch();
    CLayout*            pLayoutOld;
    CLayout*            pLayoutNew;
    CDocument*          pDoc = Doc();
    CCollectionCache*   pCollectionCache;
    CElement*           pElemCurrent = pDoc->_pElemCurrent;

    // if this is an event property that has just been hooked up then we need to 
    // start firing events for this element. we want to do this event if this 
    // element is not yet in the tree (i.e. no pNode) so that when it IS put into
    // the tree we can fire events.
    //
    // BUGBUG a good alternative implementation is to have this flag on the AttrArray.
    // then with dword stored, we could have event level granularity and control.
    // to support this all we need to do is change the implementation fo SetEventsShouldFire()
    // adn ShouldEventsFire() to use the AA.
    if(dispid>=DISPID_EVENTS && dispid<DISPID_EVENTS+DISPID_EVPROPS_COUNT)
    {
        SetEventsShouldFire();

        // don't expose this to the outside.
        if(dispid == DISPID_EVPROP_ONATTACHEVENT)
        {
            goto Cleanup;
        }
    }

    if(!IsInMarkup() || !pNode)
    {
        goto Cleanup;
    }

    Verify(OpenView());

    Trace1("Property changed, flags:%ld\n", dwFlags);

    if(DISPID_A_BEHAVIOR==dispid || DISPID_CElement_className==dispid || DISPID_UNKNOWN==dispid)
    {
        /*if(DISPID_A_BEHAVIOR == dispid)
        {
            pDoc->SetPeersPossible();
        }

        hr = ProcessPeerTask(PEERTASK_RECOMPUTEBEHAVIORS);
        if(hr)
        {
            goto Cleanup;
        } wlw note*/
    }

    pLayoutOld = GetUpdatedLayout();

    // some changes invalidate collections
    if(dwFlags & ELEMCHNG_UPDATECOLLECTION)
    {
        // BUGBUG rgardner, for now Inval all the collections, whether they are filtered on property values
        // or not. We should tweak the PDL code to indicate which collections should be invaled, or do
        // this intelligently through some other mechanism, we should tweak this
        // when we remove the all collection.
        Assert(IsInMarkup());

        pCollectionCache = GetMarkup()->CollectionCache();
        if(pCollectionCache)
        {
            pCollectionCache->InvalidateAllSmartCollections();
        }

        // Clear this flag: exclusive or
        dwFlags ^= ELEMCHNG_UPDATECOLLECTION;
    }

    switch(dispid)
    {
    case DISPID_A_BACKGROUNDIMAGE:
    case DISPID_A_LISTSTYLEIMAGE:
        // Release any dispid's which hold image contexts
        DeleteImageCtx(dispid);
        break;
    }

    // Make sure our caches are up to date - if we haven't already
    hr = EnsureFormatCacheChange(dwFlags);
    if(hr)
    {
        goto Cleanup;
    }

    fHasLayout = NeedsLayout();
    pLayoutNew = GetUpdatedNearestLayout();

    if(!pLayoutNew || !pLayoutNew->ElementOwner()->IsInMarkup())
    {
        goto Cleanup;
    }

    switch(dispid)
    {
    case DISPID_A_POSITION:
        //BUGBUG tomfakes  EnsureInMarkup?
        Assert(IsInMarkup());

        pCollectionCache = GetMarkup()->CollectionCache();
        if(pCollectionCache)
        {
            pCollectionCache->InvalidateItem(CMarkup::REGION_COLLECTION);
        }

        SendNotification(NTYPE_ZPARENT_CHANGE);
        break;

    case STDPROPID_XOBJ_LEFT:
    case STDPROPID_XOBJ_RIGHT:
        {
            const CFancyFormat* pFF = GetFirstBranch()->GetFancyFormat();

            // if width is auto, left & right control the width of the
            // element if it is absolute, so fire a resize instead of
            // reposition
            if(pFF->IsAbsolute() && pFF->IsWidthAuto())
            {
                dwFlags |= ELEMCHNG_SIZECHANGED;
            }
            else if(!pFF->IsPositionStatic())
            {
                RepositionElement();
            }

            if(dispid == STDPROPID_XOBJ_LEFT)
            {
                FireOnChanged(DISPID_IHTMLSTYLE_PIXELLEFT);
                FireOnChanged(DISPID_IHTMLSTYLE_POSLEFT);
            }
            else
            {
                FireOnChanged(DISPID_IHTMLSTYLE2_PIXELRIGHT);
                FireOnChanged(DISPID_IHTMLSTYLE2_POSRIGHT);
            }
        }
        break;

    case STDPROPID_XOBJ_TOP:
    case STDPROPID_XOBJ_BOTTOM:
        {
            const CFancyFormat* pFF = GetFirstBranch()->GetFancyFormat();

            // if height is auto, and the bottom is not auto then the size
            // of the element if absolute can change based on it's top & bottom.
            // So, fire a resize instead of reposition. If bottom is auto,
            // then the element is sized to content.
            if(pFF->IsAbsolute() && pFF->IsHeightAuto()
                && (dispid!=STDPROPID_XOBJ_TOP || !pFF->IsBottomAuto()))
            {
                dwFlags |= ELEMCHNG_SIZECHANGED;
            }
            else if(!pFF->IsPositionStatic())
            {
                RepositionElement();
            }

            if(dispid == STDPROPID_XOBJ_TOP)
            {
                FireOnChanged(DISPID_IHTMLSTYLE_PIXELTOP);
                FireOnChanged(DISPID_IHTMLSTYLE_POSTOP);
            }
            else
            {
                FireOnChanged(DISPID_IHTMLSTYLE2_PIXELBOTTOM);
                FireOnChanged(DISPID_IHTMLSTYLE2_POSBOTTOM);
            }
        }
        break;

    case STDPROPID_XOBJ_WIDTH:
        FireOnChanged(DISPID_IHTMLSTYLE_PIXELWIDTH);
        FireOnChanged(DISPID_IHTMLSTYLE_POSWIDTH);
        break;

    case STDPROPID_XOBJ_HEIGHT:
        FireOnChanged(DISPID_IHTMLSTYLE_PIXELHEIGHT);
        FireOnChanged(DISPID_IHTMLSTYLE_POSHEIGHT);
        break;

    case DISPID_A_ZINDEX:
        if(!IsPositionStatic())
        {
            ZChangeElement();
        }

        pDoc->FixZOrder();
        break;


    case DISPID_CElement_accessKey:
        // if accessKey property is changed, always insert it into
        // CDoc::_aryAccessKeyItems, we are not considering removing CElements
        // from CDoc::_aryAccessKeyItems as of now
        if(pDoc->_aryAccessKeyItems.Find(this) < 0)
        {
            hr = pDoc->InsertAccessKeyItem(this);
            if(hr)
            {
                goto Cleanup;
            }
        }
        break;

    case DISPID_CElement_tabIndex:
        hr = OnTabIndexChange();
        break;

    case DISPID_CElement_disabled:
        {
            BOOL        fEnabled    = !GetAAdisabled();
            CElement*   pNewDefault = 0;
            CElement*   pOldDefault = 0;
            CElement*   pSavDefault = 0;

            // we should not be the default button before becoming enabled
            Assert(!(fEnabled && _fDefault));

            if(!_fDefault && fEnabled)
            {
                // this is not the previous default button
                // look for it
                pOldDefault = FindDefaultElem(TRUE);
            }

            if(pDoc->_pElemCurrent==this && !fEnabled)
            {
                // this is the case where the button disables itself
                pOldDefault = this;
            }

            // if this element act like a button
            // becomes disabled or enabled
            // we need to make sure this is recorded in the cached
            // default element of the doc or form
            if(TestClassFlag(CElement::ELEMENTDESC_DEFAULT))
            {
                // try to find a new default
                // set _fDefault to FALSE in order to avoid
                // FindDefaultLayout returning this site again.
                // because FindDefaultLayout will use _fDefault
                _fDefault = FALSE;
                pSavDefault = FindDefaultElem(TRUE, TRUE);
                if(pSavDefault==this || !fEnabled)
                {
                    if(pSavDefault)
                    {
                        pSavDefault->_fDefault = TRUE;
                    }
                    pDoc->_pElemDefault = pSavDefault;
                }
            }
            if(!pDoc->_pElemCurrent->_fActsLikeButton || _fDefault)
            {
                _fDefault = FALSE;
                pNewDefault = pSavDefault ? pSavDefault : FindDefaultElem(TRUE, TRUE);
            }

            // if this was the default, and now becoming disabled
            _fDefault = FALSE;

            if(pNewDefault != pOldDefault)
            {
                if(pOldDefault)
                {
                    CNotification nf;

                    nf.AmbientPropChange(pOldDefault, (void*)DISPID_AMBIENT_DISPLAYASDEFAULT);

                    // refresh the old button
                    pOldDefault->_fDefault = FALSE;
                    pOldDefault->Notify(&nf);
                    pOldDefault->Invalidate();
                }

                if(pNewDefault)
                {
                    CNotification nf;

                    Assert(pNewDefault->_fActsLikeButton);
                    nf.AmbientPropChange(pNewDefault, (void*)DISPID_AMBIENT_DISPLAYASDEFAULT);
                    pNewDefault->_fDefault = TRUE;
                    pNewDefault->Notify(&nf);
                    pNewDefault->Invalidate();
                }
            }
        }
        break;

    case DISPID_A_VISIBILITY:
        // Notify element and all descendents of the change
        SendNotification(NTYPE_VISIBILITY_CHANGE);

        // If the element is being hidden, ensure it and any descendent which inherits visibility are the current element
        // (Do this by, within this routine, pretending that this element is the current element)
        fYieldCurrency = !!pNode->GetCharFormat()->IsVisibilityHidden();

        if(fYieldCurrency
            && pElemCurrent->GetFirstBranch()->SearchBranchToRootForScope(this)
            && pElemCurrent->GetFirstBranch()->GetCharFormat()->IsVisibilityHidden())
        {
            pElemCurrent = this;
        }
        break;

    case DISPID_A_DISPLAY:
        // If the element is being hidden, ensure it and none of its descendents are the current element
        // (Do this by, within this routine, pretending that this element is the current element)
        if(pNode->GetCharFormat()->IsDisplayNone())
        {
            fYieldCurrency = TRUE;

            if(pElemCurrent->GetFirstBranch()->SearchBranchToRootForScope(this))
            {
                pElemCurrent = this;
            }
        }
        break;

    case DISPID_A_MARGINTOP:
    case DISPID_A_MARGINLEFT:
    case DISPID_A_MARGINRIGHT:
    case DISPID_A_MARGINBOTTOM:
        if(Tag() == ETAG_BODY)
        {
            dwFlags |= ELEMCHNG_REMEASURECONTENTS;
            dwFlags &= ~(ELEMCHNG_SIZECHANGED | ELEMCHNG_REMEASUREINPARENT);
        }
        break;
    case DISPID_A_CLIP:
    case DISPID_A_CLIPRECTTOP:
    case DISPID_A_CLIPRECTRIGHT:
    case DISPID_A_CLIPRECTBOTTOM:
    case DISPID_A_CLIPRECTLEFT:
        if(fHasLayout)
        {
            CDispNode* pDispNode = pLayoutNew->GetElementDispNode(this);
            if(pDispNode)
            {
                if(pDispNode->HasUserClip())
                {
                    CSize size;

                    pLayoutNew->GetSize(&size);
                    pLayoutNew->SizeDispNodeUserClip(&(pDoc->_dci), size);
                }
                // we need to create a display node that can have user
                // clip information, and ResizeElement will force this
                // node to be created.  Someday, we might be able to morph
                // the existing display node for greater efficiency.
                else
                {
                    ResizeElement();
                }
            }
        }
        break;
    }

    // Notify the layout of the property change, layout fixes up
    // its display node to handle visibility/background changes.
    if(fHasLayout)
    {
        pLayoutNew->OnPropertyChange(dispid, dwFlags);
    }

    if((pLayoutOld && !fHasLayout) || (!pLayoutOld && fHasLayout))
    {
        if(this == pDoc->_pElemCurrent)
        {
            pLayoutNew->ElementOwner()->BecomeCurrent(pDoc->_lSubCurrent);
        }

        dwFlags |= ELEMCHNG_REMEASUREINPARENT;
        dwFlags &= ~ELEMCHNG_SIZECHANGED;
    }

    if(dwFlags & (ELEMCHNG_REMEASUREINPARENT|ELEMCHNG_SIZECHANGED))
    {
        MinMaxElement();
    }

    if(dwFlags & ELEMCHNG_REMEASUREINPARENT)
    {
        RemeasureInParentContext();
    }
    else if(dwFlags&ELEMCHNG_SIZECHANGED
        || (!fHasLayout && dwFlags&ELEMCHNG_RESIZENONSITESONLY))
    {
        ResizeElement();
    }

    if(dwFlags & (ELEMCHNG_REMEASURECONTENTS|ELEMCHNG_REMEASUREALLCONTENTS))
    {
        RemeasureElement(dwFlags&ELEMCHNG_REMEASUREALLCONTENTS ? NFLAGS_FORCE : 0);
    }

    // we need to send the display change notification after sending
    // the RemeasureInParent notification. RemeasureInParent marks the
    // ancestors dirty, therefore any ZParentChange notifications fired
    // from DisplayChange are queued up until the ZParent is calced.
    if(dispid==DISPID_A_DISPLAY || dispid==DISPID_CElement_className || dispid==DISPID_UNKNOWN)
    {
        CNotification nf;

        nf.DisplayChange(this);

        SendNotification(&nf);
    }

    if(this == pElemCurrent)
    {
        if(!IsEnabled() || fYieldCurrency)
        {
            pNode->Parent()->Element()->BubbleBecomeCurrentAndActive();
        }
        else if(dispid==DISPID_A_VISIBILITY || dispid==DISPID_A_DISPLAY)
        {
            pDoc->GetView()->SetFocus(pDoc->_pElemCurrent, pDoc->_lSubCurrent);
        }
    }

    // Invalidation on z-index changes will be handled by FixZOrder in CSite
    // unless the change is on a non-site
    if(!(dispid==DISPID_A_ZINDEX && fHasLayout))
    {
        if((dwFlags&(ELEMCHNG_SITEREDRAW|ELEMCHNG_CLEARCACHES|ELEMCHNG_CLEARFF))
            && !(dwFlags&ELEMCHNG_SITEPOSITION && fHasLayout))
        {
            Invalidate();

            // Invalidate() sends a notification ot the parent, however in OPC
            // descendant elemetns may inherit what has just changed, so we need
            // a notification that goes down to the positioned children
            // so that they know to invalidate.  This is necessary since
            // when a property is changed they may have a change in an 
            // inherited value that they need to display
            //
            //  if we wind up needing this notification fired from other 
            // places, consider moving it into InvalidateElement()
            SendNotification(NTYPE_ELEMENT_INVAL_Z_DESCENDANTS);
        }
    }

    // once we start firing events, the pLayout Old is not reliable since script can
    // change the page.
    FireOnChanged(dispid);

    // fire the onpropertychange script event, but only if it is possible
    //   that someone is actually listeing. otherwise, don't waste the time.
    if(ShouldFireEvents() && DISPID_UNKNOWN!=dispid)
    {
        hr = Fire_PropertyChangeHelper(dispid, dwFlags);
        if(hr)
        {
            goto Cleanup;
        }
    }

    // Accessibility state change check and event firing
    // wlw note
    //if((dwFlags&ELEMCHNG_ACCESSIBILITY) && HasAccObjPtr())
    //{
    //    // fire accesibility state change event.
    //    hr = pDoc->FireAccesibilityEvents(NULL, GetSourceIndex());
    //}

    dwFlags &= ~(ELEMCHNG_CLEARCACHES|ELEMCHNG_CLEARFF);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::put_dir
//
//  Synopsis: Object model entry point to put the dir property.
//
//  Arguments:
//            [v]  - BSTR containing the new property value (ltr or rtl)
//
//  Returns:  S_OK                  - successful
//            DISP_E_MEMBERNOTFOUND - this element doesn't support this
//                                    property
//            E_INVALIDARG          - the argument is not one of the enum string
//                                    values (ltr or rtl)
//
//----------------------------------------------------------------------------
STDMETHODIMP CElement::put_dir(BSTR v)
{
    HRESULT hr;

    hr = s_propdescCElementdir.b.SetEnumStringProperty(
        v, this, (CVoid*)(void*)(GetAttrArray()));

    RRETURN(hr);
}

//+-------------------------------------------------------------------
//      Member : get_style
//
//      Synopsis : for use by IDispatch Invoke to retrieve the
//      IHTMLStyle for this object's inline style.  Get it from
//      CStyle. If none currently exists, make one.
//+-------------------------------------------------------------------
HRESULT CElement::get_style(IHTMLStyle** ppISTYLE)
{
    HRESULT hr = S_OK;
    CStyle* pStyleInline = NULL;
    *ppISTYLE = NULL;

    hr = GetStyleObject(&pStyleInline);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pStyleInline->QueryInterface(IID_IHTMLStyle, (LPVOID*)ppISTYLE);

Cleanup:
    RRETURN(SetErrorInfoPGet(hr, STDPROPID_XOBJ_STYLE));
}

//+---------------------------------------------------------------------------
//
//  Member: CElement::GetInfo
//
//  Params: [gi]: The GETINFO enumeration.
//
//  Descr:  Returns the information requested in the enum
//
//----------------------------------------------------------------------------
DWORD CElement::GetInfo(GETINFO gi)
{
    switch(gi)
    {
    case GETINFO_ISCOMPLETED:
        return TRUE;

    case GETINFO_HISTORYCODE:
        // BUGBUG (MohanB) Encode _uIndex also ? Need to reserve space
        // for GetAAtype() though, for input controls.
        return Tag();
    }

    return 0;
}

//+------------------------------------------------------------------------
//
//  Member:     GettagName
//
//  Synopsis:   Returns the tag name of the current node.
//
//-------------------------------------------------------------------------
HRESULT CElement::get_tagName(BSTR* pTagName)
{
    *pTagName = SysAllocString(TagName());

    RRETURN(SetErrorInfoPGet(*pTagName?S_OK:E_OUTOFMEMORY, DISPID_CElement_tagName));
}

//+------------------------------------------------------------------------
//
//  Member:     GettagName
//
//  Synopsis:   Returns the tag name of the current node.
//
//-------------------------------------------------------------------------
HRESULT CElement::GettagName(BSTR* pTagName)
{
    *pTagName = SysAllocString(TagName());

    RRETURN(SetErrorInfoPGet(*pTagName?S_OK:E_OUTOFMEMORY, DISPID_CElement_tagName));
}

//+----------------------------------------------------------------------------
//
//  Method : CElement :: QueryService
//
//  Synopsis : IServiceProvider methoid Implementaion.
//
//-----------------------------------------------------------------------------
HRESULT CElement::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
{
    HRESULT hr = E_POINTER;

    if(!ppvObject)
    {
        goto Cleanup;
    }

    if(IsEqualGUID(guidService, SID_ELEMENT_SCOPE_OBJECT))
    {
        hr = QueryInterface(riid, ppvObject);
    }
    else
    {
        hr = super::QueryService(guidService, riid, ppvObject);
        if(hr == E_NOINTERFACE)
        {
            hr = Doc()->QueryService(guidService, riid, ppvObject);
        }
    }

Cleanup:
    RRETURN1(hr, E_NOINTERFACE);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetAtomTable, virtual override from CBase
//
//-------------------------------------------------------------------------
CAtomTable* CElement::GetAtomTable(BOOL* pfExpando)
{
    CAtomTable* pat = NULL;
    CDocument*  pDoc;

    pDoc = Doc();
    if(pDoc)
    {
        pat = &(pDoc->_AtomTable);
        if(pfExpando)
        {
            *pfExpando = pDoc->_fExpando;
        }
    }

    // We should have an atom table otherwise there's a problem.
    Assert(pat && "Element is not associated with a CDocument.");

    return pat;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::Init
//
//  Synopsis:   Do any element initialization here, called after the element is
//              created from CreateElement()
//
//----------------------------------------------------------------------------
HRESULT CElement::Init()
{
    HRESULT hr;

    hr = super::Init();

    if(hr)
    {
        goto Cleanup;
    }

    _fLayoutAlwaysValid = _fSite || TestClassFlag(ELEMENTDESC_NOLAYOUT);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::Init2
//
//  Synopsis:   Perform any element level initialization
//
//-------------------------------------------------------------------------
HRESULT CElement::Init2(CInit2Context* pContext)
{
    HRESULT     hr = S_OK;
    LPCTSTR     pch;
    CDocument*  pDoc = Doc();
    CAttrArray* pAAInline;

    pch = GetIdentifier();
    if(pch)
    {
        hr = pDoc->_AtomTable.AddNameToAtomTable(pch, NULL);
        if(hr)
        {
            goto Cleanup;
        }
    }

    // behaviors support
    // BUGBUG (alexz) see if it is possible to make it without parsing inline styles here
    pAAInline = GetInLineStyleAttrArray();
    if(pAAInline && pAAInline->Find(DISPID_A_BEHAVIOR))
    {
        pDoc->SetPeersPossible();
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::InitAttrBag
//
//  Synopsis:   Fetch values from CHtmTag and put into the bag
//
//  Arguments:  pht : parsed attributes in text format
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CElement::InitAttrBag(CHtmTag* pht)
{
    int             i;
    CHtmTag::CAttr* pattr = NULL;
    HRESULT         hr = S_OK;
    const PROPERTYDESC* ppropdesc;
    CBase*          pBaseObj;
    WORD            wMaxstrlen = 0;
    TCHAR           chOld = _T('\0');
    CDocument*      pDoc = Doc();
    THREADSTATE*    pts = GetThreadState();

    if(*GetAttrArray())
    {
        // InitAttrBag must have already been called.
        return S_OK;
    }

    // Loop over all attr pairs in the tag, and see if their BYTE in the known-attr array
    // is set; if not, add the attr and any val to the CaryUnknownAttrs in the attr bag
    // (create one if needed).
    for(i=pht?pht->GetAttrCount():0; --i>=0; )
    {
        pattr = pht->GetAttr(i);

        if(!pattr->_pchName)
        {
            continue;
        }

        if((ppropdesc=FindPropDescForName(pattr->_pchName)) != NULL)
        {
            // Allow some elements to redirect to another attr array
            pBaseObj = GetBaseObjectFor(ppropdesc->GetDispid());

            if(!pBaseObj)
            {
                continue;
            }

            wMaxstrlen = (ppropdesc->GetBasicPropParams()->wMaxstrlen==pdlNoLimit) ? 0 :
                (ppropdesc->GetBasicPropParams()->wMaxstrlen ? ppropdesc->GetBasicPropParams()->wMaxstrlen : DEFAULT_ATTR_SIZE);

            if(wMaxstrlen && pattr->_pchVal && _tcslen(pattr->_pchVal)>wMaxstrlen)
            {
                chOld = pattr->_pchVal[wMaxstrlen];
                pattr->_pchVal[wMaxstrlen] = _T('\0');
            }
            hr = ppropdesc->HandleLoadFromHTMLString(pBaseObj, pattr->_pchVal);

            if(hr)
            {
                // Create an "unknown" attribute containing the original string from the HTML
                // SetString with fIsUnkown set to TRUE
                if(chOld)
                {
                    pattr->_pchVal[wMaxstrlen] = chOld;
                    chOld = 0;
                }
                hr = CAttrArray::SetString(pBaseObj->GetAttrArray(), ppropdesc, pattr->_pchVal, TRUE, CAttrValue::AA_Extra_DefaultValue);
            }

            // If the parameter was invalid, value will get set to default &
            // parameter will go into the unknown bag
            if(!hr)
            {
                if(ppropdesc->GetDispid() == DISPID_A_BACKGROUNDIMAGE)
                {
                    // Fork off an early download for background images
                    LPCTSTR lpszURL;
                    if(!(*(pBaseObj->GetAttrArray()))->FindString(DISPID_A_BACKGROUNDIMAGE, &lpszURL))
                    {
                        LONG lCookie;

                        if(GetImageUrlCookie(lpszURL, &lCookie) == S_OK)
                        {
                            AddImgCtx(DISPID_A_BGURLIMGCTXCACHEINDEX, lCookie);
                        }
                    }
                }
            }
            else if(hr == E_OUTOFMEMORY)
            {
                goto Cleanup;
            }
        }
        else
        {
            DISPID expandoDISPID;

            // Create an expando
            hr = AddExpando(pattr->_pchName, &expandoDISPID);

            if(hr == DISP_E_MEMBERNOTFOUND)
            {
                hr = S_OK;
                continue; // Expando not turned on
            }
            if(hr)
            {
                goto Cleanup;
            }

            if(TestClassFlag(ELEMENTDESC_OLESITE))
            {
                expandoDISPID = expandoDISPID - DISPID_EXPANDO_BASE +
                    DISPID_ACTIVEX_EXPANDO_BASE;
            }

            // Note that we always store expandos in the current object - we never redirect them
            hr = AddString(expandoDISPID, pattr->_pchVal, CAttrValue::AA_Expando);
            if(hr)
            {
                goto Cleanup;
            }

            // if begins with "on", this can be a peer registered event - need to store line/offset numbers
            if(0 == StrCmpNIC(_T("on"), pattr->_pchName, 2))
            {
                hr = StoreLineAndOffsetInfo(this, expandoDISPID, pattr->_ulLine, pattr->_ulOffset);
                if(hr)
                {
                    goto Cleanup;
                }
            }
        }

        if(chOld)
        {
            pattr->_pchVal[wMaxstrlen] = chOld;
            chOld = _T('\0');
        }
    }

Cleanup:
    if(chOld)
    {
        pattr->_pchVal[wMaxstrlen] = chOld;
    }

    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::MergeAttrBag
//
//  Synopsis:   Add any value from CHtmTag that are not already present
//              in the attrbag
//
//              Note: currently, expandos are not merged.
//
//  Arguments:  pht : parsed attributes in text format
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CElement::MergeAttrBag(CHtmTag* pht)
{
    int                 i;
    CHtmTag::CAttr*     pattr = NULL;
    HRESULT             hr = S_OK;
    const PROPERTYDESC* ppropdesc;
    CBase*              pBaseObj;
    WORD                wMaxstrlen = 0;
    TCHAR               chOld = _T('\0');
    CDocument*          pDoc = Doc();

    // Loop over all attr pairs in the tag, and see if their BYTE in the known-attr array
    // is set; if not, add the attr and any val to the CaryUnknownAttrs in the attr bag
    // (create one if needed).
    for(i=pht?pht->GetAttrCount():0; --i>=0; )
    {
        pattr = pht->GetAttr(i);

        if(!pattr->_pchName)
        {
            continue;
        }

        if((ppropdesc=FindPropDescForName(pattr->_pchName)) != NULL)
        {
            // Allow some elements to redirect to another attr array
            pBaseObj = GetBaseObjectFor(ppropdesc->GetDispid());

            // Only add the attribute if it has not been previously defined
            // style attribute requires special handling
            if( !pBaseObj
                || (AA_IDX_UNKNOWN!=pBaseObj->FindAAIndex(ppropdesc->GetDispid(), CAttrValue::AA_Attribute))
                || (AA_IDX_UNKNOWN!=pBaseObj->FindAAIndex(ppropdesc->GetDispid(), CAttrValue::AA_UnknownAttr))
                || (AA_IDX_UNKNOWN!=pBaseObj->FindAAIndex(ppropdesc->GetDispid(), CAttrValue::AA_Internal))
                || (AA_IDX_UNKNOWN!=pBaseObj->FindAAIndex(ppropdesc->GetDispid(), CAttrValue::AA_AttrArray))
                || (ppropdesc==(PROPERTYDESC*)&s_propdescCElementstyle_Str && AA_IDX_UNKNOWN!=pBaseObj->FindAAIndex(DISPID_INTERNAL_INLINESTYLEAA, CAttrValue::AA_AttrArray)))
            {
                continue;
            }

            wMaxstrlen = (ppropdesc->GetBasicPropParams()->wMaxstrlen==pdlNoLimit) ? 0 :
                (ppropdesc->GetBasicPropParams()->wMaxstrlen?ppropdesc->GetBasicPropParams()->wMaxstrlen:DEFAULT_ATTR_SIZE);

            if(wMaxstrlen && pattr->_pchVal && _tcslen(pattr->_pchVal)>wMaxstrlen)
            {
                chOld = pattr->_pchVal[wMaxstrlen];
                pattr->_pchVal[wMaxstrlen] = _T('\0');
            }
            hr = ppropdesc->HandleMergeFromHTMLString(pBaseObj, pattr->_pchVal);

            if(hr)
            {
                // Create an "unknown" attribute containing the original string from the HTML
                // SetString with fIsUnkown set to TRUE
                hr = CAttrArray::SetString(pBaseObj->GetAttrArray(), ppropdesc,
                    pattr->_pchVal, TRUE, CAttrValue::AA_Extra_DefaultValue);
            }

            // If the parameter was invalid, value will get set to default &
            // parameter will go into the unknown bag
            if(!hr)
            {
                if(ppropdesc->GetDispid() == DISPID_A_BACKGROUNDIMAGE)
                {
                    // Fork off an early download for background images
                    LPCTSTR lpszURL;
                    if(!(*(pBaseObj->GetAttrArray()))->FindString(DISPID_A_BACKGROUNDIMAGE, &lpszURL))
                    {
                        LONG lCookie;

                        if(GetImageUrlCookie(lpszURL, &lCookie) == S_OK)
                        {
                            AddImgCtx(DISPID_A_BGURLIMGCTXCACHEINDEX, lCookie);
                        }
                    }
                }
            }
            else if(hr == E_OUTOFMEMORY)
            {
                goto Cleanup;
            }
        }

        if(chOld)
        {
            pattr->_pchVal[wMaxstrlen] = chOld;
            chOld = _T('\0');
        }
    }

Cleanup:
    if(chOld)
    {
        pattr->_pchVal[wMaxstrlen] = chOld;
    }

    RRETURN(hr);
}

void CElement::Notify(CNotification* pnf)
{
    Assert(pnf);

    /*if(HasPeerHolder())
    {
        GetPeerHolder()->OnElementNotification(pnf);
    } wlw note*/

    switch(pnf->Type())
    {
    case NTYPE_AMBIENT_PROP_CHANGE:
        Invalidate();
        break;

    case NTYPE_COMMAND:
        break;

    case NTYPE_ELEMENT_EXITTREE_1:
        /*if(HasPeerHolder() || HasAccObjPtr())
        {
            pnf->SetData(pnf->DataAsDWORD()|EXITTREE_DELAYRELEASENEEDED);
        } wlw note*/
        ExitTree(pnf->DataAsDWORD());

        break;

    case NTYPE_RELEASE_EXTERNAL_OBJECTS:
        /*if(HasPeerHolder())
        {
            DelPeerHolder()->Release(); // delete the ptr and release peer holder
        } wlw note*/
        break;

    case NTYPE_RECOMPUTE_BEHAVIOR:
        /*ProcessPeerTask(PEERTASK_RECOMPUTEBEHAVIORS); wlw note*/
        break;

    case NTYPE_ELEMENT_ENTERTREE:
        EnterTree();
        break;

    case NTYPE_VISIBILITY_CHANGE:
        // BUGBUG: Sanitize this by making official throughout the code that layout-like notifications
        //         targeted at positioned elements without layouts are re-directed to the nearest layout.
        //         This redirection should be done in CMarkup::NotifyElement rather than ad hoc. (brendand)
        Assert(GetFirstBranch()->GetCascadedposition() == stylePositionrelative);
        Assert(!HasLayout());
        {
            CLayout* pLayout = GetCurNearestLayout();

            if(pLayout)
            {
                pLayout->Notify(pnf);
            }
        }
        break;

    case NTYPE_ELEMENT_INVAL_Z_DESCENDANTS:
        {
            // bugbug (carled) this notification should really be unnecessary. what it
            // (and the one above) are crying for is a Notifcation for OnPropertyChange
            // so that descendants can take specific action.  More generally, there are
            // other OPP things like VoidCachedInfo that are duplicating the Notification
            // logic (walking the subtree) which could get rolled into this.  unifiying 
            // these things could go a long way to streamlining the code, and preventing
            // inconsistencies and lots of workarounds. (like this notification)
            //
            // This notification is targeted at positioned elements which are children of 
            // the element that sent it. We do not want to delegate to the nearest Layout
            // since if they are positioned, they have thier own. Since the source element2
            // may not have a layout we have to use the element tree for the routing.
            CLayout* pLayout = GetUpdatedLayout();

            if(pLayout && IsPositioned())
            {
                pLayout->Notify(pnf);
            }
        }
        break;
    }

    return;
}

HRESULT CElement::EnterTree()
{
    HRESULT     hr = S_OK;
    CDocument*  pDoc = Doc();

    Assert(IsInMarkup());

    /*hr = ProcessPeerTask(PEERTASK_ENTERTREE_UNSTABLE);
    if(hr)
    {
        goto Cleanup;
    } wlw note*/

    if(IsInPrimaryMarkup())
    {
        // These things only make sense on elements which are in the main
        // tree of the document, not on any old random tree that's been
        // created.
        pDoc->OnElementEnter(this);
    }

    if(HasLayout())
    {
        GetCurLayout()->Init();
    }

    if(GetMarkup()->CollectionCache())
    {
        OnEnterExitInvalidateCollections(FALSE);
    }

Cleanup:
    RRETURN(hr);
}

void CElement::ExitTree(DWORD dwExitFlags)
{
    CMarkup*    pMarkup = GetMarkup();
    CDocument*  pDoc = pMarkup->Doc();
    CLayout*    pLayout;

    Assert(IsInMarkup());

    pDoc->OnElementExit(this, dwExitFlags);
    if(!(dwExitFlags & EXITTREE_DESTROY))
    {
        pLayout = GetUpdatedLayout();

        if(pLayout && pLayout->_fAdorned)
        {
            // We now send the notification for adorned elements - or 
            // any LayoutElement leaving the tree.
            //
            // BUGBUG marka - make this work via the adorner telling the selection manager
            // to change state. Also make sure _fAdorned is cleared on removal of the Adorner
            IUnknown* pUnknown = NULL;
            this->QueryInterface(IID_IUnknown, (void**)&pUnknown);
            pDoc->NotifySelection(SELECT_NOTIFY_EXIT_TREE, pUnknown);
            ReleaseInterface(pUnknown);
        }
        pLayout = NULL;
    }

    if(pDoc->GetView()->HasAdorners())
    {
        Verify(pDoc->GetView()->OpenView());
        pDoc->GetView()->RemoveAdorners(this);
    }

    if(pMarkup->HasCollectionCache())
    {
        OnEnterExitInvalidateCollections(FALSE);
    }

    // Filters do not survive tree changes
    if(_fHasPendingFilterTask)
    {
        pDoc->RemoveFilterTask(this);
        Assert(!_fHasPendingFilterTask);
    }

    // If we don't have any lookasides we don't have to do *any* of the following tests!
    if(_fHasLookasidePtr)
    {
        // wlw note
        //if(HasPeerHolder() && !(dwExitFlags&EXITTREE_PASSIVATEPENDING))
        //{
        //    IGNORE_HR(ProcessPeerTask(PEERTASK_EXITTREE_UNSTABLE));
        //}

        // wlw note
        //if(HasPeerMgr())
        //{
        //    CPeerMgr* pPeerMgr = GetPeerMgr();
        //    pPeerMgr->OnExitTree();
        //}
    }

    if(HasLayoutPtr())
    {
        CLayout* pLayout;

        GetCurLayout()->OnExitTree();

        if(!_fSite)
        {
            pLayout = DelLayoutPtr();

            pLayout->Detach();

            Assert(pLayout->_fDetached);

            pLayout->Release();
        }
    }
}

CBase* CElement::GetOmWindow(void)
{
    CDocument* pDoc = Doc();

    /*pDoc->EnsureOmWindow();
    return pDoc->_pOmWindow; wlw note*/
    return NULL;
}

// N.B. (CARLED) in order for body.onfoo = new Function("") to work you
// need to add the BodyAliasForWindow flag in the PDL files for each of
// these properties
CBase* CElement::GetBaseObjectFor(DISPID dispID)
{
    return (((ETAG_BODY == Tag())&&
        (dispID == DISPID_EVPROP_ONLOAD ||
        dispID==DISPID_EVPROP_ONUNLOAD ||
        dispID==DISPID_EVPROP_ONRESIZE ||
        dispID==DISPID_EVPROP_ONSCROLL ||
        dispID==DISPID_EVPROP_ONBEFOREUNLOAD ||
        dispID==DISPID_EVPROP_ONHELP ||
        dispID==DISPID_EVPROP_ONBLUR ||
        dispID==DISPID_EVPROP_ONFOCUS ||
        dispID==DISPID_EVPROP_ONBEFOREPRINT ||
        dispID==DISPID_EVPROP_ONAFTERPRINT)) ? GetOmWindow() : this);
}

//+------------------------------------------------------------------------
//
//  Member:     CloseErrorInfo
//
//  Synopsis:   Pass the call to the form so it can return its clsid
//              instead of the object's clsid as in CBase.
//
//-------------------------------------------------------------------------
HRESULT CElement::CloseErrorInfo(HRESULT hr)
{
    Doc()->CloseErrorInfo(hr);

    return hr;
}

HRESULT CElement::ScrollIntoView(SCROLLPIN spVert, SCROLLPIN spHorz, BOOL fScrollBits)
{
    HRESULT hr = S_OK;

    if(GetFirstBranch())
    {
        CLayout* pLayout;

        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedParentLayout();

        if(pLayout)
        {
            hr = pLayout->ScrollElementIntoView(this, spVert, spHorz, fScrollBits);
        }
    }

    return hr;
}

HRESULT CElement::DeferScrollIntoView(SCROLLPIN spVert, SCROLLPIN spHorz )
{
    HRESULT hr;

    extern void GWKillMethodCall(void* pvObject, PFN_VOID_ONCALL pfnOnCall, DWORD_PTR dwContext);
    GWKillMethodCall(this, ONCALL_METHOD(CElement, DeferredScrollIntoView, DeferredScrollIntoView), 0);
    hr = GWPostMethodCall(this, ONCALL_METHOD(CElement, DeferredScrollIntoView, DeferredScrollIntoView),
        (DWORD_PTR)(spVert|(spHorz<<16)), FALSE, "CElement::DeferredScrollIntoView");
    return hr;
}

void CElement::DeferredScrollIntoView(DWORD_PTR dwParam)
{
    SCROLLPIN spVert = (SCROLLPIN)((DWORD)dwParam&0xffff);
    SCROLLPIN spHorz = (SCROLLPIN)((DWORD)dwParam>>16);

    ScrollIntoView(spVert, spHorz);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::HitTestPoint, public
//
//  Synopsis:   Determines if this element is hit by the given point
//
//  Arguments:  [pt]        -- Point to check against.
//              [ppSite]    -- If there's a hit, the site that was hit.
//              [ppElement] -- If there's a hit, the element that was hit.
//              [dwFlags]   -- HitTest flags.
//
//  Returns:    HTC
//
//  Notes:      Only ever returns a hit if this element is a relatively
//              positioned element.
//
//              BUGBUG -- This will be changed to call CTxtSite soon. (lylec)
//
//----------------------------------------------------------------------------
HTC CElement::HitTestPoint(CMessage* pMessage, CTreeNode** ppNodeElement, DWORD dwFlags)
{
    CLayout*    pLayout = GetUpdatedNearestLayout();
    HTC         htc     = HTC_NO;

    if(pLayout)
    {
        CDispNode*  pDispNodeOut = NULL;
        CView*      pView        = pLayout->GetView();
        POINT       ptContent;

        if(pView)
        {
            *ppNodeElement = GetFirstBranch();

            htc = pLayout->GetView()->HitTestPoint(
                pMessage->pt,
                COORDSYS_GLOBAL,
                this,
                dwFlags,
                &pMessage->resultsHitTest,
                ppNodeElement,
                ptContent,
                &pDispNodeOut);
        }
    }

    return htc;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::HandleMessage
//
//  Synopsis:   Perform any element specific mesage handling
//
//  Arguments:  pmsg    Ptr to incoming message
//
//  Notes:      pBranch should always be non-scoped. This is the only way
//              that the context information can be maintained. each element
//              HandleMessage which uses pBranch needs to be very careful
//              to scope first.
//
//-------------------------------------------------------------------------
HRESULT CElement::HandleMessage(CMessage* pmsg)
{
    HRESULT hr = S_FALSE;

    // Only the marquee is allowed to cheat and pass the wrong
    // context in
    Assert(IsInMarkup());

    CLayout* pLayout = GetUpdatedLayout();

    if(pLayout)
    {
        Assert(!pmsg->fStopForward);
        hr = pLayout->HandleMessage(pmsg);
        if(hr!=S_FALSE || pmsg->fStopForward)
        {
            goto Cleanup;
        }
    }

    if(pmsg->message == WM_SETCURSOR)
    {
        hr = SetCursorStyle((LPTSTR)NULL, GetFirstBranch());
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     DisallowSelection
//
//  Synopsis:   Returns TRUE iff selection must be diallowed.
//
//-------------------------------------------------------------------------
BOOL CElement::DisallowSelection()
{
    // In dialogs, only editable controls can be selected
    return ((Doc()->_dwFlagsHostInfo&DOCHOSTUIFLAG_DIALOG)
        && !(HasLayout() && GetCurLayout()->_fAllowSelectionInDialog && IsEnabled()));
}

//+-------------------------------------------------------------------------------
//
//  Memeber:    SetImeState
//
//  Synopsis:   Check imeMode to set state of IME.
//
//+-------------------------------------------------------------------------------
HRESULT CElement::SetImeState()
{
    HRESULT         hr  = S_OK;
    CDocument*      pDoc = Doc();
    Assert(pDoc->_pInPlace->_hwnd);
    HIMC            himc = ImmGetContext(pDoc->_pInPlace->_hwnd);
    styleImeMode    sty;
    BOOL            fSuccess;
    VARIANT         varValue;

    if(!himc)
    {
        goto Cleanup;
    }

    hr = ComputeExtraFormat(
        DISPID_A_IMEMODE,
        FALSE,
        GetUpdatedNearestLayoutNode(),
        &varValue);
    if(hr)
    {
        goto Cleanup;
    }

    sty = (((CVariant*)&varValue)->IsEmpty()) ? styleImeModeNotSet : (styleImeMode)V_I4(&varValue);

    switch(sty)
    {
    case styleImeModeActive:
        fSuccess = ImmSetOpenStatus(himc, TRUE);
        if(!fSuccess)
        {
            goto Cleanup;
        }
        break;

    case styleImeModeInactive:
        fSuccess = ImmSetOpenStatus(himc, FALSE);
        if(!fSuccess)
        {
            goto Cleanup;
        }
        break;

    case styleImeModeDisabled:
        pDoc->_himcCache = ImmAssociateContext(pDoc->_pInPlace->_hwnd, NULL);
        break;

    case styleImeModeNotSet:
    case styleImeModeAuto:
    default:
        break;
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------
//
//  member : ComputeExtraFormat
//
//  Synopsis : Uses a miodified ComputeFomrats call to return requested
//              property. Only some of the properties can be returned this way.
//+----------------------------------------------------------------
HRESULT CElement::ComputeExtraFormat(DISPID dispID, BOOL fInherits, CTreeNode* pTreeNode, VARIANT* pVarReturn)
{
    BYTE            ab[sizeof(CFormatInfo)];
    CFormatInfo*    pInfo = (CFormatInfo*)ab;
    HRESULT         hr;

    Assert(pVarReturn);
    Assert(pTreeNode);

    // Make sure that the formats are calculated
    pTreeNode->GetCharFormatIndex();
    pTreeNode->GetFancyFormatIndex();

    VariantInit(pVarReturn);

    // Set the special mode flag so that ComputeFormats does not use
    // cached info,
    pInfo->_eExtraValues = (fInherits) ? ComputeFormatsType_GetInheritedValue : ComputeFormatsType_GetValue;

    // Save the requested property dispID
    pInfo->_dispIDExtra = dispID;
    pInfo->_pvarExtraValue = pVarReturn;

    hr = ComputeFormats(pInfo, pTreeNode);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    pInfo->Cleanup();
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::DoClick
//
//  Arguments:  pMessage    Message that resulted in the click. NULL when
//                          called by OM method click().
//              if this is called on a disabled site, we still want to
//              fire the event (the event code knows to do the right thing
//              and start above the disabled site). and we still want to
//              call click action... but not on us, on the parent
//
//-------------------------------------------------------------------------
HRESULT CElement::DoClick(CMessage* pMessage/*=NULL*/, CTreeNode* pNodeContext/*=NULL*/, BOOL fFromLabel/*=FALSE*/)
{
    HRESULT hr = S_OK;

    if(!pNodeContext)
    {
        pNodeContext = GetFirstBranch();
    }

    if(!pNodeContext)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(!TestLock(ELEMENTLOCK_CLICK))
    {
        CLock Lock(this, ELEMENTLOCK_CLICK);
        CTreeNode::CLock NodeLock(pNodeContext);

        if(Fire_onclick(pNodeContext, pMessage ? pMessage->lSubDivision : 0))
        {
            // Bubble clickaction up the parent chain,
            CTreeNode* pNode = pNodeContext;
            while(pNode)
            {
                if(!(pNode->HasLayout() && pNode->Element()->GetAAdisabled()))
                {
                    // we're not a disabled site
                    hr = pNode->Element()->ClickAction(pMessage);
                    if(hr != S_FALSE)
                    {
                        break;
                    }
                }

                // We need to break if we are called because
                // a label was clicked in case the element is a child
                // of the label. (BUG 19132 - krisma)
                if(fFromLabel == TRUE)
                {
                    // If we're the map, break out right now, since the associated
                    // IMG has already fired its events.
                    break;
                }

                if(pNode->Element()->IsRoot())
                {
                    // jump to the master element
                    CElement* pElemMaster = pNode->Element()->MarkupMaster();

                    pNode = (pElemMaster) ? pElemMaster->GetFirstBranch() : NULL;
                }
                else if(pNode->Parent())
                {
                    pNode = pNode->Parent();
                }
                else
                {
                    pNode = pNode->Element()->GetFirstBranch();
                    if(pNode)
                    {
                        pNode = pNode->Parent();
                    }
                }
            }

            // Propagate error codes from ClickAction, but not S_FALSE
            // because that could confuse callers into thinking that the
            // message was not handled.?
            if(S_FALSE == hr)
            {
                hr = S_OK;
            }
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::ClickAction
//
//  Arguments:  pMessage    Message that resulted in the click. NULL when
//                          called by OM method click().
//
//  Synopsis:   Returns S_FALSE if this should bubble up to the parent.
//              Returns S_OK otherwise.
//
//-------------------------------------------------------------------------
HRESULT CElement::ClickAction(CMessage* pMessage)
{
    return S_FALSE;
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::ShowTooltip
//
//  Synopsis:   Displays the tooltip for the site.
//
//  Arguments:  [pt]    Mouse position in container window coordinates
//              msg     Message passed to tooltip for Processing
//
//--------------------------------------------------------------------------
HRESULT CElement::ShowTooltip(CMessage* pmsg, POINT pt)
{
    HRESULT     hr = S_FALSE;
    RECT        rc;
    CDocument*  pDoc = Doc();
    TCHAR*      pchString;
    BOOL        fRTL = FALSE;

    // if there is a title property, use it as tooltip
    pchString = (LPTSTR)GetAAtitle();
    if(pchString != NULL)
    {
        GetElementRc(&rc, GERC_CLIPPED|GERC_ONALINE, &pt);

        //It is possible to have an empty rect when an
        //element doesn't have a TxtSite above it.
        //Should this be an ASSERT or BUGBUG?
        if(IsRectEmpty(&rc))
        {
            rc.left = pt.x - 10;
            rc.right = pt.x + 10;
            rc.top = pt.y - 10;
            rc.bottom = pt.y + 10;
        }

        // Ignore spurious WM_ERASEBACKGROUNDs generated by tooltips
        CServer::CLock Lock(pDoc, SERVERLOCK_IGNOREERASEBKGND);

        // COMPLEXSCRIPT - determine if element is right to left for tooltip style setting
        if(GetFirstBranch())
        {
            fRTL = GetFirstBranch()->GetCharFormat()->_fRTL;
        }

        FormsShowTooltip(pchString, pDoc->_pInPlace->_hwnd, *pmsg, &rc, (DWORD_PTR)this, fRTL);
        hr = S_OK;
    }

Cleanup:
    return hr;
}

//+------------------------------------------------------------------------
//
//  Member  :   SetCursorStyle
//
//  Synopsis : if the element.style.cursor property is set then on the handling
//      of the WM_CURSOR we should set the cursor to the one specified in
//      the style.
//
//-----------------------------------------------------------------------------
HRESULT CElement::SetCursorStyle(LPCTSTR idcArg, CTreeNode* pContext/*=NULL*/)
{
    HRESULT     hr = E_FAIL;
    LPCTSTR     idc;
    CDocument*  pDoc = Doc();
    static const LPCTSTR aStyleToCursor[] =
    {
        IDC_ARROW,                      // auto map to arrow
        IDC_CROSS,                      // map to crosshair
        IDC_ARROW,                      // default map to arrow
        IDC_HAND,                       // hand map to IDC_HYPERLINK
        IDC_SIZEALL,                    // move map to SIZEALL
        MAKEINTRESOURCE(0),             // e-resize
        MAKEINTRESOURCE(0),             // ne-resize
        MAKEINTRESOURCE(0),             // nw-resize
        MAKEINTRESOURCE(0),             // n-resize
        MAKEINTRESOURCE(0),             // se-resize
        MAKEINTRESOURCE(0),             // sw-resize
        MAKEINTRESOURCE(0),             // s-resize
        MAKEINTRESOURCE(0),             // w-resize
        IDC_IBEAM,                      // text
        IDC_WAIT,                       // wait
        IDC_HELP,                       // help as IDC_help
    };

    const CCharFormat* pCF;

    pContext = pContext ? pContext : GetFirstBranch();
    if(!pContext)
    {
        goto Cleanup;
    }

    pCF = pContext->GetCharFormat();
    Assert(pCF);

    if(!pDoc->_fEnableInteraction)
    {
        // Waiting for page to navigate.  Show wait cursor.
        idc = IDC_APPSTARTING;
    }
    else if(pCF->_bCursorIdx)
    {
        // The style is set to something other than auto
        // so use aStyleToCursor to map the enum to a cursor id.
        Assert(pCF->_bCursorIdx>=0 && pCF->_bCursorIdx<ARRAYSIZE(aStyleToCursor));
        idc = aStyleToCursor[pCF->_bCursorIdx];
    }
    else if(idcArg)
    {
        idc = idcArg;
    }
    else
    {
        // We didn't handle it.
        idc = NULL;
    }

    if(idc)
    {
        SetCursorIDC(idc);
        hr = S_OK;
    }
    else
    {
        hr = S_FALSE;
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::Invalidate
//
//  Synopsis:   Invalidates the area occupied by the element.
//
//-------------------------------------------------------------------------
void CElement::Invalidate()
{
    if(!GetFirstBranch())
    {
        return;
    }

    InvalidateElement();
}

//+-----------------------------------------------------------------
//
//  Member :    CElement::GetElementRegion
//
//  Synopsis:   Get the region corresponding to the element
//
//  Arguments:  paryRects - array to hold the rects corresponding to
//                          region
//              dwFlags   - flags to control the coordinate system
//                          the rects are returned in.
//                          0 - returns the region relative to the
//                              parent content
//                          RFE_SCREENCOORD - window/document/global
//
//------------------------------------------------------------------
void CElement::GetElementRegion(CDataAry<RECT>* paryRects, RECT* prcBound, DWORD dwFlags)
{
    CRect rect;
    BOOL fAppendRect = FALSE;

    if(!prcBound)
    {
        prcBound = &rect;
    }

    // make sure that current is calced
    SendNotification(NTYPE_ELEMENT_ENSURERECALC);

    // BUGBUG (MohanB) The top-left corner is set to (0,0) because
    // we don't know the context, i.e. the image to which this MAP
    // or AREA is bound.
    switch(Tag())
    {
    case ETAG_AREA:
        AssertSz(FALSE, "must improve");
        break;

    case ETAG_OPTION:
        *prcBound = _afxGlobalData._Zero.rc;
        fAppendRect = TRUE;
        break;

    case ETAG_HTML:
        {
            CSize size;

            Doc()->GetView()->GetViewSize(&size);
            prcBound->top = prcBound->left = 0;
            prcBound->right = size.cx;
            prcBound->bottom = size.cy;

            fAppendRect = TRUE;
        }
        break;

    default:
        {
            CLayout* pLayout = GetUpdatedNearestLayout();

            if(pLayout)
            {
                // Get the array that contains bounding rects for each line of the text
                // We want to account for aligned content contained within the element
                // when computing the region.
                dwFlags |= RFE_ELEMENT_RECT;
                pLayout->RegionFromElement(this, paryRects, prcBound, dwFlags);
            }
        }
        break;
    }

    if(fAppendRect)
    {
        paryRects->AppendIndirect((RECT*)prcBound);
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     GetElementRc
//
//  Synopsis:   Get the bounding rect for the element
//
//  Arguments:  prc:      the rc to be returned
//              dwFlags:  flags indicating desired behaviour
//                           GERC_ONALINE: rc on a line, line indicated by ppt.
//                           GERC_CLIPPED: rc clipped by visible client rect.
//              ppt:      the point around which we want the rect
//
//  Returns:    hr
//
//--------------------------------------------------------------------------
HRESULT CElement::GetElementRc(RECT* prc, DWORD dwFlags, POINT* ppt)
{
    CDataAry<RECT> aryRects;
    HRESULT hr = S_FALSE;
    LONG i;

    Assert(prc);

    // Get the region for the element
    GetElementRegion(
        &aryRects,
        !(dwFlags&GERC_ONALINE)?prc:NULL,
        RFE_SCREENCOORD);

    if(dwFlags & GERC_ONALINE)
    {
        Assert(ppt);
        for(i=0; i<aryRects.Size(); i++)
        {
            if(PtInRect(&aryRects[i], *ppt))
            {
                *prc = aryRects[i];
                hr = S_OK;
                break;
            }
        }
    }

    if((S_OK==hr) && (dwFlags&GERC_CLIPPED))
    {
        RECT        rcVisible;
        CDispNode*  pDispNode;
        CLayout*    pLayout = GetUpdatedNearestLayout();

        Assert(Doc()->GetPrimaryElementClient());
        pLayout = pLayout ? pLayout : Doc()->GetPrimaryElementClient()->GetUpdatedNearestLayout();

        Assert(pLayout);
        pDispNode = pLayout->GetElementDispNode(this);

        if(!pDispNode)
        {
            pDispNode = pLayout->GetElementDispNode(pLayout->ElementOwner());
        }


        if(pDispNode)
        {
            pDispNode->GetClippedBounds(&rcVisible, COORDSYS_GLOBAL);
            IntersectRect(prc, prc, &rcVisible);
        }
    }

    return hr;
}

// helper function for offset height and width
void CElement::GetBoundingSize(SIZE& sizeBounds)
{
    CRect rcBound;

    sizeBounds.cx=0;
    sizeBounds.cy=0;

    GetBoundingRect(&rcBound);

    sizeBounds = rcBound.Size();
}

//+-----------------------------------------------------------------
//
//  member : CElement::GetBoundingRect
//
//  Synopsis:   Get the region corresponding to the element
//
//  Arguments:  pRect   - bounding rect of the element
//              dwFlags - flags to control the coordinate system
//                        the rect.
//                        0 - returns the region relative to the
//                            parent content
//                        RFE_SCREENCOORD - window/document/global
//
//------------------------------------------------------------------
HRESULT CElement::GetBoundingRect(CRect* pRect, DWORD dwFlags)
{
    HRESULT         hr      = S_OK;
    CTreeNode*      pNode   = GetFirstBranch();
    CDataAry<RECT>  aryRects;

    Assert(pRect);
    pRect->SetRectEmpty();

    if(!pNode)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    GetElementRegion(&aryRects, pRect, dwFlags);

Cleanup:
    RRETURN(hr);
}

// helper function for offset left and top
HRESULT CElement::GetElementTopLeft(POINT & pt)
{
    HRESULT hr = S_OK;

    pt = _afxGlobalData._Zero.pt;

    if(IsInMarkup())
    {
        Assert(GetFirstBranch());

        CLayout* pLayout = GetUpdatedNearestLayout();

        if(pLayout)
        {
            if(HasLayout())
            {
                pLayout->GetPosition(&pt);

                if(!pLayout->ElementOwner()->IsAbsolute())
                {
                    // Absolutely positined elements are already properly reporting their position
                    //
                    // we are not a TD/TH, but our PARENT might be!
                    // if we are in a table cell, then we need to adjust for the cell insets,
                    // in case the content is vertically aligned.
                    //-----------------------------------------------------------
                    CLayout* pParentLayout = pLayout->GetUpdatedParentLayout();

                    if(pParentLayout && ((pParentLayout->Tag()==ETAG_TD) || (pParentLayout->Tag()==ETAG_TH) || (pParentLayout->Tag()==ETAG_CAPTION)))
                    {
                        CDispNode* pDispNode = pParentLayout->GetElementDispNode();

                        if(pDispNode && pDispNode->HasInset())
                        {
                            const CSize& sizeInset = pDispNode->GetInset();
                            pt.x += sizeInset.cx;
                            pt.y += sizeInset.cy;
                        }
                    }

                }
            }
            else
            {
                hr = pLayout->GetChildElementTopLeft(pt, this);
            }
        }
    }

    RRETURN(hr);
}

//+----------------------------------------------------
//
//  member : GetInheritedBackgroundColor
//
//  synopsis : returns the actual background color
//
//-----------------------------------------------------
COLORREF CElement::GetInheritedBackgroundColor(CTreeNode* pNodeContext/*=NULL*/)
{
    CColorValue ccv;
    GetInheritedBackgroundColorValue(&ccv, pNodeContext);

    return ccv.GetColorRef();
}

//+----------------------------------------------------
//
//  member : GetInheritedBackgroundColorValue
//
//  synopsis : returns the actual background color as a CColorValue
//
//-----------------------------------------------------
HRESULT CElement::GetInheritedBackgroundColorValue(CColorValue* pVal, CTreeNode* pNodeContext /* = NULL */ )
{
    HRESULT hr = S_OK;
    CTreeNode* pNode = pNodeContext ? pNodeContext : GetFirstBranch();

    if(!pNode)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    Assert(pVal != NULL);

    do
    {
        *pVal = pNode->GetFancyFormat()->_ccvBackColor;
        pNode = pNode->Parent();
    } while(!pVal->IsDefined() && pNode);

    // The root site should always have a background color defined.
    Assert(pVal->IsDefined());

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetColors
//
//  Synopsis:   Gets the color set for the Element.
//
//  Returns:    The return code is as per GetColorSet.
//              If any child site fails, that is our return code.  If any
//              child site returns S_OK, that is our return code, otherwise
//              we return S_FALSE.
//
//----------------------------------------------------------------------------
HRESULT CElement::GetColors(CColorInfo* pCI)
{
    DWORD_PTR dw = 0;
    HRESULT hr = S_FALSE;

    CLayout* pLayoutThis = GetUpdatedLayout();
    CLayout* pLayout;

    Assert(pLayoutThis && "CElement::GetColors() should not be called here !!!");
    if(!pLayoutThis)
    {
        goto Error;
    }

    for(pLayout=pLayoutThis->GetFirstLayout(&dw);
        pLayout && !pCI->IsFull();
        pLayout=pLayoutThis->GetNextLayout(&dw))
    {
        HRESULT hrTemp = pLayout->ElementOwner()->GetColors(pCI);
        if(FAILED(hrTemp) && hrTemp!=E_NOTIMPL)
        {
            hr = hrTemp;
            goto Error;
        }
        else if(hrTemp == S_OK)
        {
            hr = S_OK;
        }
    }

Error:
    if(pLayoutThis)
    {
        pLayoutThis->ClearLayoutIterator(dw, FALSE);
    }
    RRETURN1(hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::FireCancelableEvent
//
//  Synopsis:   Fire the event, called from PDL-generated stubs.
//
//  Returns:    TRUE: Take default action, FALSE: Don't take default action
//
//-----------------------------------------------------------------------------
BOOL CElement::FireCancelableEvent(
           DISPID dispidEvent,
           DISPID dispidProp,
           LPCTSTR pchEventType,
           BYTE* pbTypes, ...)
{
    CLayout* pLayout = GetUpdatedNearestLayout();
    BOOL fRet = TRUE;

    // No nearest layout->assume enabled (dbau)
    // wlw note
    //if((!pLayout || !pLayout->ElementOwner()->GetAAdisabled()) && ShouldFireEvents())
    //{
    //    VARIANT_BOOL    vb;
    //    CVariant        Var;
    //    CDocument*      pDoc = Doc();
    //    EVENTPARAM      param(pDoc, TRUE);
    //    CElement*       pElemFire;
    //    IDispatch*      pEventObj = NULL;
    //    CDocument::CLock Lock(pDoc);
    //    va_list         valParms;

    //    // The first element in a slave tree is just a proxy for its master,
    //    // so it lets the master fire the event.
    //    pElemFire = FireEventWith();

    //    // NEWTREE: GetFirstBranch is iffy here
    //    param.SetNodeAndCalcCoordinates(pElemFire->GetFirstBranch());
    //    param.SetType(pchEventType);

    //    va_start(valParms, pbTypes);

    //    // Get the eventObject.
    //    if(!pDoc->EnsureOmWindow())
    //    {
    //        pDoc->_pOmWindow->get_event((IHTMLEventObj**)&pEventObj);
    //    }

    //    (void)pElemFire->FireEventV(
    //        dispidEvent,
    //        dispidProp,
    //        pEventObj,
    //        &Var,
    //        pbTypes,
    //        valParms);

    //    vb = (V_VT(&Var)==VT_BOOL) ? V_BOOL(&Var) : VB_TRUE;

    //    fRet = !param.IsCancelled() && VB_TRUE == vb;

    //    va_end(valParms);

    //    ReleaseInterface(pEventObj);
    //}

    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::BubbleCancelableEvent
//
//  Synopsis:   Fire the bubbling event, called from PDL-generated stubs.
//
//  Returns:    TRUE: Take default action, FALSE: Don't take default action
//
//-----------------------------------------------------------------------------
BOOL CElement::BubbleCancelableEvent(
         CTreeNode* pNodeContext,
         long lSubDivision,
         DISPID dispidMethod,
         DISPID dispidProp,
         LPCTSTR pchEventType,
         BYTE* pbTypes, ...)
{
    BOOL fRet = TRUE;
    CDocument* pDoc = Doc();

    if(!pNodeContext)
    {
        pNodeContext = GetFirstBranch();
    }

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(pDoc)
    {
        CVariant vb;
        EVENTPARAM param(pDoc, TRUE);

        va_list valParms;

        param.SetNodeAndCalcCoordinates(pDoc->_pElementOMCapture
            && pDoc->_pNodeLastMouseOver
            && ((dispidMethod==DISPID_EVMETH_ONCLICK)
            || (dispidMethod==DISPID_EVMETH_ONDBLCLICK))
            ? pDoc->_pNodeLastMouseOver : pNodeContext);
        param.SetType(pchEventType);
        param.lSubDivision = lSubDivision;

        if(dispidMethod == DISPID_EVMETH_ONBEFOREEDITFOCUS)
        {
            param._pNodeTo = pNodeContext;
        }

        va_start(valParms, pbTypes);

        (void)BubbleEventHelper(pNodeContext, lSubDivision, dispidMethod, dispidProp, FALSE, &vb, pbTypes, valParms);

        if(V_VT(&vb) != VT_EMPTY)
        {
            fRet = (V_VT(&vb)==VT_BOOL) ? V_BOOL(&vb)==VB_TRUE : TRUE;
        }

        fRet = fRet && !param.IsCancelled();
    }
    return fRet;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::FireEvent, public
//
//  Synopsis:   Fires an event out the primary dispatch event connection point.
//              Called from PDL-generated stubs.  Calls FireEvent Helper For
//              the real work, and is mainly responsible for creating the
//              eventparam
//
//  Arguments:  [dispidEvent]   -- DISPID of event to fire
//              [dispidProp]    -- Dispid of prop storing event function
//              [pchEventType]  -- String of type of event
//              [pbTypes]       -- Pointer to array giving the types of parms
//              [...]           -- Parameters
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CElement::FireEvent(
        DISPID dispidEvent,
        DISPID dispidProp,
        LPCTSTR pchEventType,
        BYTE* pbTypes, ...)
{
    HRESULT         hr = S_OK;
    CDocument*      pDoc = Doc();
    CElement*       pElemFire = NULL;

    //pElemFire = FireEventWith();

    //if(ShouldFireEvents())
    //{
    //    CPeerHolder::CLock lock(pPeerHolder);

    //    EVENTPARAM param(pDoc, TRUE);

    //    // NEWTREE: GetFirstBranch is iffy here
    //    //          I've seen this be null on a timer event.
    //    param.SetNodeAndCalcCoordinates(pElemFire->GetFirstBranch());
    //    param.SetType(pchEventType);
    //    hr = pElemFire->FireEventHelper(dispidEvent, dispidProp, pbTypes);
    //}

    //// since focus/blur do not bubble, and things like accessibility
    //// need to have a centralized place to handle focus changes, if
    //// we just fired focus or blur for someone other than the doc,
    //// then fire the doc's onfocuschange/onblurchange methods, keeping
    //// the event param structure intact.
    ////
    //// BUGBUG (carled) the ONCHANGEFOCUS/ONCHANGEBLUR need to be removed
    ////   from the document's event interface
    //if(!hr && ((dispidEvent==DISPID_EVMETH_ONFOCUS)||(dispidEvent==DISPID_EVMETH_ONBLUR)) && 
    //    (pElemFire->Tag()!=ETAG_ROOT))
    //{
    //    // BUGBUG (carled) the sourceindex is not necessarily up
    //    //  to date and updateTreeCahce() would need to be called.
    //    //  DO NOT DO THIS. Accessibility1.0 doen't support DHTML
    //    //  and we don't want the pref hit.  Also the srcIndex is
    //    //  unreliable even if ensured and this needs to be changed
    //    //  when DHTML is actually supported.

    //    // Since carled doesn't want a monster walk (why?!?) I'll
    //    // just use __iSourceIndex directly here to avoid annoying
    //    // asserts that everyone has been complaining about.  (jbeda)

    //    // FerhanE
    //    // We have to check is the document has the focus, before firing this, 
    //    // to prevent firing of accessible focus/blur events when the document
    //    // does not have the focus. 
    //    if(pDoc->HasFocus())
    //    {
    //        hr = pDoc->FireAccesibilityEvents(dispidEvent, pElemFire->GetSourceIndex());
    //    }
    //}

    RRETURN(hr);
}

HRESULT CElement::FireEventHelper(DISPID dispidEvent, DISPID dispidProp, BYTE* pbTypes, ...)
{
    HRESULT hr = S_OK;

    Assert(FireEventWith() == this);
    if(!GetAAdisabled())
    {
        va_list         valParms;
        CDocument*      pDoc = Doc();
        CDocument::CLock Lock(pDoc);
        IDispatch*      pEventObj = NULL;

        va_start(valParms, pbTypes);

        // wlw note
        //// Get the eventObject.
        //if(!pDoc->EnsureOmWindow())
        //{
        //    pDoc->_pOmWindow->get_event((IHTMLEventObj **)&pEventObj);
        //}

        hr = FireEventV(
            dispidEvent,
            dispidProp,
            pEventObj,
            NULL,
            pbTypes,
            valParms);

        ReleaseInterface(pEventObj);

        va_end(valParms);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::BubbleEvent, public
//
//  Synopsis:   Fires an event out the primary dispatch event connection point,
//              and bubbles up the parent chain firing the event from
//              parent sites, as well.   Called from PDL-generated stubs.
//
//  Arguments:  [dispid]  -- DISPID of event to fire
//              [pbTypes] -- Pointer to array giving the types of the params
//              [...]     -- Parameters
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CElement::BubbleEvent(
          CTreeNode* pNodeContext,
          long lSubDivision,
          DISPID dispidMethod,
          DISPID dispidProp,
          LPCTSTR pchEventType,
          BYTE* pbTypes, ...)
{
    HRESULT hr = S_OK;
    CDocument* pDoc = Doc();

    if(!pNodeContext)
    {
        pNodeContext = GetFirstBranch();
    }

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(!pNodeContext)
    {
        goto Cleanup;
    }

    if(pDoc)
    {
        EVENTPARAM param(pDoc, TRUE);

        va_list valParms;

        param.SetNodeAndCalcCoordinates(pNodeContext);
        param.SetType(pchEventType);
        param.lSubDivision = lSubDivision;

        va_start(valParms, pbTypes);

        hr = BubbleEventHelper(pNodeContext, lSubDivision, dispidMethod, dispidProp, FALSE, NULL, pbTypes, valParms);
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Function:   BubbleEventHelper
//
//  Synopsis:   Fire the specified event. All the sites in the parent chain
//              are supposed to fire the events (if they can).  Caller has
//              responsibility for setting up any EVENTPARAM.
//
//  Arguments:  [dispidEvent]   -- dispid of the event to fire.
//              [dispidProp]    -- dispid of prop containing event func.
//              [pvb]           -- Boolean return value
//              [pbTypes]       -- Pointer to array giving the types of parms
//              [...]           -- Parameters
//
//  Returns:    S_OK if successful
//
//-----------------------------------------------------------------------------
HRESULT CElement::BubbleEventHelper(
            CTreeNode*  pNodeContext,
            long        lSubDivision,
            DISPID      dispidEvent,
            DISPID      dispidProp,
            BOOL        fRaisedByPeer,
            VARIANT*    pvb,
            BYTE*       pbTypes, ...)
{
    // wlw note
    //CPeerHolder* pPeerHolder = GetPeerHolder();

    //// don't fire standard events if there is a behavior attached that wants to fire them instead
    //if(pPeerHolder && !fRaisedByPeer &&
    //    IsStandardDispid(dispidProp) &&
    //    pPeerHolder->HasCustomEventMulti(dispidProp))
    //{
    //    return S_OK;
    //}

    //CPeerHolder::CLock lock(pPeerHolder);

    CDocument* pDoc = Doc();

    Assert(pNodeContext && pNodeContext->Element()==this);
    Assert(pDoc->_pparam);

    if(!pNodeContext)
    {
        return S_OK;
    }

    CTreeNode*      pNode = pNodeContext;
    CElement*       pElementReal;
    va_list         valParms;
    CVariant        vb;
    CVariant        Var;
    CLayout*        pLayout = pNodeContext->GetUpdatedNearestLayout();
    unsigned int    cInvalOld = pDoc->_cInval;
    long            cSub = 0;

    // By default do not cancel events
    pDoc->_pparam->fCancelBubble = FALSE;

    // If there are any subdivisions, let the subdivisions handle the
    // event.
    if(OK(GetSubDivisionCount(&cSub)) && cSub)
    {
        va_list valParms;

        va_start(valParms, pbTypes);
        DoSubDivisionEvents(
            lSubDivision,
            dispidEvent,
            dispidProp,
            &Var,
            pbTypes,
            valParms);

        if(V_VT(&Var) != VT_EMPTY)
        {
            V_VT(&vb)   = VT_BOOL;
            V_BOOL(&vb) = (V_VT(&Var)==VT_BOOL) ? V_BOOL(&Var) : VB_TRUE;
        }
    }

    // if the srcElement is the rootsite, return the HTML element
    Assert(pDoc->_pparam->_pNode);

    if((pDoc->_pparam->_pNode->Element()->IsRoot()) && (pDoc->_pElementOMCapture!=this))
    {
        pDoc->_pparam->SetNodeAndCalcCoordinates(NULL);
    }

    CDocument::CLock Lock(pDoc);

    va_start(valParms, pbTypes);

    // if Bubbling cancelled by a sink. Don't bubble anymore
    while(pNode && !pDoc->_pparam->fCancelBubble)
    {
        BOOL fListenerPresent = FALSE;

        // if we're disabled () or in the scope of a disabled site
        // then pass the event to the parent
        pLayout = pNode->GetUpdatedNearestLayout();
        if(pLayout && pLayout->ElementOwner()->GetAAdisabled())
        {
            pNode = pNode->GetUpdatedNearestLayoutNode()->Parent();
            continue;
        }

        pElementReal = pNode->Element();
        CBase* pBase = pElementReal;

        // BUGBUG (MohanB) This logic should change once the root element
        // is gone.

        // If this is the root element in the tree, then the
        // event is to be fired by the CDoc containing the
        // rootsite, rather than the rootsite itself.
        // we do this so that the top of the bubble goes to the doc
        if(pElementReal->IsRoot())
        {
            CElement* pElementMaster = pElementReal->MarkupMaster();

            if(pElementMaster)
            {
                // Bubble through the master's tree
                pNode = pElementMaster->GetFirstBranch();
                if(!pNode)
                {
                    break;
                }
                pBase = pElementReal = pElementMaster;
                pDoc->_pparam->SetNodeAndCalcCoordinates(pNode);

                // BUGBUG (MohanB) What to about _pNodeFrom && _pNodeTo? For now,
                // transform them to respective master elements
                TransformSlaveToMaster(&pDoc->_pparam->_pNodeFrom);
                TransformSlaveToMaster(&pDoc->_pparam->_pNodeTo);
            }
            else
            {
                pBase = pDoc;
            }

            fListenerPresent = TRUE;
        }
        else
        {
            // this is an element in the tree. Check to see if there
            // are any possible listeners for this event. if so, continue
            // if not, don't call FireEventV and let the event continue 
            // in its bubbling.
            fListenerPresent = pElementReal->ShouldFireEvents();
        }

        pNode->NodeAddRef();

        if(fListenerPresent)
        {
            IDispatch* pEventObj = NULL;

            // Get the eventObject.
            /*if(!pDoc->EnsureOmWindow())
            {
                pDoc->_pOmWindow->get_event((IHTMLEventObj**)&pEventObj);
            } wlw note*/

            pBase->FireEventV(
                dispidEvent,
                dispidProp,
                pEventObj,
                &Var,
                pbTypes,
                valParms);

            ReleaseInterface(pEventObj);

            if(V_VT(&Var) != VT_EMPTY)
            {
                V_VT(&vb)   = VT_BOOL;
                V_BOOL(&vb) = (V_VT(&Var)==VT_BOOL) ? V_BOOL(&Var) : VB_TRUE;
            }
        }

        // If the node is no longer valid we're done.  Script in the event handler caused
        // the tree to change.
        if(!pNode->IsDead())
        {
            pNode->NodeRelease();
            pNode = pNode->Parent();
            if(!pNode)
            {
                pNode = pElementReal->GetFirstBranch();
                if(pNode)
                {
                    pNode = pNode->Parent();
                }
            }
        }
        else
        {
            pNode->NodeRelease();
            pNode = NULL;
        }
    }

    // if we're still bubbling, we need to go all the way up to the window also.
    // Currently the only event that does this, is onhelp so this test is here to
    // minimize the work
    if(!pDoc->_pparam->fCancelBubble && GetOmWindow() && dispidEvent==DISPID_EVMETH_ONHELP)
    {
        // wlw note
        //IDispatch* pEventObj = NULL;
        //CBase* pBase = GetOmWindow();

        //// Get the eventObject.
        //if(!pDoc->EnsureOmWindow())
        //{
        //    pDoc->_pOmWindow->get_event((IHTMLEventObj**)&pEventObj);
        //}

        //pBase->FireEventV(
        //    dispidEvent,
        //    dispidProp,
        //    pEventObj,
        //    &Var,
        //    pbTypes,
        //    valParms);

        //ReleaseInterface(pEventObj);

        //if(V_VT(&Var) != VT_EMPTY)
        //{
        //    V_VT(&vb)   = VT_BOOL;
        //    V_BOOL(&vb) = (V_VT(&Var)==VT_BOOL) ? V_BOOL(&Var) : VB_TRUE;
        //}
    }

    va_end(valParms);

    if(pvb)
    {
        ((CVariant*)pvb)->Copy(&vb);
    }

    // set a flag in doc, if the script caused an invalidation
    if(cInvalOld != pDoc->_cInval)
    {
        pDoc->_fInvalInScript = TRUE;
    }

    return S_OK;
}

//+----------------------------------------------------------------------------
//
//  Function:   DoSubDivisionEvents
//
//  Synopsis:   Fire the specified event on the given subdivision.
//
//  Arguments:  [dispidEvent]   -- dispid of the event to fire.
//              [dispidProp]    -- dispid of prop containing event func.
//              [pvb]           -- Boolean return value
//              [pbTypes]       -- Pointer to array giving the types of parms
//              [...]           -- Parameters
//
//-----------------------------------------------------------------------------
HRESULT CElement::DoSubDivisionEvents(
          long      lSubDivision,
          DISPID    dispidEvent,
          DISPID    dispidProp,
          VARIANT*  pvb,
          BYTE*     pbTypes, ...)
{
    return S_OK;
}

BOOL AllowCancelKeydown(CMessage* pMessage)
{
    Assert(pMessage->message==WM_SYSKEYDOWN || pMessage->message==WM_KEYDOWN);

    WPARAM wParam = pMessage->wParam;
    DWORD  dwKeyState = pMessage->dwKeyState;
    int i;

    struct KEY
    { 
        WPARAM  wParam;
        DWORD   dwKeyState;
    };

    static KEY s_aryVK[] =
    {
        VK_F1,      0,
        VK_F2,      0,
        VK_F3,      0,
        VK_F4,      0,
        VK_F5,      0,
        VK_F6,      0,
        VK_F7,      0,
        VK_F8,      0,
        VK_F9,      0,
        VK_F10,     0,
        VK_F11,     0,
        VK_F12,     0,
        VK_SHIFT,   MK_SHIFT,
        VK_F4,      MK_CONTROL,
        VK_TAB,     MK_CONTROL,
        VK_TAB,     MK_CONTROL | MK_SHIFT,
        65,         MK_CONTROL, // ctrl-a
        70,         MK_CONTROL, // ctrl-f
        79,         MK_CONTROL, // ctrl-o
        80,         MK_CONTROL, // ctrl-p
    };

    if(dwKeyState == MK_ALT)
    {
        return FALSE;
    }

    for(i=0; i<ARRAYSIZE(s_aryVK); i++)
    {
        if(s_aryVK[i].wParam==wParam && s_aryVK[i].dwKeyState==dwKeyState)
        {
            return FALSE;
        }
    }

    return TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     FireEventOnMessage
//
//  Synopsis:   fires event corresponding to message
//
//  Arguments:  [pMessage]   -- message
//
//  Returns:    return value satisfies same rules as HandleMessage:
//              S_OK        don't do anything - event was cancelled;
//              S_FALSE     keep processing the message - event was not cancelled
//              other       error
//
//----------------------------------------------------------------------------
HRESULT CElement::FireStdEventOnMessage(
        CTreeNode*  pNodeContext,
        CMessage*   pMessage,
        CTreeNode*  pNodeBeginBubbleWith/*=NULL*/,
        CTreeNode*  pNodeEvent/*=NULL*/)
{
    Assert(pNodeContext && pNodeContext->Element()==this);

    if(pMessage->fEventsFired)
    {
        return S_FALSE;
    }

    HRESULT     hr = S_FALSE;
    POINT       ptEvent = pMessage->pt;
    CDocument*  pDoc = Doc();
    CTreeNode*  pNodeThisCanFire = pNodeContext;
    CTreeNode::CLock lock(pNodeThisCanFire);

    // BUGBUG (alexz) (anandra) need this done in a generic way for all events.
    // about to fire an event; if there are deferred scripts, commit them now.
    /*pDoc->CommitDeferredScripts(TRUE); wlw note*/ // TRUE - early, so don't commit downloaded deferred scripts

    // Keyboard events are fired only once before the message is
    // dispatched.
    switch(pMessage->message)
    {
    case WM_HELP:
        hr = pNodeThisCanFire->Element()->Fire_onhelp(
            pNodeThisCanFire, pMessage?pMessage->lSubDivision:0) ? S_FALSE : S_OK;
        break;

    case WM_SYSKEYDOWN:
    case WM_KEYDOWN:
        hr = pNodeThisCanFire->Element()->FireStdEvent_KeyDown(pNodeThisCanFire, pMessage, (int*)&pMessage->wParam, VBShiftState()) ? S_FALSE : S_OK;
        break;

    case WM_SYSKEYUP:
    case WM_KEYUP:
        pNodeThisCanFire->Element()->FireStdEvent_KeyUp(pNodeThisCanFire, pMessage, (int*)&pMessage->wParam, VBShiftState());
        break;

    case WM_CHAR:
        hr = pNodeThisCanFire->Element()->FireStdEvent_KeyPress(pNodeThisCanFire, pMessage, (int*)&pMessage->wParam) ? S_FALSE : S_OK;
        break;

    case WM_LBUTTONDBLCLK:
        pDoc->_fGotDblClk = TRUE;
        goto Cleanup; // To not set the event fired bit.

    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_MBUTTONDOWN:
        {
            FireStdEvent_MouseHelper(
                pNodeContext,
                pMessage,
                DISPID_EVMETH_ONMOUSEDOWN,
                DISPID_EVPROP_ONMOUSEDOWN,
                VBButtonState((short)pMessage->wParam),
                VBShiftState(),
                (float)ptEvent.x, (float)ptEvent.y,
                pNodeContext,
                pNodeContext,
                pNodeBeginBubbleWith,
                pNodeEvent);
            break;
        }

    case WM_LBUTTONUP:
    case WM_RBUTTONUP:
    case WM_MBUTTONUP:
        {
            short sParam;
            CTreeNode* pNodeFireWith = NULL;

            // pMessage->wParam represents the button state
            // not button transition, on up, nothing is down.
            if(pMessage->message == WM_LBUTTONUP)
            {
                sParam = VB_LBUTTON;
            }
            else if(pMessage->message == WM_RBUTTONUP)
            {
                sParam = VB_RBUTTON;
            }
            else if(pMessage->message == WM_MBUTTONUP)
            {
                sParam = VB_MBUTTON;
            }
            else
            {
                sParam = 0;
            }

            // the real element under the mouse should fire mouse up.
            // not the object which had capture. (pElemThisCanFire)
            if(pNodeEvent)
            {
                pNodeFireWith = pNodeThisCanFire;
            }
            else
            {
                pNodeFireWith = pMessage->pNodeHit;
            }

            if(pNodeFireWith == NULL)
            {
                break;
            }

            pNodeFireWith->Element()->FireStdEvent_MouseHelper(
                pNodeFireWith,
                pMessage,
                DISPID_EVMETH_ONMOUSEUP,
                DISPID_EVPROP_ONMOUSEUP,
                sParam,
                VBShiftState(),
                (float)ptEvent.x, (float)ptEvent.y,
                NULL,
                NULL,
                pNodeBeginBubbleWith,
                pNodeEvent);
        }
        break;

    case WM_MOUSELEAVE: // fired for MouseOut event
        pNodeThisCanFire->Element()->FireStdEvent_MouseHelper(
            pNodeThisCanFire,
            pMessage,
            DISPID_EVMETH_ONMOUSEOUT,
            DISPID_EVPROP_ONMOUSEOUT,
            VBButtonState((short)pMessage->wParam),
            VBShiftState(),
            (float)ptEvent.x,
            (float)ptEvent.y,
            pNodeContext,
            pDoc->_pNodeLastMouseOver,
            pNodeBeginBubbleWith,
            pNodeEvent);
        break;

    case WM_MOUSEOVER: // essentially an Enter event
        pNodeThisCanFire->Element()->FireStdEvent_MouseHelper(
            pNodeThisCanFire,
            pMessage,
            DISPID_EVMETH_ONMOUSEOVER,
            DISPID_EVPROP_ONMOUSEOVER,
            VBButtonState((short)pMessage->wParam),
            VBShiftState(),
            (float)ptEvent.x,
            (float)ptEvent.y,
            pDoc->_pNodeLastMouseOver,
            pNodeContext,
            pNodeBeginBubbleWith,
            pNodeEvent);
        break;

    case WM_MOUSEMOVE:
        // now fire the mousemove event
        hr = (pNodeThisCanFire->Element()->FireStdEvent_MouseHelper(
            pNodeThisCanFire,
            pMessage,
            DISPID_EVMETH_ONMOUSEMOVE,
            DISPID_EVPROP_ONMOUSEMOVE,
            VBButtonState((short)pMessage->wParam),
            VBShiftState(),
            (float)ptEvent.x,
            (float)ptEvent.y,
            pDoc->_pNodeLastMouseOver,
            pNodeContext,
            pNodeBeginBubbleWith,
            pNodeEvent))
            ? S_FALSE : S_OK;
        break;

    case WM_CONTEXTMENU:
        hr = pNodeThisCanFire->Element()->Fire_oncontextmenu() ? S_FALSE : S_OK;
        break;

    case WM_MOUSEHOVER:
        hr = pNodeThisCanFire->Element()->Fire_onmousehover() ? S_FALSE : S_OK;
        break;

    default:
        goto Cleanup; // don't set fStdEventsFired
    }

    pMessage->fEventsFired = TRUE;

Cleanup:
    RRETURN1 (hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::FireStdEvent_KeyDown
//
//  Synopsis:   Fire a key event.
//
//  Returns:    TRUE: Take default action, FALSE: Don't take default action
//
//-----------------------------------------------------------------------------
BOOL CElement::FireStdEvent_KeyDown(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode, short shift)
{
    BOOL        fRet = TRUE;
    CDocument*  pDoc = Doc();
    HRESULT     hr;

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(pDoc)
    {
        CVariant    varRet;
        EVENTPARAM  param(pDoc, TRUE);

        param.SetNodeAndCalcCoordinates(pNodeContext);
        param._lKeyCode = (long)*piKeyCode;
        param.SetType(_T("keydown"));
        // the 30th bit of the lparam indicates whether this is a repeated WM_
        param.fRepeatCode = !!(HIWORD(pMessage->lParam) & KF_REPEAT);

        hr = BubbleEventHelper(
            pNodeContext,
            pMessage->lSubDivision,
            DISPID_EVMETH_ONKEYDOWN,
            DISPID_EVPROP_ONKEYDOWN,
            FALSE,
            &varRet,
            (BYTE*) VTS_NONE,
            shift);
        *piKeyCode = (int)param._lKeyCode;
        fRet = !param.IsCancelled() && (V_VT(&varRet)==VT_EMPTY || VB_TRUE==V_BOOL(&varRet));
    }

    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::FireStdEvent_KeyUp
//
//  Synopsis:   Fire the key up event.
//
//  Returns:    TRUE: Take default action, FALSE: Don't take default action
//
//-----------------------------------------------------------------------------
BOOL CElement::FireStdEvent_KeyUp(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode, short shift)
{
    BOOL        fRet = TRUE;
    CDocument*  pDoc = Doc();
    HRESULT     hr;

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(pDoc)
    {
        CVariant    varRet;
        EVENTPARAM  param(pDoc, TRUE);

        param.SetNodeAndCalcCoordinates(pNodeContext);
        param._lKeyCode = (long)*piKeyCode;
        param.SetType(_T("keyup"));

        hr = BubbleEventHelper(
            pNodeContext,
            pMessage->lSubDivision,
            DISPID_EVMETH_ONKEYUP,
            DISPID_EVPROP_ONKEYUP,
            FALSE,
            &varRet,
            (BYTE*)VTS_NONE,
            shift);
        *piKeyCode = (int)param._lKeyCode;
        fRet = !param.IsCancelled() && (V_VT(&varRet)==VT_EMPTY || VB_TRUE==V_BOOL(&varRet));
    }

    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::FireStdEvent_KeyPress
//
//  Synopsis:   Fire the key press event.
//
//  Returns:    TRUE: Take default action, FALSE: Don't take default action
//
//-----------------------------------------------------------------------------
BOOL CElement::FireStdEvent_KeyPress(CTreeNode* pNodeContext, CMessage* pMessage, int* piKeyCode)
{
    BOOL        fRet = TRUE;
    CDocument*  pDoc = Doc();
    HRESULT     hr;

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(pDoc)
    {
        CVariant    varRet;
        EVENTPARAM  param(pDoc, TRUE);

        param.SetNodeAndCalcCoordinates(pNodeContext);
        param._lKeyCode = (long)*piKeyCode;
        param.SetType(_T("keypress"));

        hr = BubbleEventHelper(
            pNodeContext,
            pMessage->lSubDivision,
            DISPID_EVMETH_ONKEYPRESS,
            DISPID_EVPROP_ONKEYPRESS,
            FALSE,
            &varRet,
            (BYTE*)VTS_NONE);
        *piKeyCode = (int)param._lKeyCode;
        fRet = !param.IsCancelled() && (V_VT(&varRet)==VT_EMPTY || VB_TRUE==V_BOOL(&varRet));
    }

    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::FireStdEvent_MouseHelper
//
//  Synopsis:   Fire any mouse event.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
BOOL CElement::FireStdEvent_MouseHelper(
           CTreeNode*   pNodeContext,
           CMessage*    pMessage,
           DISPID       dispidMethod,
           DISPID       dispidProp,
           short        button,
           short        shift,
           float        x,
           float        y,
           CTreeNode*   pNodeFrom,              /*=NULL*/
           CTreeNode*   pNodeTo,                /*=NULL*/
           CTreeNode*   pNodeBeginBubbleWith,   /*=NULL*/
           CTreeNode*   pNodeEvent              /*=NULL*/)
{
    BOOL        fRet = TRUE;
    CDocument*  pDoc = Doc();
    CTreeNode*  pNodeSrcElement = pNodeEvent ? pNodeEvent : pNodeContext;

    Assert(pNodeContext && pNodeContext->Element()==this);

    if(!pNodeBeginBubbleWith)
    {
        pNodeBeginBubbleWith = pNodeContext;
    }

    if(pDoc)
    {
        HRESULT     hr;
        CVariant    varRet;
        EVENTPARAM  param(pDoc, FALSE);
        POINT       pt;

        pt.x = x;
        pt.y = y;

        param._clientX = x;
        param._clientY = y;
        param.SetNodeAndCalcCoordinates(pNodeSrcElement);
        if(pDoc->_pInPlace)
        {
            ClientToScreen(pDoc->_pInPlace->_hwnd, &pt);
        }
        param._screenX = pt.x;
        param._screenY = pt.y;

        param._sKeyState = shift;
        param._lButton = button;
        param.lSubDivision = pMessage->lSubDivision;

        switch(dispidMethod)
        {
            // these have a different parameter lists. i.e. none,
            // and, they initialize two more members of EVENTPARAM
        case DISPID_EVMETH_ONMOUSEOVER:
            param.SetType(_T("mouseover"));
            param._pNodeFrom   = pNodeFrom;
            param._pNodeTo     = pNodeTo;
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                &varRet,
                (BYTE*)VTS_NONE);
            break;

        case DISPID_EVMETH_ONMOUSEOUT:
            param.SetType(_T("mouseout"));
            param._pNodeFrom   = pNodeFrom;
            param._pNodeTo     = pNodeTo;
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                NULL,
                (BYTE*)VTS_NONE);
            break;

            // BUGBUG (MohanB) For now, allow MouseMove to be cancellable so
            // that script writers can override default actions like drag-drop.
        case DISPID_EVMETH_ONMOUSEMOVE:
            param.SetType(_T("mousemove"));
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                &varRet,
                (BYTE*)VTS_NONE,
                button,
                shift,
                x, y);

            fRet = !param.IsCancelled() && (V_VT(&varRet)==VT_EMPTY || VB_TRUE==V_BOOL(&varRet));
            break;

        case DISPID_EVMETH_ONMOUSEUP:
            param.SetType(_T("mouseup"));
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                NULL,
                (BYTE*)VTS_NONE,
                button,
                shift,
                x, y);
            break;

        case DISPID_EVMETH_ONMOUSEDOWN:
            param.SetType(_T("mousedown"));
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                NULL,
                (BYTE*)VTS_NONE,
                button,
                shift,
                x, y);
            break;

        default:
            Assert(0 && "ignorable assert, if we are here there is no event type");
            hr = pNodeBeginBubbleWith->Element()->BubbleEventHelper(
                pNodeBeginBubbleWith,
                pMessage->lSubDivision,
                dispidMethod,
                dispidProp,
                FALSE,
                NULL,
                (BYTE*)VTS_NONE,
                button,
                shift,
                x, y);
            break;
        }
    }
    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::Fire_onpropertychange
//
//  Synopsis:   Fires the onpropertychange event, sets up the event param
//
//+----------------------------------------------------------------------------
void CElement::Fire_onpropertychange(LPCTSTR strPropName)
{
    EVENTPARAM param(Doc(), TRUE);

    param.SetType(_T("propertychange"));
    param.SetPropName(strPropName);
    param.SetNodeAndCalcCoordinates(GetFirstBranch());

    FireEventHelper(DISPID_EVMETH_ONPROPERTYCHANGE,
        DISPID_EVPROP_ONPROPERTYCHANGE, (BYTE*)VTS_NONE);
}

//+-----------------------------------------------------------------
//
//  member : Fire_onscroll
//
//  synopsis : IHTMLTextContainer event implementation
//
//------------------------------------------------------------------
void CElement::Fire_onscroll()
{
    if(Tag() == ETAG_BODY)
    {
        CDocument* pDoc = Doc();

        /*if(pDoc->_pOmWindow)
        {
            pDoc->_pOmWindow->Fire_onscroll();
        } wlw note*/
    }
    else
    {
        FireOnChanged(DISPID_IHTMLELEMENT2_SCROLLTOP);
        FireOnChanged(DISPID_IHTMLELEMENT2_SCROLLLEFT);

        FireEvent(DISPID_EVMETH_ONSCROLL,
            DISPID_EVPROP_ONSCROLL, _T("scroll"), (BYTE*)VTS_NONE);
    }
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::Fire_PropertyChangeHelper
//
//  Synopsis:   Fires the onpropertychange event, sets up the event param
//
//+----------------------------------------------------------------------------
HRESULT CElement::Fire_PropertyChangeHelper(DISPID dispid, DWORD dwFlags)
{
    BSTR            strName = NULL;
    PROPERTYDESC*   ppropdesc;
    DISPID          expDispid;

    // first, find the appropriate propdesc for this property.
    if(dwFlags & ELEMCHNG_INLINESTYLE_PROPERTY)
    {
        CBase* pStyleObj = GetInLineStylePtr();

        if(pStyleObj)
        {
            CBufferedString cBuf;

            cBuf.Set((dwFlags&ELEMCHNG_INLINESTYLE_PROPERTY) ? 
                _T("style.") : _T("runtimeStyle."));
            // if we still can't find it, or have no inline
            // then bag this, and continue
            if(S_OK == pStyleObj->FindPropDescFromDispID(dispid, &ppropdesc, NULL, NULL))
            {

                cBuf.QuickAppend((ppropdesc->pstrExposedName)?
                    ppropdesc->pstrExposedName : ppropdesc->pstrName);
                strName = SysAllocString(cBuf);
            }
            else if(IsExpandoDISPID(dispid, &expDispid))
            {
                LPCTSTR pszName;
                if(S_OK == pStyleObj->GetExpandoName(expDispid, &pszName))
                {
                    cBuf.QuickAppend(pszName);
                    strName = SysAllocString(cBuf);
                }
            }
        }
    }
    else
    {
        AssertSz(FALSE, "must improve");
        strName = NULL;
    }

    // we have a property name, and can fire the event
    if(strName)
    {
        Fire_onpropertychange(strName);
        SysFreeString(strName);
    }

    return S_OK;
}

void CElement::Fire_onfocus(DWORD_PTR dwContext)
{
    CDocument* pDoc = Doc();
    if(!IsInMarkup() || !pDoc->_pPrimaryMarkup)
    {
        return;
    }

    CDocument::CLock LockForm(pDoc, FORMLOCK_CURRENT);
    CLock LockFocus(this, ELEMENTLOCK_FOCUS);
    FireEvent(DISPID_EVMETH_ONFOCUS, DISPID_EVPROP_ONFOCUS, _T("focus"), (BYTE*)VTS_NONE);
}

void CElement::Fire_onblur(DWORD_PTR dwContext)
{
    CDocument* pDoc = Doc();
    if(!IsInMarkup() || !pDoc->_pPrimaryMarkup)
    {
        return;
    }

    CDocument::CLock LockForm(pDoc, FORMLOCK_CURRENT);
    CLock LockBlur(this, ELEMENTLOCK_BLUR);
    pDoc->_fModalDialogInOnblur = (BOOL)dwContext;
    FireEvent(DISPID_EVMETH_ONBLUR, DISPID_EVPROP_ONBLUR, _T("blur"), (BYTE*)VTS_NONE);
    pDoc->_fModalDialogInOnblur = FALSE;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::DragElement
//
//  Synopsis:   calls Fire_ondragstart then executes DoDrag on the
//              layout
//
//-----------------------------------------------------------------------------
BOOL CElement::DragElement(CLayout* pLayout, DWORD dwStateKey,
                           IUniformResourceLocator* pUrlToDrag, long lSubDivision)
{
    CTreeNode::CLock Lock(GetFirstBranch());
    BOOL fRet = FALSE;
    CDocument* pDoc = Doc();

    Assert(!pDoc->_pDragStartInfo);
    pDoc->_pDragStartInfo = new CDragStartInfo(this, dwStateKey, pUrlToDrag);

    if (!pDoc->_pDragStartInfo)
        goto Cleanup;

    fRet = Fire_ondragstart(NULL, lSubDivision);

    if(!GetFirstBranch())
    {
        fRet = FALSE;
        goto Cleanup;
    }

    if(fRet)
    {
        Assert(pLayout->ElementOwner()->IsInMarkup());
        pLayout->DoDrag(dwStateKey, pUrlToDrag);
    }

Cleanup:
    if(pDoc->_pDragStartInfo)
    {
        delete pDoc->_pDragStartInfo;
        pDoc->_pDragStartInfo = NULL;
    }
    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::Fire_ondragenterover
//
//  Synopsis:   Fires the ondragenter or ondragover event, sets up the event param
//              and returns the dropeffect
//
//+----------------------------------------------------------------------------
BOOL CElement::Fire_ondragHelper(
        long    lSubDivision,
        DISPID  dispidEvent,
        DISPID  dispidProp,
        LPCTSTR pchType,
        DWORD*  pdwDropEffect)
{
    EVENTPARAM param(Doc(), TRUE);
    CVariant vb;
    BOOL fRet;
    CTreeNode* pNodeContext = GetFirstBranch();

    Assert(pdwDropEffect);

    param.dwDropEffect = *pdwDropEffect;

    param.SetNodeAndCalcCoordinates(pNodeContext);
    param.SetType(pchType);

    BubbleEventHelper(
        pNodeContext,
        lSubDivision,
        dispidEvent,
        dispidProp,
        FALSE,
        &vb,
        (BYTE*)VTS_NONE);

    fRet = !param.IsCancelled() && (V_VT(&vb)==VT_EMPTY || VB_TRUE==V_BOOL(&vb));

    if(!fRet)
    {
        *pdwDropEffect = param.dwDropEffect;
    }

    return fRet;
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::Fire_ondragend
//
//  Synopsis:   Fires the ondragend event
//
//+----------------------------------------------------------------------------
void CElement::Fire_ondragend(long lSubDivision, DWORD dwDropEffect)
{
    EVENTPARAM param(Doc(), TRUE);
    CTreeNode* pNodeContext = GetFirstBranch();

    param.dwDropEffect = dwDropEffect;

    param.SetNodeAndCalcCoordinates(pNodeContext);
    param.SetType(_T("dragend"));

    BubbleEventHelper(
        pNodeContext,
        lSubDivision,
        DISPID_EVMETH_ONDRAGEND,
        DISPID_EVPROP_ONDRAGEND,
        FALSE,
        NULL,
        (BYTE*)VTS_NONE);
}

//+------------------------------------------------------------------------
//
//  CElement::GetImageUrlCookie
//
//  Returns a Adds the specified URL to the url image cache on the doc
//
//-------------------------------------------------------------------------
HRESULT CElement::GetImageUrlCookie(LPCTSTR lpszURL, LONG* plCtxCookie)
{
    HRESULT         hr = S_OK;
    CDocument*      pDoc = Doc();
    LONG            lNewCookie = 0;

    // Element better be in the tree when this function is called
    Assert(pDoc);

    if(lpszURL && *lpszURL)
    {
        hr = pDoc->AddRefUrlImgCtx(lpszURL, this, &lNewCookie);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    *plCtxCookie = lNewCookie;

    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  CElement::AddImgCtx
//
//  Adds the info specified in ImgCtxInfo on the attr array, releasing
//  the current value if there is one.
//
//-------------------------------------------------------------------------
HRESULT CElement::AddImgCtx(DISPID dispID, LONG lCookie)
{
    HRESULT     hr = S_OK;
    CDocument*  pDoc = Doc();
    AAINDEX     iCookieIndex;

    iCookieIndex = FindAAIndex(dispID, CAttrValue::AA_Internal);

    if(iCookieIndex != AA_IDX_UNKNOWN)
    {
        // Remove the current entry
        DWORD dwCookieOld = 0;

        if(GetSimpleAt(iCookieIndex, &dwCookieOld) == S_OK)
        {
            pDoc->ReleaseUrlImgCtx(LONG(dwCookieOld), this);
        }
    }

    hr = AddSimple(dispID, DWORD(lCookie), CAttrValue::AA_Internal);

    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Function:   ReleaseImageCtxts
//
//  Synopsis:   Finds any image contexts associated with this element,
//              and frees them up.  Cookies can be held for LI bullets
//              and for background images.
//
//-----------------------------------------------------------------------------
void CElement::ReleaseImageCtxts()
{
    CDocument*  pDoc = Doc();
    AAINDEX     iCookieIndex;
    DWORD       dwCookieOld = 0;
    int         n;

    if(!_fHasImage)
    {
        return; // nothing to do, bail
    }

    for(n=0; n<ARRAYSIZE(s_aryImgDispID); ++n)
    {
        // Check for a bg url image cookie in the standard attr array
        iCookieIndex = FindAAIndex(s_aryImgDispID[n].cacheID, CAttrValue::AA_Internal);

        if(iCookieIndex!=AA_IDX_UNKNOWN && GetSimpleAt(iCookieIndex, &dwCookieOld)==S_OK)
        {
            pDoc->ReleaseUrlImgCtx((LONG)dwCookieOld, this);
        }
    }
}

//+----------------------------------------------------------------------------
//
//  Function:   DeleteImageCtx
//
//  Synopsis:   Finds any image contexts associated with this element,
//              corresponding to the dispid and free it up.  Cookies
//              can be held for LI bullets and for background images.
//
//-----------------------------------------------------------------------------
void CElement::DeleteImageCtx(DISPID dispid)
{
    CDocument* pDoc = Doc();
    CAttrArray* pAA;

    if(_fHasImage && (pAA=*GetAttrArray())!=NULL)
    {
        for(int n=0; n<ARRAYSIZE(s_aryImgDispID); ++n)
        {
            if(dispid == s_aryImgDispID[n].propID)
            {
                long lCookie;

                if(pAA->FindSimpleInt4AndDelete(s_aryImgDispID[n].cacheID, (DWORD*)&lCookie))
                {
                    // Release UrlImgCtxCacheEntry
                    pDoc->ReleaseUrlImgCtx(lCookie, this);
                }
                break;
            }
        }
    }
}

HRESULT CElement::EnsureFormatCacheChange(DWORD dwFlags)
{
    HRESULT hr = S_OK;

    // If we're not in the tree, it really isn't
    // very safe to do what we do.  Since putting
    // the element in the tree will call us again,
    // simply returning should be safe
    if(GetFirstBranch() == 0)
    {
        goto Cleanup;
    }

    if(dwFlags & (ELEMCHNG_CLEARCACHES|ELEMCHNG_CLEARFF))
    {
        hr = ClearRunCaches(dwFlags);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:

    RRETURN(hr);
}

BOOL CElement::IsFormatCacheValid()
{
    CTreeNode* pNode;

    pNode = GetFirstBranch();
    while(pNode)
    {
        if(!pNode->IsFancyFormatValid() || !pNode->IsCharFormatValid() || !pNode->IsParaFormatValid())
        {
            return FALSE;
        }
        pNode = pNode->NextBranch();
    }

    return TRUE;
}

BOOL CElement::IsAligned(void)
{
    return GetFirstBranch()->IsAligned();
}

BOOL CElement::IsContainer()
{
    return HasFlag(TAGDESC_CONTAINER);
}

BOOL CElement::IsNoScope()
{
    return HasFlag(TAGDESC_TEXTLESS);
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::IsBlockElement
//
//  Synopsis:   Describes whether or not this node is a block element
//
//  Returns:    BOOL indicating a block element
//
//-----------------------------------------------------------------------------
BOOL CElement::IsBlockElement()
{
    CTreeNode* pTreeNode = GetFirstBranch();

    if(pTreeNode->_iFF == -1)
    {
        pTreeNode->GetFancyFormat();
    }

    return BOOL(pTreeNode->_fBlockNess);
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::IsBlockTag
//
//  Synopsis:   Describes whether or not this element is a block tag
//              This should rarely be used - it returns the same value no
//              matter what the display: style setting on the element is.
//              To determine if you should break lines before and after the
//              element use IsBlockElement().
//
//  Returns:    BOOL indicating a block tag
//
//-----------------------------------------------------------------------------
BOOL CElement::IsBlockTag(void)
{
    return (HasFlag(TAGDESC_BLOCKELEMENT) || Tag()==ETAG_OBJECT);
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement::IsOwnLineElement
//
//  Synopsis:   Tells us if the element is a ownline element
//
//  Returns:    BOOL indicating an ownline element
//
//-----------------------------------------------------------------------------
BOOL CElement::IsOwnLineElement(CFlowLayout* pFlowLayoutContext)
{
    BOOL fRet;

    if(IsInlinedElement() && (HasFlag(TAGDESC_OWNLINE) || pFlowLayoutContext->IsElementBlockInContext(this)))
    {
        fRet = TRUE;
    }
    else
    {
        fRet = FALSE;
    }

    return fRet;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::BreaksLine
//
//  Synopsis:   Describes whether or not this node starts a new line
//
//  Returns:    BOOL indicating start of new line
//
//-------------------------------------------------------------------------
BOOL CElement::BreaksLine(void)
{
    return (IsBlockElement()&&!HasFlag(TAGDESC_WONTBREAKLINE));
}

HRESULT CElement::ClearRunCaches(DWORD dwFlags)
{
    CMarkup* pMarkup = GetMarkup();

    if(pMarkup)
    {
        pMarkup->ClearRunCaches(dwFlags, this);
    }

    RRETURN(S_OK);
}

//+------------------------------------------------------------------------
//
//  Member:     GetLastBranch
//
//  Synopsis:   Like GetFirstBranch, but gives the last one.
//
//-------------------------------------------------------------------------
CTreeNode* CElement::GetLastBranch()
{
    CTreeNode* pNode = GetFirstBranch();
    CTreeNode* pNodeLast = pNode;

    while(pNode)
    {
        pNodeLast = pNode;
        pNode = pNode->NextBranch();
    }

    return pNodeLast;
}

BOOL CElement::IsOverlapped()
{
    CTreeNode* pNode = GetFirstBranch();

    return pNode&&!pNode->IsLastBranch();
}

//+------------------------------------------------------------------------
//
//  Member:     GetTreeExtent
//
//  Synopsis:   Return the edge node pos' for this element.  Pretty
//              much just walks the context chain and gets the first
//              and last node pos'.
//
//-------------------------------------------------------------------------
void CElement::GetTreeExtent(CTreePos** pptpStart, CTreePos** pptpEnd)
{
    CTreeNode* pNodeCurr = GetFirstBranch();

    if(pptpStart)
    {
        *pptpStart = NULL;
    }

    if(pptpEnd)
    {
        *pptpEnd = NULL;
    }

    if(!pNodeCurr)
    {
        goto Cleanup;
    }

    Assert(!pNodeCurr->GetBeginPos()->IsUninit() && !pNodeCurr->GetEndPos()->IsUninit());

    if(pptpStart)
    {
        *pptpStart = pNodeCurr->GetBeginPos();

        Assert(*pptpStart);
        Assert((*pptpStart)->IsBeginNode() && (*pptpStart)->IsEdgeScope());
        Assert((*pptpStart)->Branch() == pNodeCurr);
    }

    if(pptpEnd)
    {
        while(pNodeCurr->NextBranch())
        {
            pNodeCurr = pNodeCurr->NextBranch();
        }

        Assert(pNodeCurr);

        *pptpEnd = pNodeCurr->GetEndPos();

        Assert(*pptpEnd);
        Assert((*pptpEnd)->IsEndNode() && (*pptpEnd)->IsEdgeScope());
        Assert((*pptpEnd)->Branch() == pNodeCurr);
    }

Cleanup:
    return;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::HasLayoutLazy
//
//  Synopsis:   test for layoutNess requires the formats to be computed,
//              verify that formats are computed and test for layoutness.
//
//----------------------------------------------------------------------------
BOOL CElement::HasLayoutLazy()
{
    Assert(!TestClassFlag(ELEMENTDESC_NOLAYOUT) && !_fSite);

    CTreeNode* pNode = GetFirstBranch();
    return pNode?pNode->GetFancyFormat()->_fHasLayout:!!HasLayoutPtr();
}

inline CLayout* CElement::GetLayoutLazy()
{
    Assert(!HasLayoutPtr());
    Assert(GetFirstBranch()->GetFancyFormat()->_fHasLayout == GetFirstBranch()->_fHasLayout);

    CreateLayout();

    return GetLayoutPtr();
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::CreateLayout
//
//  Synopsis:   Creates the layout object to be associated with the current element
//
//--------------------------------------------------------------------------
HRESULT CElement::CreateLayout()
{
    CLayout*    pLayout = NULL;
    HRESULT     hr = S_OK;

    Assert(!HasLayoutPtr() && !_fSite);

    pLayout = new C1DLayout(this);

    _fOwnsRuns = TRUE;

    if(!pLayout)
    {
        hr = E_OUTOFMEMORY;
    }
    else
    {
        SetLayoutPtr(pLayout);
        pLayout->Init();
    }

    RRETURN(hr);
}

CLayout* CElement::GetCurNearestLayout()
{
    CLayout* pLayout = GetCurLayout();

    return (pLayout ? pLayout : GetCurParentLayout());
}

CTreeNode* CElement::GetCurNearestLayoutNode()
{
    return (HasLayout() ? GetFirstBranch() : GetCurParentLayoutNode());
}

CElement* CElement::GetCurNearestLayoutElement()
{
    return (HasLayout() ? this : GetCurParentLayoutElement());
}

CLayout* CElement::GetCurParentLayout()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetCurParentLayout();
}

CTreeNode* CElement::GetCurParentLayoutNode()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetCurParentLayoutNode();
}

CElement* CElement::GetCurParentLayoutElement()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetCurParentLayoutElement();
}

CLayout* CElement::GetUpdatedLayout()
{
    return NeedsLayout()?GetUpdatedLayoutPtr():NULL;
}

CLayout* CElement::GetUpdatedLayoutPtr()
{
    Assert(NeedsLayout());
    return HasLayoutPtr()?GetLayoutPtr():GetLayoutLazy();
}

BOOL CElement::NeedsLayout()
{
    if(_fLayoutAlwaysValid)
    {
        return HasLayoutPtr();
    }

    CTreeNode* pNode = GetFirstBranch();
    if(pNode && pNode->_iFF!=-1)
    {
        return pNode->_fHasLayout;
    }

    return HasLayoutLazy();
}

CLayout* CElement::GetUpdatedNearestLayout()
{
    CLayout* pLayout = GetUpdatedLayout();

    return (pLayout ? pLayout : GetUpdatedParentLayout());
}

CTreeNode* CElement::GetUpdatedNearestLayoutNode()
{
    return (NeedsLayout() ? GetFirstBranch() : GetUpdatedParentLayoutNode());
}

CElement* CElement::GetUpdatedNearestLayoutElement()
{
    return (NeedsLayout() ? this : GetUpdatedParentLayoutElement());
}

CLayout* CElement::GetUpdatedParentLayout()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetUpdatedParentLayout();
}

CTreeNode* CElement::GetUpdatedParentLayoutNode()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetUpdatedParentLayoutNode();
}

CElement* CElement::GetUpdatedParentLayoutElement()
{
    Assert(GetFirstBranch());

    return GetFirstBranch()->GetCurParentLayoutElement();
}

//+----------------------------------------------------------------------------
//  Member:     DestroyLayout
//
//  Synopsis:   Destroy the current layout attached to the element. This is
//              called from CFlowLayout::DoLayout to destroy the layout lazily
//              when an element loses layoutness.
//
//-----------------------------------------------------------------------------
void CElement::DestroyLayout()
{
    CLayout* pLayout = GetCurLayout();

    Assert(HasLayout() && !NeedsLayout());
    Assert(!_fSite && !TestClassFlag(CElement::ELEMENTDESC_NOLAYOUT));

    pLayout->ElementContent()->_fOwnsRuns = FALSE;

    Verify(Doc()->OpenView());
    Verify(pLayout == DelLayoutPtr());

    pLayout->Reset(TRUE);
    pLayout->Detach();
    pLayout->Release();
}

CElement* CElement::GetParentAncestorSafe(ELEMENT_TAG etag) const
{
    CTreeNode*  pNode = GetFirstBranch();
    CElement*   p = NULL;
    if(pNode)
    {
        pNode = pNode->Parent();
        if(pNode)
        {
            pNode = pNode->Ancestor(etag);
            if(pNode)
            {
                p = pNode->Element();
            }
        }
    }
    return p;
}

CElement* CElement::GetParentAncestorSafe(ELEMENT_TAG* arytag) const
{
    CTreeNode*  pNode = GetFirstBranch();
    CElement*   p = NULL;
    if(pNode)
    {
        pNode = pNode->Parent();
        if(pNode)
        {
            pNode = pNode->Ancestor(arytag);
            if(pNode)
            {
                p = pNode->Element();
            }
        }
    }
    return p;
}

CFlowLayout* CElement::GetFlowLayout()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetFlowLayout();
}

CTreeNode* CElement::GetFlowLayoutNode()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetFlowLayoutNode();
}

CElement* CElement::GetFlowLayoutElement()
{
    if(!GetFirstBranch())
    {
        return NULL;
    }

    return GetFirstBranch()->GetFlowLayoutElement();
}

CFlowLayout* CElement::HasFlowLayout()
{
    CLayout* pLayout = GetUpdatedLayout();

    return (pLayout ? pLayout->IsFlowLayout() : NULL);
}

//+----------------------------------------------------------------------------
//
//  Member:     InvalidateElement
//              MinMaxElement
//              ResizeElement
//              RemeasureElement
//              RemeasureInParentContext
//              RepositionElement
//              ZChangeElement
//
//  Synopsis:   Notfication send helpers
//
//  Arguments:  grfFlags - NFLAGS_xxxx flags
//
//-----------------------------------------------------------------------------
void CElement::InvalidateElement(DWORD grfFlags)
{
    // this notification goes up to the parent to request that 
    // we be invalidated
    SendNotification(NTYPE_ELEMENT_INVALIDATE, grfFlags);
}

void CElement::MinMaxElement(DWORD grfFlags)
{
    CLayout* pLayout = GetCurLayout();

    if(pLayout && pLayout->_fMinMaxValid)
    {
        SendNotification(NTYPE_ELEMENT_MINMAX, grfFlags);
    }
}

void CElement::ResizeElement(DWORD grfFlags)
{
    //  Resize notifications are only fired when:
    //    a) The element does not have a layout (and must always notify its container) or
    //    b) It has a layout, but it's currently "clean" and
    //    c) The element is not presently being sized by its container
    CLayout* pLayout = GetCurLayout();

    if(!pLayout || (!pLayout->_fSizeThis && !pLayout->IsCalcingSize()))
    {
        SendNotification(NTYPE_ELEMENT_RESIZE, grfFlags);
    }
}

void CElement::RemeasureInParentContext(DWORD grfFlags)
{
    SendNotification(NTYPE_ELEMENT_RESIZEANDREMEASURE, grfFlags);
}

void CElement::RemeasureElement(DWORD grfFlags)
{
    SendNotification(NTYPE_ELEMENT_REMEASURE, grfFlags);
}

void CElement::RepositionElement(DWORD grfFlags)
{
    Assert(!IsPositionStatic());
    SendNotification(NTYPE_ELEMENT_REPOSITION, grfFlags);
}

void CElement::ZChangeElement(DWORD grfFlags, CPoint* ppt)
{
    CMarkup*        pMarkup = GetMarkup();
    CNotification   nf;

    Assert(!IsPositionStatic() || GetFirstBranch()->GetCharFormat()->_fRelative);
    Assert(pMarkup);

    nf.Initialize(NTYPE_ELEMENT_ZCHANGE, this, GetFirstBranch(), NULL, grfFlags);
    if(ppt)
    {
        nf.SetData(*ppt);
    }

    pMarkup->Notify(nf);
}

//+----------------------------------------------------------------------------
//
//  Member:     SendNotification
//
//  Synopsis:   Send a notification associated with this element
//
//-----------------------------------------------------------------------------
void CElement::SendNotification(CNotification* pNF)
{
    CMarkup* pMarkup = GetMarkup();

    if(pMarkup)
    {
        pMarkup->Notify(pNF);
    }
}

//+----------------------------------------------------------------------------
//
//  Member:     SendNotification
//
//  Synopsis:   Send a notification associated with this element
//
//  Arguments:  ntype    - NTYPE_xxxxx flag
//              grfFlags - NFLAGS_xxxx flags
//
//-----------------------------------------------------------------------------
void CElement::SendNotification(NOTIFYTYPE  ntype, DWORD grfFlags, void* pvData)
{
    CMarkup* pMarkup = GetMarkup();

    if(pMarkup)
    {
        CNotification nf;

        Assert(GetFirstBranch());

        nf.Initialize(ntype, this, GetFirstBranch(), pvData, grfFlags);

        pMarkup->Notify(nf);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     GetSourceIndex
//
//-------------------------------------------------------------------------
long CElement::GetSourceIndex()
{
    CTreeNode* pNodeCurr;

    if(Tag() == ETAG_ROOT)
    {
        return -1;
    }

    pNodeCurr = GetFirstBranch();
    if(!pNodeCurr)
    {
        return -1;
    }
    else
    {
        Assert(!pNodeCurr->GetBeginPos()->IsUninit());
        return pNodeCurr->GetBeginPos()->SourceIndex()-1; // subtract one because of ETAG_ROOT
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CompareZOrder
//
//  Synopsis:   Compare the z-order of two elements
//
//  Arguments:  pElement - The CElement to compare against
//
//  Returns:    Greater than zero if this element is greater
//              Less than zero if this element is less
//              Zero if they are equal
//
//----------------------------------------------------------------------------
long CElement::CompareZOrder(CElement* pElement)
{
    long lCompare;

    lCompare = GetFirstBranch()->GetCascadedzIndex() - pElement->GetFirstBranch()->GetCascadedzIndex();

    if(!lCompare)
    {
        lCompare = GetSourceIndex() - pElement->GetSourceIndex();
    }

    return lCompare;
}

//+----------------------------------------------------------------------------
//
//  Member:     DirtyLayout
//
//  Synopsis:   Dirty the layout engine associated with an element
//
//-----------------------------------------------------------------------------
void CElement::DirtyLayout(DWORD grfLayout)
{
    if(HasLayout())
    {
        CLayout* pLayout = GetCurLayout();

        Assert(pLayout);

        pLayout->_fSizeThis = TRUE;

        if(grfLayout & LAYOUT_FORCE)
        {
            pLayout->_fForceLayout = TRUE;
        }
    }
}

//+----------------------------------------------------------------------------
//
//  Member:     OpenView
//
//  Synopsis:   Open the view associated with the element - That view is the one
//              associated with the nearest layout
//
//  Returns:    TRUE if the view was successfully opened, FALSE if we're in the
//              middle of rendering
//
//-----------------------------------------------------------------------------
BOOL CElement::OpenView()
{
    CLayout* pLayout = GetCurNearestLayout();
    return pLayout?pLayout->OpenView():Doc()->OpenView();
}

//+---------------------------------------------------------------------------
//
// Method:      CElement::HaPercentBgImg
//
// Synopsis:    Does this element have a background image whose width or
//              height is percent based.
//
//----------------------------------------------------------------------------
BOOL CElement::HasPercentBgImg()
{
    CImgCtx* pImgCtx = GetBgImgCtx();
    const CFancyFormat* pFF = GetFirstBranch()->GetFancyFormat();

    return pImgCtx&&
        (pFF->_cuvBgPosX.GetUnitType()==CUnitValue::UNIT_PERCENT||
        pFF->_cuvBgPosY.GetUnitType()==CUnitValue::UNIT_PERCENT);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetFirstCp
//
//  Synopsis:   Get the first character position of this element in the text
//              flow. (relative to the Markup)
//
//  Returns:    LONG        - start character position in the flow. -1 if the
//                            element is not found in the tree
//
//----------------------------------------------------------------------------
long CElement::GetFirstCp()
{
    CTreePos* ptpStart;

    GetTreeExtent(&ptpStart, NULL);

    return ptpStart?ptpStart->GetCp()+1:-1;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetLastCp
//
//  Synopsis:   Get the last character position of this element in the text
//              flow. (relative to the Markup)
//
//  Returns:    LONG        - end character position in the flow. -1 if the
//                            element is not found in the tree
//
//----------------------------------------------------------------------------
long CElement::GetLastCp()
{
    CTreePos* ptpEnd;

    GetTreeExtent(NULL, &ptpEnd);

    return ptpEnd?ptpEnd->GetCp():-1;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetElementCch
//
//  Synopsis:   Get the number of charactersinfluenced by this element
//
//  Returns:    LONG        - no of characters, 0 if the element is not in the
//                            the tree.
//
//----------------------------------------------------------------------------
inline long CElement::GetElementCch()
{
    long cpStart, cpFinish;

    return GetFirstAndLastCp(&cpStart, &cpFinish);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetFirstAndLastCp
//
//  Synopsis:   Get the first and last character position of this element in
//              the text flow. (relative to the Markup)
//
//  Returns:    no of characters influenced by the element, 0 if the element
//              is not found in the tree.
//
//----------------------------------------------------------------------------
LONG CElement::GetFirstAndLastCp(long* pcpFirst, long* pcpLast)
{
    CTreePos *ptpStart, *ptpLast;
    long cpFirst, cpLast;

    Assert(pcpFirst || pcpLast);

    if(!pcpFirst)
    {
        pcpFirst = &cpFirst;
    }

    if(!pcpLast)
    {
        pcpLast = &cpLast;
    }

    GetTreeExtent(&ptpStart, &ptpLast);

    Assert((ptpStart&&ptpLast) || (!ptpStart && !ptpLast));

    if(ptpStart)
    {
        *pcpFirst = ptpStart->GetCp() + 1;
        *pcpLast = ptpLast->GetCp();
    }
    else
    {
        *pcpFirst = *pcpLast = 0;
    }

    Assert(*pcpLast-*pcpFirst >= 0);

    return *pcpLast-*pcpFirst;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetBorderInfo
//
//  Synopsis:   get the elements border information.
//
//  Arguments:  pdci        - Current CDocInfo
//              pborderinfo - pointer to return the border information
//              fAll        - (FALSE by default) return all border related
//                            related information (border style's etc.),
//                            if TRUE
//
//  Returns:    0 - if no borders
//              1 - if simple border (all sides present, all the same size)
//              2 - if complex border (present, but not simple)
//
//----------------------------------------------------------------------------
DWORD CElement::GetBorderInfo(CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll)
{
    return GetBorderInfoHelper(GetFirstBranch(), pdci, pborderinfo, fAll);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetRange
//
//  Synopsis:   Returns the range of char's under this element including the
//              end nodes
//
//----------------------------------------------------------------------------
HRESULT CElement::GetRange(long* pcpStart, long* pcch)
{
    CTreePos *ptpStart, *ptpEnd;

    GetTreeExtent(&ptpStart, &ptpEnd);

    // The range returned include the WCH_NODE characters for the element
    *pcpStart = ptpStart->GetCp();
    *pcch = ptpEnd->GetCp() - *pcpStart + 1;
    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetPlainTextInScope
//
//  Synopsis:   Returns a text string containing all the plain text in the
//              scope of this element. The caller must free the memory.
//              This function can be used to merely retrieve the text
//              length by setting ppchText to NULL.
//
//  Arguments:  pstrText    If NULL,
//                              no text is returned.
//                          If not NULL but there is no text,
//                              *pstrText is set to NULL.
//                          Otherwise
//                              *pstrText points to a new CStr
//
//-------------------------------------------------------------------------
HRESULT CElement::GetPlainTextInScope(CString* pstrText)
{
    HRESULT hr = S_OK;
    long    cp, cch;

    Assert(pstrText);

    if(!IsInMarkup())
    {
        pstrText->Set(NULL);
        goto Cleanup;
    }

    cp = GetFirstCp();
    cch = GetElementCch();

    {
        CTxtPtr tp(GetMarkup(), cp);

        cch = tp.GetPlainTextLength(cch);

        // copy text into buffer
        pstrText->SetLengthNoAlloc(0);
        pstrText->ReAlloc(cch);

        cch = tp.GetPlainText(cch, (LPTSTR)*pstrText);

        Assert(cch >= 0);

        if(cch)
        {
            // Terminate with 0. GetPlainText() does not seem to do this.
            pstrText->SetLengthNoAlloc(cch);
            *(LPTSTR(*pstrText)+cch) = 0;
        }
        else
        {
            // just making sure...
            pstrText->Free();
        }
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::Save
//
//  Synopsis:   Save the element to the stream
//
//-------------------------------------------------------------------------
HRESULT CElement::Save(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd)
{
    HRESULT hr;

    if(fEnd && HasFlag(TAGDESC_SAVEINDENT))
    {
        pStreamWrBuff->EndIndent();
    }

    hr = WriteTag(pStreamWrBuff, fEnd);
    if(hr)
    {
        goto Cleanup;
    }

    if(!fEnd && HasFlag(TAGDESC_SAVEINDENT))
    {
        pStreamWrBuff->BeginIndent();
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Function:   WriteTag
//
//  Synopsis:   writes an open/end tag to the stream buffer
//
//  Arguments:  pStreamWrBuff   -   stream buffer
//              fEnd            -   TRUE if End tag is to be written out
//
//  Returns:    S_OK    if  successful
//
//-----------------------------------------------------------------------------
HRESULT CElement::WriteTag(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd, BOOL fForce)
{
    HRESULT         hr = S_OK;
    DWORD           dwOldFlags = pStreamWrBuff->ClearFlags(WBF_ENTITYREF);
    const TCHAR*    pszTagName = TagName();
    const TCHAR*    pszScopeName;
    ELEMENT_TAG     etag = Tag();

    // Do not write tags out in plaintext mode or when we are
    // explicitly asked not to.
    if(pStreamWrBuff->TestFlag(WBF_SAVE_PLAINTEXT)
        || !pszTagName[0] ||
        (this==pStreamWrBuff->GetElementContext()
        && pStreamWrBuff->TestFlag(WBF_NO_TAG_FOR_CONTEXT))
        || (fEnd
        && (TagHasNoEndTag(Tag())
        || (!_fExplicitEndTag
        && !HasFlag(TAGDESC_SAVEALWAYSEND)
        && !fForce))
        )
        )
    {
        goto Cleanup;
    }

    if(fEnd)
    {
        hr = pStreamWrBuff->Write(_T("</"), 2);
    }
    else
    {
        if(pStreamWrBuff->TestFlag(WBF_FOR_RTF_CONV) && Tag()==ETAG_DIV)
        {
            // For the RTF converter, transform DIV tags into P tags.
            pszTagName = SZTAG_P;
        }

        // NOTE: In IE4, we would save a NewLine before every
        // element that was a ped.  This is roughly equivalent
        // to saving it before every element that is a container now
        // However, this does not round trip properly so I'm taking
        // that check out.
        if(!pStreamWrBuff->TestFlag(WBF_NO_PRETTY_CRLF)
            && (HasFlag(TAGDESC_SAVETAGOWNLINE)||IsBlockTag()))
        {
            hr = pStreamWrBuff->NewLine();
            if(hr)
            {
                goto Cleanup;
            }
        }
        hr = pStreamWrBuff->Write(_T("<"), 1);
    }

    if(hr)
    {
        goto Cleanup;
    }

    pszScopeName = pStreamWrBuff->TestFlag(WBF_SAVE_FOR_PRINTDOC) && 
        pStreamWrBuff->TestFlag(WBF_SAVE_FOR_XML) &&
        Tag()!=ETAG_GENERIC ? NamespaceHtml() : NULL;
    if(pszScopeName)
    {
        hr = pStreamWrBuff->Write(pszScopeName);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pStreamWrBuff->Write(_T(":"));
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = pStreamWrBuff->Write(pszTagName);
    if(hr)
    {
        goto Cleanup;
    }

    if(!fEnd)
    {
        BOOL fAny;

        hr = SaveAttributes(pStreamWrBuff, &fAny);
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = pStreamWrBuff->Write(_T(">"), 1);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    pStreamWrBuff->RestoreFlags(dwOldFlags);
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::SaveAttributes
//
//  Synopsis:   Save the attributes to the stream
//
//-------------------------------------------------------------------------
HRESULT CElement::SaveAttributes(CStreamWriteBuff* pStreamWrBuff, BOOL* pfAny)
{
    const PROPERTYDESC* const* pppropdescs;
    HRESULT hr = S_OK;
    BOOL fSave;
    LPCTSTR lpstrUnknownValue;
    CBase* pBaseObj;
    BOOL fAny = FALSE;

    if(GetPropDescArray())
    {
        for(pppropdescs=GetPropDescArray(); *pppropdescs; pppropdescs++)
        {
            const PROPERTYDESC* ppropdesc = (*pppropdescs);
            // BUGBUG for now check for the method pointer because of old property implementation...
            if(!ppropdesc->pfnHandleProperty)
            {
                continue;
            }

            pBaseObj = GetBaseObjectFor(ppropdesc->GetDispid());

            if(!pBaseObj)
            {
                continue;
            }

            lpstrUnknownValue = NULL;
            if(ppropdesc->GetPPFlags() & PROPPARAM_ATTRARRAY)
            {
                AAINDEX aaIx = AA_IDX_UNKNOWN;
                CAttrValue* pAV = NULL;
                CAttrArray* pAA = *(pBaseObj->GetAttrArray());

                if(pAA)
                {
                    pAV = pAA->Find(ppropdesc->GetDispid(), CAttrValue::AA_Attribute, &aaIx);
                }

                if(pAA && (!pAV || pAV->IsDefault()))
                {
                    if(pAV)
                    {
                        aaIx++;
                    }

                    pAV = pAA->FindAt(aaIx);
                    if(pAV)
                    {
                        if((pAV->GetDISPID()==ppropdesc->GetDispid()) &&
                            (pAV->GetAAType()==CAttrValue::AA_UnknownAttr))
                        {
                            // Unknown attrs are always strings
                            lpstrUnknownValue = pAV->GetLPWSTR();
                        }
                        else
                        {
                            pAV = NULL;
                        }
                    }
                }
                fSave = !!pAV;
            }
            else
            {
                // Save the property if it was not the same as the default.
                // Do not save if we got an error retrieving it.
                fSave = ppropdesc->HandleCompare(pBaseObj,
                    (void*)&ppropdesc->ulTagNotPresentDefault) == S_FALSE;
            }

            if(fSave)
            {
                fAny = TRUE;

                if(lpstrUnknownValue)
                {
                    hr = SaveAttribute(
                        pStreamWrBuff,
                        (LPTSTR)ppropdesc->pstrName,
                        (LPTSTR)lpstrUnknownValue,  // pchValue
                        NULL,                       // ppropdesc
                        NULL,                       // pBaseObj
                        FALSE);                     // fEqualSpaces
                }
                else
                {
                    if(ppropdesc->IsBOOLProperty())
                    {
                        hr = SaveAttribute(
                            pStreamWrBuff,
                            (LPTSTR)ppropdesc->pstrName,
                            NULL,                       // pchValue
                            NULL,                       // ppropdesc
                            NULL,                       // pBaseObj
                            FALSE);                     // fEqualSpaces
                    }
                    else
                    {
                        hr = SaveAttribute(
                            pStreamWrBuff,
                            (LPTSTR)ppropdesc->pstrName,
                            NULL,                       // pchValue
                            ppropdesc,                  // ppropdesc
                            pBaseObj,                   // pBaseObj
                            FALSE);                     // fEqualSpaces
                    }
                }
            }
        }
    }

    hr = SaveUnknown(pStreamWrBuff, fAny ? NULL : &fAny);
    if(hr)
    {
        goto Cleanup;
    }

    /*if(HasPeerHolder())
    {
        GetPeerHolder()->SaveMulti(pStreamWrBuff, fAny?NULL:&fAny);
    } wlw note*/

    if(pfAny)
    {
        *pfAny = fAny;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::SaveAttributes
//
//  Synopsis:   Save the attributes into property bag
//
//-------------------------------------------------------------------------
HRESULT CElement::SaveAttributes(IPropertyBag* pPropBag, BOOL fSaveBlankAttributes)
{
    const PROPERTYDESC* const* pppropdescs;
    HRESULT     hr = S_OK;
    CVariant    Var;
    BOOL        fSave;
    CBase*      pBaseObj;

    if(GetPropDescArray())
    {
        for(pppropdescs=GetPropDescArray(); *pppropdescs; pppropdescs++)
        {
            const PROPERTYDESC* ppropdesc = (*pppropdescs);
            // BUGBUG for now check for the method pointer because of old property implementation...
            if(!ppropdesc->pfnHandleProperty)
            {
                continue;
            }

            pBaseObj = GetBaseObjectFor(ppropdesc->GetDispid());

            if(!pBaseObj)
            {
                continue;
            }

            if(ppropdesc->GetPPFlags() & PROPPARAM_ATTRARRAY)
            {
                AAINDEX aaIx;
                aaIx = pBaseObj->FindAAIndex(ppropdesc->GetDispid(), CAttrValue::AA_Attribute);
                fSave = (aaIx==AA_IDX_UNKNOWN) ? FALSE : TRUE;
            }
            else
            {
                // Save the property if it was not the same as the default.
                // Do not save if we got an error retrieving it.
                fSave = ppropdesc->HandleCompare(pBaseObj, (void*)&ppropdesc->ulTagNotPresentDefault) == S_FALSE;
            }

            if(fSave)
            {
                // If we're dealing with a BOOL type, don't put a value
                if(ppropdesc->IsBOOLProperty())
                {
                    // Boolean (flag), skip the =<val>
                    Var.vt = VT_EMPTY;
                }
                else
                {
                    hr = ppropdesc->HandleGetIntoBSTR(pBaseObj, &V_BSTR(&Var));
                    if(hr)
                    {
                        continue;
                    }
                    V_VT(&Var) = VT_BSTR;
                }
                hr = pPropBag->Write(ppropdesc->pstrName, &Var);
                if(hr)
                {
                    goto Cleanup;
                }
                // if Var has an Allocated value, we need to free it before
                // going around the loop again.
                VariantClear(&Var);
            }
        }
    }

    hr = SaveUnknown(pPropBag, fSaveBlankAttributes);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::SaveAttribute
//
//  Synopsis:   Save a single attribute to the stream
//
//-------------------------------------------------------------------------
HRESULT CElement::SaveAttribute(
         CStreamWriteBuff*      pStreamWrBuff,
         LPTSTR                 pchName,
         LPTSTR                 pchValue,
         const PROPERTYDESC*    pPropDesc/*=NULL*/,
         CBase*                 pBaseObj/*=NULL*/,
         BOOL                   fEqualSpaces/*=TRUE*/,  // BUGBUG (dbau) fix all the test cases so that it never has spaces
         BOOL                   fAlwaysQuote/*=FALSE*/)
{
    HRESULT hr;
    DWORD   dwOldFlags;

    hr = pStreamWrBuff->Write(_T(" "), 1);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pStreamWrBuff->Write(pchName, _tcslen(pchName));
    if(hr)
    {
        goto Cleanup;
    }

    if(pchValue || pPropDesc)
    {
        // Quotes are necessary for pages like ASP, that might have
        // <% =x %>. This will screw up the parser if we don't output
        // the quotes around such ASP expressions.
        BOOL fForceQuotes = fAlwaysQuote || !pchValue || !pchValue[0] || 
            (pchValue && (StrChr(pchValue, _T('<')) || StrChr(pchValue, _T('>'))));

        if(fEqualSpaces)
        {
            hr = pStreamWrBuff->Write(_T(" = "));
        }
        else
        {
            hr = pStreamWrBuff->Write(_T("="));
        }

        if(hr)
        {
            goto Cleanup;
        }

        // We dont want to break the line in the middle of an attribute value
        dwOldFlags = pStreamWrBuff->SetFlags(WBF_NO_WRAP);

        if(pchValue)
        {
            hr = pStreamWrBuff->WriteQuotedText(pchValue, fForceQuotes);
        }
        else
        {
            Assert(pPropDesc && pBaseObj);
            hr = pPropDesc->HandleSaveToHTMLStream(pBaseObj, (void*)pStreamWrBuff);
        }
        if(hr)
        {
            goto Cleanup;
        }

        pStreamWrBuff->RestoreFlags(dwOldFlags);
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::TagName
//
//  Synopsis:   Chases the proxy and the returns the tag name of
//
//  Returns:    const TCHAR *
//
//-------------------------------------------------------------------------
const TCHAR* CElement::TagName()
{
    return NameFromEtag(Tag());
}

const TCHAR* CElement::NamespaceHtml()
{
    return _T("HTML");
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::CanShow
//
//  Synopsis:   Determines whether an element can be shown at this moment.
//
//--------------------------------------------------------------------------
BOOL CElement::CanShow()
{
    BOOL fRet = TRUE;
    CTreeNode* pNodeSite = GetFirstBranch()->GetCurNearestLayoutNode();

    while(pNodeSite)
    {
        if(!pNodeSite->Element()->GetInfo(GETINFO_ISCOMPLETED))
        {
            fRet = FALSE;
            break;
        }
        pNodeSite = pNodeSite->GetCurParentLayoutNode();
    }
    return fRet;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::HasFlag
//
//  Synopsis:   Checks if the element has an given tag
//
//  Returns:    TRUE if it has an end tag else FALSE
//
//-------------------------------------------------------------------------
BOOL CElement::HasFlag(TAGDESC_FLAGS flag) const
{
    const CTagDesc* ptd = TagDescFromEtag(Tag());
    return (ptd ? (ptd->_dwTagDescFlags&flag)?TRUE:ptd->HasFlag(flag) : FALSE);
}

//+------------------------------------------------------------------------
//
//  Static Member:  CElement::ReplacePtr, CElement::ClearPtr
//
//  Synopsis:   Do a CElement* assignment, but worry about refcounts
//
//-------------------------------------------------------------------------
void CElement::ReplacePtr(CElement** pplhs, CElement* prhs)
{
    if(pplhs)
    {
        if(prhs)
        {
            prhs->AddRef();
        }
        if(*pplhs)
        {
            (*pplhs)->Release();
        }
        *pplhs = prhs;
    }
}

//+------------------------------------------------------------------------
//
//  Static Member:  CElement::ReplacePtrSub, CElement::ClearPtr
//
//  Synopsis:   Do a CElement* assignment, but worry about weak refcounts
//
//-------------------------------------------------------------------------
void CElement::ReplacePtrSub(CElement** pplhs, CElement* prhs)
{
    if(pplhs)
    {
        if(prhs)
        {
            prhs->SubAddRef();
        }
        if(*pplhs)
        {
            (*pplhs)->SubRelease();
        }
        *pplhs = prhs;
    }
}

void CElement::SetPtr(CElement** pplhs, CElement* prhs)
{
    if(pplhs)
    {
        if(prhs)
        {
            prhs->AddRef();
        }
        *pplhs = prhs;
    }
}

void CElement::StealPtrSet(CElement** pplhs, CElement* prhs)
{
    SetPtr(pplhs, prhs);

    if(pplhs && *pplhs)
    {
        (*pplhs)->Release();
    }
}

void CElement::StealPtrReplace(CElement** pplhs, CElement* prhs)
{
    ReplacePtr(pplhs, prhs);

    if(pplhs && *pplhs)
    {
        (*pplhs)->Release();
    }
}

void CElement::ClearPtr(CElement** pplhs)
{
    if(pplhs && *pplhs)
    {
        CElement* pElement = *pplhs;
        *pplhs = NULL;
        pElement->Release();
    }
}

void CElement::ReleasePtr(CElement* pElement)
{
    if(pElement)
    {
        pElement->Release();
    }
}

//+------------------------------------------------------------------------
//
// Member:     CElement::SaveUnknown
//
// Synopsis:   Write these guys out
//
// Returns:    HRESULT
//
//+------------------------------------------------------------------------
HRESULT CElement::SaveUnknown(CStreamWriteBuff* pStreamWrBuff, BOOL* pfAny)
{
    HRESULT hr = S_OK;
    AAINDEX aaix = AA_IDX_UNKNOWN;
    LPCTSTR lpPropName;
    LPCTSTR lpszValue = NULL;
    BSTR bstrTemp = NULL;
    DISPID expandoDISPID;
    BOOL fAny = FALSE;

    // Look for all expandos & dump them out
    while((aaix=FindAAType(CAttrValue::AA_Expando, aaix)) != AA_IDX_UNKNOWN)
    {
        CAttrValue* pAV = _pAA->FindAt(aaix);

        Assert(pAV);

        // Get value into a string, but skip VT_DISPATCH & VT_UNKNOWN
        if(pAV->GetAVType()==VT_DISPATCH || pAV->GetAVType()==VT_UNKNOWN)
        {
            continue;
        }

        // BUGBUG rgardner - we should smarten this up so we don't need to allocate a string

        // Found a literal attrValue
        hr = pAV->GetIntoString( &bstrTemp, &lpszValue );

        if(hr == S_FALSE)
        {
            // Can't convert to string
            continue;
        }
        else if(hr)
        {
            goto Cleanup;
        }

        fAny = TRUE;

        expandoDISPID = GetDispIDAt(aaix);
        if(TestClassFlag(ELEMENTDESC_OLESITE))
        {
            expandoDISPID = expandoDISPID + DISPID_EXPANDO_BASE - DISPID_ACTIVEX_EXPANDO_BASE;
        }

        hr = GetExpandoName(expandoDISPID, &lpPropName);
        if(hr)
        {
            goto Cleanup;
        }

        hr = SaveAttribute(pStreamWrBuff, (LPTSTR)lpPropName, (LPTSTR)lpszValue, NULL, NULL, FALSE, TRUE); // Always quote value: IE5 57717
        if(hr)
        {
            goto Cleanup;
        }

        if(bstrTemp)
        {
            SysFreeString(bstrTemp);
            bstrTemp = NULL;
        }
    }

    if(pfAny)
    {
        *pfAny = fAny;
    }

Cleanup:
    if(bstrTemp)
    {
        FormsFreeString(bstrTemp);
    }
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
// Member:     CElement::SaveUnknown
//
// Synopsis:   Write these guys out
//
// Returns:    HRESULT
//
//+------------------------------------------------------------------------
HRESULT CElement::SaveUnknown(IPropertyBag* pPropBag, BOOL fSaveBlankAttributes)
{
    HRESULT     hr = S_OK;
    AAINDEX     aaix = AA_IDX_UNKNOWN;
    LPCTSTR     lpPropName;
    CVariant    var;
    DISPID      dispidExpando;

    while((aaix=FindAAType(CAttrValue::AA_Expando, aaix)) != AA_IDX_UNKNOWN)
    {
        var.vt = VT_EMPTY;

        hr = GetIntoBSTRAt(aaix, &(var.bstrVal));
        if(hr == S_FALSE)
        {
            // Can't convert to string
            continue;
        }
        else if(hr)
        {
            goto Cleanup;
        }

        // We do not save attributes with null string values for netscape compatibility of
        // <EMBED src=thisthat loop> attributes which have no value - the loop attribute
        // in the example.   Pluginst.cxx passes in FALSE for fSaveBlankAttributes, everybody
        // else passes in TRUE via a default param value.
        if(!fSaveBlankAttributes && (var.bstrVal==NULL || *var.bstrVal==_T('\0')))
        {
            VariantClear(&var);
            continue;
        }

        var.vt = VT_BSTR;

        dispidExpando = GetDispIDAt(aaix);
        if(TestClassFlag(ELEMENTDESC_OLESITE))
        {
            dispidExpando = dispidExpando + DISPID_EXPANDO_BASE - DISPID_ACTIVEX_EXPANDO_BASE;
        }
        hr = GetExpandoName(dispidExpando, &lpPropName);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pPropBag->Write(lpPropName, &var);
        if(hr)
        {
            goto Cleanup;
        }

        VariantClear(&var);
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::NameOrIDOfParentForm
//
//  Synopsis:   Return the name or id of a parent form if one exists.
//              NULL if not.
//
//-------------------------------------------------------------------------
LPCTSTR CElement::NameOrIDOfParentForm()
{
    CElement*   pElementForm;
    LPCTSTR     pchName = NULL;

    pElementForm = GetFirstBranch()->SearchBranchToRootForTag(ETAG_FORM)->SafeElement();

    if(pElementForm)
    {
        pchName = pElementForm->GetIdentifier();
    }
    return pchName;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetIdentifier
//
//  This fn looks in the current elements' attr array & picks out
// this
//-------------------------------------------------------------------------
LPCTSTR CElement::GetIdentifier(void)
{
    LPCTSTR     pStr;
    CAttrArray* pAA;

    if(!IsNamed())
    {
        return NULL;
    }

    pAA = *GetAttrArray();

    // We're leveraging the fact that we know the dispids of name & ID, & we
    // also have a single _pAA in CElement containing all the attributes
    if(pAA && pAA->HasAnyAttribute() &&
        ((pAA->FindString(STDPROPID_XOBJ_NAME, &pStr) && pStr) ||
        (pAA->FindString(DISPID_CElement_id, &pStr) && pStr) ||
        pAA->FindString(DISPID_CElement_uniqueName, &pStr)))
    {
        // This looks dodgy but is safe as long as the return value is treated
        // as a temporary value that is used immediatly, then discarded
        return pStr;
    }
    else
    {
        return NULL;
    }
}

HRESULT CElement::GetUniqueIdentifier(CString* pcstr, BOOL fSetWhenCreated/*=FALSE*/, BOOL* pfDidCreate/*=NULL*/)
{
    HRESULT hr;
    LPCTSTR pchUniqueName = GetAAuniqueName();

    if(pfDidCreate)
    {
        *pfDidCreate = FALSE;
    }

    if(!pchUniqueName)
    {
        CDocument* pDoc = Doc();

        hr = pDoc->GetUniqueIdentifier(pcstr);
        if(hr)
        {
            goto Cleanup;
        }

        if(fSetWhenCreated)
        {
            if(pfDidCreate)
            {
                *pfDidCreate = TRUE;
            }
            hr = SetUniqueNameHelper(*pcstr);
        }
    }
    else
    {
        hr = pcstr->Set(pchUniqueName);
    }

Cleanup:
    RRETURN(hr);
}

LPCTSTR CElement::GetAAname() const
{
    LPCTSTR     pv;
    CAttrArray* pAA;

    if(!IsNamed())
    {
        return NULL;
    }

    pAA = *GetAttrArray();

    // We're leveraging the fact that we know the dispids of name & ID, & we
    // also have a single _pAA in CElement containing all the attributes
    if(pAA && pAA->FindString(DISPID_CElement_submitName/*STDPROPID_XOBJ_NAME wlw note*/, &pv))
    {
        return pv;
    }
    else
    {
        return NULL;
    }
}

LPCTSTR CElement::GetAAsubmitname() const
{
    LPCTSTR     pv;
    CAttrArray* pAA = *GetAttrArray();

    // We're leveraging the fact that we know the dispids of name & ID, & we
    // also have a single _pAA in CElement containing all the attributes
    if(pAA)
    {
        if(pAA->FindString(DISPID_CElement_submitName, &pv))
        {
            return pv;
        }
        else if(IsNamed() && pAA->FindString(STDPROPID_XOBJ_NAME, &pv))
        {
            return pv;
        }
    }
    return NULL;
}

HRESULT CElement::Inject(Where where, BOOL fIsHtml, LPTSTR pStr, long cch)
{
    HRESULT         hr = S_OK;
    CDocument*      pDoc = Doc();
    CMarkup*        pMarkup;
    BOOL            fEnsuredMarkup = FALSE;
    CMarkupPointer  pointerStart(pDoc);
    CMarkupPointer  pointerFinish(pDoc);
    /*CParentUndo     Undo(pDoc); wlw note*/
    ELEMENT_TAG     etag = Tag();
    IHTMLEditingServices* pedserv = NULL;

    // See if one is attempting to place stuff IN a noscope element
    if((where==Inside || where==AfterBegin || where==BeforeEnd) && IsNoScope())
    {
        // Some elements can do inside, but in the slave tree.
        // Also, disallow HTML for those things with a slave markup.  If you
        // can't get to them with the DOM or makrup services, you should not
        // be able to with innerHTML.
        if(!SlaveMarkup() || fIsHtml)
        {
            hr = CTL_E_INVALIDPASTETARGET;
            goto Cleanup;
        }

        if(IsContainer() && !TestClassFlag(CElement::ELEMENTDESC_OMREADONLY))
        {
            CElement* pElementTextSlave = SlaveMarkup()->GetElementClient();

            if(pElementTextSlave && pElementTextSlave->Tag() == ETAG_TXTSLAVE)
            {
                hr = pElementTextSlave->Inject(where, fIsHtml, pStr, cch);

                if(hr)
                {
                    goto Cleanup;
                }

                goto Cleanup;
            }
        }

        hr = CTL_E_INVALIDPASTETARGET;
        goto Cleanup;
    }

    // Disallow inner/outer on the head and html elements
    if((etag==ETAG_HTML || etag==ETAG_HEAD || etag==ETAG_TITLE_ELEMENT) && (where==Inside || where==Outside))
    {
        hr = CTL_E_INVALIDPASTETARGET;
        goto Cleanup;
    }

    // Prevent the elimination of the client element
    pMarkup = GetMarkup();

    if(pMarkup && (where==Inside || where==Outside))
    {
        CElement* pElementClient = pMarkup->GetElementClient();

        // It's ok to do an inner on the client
        if(pElementClient && (where!=Inside || this!=pElementClient))
        {
            // If we can see the client above this, then the client
            // will get blown away.  Prevent this.
            if(pMarkup->SearchBranchForScopeInStory(pElementClient->GetFirstBranch(), this))
            {
                hr = CTL_E_INVALIDPASTETARGET;
                goto Cleanup;
            }
        }
    }

    // In IE4, an element had to be in a markup to do this operation.  Now,
    // we are looser.  In order to do validation, the element must be in a
    // markup.  Here we also remember is we placed the element in a markup
    // so that if the injection fails, we can restore it to its "original"
    // state.
    if(!pMarkup)
    {
        hr = EnsureInMarkup();

        if(hr)
        {
            goto Cleanup;
        }

        fEnsuredMarkup = TRUE;

        pMarkup = GetMarkup();

        Assert(pMarkup);
    }

    // Locate the pointer such that they surround the stuff which should
    // go away, and are located where the new stuff should be placed.
    {
        ELEMENT_ADJACENCY adjLeft = ELEM_ADJ_BeforeEnd;
        ELEMENT_ADJACENCY adjRight = ELEM_ADJ_BeforeEnd;

        switch(where)
        {
        case Inside:
            adjLeft = ELEM_ADJ_AfterBegin;
            adjRight = ELEM_ADJ_BeforeEnd;
            break;

        case Outside:
            adjLeft = ELEM_ADJ_BeforeBegin;
            adjRight = ELEM_ADJ_AfterEnd;
            break;

        case BeforeBegin:
            adjLeft = ELEM_ADJ_BeforeBegin;
            adjRight = ELEM_ADJ_BeforeBegin;
            break;

        case AfterBegin:
            adjLeft = ELEM_ADJ_AfterBegin;
            adjRight = ELEM_ADJ_AfterBegin;
            break;

        case BeforeEnd:
            adjLeft = ELEM_ADJ_BeforeEnd;
            adjRight = ELEM_ADJ_BeforeEnd;
            break;

        case AfterEnd:
            adjLeft = ELEM_ADJ_AfterEnd;
            adjRight = ELEM_ADJ_AfterEnd;
            break;
        }

        hr = pointerStart.MoveAdjacentToElement(this, adjLeft);

        if(hr)
        {
            goto Cleanup;
        }

        hr = pointerFinish.MoveAdjacentToElement(this, adjRight);

        if(hr)
        {
            goto Cleanup;
        }
    }

    Assert(pointerStart.IsPositioned());
    Assert(pointerFinish.IsPositioned());

    {
        CTreeNode* pNodeStart  = pointerStart.Branch();
        CTreeNode* pNodeFinish = pointerFinish.Branch();

        // For the 5.0 version, because we don't have contextual parsing,
        // make sure tables can't we screwed with.
        if(fIsHtml)
        {
            // See if the beginning of the inject is in a table thingy
            if(pNodeStart && IsInTableThingy(pNodeStart))
            {
                hr = CTL_E_INVALIDPASTETARGET;
                goto Cleanup;
            }

            // See if the end of the inject is different from the start.
            // If so, then also check it for being in a table thingy.
            if(pNodeFinish && pNodeStart!=pNodeFinish && IsInTableThingy(pNodeStart))
            {
                hr = CTL_E_INVALIDPASTETARGET;
                goto Cleanup;
            }

            // Make sure we record undo information if we should.  I believe that
            // here is where we make the decision to not remembers automation
            // like manipulation, but do remember user editing scenarios.
            //
            // Here, we check the elements above the start and finish to make sure
            // they are editable (in the user sense).
            if(pNodeStart && pNodeFinish &&
                pNodeStart->Element()->IsEditable() &&
                pNodeFinish->Element()->IsEditable())
            {
                /*Undo.Start(IDS_UNDOGENERICTEXT); wlw note*/
            }
        }
    }

    // Perform the HTML/text injection
    if(fIsHtml)
    {
        AssertSz(FALSE, "must improve");
        goto Cleanup;
    }
    else
    {
        IHTMLEditor* phtmed;

        HRESULT RemoveWithBreakOnEmpty(CMarkupPointer* pPointerStart, CMarkupPointer* pPointerFinish);

        hr = RemoveWithBreakOnEmpty(&pointerStart, &pointerFinish);

        if(hr)
        {
            goto Cleanup;
        }

        if(where == Inside)
        {
            HRESULT UnoverlapPartials(CElement*);

            hr = UnoverlapPartials(this);

            if(hr)
            {
                goto Cleanup;
            }
        }

        // Get the editing services interface with which I can
        // insert sanitized text
        phtmed = Doc()->GetHTMLEditor();

        if(!phtmed)
        {
            hr = E_FAIL;
            goto Cleanup;
        }

        hr = phtmed->QueryInterface(IID_IHTMLEditingServices, (void**)&pedserv);

        if(hr)
        {
            goto Cleanup;
        }

        hr = pedserv->InsertSanitizedText(&pointerStart, pStr, TRUE);

        if(hr)
        {
            goto Cleanup;
        }

        // BUGBUG - Launder spaces here on the edges
    }

Cleanup:
    // If we are failing, and we had to put this element into a markup
    // at the beginning, take it out now to restore to the origianl state.
    if(hr!=S_OK && fEnsuredMarkup && GetMarkup())
    {
        Doc()->RemoveElement(this);
    }

    ReleaseInterface(pedserv);
    /*Undo.Finish(hr); wlw note*/

    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     InsertAdjacent
//
//  Synopsis:   Inserts the given element into the tree, positioned relative
//              to 'this' element as specified.
//
//-----------------------------------------------------------------------------
HRESULT CElement::InsertAdjacent(Where where, CElement* pElementInsert)
{
    HRESULT hr = S_OK;
    CMarkupPointer pointer(Doc());

    Assert(IsInMarkup());
    Assert(pElementInsert && !pElementInsert->IsInMarkup());
    Assert(!pElementInsert->IsRoot());
    Assert(!IsRoot() || where==AfterBegin || where==BeforeEnd);

    // Figure out where to put the element
    switch(where)
    {
    case BeforeBegin :
        hr = pointer.MoveAdjacentToElement(this, ELEM_ADJ_BeforeBegin);
        break;

    case AfterEnd :
        hr = pointer.MoveAdjacentToElement(this, ELEM_ADJ_AfterEnd);
        break;

    case AfterBegin :
        hr = pointer.MoveAdjacentToElement(this, ELEM_ADJ_AfterBegin);
        break;

    case BeforeEnd :
        hr = pointer.MoveAdjacentToElement(this, ELEM_ADJ_BeforeEnd);
        break;
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = Doc()->InsertElement(pElementInsert, &pointer, NULL);

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     RemoveOuter
//
//  Synopsis:   Removes 'this' element and everything which 'this' element
//              influences.
//
//-----------------------------------------------------------------------------
HRESULT CElement::RemoveOuter()
{
    HRESULT hr;
    CMarkupPointer p1(Doc()), p2(Doc());

    hr = p1.MoveAdjacentToElement(this, ELEM_ADJ_BeforeBegin);

    if(hr)
    {
        goto Cleanup;
    }

    hr = p2.MoveAdjacentToElement(this, ELEM_ADJ_AfterEnd);

    if(hr)
    {
        goto Cleanup;
    }

    hr = Doc()->Remove(&p1, &p2);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     GetText
//
//  Synopsis:   Gets the specified text for the element.
//
//  Note: invokes saver.  Use WBF_NO_TAG_FOR_CONTEXT to determine whether
//  or not the element itself is saved.
//
//-----------------------------------------------------------------------------
HRESULT CElement::GetText(BSTR* pbstr, DWORD dwStmFlags)
{
    HRESULT     hr = S_OK;
    IStream*    pstm = NULL;

    if(!pbstr)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pbstr = NULL;

    hr = CreateStreamOnHGlobal(NULL, TRUE, &pstm);
    if(hr)
    {
        goto Cleanup;
    }

    {
        CStreamWriteBuff swb(pstm, CP_UCS_2);

        swb.SetFlags(dwStmFlags);
        swb.SetElementContext(this);

        // Save the begin tag of the context element
        hr = Save(&swb, FALSE);
        if(hr)
        {
            goto Cleanup;
        }

        if(IsInMarkup())
        {
            CTreeSaver ts(this, &swb);
            hr = ts.Save();
            if(hr)
            {
                goto Cleanup;
            }
        }

        // Save the end tag of the context element
        Save(&swb, TRUE);
        if(hr)
        {
            goto Cleanup;
        }

        hr = swb.Terminate();
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = GetBStrFromStream(pstm, pbstr, TRUE);

Cleanup:
    ReleaseInterface(pstm);
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Function:   GetBstrFromElement
//
//  Synopsis:   A helper for data binding, fetches the text of some Element.
//
//  Arguments:  [fHTML]     - does caller want HTML or plain text?
//              [pbstr]     - where to return the BSTR holding the contents
//
//  Returns:    S_OK if successful
//
//-----------------------------------------------------------------------------
HRESULT CElement::GetBstrFromElement(BOOL fHTML, BSTR* pbstr)
{
    HRESULT hr;

    *pbstr = NULL;

    if(fHTML)
    {
        // Go through the HTML saver
        hr = GetText(pbstr, 0);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else
    {
        //  Grab the plaintext directly from the runs
        CString cstr;

        hr = GetPlainTextInScope(&cstr);
        if(hr)
        {
            goto Cleanup;
        }

        hr = cstr.AllocBSTR(pbstr);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

// CElement Collection Helpers - manage the WINDOW_COLLECTION
HRESULT CElement::GetIDHelper(CString* pf)
{
    LPCTSTR lpszID = NULL;
    HRESULT hr;

    if(_pAA && _pAA->HasAnyAttribute())
    {
        _pAA->FindString(DISPID_CElement_id, &lpszID);
    }

    hr = pf->Set(lpszID);

    return hr;
}

HRESULT CElement::SetIDHelper(CString* pf)
{
    RRETURN(SetIdentifierHelper((LPCTSTR)(*pf), DISPID_CElement_id,
        STDPROPID_XOBJ_NAME, DISPID_CElement_uniqueName));
}

HRESULT CElement::GetnameHelper(CString* pf)
{
    LPCTSTR lpszID = NULL;
    HRESULT hr;

    if(_pAA && _pAA->HasAnyAttribute())
    {
        _pAA->FindString(STDPROPID_XOBJ_NAME, &lpszID);
    }

    hr = pf->Set(lpszID);

    return hr;
}

HRESULT CElement::SetnameHelper(CString* pf)
{
    RRETURN(SetIdentifierHelper((LPCTSTR)(*pf), STDPROPID_XOBJ_NAME,
        DISPID_CElement_id, DISPID_CElement_uniqueName));
}

HRESULT CElement::SetIdentifierHelper(LPCTSTR lpszValue, DISPID dspIDThis, DISPID dspOther1, DISPID dspOther2)
{
    HRESULT     hr;
    BOOL        fNamed;
    CDocument*  pDoc = Doc();

    hr = AddString(dspIDThis, lpszValue, CAttrValue::AA_Attribute);
    if(!hr)
    {
        // Remember if we're named, so that if this element moves into a different tree
        // we can inval the appropriate collections
        fNamed = lpszValue && lpszValue[0];

        // We're named if NAME= or ID= something or we have a unique name
        if(!fNamed)
        {
            LPCTSTR lpsz = NULL;
            if(_pAA && _pAA->HasAnyAttribute())
            {
                _pAA->FindString(dspOther1, &lpsz);
                if(!(lpsz && *lpsz))
                {
                    _pAA->FindString(dspOther2, &lpsz);
                }
            }
            fNamed = lpsz && *lpsz;
        }

        _fIsNamed = fNamed;
        // Inval all collections affected by a name change
        DoElementNameChangeCollections();
    }
    RRETURN(hr);
}

void CElement::OnEnterExitInvalidateCollections(BOOL fForceNamedBuild)
{
    // Optimized collections
    // If a named (name= or ID=) element enters the tree, inval the collections

    // DEVNOTE rgardner
    // This code is tighly couples with CMarkup::AddToCollections and needs
    // to be kept in sync with any changes in that function
    if(IsNamed() || fForceNamedBuild)
    {
        InvalidateCollection(CMarkup::WINDOW_COLLECTION);
    }

    // Inval collections based on specific element types
    switch(_etag)
    {
    case ETAG_LABEL:
        InvalidateCollection(CMarkup::LABEL_COLLECTION);
        break;

    case ETAG_IMG:
        InvalidateCollection(CMarkup::IMAGES_COLLECTION);
        if(IsNamed() || fForceNamedBuild)
        {
            InvalidateCollection(CMarkup::NAVDOCUMENT_COLLECTION);
        }
        break;

    case ETAG_OBJECT:
    case ETAG_APPLET:
        InvalidateCollection(CMarkup::APPLETS_COLLECTION);
        if(IsNamed() || fForceNamedBuild)
        {
            InvalidateCollection(CMarkup::NAVDOCUMENT_COLLECTION);
        }
        break;

    case ETAG_SCRIPT:
        InvalidateCollection(CMarkup::SCRIPTS_COLLECTION);
        break;

    case ETAG_AREA:
        InvalidateCollection(CMarkup::LINKS_COLLECTION);
        break;

    case ETAG_FORM:
        InvalidateCollection(CMarkup::FORMS_COLLECTION);
        if(IsNamed() || fForceNamedBuild)
        {
            InvalidateCollection(CMarkup::NAVDOCUMENT_COLLECTION);
        }
        break;

    case ETAG_A:
        InvalidateCollection(CMarkup::LINKS_COLLECTION);
        if(IsNamed() || fForceNamedBuild)
        {
            InvalidateCollection(CMarkup::ANCHORS_COLLECTION);
        }
        break;
    }
}

void CElement::DoElementNameChangeCollections(void)
{
    // Inval all the base collections, based on TAGNAme.
    // Force the WINDOW_COLLECTION to be built
    OnEnterExitInvalidateCollections(TRUE);

    // Artificialy update the _lTreeVersion so any named collection derived from 
    // the ELEMENT_COLLECTION are inval'ed

    // BUGBUG PERF Must not modify this unless document is changing!!!!
    // What should really be done here is have another version number on the
    // doc which this bumps, and have the collections record two version
    // numbers, one for the tree, and another they can bump here.  THen to
    // see if you are out of date, simply comapre the two pairs of version
    // numbers.
    Doc()->UpdateDocTreeVersion();
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::Clone
//
//  Synopsis:   Make a new one just like this one
//
//-------------------------------------------------------------------------
HRESULT CElement::Clone(CElement** ppElementClone, CDocument* pDoc)
{
    HRESULT         hr;
    CAttrArray*     pAA;
    CString         cstrPch;
    CElement*       pElementNew = NULL;
    const CTagDesc* ptd;
    CHtmTag         ht;
    BOOL            fDie = FALSE;
    ELEMENT_TAG     etag = Tag();

    Assert(ppElementClone);

    if(IsGenericTag(etag))
    {
        hr = cstrPch.Append(_T(":"));

        if(hr)
        {
            goto Cleanup;
        }

        hr = cstrPch.Append( TagName());

        if(hr)
        {
            goto Cleanup;
        }
    }
    else if(etag == ETAG_UNKNOWN)
    {
        hr = cstrPch.Append(TagName());

        if(hr)
        {
            goto Cleanup;
        }
    }

    ptd = TagDescFromEtag(etag);

    if(!ptd)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    ht.Reset();
    ht.SetTag(etag);
    ht.SetPch(cstrPch);
    ht.SetCch(cstrPch.Length());

    hr = ptd->_pfnElementCreator(&ht, pDoc, &pElementNew);

    if(hr)
    {
        goto Cleanup;
    }

    if(fDie)
    {
        goto Die;
    }

    hr = pElementNew->Init();

    if(hr)
    {
        goto Cleanup;
    }

    if(fDie)
    {
        goto Die;
    }

    pElementNew->_fBreakOnEmpty = _fBreakOnEmpty;
    pElementNew->_fExplicitEndTag = _fExplicitEndTag;

    pAA = *GetAttrArray();

    if(pAA)
    {
        CAttrArray** ppAAClone = pElementNew->GetAttrArray();

        hr = pAA->Clone(ppAAClone);

        if(hr)
        {
            goto Cleanup;
        }

        if(fDie)
        {
            goto Die;
        }
    }

    pElementNew->_fIsNamed = _fIsNamed;    

    {
        CInit2Context context(&ht, GetMarkup(), INIT2FLAG_EXECUTE);

        hr = pElementNew->Init2(&context);
    }
    if(hr)
    {
        goto Cleanup;
    }

    if(fDie)
    {
        goto Die;
    }

    pElementNew->SetEventsShouldFire();

Cleanup:

    if(hr && pElementNew)
    {
        CElement::ClearPtr(&pElementNew);
    }

    *ppElementClone = pElementNew;

    RRETURN(hr);

Die:
    hr = E_ABORT;
    goto Cleanup;
}

//+----------------------------------------------------------------------------
//
// Member:      CElement::ComputeHorzBorderAndPadding
//
// Synopsis:    Compute horizontal border and padding for a given element
//              The results represent cumulative border and padding up the
//              element's ancestor chain, up to but NOT INCLUDING element's
//              containing layout.  The layout's border should not be counted
//              when determining a contained element's indent, because it lies
//              outside the box boundary from which we are measuring.  The 
//              layout's padding usually does need to be accounted for; the caller
//              must do this! (via GetPadding on the layout's CDisplay).
//
//-----------------------------------------------------------------------------
void CElement::ComputeHorzBorderAndPadding(
        CCalcInfo*  pci,
        CTreeNode*  pNodeContext,
        CElement*   pElementStop,
        LONG*       pxBorderLeft,
        LONG*       pxPaddingLeft,
        LONG*       pxBorderRight,
        LONG*       pxPaddingRight)
{
    Assert(pNodeContext && SameScope(this, pNodeContext));

    Assert(pxBorderLeft || pxPaddingLeft || pxBorderRight || pxPaddingRight);

    CTreeNode*          pNode = pNodeContext;
    CBorderInfo         borderinfo;
    CElement*           pElement;
    const CFancyFormat* pFF;
    const CParaFormat*  pPF;

    Assert(pxBorderLeft && pxPaddingLeft && pxBorderRight && pxPaddingRight);

    *pxBorderLeft = 0;
    *pxBorderRight = 0;
    *pxPaddingLeft = 0;;
    *pxPaddingRight = 0;

    while(pNode && pNode->Element()!=pElementStop)
    {
        pElement = pNode->Element();
        pFF = pNode->GetFancyFormat();
        pPF = pNode->GetParaFormat();

        if(!pElement->_fDefinitelyNoBorders)
        {
            pElement->_fDefinitelyNoBorders = !GetBorderInfoHelper(pNode, pci, &borderinfo, FALSE);

            *pxBorderRight += borderinfo.aiWidths[BORDER_RIGHT];
            *pxBorderLeft += borderinfo.aiWidths[BORDER_LEFT];
        }

        *pxPaddingLeft += pFF->_cuvPaddingLeft.XGetPixelValue(pci, pci->_sizeParent.cx, pPF->_lFontHeightTwips);
        *pxPaddingRight += pFF->_cuvPaddingRight.XGetPixelValue(pci, pci->_sizeParent.cx, pPF->_lFontHeightTwips);

        pNode = pNode->Parent();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::SetDim, public
//
//  Synopsis:   Sets a given property (either on the inline style or the
//              attribute directly) to a given pixel value, preserving the
//              original units of that attribute.
//
//  Arguments:  [dispID]       -- Property to set the value of
//              [fValue]       -- Value of the property
//              [uvt]          -- Units [fValue] is in. If UNIT_NULLVALUE then
//                                 [fValue] is assumed to be in whatever the
//                                 current units are for this property.
//              [lDimOf]       -- For percentage values, what the percent is of
//              [fInlineStyle] -- If TRUE, the inline style is changed,
//                                otherwise the HTML attribute is changed
//              [pfChanged]    -- Place to indicate if the value actually
//                                changed
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CElement::SetDim(
         DISPID                    dispID,
         float                     fValue,
         CUnitValue::UNITVALUETYPE uvt,
         long                      lDimOf,
         CAttrArray**              ppAttrArray,
         BOOL                      fInlineStyle,
         BOOL*                     pfChanged)
{
    CUnitValue  uvValue;
    HRESULT     hr;
    long        lRawValue;

    Assert(pfChanged);
    uvValue.SetNull();

    if(!ppAttrArray)
    {
        if(fInlineStyle)
        {
            ppAttrArray = CreateStyleAttrArray(DISPID_INTERNAL_INLINESTYLEAA);
        }
        else
        {
            ppAttrArray = GetAttrArray();
        }
    }

    Assert(ppAttrArray);

    if(*ppAttrArray)
    {
        (*ppAttrArray)->GetSimpleAt((*ppAttrArray)->FindAAIndex(dispID, CAttrValue::AA_Attribute), (DWORD*)&uvValue);
    }

    lRawValue = uvValue.GetRawValue();

    if(uvt == CUnitValue::UNIT_NULLVALUE)
    {
        uvt = uvValue.GetUnitType();

        if(uvt == CUnitValue::UNIT_NULLVALUE)
        {
            uvt = CUnitValue::UNIT_PIXELS;
        }
    }

    if(dispID==STDPROPID_XOBJ_HEIGHT || dispID==STDPROPID_XOBJ_TOP)
    {
        hr = uvValue.YSetFloatValueKeepUnits(fValue,
            uvt,
            lDimOf,
            GetFirstBranch()->GetFontHeightInTwips(&uvValue));
    }
    else
    {
        hr = uvValue.XSetFloatValueKeepUnits(fValue,
            uvt,
            lDimOf,
            GetFirstBranch()->GetFontHeightInTwips(&uvValue));
    }
    if(hr)
    {
        goto Cleanup;
    }

    if(uvValue.GetRawValue() == lRawValue) // Has anything changed ??
    {
        goto Cleanup;
    }

    hr = CAttrArray::AddSimple(ppAttrArray, dispID, uvValue.GetRawValue(), CAttrValue::AA_StyleAttribute);

    if(hr)
    {
        goto Cleanup;
    }

    *pfChanged = TRUE;

Cleanup:
    RRETURN(hr);
}

HRESULT CElement::ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget)
{
    CDocument*              pDoc = Doc();
    THREADSTATE*            pts  = GetThreadState();
    CTreeNode*              pNodeParent;
    CElement*               pElemParent;
    BOOL                    fResetPosition = FALSE;
    BOOL                    fComputeFFOnly = pNodeTarget->_iCF != -1;
    HRESULT                 hr = S_OK;
    COMPUTEFORMATSTYPE      eExtraValues = pCFI->_eExtraValues;

    Assert(pCFI);
    Assert(SameScope(this, pNodeTarget));
    Assert(eExtraValues!=ComputeFormatsType_Normal || ((pNodeTarget->_iCF==-1 && pNodeTarget->_iPF==-1) || pNodeTarget->_iFF==-1));

    // Get the format of our parent before applying our own format.
    pNodeParent = pNodeTarget->Parent();
    switch(_etag)
    {
    case ETAG_TR:
        {
            // wlw note
            //CTableSection* pSection = DYNCAST(CTableRow, this)->Section();
            //if(pSection)
            //{
            //    pNodeParent = pSection->GetFirstBranch();
            //}
        }
        break;
    case ETAG_TBODY:
    case ETAG_THEAD:
    case ETAG_TFOOT:
        {
            // wlw note
            //CTable* pTable = DYNCAST(CTableSection, this)->Table();
            //if(pTable)
            //{
            //    pNodeParent = pTable->GetFirstBranch();
            //}
            //fResetPosition = TRUE;
        }
        break;
    }

    if(pNodeParent == NULL)
    {
        AssertSz(0, "CElement::ComputeFormats should not be called on elements without a parent node");
        hr = E_FAIL;
        goto Cleanup;
    }

    // If the parent node has not computed formats yet, recursively compute them
    pElemParent = pNodeParent->Element();

    if(pNodeParent->_iCF==-1 || pNodeParent->_iFF==-1 || eExtraValues==ComputeFormatsType_GetInheritedValue)
    {
        hr = pElemParent->ComputeFormats(pCFI, pNodeParent);

        if(hr)
        {
            goto Cleanup;
        }
    }

    Assert(pNodeParent->_iCF >= 0);
    Assert(pNodeParent->_iPF >= 0);
    Assert(pNodeParent->_iFF >= 0);

    // NOTE: From this point forward any errors must goto Error instead of Cleanup!
    pCFI->Reset();
    pCFI->_pNodeContext = pNodeTarget;

    // Setup Fancy Format
    if(_fInheritFF)
    {
        pCFI->_iffSrc = pNodeParent->_iFF;
        pCFI->_pffSrc = pCFI->_pff = &(*pts->_pFancyFormatCache)[pCFI->_iffSrc];
        pCFI->_fHasExpandos = FALSE;

        if( pCFI->_pff->_bPositionType!=stylePositionNotSet
            || pCFI->_pff->_bDisplay!=styleDisplayNotSet
            || pCFI->_pff->_bVisibility!=styleVisibilityNotSet
            || pCFI->_pff->_bOverflowX!=styleOverflowNotSet
            || pCFI->_pff->_bOverflowY!=styleOverflowNotSet
            || pCFI->_pff->_fPositioned
            || pCFI->_pff->_fAutoPositioned
            || pCFI->_pff->_fScrollingParent
            || pCFI->_pff->_fZParent
            || pCFI->_pff->_ccvBackColor.IsDefined()
            || pCFI->_pff->_lImgCtxCookie!=0
            || pCFI->_pff->_iExpandos!=-1
            || pCFI->_pff->_fHasExpressions!=0
            || pCFI->_pff->_pszFilters
            || pCFI->_pff->_fHasNoWrap)
        {
            pCFI->PrepareFancyFormat();
            pCFI->_ff()._bPositionType = stylePositionNotSet;
            pCFI->_ff()._bDisplay   = styleDisplayNotSet;
            pCFI->_ff()._bVisibility = styleVisibilityNotSet;
            pCFI->_ff()._bOverflowX = styleOverflowNotSet;
            pCFI->_ff()._bOverflowY = styleOverflowNotSet;
            pCFI->_ff()._pszFilters = NULL;
            pCFI->_ff()._fPositioned = FALSE;
            pCFI->_ff()._fAutoPositioned = FALSE;
            pCFI->_ff()._fScrollingParent = FALSE;
            pCFI->_ff()._fZParent = FALSE;
            pCFI->_ff()._fHasNoWrap = FALSE;

            // We never ever inherit expandos or expressions
            pCFI->_ff()._iExpandos = -1;
            pCFI->_ff()._fHasExpressions = FALSE;

            if(Tag() != ETAG_TR)
            {
                // do not inherit background from the table.
                pCFI->_ff()._ccvBackColor.Undefine();
                pCFI->_ff()._lImgCtxCookie = 0;
                pCFI->UnprepareForDebug();
            }
        }
    }
    else
    {
        pCFI->_iffSrc = pts->_iffDefault;
        pCFI->_pffSrc = pCFI->_pff = pts->_pffDefault;

        Assert(pCFI->_pffSrc->_pszFilters == NULL);
    }

    if(!fComputeFFOnly)
    {
        // Setup Char and Para formats
        if(TestClassFlag(ELEMENTDESC_DONTINHERITSTYLE))
        {
            // The CharFormat inherits a couple of attributes from the parent, the rest from defaults.
            const CCharFormat* pcfParent = &(*pts->_pCharFormatCache)[pNodeParent->_iCF];
            const CParaFormat* ppfParent = &(*pts->_pParaFormatCache)[pNodeParent->_iPF];

            pCFI->_fDisplayNone      = pcfParent->_fDisplayNone;
            pCFI->_fVisibilityHidden = pcfParent->_fVisibilityHidden;

            if(pDoc->_icfDefault < 0)
            {
                hr = pDoc->CacheDefaultCharFormat();
                if (hr)
                    goto Error;
            }

            Assert(pDoc->_icfDefault >= 0);
            Assert(pDoc->_pcfDefault != NULL);

            pCFI->_icfSrc = pDoc->_icfDefault;
            pCFI->_pcfSrc = pCFI->_pcf = pDoc->_pcfDefault;

            // Some properties are ALWAYS inherited, regardless of ELEMENTDESC_DONTINHERITSTYLE.
            // Do that here:
            if(pCFI->_fDisplayNone
                || pCFI->_fVisibilityHidden
                || pcfParent->_fHasBgColor
                || pcfParent->_fHasBgImage
                || pcfParent->_fRelative
                || pcfParent->_fNoBreakInner
                || pcfParent->_fRTL
                || pcfParent->_fBranchFiltered
                || pCFI->_pcf->_bCursorIdx!=pcfParent->_bCursorIdx
                || pcfParent->_fDisabled)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fDisplayNone       = pCFI->_fDisplayNone;
                pCFI->_cf()._fVisibilityHidden  = pCFI->_fVisibilityHidden;
                pCFI->_cf()._fHasBgColor        = pcfParent->_fHasBgColor;
                pCFI->_cf()._fHasBgImage        = pcfParent->_fHasBgImage;
                pCFI->_cf()._fRelative          = pNodeParent->Element()->NeedsLayout() ? FALSE : pcfParent->_fRelative;
                pCFI->_cf()._fNoBreak           = pcfParent->_fNoBreakInner;
                pCFI->_cf()._fRTL               = pcfParent->_fRTL;
                pCFI->_cf()._fBranchFiltered    = pcfParent->_fBranchFiltered;
                pCFI->_cf()._bCursorIdx         = pcfParent->_bCursorIdx;
                pCFI->_cf()._fDisabled          = pcfParent->_fDisabled;
                pCFI->UnprepareForDebug();
            }


            pCFI->_ipfSrc = pNodeParent->_iPF;
            pCFI->_ppfSrc = pCFI->_ppf = ppfParent;

            if(pCFI->_ppf->_fPreInner || pCFI->_ppf->_fInclEOLWhiteInner || pCFI->_ppf->_bBlockAlignInner!=htmlBlockAlignNotSet)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fPreInner              = FALSE;
                // BUGBUG: (KTam) shouldn't this be _fInclEOLWhiteInner? fix in IE6
                pCFI->_pf()._fInclEOLWhite          = FALSE;
                pCFI->_pf()._bBlockAlignInner       = htmlBlockAlignNotSet;
                pCFI->UnprepareForDebug();
            }

            if(pcfParent->_fHasGridValues)
            {
                pCFI->PrepareCharFormat();
                pCFI->PrepareParaFormat();
                pCFI->_cf()._fHasGridValues = TRUE;
                pCFI->_cf()._uLayoutGridMode = pcfParent->_uLayoutGridModeInner;
                pCFI->_cf()._uLayoutGridType = pcfParent->_uLayoutGridTypeInner;
                pCFI->_pf()._cuvCharGridSize = ppfParent->_cuvCharGridSizeInner;
                pCFI->_pf()._cuvLineGridSize = ppfParent->_cuvLineGridSizeInner;

                pCFI->UnprepareForDebug();
            }

            // outer block alignment should still be inherited from
            // parent, but reset the inner block alignment
            pCFI->_bCtrlBlockAlign  = ppfParent->_bBlockAlignInner;
            pCFI->_bBlockAlign      = htmlBlockAlignNotSet;


            // outer direction should still be inherited from parent
            if(pCFI->_ppf->_fRTL != pCFI->_ppf->_fRTLInner)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fRTL = pCFI->_ppf->_fRTLInner;
                pCFI->UnprepareForDebug();
            }
        }
        else
        {
            // Inherit the Char and Para formats from the parent node

            pCFI->_icfSrc = pNodeParent->_iCF;
            pCFI->_pcfSrc = pCFI->_pcf = &(*pts->_pCharFormatCache)[pCFI->_icfSrc];
            pCFI->_ipfSrc = pNodeParent->_iPF;
            pCFI->_ppfSrc = pCFI->_ppf = &(*pts->_pParaFormatCache)[pCFI->_ipfSrc];

            pCFI->_fDisplayNone      = pCFI->_pcf->_fDisplayNone;
            pCFI->_fVisibilityHidden = pCFI->_pcf->_fVisibilityHidden;

            // If the parent had layoutness, clear the inner formats
            if(pNodeParent->Element()->NeedsLayout())
            {
                if(pCFI->_pcf->_fHasDirtyInnerFormats)
                {
                    pCFI->PrepareCharFormat();
                    pCFI->_cf().ClearInnerFormats();
                    pCFI->UnprepareForDebug();
                }

                if(pCFI->_ppf->_fHasDirtyInnerFormats)
                {
                    pCFI->PrepareParaFormat();
                    pCFI->_pf().ClearInnerFormats();
                    pCFI->UnprepareForDebug();
                }

                // copy parent's inner formats to current elements outer
                if(pCFI->_ppf->_fPre!=pCFI->_ppf->_fPreInner
                    || pCFI->_ppf->_fInclEOLWhite!=pCFI->_ppf->_fInclEOLWhiteInner
                    || pCFI->_ppf->_bBlockAlign!=pCFI->_ppf->_bBlockAlignInner)
                {
                    pCFI->PrepareParaFormat();
                    pCFI->_pf()._fPre = pCFI->_pf()._fPreInner;
                    pCFI->_pf()._fInclEOLWhite = pCFI->_pf()._fInclEOLWhiteInner;
                    pCFI->_pf()._bBlockAlign = pCFI->_pf()._bBlockAlignInner;
                    pCFI->UnprepareForDebug();
                }

                if(pCFI->_pcf->_fNoBreak != pCFI->_pcf->_fNoBreakInner)
                {
                    pCFI->PrepareCharFormat();
                    pCFI->_cf()._fNoBreak = pCFI->_pcf->_fNoBreakInner;
                    pCFI->UnprepareForDebug();
                }

                // copy parent's inner formats to current elements outer
                if(pCFI->_pcf->_fHasGridValues)
                {
                    pCFI->PrepareCharFormat();
                    pCFI->PrepareParaFormat();
                    pCFI->_cf()._uLayoutGridMode = pCFI->_pcf->_uLayoutGridModeInner;
                    pCFI->_cf()._uLayoutGridType = pCFI->_pcf->_uLayoutGridTypeInner;
                    pCFI->_pf()._cuvCharGridSize = pCFI->_ppf->_cuvCharGridSizeInner;
                    pCFI->_pf()._cuvLineGridSize = pCFI->_ppf->_cuvLineGridSizeInner;
                    pCFI->UnprepareForDebug();
                }
            }

            if(pCFI->_pcf->_fBranchFiltered)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fBranchFiltered = TRUE;
                pCFI->UnprepareForDebug();
            }

            // outer block alignment should still be inherited from
            // parent
            pCFI->_bCtrlBlockAlign  = pCFI->_ppf->_bBlockAlign;
            pCFI->_bBlockAlign      = pCFI->_ppf->_bBlockAlign;

            // outer direction should still be inherited from parent
            if(pCFI->_ppf->_fRTL != pCFI->_ppf->_fRTLInner)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fRTL = pCFI->_ppf->_fRTLInner;
                pCFI->UnprepareForDebug();
            }

        }

        pCFI->_bControlAlign = htmlControlAlignNotSet;
    }
    else
    {
        pCFI->_icfSrc = pDoc->_icfDefault;
        pCFI->_pcfSrc = pCFI->_pcf = pDoc->_pcfDefault;
        pCFI->_ipfSrc = pts->_ipfDefault;
        pCFI->_ppfSrc = pCFI->_ppf = pts->_ppfDefault;
    }

    hr = ApplyDefaultFormat(pCFI);
    if(hr)
    {
        goto Error;
    }

    hr = ApplyInnerOuterFormats(pCFI);
    if(hr)
    {
        goto Error;
    }

    if(eExtraValues==ComputeFormatsType_Normal || eExtraValues==ComputeFormatsType_ForceDefaultValue)
    {
        if(fResetPosition)
        {
            if(pCFI->_pff->_bPositionType != stylePositionstatic)
            {
                pCFI->PrepareFancyFormat();
                pCFI->_ff()._bPositionType = stylePositionstatic;
                pCFI->UnprepareForDebug();
            }
        }

        if(GetMarkup()->IsPrimaryMarkup() && (_fHasFilterCollectionPtr||pCFI->_fHasFilters))
        {
            AssertSz(FALSE, "must improve");
        }

        hr = pNodeTarget->CacheNewFormats(pCFI);
        pCFI->_cstrFilters.Free();  // Arrggh!!! BUGBUG (michaelw)  This should really happen 
                                    // somewhere else (when you know where, put it there)
                                    // Fix CTableCell::ComputeFormats also
        if(hr)
        {
            goto Error;
        }

        // Cache whether an element is a block element or not for fast retrieval.
        pNodeTarget->_fBlockNess = pCFI->_pff->_fBlockNess;

        // BUGBUG (srinib) - we need to move this bit onto element.
        if(HasLayoutPtr())
        {
            GetCurLayout()->_fEditableDirty = TRUE;

            // If the current element has a layout attached to but does not
            // need one, post a layout request to lazily destroy it.
            if(!pCFI->_pff->_fHasLayout)
            {
                GetCurLayout()->PostLayoutRequest(LAYOUT_MEASURE|LAYOUT_POSITION);
            }
        }

        pNodeTarget->_fHasLayout = pCFI->_pff->_fHasLayout;

        // Update expressions in the recalc engine
        //
        // If we had expressions or have expressions then we need to tell the recalc engine
        if(_fHasStyleExpressions || pCFI->_pff->_fHasExpressions)
        {
            AssertSz(FALSE, "must improve");
        }
    }

Cleanup:
    RRETURN(hr);

Error:
    pCFI->Cleanup();
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::ApplyDefaultFormat
//
//  Synopsis:   Applies default formatting properties for that element to
//              the char and para formats passed in
//
//  Arguments:  pCFI - Format Info needed for cascading
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CElement::ApplyDefaultFormat(CFormatInfo* pCFI)
{
    HRESULT             hr = S_OK;
    CDocument*          pDoc = Doc();
    OPTIONSETTINGS*     pos  = pDoc->_pOptionSettings;
    BOOL                fUseStyleSheets = !pos || pos->fUseStylesheets;

    if(_pAA
        || (fUseStyleSheets && GetMarkup()->HasStyleSheets())
        || (pos && pos->fUseMyStylesheet) 
        /*|| (HasPeerHolder() wlw note)*/)
    {
        CAttrArray* pInLineStyleAA;
        BOOL        fUserImportant = FALSE;
        BOOL        fDocumentImportant = FALSE;
        BOOL        fInlineImportant = FALSE;
        BOOL        fRuntimeImportant = FALSE;

        Assert(pCFI && pCFI->_pNodeContext && SameScope(this, pCFI->_pNodeContext));

        // Apply any HTML formatting properties
        if(_pAA)
        {
            hr = ApplyAttrArrayValues(pCFI, &_pAA, APPLY_All, NULL, FALSE);
            if(hr)
            {
                goto Cleanup;
            }
        }

        // Skip author stylesheet properties if they're turned off.
        if(fUseStyleSheets)
        {
            Trace0("Applying author style sheets\n");

            hr = GetMarkup()->ApplyStyleSheets(pCFI, APPLY_NoImportant, &fDocumentImportant);
            if(hr)
            {
                goto Cleanup;
            }

            if(_pAA)
            {
                // Apply any inline STYLE rules
                pInLineStyleAA = GetInLineStyleAttrArray();

                if(pInLineStyleAA)
                {
                    Trace0("Applying inline style attr array\n");

                    // The last parameter to ApplyAttrArrayValues is used to prevent the expression of that dispid from being
                    // overwritten.
                    //
                    // BUGBUG (michaelw) this hackyness should go away when we store both the expression and the value in a single CAttrValue
                    hr = ApplyAttrArrayValues(pCFI, &pInLineStyleAA, APPLY_NoImportant, &fInlineImportant, TRUE, 0/*Doc()->_recalcHost.GetSetValueDispid(this) wlw note*/);
                    if(hr)
                    {
                        goto Cleanup;
                    }
                }

                if(GetRuntimeStylePtr())
                {
                    CAttrArray* pRuntimeStyleAA;

                    pRuntimeStyleAA = *GetRuntimeStylePtr()->GetAttrArray();
                    if(pRuntimeStyleAA)
                    {
                        Trace0("Applying runtime style attr array\n");

                        // The last parameter to ApplyAttrArrayValues is used to 
                        // prevent the expression of that dispid from being
                        // overwritten.
                        //
                        // BUGBUG (michaelw) this hackyness should go away when 
                        // we store both the expression and the value in a single CAttrValue
                        hr = ApplyAttrArrayValues(pCFI, &pRuntimeStyleAA,
                            APPLY_NoImportant, &fRuntimeImportant, TRUE, 
                            0/*Doc()->_recalcHost.GetSetValueDispid(this) wlw note*/);
                        if(hr)
                        {
                            goto Cleanup;
                        }
                    }
                }
            }
        }

        // Now handle any "!important" properties.
        // Order: document !important, inline, runtime, user !important.

        // Apply any document !important rules
        if(fDocumentImportant)
        {
            Trace0("Applying important doc styles\n");

            AssertSz(FALSE, "must improve");
            goto Cleanup;
        }

        // Apply any inline STYLE rules
        if(fInlineImportant)
        {
            Trace0("Applying important inline styles\n");

            hr = ApplyAttrArrayValues(pCFI, &pInLineStyleAA, APPLY_ImportantOnly);
            if(hr)
            {
                goto Cleanup;
            }
        }

        // Apply any runtime important STYLE rules
        if(fRuntimeImportant)
        {
            Trace0("Applying important runtimestyles\n");

            hr = ApplyAttrArrayValues(
                pCFI, GetRuntimeStylePtr()->GetAttrArray(), APPLY_ImportantOnly);
            if(hr)
            {
                goto Cleanup;
            }
        }

        // Apply user !important rules last for accessibility
        if(fUserImportant)
        {
            Trace0("Applying important user styles\n");

            AssertSz(FALSE, "must improve");
            goto Cleanup;
        }
    }

    if(ComputeFormatsType_ForceDefaultValue == pCFI->_eExtraValues)
    {
        Assert(pCFI->_pStyleForce);
        hr = ApplyAttrArrayValues(
            pCFI, pCFI->_pStyleForce->GetAttrArray(), APPLY_All);
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::NeedsLayout, public
//
//  Synopsis:   Determines, based on attributes, whether this element needs
//              layout and if so what type.
//
//  Arguments:  [pCFI]  -- Pointer to FormatInfo with applied properties
//              [pNode] -- Context Node for this element
//              [plt]   -- Returns the type of layout we need.
//
//  Returns:    TRUE if this element needs a layout, FALSE if not.
//
//----------------------------------------------------------------------------
BOOL CElement::ElementNeedsLayout(CFormatInfo* pCFI)
{
    if(_fSite || (!TestClassFlag(ELEMENTDESC_NOLAYOUT) && pCFI->_pff->ElementNeedsFlowLayout()))
    {
        return TRUE;
    }

    return FALSE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::DetermineBlockness, public
//
//  Synopsis:   Determines if this element is a block element based on
//              inherent characteristics of the tag or current properties.
//
//  Arguments:  [pCFI] -- Pointer to current FormatInfo with applied properties
//
//  Returns:    TRUE if this element has blockness.
//
//----------------------------------------------------------------------------
BOOL CElement::DetermineBlockness(CFormatInfo* pCFI)
{
    BOOL fHasBlockFlag = HasFlag(TAGDESC_BLOCKELEMENT); 
    BOOL fIsBlock      = fHasBlockFlag;
    styleDisplay disp  = (styleDisplay)(pCFI->_pff->_bDisplay);

    // BUGBUG -- Are there any elements for which we don't want to override
    // blockness? (lylec)
    if(disp == styleDisplayBlock)
    {
        fIsBlock = TRUE;
    }
    else if(disp == styleDisplayInline)
    {
        fIsBlock = FALSE;
    }

    return fIsBlock;
}

BOOL IsBlockListElement(CTreeNode* pNode)
{
    return pNode->Element()->IsFlagAndBlock(TAGDESC_LIST);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::ApplyInnerOuterFormats, public
//
//  Synopsis:   Takes the current FormatInfo and creates the correct
//              inner and (if appropriate) outer formats.
//
//  Arguments:  [pCFI]     -- FormatInfo with applied properties
//              [pCFOuter] -- Place to store Outer format properties
//              [pPFOuter] -- Place to store Outer format properties
//
//  Returns:    HRESULT
//
//  Notes:      Inner/Outer sensitive formats are put in the _fXXXXInner
//              for inner and outer are held in _fXXXXXX
//
//----------------------------------------------------------------------------
HRESULT CElement::ApplyInnerOuterFormats(CFormatInfo* pCFI)
{
    HRESULT hr = S_OK;
    BOOL    fHasLayout      = ElementNeedsLayout(pCFI);
    BOOL    fNeedsOuter     = fHasLayout && (!_fSite || (TestClassFlag(ELEMENTDESC_TEXTSITE) && !TestClassFlag(ELEMENTDESC_TABLECELL)));
    BOOL    fHasLeftIndent  = FALSE;
    BOOL    fHasRightIndent = FALSE;
    BOOL    fIsBlockElement = DetermineBlockness(pCFI);
    LONG    lFontHeight     = 1;
    CDocument* pDoc         = Doc();
    BOOL    fComputeFFOnly  = pCFI->_pNodeContext->_iCF != -1;

    Assert(pCFI->_pNodeContext->Element() == this);

    if((!!pCFI->_pff->_fHasLayout != !!fHasLayout) || (!!pCFI->_pff->_fBlockNess != !!fIsBlockElement))
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fHasLayout = fHasLayout;
        pCFI->_ff()._fBlockNess = fIsBlockElement;
        pCFI->UnprepareForDebug();
    }

    if(!fIsBlockElement && HasFlag(TAGDESC_LIST))
    {
        // If the current list is not a block element, and it is not parented by
        // any block-list elements, then we want the LI's inside to be treated
        // like naked LI's. To do this we have to set cuvOffsetPoints to 0.
        CTreeNode* pNodeList = GetMarkup()->SearchBranchForCriteria(
            pCFI->_pNodeContext->Parent(), IsBlockListElement);
        if(!pNodeList)
        {
            pCFI->PrepareParaFormat();
            pCFI->_pf()._cuvOffsetPoints.SetValue(0, CUnitValue::UNIT_POINT);
            pCFI->_pf()._cListing.SetNotInList();
            pCFI->_pf()._cListing.SetStyle(styleListStyleTypeDisc);
            pCFI->UnprepareForDebug();
        }
        else
        {
            styleListStyleType listType = pNodeList->GetFancyFormat()->_ListType;
            WORD wLevel = (WORD)pNodeList->GetParaFormat()->_cListing.GetLevel();

            pCFI->PrepareParaFormat();
            pCFI->UnprepareForDebug();
        }
    }

    if(!fComputeFFOnly)
    {
        if(pCFI->_fDisplayNone && !pCFI->_pcf->_fDisplayNone)
        {
            pCFI->PrepareCharFormat();
            pCFI->_cf()._fDisplayNone = TRUE;
            pCFI->UnprepareForDebug();
        }

        if(pCFI->_fVisibilityHidden != unsigned(pCFI->_pcf->_fVisibilityHidden))
        {
            pCFI->PrepareCharFormat();
            pCFI->_cf()._fVisibilityHidden = pCFI->_fVisibilityHidden;
            pCFI->UnprepareForDebug();
        }

        if(fNeedsOuter)
        {
            if(pCFI->_fPre != pCFI->_ppf->_fPreInner)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fPreInner = pCFI->_fPre;
                pCFI->UnprepareForDebug();
            }

            if(!!pCFI->_ppf->_fInclEOLWhiteInner != (pCFI->_fInclEOLWhite||TestClassFlag(ELEMENTDESC_SHOWTWS)))
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fInclEOLWhiteInner = !pCFI->_pf()._fInclEOLWhiteInner;
                pCFI->UnprepareForDebug();
            }

            // NO WRAP
            if(pCFI->_fNoBreak)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fNoBreakInner = TRUE;
                pCFI->UnprepareForDebug();
            }

        }
        else
        {
            if(pCFI->_fPre && (!pCFI->_ppf->_fPre || !pCFI->_ppf->_fPreInner))
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fPre = pCFI->_pf()._fPreInner = TRUE;
                pCFI->UnprepareForDebug();
            }

            if(pCFI->_fInclEOLWhite && (!pCFI->_ppf->_fInclEOLWhite || !pCFI->_ppf->_fInclEOLWhiteInner))
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fInclEOLWhite = pCFI->_pf()._fInclEOLWhiteInner = TRUE;
                pCFI->UnprepareForDebug();
            }

            if(pCFI->_fNoBreak && (!pCFI->_pcf->_fNoBreak || !pCFI->_pcf->_fNoBreakInner))
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fNoBreak = pCFI->_cf()._fNoBreakInner = TRUE;
                pCFI->UnprepareForDebug();
            }
        }

        if(pCFI->_fRelative)
        {
            pCFI->PrepareCharFormat();
            pCFI->_cf()._fRelative = TRUE;
            pCFI->UnprepareForDebug();
        }
    }

    // Note: (srinib) Currently table cell's do not support margins,
    // based on the implementation, this may change.
    if(fIsBlockElement && pCFI->_pff->_fHasMargins && Tag()!=ETAG_BODY)
    {
        pCFI->PrepareFancyFormat();

        // MARGIN-TOP - on block elements margin top is treated as before space
        if(!pCFI->_pff->_cuvMarginTop.IsNullOrEnum())
        {
            pCFI->_ff()._cuvSpaceBefore = pCFI->_pff->_cuvMarginTop;
        }

        // MARGIN-BOTTOM - on block elements margin top is treated as after space
        if(!pCFI->_pff->_cuvMarginBottom.IsNullOrEnum())
        {
            pCFI->_ff()._cuvSpaceAfter = pCFI->_pff->_cuvMarginBottom;
        }

        // MARGIN-LEFT - on block elements margin left is treated as left indent
        if(!pCFI->_pff->_cuvMarginLeft.IsNullOrEnum())
        {
            // We handle the various data types below when we accumulate values.
            fHasLeftIndent = TRUE;
        }

        // MARGIN-RIGHT - on block elements margin right is treated as right indent
        if(!pCFI->_pff->_cuvMarginRight.IsNullOrEnum())
        {
            // We handle the various data types below when we accumulate values.
            fHasRightIndent = TRUE;
        }

        pCFI->UnprepareForDebug();
    }

    if(!fComputeFFOnly)
    {
        if(!fHasLayout)
        {
            // PADDING / BORDERS
            //
            // For padding, set _fPadBord flag if CFI _fPadBord is set. Values have
            // already been copied. It always goes on inner.
            if(pCFI->_fPadBord && fIsBlockElement && !pCFI->_ppf->_fPadBord)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fPadBord = TRUE;
                pCFI->UnprepareForDebug();
            }

            // BACKGROUND
            //
            // Sites draw their own background, so we don't have to inherit their
            // background info. Always goes on inner.
            if(pCFI->_fHasBgColor && !pCFI->_pcf->_fHasBgColor)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fHasBgColor = TRUE;
                pCFI->UnprepareForDebug();
            }

            if(pCFI->_fHasBgImage)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fHasBgImage = TRUE;
                pCFI->UnprepareForDebug();
            }
        }
    }

    // FLOAT
    //
    // BUGBUG -- This code needs to move back to ApplyFormatInfoProperty to ensure
    // that the alignment follows the correct ordering. (lylec)
    if(fHasLayout && pCFI->_pff->_bStyleFloat!=styleStyleFloatNotSet)
    {
        htmlControlAlign    hca   = htmlControlAlignNotSet;
        BOOL                fDoIt = TRUE;

        switch(pCFI->_pff->_bStyleFloat)
        {
        case styleStyleFloatLeft:
            hca = htmlControlAlignLeft;
            if(fIsBlockElement)
            {
                pCFI->_bCtrlBlockAlign = htmlBlockAlignLeft;
            }
            break;

        case styleStyleFloatRight:
            hca = htmlControlAlignRight;
            if(fIsBlockElement)
            {
                pCFI->_bCtrlBlockAlign = htmlBlockAlignRight;
            }
            break;

        case styleStyleFloatNone:
            hca = htmlControlAlignNotSet;
            break;

        default:
            fDoIt = FALSE;
        }

        if(fDoIt)
        {
            ApplySiteAlignment(pCFI, hca, this);
            pCFI->_fCtrlAlignLast = TRUE;

            // Autoclear works for float from CSS.  Navigator doesn't
            // autoclear for HTML floating.  Another annoying Nav compat hack.
            if(!pCFI->_pff->_fCtrlAlignFromCSS)
            {
                pCFI->PrepareFancyFormat();
                pCFI->_ff()._fCtrlAlignFromCSS = TRUE;
                pCFI->UnprepareForDebug();
            }
        }
    }

    // ALIGNMENT
    //
    // Alignment is tricky because DISPID_CONTROLALIGN should only set the
    // control align if it has layout, but sets the block alignment if it's
    // not.  Also, if the element has TAGDESC_OWNLINE then DISPID_CONTROLALIGN
    // sets _both_ the control alignment and block alignment.  However, you
    // can still have inline sites (that are not block elements) that have the
    // OWNLINE flag set.  Also, if both CONTROLALIGN and BLOCKALIGN are set,
    // we must remember the order they were applied.  The last kink is that
    // HR's break the pattern because they're not block elements but
    // DISPID_BLOCKALIGN does set the block align for them.
    BOOL fOwnLine = HasFlag(TAGDESC_OWNLINE) && fHasLayout;

    if(fHasLayout)
    {
        if(pCFI->_pff->_bControlAlign != pCFI->_bControlAlign)
        {
            pCFI->PrepareFancyFormat();
            pCFI->_ff()._bControlAlign = pCFI->_bControlAlign;
            pCFI->UnprepareForDebug();
        }

        // If a site is positioned explicitly (absolute) it
        // overrides control alignment. We do that by simply turning off
        // control alignment.
        if(pCFI->_pff->_bControlAlign!=htmlControlAlignNotSet && (pCFI->_pff->_bControlAlign==htmlControlAlignRight || pCFI->_pff->_bControlAlign==htmlControlAlignLeft))
        {
            pCFI->PrepareFancyFormat();

            if(pCFI->_pff->_bPositionType==stylePositionabsolute && Tag()!=ETAG_BODY)
            {
                pCFI->_ff()._bControlAlign = htmlControlAlignNotSet;
                pCFI->_ff()._fAlignedLayout = FALSE;
            }
            else
            {
                pCFI->_ff()._fAlignedLayout = (Tag()!=ETAG_HR && Tag()!=ETAG_LEGEND);
            }

            pCFI->UnprepareForDebug();

        }
    }

    if(!fComputeFFOnly)
    {
        if(fHasLayout && (fNeedsOuter || IsRunOwner()))
        {
            if(pCFI->_ppf->_bBlockAlign!=pCFI->_bCtrlBlockAlign || pCFI->_ppf->_bBlockAlignInner!=pCFI->_bBlockAlign)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._bBlockAlign = pCFI->_bCtrlBlockAlign;
                pCFI->_pf()._bBlockAlignInner = pCFI->_bBlockAlign;
                pCFI->UnprepareForDebug();
            }
        }
        else if(fIsBlockElement || fOwnLine)
        {
            BYTE bAlign = pCFI->_bBlockAlign;

            if(((!fIsBlockElement && Tag()!=ETAG_HR) || pCFI->_fCtrlAlignLast)  && (fOwnLine||!fHasLayout))
            {
                bAlign = pCFI->_bCtrlBlockAlign;
            }

            if(pCFI->_ppf->_bBlockAlign!=bAlign || pCFI->_ppf->_bBlockAlignInner!=bAlign)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._bBlockAlign = bAlign;
                pCFI->_pf()._bBlockAlignInner = bAlign;
                pCFI->UnprepareForDebug();
            }
        }

        // DIRECTION
        if(fHasLayout && (fNeedsOuter || IsRunOwner()))
        {
            if((fIsBlockElement && pCFI->_ppf->_fRTL!=pCFI->_pcf->_fRTL) || (pCFI->_ppf->_fRTLInner!=pCFI->_pcf->_fRTL))
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fRTLInner = pCFI->_pcf->_fRTL;
                // paulnel - we only set the inner for these guys because the
                //           positioning of the layout is determined by the
                //           parent and not by the outter _fRTL
                pCFI->UnprepareForDebug();
            }
        }
        else if(fIsBlockElement || fOwnLine)
        {
            if(pCFI->_ppf->_fRTLInner!=pCFI->_pcf->_fRTL || pCFI->_ppf->_fRTL!=pCFI->_pcf->_fRTL)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fRTLInner = pCFI->_pcf->_fRTL;
                pCFI->_pf()._fRTL = pCFI->_pcf->_fRTL;
                pCFI->UnprepareForDebug();
            }
        }
        if(pCFI->_fBidiEmbed!=pCFI->_pcf->_fBidiEmbed || pCFI->_fBidiOverride!=pCFI->_pcf->_fBidiOverride)
        {
            pCFI->PrepareCharFormat();
            pCFI->_cf()._fBidiEmbed = pCFI->_fBidiEmbed;
            pCFI->_cf()._fBidiOverride = pCFI->_fBidiOverride;
            pCFI->UnprepareForDebug();
        }

        // TEXTINDENT

        // We used to apply text-indent only to block elems; now we apply regardless because text-indent
        // is always inherited, meaning inline elems can end up having text-indent in their PF 
        // (via format inheritance and not Apply).  If we don't allow Apply() to set it on inlines,
        // there's no way to change what's inherited.  This provides a workaround for bug #67276.
        if(!pCFI->_cuvTextIndent.IsNull())
        {
            pCFI->PrepareParaFormat();
            pCFI->_pf()._cuvTextIndent = pCFI->_cuvTextIndent;
            pCFI->UnprepareForDebug();
        }

        if(fIsBlockElement)
        {
            // TEXTJUSTIFY
            if(pCFI->_uTextJustify)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._uTextJustify = pCFI->_uTextJustify;
                pCFI->UnprepareForDebug();
            }

            // TEXTJUSTIFYTRIM
            if(pCFI->_uTextJustifyTrim)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._uTextJustifyTrim = pCFI->_uTextJustifyTrim;
                pCFI->UnprepareForDebug();
            }

            // TEXTKASHIDA
            if(!pCFI->_cuvTextKashida.IsNull())
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._cuvTextKashida = pCFI->_cuvTextKashida;
                pCFI->UnprepareForDebug();
            }
        }
    }

    // Process any image urls (for background images, li's, etc).
    if(!pCFI->_cstrBgImgUrl.IsNull())
    {
        pCFI->PrepareFancyFormat();
        pCFI->ProcessImgUrl(this, pCFI->_cstrBgImgUrl,
            DISPID_A_BGURLIMGCTXCACHEINDEX, &pCFI->_ff()._lImgCtxCookie, fHasLayout);
        pCFI->UnprepareForDebug();
        pCFI->_cstrBgImgUrl.Free();
    }

    if(!pCFI->_cstrLiImgUrl.IsNull())
    {
        pCFI->PrepareParaFormat();
        pCFI->ProcessImgUrl(this, pCFI->_cstrLiImgUrl,
            DISPID_A_LIURLIMGCTXCACHEINDEX, &pCFI->_pf()._lImgCookie, fHasLayout);
        pCFI->UnprepareForDebug();
        pCFI->_cstrLiImgUrl.Free();
    }

    // ******** ACCUMULATE VALUES **********
    if(!fComputeFFOnly)
    {
        // LEFT/RIGHT indents
        if(fHasLeftIndent)
        {
            CUnitValue cuv = pCFI->_pff->_cuvMarginLeft;

            pCFI->PrepareParaFormat();

            Assert (!cuv.IsNullOrEnum());

            // LEFT INDENT
            switch(cuv.GetUnitType() )
            {
            case CUnitValue::UNIT_PERCENT:
                pCFI->_pf()._cuvLeftIndentPercent.SetValue(
                    pCFI->_pf()._cuvLeftIndentPercent.GetUnitValue()+cuv.GetUnitValue(),
                    CUnitValue::UNIT_PERCENT);
                break;

            case CUnitValue::UNIT_EM:
            case CUnitValue::UNIT_EX:
                if(lFontHeight == 1)
                {
                    lFontHeight = pCFI->_pcf->GetHeightInTwips(pDoc);
                }
                // Intentional fall-through...
            default:
                hr = cuv.ConvertToUnitType(CUnitValue::UNIT_POINT, 0,
                    CUnitValue::DIRECTION_CX, lFontHeight);
                if(hr)
                {
                    goto Cleanup;
                }

                pCFI->_pf()._cuvLeftIndentPoints.SetValue(
                    pCFI->_pf()._cuvLeftIndentPoints.GetUnitValue()+cuv.GetUnitValue(),
                    CUnitValue::UNIT_POINT);
            }
            pCFI->UnprepareForDebug();
        }

        // RIGHT INDENT
        if(fHasRightIndent)
        {
            CUnitValue cuv = pCFI->_pff->_cuvMarginRight;

            pCFI->PrepareParaFormat();

            Assert(!cuv.IsNullOrEnum());

            switch(cuv.GetUnitType())
            {
            case CUnitValue::UNIT_PERCENT:
                pCFI->_pf()._cuvRightIndentPercent.SetValue(
                    pCFI->_pf()._cuvRightIndentPercent.GetUnitValue()+cuv.GetUnitValue(),
                    CUnitValue::UNIT_PERCENT);
                break;

            case CUnitValue::UNIT_EM:
            case CUnitValue::UNIT_EX:
                if(lFontHeight == 1)
                {
                    lFontHeight = pCFI->_pcf->GetHeightInTwips(pDoc);
                }
                // Intentional fall-through...
            default:
                hr = cuv.ConvertToUnitType(CUnitValue::UNIT_POINT, 0,
                    CUnitValue::DIRECTION_CX, lFontHeight);
                if(hr)
                {
                    goto Cleanup;
                }

                pCFI->_pf()._cuvRightIndentPoints.SetValue(
                    pCFI->_pf()._cuvRightIndentPoints.GetUnitValue()+cuv.GetUnitValue(),
                    CUnitValue::UNIT_POINT);
            }

            pCFI->UnprepareForDebug();
        }

        if(Tag() == ETAG_LI)
        {
            pCFI->PrepareParaFormat();
            pCFI->_pf()._cuvNonBulletIndentPoints.SetValue(0, CUnitValue::UNIT_POINT);
            pCFI->_pf()._cuvNonBulletIndentPercent.SetValue(0, CUnitValue::UNIT_PERCENT);
            pCFI->UnprepareForDebug();
        }

        // LINE HEIGHT
        switch(pCFI->_pcf->_cuvLineHeight.GetUnitType())
        {
        case CUnitValue::UNIT_EM:
        case CUnitValue::UNIT_EX:
            pCFI->PrepareCharFormat();
            if(lFontHeight == 1)
            {
                lFontHeight = pCFI->_cf().GetHeightInTwips(pDoc);
            }
            hr = pCFI->_cf()._cuvLineHeight.ConvertToUnitType(CUnitValue::UNIT_POINT, 1,
                CUnitValue::DIRECTION_CX, lFontHeight);
            pCFI->UnprepareForDebug();
            break;

        case CUnitValue::UNIT_PERCENT:
            {
                pCFI->PrepareCharFormat();
                if(lFontHeight == 1)
                {
                    lFontHeight = pCFI->_cf().GetHeightInTwips(pDoc);
                }

                // The following line of code does multiple things:
                //
                // 1) Takes the height in twips and applies the percentage scaling to it
                // 2) However, the percentages are scaled so we divide by the unit_percent
                //    scale multiplier
                // 3) Remember that its percent, so we need to divide by 100. Doing this
                //    gives us the desired value in twips.
                // 4) Dividing that by 20 and we get points.
                // 5) This value is passed down to SetPoints which will then scale it by the
                //    multiplier for points.
                //
                // (whew!)
                pCFI->_cf()._cuvLineHeight.SetPoints(MulDivQuick(lFontHeight,
                    pCFI->_cf()._cuvLineHeight.GetUnitValue(),
                    20*100*LONG(CUnitValue::TypeNames[CUnitValue::UNIT_PERCENT].wScaleMult)));

                pCFI->UnprepareForDebug();
            }
            break;
        }
    }

    if(pCFI->_pff->_bPositionType==stylePositionrelative
        || pCFI->_pff->_bPositionType==stylePositionabsolute
        || Tag()==ETAG_ROOT || Tag()==ETAG_BODY)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fPositioned = TRUE;
        pCFI->UnprepareForDebug();
    }

    // cache important values
    if(pCFI->_pff->_fHasLayout)
    {
        if(!TestClassFlag(ELEMENTDESC_NEVERSCROLL)
            && (pCFI->_pff->_bOverflowX==styleOverflowAuto
            || pCFI->_pff->_bOverflowX==styleOverflowScroll
            || pCFI->_pff->_bOverflowY==styleOverflowAuto
            || pCFI->_pff->_bOverflowY==styleOverflowScroll
            || (pCFI->_pff->_bOverflowX==styleOverflowHidden)
            || (pCFI->_pff->_bOverflowY==styleOverflowHidden)
            || (TestClassFlag(ELEMENTDESC_CANSCROLL) && !pCFI->_pff->_fNoScroll)))
        {
            pCFI->PrepareFancyFormat();
            pCFI->_ff()._fScrollingParent = TRUE;
            pCFI->UnprepareForDebug();
        }
    }

    if(pCFI->_pff->_fScrollingParent || pCFI->_pff->_fPositioned)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fZParent = TRUE;
        pCFI->UnprepareForDebug();
    }

    if(pCFI->_pff->_fPositioned
        && (pCFI->_pff->_bPositionType!=stylePositionabsolute
        || (pCFI->_pff->_cuvTop.IsNullOrEnum() && pCFI->_pff->_cuvBottom.IsNullOrEnum())
        || (pCFI->_pff->_cuvLeft.IsNullOrEnum() && pCFI->_pff->_cuvRight.IsNullOrEnum())))
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fAutoPositioned = TRUE;
        pCFI->UnprepareForDebug();
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
// Member:      GetBgImgCtx()
//
// Synopsis:    Return the image ctx if the element has a background Image
//
//-----------------------------------------------------------------------------
CImgCtx* CElement::GetBgImgCtx()
{
    if(_afxGlobalData._fHighContrastMode)
    {
        return NULL; // in high contrast mode there is no background image
    }

    long lCookie = GetFirstBranch()->GetFancyFormat()->_lImgCtxCookie;

    return lCookie?Doc()->GetUrlImgCtx(lCookie):NULL;
}

FOCUS_ITEM CElement::GetMnemonicTarget()
{
    BOOL fNotify = FALSE;
    FOCUS_ITEM fi;

    fi.pElement = this;
    fi.lSubDivision = 0;

    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        fNotify = TRUE;
    }
    else
    {
        switch(Tag())
        {
        case ETAG_AREA:
        case ETAG_LABEL:
        case ETAG_LEGEND:
            fNotify = TRUE;
            break;
        }
    }
    if(fNotify)
    {
        SendNotification(NTYPE_ELEMENT_QUERYMNEMONICTARGET, &fi);
    }
    return fi;
}

//+---------------------------------------------------------------------------
//
// Member:      HandleMnemonic
//
//  Synopsis:   This function is called because the user tried to navigate to
//              this element. There at least four ways the user can do this:
//                  1) by pressing Tab or Shift+Tab
//                  2) by pressing the access key for this element
//                  3) by clicking on a label associated with this element
//                  4) by pressing the access key for a label associated with
//                     this element
//
//              Typically, this function sets focus/currency to this element.
//              Click-able elements usually override this function to call
//              DoClick() on themselves if fTab is FALSE (i.e. navigation
//              happened due to reasons other than tabbing).
//
//----------------------------------------------------------------------------
HRESULT CElement::HandleMnemonic(CMessage* pmsg, BOOL fDoClick, BOOL* pfYieldFailed/*=NULL*/)
{
    HRESULT hr = S_FALSE;
    CDocument* pDoc = Doc();

    Assert(IsInMarkup());
    Assert(pmsg);

    if(IsFrameTabKey(pmsg))
    {
        goto Cleanup;
    }

    if(pDoc->_pElemCurrent)
    {
        pDoc->_pElemCurrent->LostMnemonic(pmsg); // tell the element that it's losing focus due to a mnemonic
    }

    if(IsEditable(TRUE))
    {
        if(Tag()==ETAG_ROOT || Tag()==ETAG_BODY)
        {
            Assert(Tag()!=ETAG_ROOT || this==pDoc->PrimaryRoot());
            hr = BecomeCurrentAndActive(NULL, pmsg->lSubDivision, TRUE);
        }
        else
        {
            // BUGBUG (jenlc)
            // Currently in design mode, only element with layout (site)
            // can be tabbed to.
            CMarkup*    pMarkup = GetMarkup();
            CElement*   pRootElement = pMarkup->Root();

            // BugFix 14600 / 14496 (JohnBed) 03/26/98
            // If the current element does not have layout, just fall
            // out of this so the parent element can handle the mnemonic
            CLayout* pLayout = GetUpdatedLayout();

            if(!pLayout)
            {
                goto Cleanup;
            }

            Assert(pRootElement);

            hr = pRootElement->BecomeCurrentAndActive();

            if(hr)
            {
                goto Cleanup;
            }

            // Site-select the element
            // BUGBUG (MohanB) Make this into a function
            {
                CMarkupPointer  ptrStart(pDoc);
                CMarkupPointer  ptrEnd(pDoc);
                IMarkupPointer* pIStart;
                IMarkupPointer* pIEnd;

                hr = ptrStart.MoveAdjacentToElement(this, ELEM_ADJ_BeforeBegin);
                if(hr)
                {
                    goto Cleanup;
                }
                hr = ptrEnd.MoveAdjacentToElement(this, ELEM_ADJ_AfterEnd);
                if(hr)
                {
                    goto Cleanup;
                }

                Verify(S_OK == ptrStart.QueryInterface(IID_IMarkupPointer, (void**)&pIStart));
                Verify(S_OK == ptrEnd.QueryInterface(IID_IMarkupPointer, (void**)&pIEnd));
                hr = pDoc->Select(pIStart, pIEnd, SELECTION_TYPE_Control);
                if(hr)
                {
                    if(hr == E_INVALIDARG)
                    {
                        // This element was not site-selectable. Return S_FALSE, so
                        // that we can try the next element
                        hr = S_FALSE;
                    }
                }
                pIStart->Release();
                pIEnd->Release();
                if(hr)
                {
                    goto Cleanup;
                }
            }

            hr = ScrollIntoView();
        }
    }
    else
    {
        Assert(IsFocussable(pmsg->lSubDivision));

        hr = BecomeCurrentAndActive(pmsg, pmsg->lSubDivision, TRUE, pfYieldFailed);
        if(hr)
        {
            goto Cleanup;
        }

        hr = ScrollIntoView();
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        if(fDoClick && _fActsLikeButton)
        {
            DoClick(pmsg);
        }

        hr = GotMnemonic(pmsg);
    }

Cleanup:

    RRETURN1(hr, S_FALSE);
}

HRESULT CElement::GotMnemonic(CMessage* pMessage)
{
    // Send only to the listeners
    switch(Tag())
    {
    case ETAG_INPUT:
    case ETAG_TEXTAREA:
        {
            CSetFocus sf;

            sf._pMessage = pMessage;
            sf._hr = S_OK;
            SendNotification(NTYPE_ELEMENT_GOTMNEMONIC, &sf);
            return sf._hr;
        }
    }
    return S_OK;
}

HRESULT CElement::LostMnemonic(CMessage* pMessage)
{
    // Send only to the listeners
    switch(Tag())
    {
    case ETAG_INPUT:
    case ETAG_TEXTAREA:
        {
            CSetFocus sf;

            sf._pMessage = pMessage;
            sf._hr = S_OK;
            SendNotification(NTYPE_ELEMENT_LOSTMNEMONIC, &sf);
            return sf._hr;
        }
    }
    return S_OK;
}

BOOL CElement::MatchAccessKey(CMessage* pmsg)
{
    BOOL    fMatched = FALSE;
    LPCTSTR lpAccKey = GetAAaccessKey();
    WCHAR   chKey = (TCHAR)pmsg->wParam;

    // Raid 57053
    // If we are in HTML dialog, accessKey can be matched with/without
    // SHIFT/CTRL/ALT keys.

    // 60711 - Translate the virtkey to unicode for foreign languages.
    // We only test for 0x20 (space key) or above. This way we avoid
    // coming in here for other system type keys
    if(pmsg->message==WM_SYSKEYDOWN && pmsg->wParam>31)
    {
        BYTE bKeyState[256];
        if(GetKeyboardState(bKeyState))
        {
            WORD cBuf[2];
            int cchBuf;

            UINT wScanCode = MapVirtualKeyEx(pmsg->wParam, VIRTKEY_TO_SCAN, GetKeyboardLayout(0));
            cchBuf = ToAscii(pmsg->wParam, wScanCode, bKeyState, cBuf, 0);

            if(cchBuf == 1)
            {
                WCHAR wBuf[2];
                UINT  uKbdCodePage = CP_ACP;

                MultiByteToWideChar(uKbdCodePage, 0, (char*)cBuf, 2, wBuf, 2);
                chKey = wBuf[0];
            }
        }
    }   

    if((pmsg->message==WM_SYSKEYDOWN)
        && lpAccKey
        && lpAccKey[0] &&
        TOTLOWER((TCHAR)chKey)==TOTLOWER(lpAccKey[0])
        && ISVALIDKEY(pmsg->lParam))
    {
        fMatched = TRUE;
    }
    return fMatched;
}

HRESULT CElement::OnTabIndexChange()
{
    long        i;
    HRESULT     hr = S_OK;
    CDocument*  pDoc = Doc();

    for(i=0; i<pDoc->_aryFocusItems.Size(); i++)
    {
        if(pDoc->_aryFocusItems[i].pElement == this)
        {
            pDoc->_aryFocusItems.Delete(i);
        }
    }

    hr = pDoc->InsertFocusArrayItem(this);
    if(hr) 
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//----------------------------------------------------------------
//
//      Member:         CElement::GetStyleObject
//
//  Description Helper function to create the .style sub-object
//
//----------------------------------------------------------------
HRESULT CElement::GetStyleObject(CStyle** ppStyle)
{
    HRESULT hr;
    CStyle* pStyle = 0;

    hr = GetPointerAt(FindAAIndex(DISPID_INTERNAL_CSTYLEPTRCACHE,CAttrValue::AA_Internal), (void**)&pStyle);

    if(!pStyle)
    {
        pStyle = new CStyle(this, DISPID_INTERNAL_INLINESTYLEAA, 0);
        if(!pStyle)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }
        hr = AddPointer(DISPID_INTERNAL_CSTYLEPTRCACHE, (void*)pStyle, CAttrValue::AA_Internal);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    if(!hr)
    {
        *ppStyle = pStyle;
    }
    else
    {
        *ppStyle = 0;
        delete pStyle;
    }

    RRETURN(hr);
}

CAttrArray* CElement::GetInLineStyleAttrArray(void)
{
    CAttrArray* pAA = NULL;

    // Apply the in-line style attributes
    AAINDEX aaix = FindAAIndex(DISPID_INTERNAL_INLINESTYLEAA,
        CAttrValue::AA_AttrArray);
    if(aaix != AA_IDX_UNKNOWN)
    {
        CAttrValue* pAttrValue = (CAttrValue*)**GetAttrArray();
        pAA = pAttrValue[aaix].GetAA();
    }
    return pAA;
}

CAttrArray** CElement::CreateStyleAttrArray(DISPID dispID)
{
    AAINDEX aaix = AA_IDX_UNKNOWN;
    if((aaix=FindAAIndex(dispID, CAttrValue::AA_AttrArray))
        == AA_IDX_UNKNOWN)
    {
        CAttrArray* pAA = new CAttrArray;
        AddAttrArray(dispID, pAA, CAttrValue::AA_AttrArray);
        aaix = FindAAIndex(dispID, CAttrValue::AA_AttrArray);
    }
    if(aaix == AA_IDX_UNKNOWN)
    {
        return NULL;
    }
    else
    {
        CAttrValue* pAttrValue = (CAttrValue*)**GetAttrArray();
        return (CAttrArray**)(pAttrValue[aaix].GetppAA());
    }
}

// abstract name property helpers
STDMETHODIMP CElement::put_name(BSTR v)
{
    return s_propdescCElementsubmitName.b.SetStringProperty(v, this, (CVoid*)(void*)(GetAttrArray()));
}

STDMETHODIMP CElement::get_name(BSTR* p)
{
    RRETURN(SetErrorInfo(FormsAllocString(GetAAsubmitname(), p)));
}

htmlBlockAlign CElement::GetParagraphAlign(BOOL fOuter)
{
    return GetFirstBranch()->GetParagraphAlign(fOuter);
}

htmlControlAlign CElement::GetSiteAlign()
{
    return GetFirstBranch()->GetSiteAlign();
}

BOOL CElement::IsInlinedElement()
{
    return GetFirstBranch()->IsInlinedElement();
}

BOOL CElement::IsPositioned(void)
{
    return GetFirstBranch()->IsPositioned();
}

BOOL CElement::IsPositionStatic(void)
{
    return GetFirstBranch()->IsPositionStatic();
}

BOOL CElement::IsAbsolute(void)
{
    return GetFirstBranch()->IsAbsolute();
}

BOOL CElement::IsRelative(void)
{
    return GetFirstBranch()->IsRelative();
}

BOOL CElement::IsInheritingRelativeness(void)
{
    return GetFirstBranch()->IsInheritingRelativeness();
}

// Returns TRUE if we're the body or a scrolling DIV or SPAN
BOOL CElement::IsScrollingParent(void)
{
    return GetFirstBranch()->IsScrollingParent();
}

BOOL CElement::IsClipParent(void)
{
    return GetFirstBranch()->IsClipParent();
}

BOOL CElement::IsZParent(void)
{
    return GetFirstBranch()->IsZParent();
}

BOOL CElement::IsLocked()
{
    // only absolute sites can be locked.
    if(IsAbsolute())
    {
        return FALSE;
    }
    return FALSE;
}

BOOL CElement::IsDisplayNone()
{
    CTreeNode* pNode = GetFirstBranch();

    return pNode?pNode->IsDisplayNone():TRUE;
}

BOOL CElement::IsVisibilityHidden()
{
    CTreeNode* pNode = GetFirstBranch();

    return pNode ? pNode->IsVisibilityHidden() : TRUE;
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::QueryStatus
//
//  Synopsis:   Called to discover if a given command is supported
//              and if it is, what's its state.  (disabled, up or down)
//
//--------------------------------------------------------------------------
HRESULT CElement::QueryStatus(
      GUID*         pguidCmdGroup,
      ULONG         cCmds,
      MSOCMD        rgCmds[],
      MSOCMDTEXT*   pcmdtext)
{
    Trace0("CSite::QueryStatus\n");

    Assert(IsCmdGroupSupported(pguidCmdGroup));
    Assert(cCmds == 1);

    MSOCMD* pCmd = &rgCmds[0];
    HRESULT hr = S_OK;

    Assert(!pCmd->cmdf);

    switch(IDMFromCmdID(pguidCmdGroup, pCmd->cmdID))
    {
    case IDM_DYNSRCPLAY:
    case IDM_DYNSRCSTOP:
        // The selected site wes not an image site, return disabled
        pCmd->cmdf = MSOCMDSTATE_DISABLED;
        break;

    case  IDM_SIZETOCONTROL:
    case  IDM_SIZETOCONTROLHEIGHT:
    case  IDM_SIZETOCONTROLWIDTH:
        // will be executed only if the selection is not a control range
        pCmd->cmdf = MSOCMDSTATE_DISABLED;
        break;

    case IDM_SAVEBACKGROUND:
    case IDM_COPYBACKGROUND:
    case IDM_SETDESKTOPITEM:
        pCmd->cmdf = GetNearestBgImgCtx() ? MSOCMDSTATE_UP : MSOCMDSTATE_DISABLED;
        break;

    case IDM_SELECTALL:
        if(HasFlag(TAGDESC_CONTAINER))
        {
            // Do not bubble to parent if this is a container.
            pCmd->cmdf =  DisallowSelection() ? MSOCMDSTATE_DISABLED : MSOCMDSTATE_UP;
        }
        break;
    }

    RRETURN_NOTRACE(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CSite::Exec
//
//  Synopsis:   Called to execute a given command.  If the command is not
//              consumed, it may be routed to other objects on the routing
//              chain.
//
//--------------------------------------------------------------------------
HRESULT CElement::Exec(
       GUID*        pguidCmdGroup,
       DWORD        nCmdID,
       DWORD        nCmdexecopt,
       VARIANTARG*  pvarargIn,
       VARIANTARG*  pvarargOut)
{
    Trace0("CSite::Exec\n");

    /*Assert(IsCmdGroupSupported(pguidCmdGroup)); wlw note*/

    UINT    idm;
    HRESULT hr = OLECMDERR_E_NOTSUPPORTED;

    // default processing
    switch(idm = IDMFromCmdID(pguidCmdGroup, nCmdID))
    {
    case IDM_COPYBACKGROUND:
        {
            /*CImgCtx* pImgCtx = GetNearestBgImgCtx();

            if(pImgCtx)
            {
                CreateImgDataObject(Doc(), pImgCtx, NULL, NULL, NULL);
            } wlw note*/

            hr = S_OK;
            break;
        }
    }
    if(hr != OLECMDERR_E_NOTSUPPORTED)
    {
        goto Cleanup;
    }

    // behaviors
    /*if(HasPeerHolder())
    {
        hr = GetPeerHolder()->ExecMulti(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut);
        if(hr != OLECMDERR_E_NOTSUPPORTED)
        {
            goto Cleanup;
        }
    } wlw note*/

Cleanup:
    RRETURN_NOTRACE(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     IsEditable()
//              IsEditableSlow()==IsEditable(TRUE) (but does not use cached info)
//
//  Synopsis:   returns if the site is editable (in design mode or editAtBrowse)
//  Note:
//       Check if the current element is editable
//       An element is editable has two meanings:
//        1) the element is editable in design mode, can be resized,
//            content can be edited
//        2) the content of the site is editable,
//            usually in browse mode
//
//       IsEditable(FALSE) or IsEditable() checks if the element is
//        Editable as a design mode item like the case 1)
//
//       IsEditable(TRUE) checks if the element is is NOT in case 2)
//       which means that the parent of element is in case 2)
//
//       this function has the following equivalences:
//        IsEditable(TRUE) = Doc()->_fDesignMode
//        IsEditable()     = Doc()->_fDesignMode || _fEditAtBrowse
//                         = !IsInBrowseMode() //CTxtsite method being removed
//
//       If you insist on the Document level editability, you should keep using
//        Doc()->_fDesignMode
//
//-----------------------------------------------------------------------------
BOOL CElement::IsEditable(BOOL fCheckContainerOnly /*=FALSE*/)
{
    // BUGBUIG (MohanB) Since HTMLAREA is gone for now, this function
    // can be made really fast
    return (!fCheckContainerOnly && (_fEditAtBrowse
        || Tag()==ETAG_TXTSLAVE
        && IsInMarkup()
        && GetMarkup()->IsEditable()));
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::TakeCapture
//
//  Synopsis:   Set the element with the mouse capture.
//
//----------------------------------------------------------------------------
void CElement::TakeCapture(BOOL fTake)
{
    CDocument* pDoc = Doc();

    if(fTake)
    {
        pDoc->SetMouseCapture(
            MOUSECAPTURE_METHOD(CElement, HandleCaptureMessage,
            handlecapturemessage),
            (void*)this,
            TRUE);
    }
    else if(pDoc->GetCaptureObject()==this &&::GetCapture()==pDoc->_pInPlace->_hwnd)
    {
        pDoc->ClearMouseCapture((void*)this);
        if(!pDoc->_pElementOMCapture)
        {
            ::ReleaseCapture();
        }
    }
}

BOOL CElement::HasCapture()
{
    CDocument* pDoc = Doc();

    return ((pDoc->GetCaptureObject()==(void*)this) && (::GetCapture()==pDoc->_pInPlace->_hwnd));
}

//+----------------------------------------------------------------------------
//
//  Function:   EnsureInMarkup()
//
//  Synopsis:   Creates a private markup for the element, if it is
//                  outside any markup.
//
//
//  Returns:    S_OK if successful
//
//-----------------------------------------------------------------------------
HRESULT CElement::EnsureInMarkup()
{
    HRESULT hr = S_OK;

    if(!IsInMarkup())
    {
        hr = Doc()->CreateMarkupWithElement(NULL, this, NULL);

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
//  Member:     CElement::HasCurrency
//
//  Synopsis:   Checks if this element has currency. If this element is a slave,
//              check if its master has currency.
//
//  Notes:      Can't inline, because that would need including the defn of
//              CDoc in element.hxx.
//
//----------------------------------------------------------------------------
BOOL CElement::HasCurrency()
{
    CElement* pElemCurrent = Doc()->_pElemCurrent;

    // BUGBUG (MohanB) Need to generalize this for super-slaves
    // of arbitrary depth, by moving currency to CMarkup from CDoc
    return (this && pElemCurrent && (pElemCurrent==this||pElemCurrent==MarkupMaster()));
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::RequestYieldCurrency
//
//  Synopsis:   Check if OK to Relinquish currency
//
//  Arguments:  BOOl fForce -- if TRUE, force change and don't ask user about
//                             usaveable data.
//
//  Returns:    S_OK        ok to yield currency
//              S_FALSE     ok to yield currency, but user explicitly reverted
//                          the value to what the database has
//              E_*         not ok to yield currency
//
//--------------------------------------------------------------------------
HRESULT CElement::RequestYieldCurrency(BOOL fForce)
{
    HRESULT hr = S_OK;
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::YieldCurrency
//
//  Synopsis:   Relinquish currency
//
//  Arguments:  pElementNew    New Element that wants currency
//
//  Returns:    HRESULT
//
//--------------------------------------------------------------------------
HRESULT CElement::YieldCurrency(CElement* pElemNew)
{
    HRESULT     hr = S_OK;
    CLayout*    pLayout = GetUpdatedLayout();
    CDocument*  pDoc = Doc();
    if(pLayout)
    {
        CFlowLayout* pFlowLayout = pLayout->IsFlowLayout();
        if(pFlowLayout)
        {
            hr = pFlowLayout->YieldCurrencyHelper(pElemNew);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    // Restore the IMC if we've temporarily disabled it.  See BecomeCurrent().
    if(pDoc && pDoc->_pInPlace && !pDoc->_pInPlace->_fDeactivating
        && pDoc->_pInPlace->_hwnd && pElemNew)
    {
        // We don't know if ETAG_OBJECT is editable, so enable input
        // to be ie401 compat.
        if((pElemNew->IsEditable() || pElemNew->Tag()==ETAG_OBJECT) && Doc()->_himcCache)
        {
            ImmAssociateContext(Doc()->_pInPlace->_hwnd, Doc()->_himcCache);
            Doc()->_himcCache = NULL;
        }
    }

    // Hide caret (but don't clear selection!) when focus changes
    {
        CMarkup* pMarkup = GetMarkup();
        if(_etag != ETAG_ROOT)
        {
            if(pDoc->_pCaret)
            {
                pDoc->_pCaret->Hide();
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::YieldUI
//
//  Synopsis:   Relinquish UI, opposite of BecomeUIActive
//
//  Arguments:  pElementNew    New site that wants UI
//
//--------------------------------------------------------------------------
void CElement::YieldUI(CElement* pElemNew)
{
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::BecomeUIActive
//
//  Synopsis:   Force ui activity on the site.
//
//  Notes:      This is the method that external objects should call
//              to force sites to become ui active.
//
//--------------------------------------------------------------------------
HRESULT CElement::BecomeUIActive()
{
    HRESULT hr = S_FALSE;
    CLayout* pLayout = GetUpdatedParentLayout();

    // This site does not care about grabbing the UI.
    // Give the parent a chance.
    if(pLayout)
    {
        hr = pLayout->ElementOwner()->BecomeUIActive();
    }
    else
    {
        hr = GetMarkup()->Root()->BecomeUIActive();
    }
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::NoUIActivate
//
//  Synopsis:   Determines if this site does not allow UIActivation
//
//  Arguments:  none
//
//--------------------------------------------------------------------------
BOOL CElement::NoUIActivate()
{
    CLayout* pLayout = GetUpdatedLayout();

    return pLayout&&pLayout->_fNoUIActivate;
}

BOOL CElement::IsFocussable(long lSubDivision)
{
    // avoid visibilty and other checks for special elements
    if(Tag()==ETAG_ROOT || Tag()==ETAG_DEFAULT)
    {
        return TRUE;
    }

    CDocument* pDoc = Doc();

    FOCUSSABILITY fcDefault = GetDefaultFocussability(Tag());

    if(fcDefault<=FOCUSSABILITY_NEVER
        || !IsInMarkup()
        || !IsEnabled()
        || !IsVisible(TRUE)
        || NoUIActivate()
        || !(this==GetMarkup()->GetElementClient() || GetUpdatedParentLayoutNode()))
    {
        return FALSE;
    }

    // BUGBUG (MohanB) IsVisible() does not work for elements without layout.
    // Until, that is fixed, let's use a partial workaround (to do this right,
    // I probably need to walk up the tree until I hit the first layout parent)
    const CCharFormat* pCF = GetFirstBranch()->GetCharFormat();
    if(pCF->IsDisplayNone() || pCF->IsVisibilityHidden())
    {
        return FALSE;
    }

    // do not  query for focussability if tabIndex is set.
    if(GetAAtabIndex() != htmlTabIndexNotSet)
    {
        return TRUE;
    }

    if(fcDefault < FOCUSSABILITY_FOCUSSABLE)
    {
        return FALSE;
    }

    // Hack for DIV and SPAN which want focus only they have a layout
    // I don't want to send them a queryfocussable because the hack is
    // more obvious here and we will try to get rid of it in IE6
    if((Tag()==ETAG_DIV || Tag()==ETAG_SPAN) && !GetCurLayout())
    {
        return FALSE;
    }

    BOOL fNotify = FALSE;

    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        fNotify = TRUE;
    }
    else
    {
        switch(Tag())
        {
        case ETAG_A:
        case ETAG_IMG:
        case ETAG_SELECT:
            fNotify = TRUE;
            break;
        }
    }
    if(fNotify)
    {
        CQueryFocus qf;

        qf._lSubDivision = lSubDivision;
        qf._fRetVal = TRUE;

        SendNotification(NTYPE_ELEMENT_QUERYFOCUSSABLE, &qf);
        return qf._fRetVal;
    }
    else
    {
        return TRUE;
    }
}

BOOL CElement::IsTabbable(long lSubDivision)
{
    FOCUSSABILITY fcDefault = GetDefaultFocussability(Tag());
    BOOL fDesignMode = IsEditable(TRUE);

    // browse-time tabbing checks for focussability
    if(!IsFocussable(lSubDivision))
    {
        return FALSE;
    }

    long lTabIndex = GetAAtabIndex();

    // Specifying an explicit tabIndex overrides the rest of the checks
    if(lTabIndex != htmlTabIndexNotSet)
    {
        return (lTabIndex>=0);
    }

    if(fcDefault<FOCUSSABILITY_TABBABLE
        && !(fDesignMode && GetCurLayout()/* && Doc()->IsElementSiteSelectable(this)*/))
    {      
        return FALSE;
    }

    BOOL fNotify = FALSE;

    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        fNotify = TRUE;
    }
    else
    {
        switch(Tag())
        {
        case ETAG_BODY:
        case ETAG_IMG:
        case ETAG_INPUT:
            fNotify = TRUE;
            break;
        }
    }
    if(fNotify)
    {
        CQueryFocus qf;

        qf._lSubDivision = lSubDivision;
        qf._fRetVal = TRUE;
        SendNotification(NTYPE_ELEMENT_QUERYTABBABLE, &qf);
        return qf._fRetVal;
    }
    else
    {
        return TRUE;
    }
}

HRESULT CElement::PreBecomeCurrent(long lSubDivision, CMessage* pMessage)
{
    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        CSetFocus sf;

        sf._pMessage = pMessage;
        sf._lSubDivision = lSubDivision;
        sf._hr = S_OK;
        SendNotification(NTYPE_ELEMENT_SETTINGFOCUS, &sf);
        return sf._hr;
    }
    else
    {
        return S_OK;
    }
}

HRESULT CElement::BecomeCurrentFailed(long lSubDivision, CMessage* pMessage)
{
    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        CSetFocus sf;

        sf._pMessage = pMessage;
        sf._lSubDivision = lSubDivision;
        sf._hr = S_OK;
        SendNotification(NTYPE_ELEMENT_SETFOCUSFAILED, &sf);
        return sf._hr;
    }
    else
    {
        return S_OK;
    }
}

HRESULT CElement::PostBecomeCurrent(CMessage* pMessage)
{
    BOOL fNotify = FALSE;

    // Send only to the listeners
    if(TestClassFlag(ELEMENTDESC_OLESITE))
    {
        fNotify = TRUE;
    }
    else
    {
        switch(Tag())
        {
        case ETAG_A:
        case ETAG_BODY:
        case ETAG_BUTTON:
        case ETAG_IMG:
        case ETAG_INPUT:
        case ETAG_SELECT:
            fNotify = TRUE;
            break;
        }
    }
    if(fNotify)
    {
        CSetFocus sf;

        sf._pMessage = pMessage;
        sf._hr = S_OK;
        SendNotification(NTYPE_ELEMENT_SETFOCUS, &sf);
        return sf._hr;
    }
    else
    {
        return S_OK;
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::BecomeCurrent
//
//  Synopsis:   Force currency on the site.
//
//  Notes:      This is the method that external objects should call
//              to force sites to become current.
//
//--------------------------------------------------------------------------
HRESULT CElement::BecomeCurrent(
            long        lSubDivision,
            BOOL*       pfYieldFailed,
            CMessage*   pmsg,
            BOOL        fTakeFocus/*=FALSE*/)
{
    HRESULT     hr = S_FALSE;
    CDocument*  pDoc = Doc();
    CElement*   pElemOld = pDoc->_pElemCurrent;
    long        lSubOld = pDoc->_lSubCurrent;

    Assert(IsInMarkup());

    // Hack for txtslave. The editing code calls BecomeCurrent() on txtslave.
    // For now, make the master current, Need to think about the nested slave tree case
    if(Tag() == ETAG_TXTSLAVE)
    {
        return MarkupMaster()->BecomeCurrent(lSubDivision, pfYieldFailed, pmsg, fTakeFocus);
    }

    Assert(pDoc->_fForceCurrentElem || pDoc->_pInPlace);
    if(!pDoc->_fForceCurrentElem && pDoc->_pInPlace->_fDeactivating)
    {
        goto Cleanup;
    }

    // For now, only the elements in the primary markup can become current
    Assert(GetMarkup() == pDoc->_pPrimaryMarkup);

    if(!IsFocussable(lSubDivision))
    {
        goto Cleanup;
    }

    hr = PreBecomeCurrent(lSubDivision, pmsg);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDoc->SetCurrentElem(this, lSubDivision, pfYieldFailed);
    if(hr)
    {
        BecomeCurrentFailed(lSubDivision, pmsg);
        goto Cleanup;
    }

    // The event might have killed the element
    if(!IsInMarkup())
    {
        goto Cleanup;
    }

    // Do not inhibit if Currency did not change. Onfocus needs to be fired from
    // WM_SETFOCUS handler if clicking in the address bar and back to the same element
    // that previously had the focus.
    pDoc->_fInhibitFocusFiring = (this != pElemOld);

    if(pDoc->_pInPlace && !pDoc->_pInPlace->_fDeactivating && pDoc->_pInPlace->_hwnd)
    {
        if(!IsEditable() && Tag()!=ETAG_OBJECT)
        {
            HIMC himc = ImmGetContext(Doc()->_pInPlace->_hwnd);

            if(himc)
            {
                IUnknown* pUnknown = NULL;
                this->QueryInterface(IID_IUnknown, (void**)&pUnknown);
                pDoc->NotifySelection(SELECT_NOTIFY_DISABLE_IME , pUnknown);
                ReleaseInterface(pUnknown);

                ImmAssociateContext(Doc()->_pInPlace->_hwnd, NULL);
                Doc()->_himcCache = himc;
            }
        }
        else
        {
            if(Doc()->_himcCache)
            {
                ImmAssociateContext(Doc()->_pInPlace->_hwnd, Doc()->_himcCache);
                Doc()->_himcCache = NULL;        
            }

            if(Tag()==ETAG_INPUT || Tag()==ETAG_TEXTAREA)
            {
                SetImeState();
            }
        }
    }

    if(IsEditable(FALSE))
    {
        // An editable element has just become current.
        // We tell the mshtmled.dll about this change in editing "context"
        if(_etag != ETAG_ROOT) // make sure we're done parsing
        {
            // marka - we need to put in whether to select the text or not.
            hr = pDoc->SetEditContext((HasSlaveMarkupPtr()? GetSlaveMarkupPtr()->FirstElement() : this), TRUE, FALSE, TRUE);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    // Take focus only if told to do so, and the element becoming current is not 
    // the root and not an olesite.  Olesite's will do it themselves.  
    if(fTakeFocus && Tag()!=ETAG_ROOT && !TestClassFlag(ELEMENTDESC_OLESITE))
    {
        pDoc->TakeFocus();
    }


    pDoc->_fInhibitFocusFiring = FALSE;
    hr = S_OK;

    if(Tag()!=ETAG_ROOT && pDoc->_pInPlace)
    {
        CLayout* pLayout = GetUpdatedLayout();

        if(pLayout && pLayout->IsFlowLayout())
        {
            hr = pLayout->IsFlowLayout()->BecomeCurrentHelper(lSubDivision, pfYieldFailed, pmsg);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    hr = PostBecomeCurrent(pmsg);

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Method:     CElement::BubbleBecomeCurrent
//
//  Synopsis:   Bubble up BecomeCurrent requests through parent chain.
//
//----------------------------------------------------------------------------
HRESULT CElement::BubbleBecomeCurrent(
          long      lSubDivision,
          BOOL*     pfYieldFailed,
          CMessage* pMessage,
          BOOL      fTakeFocus)
{
    CTreeNode*  pNode = NULL;
    HRESULT     hr=S_OK;

    if(Tag() == ETAG_TXTSLAVE)
    {
        return MarkupMaster()->BubbleBecomeCurrent(lSubDivision, pfYieldFailed, pMessage, fTakeFocus);
    }

    if(!GetFirstBranch())
    {
        goto Cleanup;
    }

    pNode = GetFirstBranch()->Parent();

    hr = BecomeCurrent(lSubDivision, pfYieldFailed, pMessage, fTakeFocus);

    while(hr==S_FALSE && pNode)
    {
        hr = pNode->Element()->BecomeCurrent(0, pfYieldFailed, pMessage, fTakeFocus);
        if(hr == S_FALSE)
        {
            pNode = pNode->Parent();
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

CElement* CElement::GetFocusBlurFireTarget(long lSubDiv)
{
    return this;
}

HRESULT DetectSiteState(CLayout* pLayout)
{
    CElement*   pElement = pLayout->ElementOwner();
    CDocument*  pDoc = pElement->Doc();

    if(!pElement->IsInMarkup())
    {
        return E_UNEXPECTED;
    }

    if(!pElement->IsVisible(TRUE) || !pElement->IsEnabled())
    {
        return CTL_E_CANTMOVEFOCUSTOCTRL;
    }

    return S_OK;
}

HRESULT CElement::focusHelper(long lSubDivision)
{
    HRESULT     hr          = S_OK;
    CLayout*    pLayout     = GetUpdatedLayout();
    CDocument*  pDoc        = Doc();

    if(pLayout)
    {
        // bail out if the site has been detached, or the doc is not yet inplace, etc.
        hr = DetectSiteState(pLayout);
        if(hr)
        {
            goto Cleanup;
        }
    }

    // if called on body, delegate to the window
    if(Tag() == ETAG_BODY)
    {
        /*pDoc->EnsureOmWindow();
        if(pDoc->_pOmWindow)
        {
            hr = pDoc->_pOmWindow->focus();
        } wlw note*/

        goto Cleanup;
    }

    if(!IsFocussable(lSubDivision))
    {
        goto Cleanup;
    }

    pDoc->_fFirstTimeTab = FALSE;
    BecomeCurrent(lSubDivision, NULL, NULL, TRUE);
    SendNotification(NTYPE_ELEMENT_ENSURERECALC);
    ScrollIntoView(SP_MINIMAL, SP_MINIMAL);

    hr = BecomeUIActive();

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetFocusShape
//
//  Synopsis:   Returns the shape of the focus outline that needs to be drawn
//              when this element has focus. This function creates a new
//              CShape-derived object. It is the caller's responsibility to
//              release it.
//
//  Notes:      This default implementation assumes that the element encloses
//              a text range (e.g. anchor, label). Other elements (buttons,
//              body, image, checkbox, radio button, input file, image map
//              area) must override this function to supply the correct shape.
//
//
//----------------------------------------------------------------------------
HRESULT CElement::GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape)
{
    HRESULT hr = S_FALSE;

    Assert(ppShape);

    *ppShape = NULL;

    // By default, provide focus shape only for 
    // 1) elements that had focus shape in IE4 (compat)
    // 2) elements that have a tab index specified (#40434)
    //
    // In IE6, we will introduce a CSS style that lets elements override this
    // this behavior to turn on/off focus shapes.
    if(GetAAtabIndex() < 0)
    {
        switch(Tag())
        {
        case ETAG_A:
        case ETAG_LABEL:
        case ETAG_IMG:
            break;
        default:
            goto Cleanup;
        }
    }

    if(HasLayout())
    {
        CRect rc;
        CLayout* pLayout = GetCurLayout();

        if(!pLayout)
        {
            goto Cleanup;
        }

        pLayout->GetClientRect(&rc);
        if(rc.IsEmpty())
        {
            goto Cleanup;
        }

        CRectShape* pShape = new CRectShape;
        if(!pShape)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        pShape->_rect = rc;
        *ppShape = pShape;

        hr = S_OK;
    }
    else
    {
        long            cpStart, cch;
        CFlowLayout*    pFlowLayout = GetFlowLayout();

        if(!pFlowLayout)
        {
            goto Cleanup;
        }

        cch = GetFirstAndLastCp(&cpStart, NULL);
        // GetFirstAndLast gets a cpStart + 1. We need the element's real cpStart
        // Also, the cch is less than the real cch by 2. We need to add this back
        // in so RegionFromElement can get the correct width of this element.
        hr = pFlowLayout->GetDisplay()->GetWigglyFromRange(pdci, cpStart-1, cch+2, ppShape);
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::BecomeCurrentAndActive
//
//  Synopsis:   Force currency and uiactivity on the site.
//
//--------------------------------------------------------------------------
HRESULT CElement::BecomeCurrentAndActive(
    CMessage*   pmsg,
    long        lSubDivision,
    BOOL        fTakeFocus/*=FALSE*/,
    BOOL*       pfYieldFailed/*=NULL*/)
{
    HRESULT     hr          = S_FALSE;
    CDocument*  pDoc        = Doc();
    CElement*   pElemOld    = pDoc->_pElemCurrent;
    long        lSubOld     = pDoc->_lSubCurrent;

    // Store the old current site in case the new current site cannot
    // become ui-active.  If the becomeuiactive call fails, then we
    // must reset currency to the old guy.
    hr = BecomeCurrent(lSubDivision, pfYieldFailed, pmsg, fTakeFocus);
    if(hr)
    {
        goto Cleanup;
    }

    hr = BecomeUIActive();
    if(OK(hr))
    {
        hr = S_OK;
    }
    else
    {
        if(pElemOld)
        {
            // Don't take focus in this case
            Verify(!pElemOld->BecomeCurrent(lSubOld));
        }
    }
Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Method:     CElement::BubbleBecomeCurrentAndActive
//
//  Synopsis:   Bubble up BecomeCurrentAndActive requests through parent chain.
//
//----------------------------------------------------------------------------
HRESULT CElement::BubbleBecomeCurrentAndActive(CMessage* pMessage, BOOL fTakeFocus)
{
    HRESULT     hr = S_FALSE;
    CTreeNode*  pNode = GetFirstBranch();

    while(hr==S_FALSE && pNode)
    {
        hr = pNode->Element()->BecomeCurrentAndActive(pMessage, 0, fTakeFocus);
        if(hr == S_FALSE)
        {
            pNode = pNode->Parent();
        }
    }
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetSubDivisionCount
//
//  Synopsis:   Get the count of subdivisions for this element
//
//-------------------------------------------------------------------------
HRESULT CElement::GetSubDivisionCount(long* pc)
{
    HRESULT                 hr = S_OK;
    *pc = 0;
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetSubDivisionTabs
//
//  Synopsis:   Get the tabindices of subdivisions for this element
//
//-------------------------------------------------------------------------
HRESULT CElement::GetSubDivisionTabs(long* pTabs, long c)
{
    HRESULT                 hr = S_OK;
    Assert(c == 0);
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::SubDivisionFromPt
//
//  Synopsis:   Perform a hittest of subdivisions for this element
//
//-------------------------------------------------------------------------
HRESULT CElement::SubDivisionFromPt(POINT pt, long* plSub)
{
    HRESULT hr = S_OK;

    *plSub = 0;

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::FindDefaultElem
//
//  Synopsis:   Find the default element
//
//-------------------------------------------------------------------------
CElement* CElement::FindDefaultElem(BOOL fDefault, BOOL fFull/*=FALSE*/)
{
    CElement*       pElem = NULL;
    CDocument*      pDoc = Doc();

    if(!pDoc || !pDoc->_pInPlace || pDoc->_pInPlace->_fDeactivating)
    {
        return NULL;
    }

    if(!fFull && (fDefault?_fDefault:_fCancel))
    {
        return this;
    }

    if(fFull)
    {
        pElem = pDoc->FindDefaultElem(fDefault, FALSE);
        if(fDefault)
        {
            pDoc->_pElemDefault = pElem;
        }
    }
    else
    {
        Assert(fDefault);
        pElem = pDoc->_pElemDefault;
    }

    return pElem;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::SetDefaultElem
//
//  Synopsis:   Set the default element
//  Parameters:
//              fFindNew = TRUE: we need to find a new one
//                       = FALSE: set itself to be the default if appropriate
//
//-------------------------------------------------------------------------
void CElement::SetDefaultElem(BOOL fFindNew /*=FALSE*/)
{
    CElement**      ppElem;
    CDocument*      pDoc = Doc();

    ppElem = &pDoc->_pElemDefault;

    Assert(TestClassFlag(ELEMENTDESC_DEFAULT));
    if(fFindNew)
    {
        // Only find the new when the current is the default and
        // the document is not deactivating
        *ppElem = (*ppElem==this && pDoc->_pInPlace && !pDoc->_pInPlace->_fDeactivating) ?
            FindDefaultElem(TRUE, TRUE) : 0;
        Assert(*ppElem != this);
    }
    else if(!*ppElem || (*ppElem)->GetSourceIndex()>GetSourceIndex())
    {
        if(IsEnabled() && IsVisible(TRUE))
        {
            *ppElem = this;
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetNextSubdivision
//
//  Synopsis:   Finds the next tabbable subdivision which has tabindex
//              == 0.  This is a helper meant to be used by SearchFocusTree.
//
//  Returns:    The next subdivision in plSubNext.  Set to -1, if there is no
//              such subdivision.
//
//  Notes:      lSubDivision coming in can be set to -1 to search for the
//              first possible subdivision.
//
//----------------------------------------------------------------------------
HRESULT CElement::GetNextSubdivision(FOCUS_DIRECTION dir, long lSubDivision, long* plSubNext)
{
    HRESULT hr = S_OK;
    long    c;
    long*   pTabs = NULL;
    long    i;

    hr = GetSubDivisionCount(&c);
    if(hr)
    {
        goto Cleanup;
    }

    if(!c)
    {
        *plSubNext = (lSubDivision==-1) ? 0 : -1;
        goto Cleanup;
    }

    pTabs = new long[c];
    if(!pTabs)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = GetSubDivisionTabs(pTabs, c);
    if(hr)
    {
        goto Cleanup;
    }

    if(lSubDivision<0 && dir==DIRECTION_BACKWARD)
    {
        lSubDivision = c;
    }

    // Search for the next subdivision if possible to tab to it.
    for(i=(DIRECTION_FORWARD==dir)?lSubDivision+1:lSubDivision-1;
        (DIRECTION_FORWARD==dir)?i<c:i>=0;
        (DIRECTION_FORWARD==dir)?i++:i--)
    {
        if(pTabs[i]==0 || pTabs[i]==htmlTabIndexNotSet)
        {
            // Found something to tab to! Return it.  We're
            // only checking for zero here because negative
            // tab indices means cannot tab to them.  Positive
            // ones would already have been put in the focus array
            *plSubNext = i;
            goto Cleanup;
        }
    }

    // To reach here means that there are no further tabbable
    // subdivisions in this element.
    *plSubNext = -1;

Cleanup:
    delete[] pTabs;
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::IsEnabled
//
//  Synopsis:   Is this element enabled? Try to get the cascaded info cached
//              in the char format. If the element is not in any markup,
//              use the 'disabled attribute directly.
//
//  Returns:    BOOL
//
//----------------------------------------------------------------------------
BOOL CElement::IsEnabled()
{
    CTreeNode* pNode = GetFirstBranch();

    return !(pNode ? pNode->GetCharFormat()->_fDisabled : GetAAdisabled());
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::IsVisible
//
//  Synopsis:   Is this layout element visible?
//
//  Parameters: fCheckParent - check the parent first
//
//----------------------------------------------------------------------------
BOOL CElement::IsVisible(BOOL fCheckParent)
{
    CLayout*    pLayout         = GetUpdatedNearestLayout();

    if(!pLayout)
    {
        if(Tag() == ETAG_TXTSLAVE)
        {
            CElement* pElemMaster = MarkupMaster();

            if(pElemMaster)
            {
                return pElemMaster->IsVisible(fCheckParent);
            }
        }
        return FALSE;
    }

    CTreeNode*  pNode       = pLayout->GetFirstBranch();
    BOOL        fVisible    = !!pNode;

    if(pLayout->_fInvisibleAtRuntime)
    {
        fVisible = FALSE;
    }

    while(fVisible)
    {
        const CCharFormat* pCF;

        if(!pNode->Element()->IsInMarkup())
        {
            fVisible = FALSE;
            break;
        }

        if(!pNode->GetUpdatedLayout()->_fVisible)
        {
            fVisible = FALSE;
            break;
        }

        // BUGBUG This code was originally coded to use _this_ GetCharFormat but
        // do we want to use pNodeSite's instead? (jbeda)
        pCF = pNode->GetCharFormat();

        if(pCF->IsDisplayNone() || pCF->IsVisibilityHidden())
        {
            fVisible = FALSE;
            break;
        }

        if(!fCheckParent)
        {
            fVisible = TRUE;
            break;
        }

        pNode = pNode->GetUpdatedParentLayoutNode();

        if(!pNode || pNode->Element()->_fEditAtBrowse)
        {
            fVisible = TRUE;
            break;
        }
    }

    return fVisible;
}

//+---------------------------------------------------------------------------
//
//  Member: CElement::IsParent
//
//  Params: pElement: Check if pElement is the parent of this site
//
//  Descr:  Returns TRUE is pSite is a parent of this site, FALSE otherwise
//
//----------------------------------------------------------------------------
BOOL CElement::IsParent(CElement* pElement)
{
    CTreeNode* pNodeSiteTest = GetFirstBranch();

    while(pNodeSiteTest)
    {
        if(SameScope(pNodeSiteTest, pElement))
        {
            return TRUE;
        }
        pNodeSiteTest = pNodeSiteTest->GetUpdatedParentLayoutNode();
    }
    return FALSE;
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::OnContextMenu
//
//  Synopsis:   Handles WM_CONTEXTMENU message.
//
//--------------------------------------------------------------------------
HRESULT CElement::OnContextMenu(int x, int y, int id)
{
    int cx = x;
    int cy = y;
    CDocument* pDoc = Doc();


    if(cx==-1 && cy==-1)
    {
        RECT rcWin;

        GetWindowRect(pDoc->InPlace()->_hwnd, &rcWin);
        cx = rcWin.left;
        cy = rcWin.top;
    }

    {
        EVENTPARAM param(pDoc, TRUE);
        CTreeNode* pNode;

        pNode = (pDoc->_pElementOMCapture && pDoc->_pNodeLastMouseOver) ? pDoc->_pNodeLastMouseOver : GetFirstBranch();

        param.SetNodeAndCalcCoordinates(pNode);
        param.SetType(_T("MenuExtUnknown"));

        /*RRETURN1(pDoc->ShowContextMenu(cx, cy, id, this), S_FALSE); wlw note*/
        RRETURN(S_OK);
    }
}

//+---------------------------------------------------------------
//
//  Member:     CElement::OnMenuSelect
//
//  Synopsis:   Handle WM_MENUSELECT by updating status line text.
//
//----------------------------------------------------------------
HRESULT CElement::OnMenuSelect(UINT uItem, UINT fuFlags, HMENU hmenu)
{
    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     CElement::OnInitMenuPopup
//
//  Synopsis:   Handles WM_CONTEXTMENU message.
//
//---------------------------------------------------------------
BOOL IsMenuItemFontOrEncoding(ULONG idm)
{
    return (idm>=IDM_MIMECSET__FIRST__ && idm<=IDM_MIMECSET__LAST__)
        || ((idm>=IDM_BASELINEFONT1 && idm<=IDM_BASELINEFONT5)
        || (idm==IDM_DIRLTR || idm==IDM_DIRRTL));
}

HRESULT CElement::OnInitMenuPopup(HMENU hmenu, int item, BOOL fSystemMenu)
{
    int     i;
    MSOCMD  msocmd;
    UINT    mf;

    for(i=0; i<GetMenuItemCount(hmenu); i++)
    {
        msocmd.cmdID = GetMenuItemID(hmenu, i);
        if(msocmd.cmdID>0 && !IsMenuItemFontOrEncoding(msocmd.cmdID))
        {
            Doc()->QueryStatus((GUID*)&CGID_MSHTML, 1, &msocmd, NULL);
            switch(msocmd.cmdf)
            {
            case MSOCMDSTATE_UP:
            case MSOCMDSTATE_NINCHED:
                mf = MF_BYCOMMAND | MF_ENABLED | MF_UNCHECKED;
                break;

            case MSOCMDSTATE_DOWN:
                mf = MF_BYCOMMAND | MF_ENABLED | MF_CHECKED;
                break;

            case MSOCMDSTATE_DISABLED:
            default:
                mf = MF_BYCOMMAND | MF_DISABLED | MF_GRAYED;
                break;
            }
            CheckMenuItem(hmenu, msocmd.cmdID, mf);
            EnableMenuItem(hmenu, msocmd.cmdID, mf);
        }
    }
    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     CElement::PerformTA
//
//  Synopsis:   Forms implementation of TranslateAccelerator
//              Check against a list of accelerators for the incoming
//              message and if a match is found, fire the appropriate
//              command.  Return true if match found, false otherwise.
//
//  Input:      pMessage    Ptr to incoming message
//
//---------------------------------------------------------------
HRESULT CElement::PerformTA(CMessage* pMessage)
{
    HRESULT     hr = S_FALSE;
    DWORD       cmdID;
    MSOCMD      msocmd;
    CDocument*  pDoc = Doc();

    cmdID = GetCommandID(pMessage);

    if(cmdID == IDM_UNKNOWN)
    {
        goto Cleanup;
    }

    // CONSIDER: (anandra) Think about using an Exec
    // call directly here, instead of sendmessage.
    msocmd.cmdID = cmdID;
    msocmd.cmdf  = 0;

    hr = pDoc->QueryStatus((GUID*)&CGID_MSHTML, 1, &msocmd, NULL);
    if(hr)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    if(msocmd.cmdf == 0)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    if(msocmd.cmdf == MSOCMDSTATE_DISABLED)
    {
        goto Cleanup;
    }

    SendMessage(pDoc->_pInPlace->_hwnd, WM_COMMAND,
        GET_WM_COMMAND_MPS(cmdID, NULL, 1));
    hr = S_OK;

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CElement::GetNearestBgImgCtx
//
//--------------------------------------------------------------------------
CImgCtx* CElement::GetNearestBgImgCtx()
{
    CTreeNode*  pNodeSite;
    CImgCtx*    pImgCtx;

    for(pNodeSite=GetFirstBranch();
        pNodeSite;
        pNodeSite=pNodeSite->GetUpdatedParentLayoutNode())
    {
        pImgCtx = pNodeSite->Element()->GetBgImgCtx();
        if(pImgCtx && (pImgCtx->GetState()&IMGLOAD_COMPLETE))
        {
            return pImgCtx;
        }
    }

    return NULL;
}

//+---------------------------------------------------------------
//
//  Member:     CElement::GetCommandID
//
//---------------------------------------------------------------
DWORD CElement::GetCommandID(LPMSG lpmsg)
{
    DWORD   nCmdID;
    ACCELS* pAccels;

    pAccels = IsEditable() ? ElementDesc()->_pAccelsDesign : ElementDesc()->_pAccelsRun;

    nCmdID = pAccels ? pAccels->GetCommandID(lpmsg) : IDM_UNKNOWN;

    // IE5 73627 -- If the SelectAll accelerator is received by an element
    // that does not allow selection (possibly because it is in a dialog),
    // treat the accelerator as unknown instead of disabled. This allows the
    // accelerator to bubble up.
    if(nCmdID==IDM_SELECTALL && DisallowSelection())
    {
        nCmdID = IDM_UNKNOWN;
    }

    return nCmdID;
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::CLock::CLock
//
//  Synopsis:   Lock resources in CElement object.
//
//-------------------------------------------------------------------------
CElement::CLock::CLock(CElement* pElement, ELEMENTLOCK_FLAG enumLockFlags)
{
    Assert(enumLockFlags < (1<<(sizeof(_wLockFlags)*8)));
    _pElement = pElement;
    if(_pElement)
    {
        _wLockFlags = pElement->_wLockFlags;
        pElement->_wLockFlags |= (WORD)enumLockFlags;
        {
            pElement->PrivateAddRef();
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::CLock::~CLock
//
//  Synopsis:   Unlock resources in CElement object.
//
//-------------------------------------------------------------------------
CElement::CLock::~CLock()
{
    if(_pElement)
    {
        _pElement->_wLockFlags = _wLockFlags;
        {
            _pElement->PrivateRelease();
        }
    }
}

BOOL CElement::WantTextChangeNotifications()
{
    if(!HasLayout())
    {
        return FALSE;
    }

    CFlowLayout* pFlowLayout = GetCurLayout()->IsFlowLayout();
    if(pFlowLayout && pFlowLayout->IsListening())
    {
        return TRUE;
    }

    return FALSE;
}

CElement::ACCELS::ACCELS(ACCELS* pSuper, WORD wAccels)
{
    _pSuper = pSuper;
    _wAccels = wAccels;
    _fResourcesLoaded = FALSE;
    _pAccels = NULL;
}

CElement::ACCELS::~ACCELS()
{
    if(_pAccels)
    {
        delete[] _pAccels;
        _pAccels = NULL;
    }
}

HRESULT CElement::ACCELS::LoadAccelTable()
{
    HRESULT hr = S_OK;
    HACCEL  hAccel;
    int     cLoaded;

    if(!_wAccels)
    {
        goto Cleanup;
    }

    hAccel = LoadAccelerators(_afxGlobalData._hInstResource, MAKEINTRESOURCE(_wAccels));
    if(!hAccel)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    _cAccels = CopyAcceleratorTable(hAccel, NULL, 0);
    Assert(_cAccels);

    _pAccels = new ACCEL[_cAccels];
    if(!_pAccels)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    cLoaded = CopyAcceleratorTable(hAccel, _pAccels, _cAccels);
    DestroyAcceleratorTable(hAccel);
    if(cLoaded != _cAccels)
    {
        hr = E_OUTOFMEMORY;
    }

Cleanup:
    if(hr)
    {
        _cAccels = 0;
        _pAccels = NULL;
    }
    RRETURN (hr);
}

HRESULT CElement::ACCELS::EnsureResources()
{
    if(!_fResourcesLoaded)
    {
        CGlobalLock glock;

        if(!_fResourcesLoaded)
        {
            HRESULT hr;
            hr = LoadAccelTable();
            _fResourcesLoaded = TRUE;
            RRETURN(hr);
        }
    }

    return S_OK;
}

DWORD CElement::ACCELS::GetCommandID(LPMSG pmsg)
{
    HRESULT hr;
    DWORD   nCmdID = IDM_UNKNOWN;
    WORD    wVKey;
    ACCEL*  pAccel;
    int     i;
    DWORD   dwKeyState = FormsGetKeyState();

    if(WM_KEYDOWN!=pmsg->message && WM_SYSKEYDOWN!=pmsg->message)
    {
        goto Cleanup;
    }

    if(_pSuper)
    {
        nCmdID = _pSuper->GetCommandID(pmsg);
        if(IDM_UNKNOWN != nCmdID) // found id, nothing more to do
        {
            goto Cleanup;
        }
    }

    hr = EnsureResources();
    if(hr)
    {
        goto Cleanup;
    }

    // loop through the table
    for(i=0,pAccel=_pAccels; i<_cAccels; i++,pAccel++)
    {
        if(!(pAccel->fVirt & FVIRTKEY))
        {
            wVKey = LOBYTE(VkKeyScan(pAccel->key));
        }
        else
        {
            wVKey = pAccel->key;
        }

        if(wVKey==pmsg->wParam &&
            EQUAL_BOOL(pAccel->fVirt&FCONTROL, dwKeyState&MK_CONTROL) &&
            EQUAL_BOOL(pAccel->fVirt&FSHIFT,   dwKeyState&MK_SHIFT) &&
            EQUAL_BOOL(pAccel->fVirt&FALT,     dwKeyState&MK_ALT))
        {
            nCmdID = pAccel->cmd;
            break;
        }
    }

Cleanup:
    return nCmdID;
}

CElement::ACCELS CElement::s_AccelsElementDesign = CElement::ACCELS(NULL, 0);
CElement::ACCELS CElement::s_AccelsElementRun    = CElement::ACCELS(NULL, 0);

void* CElement::GetLookasidePtr(int iPtr)
{
    return (HasLookasidePtr(iPtr) ? Doc()->GetLookasidePtr((DWORD*)this+iPtr) : NULL);
}

HRESULT CElement::SetLookasidePtr(int iPtr, void * pvVal)
{
    Assert(!HasLookasidePtr(iPtr) && "Can't set lookaside ptr when the previous ptr is not cleared");

    HRESULT hr = Doc()->SetLookasidePtr((DWORD*)this+iPtr, pvVal);

    if(hr == S_OK)
    {
        _fHasLookasidePtr |= 1 << iPtr;
    }

    RRETURN(hr);
}

void* CElement::DelLookasidePtr(int iPtr)
{
    if(HasLookasidePtr(iPtr))
    {
        void* pvVal = Doc()->DelLookasidePtr((DWORD*)this+iPtr);
        _fHasLookasidePtr &= ~(1 << iPtr);
        return(pvVal);
    }

    return (NULL);
}

CLayout* CElement::GetLayoutPtr() const
{
    return HasLayoutPtr()?(CLayout*)__pvChain:NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::SetLayoutPtr
//
//  Synopsis:   Set the layout pointer
//
//----------------------------------------------------------------------------
void CElement::SetLayoutPtr(CLayout* pLayout)
{
    Assert(!HasLayoutPtr());
    Assert(!pLayout->_fHasMarkupPtr);

    pLayout->__pvChain = __pvChain;
    __pvChain = pLayout;
    pLayout->_fHasMarkupPtr = _fHasMarkupPtr;
    _fHasLayoutPtr = TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::DelLayoutPtr
//
//  Synopsis:   Remove the layout pointer
//
//----------------------------------------------------------------------------
CLayout* CElement::DelLayoutPtr()
{
    Assert(HasLayoutPtr());

    CLayout* pLayoutRet = (CLayout*)__pvChain;

    __pvChain = pLayoutRet->__pvChain;
    _fHasLayoutPtr = FALSE;

    pLayoutRet->__pvChain = HasMarkupPtr() ? ((CMarkup*)__pvChain)->Doc() : (CDocument*)__pvChain;

    pLayoutRet->_fHasMarkupPtr = FALSE;

    return pLayoutRet;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetMarkupPtr
//
//  Synopsis:   Get the markup pointer if any
//
//----------------------------------------------------------------------------
CMarkup* CElement::GetMarkupPtr() const
{
    void* pv = __pvChain;

    if(HasLayoutPtr())
    {
        pv = ((CLayout*)pv)->__pvChain;
    }

    if(HasMarkupPtr())
    {
        return (CMarkup*)pv;
    }

    return NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::SetMarkupPtr
//
//  Synopsis:   Set the markup pointer
//
//----------------------------------------------------------------------------
void CElement::SetMarkupPtr(CMarkup* pMarkup)
{
    Assert(!HasMarkupPtr());
    Assert(pMarkup);

    if(HasLayoutPtr())
    {
        ((CLayout*)__pvChain)->__pvChain = pMarkup;
        ((CLayout*)__pvChain)->_fHasMarkupPtr = TRUE;
    }
    else
    {
        __pvChain = pMarkup;
    }

    _fHasMarkupPtr = TRUE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::DelMarkupPtr
//
//  Synopsis:   Remove the markup pointer
//
//----------------------------------------------------------------------------
void CElement::DelMarkupPtr()
{
    Assert(HasMarkupPtr());

    if(HasLayoutPtr())
    {
        ((CLayout*)__pvChain)->__pvChain = ((CMarkup*)(((CLayout*)__pvChain)->__pvChain))->Doc();
        ((CLayout*)__pvChain)->_fHasMarkupPtr = FALSE;
    }
    else
    {
        __pvChain = ((CMarkup*)__pvChain)->Doc();
    }

    _fHasMarkupPtr = FALSE;
}

BOOL CElement::IsInPrimaryMarkup() const
{
    CMarkup* pMarkup = GetMarkup();

    return pMarkup?pMarkup->IsPrimaryMarkup():FALSE;
}

BOOL CElement::IsInThisMarkup(CMarkup* pMarkupIn) const
{
    CMarkup* pMarkup = GetMarkup();

    return pMarkup==pMarkupIn;
}

CRootElement* CElement::MarkupRoot()
{
    CMarkup* pMarkup = GetMarkup();

    return (pMarkup?pMarkup->Root():NULL);
}

CElement* CElement::MarkupMaster() const
{
    CMarkup* pMarkup = GetMarkup();

    return pMarkup?pMarkup->Master():NULL;
}

CElement* CElement::FireEventWith()
{
    CMarkup* pMarkup = GetMarkup();

    if(pMarkup && pMarkup->FirstElement()==this && pMarkup->Master())
    {
        return pMarkup->Master();
    }
    return this;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetDocPtr
//
//  Synopsis:   Get the CDocument pointer
//
//----------------------------------------------------------------------------
CDocument* CElement::GetDocPtr() const
{
    void* pv = __pvChain;

    if(HasLayoutPtr())
    {
        pv = ((CLayout*)pv)->__pvChain;
    }
    if(HasMarkupPtr())
    {
        pv = ((CMarkup*)pv)->Doc();
    }

    return (CDocument*)pv;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::GetReadyState
//
//----------------------------------------------------------------------------
long CElement::GetReadyState()
{
    /*CPeerMgr* pPeerMgr = GetPeerMgr();

    if(pPeerMgr)
    {
        return pPeerMgr->_readyState;
    }
    else wlw note*/
    {
        return READYSTATE_COMPLETE;
    } 
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::OnReadyStateChange
//
//----------------------------------------------------------------------------
void CElement::OnReadyStateChange()
{
    Fire_onreadystatechange();
}

//+----------------------------------------------------------------------------
//
//  member  :   click()   IHTMLElement method
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CElement::click(CTreeNode* pNodeContext)
{
    HRESULT hr = DoClick(NULL, pNodeContext);
    if(hr == S_FALSE)
    {
        hr = S_OK;
    }
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::releaseCapture()
{
    HRESULT hr = S_OK;
    CDocument* pDoc = Doc();

    if(!pDoc)
    {
        goto Cleanup;
    }

    if(pDoc->_pElementOMCapture == this)
    {
        hr = pDoc->releaseCapture();
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT GetAll(CElement* pel, IDispatch** ppDispChildren)
{
    HRESULT hr = S_OK;
    CMarkup* pMarkup;

    if (!ppDispChildren)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    Assert(pel);
    *ppDispChildren = NULL;

    hr = pel->EnsureInMarkup();
    if(hr)
    {
        goto Cleanup;
    }

    pMarkup = pel->GetMarkupPtr();

    hr = pMarkup->InitCollections();
    if(hr)
    {
        goto Cleanup;
    }

    hr = pMarkup->CollectionCache()->CreateChildrenCollection(CMarkup::ELEMENT_COLLECTION, pel, ppDispChildren, TRUE);

Cleanup:
    RRETURN(hr);
}

STDMETHODIMP CElement::get_all(IDispatch** ppDispChildren)
{
    HRESULT hr;
    hr = GetAll(this, ppDispChildren);
    RRETURN(SetErrorInfoPGet(hr, DISPID_CElement_all));
}

STDMETHODIMP CElement::get_children(IDispatch** ppDispChildren)
{
    HRESULT hr = S_OK;
    CMarkup* pMarkup;

    if(!ppDispChildren)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ppDispChildren = NULL;

    hr = EnsureInMarkup();
    if(hr)
    {
        goto Cleanup;
    }

    pMarkup = GetMarkupPtr();

    hr = pMarkup->InitCollections();
    if(hr)
    {
        goto Cleanup;
    }

    hr = pMarkup->CollectionCache()->CreateChildrenCollection(CMarkup::ELEMENT_COLLECTION, this, ppDispChildren, FALSE);

Cleanup:
    RRETURN(SetErrorInfoPGet(hr, DISPID_CElement_children));
}

HRESULT CElement::toString(BSTR* String)
{
    RRETURN(super::toString(String));
};

static CElement::Where ConvertAdjacent(htmlAdjacency where)
{
    switch(where)
    {
    case htmlAdjacencyBeforeBegin : return CElement::BeforeBegin;
    case htmlAdjacencyAfterBegin  : return CElement::AfterBegin;
    case htmlAdjacencyBeforeEnd   : return CElement::BeforeEnd;
    case htmlAdjacencyAfterEnd    : return CElement::AfterEnd;
    default                       : Assert(0);
    }

    return CElement::BeforeBegin;
}

//+----------------------------------------------------------------------------
//
//  Member:     put_outerText
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::put_outerText(BSTR bstrText)
{
    HRESULT hr = S_OK;

    hr = Inject(Outside, FALSE, bstrText, FormsStringLen(bstrText));

    if(hr)
    {
        goto Cleanup;
    }

    hr = OnPropertyChange(s_propdescCElementouterText.b.dispid,
        s_propdescCElementouterText.b.dwFlags);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfoPSet(hr, DISPID_CElement_outerText));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_outerText
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::get_outerText(BSTR* pbstrText)
{
    HRESULT hr = GetText(pbstrText, WBF_SAVE_PLAINTEXT|WBF_NO_WRAP);

    RRETURN(SetErrorInfoPGet(hr, DISPID_CElement_outerText));
}

//+----------------------------------------------------------------------------
//
//  Member:     put_innerText
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::put_innerText(BSTR bstrText)
{
    HRESULT hr = S_OK;

    hr = Inject(Inside, FALSE, bstrText, FormsStringLen(bstrText));

    if(hr)
    {
        goto Cleanup;
    }

    hr = OnPropertyChange(s_propdescCElementinnerText.b.dispid,
        s_propdescCElementinnerText.b.dwFlags);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfoPSet(hr, DISPID_CElement_innerText));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_innerText
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::get_innerText(BSTR* pbstrText)
{
    HRESULT hr = GetText(pbstrText, WBF_SAVE_PLAINTEXT|WBF_NO_WRAP|WBF_NO_TAG_FOR_CONTEXT);

    RRETURN(SetErrorInfoPGet(hr, DISPID_CElement_innerText));
}

//+----------------------------------------------------
//
//  member : get_offsetParent, IHTMLElement
//
//  synopsis : returns the parent container (site) which
//      defines the coordinate system of the offset*
//      properties above.
//              returns NULL when no parent makes sense
//
//-----------------------------------------------------
HRESULT CElement::get_offsetParent(IHTMLElement** ppIElement)
{
    HRESULT  hr = S_OK;
    CTreeNode* pNodeContext = GetFirstBranch();
    CTreeNode* pNodeRet = NULL;

    if(!ppIElement)
    {
        hr = E_POINTER;
        goto Cleanup;
    }
    else if(!GetFirstBranch())
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    *ppIElement= NULL;

    switch(_etag)
    {

    case ETAG_BODY:
        break;

    case ETAG_AREA:
        // the area's parent is the map
        if(pNodeContext->Parent())
        {
            pNodeRet = pNodeContext->Parent();
        }
        break;

    case ETAG_TD:
    case ETAG_TH:
        if(pNodeContext->GetUpdatedNearestLayout())
        {
            pNodeRet = NeedsLayout() ? pNodeContext->ZParentBranch()
                : pNodeContext->GetUpdatedNearestLayoutNode();

            if(pNodeRet && pNodeRet->IsPositionStatic())
            {
                // return pNodeRet's parent
                pNodeRet = pNodeRet->ZParentBranch();
                Assert(pNodeRet && pNodeRet->Tag()==ETAG_TABLE);
            }
        }
        break;

    default:
        if(pNodeContext->GetUpdatedNearestLayout())
        {
            // If this changes, make sure to update GetElementTopLeft()
            pNodeRet = NeedsLayout() ? pNodeContext->ZParentBranch()
                : pNodeContext->GetUpdatedNearestLayoutNode();

        }
        break;
    }


    if(pNodeRet && pNodeRet->Tag()!=ETAG_ROOT)
    {
        hr = pNodeRet->GetElementInterface(IID_IHTMLElement, (void**)ppIElement);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+------------------------------------------------------------
//
//  member  :   get_offsetWidth,  IHTMLElement
//
//  Synopsis:   Get the calculated height in doc units. if *this*
//          is a site then just return based on the size
//          if it is an element, then we need to get the regions
//          of its parts and add it up.
//
//-------------------------------------------------------------
HRESULT CElement::get_offsetWidth(long* plValue)
{
    HRESULT hr = S_OK;
    SIZE size;

    if(!plValue)
    {
        hr = E_POINTER;
    }
    else if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        GetBoundingSize(size);

        *plValue = size.cx;
    }
    else
    {
        *plValue = 0;
    }

    RRETURN(SetErrorInfo(hr));
}

//+------------------------------------------------------------
//
//  member  :   get_offsetHeight,  IHTMLElement
//
//  Synopsis:   Get the calculated height in doc units. if *this*
//          is a site then just return based on the size
//          if it is an element, then we need to get the regions
//          of its parts and add it up.
//
//-------------------------------------------------------------
HRESULT CElement::get_offsetHeight(long* plValue)
{
    HRESULT hr = S_OK;
    SIZE size;

    if(!plValue)
    {
        hr = E_POINTER;
    }
    else if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        GetBoundingSize(size);

        *plValue = size.cy;
    }
    else
    {
        *plValue = 0;
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------
//
//  member : get_offsetTop, IHTMLElement
//
//  synopsis : returns the top, coordinate of the
//      element
//
//-----------------------------------------------------
HRESULT CElement::get_offsetTop(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_POINTER;
    }
    else if(!IsInMarkup() || !Doc()->GetView()->IsActive())
    {
        *plValue = 0;
    }
    else
    {
        *plValue = 0;

        switch(_etag)
        {
        case ETAG_AREA:
            {
                AssertSz(FALSE, "must improve");
            }
            break;

        default:
            {
                POINT pt;
                SendNotification(NTYPE_ELEMENT_ENSURERECALC);
                hr = GetElementTopLeft(pt);
                *plValue = pt.y;
            }
            break;
        }
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------
//
//  member : get_OffsetLeft, IHTMLElement
//
//  synopsis : returns the left coordinate of the
//      element
//
//-----------------------------------------------------
HRESULT CElement::get_offsetLeft(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_POINTER;
    }
    else if(!IsInMarkup() || !Doc()->GetView()->IsActive())
    {
        *plValue = 0;
    }
    else
    {
        // (paulnel) Are we right to left? If so, we need to convert the offset left
        // amount, that will be a negative value from the parent's top right, to a 
        // positive offset from the parent's top left. To do this we will
        // use the formula "left = parent.width + left"
        CTreeNode* pNodeContext = GetFirstBranch();
        CTreeNode* pNodeParent = NULL;
        long cxParent = 0;
        BOOL fRTL = FALSE;
        if(pNodeContext->GetUpdatedNearestLayout())
        {
            pNodeParent = (NeedsLayout()) ? pNodeContext->ZParentBranch()
                : pNodeContext->GetUpdatedNearestLayoutNode();

            fRTL = (pNodeParent!=NULL && pNodeParent->GetCharFormat()->_fRTL);

            if(fRTL)
            {
                CLayout* pLayoutParent = pNodeParent->GetUpdatedNearestLayout();

                // if we have scroll bars, we need to take of the size of the scrolls.
                Assert(pLayoutParent);
                if(pLayoutParent && pLayoutParent->GetElementDispNode() &&
                    pLayoutParent->GetElementDispNode()->IsScroller())
                {
                    CRect rcClient;
                    pLayoutParent->GetClientRect(&rcClient);
                    cxParent = rcClient.Width();
                }
                else
                {
                    CElement* pElemParent;
                    CSize sizeParent;

                    pElemParent = pNodeParent->Element();
                    Assert(pElemParent && pElemParent != this);
                    if(pElemParent)
                    {
                        pElemParent->GetBoundingSize(sizeParent);
                        cxParent = sizeParent.cx;
                    }
                }
            }
        }

        *plValue = 0;
        switch(_etag)
        {
        case ETAG_AREA:
            {
                AssertSz(FALSE, "must improve");
            }
            break;

        default :
            {
                POINT pt;
                SendNotification(NTYPE_ELEMENT_ENSURERECALC);
                hr = GetElementTopLeft(pt);

                // convert for Right To Left elements
                if(fRTL)
                {
                    pt.x += cxParent;
                }

                *plValue = pt.x;
            }
            break;
        }

    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     get_sourceIndex
//
//  Synopsis:   Returns the source index (order of appearance) of this element
//              If the element is no longer in the source tree, return -1
//              as source index and hr = S_OK.
//
//-----------------------------------------------------------------------------
STDMETHODIMP CElement::get_sourceIndex(long* pSourceIndex)
{
    HRESULT hr = S_OK;

    if(!pSourceIndex)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pSourceIndex = GetSourceIndex();

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     contains
//
//  Synopsis:   IHTMLElement method. returns a boolean  if PIelement is within
//              the scope of this
//
//-----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CElement::contains(IHTMLElement* pIElement, VARIANT_BOOL* pfResult)
{
    CTreeNode* pNode = NULL;
    HRESULT hr = S_OK;

    Assert(pfResult);

    if(!pfResult)
    {
        goto Cleanup;
    }

    *pfResult = VB_FALSE;
    if(!pIElement)
    {
        goto Cleanup;
    }

    // get a CTreeNode pointer
    hr = pIElement->QueryInterface(CLSID_CTreeNode, (void**)&pNode);
    if(hr == E_NOINTERFACE)
    {
        CElement* pElement;
        hr =pIElement->QueryInterface(CLSID_CElement, (void**)&pElement);
        if(hr)
        {
            goto Cleanup;
        }

        pNode = pElement->GetFirstBranch();
    }
    else if(hr)
    {
        goto Cleanup;
    }

    while(pNode && DifferentScope(pNode, this))
    {
        // stop after the HTML tag
        if(pNode->Tag() == ETAG_ROOT)
        {
            pNode = NULL;
        }
        else
        {
            pNode = pNode->Parent();
        }
    }

    if(pNode)
    {
        *pfResult = VB_TRUE;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT STDMETHODCALLTYPE CElement::scrollIntoView(VARIANTARG varargStart)
{
    HRESULT     hr;
    SCROLLPIN   spVert;
    BOOL        fStart;
    CVariant    varBOOLStart;

    hr = varBOOLStart.CoerceVariantArg(&varargStart, VT_BOOL);
    if(hr == S_OK)
    {
        fStart = V_BOOL(&varBOOLStart);
    }
    else if(hr == S_FALSE)
    {
        // when no argument
        fStart = TRUE;
    }
    else
    {
        goto Cleanup;
    }

    spVert = fStart ? SP_TOPLEFT : SP_BOTTOMRIGHT;

    hr = ScrollIntoView(spVert, SP_TOPLEFT);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::get_onhelp(VARIANT* p)
{
    HRESULT hr;

    // only delegate if this is body
    if(Tag() == ETAG_BODY)
    {
        // wlw note
        //CDocument* pDoc = Doc();
        //hr = pDoc->EnsureOmWindow();
        //if(hr)
        //{
        //    goto Cleanup;
        //}

        //hr = pDoc->_pOmWindow->get_onhelp(p);
    }
    else
    {
        hr = s_propdescCElementonhelp.a.HandleCodeProperty(
            HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            p,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::put_onhelp(VARIANT v)
{
    HRESULT hr;

    // only delegate if this is body
    if(Tag() == ETAG_BODY)
    {
        // wlw note
        //CDocument* pDoc = Doc();
        //hr = pDoc->EnsureOmWindow();
        //if(hr)
        //{
        //    goto Cleanup;
        //}

        //hr = pDoc->_pOmWindow->put_onhelp(v);
    }
    else
    {
        hr = s_propdescCElementonhelp.a.HandleCodeProperty(
            HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

void CElement::InvalidateCollection(long lIndex)
{
    CMarkup* pMarkup;
    CCollectionCache* pCollCache;

    pMarkup = GetMarkup();
    if(pMarkup)
    {
        pCollCache = pMarkup->CollectionCache();
        if(pCollCache)
        {
            pCollCache->InvalidateItem(lIndex);
        }
    }
}

// Replace CBase::removeAttribute, need to special case some DISPID's
HRESULT CElement::removeAttribute(BSTR strPropertyName, LONG lFlags, VARIANT_BOOL* pfSuccess)
{
    DISPID dispID;
    IDispatchEx* pDEX = NULL;

    if(pfSuccess)
    {
        *pfSuccess = VB_FALSE;
    }

    // BUGBUG rgardner should move the STDPROPID_XOBJ_STYLE
    // code from CBase::removeAttribute to here
    if(PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEX))
    {
        goto Cleanup;
    }

    if(pDEX->GetDispID(strPropertyName,
        lFlags&GETMEMBER_CASE_SENSITIVE?fdexNameCaseSensitive:0,
        &dispID))
    {
        goto Cleanup;
    }

    if(!removeAttributeDispid(dispID))
    {
        goto Cleanup;
    }

    // If we remove name or ID, update the WINDOW_COLLECTION if needed
    // Don't need to deal with uniqueName here - it's not exposed externaly
    if(dispID==DISPID_CElement_id || dispID==STDPROPID_XOBJ_NAME)
    {
        LPCTSTR lpszNameID = NULL;
        BOOL fNamed;

        // Named if ID or name is present
        if(_pAA && _pAA->HasAnyAttribute())
        {
            _pAA->FindString(
                dispID==DISPID_CElement_id?STDPROPID_XOBJ_NAME:DISPID_CElement_id, 
                &lpszNameID);
            if(!(lpszNameID && *lpszNameID))
            {
                _pAA->FindString(DISPID_CElement_uniqueName, &lpszNameID);
            }
        }

        fNamed = lpszNameID && *lpszNameID;

        if(fNamed != !!_fIsNamed)
        {
            // Inval all collections affected by a name change
            DoElementNameChangeCollections();
            _fIsNamed = fNamed;
        }
    }

    if(pfSuccess)
    {
        *pfSuccess = VB_TRUE;
    }

Cleanup:
    ReleaseInterface(pDEX);

    RRETURN(SetErrorInfo(S_OK));
}

HRESULT CElement::SetUniqueNameHelper(LPCTSTR szUniqueName)
{
    RRETURN(SetIdentifierHelper(szUniqueName, DISPID_CElement_uniqueName,
        DISPID_CElement_id, STDPROPID_XOBJ_NAME));
}

HRESULT CElement::getElementsByTagName(BSTR v, IHTMLElementCollection** ppDisp)
{
    HRESULT hr = E_INVALIDARG;
    IDispatch* pDispChildren = NULL;
    CElementCollection* pelColl = NULL;

    if(!ppDisp || !v)
    {
        goto Cleanup;
    }

    *ppDisp = NULL;

    hr = GetAll(this, &pDispChildren);
    if(hr)
    {
        goto Cleanup;
    }

    Assert(pDispChildren);
    hr = pDispChildren->QueryInterface(CLSID_CElementCollection, (void**)&pelColl);
    if(hr)
    {
        goto Cleanup;
    }

    Assert(pelColl);

    // Get a collection of the specified tags.
    hr = pelColl->Tags(v, (IDispatch**)ppDisp);

Cleanup:
    ReleaseInterface(pDispChildren);
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::get_readyStateValue(long* plRetValue)
{
    HRESULT hr = S_OK;

    if(!plRetValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plRetValue = READYSTATE_COMPLETE;

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Member:     CElement:get_readyState
//
//+------------------------------------------------------------------------------
HRESULT CElement::get_readyState(VARIANT* pVarRes)
{
    HRESULT hr = S_OK;

    if(!pVarRes)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    /*Doc()->PeerDequeueTasks(0); wlw note*/

    hr = s_enumdeschtmlReadyState.StringFromEnum(GetReadyState(), &V_BSTR(pVarRes));
    if(!hr)
    {
        V_VT(pVarRes) = VT_BSTR;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::get_canHaveChildren(VARIANT_BOOL* pvb)
{
    *pvb = IsNoScope() ? VARIANT_FALSE : VARIANT_TRUE;
    return S_OK;
}

static HRESULT SetAdjacentTextPointer(
      CElement*             pElem,
      htmlAdjacency         where,
      MARKUP_CONTEXT_TYPE*  pContext,
      CMarkupPointer*       pmkptrStart,
      long*                 plCharCount)
{
    ELEMENT_ADJACENCY adj;
    BOOL fLeft;
    HRESULT hr;

    switch(where)
    {
    case htmlAdjacencyBeforeBegin:
    default:
        adj = ELEM_ADJ_BeforeBegin;
        fLeft = TRUE;
        break;
    case htmlAdjacencyAfterBegin:
        adj = ELEM_ADJ_AfterBegin;
        fLeft = FALSE;
        break;
    case htmlAdjacencyBeforeEnd:
        adj = ELEM_ADJ_BeforeEnd;
        fLeft = TRUE;
        break;
    case htmlAdjacencyAfterEnd:
        adj = ELEM_ADJ_AfterEnd;
        fLeft = FALSE;
        break;
    }

    hr = pmkptrStart->MoveAdjacentToElement(pElem, adj);
    if(hr)
    {
        goto Cleanup;
    }

    if(fLeft)
    {
        // Need to move the pointer to the start of the text
        hr = pmkptrStart->Left(TRUE, pContext, NULL, plCharCount, NULL, NULL);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else if(plCharCount)
    {
        // Need to do a non-moving Right to get the text length
        hr = pmkptrStart->Right(FALSE, pContext, NULL, plCharCount, NULL, NULL);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------------
//
//  Member:     clearAttributes
//
//+-------------------------------------------------------------------------------
HRESULT CElement::clearAttributes()
{
    HRESULT hr = S_OK;
    CAttrArray* pAA = *(GetAttrArray());
    // wlw note
    //CMergeAttributesUndo Undo(this);
    //Undo.SetClearAttr(TRUE);
    //Undo.SetWasNamed(_fIsNamed);

    if(pAA)
    {
        //pAA->Clear(Undo.GetAA());

        hr = OnPropertyChange(DISPID_UNKNOWN, ELEMCHNG_REMEASUREINPARENT|ELEMCHNG_CLEARCACHES|ELEMCHNG_REMEASUREALLCONTENTS);
        if(hr)
        {
            goto Cleanup;
        }
    }

    //Undo.CreateAndSubmit();

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+------------------------------------------------------------------
//
//  Members: [get/put]_scroll[top/left] and get_scroll[height/width]
//
//  Synopsis : CElement members. _dp is in pixels.
//
//------------------------------------------------------------------
HRESULT CElement::get_scrollHeight(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plValue = 0;

    if(IsInMarkup())
    {
        CLayout* pLayout;

        // make sure that current is calced
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();

        if(pLayout)
        {
            *plValue = pLayout->GetContentHeight();
        }

        // we don't want to return a zero for the Height (only happens
        // when there is no content). so default to the offsetHeight
        if(!pLayout || *plValue==0)
        {
            // return the offsetWidth instead
            hr = get_offsetHeight(plValue);
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::get_scrollWidth(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plValue = 0;

    if(IsInMarkup())
    {
        CLayout* pLayout;

        // make sure that current is calced
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();

        if(pLayout)
        {
            *plValue = pLayout->GetContentWidth();
        }

        // we don't want to return a zero for teh width (only haoppens
        // when there is no content). so default to the offsetWidth
        if(!pLayout || *plValue==0)
        {
            // return the offsetWidth instead
            hr = get_offsetWidth(plValue);
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::get_scrollTop(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plValue = 0;

    if(IsInMarkup())
    {
        CLayout* pLayout;
        CDispNode* pDispNode;

        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        if((pLayout=GetUpdatedLayout())!=NULL &&
            (pDispNode=pLayout->GetElementDispNode())!=NULL)
        {
            if(pDispNode->IsScroller())
            {
                CSize sizeOffset;
                DYNCAST(CDispScroller, pDispNode)->GetScrollOffset(&sizeOffset);
                *plValue = sizeOffset.cy;
            }
            else
            {
                // if this isn't a scrolling element, then the scrollTop must be 0
                // for IE4 compatability
                *plValue = 0;
            }
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::get_scrollLeft(long* plValue)
{
    HRESULT hr = S_OK;

    if(!plValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plValue = 0;

    if(IsInMarkup())
    {
        CLayout* pLayout;
        CDispNode* pDispNode;

        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        if((pLayout = GetUpdatedLayout())!=NULL &&
            (pDispNode = pLayout->GetElementDispNode())!=NULL)
        {
            if(pDispNode->IsScroller())
            {
                CSize sizeOffset;
                DYNCAST(CDispScroller, pDispNode)->GetScrollOffset(&sizeOffset);
                *plValue = sizeOffset.cx;
            }
            else
            {
                // if this isn't a scrolling element, then the scrollTop must be 0
                // for IE$ compatability
                *plValue = 0;
            }
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::put_scrollTop(long lPixels)
{
    if(!GetFirstBranch())
    {
        RRETURN(E_FAIL);
    }
    else
    {
        CLayout* pLayout;
        CDispNode* pDispNode;

        // make sure that the element is calc'd
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();
        pDispNode = pLayout ? pLayout->GetElementDispNode() : NULL;

        if(pDispNode && pDispNode->IsScroller())
        {
            // the display tree uses negative numbers to indicate nochange,
            // but the OM uses negative nubers to mean scrollto the top.
            pLayout->ScrollToY((lPixels <0)?0:lPixels);
        }
    }

    return S_OK;
}

HRESULT CElement::put_scrollLeft(long lPixels)
{
    if(!GetFirstBranch())
    {
        RRETURN(E_FAIL);
    }
    else
    {
        CLayout* pLayout;
        CDispNode* pDispNode;

        // make sure that the element is calc'd
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();
        pDispNode = pLayout ? pLayout->GetElementDispNode() : NULL;

        if(pDispNode && pDispNode->IsScroller())
        {
            pLayout->ScrollToX((lPixels<0)?0:lPixels);
        }
    }

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CElement::get_dir
//
//  Synopsis: Object model entry point to get the dir property.
//
//  Arguments:
//            [p]  - where to return BSTR containing the string
//
//  Returns:  S_OK                  - this element supports this property
//            DISP_E_MEMBERNOTFOUND - this element doesn't support this
//                                    property
//            E_OUTOFMEMORY         - memory allocation failure
//            E_POINTER             - NULL pointer to receive BSTR
//
//----------------------------------------------------------------------------
STDMETHODIMP CElement::get_dir(BSTR* p)
{
    HRESULT hr;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    if(p)
    {
        *p = NULL;
    }

    hr = s_propdescCElementdir.b.GetEnumStringProperty(p, this, (CVoid*)(void*)(GetAttrArray()));

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------
//
//  member  :   get_clientWidth, IHTMLControlElement
//
//  synopsis    :   returns a long value of the client window
//      width (not counting scrollbar, borders..)
//
//-----------------------------------------------------------
HRESULT CElement::get_clientWidth(long* pl)
{
    HRESULT hr = S_OK;

    if(!pl)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pl = 0;

    if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        CLayout* pLayout = GetUpdatedLayout();

        if(pLayout)
        {
            RECT rect;

            // TR's are strange beasts since they have layout but
            // no display (unless they are positioned....
            if(Tag() == ETAG_TR)
            {
                CDispNode* pDispNode = pLayout->GetElementDispNode();
                if(!pDispNode)
                {
                    // we don't have a display so GetClientRect will return 0
                    // so rather than this, default to the offsetWidth.  This
                    // is the same behavior as scrollWidth
                    hr = get_offsetWidth(pl);
                    goto Cleanup;
                }
            }

            pLayout->GetClientRect(&rect, CLIENTRECT_VISIBLECONTENT);

            *pl = rect.right - rect.left;
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------
//
//  member  :   get_clientHeight, IHTMLControlElement
//
//  synopsis    :   returns a long value of the client window
//      Height of the body
//
//-----------------------------------------------------------
HRESULT CElement::get_clientHeight(long* pl)
{
    HRESULT hr = S_OK;

    if(!pl)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pl = 0;

    if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        CLayout* pLayout = GetUpdatedLayout();

        if(pLayout)
        {
            RECT rect;

            // TR's are strange beasts since they have layout but
            // no display (unless they are positioned....
            if(Tag() == ETAG_TR)
            {
                CDispNode* pDispNode = pLayout->GetElementDispNode();
                if(!pDispNode)
                {
                    // we don't have a display so GetClientRect will return 0
                    // so rather than this, default to the offsetHeight. This
                    // is the same behavior as scrollHeight
                    hr = get_offsetHeight(pl);
                    goto Cleanup;
                }
            }

            pLayout->GetClientRect(&rect, CLIENTRECT_VISIBLECONTENT);

            *pl = rect.bottom - rect.top;
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------
//
//  member  :   get_clientTop, IHTMLControlElement
//
//  synopsis    :   returns a long value of the client window
//      Top (inside borders)
//
//-----------------------------------------------------------
HRESULT CElement::get_clientTop(long* pl)
{
    HRESULT hr = S_OK;

    if(!pl)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pl = 0;

    if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        CLayout* pLayout;

        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();
        if(pLayout)
        {
            CDispNode* pDispNode= pLayout->GetElementDispNode();

            if(pDispNode)
            {
                CRect rcBorders;
                pDispNode->GetBorderWidths(&rcBorders);
                *pl = rcBorders.top;
            }
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------
//
//  member  :   get_clientLeft, IHTMLControlElement
//
//  synopsis    :   returns a long value of the client window
//      Left (inside borders)
//
//-----------------------------------------------------------
HRESULT CElement::get_clientLeft(long* pl)
{
    HRESULT hr = S_OK;

    if(!pl)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pl = 0;

    if(IsInMarkup() && Doc()->GetView()->IsActive())
    {
        CLayout* pLayout;

        SendNotification(NTYPE_ELEMENT_ENSURERECALC);

        pLayout = GetUpdatedLayout();

        if(pLayout)
        {
            CDispNode* pDispNode= pLayout->GetElementDispNode();

            if(pDispNode)
            {
                // border and scroll widths are dynamic. This method
                // provides a good way to get the client left amount
                // without having to have special knowledge of the
                // display tree workings. We are getting the distance
                // between the top left of the client rect and the
                // top left of the container rect.
                CRect rcClient;
                pLayout->GetClientRect(&rcClient);
                CPoint pt(rcClient.TopLeft());

                pDispNode->TransformPoint(&pt, COORDSYS_CONTENT, COORDSYS_CONTAINER);

                *pl = pt.x;
            }
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::blur()
{
    HRESULT hr = S_OK;
    CDocument* pDoc = Doc();

    if(ETAG_AREA == Tag())
    {
        AssertSz(FALSE, "must improve");
        goto Cleanup;
    }

    if(HasLayout() && !IsInMarkup())
    {
        hr = E_UNEXPECTED;
        goto Cleanup;
    }

    // if called on body, delegate to the window
    if(Tag() == ETAG_BODY)
    {
        /*pDoc->EnsureOmWindow();
        if(pDoc->_pOmWindow)
        {
            hr = pDoc->_pOmWindow->blur();
        } wlw note*/

        goto Cleanup;
    }

    // Don't blur current object in focus if called on
    // another object that does not have the focus or if the
    // frame in which this object is does not currently have the focus
    if(this!=pDoc->_pElemCurrent ||
        ((::GetFocus()!=pDoc->_pInPlace->_hwnd) && Tag()!=ETAG_SELECT))
    {
        goto Cleanup;
    }

    hr = pDoc->GetPrimaryElementTop()->BecomeCurrent(0);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::put_onresize(VARIANT v)
{
    HRESULT hr;

    // only delegate to the window if this is body
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();

        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->put_onresize(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonresize.a.HandleCodeProperty(
            HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::get_onresize(VARIANT* p)
{
    HRESULT hr;

    // only delegate to the window if this is body
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();

        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->get_onresize(p); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonresize.a.HandleCodeProperty(
            HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            p,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement:: put_onfocus(VARIANT v)
{
    HRESULT hr = E_FAIL;

    // anchors and areas override this, body delgate, all
    // other cases just go into the scite propdesc for now.. .since
    // those are the only elements that currently have this defined
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();
        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->put_onfocus(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonfocus.a.HandleCodeProperty(
            HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement:: get_onfocus(VARIANT* v)
{
    HRESULT hr = E_FAIL;

    // anchors and areas override this, body delgate, all
    // other cases just go into the scite propdesc for now.. .since
    // those are the only elements that currently have this defined
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();
        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->get_onfocus(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonfocus.a.HandleCodeProperty(
            HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement:: put_onblur(VARIANT v)
{
    HRESULT hr = E_FAIL;

    // anchors and areas override this, body delgate, all
    // other cases just go into the scite propdesc for now.. .since
    // those are the only elements that currently have this defined
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();
        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->put_onblur(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonblur.a.HandleCodeProperty(
            HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement:: get_onblur(VARIANT* v)
{
    HRESULT hr = E_FAIL;

    // anchors and areas override this, body delgate, all
    // other cases just go into the scite propdesc for now.. .since
    // those are the only elements that currently have this defined
    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();
        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->get_onblur(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonblur.a.HandleCodeProperty(
            HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::focus()
{
    return focusHelper(0);
}

HRESULT CElement::get_tabIndex(short* puTabIndex)
{
    short tabIndex = GetAAtabIndex();

    *puTabIndex = (tabIndex==htmlTabIndexNotSet) ? 0 : tabIndex;
    return S_OK;
}

STDMETHODIMP CElement::put_onscroll(VARIANT v)
{
    HRESULT hr;

    if(Tag() == ETAG_BODY)
    {
        /*CDocument*  pDoc = Doc();

        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->put_onscroll(v); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonscroll.a.HandleCodeProperty(
            HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

STDMETHODIMP CElement::get_onscroll(VARIANT* p)
{
    HRESULT hr;

    if(Tag() == ETAG_BODY)
    {
        /*CDocument* pDoc = Doc();

        hr = pDoc->EnsureOmWindow();
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->_pOmWindow->get_onscroll(p); wlw note*/
    }
    else
    {
        hr = s_propdescCElementonscroll.a.HandleCodeProperty(
            HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            p,
            this,
            CVOID_CAST(GetAttrArray()));
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+-------------------------------------------------------------------------------
//
//  Member:     doScroll
//
//  Synopsis:   Implementation of the automation interface method.
//              this simulates a click on the particular component of the scrollbar
//              (if this txtsite has one)
//
//+-------------------------------------------------------------------------------
STDMETHODIMP CElement::doScroll(VARIANT varComponent)
{
    HRESULT         hr = E_PENDING;
    htmlComponent   eComp;
    int             iDir;
    WORD            wComp;
    CVariant        varCompStr;
    CLayout*        pLayout = GetUpdatedLayout();
    CDocument*      pDoc = Doc();

    SendNotification(NTYPE_ELEMENT_ENSURERECALC);

    // the paramter is optional, and if nothing was provide then use the defualt
    hr = varCompStr.CoerceVariantArg(&varComponent, VT_BSTR);
    if(hr == S_OK)
    {
        if(!SysStringLen(V_BSTR(&varCompStr)))
        {
            eComp = htmlComponentSbPageDown;
        }
        else
        {
            hr = ENUMFROMSTRING(
                htmlComponent,
                V_BSTR(&varCompStr),
                (long*)&eComp);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }
    else if(hr == S_FALSE)
    {
        // when no argument
        eComp = htmlComponentSbPageDown;
        hr = S_OK;
    }
    else
    {
        goto Cleanup;
    }

    // no that we know what we are doing, initialize the parametes for
    // the onscroll helper fx
    switch(eComp)
    {
    case htmlComponentSbLeft:
    case htmlComponentSbLeft2:
        iDir = 0;
        wComp = SB_LINELEFT;
        break;
    case htmlComponentSbPageLeft:
    case htmlComponentSbPageLeft2:
        iDir = 0;
        wComp = SB_PAGELEFT;
        break;
    case htmlComponentSbPageRight:
    case htmlComponentSbPageRight2:
        iDir = 0;
        wComp = SB_PAGERIGHT;
        break;
    case htmlComponentSbRight:
    case htmlComponentSbRight2:
        iDir = 0;
        wComp = SB_LINERIGHT;
        break;
    case htmlComponentSbUp:
    case htmlComponentSbUp2:
        // equivalent to up arrow
        iDir = 1;
        wComp = SB_LINEUP;
        break;
    case htmlComponentSbPageUp:
    case htmlComponentSbPageUp2:
        iDir = 1;
        wComp = SB_PAGEUP;
        break;
    case htmlComponentSbPageDown:
    case htmlComponentSbPageDown2:
        iDir = 1;
        wComp = SB_PAGEDOWN;
        break;
    case htmlComponentSbDown:
    case htmlComponentSbDown2:
        iDir = 1;
        wComp = SB_LINEDOWN;
        break;
    case htmlComponentSbTop:
        iDir = 1;
        wComp = SB_TOP;
        break;
    case htmlComponentSbBottom:
        iDir = 1;
        wComp = SB_BOTTOM;
        break;
    default:
        // nothing to do in this case.  hr is S_OK
        goto Cleanup;
    }

    // Send the request to the layout, if any
    if(pLayout)
    {
        hr = pLayout->OnScroll(iDir, wComp, 0);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+-------------------------------------------------------------------------------
//
//  Member:     componentFromPoint
//
//  Synopsis:   Base Implementation of the automation interface IHTMLELEMENT2 method.
//              This method returns none, meaning no component is hit, and is here
//              for future expansion when the component list includes borders and
//              margins and such.
//              Currently, only scrollbar components are implemented, so there is an
//              overriding implementation in CLayout which handles the hit testing
//              against the scrollbar, and determines which component is hit.
//
//+-------------------------------------------------------------------------------
STDMETHODIMP CElement::componentFromPoint(long x, long y, BSTR* pbstrComponent)
{
    HRESULT hr;
    WORD eComp = htmlComponentOutside;

    if(!pbstrComponent)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    if(HasLayout())
    {
        eComp = GetUpdatedLayout()->ComponentFromPoint(x, y);
    }

    // if we already know what we have, return
    if(eComp == htmlComponentOutside)
    {
        CTreeNode*  pNodeElement;
        HTC         htc;
        CMessage    msg;

        msg.pt.x = x;
        msg.pt.y = y;

        // is this point over us?
        htc = Doc()->HitTestPoint(&msg, &pNodeElement, HT_VIRTUALHITTEST);

        if(htc == HTC_NO)
        {
            eComp = htmlComponentOutside;
        }
        else
        {
            // if we are in design mode, worry about grab handles, but only
            // if this element is the current one.   but since these HTC's
            // don't get returned unless we are in debug mode this is fine
            // to test straight.
            switch(htc)
            {
            case HTC_TOPLEFTHANDLE:
                eComp = htmlComponentGHTopLeft;
                break;

            case HTC_LEFTHANDLE:
                eComp = htmlComponentGHLeft;
                break;

            case HTC_TOPHANDLE:
                eComp = htmlComponentGHTop;
                break;

            case HTC_BOTTOMLEFTHANDLE:
                eComp = htmlComponentGHBottomLeft;
                break;

            case HTC_TOPRIGHTHANDLE:
                eComp = htmlComponentGHTopRight;
                break;

            case HTC_BOTTOMHANDLE:
                eComp = htmlComponentGHBottom;
                break;

            case HTC_RIGHTHANDLE:
                eComp = htmlComponentGHRight;
                break;

            case HTC_BOTTOMRIGHTHANDLE:
                eComp = htmlComponentGHBottomRight;
                break;

            default:
                // do nothing
                if(this == pNodeElement->Element())
                {
                    eComp = htmlComponentClient;
                }
                break;
            }
        }
    }

    hr = STRINGFROMENUM(htmlComponent, eComp, pbstrComponent);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CElement::setCapture(VARIANT_BOOL containerCapture)
{
    HRESULT hr = S_OK;
    CDocument* pDoc = Doc();

    if(!pDoc || pDoc->_fOnLoseCapture)
    {
        goto Cleanup;
    }

    if(pDoc->_pElementOMCapture == this)
    {
        goto Cleanup;
    }

    hr = pDoc->releaseCapture();
    if(hr)
    {
        goto Cleanup;
    }

    pDoc->_fContainerCapture = (containerCapture == VB_TRUE);
    pDoc->_pElementOMCapture = this;
    SubAddRef();

    if(::GetCapture() != pDoc->_pInPlace->_hwnd)
    {
        ::SetCapture(pDoc->_pInPlace->_hwnd);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}




BOOL SameScope(CTreeNode* pNode1, const CElement* pElement2)
{
    // Both NULL
    if(!pNode1 && !pElement2)
    {
        return TRUE;
    }

    return (pNode1&&pElement2 ? pNode1->Element()==pElement2 : FALSE);
}

BOOL SameScope(const CElement* pElement1, CTreeNode* pNode2)
{
    // Both NULL
    if(!pElement1 && !pNode2)
    {
        return TRUE;
    }

    return (pElement1&&pNode2 ? pElement1==pNode2->Element() : FALSE);
}

BOOL SameScope(CTreeNode* pNode1, CTreeNode* pNode2)
{
    if(pNode1 == pNode2)
    {
        return TRUE;
    }

    return (pNode1&&pNode2 ? pNode1->Element()==pNode2->Element() : FALSE);
}

void TransformSlaveToMaster(CTreeNode ** ppNode)
{
    Assert(ppNode);
    if(!*ppNode || (*ppNode)->Element()->Tag()!=ETAG_TXTSLAVE)
    {
        return;
    }

    CElement* pElem = (*ppNode)->Element()->MarkupMaster();

    *ppNode = (pElem) ? pElem->GetFirstBranch() : NULL;
}

//+------------------------------------------------------------------------
//
//  Helper:     StoreLineAndOffsetInfo
//
//  Synopsis:   stores line and offset information for a property in attr array of the object
//
//-------------------------------------------------------------------------
HRESULT StoreLineAndOffsetInfo(CBase* pBaseObj, DISPID dispid, ULONG uLine, ULONG uOffset)
{
    HRESULT hr;
    // pchData will be of the form "ulLine ulOffset", for example: "13 1313"
    TCHAR   pchData [30];   // in only needs to be 21, but let's be safe

    hr = Format(_afxGlobalData._hInstResource, 0, &pchData, 30, _T("<0du> <1du>"), uLine, uOffset);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pBaseObj->AddString(dispid, pchData, CAttrValue::AA_Internal);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CElement::GetLineAndOffsetInfo
//
//  Synopsis:   retrieves line and offset information for a property from attr array
//
//  Returns:    S_OK        successfully retrieved the data
//              S_FALSE     no error, but the information in attr array is not
//                          a line/offset string
//              FAILED(hr)  generic error condition
//
//-------------------------------------------------------------------------
HRESULT GetLineAndOffsetInfo(CAttrArray* pAA, DISPID dispid, ULONG* puLine, ULONG* puOffset)
{
    HRESULT     hr;
    CAttrValue* pAV;
    AAINDEX     aaIdx;
    LPTSTR      pchData;
    TCHAR*      pchTempStart;
    TCHAR*      pchTempEnd;

    Assert(puOffset && puLine);

    (*puOffset) = (*puLine) = 0; // set defaults

    // get the information string
    aaIdx = AA_IDX_UNKNOWN;
    pAV = pAA->Find(dispid, CAttrValue::AA_Internal, &aaIdx);
    if(!pAV || VT_LPWSTR!=pAV->GetAVType())
    {
        hr = S_FALSE;
        goto Cleanup;
    }

    pchData = pAV->GetString();

    // pchData is of the format: "lLine lOffset", for example: "13 1313"
    // Here we crack the string apart.
    pchTempStart = pchData;
    pchTempEnd = _tcschr(pchData, _T(' '));

    Assert (pchTempEnd);

    *pchTempEnd = _T('\0');
    hr = ttol_with_error(pchTempStart, (LONG*)puLine);
    *pchTempEnd = _T(' ');
    if(hr)
    {
        goto Cleanup;
    }

    pchTempStart = ++pchTempEnd;

    hr = ttol_with_error(pchTempStart, (LONG*)puOffset);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}
