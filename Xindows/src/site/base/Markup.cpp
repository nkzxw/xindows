
#include "stdafx.h"
#include "Markup.h"

#define _cxx_
#include "../gen/markup.hdl"

const GUID FAR CLSID_CMarkup = { 0x3050f233, 0x98b5, 0x11cf, { 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };
const GUID FAR IID_IHTMLViewServices = { 0x3050f603, 0x98b5, 0x11cf, { 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };

// allocate this many CTreePos objects at a time
static const size_t s_cTreePosPoolSize = 64;

// splay when accessing something deeper than this
static const long s_cSplayDepth = 10;



class CCollectionBuildContext
{
public:
    CDocument* _pDoc;
    // Collection State
    long _lCollection;
    BOOL _fNeedNameID;
    BOOL _fNeedForm;

    CCollectionBuildContext(CDocument* pDoc)
    {
        _pDoc = pDoc;
        _lCollection = 0;
        _fNeedNameID = FALSE;
        _fNeedForm = FALSE;
    };
};

class CAllCollectionCacheItem : public CCollectionCacheItem
{
private:
    LONG _lCurrentIndex;
    CMarkup* GetMarkup(void) { return _pMarkup; }
    CMarkup* _pMarkup;
public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    CElement* GetNext(void);
    CElement* MoveTo(long lIndex);
    CElement* GetAt(long lIndex);
    long Length(void);

    void SetMarkup(CMarkup* pMarkup) { _pMarkup = pMarkup; }
};

CElement* CAllCollectionCacheItem::GetNext(void)
{
    return GetAt ( _lCurrentIndex++ );
}

CElement* CAllCollectionCacheItem::MoveTo(long lIndex)
{
    _lCurrentIndex = lIndex;
    return NULL;
}

CElement* CAllCollectionCacheItem::GetAt(long lIndex)
{
    CMarkup* pMarkup = GetMarkup();
    CTreePos* ptpBegin;

    Assert(pMarkup);

    Assert(lIndex >= 0);

    // Skip ETAG_ROOT which is always the zero'th element
    lIndex++;

    if(lIndex >= pMarkup->NumElems())
    {
        return NULL;
    }

    ptpBegin = pMarkup->TreePosAtSourceIndex(lIndex);

    Assert(ptpBegin && ptpBegin->IsBeginElementScope());

    return ptpBegin->Branch()->Element();
}

long CAllCollectionCacheItem::Length(void)
{
    CMarkup* pMarkup = GetMarkup();
    Assert(pMarkup && pMarkup->NumElems()>=1);
    return pMarkup->NumElems()-1;
}



CMarkupTextFragContext::~CMarkupTextFragContext()
{
    int cFrag;
    MarkupTextFrag* ptf;

    for(cFrag=_aryMarkupTextFrag.Size(),ptf=_aryMarkupTextFrag; cFrag>0; cFrag--,ptf++)
    {
        MemFreeString(ptf->_pchTextFrag);
    }
}

HRESULT CMarkupTextFragContext::AddTextFrag(CTreePos* ptpTextFrag, TCHAR* pchTextFrag, ULONG cchTextFrag, long iTextFrag)
{
    HRESULT         hr = S_OK;
    TCHAR*          pchCopy = NULL;
    MarkupTextFrag* ptf = NULL;

    Assert(iTextFrag>=0 && iTextFrag<=_aryMarkupTextFrag.Size());

    // Allocate and copy the string
    hr = MemAllocStringBuffer(cchTextFrag, pchTextFrag, &pchCopy);
    if(hr)
    {
        goto Cleanup;
    }

    // Allocate the TextFrag object in the array
    hr = _aryMarkupTextFrag.InsertIndirect(iTextFrag, NULL);
    if(hr)
    {
        MemFreeString(pchCopy);
        goto Cleanup;
    }

    // Fill in the text frag
    ptf = &_aryMarkupTextFrag[iTextFrag];
    Assert(ptf);

    ptf->_ptpTextFrag = ptpTextFrag;
    ptf->_pchTextFrag = pchCopy;

Cleanup:
    RRETURN(hr);
}

HRESULT CMarkupTextFragContext::RemoveTextFrag(long iTextFrag, CMarkup* pMarkup)
{
    HRESULT         hr = S_OK;
    MarkupTextFrag* ptf;

    Assert(iTextFrag>=0 && iTextFrag<=_aryMarkupTextFrag.Size());

    // Get the text frag
    ptf = &_aryMarkupTextFrag[iTextFrag];

    // Free the string
    MemFreeString(ptf->_pchTextFrag);

    hr = pMarkup->RemovePointerPos(ptf->_ptpTextFrag, NULL, NULL);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    // Remove the entry from the array
    _aryMarkupTextFrag.Delete(iTextFrag);

    RRETURN( hr );
}

long CMarkupTextFragContext::FindTextFragAtCp(long cpFind, BOOL* pfFragFound)
{
    // Do a binary search through the array to find the spot in the array
    // where cpFind lies
    int             iFragLow, iFragHigh, iFragMid;
    MarkupTextFrag* pmtfAry = _aryMarkupTextFrag;
    MarkupTextFrag* pmtfMid;
    BOOL            fResult;
    long            cpMid;

    iFragLow  = 0;
    fResult   = FALSE;
    iFragHigh = _aryMarkupTextFrag.Size() - 1;

    while(iFragLow <= iFragHigh)
    {
        iFragMid = (iFragLow + iFragHigh) >> 1;

        pmtfMid  = pmtfAry + iFragMid;
        cpMid    = pmtfMid->_ptpTextFrag->GetCp();

        if(cpMid == cpFind)
        {
            iFragLow = iFragMid;
            fResult = TRUE;
            break;
        }
        else if(cpMid < cpFind)
        {
            iFragLow = iFragMid + 1;
        }
        else
        {
            iFragHigh = iFragMid - 1;
        }
    }

    if(fResult)
    {
        // Search backward through all of the entries
        // at the same cp so we return the first one
        iFragLow--;
        while(iFragLow)
        {
            if(pmtfAry[iFragLow]._ptpTextFrag->GetCp() != cpFind)
            {
                iFragLow++;
                break;
            }
            iFragLow--;
        }
    }

    if(pfFragFound)
    {
        *pfFragFound = fResult;
    }

    Assert(iFragLow==0 || pmtfAry[iFragLow-1]._ptpTextFrag->GetCp()<cpFind);
    Assert(iFragLow==_aryMarkupTextFrag.Size()-1 
        || _aryMarkupTextFrag.Size()==0
        || cpFind<=pmtfAry[iFragLow]._ptpTextFrag->GetCp());
    Assert(!fResult || pmtfAry[iFragLow]._ptpTextFrag->GetCp()==cpFind);

    return iFragLow;
}




const CBase::CLASSDESC CMarkup::s_classdesc =
{
    &CLSID_HTMLDocument,            // _pclsid
    CDocument::s_acpi,              // _pcpi
    0,                              // _dwFlags
    &IID_IHTMLDocument2,            // _piidDispinterface
    &CDocument::s_apHdlDescs,       // _apHdlDesc
};

BEGIN_TEAROFF_TABLE(CMarkup, IMarkupContainer)
    TEAROFF_METHOD(CMarkup, &OwningDoc, owningdoc, (IHTMLDocument2** ppDoc))
END_TEAROFF_TABLE()

BEGIN_TEAROFF_TABLE(CMarkup, ISelectionRenderingServices)
    TEAROFF_METHOD(CMarkup, &MovePointersToSegment, movepointerstosegment, (int iSegmentIndex, IMarkupPointer* pStart, IMarkupPointer* pEnd))
    TEAROFF_METHOD(CMarkup, &GetSegmentCount, getsegmentcount, (int* piSegmentCount, SELECTION_TYPE* peType))
    TEAROFF_METHOD(CMarkup, &AddSegment, addsegment, (IMarkupPointer* pStart, IMarkupPointer* pEnd, HIGHLIGHT_TYPE HighlightType, int* iSegmentIndex))
    TEAROFF_METHOD(CMarkup, &AddElementSegment, addelementsegment, (IHTMLElement* pElement , int * iSegmentIndex))
    TEAROFF_METHOD(CMarkup, &MoveSegmentToPointers, movesegmenttopointers, (int iSegmentIndex, IMarkupPointer* pStart, IMarkupPointer* pEnd, HIGHLIGHT_TYPE HighlightType))
    TEAROFF_METHOD(CMarkup, &GetElementSegment, getelementsegment, (int iSegmentIndex, IHTMLElement** ppElement))
    TEAROFF_METHOD(CMarkup, &SetElementSegment, setelementsegment, (int iSegmentIndex,    IHTMLElement* pElement))
    TEAROFF_METHOD(CMarkup, &ClearSegment, clearsegment, (int iSegmentIndex, BOOL fInvalidate))
    TEAROFF_METHOD(CMarkup, &ClearSegments, clearsegments, (BOOL fInvalidate))
    TEAROFF_METHOD(CMarkup, &ClearElementSegments, clearelementsegments, ())
END_TEAROFF_TABLE()


CMarkup::CMarkup(CDocument* pDoc, CElement* pElementMaster)
{
	__lMarkupTreeVersion = 1;
	__lMarkupContentsVersion = 1;

	_pDoc = pDoc;

	Assert(_pDoc && _pDoc->AreLookasidesClear(this, LOOKASIDE_MARKUP_NUMBER));

	_pElementMaster = pElementMaster;

	if(_pElementMaster)
	{
		_fIncrementalAlloc = TRUE;
	}
}

CMarkup::~CMarkup()
{
	ClearLookasidePtrs();
}

HRESULT CMarkup::Init(CRootElement* pElementRoot)
{
	HRESULT hr;

	Assert(pElementRoot);

	_pElementRoot = pElementRoot;

	_tpRoot.MarkLast();
	_tpRoot.MarkRight();

	hr = CreateInitialMarkup(_pElementRoot);
	if(hr)
	{
		goto Cleanup;
	}

	hr = super::Init();
	if(hr)
	{
		goto Cleanup;
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::CreateInitialMarkup(CRootElement* pElementRoot)
{
	HRESULT		hr = S_OK;
	CTreeNode*	pNodeRoot;
	CTreePos*	ptpNew;

	// Assert that there is nothing in the splay tree currently
	Assert(FirstTreePos() == NULL);

	pNodeRoot = new CTreeNode(NULL, pElementRoot);
	if(!pNodeRoot)
	{
		hr = E_OUTOFMEMORY;
		goto Cleanup;
	}

	// The initial ref on the node will transfer to the tree
	ptpNew = pNodeRoot->InitBeginPos(TRUE);
	Verify(!Append(ptpNew));

	ptpNew = pNodeRoot->InitEndPos(TRUE);
	Verify(!Append(ptpNew));

	pNodeRoot->PrivateEnterTree();

	pElementRoot->SetMarkupPtr(this);
	pElementRoot->__pNodeFirstBranch = pNodeRoot;
	pElementRoot->PrivateEnterTree();

	// Insert the chars for the node poses
	Verify(CTxtPtr(this, 0).InsertRepeatingChar(2, WCH_NODE) == 2);

	{
		CNotification nf;
		nf.ElementEntertree(pElementRoot);
		pElementRoot->Notify(&nf);
	}

	UpdateMarkupTreeVersion();

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::UnloadContents(BOOL fForPassivate/*= FALSE*/)
{
	HRESULT hr = S_OK;
	CElement::CLock lock(_pElementRoot);

	// text and tree
	delete _pSelRenSvcProvider;
	_pSelRenSvcProvider = NULL;

	_TxtArray.RemoveAll();

	hr = DestroySplayTree(!fForPassivate);
	if(hr)
	{
		goto Cleanup;
	}

	// Restore initial state
	_fIncrementalAlloc = !!_pElementMaster;

	_fNoUndoInfo = FALSE;
	_fLoaded = FALSE;
	_fStreaming = FALSE;

    if(_pProgSink)
    {
        _pProgSink->Detach();
        _pProgSink->Release();
        _pProgSink = NULL;
    }

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::DestroySplayTree(BOOL fReinit)
{
	HRESULT		hr = S_OK;
	CTreePos*	ptp;
	CTreePos*	ptpNext;
	CTreeNode*	pDelayReleaseList = NULL;
	CTreeNode*	pExitTreeScList = NULL;
	CNotification nfExit;

	// Prime the exittree notification
	nfExit.ElementExittree1(NULL);
	Assert(nfExit.IsSecondChanceAvailable());

	// Unposition any unembedded markup pointers.  Because there is
	// no pointer pos for these pointers, they will not get notified.
	while(_pmpFirst)
	{
		hr = _pmpFirst->Unposition();

		if(hr)
		{
			goto Cleanup;
		}
	}

	// The walk used here is destructive.  This is needed
	// because once a CTreePos is released (via ReleaseContents)
	// we can't use it any more.  So before we release a pos,
	// we short circuit other pointers in the tree to not point
	// to it any more.  This way, we only visit (and release) each
	// node only once.

	// set a sentinel value to make the traversal end properly
	if(_tpRoot.FirstChild())
	{
		_tpRoot.FirstChild()->SetNext(NULL);
	}

	// release the tree
	for(ptp=_tpRoot.FirstChild(); ptp;)
	{
		// figure out the next to be deleted
		if(ptp->FirstChild())
		{
			ptpNext = ptp->FirstChild();

			// this is the short circuit
			if(ptpNext->IsLastChild())
			{
				ptpNext->SetNext(ptp->Next());
			}
			else
			{
				Assert(ptpNext->Next()->IsLastChild());
				ptpNext->Next()->SetNext(ptp->Next());
			}
		}
		else
		{
			ptpNext = ptp->Next();
		}

		if(ptp->IsNode())
		{
			CTreeNode* pNode = ptp->Branch();

			ReleaseTreePos(ptp, TRUE);

			if(!pNode->_fInMarkupDestruction)
			{
				pNode->_fInMarkupDestruction = TRUE;
			}
			else
			{
				BOOL fDelayRelease = FALSE;
				BOOL fExitTreeSc = FALSE;

				if(pNode->IsFirstBranch())
				{
					// Prepare and send the exit tree notification
					CElement*	pElement = pNode->Element();
					DWORD		dwData = EXITTREE_DESTROY;

					if(pElement->GetObjectRefs() == 1)
					{
						dwData |= EXITTREE_PASSIVATEPENDING;
					}

					pElement->_fExittreePending = TRUE;

					// NOTE: we may want to bump up the notification SN here
					nfExit.SetElement(pElement);
					nfExit.SetData(dwData);
					pElement->Notify(&nfExit);

					pElement->_fExittreePending = FALSE;

					// Check to see if we have to delay release
					fDelayRelease = nfExit.DataAsDWORD() & EXITTREE_DELAYRELEASENEEDED;

					// Check to see if we have to send ElementExittreeSc
					fExitTreeSc = nfExit.IsSecondChanceRequested();

					if(!fDelayRelease && !fExitTreeSc)
					{
						// The node is now no longer in the tree so release
						// the tree's ref on the node
						pNode->PrivateExitTree();
					}
					else
					{
						// We will use the node as an item on our delay release
						// linked list.  Make it dead and use useless fields
						pNode->PrivateMakeDead();
						pNode->GetBeginPos()->_pNext = (CTreePos*)pDelayReleaseList;
						pDelayReleaseList = pNode;
						pElement->AddRef();


						// If we need an exit tree second chance, link it into that list.
						if(fExitTreeSc)
						{
							pNode->GetBeginPos()->_pFirstChild = (CTreePos*)pExitTreeScList;
							pExitTreeScList = pNode;

							nfExit.ClearSecondChanceRequested();
						}
					}

					// Remove the element from the tree
					pElement->__pNodeFirstBranch = NULL;
					pElement->DelMarkupPtr();
					pElement->PrivateExitTree(this);
				}
				else
				{
					// The node is now no longer in the tree so release
					// the tree's ref on the node
					pNode->PrivateExitTree();
				}
			}
		}
		else
		{
			ReleaseTreePos(ptp, TRUE);
		}

		ptp = ptpNext;
	}

	// release all the TreePos pools
	while(_pvPool)
	{
		void* pvNextPool = *((void**)_pvPool);

		if(_pvPool != &(_abPoolInitial[0]))
		{
			_MemFree(_pvPool);
		}
		_pvPool = pvNextPool;
	}

	memset(&_tpRoot, 0, sizeof(CTreePos));
	_ptpFirst = NULL;
	Assert(_pvPool == NULL);
	_ptdpFree = NULL;
	memset(&_abPoolInitial, 0, sizeof(_abPoolInitial));

	_tpRoot.MarkLast();
	_tpRoot.MarkRight();

	if(fReinit)
	{
		hr = CreateInitialMarkup(_pElementRoot);
		if(hr)
		{
			_pElementRoot = NULL;
			goto Cleanup;
		}
	}
	else
	{
		_pElementRoot = NULL;
	}

	// Release and delete the text frag context
	delete DelTextFragContext();
	delete DelTopElemCache();

	UpdateMarkupTreeVersion();

	// Send all of the exit tree second chance notifications that we need to
	if(pExitTreeScList)
	{
		nfExit.ElementExittree2(NULL);
	}

	while(pExitTreeScList)
	{
		CElement* pElement = pExitTreeScList->Element();

		// Get the next link in the list
		pExitTreeScList = (CTreeNode*)pExitTreeScList->GetBeginPos()->_pFirstChild;

		// Send the after exit tree notifications
		// NOTE: we may want to bump up the notification SN here
		nfExit.SetElement(pElement);
		pElement->Notify(&nfExit);
	}

	// Release all of the elements (and nodes) on our delay release list
	while(pDelayReleaseList)
	{
		CElement*	pElement = pDelayReleaseList->Element();
		CTreeNode*	pNode = pDelayReleaseList;

		// Get the next link in the list
		pDelayReleaseList = (CTreeNode*)pDelayReleaseList->GetBeginPos()->_pNext;

		// Release the element
		pElement->Release();

		// Release any hold the markup has on the node
		pNode->PrivateMarkupRelease();
	}

Cleanup:
	RRETURN(hr);
}

void CMarkup::Passivate()
{
	// Release everything
	UnloadContents(TRUE);

	// super
	super::Passivate();
}

//+---------------------------------------------------------------------------
//
//  Member:     CreateElement
//
//  Note:       If the etag is ETAG_NULL, the the string contains the name
//              of the tag to create (like "h1") or it can contain an
//              actual tag (like "<h1 foo=bar>").
//
//              If etag is not ETAG_NULL, then the string contains the
//              arguments for the new element.
//
//----------------------------------------------------------------------------
HRESULT CMarkup::CreateElement(ELEMENT_TAG etag, CElement** ppElementNew)
{
    HRESULT hr = S_OK;
    CHtmTag ht, *pht = NULL;
    TCHAR* pchTag = NULL;
    long cchTag = 0;
    CString strTag; // In case we have to build "<name attrs>"

    ht.Reset();

    Assert(ppElementNew);

    *ppElementNew = NULL;

    // Here we check for the various kinds of input to this function.
    //
    // After this checking, we will either have pht != NULL which is
    // ready for use to pass to create element, or we will have a string
    // (pchTag, cchTag) which is appropriate for passing to the 'parser'
    // for creating an element.
    if(etag != ETAG_NULL)
    {
        pht = &ht;

        if(pht)
        {
            Assert(pht = &ht);
            ht.Reset();
            ht.SetTag(etag);
        }
    }

    if(!pht && pchTag && cchTag)
    {
        Assert(*pchTag == _T('<'));
        Assert(long( _tcslen( pchTag)) == cchTag);
    }

    if(!pht)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    if(pht->GetTag() == ETAG_NULL)
    {
    }

    pht->SetScriptCreated();

    hr = ::CreateElement(pht, ppElementNew, Doc(), this, NULL, INIT2FLAG_EXECUTE);

    if(hr)
    {
        goto Cleanup;
    }

    if(!(*ppElementNew)->IsNoScope())
    {
        (*ppElementNew)->_fExplicitEndTag = TRUE;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CMarkup::get_Script(IDispatch** ppWindow)
{
    if(ppWindow)
    {
        *ppWindow = NULL;
    }

    return _pDoc->get_Script(ppWindow);
}

HRESULT CMarkup::get_all(IHTMLElementCollection** ppDisp)
{
    RRETURN(SetErrorInfo(GetCollection(ELEMENT_COLLECTION, ppDisp)));
}

HRESULT CMarkup::get_body(IHTMLElement** ppDisp)
{
    HRESULT hr = S_OK;;
    CElement* pElementClient;

    if(!ppDisp)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ppDisp = NULL;

    pElementClient = GetElementClient();

    if(pElementClient)
    {
        Assert(pElementClient->Tag() != ETAG_ROOT);

        hr = pElementClient->QueryInterface(IID_IHTMLElement, (void**)ppDisp);

        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN1(SetErrorInfo(hr), S_FALSE);
}

HRESULT CMarkup::get_activeElement(IHTMLElement** ppElement)
{
    // BUGBUG(sramani) need to implement for view-slave elements
    if(ppElement)
    {
        *ppElement = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT STDMETHODCALLTYPE CMarkup::put_title(BSTR v)
{
    RRETURN(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE CMarkup::get_title(BSTR* p)
{
    RRETURN(E_NOTIMPL);
}

HRESULT CMarkup::get_readyState(BSTR* p)
{
    if(p)
    {
        *p = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_expando(VARIANT_BOOL* pfExpando)
{
    return _pDoc->get_expando(pfExpando);
}

HRESULT CMarkup::put_expando(VARIANT_BOOL fExpando)
{
    return _pDoc->put_expando(fExpando);
}

HRESULT CMarkup::get_parentWindow(IHTMLWindow2** ppWindow)
{
    if(ppWindow)
    {
        *ppWindow = NULL;
    }
    return _pDoc->get_parentWindow(ppWindow);
}

HRESULT CMarkup::get_nameProp(BSTR* pName)
{
    if(pName)
    {
        *pName = NULL;
    }
    return get_title(pName);
}

HRESULT CMarkup::queryCommandSupported(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::queryCommandEnabled(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::queryCommandState(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::queryCommandIndeterm(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::queryCommandText(BSTR bstrCmdId, BSTR* pcmdText)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::queryCommandValue(BSTR bstrCmdId, VARIANT* pvarRet)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::execCommand(BSTR bstrCmdId, VARIANT_BOOL showUI, VARIANT value, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::execCommandShowHelp(BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    if(pfRet)
    {
        *pfRet = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::createElement(BSTR bstrTag, IHTMLElement** pIElementNew)
{
    HRESULT hr = S_OK;
    CElement* pElement = NULL;

    if(!bstrTag || !pIElementNew)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = CreateElement(ETAG_NULL, &pElement);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pElement->QueryInterface(IID_IHTMLElement, (void**)pIElementNew);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    if(pElement)
    {
        pElement->Release();
    }

    RRETURN(hr);
}

HRESULT CMarkup::elementFromPoint(long x, long y, IHTMLElement** ppElement)
{
    if(ppElement)
    {
        *ppElement = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::toString(BSTR* pbstrString)
{
    RRETURN(super::toString(pbstrString));
}

HRESULT CMarkup::releaseCapture()
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_documentElement(IHTMLElement** ppRootElem)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_uniqueID(BSTR* pID)
{
    if(pID)
    {
        *pID = NULL;
    }
    return _pDoc->get_uniqueID(pID);
}

HRESULT CMarkup::attachEvent(BSTR event, IDispatch* pDisp, VARIANT_BOOL* pResult)
{
    RRETURN(super::attachEvent(event, pDisp, pResult));
}

HRESULT CMarkup::detachEvent(BSTR event, IDispatch* pDisp)
{
    RRETURN(super::detachEvent(event, pDisp));
}

HRESULT CMarkup::get_bgColor(VARIANT* p)
{
    if(p)
    {
        p->vt = VT_NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_bgColor(VARIANT p)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_fgColor(VARIANT* p)
{
    if(p)
    {
        p->vt = VT_NULL;
    }
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_fgColor(VARIANT p)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_linkColor(VARIANT* p)
{
    if(p)
    {
        p->vt = VT_NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_linkColor(VARIANT p)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_alinkColor(VARIANT* p)
{
    if(p)
    {
        p->vt = VT_NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_alinkColor(VARIANT p)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_vlinkColor(VARIANT* p)
{
    if(p)
    {
        p->vt = VT_NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_vlinkColor(VARIANT p)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_parentDocument(IHTMLDocument2** ppParentDoc)
{
    if(ppParentDoc)
    {
        *ppParentDoc = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_enableDownload(VARIANT_BOOL* pfDownload)
{
    if(pfDownload)
    {
        *pfDownload = VARIANT_FALSE;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_enableDownload(VARIANT_BOOL fDownload)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_baseUrl(BSTR* p)
{
    if(p)
    {
        *p = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_baseUrl(BSTR b)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::get_inheritStyleSheets(VARIANT_BOOL* pfInherit)
{
    if(pfInherit)
    {
        *pfInherit = NULL;
    }

    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::put_inheritStyleSheets(VARIANT_BOOL fInherit)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

HRESULT CMarkup::getElementsByName(BSTR v, IHTMLElementCollection** ppDisp)
{
    HRESULT hr = E_INVALIDARG;

    if(!ppDisp || !v)
    {
        goto Cleanup;
    }

    *ppDisp = NULL;

    hr = GetDispByNameOrID(v, (IDispatch**)ppDisp, TRUE);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CMarkup::getElementsByTagName(BSTR v, IHTMLElementCollection** ppDisp)
{
    HRESULT hr = E_INVALIDARG;
    if(!ppDisp || !v)
    {
        goto Cleanup;
    }

    *ppDisp = NULL;

    // Make sure our collection is up-to-date.
    hr = EnsureCollectionCache(ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    // Get a collection of the specified tags.
    hr = CollectionCache()->GetDisp(ELEMENT_COLLECTION,
        v,
        CacheType_Tag,
        (IDispatch**)ppDisp,
        FALSE); // Case sensitivity ignored for TagName
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CMarkup::getElementById(BSTR v, IHTMLElement** ppel)
{
    HRESULT hr = E_INVALIDARG;
    CElement* pel = NULL;

    if(!ppel || !v)
    {
        goto Cleanup;
    }

    *ppel = NULL;

    hr = GetElementByNameOrID(v, &pel);
    // Didn't find elem with id v, return null, if hr == S_FALSE, more than one elem, return first
    if(FAILED(hr))
    {
        hr = S_OK;
        goto Cleanup;
    }

    Assert(pel);
    hr = pel->QueryInterface(IID_IHTMLElement, (void**)ppel);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CMarkup::QueryService(REFGUID rguidService, REFIID riid, void** ppvObject)
{
	RRETURN(_pDoc->QueryService(rguidService, riid, ppvObject));
}

//+-------------------------------------------------------------------
//
//  Member:     CMarkup::GetProgSinkC
//
//--------------------------------------------------------------------
CProgSink* CMarkup::GetProgSinkC()
{
    return _pProgSink;
}

//+-------------------------------------------------------------------
//
//  Member:     CMarkup::GetProgSink
//
//--------------------------------------------------------------------
IProgSink* CMarkup::GetProgSink()
{
    return (IProgSink*)_pProgSink;
}

CElement* CMarkup::FirstElement()
{
	// BUGBUG: when we get rid of the root elem, we want
	// to take these number down by one
	if(NumElems() >= 2)
	{
		CTreePos* ptp;

		ptp = TreePosAtSourceIndex(1);
		Assert(ptp);
		return ptp->Branch()->Element();
	}

	return NULL;
}

void CMarkup::GetContentTreeExtent(CTreePos** pptpStart, CTreePos** pptpEnd)
{
	// BUGBUG: this implementation is a hack and should change when we
	// get rid of the root and slave element
	CTreePos* ptpStart;

	Assert(Master());
	ptpStart = FirstTreePos()->NextTreePos();

	Assert(ptpStart);

	// Find first element (besides the root) in the markup.
	while(!ptpStart->IsBeginNode())
	{
		Assert(ptpStart);
		ptpStart = ptpStart->NextTreePos();
	}

	Assert(ptpStart);

	if(pptpStart)
	{
		*pptpStart = ptpStart;
	}

	if(pptpEnd)
	{
		ptpStart->Branch()->Element()->GetTreeExtent(NULL, pptpEnd);
	}
}

BOOL CMarkup::IsEditable(BOOL fCheckContainerOnly) const
{
	if(_pElementMaster)
	{
		return _pElementMaster->IsEditable(fCheckContainerOnly);
	}

	return FALSE;
}

void CMarkup::EnsureTopElems()
{
	CElement*		pElementClientCached;

	if(Doc()->GetDocTreeVersion() == _lTopElemsVersion)
	{
		return;
	}

	_lTopElemsVersion = Doc()->GetDocTreeVersion();

	static ELEMENT_TAG atagStop[1] = { ETAG_BODY };

	CChildIterator ci(
		Root(),
		NULL,
		CHILDITERATOR_USETAGS,
		atagStop, ARRAYSIZE(atagStop));
	int nFound = 0;

	CTreeNode* pNode;

	//pHtmlElementCached = NULL;
	//pHeadElementCached = NULL;
	//pTitleElementCached = NULL;
	pElementClientCached = NULL;

	while(nFound<4 && (pNode=ci.NextChild())!=NULL )
	{
		CElement* pElement = pNode->Element();

		switch(pElement->Tag())
		{
		case ETAG_BODY :
		case ETAG_TXTSLAVE:
			{
				if(!pElementClientCached)
				{
					pElementClientCached = pElement;
					nFound++;
				}

				break;
			}
		}
	}

	if(nFound)
	{
		CMarkupTopElemCache* ptec = EnsureTopElemCache();

		if(ptec)
		{
			ptec->__pElementClientCached = pElementClientCached;
		}
	}
	else
	{
		delete DelTopElemCache();
	}
}

CElement* CMarkup::GetElementTop()
{
	CElement* pElement = GetElementClient();

	return pElement?pElement:Root();
}

CLayout* CMarkup::GetRunOwner(CTreeNode* pNode, CLayout* pLayoutParent/*= NULL*/)
{
	CTreeNode* pNodeRet = GetRunOwnerBranch(pNode, pLayoutParent);
	if(pNodeRet)
	{
		Assert(pNodeRet->GetUpdatedLayout());

		return pNodeRet->GetUpdatedLayout();
	}
	else
	{
		return NULL;
	}
}

CTreeNode* CMarkup::GetRunOwnerBranch(CTreeNode* pNode, CLayout* pLayoutContext)
{
	CTreeNode*	pNodeLayoutOwnerBranch;
	CTreeNode*	pNodeLayoutBranch;
	CElement*	pElementLytCtx = pLayoutContext ? pLayoutContext->ElementOwner() : NULL;

	Assert(pNode);
	Assert(pNode->GetUpdatedNearestLayout());
	Assert(!pElementLytCtx || pElementLytCtx->IsRunOwner());

	pNodeLayoutBranch = pNode->GetUpdatedNearestLayoutNode();
	pNodeLayoutOwnerBranch = NULL;

	do
	{
		CLayout* pLayout;

		// set this to the new branch every step
		pLayout = pNodeLayoutBranch->GetUpdatedLayout();

		// if we hit the context get out
		if(SameScope(pNodeLayoutBranch, pElementLytCtx))
		{
			// if we don't have an owner yet the context
			// branch is returned
			if(!pNodeLayoutOwnerBranch)
			{
				pNodeLayoutOwnerBranch = pNodeLayoutBranch;
			}
			break;
		}

		if(pNodeLayoutBranch->Element()->IsRunOwner())
		{
			pNodeLayoutOwnerBranch = pNodeLayoutBranch;

			if(!pElementLytCtx)
			{
				break;
			}
		}

		pNodeLayoutBranch = pNodeLayoutBranch->GetUpdatedParentLayoutNode();
	} while(pNodeLayoutBranch);

	return (pNodeLayoutBranch?pNodeLayoutOwnerBranch:NULL);
}

//+----------------------------------------------------------------------------
//
//    Member:   RemoveElement
//
//  Synopsis:   Removes the influence of an element
//
//-----------------------------------------------------------------------------
static void RemoveNodeChars(CMarkup* pMarkup, long cp, long cch, CTreeNode* pNode)
{
    CNotification nf;

    Assert(cp>=0 && cp<pMarkup->GetTextLength());
    Assert(cch>0 && cp+cch<=pMarkup->GetTextLength());

    nf.CharsDeleted(cp, cch, pNode);

#ifdef _DEBUG
    {
        CTxtPtr tp2(pMarkup, cp);

        for(int i=0; i<cch; i++,tp2.AdvanceCp(1))
        {
            Assert(tp2.GetChar() == WCH_NODE);
        }
    }
#endif

    CTxtPtr(pMarkup, cp).DeleteRange(cch);
    pMarkup->Notify(&nf);
}

HRESULT CMarkup::RemoveElementInternal(CElement* pElementRemove, DWORD dwFlags)
{
	HRESULT				hr = S_OK;
	BOOL				fDOMOperation = dwFlags&MUS_DOMOPERATION;
	CTreePosGap			tpgElementBegin(TPG_LEFT);
	CTreePosGap			tpgEnd(TPG_RIGHT);
	CTreeNode*			pNodeCurr, *pNodeNext;
	long				siStart;
	BOOL				fDelayRelease=FALSE, fExitTreeSc=FALSE;
	CMarkup::CLock		MarkupLock(this);
	//CRemoveElementUndo	Undo(this, pElementRemove, dwFlags);

	// BUGBUG (EricVas): Specialcase the removal of an empty element so that
	//                   multiple notifications can be amortized

	// The element to be removed must be specified and in this tree.  Also,
	// We better not be trying to remove the tree element itself!
	Assert(pElementRemove);
	Assert(pElementRemove->GetMarkup());
	Assert(pElementRemove->GetMarkup() == this);
	Assert(!pElementRemove->IsRoot());

	// Notify element it is about to exit the tree

	{
		CNotification nf;

		Assert(!pElementRemove->_fExittreePending);
		pElementRemove->_fExittreePending = TRUE;

		nf.ElementExittree1(pElementRemove);

		Assert(nf.IsSecondChanceAvailable());

		// If we are in the undo queue, we will have _ulRefs>1 by now.
		if(pElementRemove->GetObjectRefs() == 1)
		{
			nf.SetData(EXITTREE_PASSIVATEPENDING);
		}

		pElementRemove->Notify(&nf);

		if(nf.IsSecondChanceRequested())
		{
			fDelayRelease = TRUE;
			fExitTreeSc = TRUE;
		}
		else if(nf.DataAsDWORD() & EXITTREE_DELAYRELEASENEEDED)
		{
			fDelayRelease = TRUE;
		}

		pElementRemove->_fExittreePending = FALSE;
	}

	hr = EmbedPointers();

	if(hr)
	{
		goto Cleanup;
	}

	// Remove tree pointers with cling around the edges of the element being removed
	{
		CTreePosGap tpgCling;
		CTreePos *ptpBegin, *ptpEnd;

		pElementRemove->GetTreeExtent(&ptpBegin, &ptpEnd);

		Assert(ptpBegin && ptpEnd);

		// Around the beginning
		Verify(!tpgCling.MoveTo(ptpBegin, TPG_LEFT));
		tpgCling.PartitionPointers(this, fDOMOperation);
		tpgCling.CleanCling(this, TPG_RIGHT, fDOMOperation);

		Verify(!tpgCling.MoveTo(ptpBegin, TPG_RIGHT));
		tpgCling.PartitionPointers(this, fDOMOperation);
		tpgCling.CleanCling(this, TPG_LEFT, fDOMOperation);

		// Around the end
		Verify(!tpgCling.MoveTo(ptpEnd, TPG_LEFT));
		Verify(!tpgCling.MoveLeft(TPG_VALIDGAP|TPG_OKNOTTOMOVE));
		tpgCling.PartitionPointers(this, fDOMOperation);
		tpgCling.CleanCling(this, TPG_RIGHT, fDOMOperation);

		Verify(!tpgCling.MoveTo(ptpEnd, TPG_RIGHT));
		Verify(!tpgCling.MoveRight(TPG_VALIDGAP|TPG_OKNOTTOMOVE));
		tpgCling.PartitionPointers(this, fDOMOperation);
		tpgCling.CleanCling(this, TPG_LEFT, fDOMOperation);
	}

	// Record where the element begins, making sure the gap used to record this
	// points to the tree pos to the left so that stuff removed from the tree
	// does not include that.
	tpgElementBegin.MoveTo(pElementRemove->GetFirstBranch()->GetBeginPos(), TPG_LEFT);

	// Remember the source index of the element that we are removing
	siStart = pElementRemove->GetFirstBranch()->GetBeginPos()->SourceIndex(); 

	// Run through the context chain removing contexts in their entirety
	for(pNodeCurr=pElementRemove->GetFirstBranch(); pNodeCurr; pNodeCurr=pNodeNext)
	{
		CTreeNode*	pNodeParent;
		CTreePosGap	tpgNodeBegin(TPG_LEFT);
		BOOL		fBeginEdge;

		// Before we nuke this node, record its parent and next context
		pNodeParent = pNodeCurr->Parent();
		pNodeNext = pNodeCurr->NextBranch();

		// Record the gaps where the node is right now
		tpgNodeBegin.MoveTo(pNodeCurr->GetBeginPos(), TPG_RIGHT);
		tpgEnd.MoveTo(pNodeCurr->GetEndPos(), TPG_LEFT);

		// Make all the immediate children which were beneath this node point
		// to its parent.
		hr = ReparentDirectChildren(pNodeParent, &tpgNodeBegin, &tpgEnd);

		if(hr)
		{
			goto Cleanup;
		}

		// Swing the pointers to the outside
		tpgNodeBegin.MoveTo(pNodeCurr->GetBeginPos(), TPG_LEFT);
		tpgEnd.MoveTo(pNodeCurr->GetEndPos(), TPG_RIGHT);

		// Remove the begin pos
		fBeginEdge = pNodeCurr->GetBeginPos()->IsEdgeScope();

		hr = Remove(pNodeCurr->GetBeginPos());

		if(hr)
		{
			goto Cleanup;
		}

		// Delete the character and send the notification
		{
			CTreeNode* pNodeNotifyBegin = pNodeParent;
			long cpNotifyBegin = tpgNodeBegin.GetCp();

			// If we removed something from inside of an inclusion, find the bottom
			// of that inclusion to send the notification to
			if(!fBeginEdge)
			{
				CTreePos* ptpBeginTest = tpgNodeBegin.AdjacentTreePos(TPG_RIGHT);

				while(ptpBeginTest->IsBeginNode() && !ptpBeginTest->IsEdgeScope())
				{
					pNodeNotifyBegin = ptpBeginTest->Branch();
					ptpBeginTest = ptpBeginTest->NextTreePos();
				}
			}

			RemoveNodeChars(this, cpNotifyBegin, 1, pNodeNotifyBegin);
		}


		// Remove the end pos
		{
			CTreeNode* pNodeNotifyEnd;
			long cchNotifyEnd;
			long cpNotifyEnd;

			hr = Remove(pNodeCurr->GetEndPos());

			if(hr)
			{
				goto Cleanup;
			}

			pNodeNotifyEnd = pNodeParent;
			cpNotifyEnd = tpgEnd.GetCp();
			cchNotifyEnd = 1;

			// If this is the last node in the content chain, there must be
			// an inclusion here.  Remove it.
			if(!pNodeNext)
			{
				long cchRemove;

				hr = CloseInclusion(&tpgEnd, &cchRemove);
				if(hr)
				{
					goto Cleanup;
				}

				cpNotifyEnd -= cchRemove/2;
				cchNotifyEnd += cchRemove;

				pNodeNotifyEnd = tpgEnd.Branch();
			}

			// Delete the chracters and send the  notification

			// If we removed something from inside of an inclusion, find the bottom
			// of that inclusion to send the notification to
			if(pNodeNext)
			{
				CTreePos* ptpEndTest = tpgEnd.AdjacentTreePos(TPG_LEFT);

				while(ptpEndTest->IsEndNode() && !ptpEndTest->IsEdgeScope())
				{
					pNodeNotifyEnd = ptpEndTest->Branch();
					ptpEndTest = ptpEndTest->PreviousTreePos();
				}
			}

			RemoveNodeChars(this, cpNotifyEnd, cchNotifyEnd, pNodeNotifyEnd);
		}

		// Kill the node
		pNodeCurr->PrivateExitTree();
		pNodeCurr = NULL;
	}


	// Release element from the tree
	{
		if(fDelayRelease)
		{
			pElementRemove->AddRef();
		}

		pElementRemove->__pNodeFirstBranch = NULL;

		pElementRemove->DelMarkupPtr();
		pElementRemove->PrivateExitTree(this);
	}

	// Clear caches, etc
	UpdateMarkupTreeVersion();

	if(tpgElementBegin != tpgEnd)
	{
		CTreePos* ptpLeft = tpgElementBegin.AdjacentTreePos(TPG_RIGHT);
		CTreePos* ptpRight = tpgEnd.AdjacentTreePos(TPG_LEFT);

		hr = RangeAffected(ptpLeft, ptpRight);
	}

	tpgEnd.UnPosition();
	tpgElementBegin.UnPosition();

	// Send the ElementsDeleted notification
	{
		CNotification nf;

		nf.ElementsDeleted(siStart, 1);

		Notify(nf);
	}

	//Undo.CreateAndSubmit();

	// Send the ExitTreeSc notification and do the delay release
	if(fExitTreeSc)
	{
		Assert(fDelayRelease);
		CNotification nf;

		nf.ElementExittree2(pElementRemove);
		pElementRemove->Notify(&nf);
	}

	if(fDelayRelease)
	{
		// Release the element
		pElementRemove->Release();
	}

Cleanup:
	RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//    Member:   InsertElement
//
//  Synopsis:   Inserts an element into the tree.  The element must not already
//              be in the tree when this is called.
//
// Arguments:
//      pElementInsertThis  -   The element to insert
//      ptpgBegin           -   Element will begin here.  Must be a valid gap.
//                              IN only
//      ptpgEnd             -   Element will end here. Must be a valid gap.
//                              IN only
//-----------------------------------------------------------------------------
static void InsertNodeChars(CMarkup* pMarkup, long cp, long cch, CTreeNode* pNode)
{
    CNotification nf;

    Assert(cp>=0 && cp<pMarkup->GetTextLength());

    nf.CharsAdded(cp, cch, pNode);

    CTxtPtr(pMarkup, cp).InsertRepeatingChar(cch, WCH_NODE);

    pMarkup->Notify(&nf);
}

HRESULT CMarkup::InsertElementInternal(
			CElement*		pElementInsertThis,
			CTreePosGap*	ptpgBegin,
			CTreePosGap*	ptpgEnd,
			DWORD			dwFlags)
{
	HRESULT		hr = S_OK;
	CDocument*  pDoc = Doc();
	BOOL		fDOMOperation = dwFlags&MUS_DOMOPERATION;
	CTreePosGap	tpgCurr(TPG_LEFT);
	CTreeNode*	pNodeNew = NULL;
	CTreeNode*	pNodeLast = NULL;
	CTreeNode*	pNodeParent = NULL;
	CTreeNode*	pNodeAboveEnd;
	BOOL		fLast = FALSE;
	CTreePosGap	tpgBegin(TPG_LEFT), tpgEnd(TPG_RIGHT);
	//CInsertElementUndo Undo(this, dwFlags);

	Assert(!HasUnembeddedPointers());

	EnsureTotalOrder(ptpgBegin, ptpgEnd);


	// BUGBUG (EricVas) : Make insertion of unscoped element amortize notifications

	//
	// Element to insert must be specified and not in any tree
	Assert(pElementInsertThis && !pElementInsertThis->GetFirstBranch());

	// The gaps specifying where to put the element must be specified and
	// in valid locations in the splay tree.  They must also both be in
	// the same tree (this tree).
	Assert(ptpgBegin && ptpgBegin->IsValid());
	Assert(ptpgEnd && ptpgEnd->IsValid());
	Assert(ptpgBegin->AttachedTreePos()->IsInMarkup(this));
	Assert(ptpgEnd->AttachedTreePos()->IsInMarkup(this));

	// Push any tree pointers at the insertion points in the correct directions
	ptpgBegin->PartitionPointers(this, fDOMOperation);
	ptpgEnd->PartitionPointers(this, fDOMOperation);

	// Make sure to split any text IDs
	if(pDoc->_lLastTextID)
	{
		SplitTextID(
			ptpgBegin->AdjacentTreePos(TPG_LEFT), ptpgBegin->AdjacentTreePos(TPG_RIGHT));

		if(*ptpgBegin != *ptpgEnd)
		{
			SplitTextID(
				ptpgEnd->AdjacentTreePos(TPG_LEFT ), ptpgEnd->AdjacentTreePos(TPG_RIGHT));
		}
	}

	// Copy the endpoints to local gaps.  This way we don't muck with the
	// arguments.
	Verify(!tpgBegin.MoveTo(ptpgBegin));
	Verify(!tpgEnd.MoveTo(ptpgEnd));

	// The var tpgCurr will walk through the insertionm left to right
	tpgCurr.MoveTo(ptpgBegin);

	pNodeAboveEnd = ptpgEnd->Branch();

	// Get the parent for the new node
	pNodeParent = tpgCurr.Branch();

	while(!fLast)
	{
		// Check to see if this is the last node we will have to insert
		if(SearchBranchForNodeInStory(pNodeAboveEnd, pNodeParent))
		{
			fLast = TRUE;
		}

		// Create the new node and hook it into the context list
		pNodeNew = new CTreeNode(pNodeParent, pElementInsertThis);

		if(!pNodeNew)
		{
			hr = E_OUTOFMEMORY;
			goto Cleanup;
		}


		if(!pNodeLast)
		{
			pElementInsertThis->__pNodeFirstBranch = pNodeNew;

			pElementInsertThis->SetMarkupPtr(this);
			pElementInsertThis->PrivateEnterTree();
		}

		// Insert the start node
		{
			CTreePos* ptpNew;
			CTreeNode* pNodeNotify = pNodeParent;

			Assert(pNodeNew->GetBeginPos()->IsUninit());

			ptpNew = pNodeNew->InitBeginPos(!pNodeLast);

			Verify(!Insert(ptpNew, &tpgCurr));

			// move tpgCurr to the right looking for the bottom of 
			// the inclusion, so that we know which node to send
			// the notification to
			if(pNodeLast)
			{
				tpgCurr.MoveRight();
				Assert(tpgCurr.AttachedTreePos() == ptpNew);

				tpgCurr.MoveRight();

				while(tpgCurr.AttachedTreePos()->IsBeginNode() 
					&& !tpgCurr.AttachedTreePos()->IsEdgeScope())
				{
					pNodeNotify = tpgCurr.AttachedTreePos()->Branch();
					tpgCurr.MoveRight();
				}
			}

			// The element is now in the tree so add a ref for the
			// tree
			pNodeNew->PrivateEnterTree();

			InsertNodeChars(this, ptpNew->GetCp(), 1, pNodeNotify);
		}

		// Insert the end node
		{
			CTreePos*	ptpNew;
			long		cpInsert;
			long		cchInsert = 0;
			CTreeNode*	pNodeNotify = NULL;

			// Create or set the insertion point
			if(fLast)
			{
				// Split the branch to create a spot for our end node pos
				pNodeNotify = pNodeAboveEnd;

				cpInsert = tpgEnd.GetCp();

				hr = CreateInclusion(
					pNodeParent, 
					&tpgEnd, 
					&tpgCurr, 
					&cchInsert,
					pNodeAboveEnd);

				if(hr)
				{
					goto Cleanup;
				}
			}
			else
			{
				CTreePosGap tpgNotify(TPG_RIGHT);

				Verify(!tpgCurr.MoveTo(pNodeParent->GetEndPos(), TPG_LEFT));

				cpInsert = tpgCurr.GetCp();

				// Move tpgNotify to the left to find the lowest 
				// node in the inclusion.
				Verify(!tpgNotify.MoveTo(&tpgCurr));
				Assert(tpgNotify.AttachedTreePos() == pNodeParent->GetEndPos());

				pNodeNotify = pNodeParent;
				tpgNotify.MoveLeft();

				while(tpgNotify.AttachedTreePos()->IsEndNode()
					&& !tpgNotify.AttachedTreePos()->IsEdgeScope())
				{
					pNodeNotify = tpgNotify.AttachedTreePos()->Branch();
					tpgNotify.MoveLeft();
				}
			}


			Assert(pNodeNew->GetEndPos()->IsUninit());

			ptpNew = pNodeNew->InitEndPos(fLast);

			Verify(!Insert(ptpNew, &tpgCurr));

			cchInsert++;

			Verify(!tpgCurr.MoveRight());
			Assert(tpgCurr.AdjacentTreePos(TPG_LEFT) == ptpNew);

			InsertNodeChars(this, cpInsert, cchInsert, pNodeNotify);
		}

		// Reparent direct children to the new node
		hr = ReparentDirectChildren(pNodeNew);

		if(hr)
		{
			goto Cleanup;
		}

		if(!fLast)
		{
			// if we are here, it implies that we are
			// inserting the end after the end of pNodeParent
			// This _can't_ happen if pNodeParent is a ped.
			Assert(!pNodeParent->IsRoot());

			// Set up for the next new node
			pNodeLast = pNodeNew;

			// This sets our beginning insertion point
			// for the next node.  What we are doing here
			// is essentially threading the new element
			// through any inclusions.
			Verify(!tpgCurr.MoveRight());

			Assert(tpgCurr.AttachedTreePos()->IsEndNode() &&
				tpgCurr.AttachedTreePos()->Branch()==pNodeParent);

			// pNodeParent is done so its parent is the new pNodeParent
			if(tpgCurr.AttachedTreePos()->IsEdgeScope())
			{
				pNodeParent = pNodeParent->Parent();
			}
			else
			{
				// Find the next branch of pNodeParent
				CElement* pElementParent = pNodeParent->Element();
				CTreeNode* pNodeCurr;

				do
				{
					Verify(!tpgCurr.MoveRight());

					Assert(tpgCurr.AttachedTreePos()->IsNode());

					pNodeCurr = tpgCurr.AttachedTreePos()->Branch();
				} while(pNodeCurr->Element() != pElementParent);

				Assert(tpgCurr.AttachedTreePos()->IsBeginNode());

				pNodeParent = pNodeCurr;
			}

		}
	}

	// Tell the element that it is now in the tree
	{
		CNotification nf;

		nf.ElementEntertree(pElementInsertThis);
		pElementInsertThis->Notify(&nf);
	}

	if(tpgBegin != tpgEnd)
	{
		CTreePos* ptpLeft = tpgBegin.AdjacentTreePos(TPG_RIGHT);
		CTreePos* ptpRight = tpgEnd.AdjacentTreePos(TPG_LEFT);

		hr = RangeAffected(ptpLeft, ptpRight);
	}

	// Send the element added notification
	{
		CTreePos* ptpBegin;
		CNotification nf;

		pElementInsertThis->GetTreeExtent(&ptpBegin, NULL);
		Assert(ptpBegin);

		nf.ElementsAdded(ptpBegin->SourceIndex(), 1);
		Notify(nf);
	}

	UpdateMarkupTreeVersion();

	//Undo.CreateAndSubmit();

Cleanup:
	RRETURN(hr);
}

BOOL AreDifferentScriptIDs(SCRIPT_ID* psidFirst, SCRIPT_ID sidSecond)
{
    if(*psidFirst==sidSecond || sidSecond==sidMerge
        || *psidFirst==sidAsciiLatin&&sidSecond==sidAsciiSym)
    {
        return FALSE;
    }

    if(*psidFirst==sidMerge || *psidFirst==sidAsciiSym&&sidSecond==sidAsciiLatin)
    {
        *psidFirst = sidSecond;
        return FALSE;
    }

    return TRUE;
}

//+----------------------------------------------------------------------------
//
//    Member:   InsertText
//
//  Synopsis:   Inserts a chunk of text into the tree before the given TreePos
//              This does not send notifications.
//
//-----------------------------------------------------------------------------
// BUGBUG: we can probably share code here between this function and
// the slow loop of CHtmRootParseCtx::AddText
HRESULT CMarkup::InsertTextInternal(
		CTreePos*		ptpAfterInsert, 
		const TCHAR*	pch,
		long			cch,
		DWORD			dwFlags)
{
	HRESULT			hr = S_OK;
	BOOL			fDOMOperation = dwFlags&MUS_DOMOPERATION;
	CTreePosGap		tpg(TPG_RIGHT);
	CTreePos*		ptpLeft;
	const TCHAR*	pchCurr;
	const TCHAR*	pchStart;
	long			cchLeft;
	SCRIPT_ID		sidLast = sidDefault;
	long			lTextIDCurr = 0;
	CNotification	nf;
	long			cpStart = ptpAfterInsert->GetCp();
	DEBUG_ONLY(long ichAfterInsert = 0);

	Assert(!HasUnembeddedPointers());

	Assert(ptpAfterInsert);
	Assert(cch >= 0);

	if(cch <= 0)
	{
		goto Cleanup;
	}

	UpdateMarkupContentsVersion();

	Assert(ptpAfterInsert);
	Assert(cch >= 0);

	// Partition the pointers to ensure gravity
	tpg.MoveTo(ptpAfterInsert, TPG_LEFT);
	tpg.PartitionPointers(this, fDOMOperation);
	ptpAfterInsert = tpg.AdjacentTreePos(TPG_RIGHT);


	nf.CharsAdded(cpStart, cch, tpg.Branch());

	// Collect some information that we'll need later

	// Get info about the previous pos
	ptpLeft = tpg.AdjacentTreePos(TPG_LEFT);
	Assert(ptpLeft);

	if(ptpLeft->IsText())
	{
		sidLast = ptpLeft->Sid();
	}

	pchCurr = pchStart = pch;
	cchLeft = cch;

	// Look around for a TextID to merge with
	{
		CTreePos* ptp;

		for(ptp=ptpLeft; !ptp->IsNode(); ptp=ptp->PreviousTreePos())
		{
			if(ptp->IsText())
			{
				lTextIDCurr = ptp->TextID();
				break;
			}
		}

		if(!lTextIDCurr)
		{
			for(ptp=ptpAfterInsert; !ptp->IsNode(); ptp=ptp->NextTreePos())
			{
				if(ptp->IsText())
				{
					lTextIDCurr = ptp->TextID();
					break;
				}
			}
		}
	}

	// Break up the text into chunks of compatible SIDs.
	while(cchLeft)
	{
		SCRIPT_ID	sidChunk, sidCurr;
		ULONG		cchChunk;
		TCHAR		chCurr = *pchCurr;

		sidChunk = ScriptIDFromCh(chCurr);

		// Find the end of this chunk of characters with the same sid
		while(cchLeft)
		{
			chCurr = *pchCurr;

			// If we find an illeagal character, then simply compute a new,
			// correct, buffer and call recursively.
			if(chCurr==0 || !IsValidWideChar(chCurr))
			{
				TCHAR* pch2;
				long i;

				AssertSz(0, "Bad char during insert");

				pch2 = new TCHAR[cch];

				if(!pch2)
				{
					hr = E_OUTOFMEMORY;
					goto Cleanup;
				}

				for(i=0; i<cch; i++)
				{
					TCHAR ch = pch[i];

					if(ch==0 || !IsValidWideChar(ch))
					{
						ch = _T('?');
					}

					pch2[i] = ch;
				}

				hr = InsertTextInternal(ptpAfterInsert, pch2, cch, dwFlags);

				delete pch2;

				goto Cleanup;
			}

			sidCurr = ScriptIDFromCh(chCurr);

			if(AreDifferentScriptIDs(&sidChunk, sidCurr))
			{
				break;
			}

			pchCurr++;
			cchLeft--;
		}

		cchChunk = pchCurr - pchStart;

		// If this is the first chunk, attempt to merge with a text pos to the left
		if(ptpLeft->IsText() && !AreDifferentScriptIDs(&sidChunk, sidLast))
		{
			Assert(pchStart == pch);
			ptpLeft->ChangeCch(cchChunk);
			ptpLeft->DataThis()->t._sid = sidChunk;
		}
		else
		{
			// If we can merge with the text pos after insertion, do so
			if(cchLeft==0 && 
				ptpAfterInsert->IsText() &&
				!AreDifferentScriptIDs(&sidChunk, ptpAfterInsert->Sid()))
			{
				DEBUG_ONLY(ichAfterInsert = cchChunk);
				ptpAfterInsert->ChangeCch(cchChunk);
				ptpAfterInsert->DataThis()->t._sid = sidChunk;
			}
			else
			{
				if(sidChunk == sidMerge)
				{
					sidChunk = sidDefault;
				}

				// No merging - we actually have to create a new text pos
				ptpLeft = NewTextPos(cchChunk, sidChunk, lTextIDCurr);
				if(!ptpLeft)
				{
					hr = E_OUTOFMEMORY;
					goto Cleanup;
				}

				hr = Insert(ptpLeft, &tpg);
				if(hr)
				{
					goto Cleanup;
				}

				Assert(tpg.AttachedTreePos() == ptpAfterInsert);
				Assert(tpg.AdjacentTreePos(TPG_LEFT) == ptpLeft);
			}
		}

		pchStart = pchCurr;
		sidLast = sidChunk;
	}

	Assert(cch == pchCurr-pch);
	Assert(cch == ptpAfterInsert->GetCp()+ichAfterInsert-cpStart);

	// Now actually stuff the characters into the story
	CTxtPtr(this, cpStart).InsertRange(cch, pch);

	// Send the notification
	Notify(&nf);

	/*Undo.CreateAndSubmit(); wlw note*/

Cleanup:
	RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     FastElemTextSet
//
//  Synopsis:   Fast way to set the entire text of an element (e.g. INPUT).
//              Assumes that there is no overlapping.
//
//-----------------------------------------------------------------------------
HRESULT CMarkup::FastElemTextSet(CElement* pElem, const TCHAR* psz, int c, BOOL fAsciiOnly)
{
	HRESULT		hr = S_OK;
	CTreePos*	ptpBegin;
	CTreePos*	ptpEnd;
	long		cpBegin, cpEnd;

	// make sure pointers are embedded
	hr = EmbedPointers();
	if(hr)
	{
		goto Cleanup;
	}

	// Assert that there is no overlapping into this element
	Assert(!pElem->GetFirstBranch()->NextBranch());

	pElem->GetTreeExtent(&ptpBegin, &ptpEnd);
	Assert(ptpBegin && ptpEnd);

	cpBegin = ptpBegin->GetCp() + 1;
	cpEnd = ptpEnd->GetCp();

	if(cpBegin!=cpEnd || !fAsciiOnly)
	{
		CTreePosGap tpgBegin(TPG_LEFT), tpgEnd(TPG_RIGHT);
		Verify(!tpgBegin.MoveTo(ptpBegin, TPG_RIGHT));
		Verify(!tpgEnd.MoveTo(ptpEnd, TPG_LEFT));

		hr = SpliceTreeInternal(&tpgBegin, &tpgEnd);
		if(hr)
		{
			goto Cleanup;
		}
#ifdef _DEBUG
		{
			CTreePos* ptpEndAfter;
			pElem->GetTreeExtent(NULL, &ptpEndAfter);
			Assert(ptpEndAfter == ptpEnd);
		}
#endif

		hr = InsertTextInternal(ptpEnd, psz, c, NULL);
		if(hr)
		{
			goto Cleanup;
		}
	}
	else if(c)
	{
		CTreePos*		ptpText = ptpEnd;
		CNotification	nf;
		//CInsertSpliceUndo Undo(Doc());

		//Undo.Init(this, NULL);

		UpdateMarkupContentsVersion();

		//Undo.SetData(cpBegin, cpBegin+c);

		// Build a notification (do we HAVE to do this?!?)
		nf.CharsAdded(cpBegin, c, pElem->GetFirstBranch());

		ptpText = NewTextPos(c);
		hr = Insert(ptpText, ptpBegin, FALSE);
		if(hr)
		{
			goto Cleanup;
		}

		// Insert the text
		Verify(CTxtPtr(this, cpBegin).InsertRange(c, psz) == c);

		Notify(&nf);

		//Undo.CreateAndSubmit();
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::ReparentDirectChildren(CTreeNode* pNodeParentNew, CTreePosGap* ptpgStart, CTreePosGap* ptpgEnd)
{
	CTreePos *ptpCurr, *ptpEnd;

	// If the start/end is specified, make sure they're sane
	Assert(!ptpgStart || ptpgStart->GetAttachedMarkup()==this);
	Assert(!ptpgEnd || ptpgEnd->GetAttachedMarkup()==this);

	// If the position where we should start and end the reparent operation
	// are specified, then go there, otherwise start/end at where the new
	// parent starts/ends.
	ptpCurr = ptpgStart
		? ptpgStart->AdjacentTreePos(TPG_RIGHT) : pNodeParentNew->GetBeginPos()->NextTreePos();

	ptpEnd = ptpgEnd
		? ptpgEnd->AdjacentTreePos(TPG_RIGHT) : pNodeParentNew->GetEndPos();

	// Make sure the start is to the left of the end
	Assert(ptpCurr->InternalCompare(ptpEnd) <= 0);

	// Loop and reparent.  Get to just direct children by skipping
	// over direct children.
	while(ptpCurr != ptpEnd)
	{
		switch(ptpCurr->Type())
		{
		case CTreePos::NodeEnd :
			// We will (most) never find an end pos because we'll skip over
			// them when we encounter the begin pos for that node.
			AssertSz(0, "Found an end pos during reparent children");

			break;

		case CTreePos::NodeBeg :
			ptpCurr->Branch()->SetParent( pNodeParentNew );
			// Skip over everything under this node
			ptpCurr = ptpCurr->Branch()->GetEndPos();

			break;

		default:
			Assert(!ptpCurr->IsUninit());
			break;
		}

		ptpCurr = ptpCurr->NextTreePos();
	}

	RRETURN(S_OK);
}

HRESULT CMarkup::CreateInclusion(
		 CTreeNode*		pNodeStop,
		 CTreePosGap*	ptpgLocation,
		 CTreePosGap*	ptpgInclusion,
		 long*			pcchNeeded,
		 CTreeNode*		pNodeAboveLocation/*=NULL*/,
		 BOOL			fFullReparent/*=TRUE*/,
		 CTreeNode**	ppNodeLastAdded/*=NULL*/)
{
	HRESULT		hr = S_OK;
	CTreePosGap	tpgInsert(TPG_LEFT);
	CTreeNode*	pNodeCurr;
	CTreeNode*	pNodeNew=NULL;
	CTreeNode*	pNodeLastNew = NULL;
	long		cpInsertNodeChars=-1;
	BOOL		fInsertChars=!pcchNeeded;
	long		cchNeededAlternate;

	// The passed in gap must be specified and positioned at a valid point
	// in this tree.
	Assert(ptpgLocation);
	Assert(ptpgLocation->IsValid());
	Assert(ptpgLocation->GetAttachedMarkup() == this);

	// The node up to which we should split must be specified and in this tree.
	// Note that this node may not actually curently influence the gap in the
	// tree specified by ptpgLocation.  This is so because the tree may be in
	// an unstable state.
	Assert(pNodeStop);
	Assert(pNodeStop->GetMarkup() == this);

	// If caller wants us to put in the node characters, then point the
	// char counter at a local variable.
	if(!pcchNeeded)
	{
		pcchNeeded = &cchNeededAlternate;
	}

	// Initially, there are no node chars needed
	*pcchNeeded = 0;

	// cpInsertNodeChars is where we may need to insert node chars
	if(fInsertChars)
	{
		cpInsertNodeChars = ptpgLocation->GetCp();
	}

	// The pNodeAboveLocation, if specified, should be assumed to be the first node
	// which influences the gap specified by ptpgLocation.  Because the tree may be
	// in an unstable state, this node may not be visible to ptpgLocation.
	//
	// In any case, here we compute the node where we start the creation of the
	// inclusion.
	pNodeCurr = pNodeAboveLocation ? pNodeAboveLocation : ptpgLocation->Branch();

	Assert(pNodeCurr);

	// Now that we have a place to start with in the tree, make sure the stop node
	// is somewhere along the parent chain from the start node (now pNodeCurr).
	Assert( pNodeCurr->SearchBranchToRootForScope(pNodeStop->Element() ) == pNodeStop);

	// Set up the spot to insert
	Verify(!tpgInsert.MoveTo(ptpgLocation));

	while(pNodeCurr != pNodeStop)
	{
		CTreePos* ptpNew;

		// Create the new node (proxy to the current element)
		pNodeNew = new CTreeNode(pNodeCurr->Parent(), pNodeCurr->Element());

		if(!pNodeNew)
		{
			hr = E_OUTOFMEMORY;
			goto Cleanup;
		}

		// Insert the end pos of the new node to be just before the end pos of
		// the node we cloned it from.  It inherits edge scopeness from the end
		// of the current node.
		ptpNew = pNodeNew->InitEndPos(pNodeCurr->GetEndPos()->IsEdgeScope());

		Verify(!Insert(ptpNew, pNodeCurr->GetEndPos(), TRUE));

		// Based on the attach preference of the tree gap which originally
		// specified where to create the inclusion, push that gap around.
		if(ptpgLocation->AttachedTreePos() == pNodeCurr->GetEndPos())
		{
			Assert(ptpgLocation->AttachDirection() == TPG_RIGHT);

			ptpgLocation->MoveTo(ptpNew, TPG_LEFT);
		}

		// Insert the begin pos to the right of the point of the inclusion.  It is
		// automatically not an edge.
		ptpNew = pNodeNew->InitBeginPos(FALSE);

		Verify(!Insert(ptpNew, tpgInsert.AttachedTreePos(), FALSE));

		// Mark this node as being in the markup
		pNodeNew->PrivateEnterTree();

		// Move the end pos of the current node to be to the left of the
		// location of the inclusion.
		Verify(!Move(pNodeCurr->GetEndPos(), tpgInsert.AttachedTreePos(), FALSE));

		// The end pos of the current node is no longer an edge
		pNodeCurr->GetEndPos()->SetScopeFlags(FALSE);

		// Move over the end of end of the current node to get back into the
		// location of the inclusion.
		Verify(!tpgInsert.Move(TPG_RIGHT));

		Assert(tpgInsert.AdjacentTreePos(TPG_LEFT) == pNodeCurr->GetEndPos());

		// Reparent the direct children in the range affected
		if(fFullReparent)
		{
			hr = ReparentDirectChildren(pNodeNew);

			if(hr)
			{
				goto Cleanup;
			}
		}
		else if(pNodeLastNew)
		{
			pNodeLastNew->SetParent(pNodeNew);
		}

		// Record the number of WCH_NODE characters needed at
		// the inclusion point
		*pcchNeeded += 2;

		// Finally, set up for the next time around
		pNodeLastNew = pNodeNew;
		pNodeCurr = pNodeCurr->Parent();
	}

	if(ppNodeLastAdded)
	{
		*ppNodeLastAdded = pNodeLastNew;
	}

	if(ptpgInclusion)
	{
		Verify(!ptpgInclusion->MoveTo(&tpgInsert));
	}

	// If requested, insert need characters here
	if(fInsertChars && *pcchNeeded)
	{
		CNotification nf;

		Assert(*pcchNeeded%2 == 0);

		nf.CharsAdded(cpInsertNodeChars, *pcchNeeded, tpgInsert.Branch());

		CTxtPtr(this, cpInsertNodeChars).InsertRepeatingChar(*pcchNeeded, WCH_NODE);

		Notify(&nf);
	}

Cleanup:

	RRETURN(hr);
}

HRESULT CMarkup::CloseInclusion(CTreePosGap* ptpgMiddle, long* pcchRemove)
{
	HRESULT		hr = S_OK;
	BOOL		fRemoveChars = ! pcchRemove;
	long		cchRemoveAlternate;
	CTreePosGap	tpgMiddle(TPG_LEFT);

	// Note: tpgMiddle is bound to the pos which is to the left of the
	// gap because it will move to the left as we delete nodes.  This
	// way, it stays out of the way (to the left) of the removals.

	// If caller wants us to remove the node characters, then point the
	// char counter at a local variable.
	if(!pcchRemove)
	{
		pcchRemove = &cchRemoveAlternate;
	}

	// No chars to delete ... yet
	*pcchRemove = 0;

	// The middle gap argument must be specified and positioned in this
	// tree.
	Assert(ptpgMiddle);
	Assert(ptpgMiddle->IsPositioned());
	Assert(ptpgMiddle->GetAttachedMarkup() == this);

	// Copy the incomming argument to our local gap, nd unposition the argument.
	// We'll reposition it later, after we have closed the inclusion.
	Verify(!tpgMiddle.MoveTo(ptpgMiddle));

	ptpgMiddle->UnPosition();

	for(;;)
	{
		CTreePos *ptpRightMiddle, *ptpLeftMiddle;
		CTreeNode *pNodeBefore, *pNodeAfter;

		// Get the pos to the left of the gap which travels accross
		// the splays.
		ptpLeftMiddle = tpgMiddle.AttachedTreePos();

		// We know we're not in an inclusion anymore when the pos to the
		// left is not a ending edge.
		//
		// Note, we don't have to check to tree pointers or text because
		// they must never be seen in inclusions.
		if(!ptpLeftMiddle->IsEndNode() || ptpLeftMiddle->IsEdgeScope())
		{
			break;
		}

		// Get the tree pos to the right of the one which is to the
		// left of the gap.  This will, then, be to the right of the gap.
		ptpRightMiddle = ptpLeftMiddle->NextTreePos();

		// Check the right edge for non matching ptp's.  We can
		// have a non edge end node pos on the left but not a non edge begin
		// node ont he right during certain circumstances in splice.
		if(!ptpRightMiddle->IsBeginNode() || ptpRightMiddle->IsEdgeScope())
		{
			break;
		}

		// Move the gap to the left (<--) by one (not right).
		Verify(!tpgMiddle.MoveLeft());

		// If we've gotten this far, then there damn well better be
		// an inclusion here.
		//
		// Thus, we must have a non edge beginning to our right.
		Assert(ptpRightMiddle->IsBeginNode() && !ptpRightMiddle->IsEdgeScope());

		// Get the nodes associated with either side of the gap in question.
		pNodeBefore = ptpLeftMiddle->Branch();
		pNodeAfter = ptpRightMiddle->Branch();

		// To be in an inclusion, the two adjacent nodes must refer to
		// the same element.  
		Assert(pNodeBefore->Element() == pNodeAfter->Element());

		// Make sure the inclusion does not refer to the root!
		Assert(!pNodeBefore->Element()->IsRoot());

		// Move the end pos of the node left of the gap to be just before
		// the end pos of the node right of the gap.
		//
		// This way, the node left of the gap will subsume "ownership"
		// of all the stuff which is under the node right of the gap.
		Verify(!Move( pNodeBefore->GetEndPos(), pNodeAfter->GetEndPos(), TRUE));

		// Now that the node right of the gap is going away, the end pos
		// of the node left of the gap needs to reflect the same edge status
		// of the end pos going away.  Because there may be yet another
		// proxy for this element to the right, we can't assume that it
		// is an edge.
		pNodeBefore->GetEndPos()->SetScopeFlags(pNodeAfter->GetEndPos()->IsEdgeScope());

		// Here we remove the node right of the gap (and its pos's)
		Verify(!Remove(pNodeAfter->GetBeginPos()));

		Verify(!Remove(pNodeAfter->GetEndPos()));

		// Two pos's gone .. two chars must go
		*pcchRemove += 2;

		// Reparent the children that were pointing to the node going away.
		//
		// Now that the node to the right of the gap is gone, we must
		// make sure that all the nodes which pointed to the node which
		// we just removed now point to the node which remains.
		//
		// We start the reparent starting at the current location of the gap
		// and extending until the remaining node goes out of scope.
		hr = ReparentDirectChildren(pNodeBefore, &tpgMiddle);

		if(hr)
		{
			goto Cleanup;
		}

		// Now, kill the node we just removed.  No one in this tree better be
		// pointing to it after the reparent operation.

		// pNodeAfter is not usable after this
		pNodeAfter->PrivateExitTree();
		pNodeAfter = NULL;
	}

	// Put the incomming argument back into the tree
	Verify(!ptpgMiddle->MoveTo(&tpgMiddle));

	// If requested, remove WCH_NODE chars here
	if(fRemoveChars && *pcchRemove)
	{
		CNotification nf;
		long cp = ptpgMiddle->GetCp();

		Assert(*pcchRemove%2 == 0);

		nf.CharsDeleted(cp, *pcchRemove, tpgMiddle.Branch());

		CTxtPtr(this, cp).DeleteRange(*pcchRemove);

		Notify(&nf);
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::RangeAffected(CTreePos* ptpLeft, CTreePos* ptpRight)
{
	HRESULT			hr;
	CNotification	nf;
	long			cp, cch;
	CTreeNode*		pNodeNotify = NULL;
	CTreePos*		ptpStart=ptpLeft, *ptpAfterLast;

	// clear all of the caches
	hr = ClearCaches(ptpLeft, ptpRight);
	if(hr)
	{
		goto Cleanup;
	}

	Assert(ptpLeft && ptpRight);
	ptpAfterLast = ptpRight->NextTreePos();

	// Send the CharsResize notification
	cp = ptpStart->GetCp();
	cch = 0;

	Assert(cp >= 0);

	// Note: notifications for the WCH_NODE characters go
	// to intersting places.  For WCH_NODE characters inside
	// of an inclusion, the notifations go to nodes at the
	// bottom of the inclusion.  Also, for edges, notifications
	// go to the parent of the element, not the element itself.
	// This way, noscope elements to not get any notifications
	// in this loop.
	while(ptpStart != ptpAfterLast)
	{
		if(ptpStart->IsNode())
		{
			// if we are entering an inclusion
			// then remember the first node as the
			// one to send the left half of 
			// the notification to
			if(!pNodeNotify && ptpStart->IsEndNode() && !ptpStart->IsEdgeScope())
			{
				pNodeNotify = ptpStart->Branch();
			}

			if(ptpStart->IsBeginNode())
			{
				cch++;
			}

			// send a notification if we hit the edge
			// of a layout
			if(cch && ptpStart->IsEdgeScope() && ptpStart->Branch()->HasLayout())
			{
				if(!pNodeNotify)
				{
					pNodeNotify = ptpStart->IsBeginNode() 
						? ptpStart->Branch()->Parent() : ptpStart->Branch();
				}

				Assert(!pNodeNotify->Element()->IsNoScope());

				nf.CharsResize(cp, cch, pNodeNotify);
				Notify(nf);

				cp += cch;
				cch = 0;
			}

			// If we hit an edge, clear pNodeNotify.  Either we sent
			// a notification or we didn't. Either way, we are done
			// with this inclusion so we don't have to remember pNodeNotify
			if(ptpStart->IsEdgeScope())
			{
				pNodeNotify = NULL;
			}

			if(ptpStart->IsEndNode())
			{
				cch++;
			}
		}
		else if(ptpStart->IsText())
		{
			cch += ptpStart->Cch();
		}

		ptpStart = ptpStart->NextTreePos();
	}

	// Finish off any notification left over
	if(cch)
	{
		if(!pNodeNotify)
		{
			CTreePosGap tpg(ptpStart, TPG_LEFT);
			pNodeNotify = tpg.Branch();
		}

		nf.CharsResize(cp, cch, pNodeNotify);
		Notify(nf);
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::ClearCaches(CTreePos* ptpStart, CTreePos* ptpFinish)
{
	CTreePos *ptpCurr, *ptpAfterFinish=ptpFinish->NextTreePos();

	for(ptpCurr=ptpStart; ptpCurr!=ptpAfterFinish; ptpCurr=ptpCurr->NextTreePos())
	{
		if(ptpCurr->IsBeginNode())
		{
			ptpCurr->Branch()->VoidCachedInfo();

			if(ptpCurr->IsEdgeScope())
			{
				CElement* pElementCur = ptpCurr->Branch()->Element();

				// Clear caches on the slave markup
				if(pElementCur->HasSlaveMarkupPtr())
				{
					CTreePos *ptpStartSlave, *ptpFinishSlave;
					CMarkup* pMarkupSlave = pElementCur->GetSlaveMarkupPtr();

					pMarkupSlave->GetContentTreeExtent(&ptpStartSlave, &ptpFinishSlave);
					pMarkupSlave->ClearCaches(ptpStartSlave, ptpFinishSlave);
				}
			}
		}
	}

	return S_OK;
}

HRESULT CMarkup::ClearRunCaches(DWORD dwFlags, CElement* pElement)
{
	CTreePos*	ptpStart = NULL;
	CTreePos*	ptpEnd;
	BOOL		fClearAllFormats = dwFlags&ELEMCHNG_CLEARCACHES;

	Assert(pElement);
	pElement->GetTreeExtent(&ptpStart, &ptpEnd);

	if(ptpStart)
	{
		CTreePos* ptpAfterFinish = ptpEnd->NextTreePos();

		for(; ptpStart!=ptpAfterFinish; ptpStart=ptpStart->NextTreePos())
		{
			if(ptpStart->IsBeginNode())
			{
				CTreeNode*	pNodeCur			= ptpStart->Branch();
				CElement *	pElementCur			= pNodeCur->Element();
				BOOL		fNotifyFormatChange	= FALSE;

				if(fClearAllFormats)
				{
					// clear the formats on the node
					pNodeCur->VoidCachedInfo();
					fNotifyFormatChange = TRUE;
				}
				else if(pElement==pElementCur || pElementCur->_fInheritFF)
				{
					pNodeCur->VoidFancyFormat();
					fNotifyFormatChange = pElement == pElementCur;
				}

				// if the node comming into scope is a new element
				// notify the element of a format cache change.
				if(fNotifyFormatChange && ptpStart->IsEdgeScope())
				{
					CLayout* pLayout = pElementCur->GetLayoutPtr();

					if(pLayout)
					{
						pLayout->OnFormatsChange(dwFlags);
					}

					// Clear caches on the slave markup
					if(pElementCur->HasSlaveMarkupPtr())
					{
						CMarkup* pMarkupSlave = pElementCur->GetSlaveMarkupPtr();
						pMarkupSlave->ClearRunCaches(dwFlags, pMarkupSlave->Root());
					}
				}
			}
		}
	}
	else if(pElement->GetFirstBranch())
	{
		// this could happen, when element's are temporarily parented to the
		// rootsite.
		pElement->GetFirstBranch()->VoidCachedInfo();
	}
	return S_OK;
}

HRESULT CMarkup::DoEmbedPointers()
{
	HRESULT hr = S_OK;

	while(_pmpFirst)
	{
		CMarkupPointer* pmp;
		CTreePos* ptpNew;

		// Remove the first pointer from the list
		pmp = _pmpFirst;

		_pmpFirst = pmp->_pmpNext;

		if(_pmpFirst)
		{
			_pmpFirst->_pmpPrev = NULL;
		}

		pmp->_pmpNext = pmp->_pmpPrev = NULL;

		Assert(pmp->_ichRef==0 || pmp->_ptpRef->IsText());

		// Consider the case where two markup pointers point in to the the
		// middle of the same text pos, where the one with the larger ich
		// occurs later in the list of unembedded pointers.  When the first
		// is encountered, it will split the text pos, leaving the second
		// with an invalid ich.
		//
		// Here we check to see if the ich is within the text pos.  If it is
		// not, then the embedding of a previous pointer must have split the
		// text pos.  In this case, we scan forward to locate the right hand
		// side of that split, and re-adjust this pointer!
		if(pmp->_ptpRef->IsText() && pmp->_ichRef>pmp->_ptpRef->Cch())
		{
			// If we are out of range, then there better very well be a pointer
			// next which did it.
			Assert(pmp->_ptpRef->NextTreePos()->IsPointer());

			while(pmp->_ichRef > pmp->_ptpRef->Cch())
			{
				Assert(pmp->_ptpRef->IsText() && pmp->_ptpRef->Cch()>0);

				pmp->_ichRef -= pmp->_ptpRef->Cch();

				do
				{
					pmp->_ptpRef = pmp->_ptpRef->NextTreePos();
				} while(!pmp->_ptpRef->IsText());
			}
		}

		// See if we have to split a text pos
		if(pmp->_ptpRef->IsText() && pmp->_ichRef<pmp->_ptpRef->Cch())
		{
			Assert(pmp->_ichRef != 0);

			hr = Split(pmp->_ptpRef, pmp->_ichRef);

			if(hr)
			{
				goto Cleanup;
			}
		}

		ptpNew = NewPointerPos(pmp, pmp->Gravity(), pmp->Cling());

		if(!ptpNew)
		{
			hr = E_OUTOFMEMORY;
			goto Cleanup;
		}

		// We should always be at the end of a text pos.
		Assert(!pmp->_ptpRef->IsText() || pmp->_ichRef==pmp->_ptpRef->Cch());

		hr = Insert(ptpNew, pmp->_ptpRef, FALSE);

		if(hr)
		{
			goto Cleanup;
		}

		pmp->_fEmbedded = TRUE;
		pmp->_ptpEmbeddedPointer = ptpNew;
		pmp->_ichRef = 0;
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::SetTextID(CTreePosGap*	ptpgStart, CTreePosGap*	ptpgEnd, long* plNewTextID)
{
	HRESULT hr = S_OK;
	long lTxtID;

	Assert(!HasUnembeddedPointers());

	EnsureTotalOrder(ptpgStart, ptpgEnd);

	Assert(ptpgStart && ptpgStart->IsPositioned() && ptpgStart->GetAttachedMarkup()==this);
	Assert(ptpgEnd && ptpgEnd->IsPositioned() && ptpgEnd->GetAttachedMarkup()==this);

	CTreePos *ptpFirst, *ptpCurr, *ptpStop;
	CDocument* pDoc = Doc();

	if(!plNewTextID)
	{
		plNewTextID = &lTxtID;
	}

	*plNewTextID = 0;

	ptpFirst = ptpgStart->AdjacentTreePos(TPG_LEFT);
	ptpCurr = ptpFirst;
	ptpStop = ptpgEnd->AdjacentTreePos(TPG_RIGHT);

	ptpgStart->UnPosition();
	ptpgEnd->UnPosition();

	SplitTextID(ptpCurr, ptpStop);

	ptpCurr = ptpCurr->NextTreePos();

	while(ptpCurr != ptpStop)
	{
		Assert(ptpCurr);

		if(ptpCurr->IsNode())
		{
			*plNewTextID = 0;
		}

		if(ptpCurr->IsText())
		{
			if(*plNewTextID == 0)
			{
				*plNewTextID = ++(pDoc->_lLastTextID);
			}

			hr = SetTextPosID(&ptpCurr, *plNewTextID);
			if(hr)
			{
				goto Cleanup;
			}
		}

		ptpCurr = ptpCurr->NextTreePos();
	}

	if(!*plNewTextID)
	{
		CTreePos* ptpNew;
		CTreePosGap tpgInsert(ptpCurr, TPG_RIGHT);

		ptpNew = NewTextPos(0, sidDefault, *plNewTextID=++(pDoc->_lLastTextID));

		hr = Insert(ptpNew, &tpgInsert);
		if(hr)
		{
			goto Cleanup;
		}
	}

	Verify(!ptpgStart->MoveTo(ptpFirst, TPG_RIGHT));
	Verify(!ptpgEnd->MoveTo(ptpStop, TPG_LEFT));

Cleanup:
	RRETURN(hr);
}

long CMarkup::GetTextID(CTreePosGap* ptpg)
{
	Assert(ptpg && ptpg->IsPositioned() && ptpg->GetAttachedMarkup()==this);

	CTreePos* ptp = ptpg->AdjacentTreePos(TPG_RIGHT);

	while(!ptp->IsNode())
	{
		if(ptp->IsText())
		{
			return ptp->TextID();
		}
		ptp = ptp->NextTreePos();
		Assert(ptp);
	}

	return -1;
}

HRESULT CMarkup::FindTextID(long lTextID, CTreePosGap* ptpgStart, CTreePosGap* ptpgEnd)
{
	Assert(ptpgStart && ptpgStart->IsPositioned() && ptpgStart->GetAttachedMarkup()==this);
	Assert(ptpgEnd);

	CTreePos *ptpLeft, *ptpRight, *ptpFound=NULL;

	ptpLeft = ptpgStart->AdjacentTreePos(TPG_LEFT);
	ptpRight = ptpgStart->AdjacentTreePos(TPG_RIGHT);

	// Start from ptpgStart and search both directions at the same time.
	while(ptpLeft || ptpRight)
	{
		if(ptpLeft)
		{
			if(ptpLeft->IsText() && ptpLeft->TextID()==lTextID)
			{
				ptpRight = ptpLeft;

				// Starting at ptpLeft, loop to the left
				// looking for all of consecutive text poses 
				// with TextID of lTextID
				do
				{
					if(ptpLeft->IsText())
					{
						if(ptpLeft->TextID() == lTextID)
						{
							ptpFound = ptpLeft;
						}
						else
						{
							break;
						}
					}
					ptpLeft = ptpLeft->PreviousTreePos();
				} while(!ptpLeft->IsNode());

				ptpLeft = ptpFound;

				break;
			}
			ptpLeft = ptpLeft->PreviousTreePos();
		}

		if(ptpRight)
		{
			if(ptpRight->IsText() && ptpRight->TextID()==lTextID)
			{
				ptpLeft = ptpRight;

				// Starting at ptpRight, loop to the right
				// looking for all of consecutive text poses 
				// with TextID of lTextID
				do
				{
					if(ptpRight->IsText())
					{
						if(ptpRight->TextID() == lTextID)
						{
							ptpFound = ptpRight;
						}
						else
						{
							break;
						}
					}
					ptpRight = ptpRight->NextTreePos();
				} while(!ptpRight->IsNode());

				ptpRight = ptpFound;

				break;
			}
			ptpRight = ptpRight->NextTreePos();
		}
	}

	if(ptpFound)
	{
		Verify(!ptpgStart->MoveTo(ptpLeft, TPG_LEFT));
		Verify(!ptpgEnd->MoveTo(ptpRight, TPG_RIGHT));

		return S_OK;
	}

	return S_FALSE;
}

void CMarkup::SplitTextID(CTreePos* ptpLeft, CTreePos* ptpRight)
{
	Assert(ptpLeft && ptpRight);

	Assert(!HasUnembeddedPointers());

	// Find a text pos to the left if any
	while(ptpLeft)
	{
		if(ptpLeft->IsNode())
		{
			ptpLeft = NULL;
			break;
		}

		if(ptpLeft->IsText())
		{
			break;
		}

		ptpLeft = ptpLeft->PreviousTreePos();
	}

	// Find one to the right
	while(ptpRight)
	{
		if(ptpRight->IsNode())
		{
			ptpRight = NULL;
			break;
		}

		if(ptpRight->IsText())
		{
			break;
		}

		ptpRight = ptpRight->NextTreePos();
	}

	// If we have one to the left and right and they
	// both have the same ID (that isn't 0) we want
	// to give the fragment to the right a new ID.
	if(ptpLeft && ptpRight && ptpRight->TextID()
		&& ptpLeft->TextID()==ptpRight->TextID())
	{
		long lCurrTextID = ptpRight->TextID();
		long lNewTextID = ++(Doc()->_lLastTextID);

		while(ptpRight && !ptpRight->IsNode())
		{
			if(ptpRight->IsText())
			{
				if(ptpRight->TextID() == lCurrTextID)
				{
					DEBUG_ONLY(CTreePos* ptpOld = ptpRight);
					Verify(!SetTextPosID(&ptpRight, lNewTextID));
					Assert(ptpOld == ptpRight);
				}
				else
				{
					break;
				}
			}

			ptpRight = ptpRight->NextTreePos();
		}
	}
}

CTreePos* CMarkup::NewTextPos(long cch, SCRIPT_ID sid/*=sidAsciiLatin */, long lTextID/*=0*/)
{
	CTreePos* ptp;
	if(!lTextID)
	{
		ptp = AllocData1Pos();
	}
	else
	{
		ptp = AllocData2Pos();
	}

	if(ptp)
	{
		Assert(ptp->IsDataPos());
		ptp->SetType(CTreePos::Text);
		ptp->DataThis()->t._cch = cch;
		ptp->DataThis()->t._sid = sid;
		if(lTextID)
		{
			ptp->DataThis()->t._lTextID = lTextID;
		}
	}

	return ptp;
}

CTreePos* CMarkup::NewPointerPos(CMarkupPointer* pPointer, BOOL fRight, BOOL fStick)
{
#ifdef _WIN64
	CTreePos* ptp = AllocData2Pos();
#else
	CTreePos* ptp = AllocData1Pos();
#endif

	if(ptp)
	{
		Assert(ptp->IsDataPos());
		ptp->SetType(CTreePos::Pointer);
		Assert((DWORD_PTR(pPointer)&0x3) == 0);
		ptp->DataThis()->p._dwPointerAndGravityAndCling = (DWORD_PTR)(pPointer)|!!fRight|(fStick?0x2:0);
	}

	return ptp;
}

HRESULT CMarkup::Append(CTreePos* ptpNew)
{
	HRESULT hr = (ptpNew) ? S_OK : E_OUTOFMEMORY;
	CTreePos* ptpOldRoot = _tpRoot.FirstChild();

	if(hr)
	{
		goto Cleanup;
	}

	ptpNew->SetFirstChild(ptpOldRoot);
	ptpNew->MarkLeft();
	ptpNew->MarkLast();
	ptpNew->SetNext(&_tpRoot);

	ptpNew->ClearCounts();
	ptpNew->IncreaseCounts(&_tpRoot, CTreePos::TP_LEFT);

	_tpRoot.SetFirstChild(ptpNew);
	_tpRoot.IncreaseCounts(ptpNew, CTreePos::TP_DIRECT);

	if(ptpOldRoot)
	{
		Assert(ptpOldRoot->IsLastChild());
#if defined(MAINTAIN_SPLAYTREE_THREADS)
		// adjust threads
		CTreePos* ptpRightmost = ptpOldRoot->RightmostDescendant();
		Assert(ptpRightmost->RightThread() == NULL);
		ptpRightmost->SetRightThread(ptpNew);
		ptpNew->SetLeftThread(ptpRightmost);
		ptpNew->SetRightThread(NULL);
#endif
		ptpOldRoot->SetNext(ptpNew);
	}
	else
	{
		_ptpFirst = ptpNew;
	}

	Trace2("%p: Append TreePos cch=%ld\n", this, ptpNew->GetCch());

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::Insert(CTreePos* ptpNew, CTreePosGap* ptpgInsert)
{
	Assert(ptpgInsert->IsPositioned());

	return Insert(ptpNew, ptpgInsert->AttachedTreePos(),
		ptpgInsert->AttachDirection()==TPG_RIGHT);
}

HRESULT CMarkup::Insert(CTreePos* ptpNew, CTreePos* ptpInsert, BOOL fBefore)
{
	CTreePos *pLeftChild, *pRightChild;
	CTreePos *ptp;
	BOOL fLeftChild;
	HRESULT hr = (ptpNew) ? S_OK : E_OUTOFMEMORY;

	if(hr)
	{
		goto Cleanup;
	}

	Assert(ptpInsert);

	if(!fBefore)
	{
		CTreePos* ptpInsertBefore = ptpInsert->NextTreePos();

		if(ptpInsertBefore)
		{
			ptpInsert = ptpInsertBefore;
		}
		else
		{
			hr = Append(ptpNew);
			goto Cleanup;
		}
	}

	ptpNew->ClearCounts();
	ptpNew->IncreaseCounts(ptpInsert, CTreePos::TP_LEFT);

	ptpInsert->GetChildren(&pLeftChild, &pRightChild);

	ptpInsert->SetFirstChild(ptpNew);

	ptpNew->SetFirstChild(pLeftChild);
	ptpNew->MarkLeft();

	if(pLeftChild)
	{
		ptpNew->MarkLast(pLeftChild->IsLastChild());
		ptpNew->SetNext(pLeftChild->Next());
		pLeftChild->MarkLast();
		pLeftChild->SetNext(ptpNew);
	}
	else if(pRightChild)
	{
		ptpNew->MarkFirst();
		ptpNew->SetNext(pRightChild);
	}
	else
	{
		ptpNew->MarkLast();
		ptpNew->SetNext(ptpInsert);
	}

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	// adjust threads
	ptpNew->SetLeftThread(ptpInsert->LeftThread());
	ptpNew->SetRightThread(ptpInsert);
	if(ptpNew->LeftThread())
	{
		ptpNew->LeftThread()->SetRightThread(ptpNew);
	}
	ptpInsert->SetLeftThread(ptpNew);
#endif

	if(ptpNew->HasNonzeroCounts(CTreePos::TP_DIRECT))
	{
		fLeftChild = TRUE;
		for(ptp=ptpInsert; ptp; ptp=ptp->Parent())
		{
			if(fLeftChild)
			{
				ptp->IncreaseCounts(ptpNew, CTreePos::TP_DIRECT);
			}
			fLeftChild = ptp->IsLeftChild();
		}
	}

	if(ptpInsert == _ptpFirst)
	{
		_ptpFirst = ptpNew;
	}

	Trace2("%p: Insert TreePos cch=%ld\n", this, ptpNew->GetCch());

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::Move(CTreePos* ptpNew, CTreePosGap* ptpgDest)
{
	Assert(ptpgDest->IsPositioned());

	return Move(ptpNew, ptpgDest->AttachedTreePos(),
		ptpgDest->AttachDirection()==TPG_RIGHT);
}

HRESULT CMarkup::Move(CTreePos* ptpMove, CTreePos* ptpDest, BOOL fBefore)
{
	Assert(ptpMove != ptpDest);
	SUBLIST sublist;
	HRESULT hr;

	hr = SpliceOut(ptpMove, ptpMove, &sublist);
	if(!hr)
	{
		Assert(sublist.FirstChild() == ptpMove);
		hr = Insert(ptpMove, ptpDest, fBefore);
	}

	return hr;
}

HRESULT CMarkup::Remove(CTreePos* ptpStart, CTreePos* ptpFinish)
{
	HRESULT hr;
	SUBLIST sublist;

	hr = SpliceOut(ptpStart, ptpFinish, &sublist);
	if(!hr && sublist.FirstChild())
	{
		FreeTreePos(sublist.FirstChild());

		Trace2("%p: Remove cch=%ld\n", this, sublist._cchLeft);
	}

	RRETURN(hr);
}

HRESULT CMarkup::Split(CTreePos* ptpSplit, long cchLeft, SCRIPT_ID sidNew/*=sidNil*/)
{
	CTreePos*	ptpNew;
	CTreePos*	pLeftChild;
	CTreePos*	pRightChild;
	HRESULT		hr = S_OK;

	Assert(ptpSplit->IsText() && ptpSplit->Cch()>=cchLeft);

	if(sidNew == sidNil)
	{
		sidNew = ptpSplit->Sid();
	}

	ptpNew = NewTextPos(ptpSplit->Cch()-cchLeft, sidNew, ptpSplit->TextID());

	if(!ptpNew)
	{
		hr = E_OUTOFMEMORY;
		goto Cleanup;
	}

	ptpSplit->GetChildren(&pLeftChild, &pRightChild);

	ptpNew->MarkLast();
	ptpNew->SetNext(ptpSplit);
	ptpNew->SetFirstChild(pRightChild);
	if(pRightChild)
	{
		pRightChild->SetNext(ptpNew);
	}

	ptpSplit->DataThis()->t._cch = cchLeft;
	if(pLeftChild)
	{
		pLeftChild->MarkFirst();
		pLeftChild->SetNext(ptpNew);
	}
	else
	{
		ptpSplit->SetFirstChild(ptpNew);
	}

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	// adjust threads
	ptpNew->SetLeftThread(ptpSplit);
	ptpNew->SetRightThread(ptpSplit->RightThread());
	ptpSplit->SetRightThread(ptpNew);
	if(ptpNew->RightThread())
	{
		ptpNew->RightThread()->SetLeftThread(ptpNew);
	}
#endif

	Trace3("%p: Split TreePos into cch %ld/%ld\n", this, ptpSplit->Cch(), ptpNew->Cch());

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::Join(CTreePos* ptpJoin)
{
	CTreePos*	ptpNext;
	CTreePos*	ptpParent;
	CTreePos*	pLeftChild;

	// put the two TreePos together at the top of the tree, next TreePos on top
	ptpJoin->Splay();
	ptpNext = ptpJoin->NextTreePos();
	ptpNext->Splay();

	Assert(ptpJoin->IsText() && ptpNext->IsText());

	// take over the next TreePos (my right subtree is empty)
	ptpJoin->DataThis()->t._cch += ptpNext->Cch();
	if(!ptpJoin->IsLastChild())
	{
		ptpJoin->Next()->SetNext(ptpJoin);
		pLeftChild = ptpJoin->LeftChild();
		if(pLeftChild)
		{
			pLeftChild->MarkFirst();
			pLeftChild->SetNext(ptpJoin->Next());
		}
		else
		{
			ptpJoin->SetFirstChild(ptpJoin->Next());
		}
	}

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	// adjust threads
	ptpJoin->SetRightThread(ptpNext->RightThread());
	if(ptpNext->RightThread())
	{
		ptpNext->RightThread()->SetLeftThread(ptpJoin);
	}
#endif

	// discard the next TreePos
	ptpParent = ptpNext->Parent();
	Assert(ptpParent->Next() == NULL);
	ptpParent->ReplaceChild(ptpNext, ptpJoin);
	ptpNext->SetFirstChild(NULL);
	FreeTreePos(ptpNext);

	Trace2("%p: Join cch=%ld\n", this, ptpJoin->Cch());

	return S_OK;
}

HRESULT CMarkup::ReplaceTreePos(CTreePos* ptpOld, CTreePos* ptpNew)
{
	HRESULT hr = S_OK;

	Assert(!HasUnembeddedPointers());

	Assert(ptpOld && ptpNew);

	hr = Insert(ptpNew, ptpOld, TRUE);
	if(hr)
	{
		goto Cleanup;
	}

	hr = Remove(ptpOld);
	if(hr)
	{
		goto Cleanup;
	}

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::MergeText(CTreePos* ptpMerge)
{
	HRESULT hr = S_OK;
	CTreePos* ptp;

	Assert(!HasUnembeddedPointers());

	if(!ptpMerge->IsText())
	{
		goto Cleanup;
	}

	ptp = ptpMerge->PreviousTreePos();

	if(ptp->IsText() && ptp->Sid()==ptpMerge->Sid()
		&& ptp->TextID()==ptpMerge->TextID())
	{
		hr = MergeTextHelper(ptp);

		if(hr)
		{
			goto Cleanup;
		}

		ptpMerge = ptp;
	}

	ptp = ptpMerge->NextTreePos();

	if(ptp->IsText() && ptp->Sid()==ptpMerge->Sid()
		&& ptp->TextID()==ptpMerge->TextID())
	{
		hr = MergeTextHelper(ptpMerge);

		if(hr)
		{
			goto Cleanup;
		}
	}

Cleanup:

	RRETURN(hr);
}

HRESULT CMarkup::SetTextPosID(CTreePos** pptpText, long lTextID)
{
	HRESULT hr = S_OK;

	Assert(pptpText && *pptpText);

	if(!(*pptpText)->IsData2Pos())
	{
		CTreePos* ptpRet;

		ptpRet = AllocData2Pos();
		if(!ptpRet)
		{
			hr = E_OUTOFMEMORY;
			goto Cleanup;
		}

		ptpRet->SetType(CTreePos::Text);
		*(LONG*)&ptpRet->DataThis()->t = *(LONG*)&(*pptpText)->DataThis()->t;

		hr = ReplaceTreePos(*pptpText, ptpRet);
		if(hr)
		{
			goto Cleanup;
		}

		*pptpText = ptpRet;
	}

	(*pptpText)->DataThis()->t._lTextID = lTextID;

Cleanup:
	RRETURN(hr);
}

HRESULT CMarkup::RemovePointerPos(CTreePos* ptp, CTreePos** pptpUpdate, long* pichUpdate)
{
	HRESULT		hr = S_OK;
	CTreePos*	ptpBefore;
	CTreePos*	ptpAfter;

	Assert(ptp->IsPointer());
	Assert(!pptpUpdate || *pptpUpdate!=ptp);

	// Remove the pos from the list, making sure to record the pos before the one
	// being removed
	ptpBefore = ptp->PreviousTreePos();

	hr = Remove(ptp);

	if(hr)
	{
		goto Cleanup;
	}

	// Now, see if there are two adjacent text pos's we can merge
	if(ptpBefore->IsText() && (ptpAfter=ptpBefore->NextTreePos())->IsText() &&
		ptpBefore->Sid()==ptpAfter->Sid() && ptpBefore->TextID()==ptpAfter->TextID())
	{
		CMarkupPointer* pmp;

		// Update the incomming ref and unembedded markup pointers
		if(pptpUpdate && *pptpUpdate==ptpAfter)
		{
			*pptpUpdate = ptpBefore;
			*pichUpdate += ptpBefore->Cch();
		}

		for(pmp=_pmpFirst; pmp; pmp=pmp->_pmpNext)
		{
			Assert(!pmp->_fEmbedded);

			if(pmp->_ptpRef == ptpAfter)
			{
				pmp->_ptpRef = ptpBefore;

				Assert(!pmp->_ptpRef->IsPointer());

				Assert(!pmp->_ptpRef->IsBeginElementScope() ||
					!pmp->_ptpRef->Branch()->Element()->IsNoScope());

				pmp->_ichRef += ptpBefore->Cch();
			}
		}

		// Then joing the adjacent text pos's
		hr = Join(ptpBefore);

		if(hr)
		{
			goto Cleanup;
		}
	}

Cleanup:

	RRETURN(hr);
}

HRESULT CMarkup::SpliceOut(CTreePos* ptpStart, CTreePos* ptpFinish, SUBLIST* pSublistSplice)
{
	CTreePos*	ptpSplice;
	CTreePos*	ptp;
	BOOL		fLeftChild;
	CTreePos*	ptpLeftEdge = ptpStart->PreviousTreePos();
	CTreePos*	ptpRightEdge = ptpFinish->NextTreePos();
#if defined(MAINTAIN_SPLAYTREE_THREADS)
	CTreePos *ptpLeftmost, *ptpRightmost;
#endif

	pSublistSplice->InitSublist();

	// isolate the splice region near the top of the tree
	if(ptpRightEdge)
	{
		ptpRightEdge->Splay();
	}
	if(ptpLeftEdge)
	{
		ptpLeftEdge->Splay();
	}

	// locate the splice region
	ptpSplice = _tpRoot.FirstChild();
	if(ptpLeftEdge)
	{
		ptpSplice = ptpLeftEdge->RightChild();
		// call RotateUp to isolate if last step was zig-zig
		if(ptpRightEdge && ptpRightEdge->Parent()==ptpSplice)
		{
			ptpRightEdge->RotateUp(ptpSplice, ptpSplice->Parent());
		}
	}
	if(ptpRightEdge)
	{
		ptpSplice = ptpRightEdge->LeftChild();
	}

	ptp = ptpSplice->Parent();
	fLeftChild = ptpSplice->IsLeftChild();

	// splice it out
	if(fLeftChild)
	{
		pSublistSplice->IncreaseCounts(ptp, CTreePos::TP_LEFT);
	}
	else
	{
		Assert(ptp->Parent() == &_tpRoot);
		pSublistSplice->IncreaseCounts(&_tpRoot, CTreePos::TP_LEFT);
		pSublistSplice->DecreaseCounts(ptp, CTreePos::TP_BOTH);
	}
	ptp->RemoveChild(ptpSplice);
	pSublistSplice->SetFirstChild(ptpSplice);
	ptpSplice->SetNext(pSublistSplice);
	ptpSplice->MarkLast();
	ptpSplice->MarkLeft();

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	// adjust threads
	ptpLeftmost = ptpSplice->LeftmostDescendant();
	ptpRightmost = ptpSplice->RightmostDescendant();
	if(ptpLeftmost->LeftThread())
	{
		ptpLeftmost->LeftThread()->SetRightThread(ptpRightmost->RightThread());
	}
	if(ptpRightmost->RightThread())
	{
		ptpRightmost->RightThread()->SetLeftThread(ptpLeftmost->LeftThread());
	}
	ptpLeftmost->SetLeftThread(NULL);
	ptpRightmost->SetRightThread(NULL);
#endif

	// adjust cumulative counts
	if(pSublistSplice->HasNonzeroCounts(CTreePos::TP_LEFT))
	{
		for(; ptp; ptp=ptp->Parent())
		{
			if(fLeftChild)
			{
				ptp->DecreaseCounts(pSublistSplice, CTreePos::TP_LEFT);
			}
			fLeftChild = ptp->IsLeftChild();
		}
	}

	// return answers
	Trace2("%p: SpliceOut cch=%ld\n", this, pSublistSplice->_cchLeft);

	if(ptpLeftEdge == NULL)
	{
		_ptpFirst = _tpRoot.LeftmostDescendant();
		if(_ptpFirst == &_tpRoot)
		{
			_ptpFirst = NULL;
		}
	}

	return S_OK;
}

HRESULT CMarkup::SpliceIn(SUBLIST* pSublistSplice, CTreePos* ptp)
{
	CTreePos*	ptpSplice = pSublistSplice->FirstChild();
	CTreePos*	pRightChild;
	CTreePos*	ptpPrev;
	BOOL		fLeftChild;
	BOOL		fNewFirst = (ptp == _ptpFirst);

	Assert(ptpSplice);

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	CTreePos* ptpLeftmost = ptpSplice->LeftmostDescendant();
	CTreePos* ptpRightmost = ptpSplice->RightmostDescendant();
#endif

	if(ptp)
	{
		ptp->Splay();

		ptpPrev = ptp->PreviousTreePos();
		if(ptpPrev)
		{
			ptpPrev->Splay();
		}

		pRightChild = ptp->RightChild();
		if(pRightChild)
		{
			ptpSplice->MarkFirst();
			ptpSplice->SetNext(pRightChild);
		}
		else
		{
			ptpSplice->MarkLast();
			ptpSplice->SetNext(ptp);
		}

		ptp->SetFirstChild(ptpSplice);
		ptpSplice->MarkLeft();

		fLeftChild = TRUE;

#if defined(MAINTAIN_SPLAYTREE_THREADS)
		// adjust threads
		if(ptp->LeftThread())
		{
			ptp->LeftThread()->SetRightThread(ptpLeftmost);
			ptpLeftmost->SetLeftThread(ptp->LeftThread());
		}
		ptpRightmost->SetRightThread(ptp);
		ptp->SetLeftThread(ptpRightmost);
#endif
	}
	else
	{
		ptp = LastTreePos();
		Assert(ptp);
		CTreePos* ptpLeft = ptp->LeftChild();

		if(ptpLeft)
		{
			ptpLeft->MarkFirst();
			ptpLeft->SetNext(ptpSplice);
		}
		else
		{
			ptp->SetFirstChild(ptpSplice);
		}
		ptpSplice->MarkLast();
		ptpSplice->MarkRight();
		ptpSplice->SetNext(ptp);

		fLeftChild = FALSE;

#if defined(MAINTAIN_SPLAYTREE_THREADS)
		// adjust threads
		ptp->SetRightThread(ptpLeftmost);
		ptpLeftmost->SetLeftThread(ptp);
#endif
	}

	for(; ptp; ptp=ptp->Parent())
	{
		if(fLeftChild)
		{
			ptp->IncreaseCounts(pSublistSplice, CTreePos::TP_LEFT);
		}
		fLeftChild = ptp->IsLeftChild();
	}

	if(fNewFirst)
	{
		_ptpFirst = _ptpFirst->LeftmostDescendant();
	}

	Trace2("%p: SpliceIn cch=%ld\n", this, pSublistSplice->_cchLeft);
	return S_OK;
}

HRESULT CMarkup::InsertPosChain(CTreePos* ptpChainHead, CTreePos* ptpChainTail, CTreePos* ptpInsertBefore)
{
	CTreePos *pLeftChild, *pRightChild;

	// We can't insert at the beginning or end of the splay tree
	Assert(ptpInsertBefore != _ptpFirst);
	Assert(ptpInsertBefore);

	ptpInsertBefore->GetChildren(&pLeftChild, &pRightChild);

	ptpInsertBefore->SetFirstChild(ptpChainTail);
	ptpChainTail->MarkLeft();

	ptpChainHead->SetFirstChild(pLeftChild);

	if(pLeftChild)
	{
		ptpChainTail->MarkLast(pLeftChild->IsLastChild());
		ptpChainTail->SetNext(pLeftChild->Next());

		pLeftChild->MarkLast();
		pLeftChild->SetNext(ptpChainHead);
	}
	else if(pRightChild)
	{
		ptpChainTail->MarkFirst();
		ptpChainTail->SetNext(pRightChild);
	}
	else
	{
		ptpChainTail->MarkLast();
		ptpChainTail->SetNext(ptpInsertBefore);
	}

#if defined(MAINTAIN_SPLAYTREE_THREADS)
	// adjust threads
	ptpChainHead->SetLeftThread(ptpInsertBefore->LeftThread());
	ptpChainTail->SetRightThread(ptpInsertBefore);
	if(ptpChainHead->LeftThread())
	{
		ptpChainHead->LeftThread()->SetRightThread(ptpChainHead);
	}
	ptpInsertBefore->SetLeftThread(ptpChainTail);
#endif

	// update the order statistics
	{
		CTreePos *ptp, *ptpPrev=NULL;
		BOOL fLeftChild=TRUE;
		CTreePos::SCounts counts;

		// accumulate/set the chain just inserted
		ptpChainHead->ClearCounts();
		ptpChainHead->IncreaseCounts(ptpInsertBefore, CTreePos::TP_LEFT);

		counts.Clear();
		counts.Increase(ptpChainHead);

		for(ptpPrev=ptpChainHead,ptp=ptpChainHead->Parent(); 
			ptpPrev!=ptpChainTail; 
			ptpPrev=ptp,ptp=ptp->Parent())
		{
			ptp->ClearCounts();
			ptp->IncreaseCounts(ptpPrev, CTreePos::TP_BOTH);

			counts.Increase(ptp);
		}

		// This method isn't used to insert pointers so
		// we must have some sort of count
		Assert(counts.IsNonzero());

		fLeftChild = TRUE;
		for(ptp=ptpInsertBefore; ptp; ptp=ptp->Parent())
		{
			if(fLeftChild)
			{
				ptp->IncreaseCounts(&counts);
			}
			fLeftChild = ptp->IsLeftChild();
		}
	}

	RRETURN(S_OK);
}

CTreePos* CMarkup::FirstTreePos() const
{
	return _ptpFirst;
}

CTreePos* CMarkup::LastTreePos() const
{
	CTreePos* ptpLeft = _tpRoot.FirstChild();

	return (ptpLeft)?ptpLeft->RightmostDescendant():NULL;
}

CTreePos* CMarkup::TreePosAtCp(long cp, long* pcchOffset) const
{
	long cDepth = 0;
	DEBUG_ONLY(long cpOrig = cp);

	// Make sure we got a valid cp.
	AssertSz(cp>=1&&cp<Cch(), "Invalid cp - out of document");

	CTreePos* ptp = _tpRoot.FirstChild();

	for(;; ++cDepth)
	{
		if(cp < long(ptp->_cchLeft))
		{
			ptp = ptp->FirstChild();

			if(!ptp || !ptp->IsLeftChild())
			{
				ptp = FirstTreePos();
				break; // we fell off the left end
			}
		}
		else
		{
			cp -= ptp->_cchLeft;

			if(ptp->IsPointer() || cp && (!ptp->IsText()||cp>=ptp->Cch()))
			{
				cp -= ptp->GetCch();
				ptp = ptp->RightChild();
			}
			else
			{
				break;
			}
		}
	}

	if(pcchOffset)
	{
		*pcchOffset = cp;
	}

	Trace((_T("%p: TreePos at cp %ld offset %ld depth %ld\n"), this, cpOrig, cp, cDepth));

	if(ShouldSplay(cDepth))
	{
		ptp->Splay();
	}

	return ptp;
}

CTreePos* CMarkup::TreePosAtSourceIndex(long iSourceIndex)
{
	Assert(0 <= iSourceIndex);
	DEBUG_ONLY(long iSIOrig = iSourceIndex);
	long cElemLeft;
	long cDepth = 0;

	CTreePos* ptp = _tpRoot.FirstChild();

	for(;; ++cDepth)
	{
		cElemLeft = ptp->GetElemLeft();

		if(iSourceIndex < cElemLeft)
		{
			ptp = ptp->FirstChild();
			Assert(ptp->IsLeftChild());
		}
		else if(iSourceIndex>cElemLeft || !ptp->IsBeginElementScope())
		{
			iSourceIndex -= cElemLeft + (ptp->IsBeginElementScope()?1:0);
			ptp = ptp->RightChild();
		}
		else
		{
			break;
		}
	}

	Trace3("%p: TreePos at sourceindex %ld  depth %ld\n", this, iSIOrig, cDepth);

	if(ShouldSplay(cDepth))
	{
		ptp->Splay();
	}

	return ptp;
}

long CMarkup::CchInRange(CTreePos* ptpFirst, CTreePos* ptpLast)
{
	Assert(ptpFirst && ptpLast);
	return  ptpLast->GetCp()+ptpLast->GetCch()-ptpFirst->GetCp();
}

CTreePos* CMarkup::AllocData1Pos()
{
	CTreeDataPos* ptdpNew;

	if(!_ptdpFree)
	{
		void* pvPoolNew;
		size_t cPoolSize;

		if(!_pvPool)
		{
			pvPoolNew = (void*)(_abPoolInitial);
			cPoolSize = INITIAL_TREEPOS_POOL_SIZE;
		}
		else
		{
			cPoolSize = s_cTreePosPoolSize;
			pvPoolNew = _MemAllocClear(sizeof(void*)+TREEDATA1SIZE*s_cTreePosPoolSize);
		}

		if(pvPoolNew)
		{
			CTreeDataPos* ptdp;
			int i;

			for(ptdp=(CTreeDataPos*)((void**)pvPoolNew+1),i=cPoolSize-1; 
				i>0; ptdp=(CTreeDataPos*)((BYTE*)ptdp+TREEDATA1SIZE),--i)
			{
				ptdp->SetFlag(CTreePos::TPF_DATA_POS);
				ptdp->SetNext((CTreeDataPos*)((BYTE*)ptdp+TREEDATA1SIZE));
			}

			ptdp->SetFlag(CTreePos::TPF_DATA_POS);

			_ptdpFree = (CTreeDataPos*)((void**)pvPoolNew+1);
			*((void**)pvPoolNew) = _pvPool;
			_pvPool = pvPoolNew;
		}
	}

	ptdpNew = _ptdpFree;
	if(ptdpNew)
	{
		_ptdpFree = (CTreeDataPos*)(ptdpNew->Next());
		ptdpNew->SetNext(NULL);
	}

	return ptdpNew;
}

CTreePos* CMarkup::AllocData2Pos()
{
	CTreeDataPos* ptdpNew;
	ptdpNew = (CTreeDataPos*)_MemAllocClear(TREEDATA2SIZE);

	if(ptdpNew)
	{
		ptdpNew->SetFlag(CTreePos::TPF_DATA_POS|CTreePos::TPF_DATA2_POS);
	}

	return ptdpNew;
}

void CMarkup::FreeTreePos(CTreePos* ptp)
{
	Assert(ptp);

	// set a sentinel to make the traversal terminate
	ptp->MarkFirst();
	ptp->SetNext(NULL);

	// release the subtree, adding its nodes to the free list
	while(ptp)
	{
		if (ptp->FirstChild())
		{
			ptp = ptp->FirstChild();
		}
		else
		{
			CTreePos* ptpNext;
			BOOL fRelease = TRUE;
			while(fRelease)
			{
				fRelease = ptp->IsLastChild();
				ptpNext = ptp->Next();
				ReleaseTreePos(ptp);
				ptp = ptpNext;
			}
		}
	}
}

void CMarkup::ReleaseTreePos(CTreePos* ptp, BOOL fLastRelease/*= FALSE*/)
{
	switch(ptp->Type())
	{
	case CTreePos::Pointer:
		if(ptp->MarkupPointer())
		{
			ptp->MarkupPointer()->OnPositionReleased();
			ptp->DataThis()->p._dwPointerAndGravityAndCling = 0;
		}

		// fall through
	case CTreePos::Text:
		Assert(ptp->IsDataPos());
		if(ptp->IsData2Pos())
		{
			MemFree(ptp);
		}
		else if(!fLastRelease)
		{
			memset(ptp, 0, TREEDATA1SIZE);

			ptp->SetFlag(CTreePos::TPF_DATA_POS);
			ptp->SetNext(_ptdpFree);
			_ptdpFree = ptp->DataThis();
		}
		break;

	case CTreePos::NodeBeg:
	case CTreePos::NodeEnd:
		// Make sure that we crash if someone
		// tries to use a dead node pos...
		ptp->_pFirstChild = NULL;
		ptp->_pNext = NULL;
		break;
	}
}

BOOL CMarkup::ShouldSplay(long cDepth) const
{
	return cDepth>4&&(cDepth>30||(0x1<<cDepth)>NumElems());
}

HRESULT CMarkup::MergeTextHelper(CTreePos* ptpMerge)
{
	HRESULT hr = S_OK;

	Assert(ptpMerge->IsText());
	Assert(ptpMerge->NextTreePos()->IsText());

	hr = Join(ptpMerge);

	if(hr)
	{
		goto Cleanup;
	}
Cleanup:
	RRETURN(hr);
}

void CMarkup::Notify(CNotification* pnf)
{
	Assert(pnf);

	// If the branch was not supplied, infer it from the affected element
	if(!pnf->_pNode && pnf->_pElement)
	{
		pnf->_pNode = pnf->_pElement->GetFirstBranch();
	}

	Assert(pnf->_pNode || pnf->_pElement || pnf->IsTreeChange()
		|| (pnf->SendTo(NFLAGS_DESCENDENTS) && pnf->IsRangeValid()));
	
	Assert(!pnf->SendTo(NFLAGS_ANCESTORS) || !pnf->SendTo(NFLAGS_DESCENDENTS));

	// If a notification is blocking the ancestor direction, queue the incoming notification
	if(_fInSendAncestor && pnf->SendTo(NFLAGS_ANCESTORS))
	{
		if(!pnf->IsFlagSet(NFLAGS_SYNCHRONOUSONLY))
		{
			if(pnf->Element())
			{
				pnf->Element()->AddRef();
				pnf->SetFlag(NFLAGS_NEEDS_RELEASE);
			}

			_aryANotification.AppendIndirect(pnf);
		}
	}
	else
	{
		CDataAry<CNotification>* paryNotification = pnf->SendTo(NFLAGS_ANCESTORS)
			? (CDataAry<CNotification>*)&_aryANotification : NULL;

		if(!pnf->IsFlagSet(NFLAGS_DONOTBLOCK))
		{
			if(pnf->SendTo(NFLAGS_ANCESTORS))
			{
				_fInSendAncestor = TRUE;
			}
		}

		SendNotification(pnf, paryNotification);

		Assert(!_fInSendAncestor || (paryNotification!=&_aryANotification));
	}
	return;
}

void CMarkup::SendNotification(CNotification* pnf, CDataAry<CNotification>* paryNotification)
{
	CNotification nf;
	int iRequest = 0;

	for(;;)
	{
		Assert(pnf);

		BOOL fNeedsRelease = pnf->IsFlagSet(NFLAGS_NEEDS_RELEASE);
		CElement* pElement = pnf->_pElement;

		// Send the notification to "self"
		if(pnf->SendTo(NFLAGS_SELF))
		{
			CElement* pElementSelf = pElement ? pElement : pnf->_pNode->Element();

			Assert(pElementSelf);
			if(ElementWantsNotification(pElementSelf, pnf))
			{
				NotifyElement(pElementSelf, pnf);
			}
		}

		// If sending to ancestors or descendents and a starting node or range exists,
		// Send the notification to "ancestors" and "descendents"
		if(!pnf->IsFlagSet(NFLAGS_SENDENDED) && (pnf->_pNode || pnf->IsRangeValid()))
		{
			if(pnf->SendTo(NFLAGS_ANCESTORS))
			{
				Assert(pnf->_pNode);
				NotifyAncestors(pnf);
			}
			else if(pnf->SendTo(NFLAGS_DESCENDENTS))
			{
				NotifyDescendents(pnf);
			}
		}

		// Send the notification to listeners at the "tree" level
		if(!pnf->IsFlagSet(NFLAGS_SENDENDED) && pnf->SendTo(NFLAGS_TREELEVEL))
		{
			NotifyTreeLevel(pnf);
		}

		// Release the element (if necessary)
		// (Elements are AddRef'd only when the associated notification is delayed)
		if(fNeedsRelease)
		{
			Assert(pElement);
			Assert(paryNotification);
			Assert(iRequest <= (*paryNotification).Size());
			DEBUG_ONLY((*paryNotification)[iRequest-1].ClearFlag(NFLAGS_NEEDS_RELEASE));
			pElement->Release();
		}

		// Leave or fetch the next notification to send
		// (Copy the notification into a local in case others should arrive
		//  and cause re-allocation of the array)
		if(!paryNotification || iRequest>=(*paryNotification).Size())
		{
			goto Cleanup;
		}

		nf = (*paryNotification)[iRequest++];
		pnf = &nf;

		// Ensure notifications no longer part of this markup are sent only to self
		// (Previously sent notifications can initiate changes such that the elements
		//  of pending notifications are no longer part of this markup)
		if(pnf->Element())
		{
			CMarkup* pMarkup = pnf->Element()->GetMarkupPtr();

			//  BUGBUG: Unfortunately, notifications forwarded from nested markups contain the element
			//          from the nested markup, hence the more complex check. What should be is, when
			//          a notification from a nested markup is forwarded (by CMarkup::NotifyTreeLevel)
			//          the contained element et. al. should be changed to the "master" element.
			//          This change cannot be made right now because it is potentially de-stabilizing. (brendand)
			if(!pMarkup || (pMarkup!=this && !pMarkup->Master())
				|| (pMarkup!=this && pMarkup->Master() && pMarkup->Master()->GetMarkupPtr()!=this))
			{
				pnf->_grfFlags &= ~NFLAGS_TARGETMASK | NFLAGS_SELF;
			}
		}
	}

Cleanup:
	if(paryNotification)
	{
#ifdef _DEBUG
		for(int i=0; i<(*paryNotification).Size(); i++)
		{
			Assert(!(*paryNotification)[i].IsFlagSet(NFLAGS_NEEDS_RELEASE));
		}
#endif

		(*paryNotification).DeleteAll();

		if(paryNotification == &_aryANotification)
		{
			_fInSendAncestor = FALSE;
		}
	}
}

BOOL CMarkup::ElementWantsNotification(CElement* pElement, CNotification* pnf)
{
	return  !pElement->_fExittreePending
		&& (pnf->IsForAllElements()
		|| (pElement->HasLayout()
		&& (pnf->IsTextChange()
		|| pnf->IsTreeChange()
		|| pnf->IsLayoutChange()
		|| pnf->IsForLayouts()
		|| (pnf->IsForPositioned()
		&& (pElement->IsZParent()
		|| pElement->GetLayoutPtr()->_fContainsRelative))
		|| (pnf->IsForWindowed()
		&& pElement->GetHwnd())))
		|| ((IsPositionNotification(pnf)
		// BUGBUG: Rework the categories to include the following:
		//         1) text change
		//         2) tree change
		//         3) layout change
		//         4) display change
		//         Then, instead of testing for NTYPE_VISIBILITY_CHANGE etc., test for
		//         display changes
		|| pnf->IsType(NTYPE_VISIBILITY_CHANGE))
		&& !pElement->IsPositionStatic())
		|| (pnf->IsForActiveX()
		&& pElement->TestClassFlag(CElement::ELEMENTDESC_OLESITE)));
}

void CMarkup::NotifyElement(CElement* pElement, CNotification* pnf)
{
	Assert(pElement);
	Assert(ElementWantsNotification(pElement, pnf));
	Assert(!pnf->IsFlagSet(NFLAGS_SENDENDED));

	// If a layout exists, pass it all notifications except those meant for ActiveX elements
	// NOTE: Notifications passed to a layout may not also sent to the element
	if(pElement->HasLayout() && !pnf->IsForActiveX())
	{
		CLayout* pLayout = pElement->GetCurLayout();

		Assert(pLayout);

		pLayout->Notify(pnf);

		if(pnf->IsFlagSet(NFLAGS_SENDUNTILHANDLED) && pnf->IsHandled())
		{
			pnf->SetFlag(NFLAGS_SENDENDED);
		}
	}

	//  If not handled, hand appropriate notifications to the element
	//
	// (KTam) These look like the same conditions that ElementWantsNotification()
	// checks.  Why are we checking them again when we've already asserted that
	// ElementWantsNotification is true?
	if(!pnf->IsFlagSet(NFLAGS_SENDENDED)
		&& (pnf->IsForAllElements()
		|| (pnf->IsForActiveX()
		&& pElement->TestClassFlag(CElement::ELEMENTDESC_OLESITE))
		|| ((IsPositionNotification(pnf)
		// BUGBUG: Rework the categories to include the follownig:
		//         1) text change
		//         2) tree change
		//         3) layout change
		//         4) display change
		//         Then, instead of testing for NTYPE_VISIBILITY_CHANGE etc., test for
		//         display changes
		|| (pnf->IsType(NTYPE_VISIBILITY_CHANGE)
		&& !pElement->NeedsLayout()))
		&& !pElement->IsPositionStatic())))
	{
		pElement->Notify(pnf);

		if(pnf->IsFlagSet(NFLAGS_SENDUNTILHANDLED) &&pnf->IsHandled())
		{
			pnf->SetFlag(NFLAGS_SENDENDED);
		}
	}
}

void CMarkup::NotifyAncestors(CNotification* pnf)
{
	CTreeNode* pNodeBranch;
	CTreeNode* pNode;

	Assert(pnf);
	Assert(pnf->_pNode);

	for(pNodeBranch=pnf->_pNode; pNodeBranch;
		//  BUGBUG: We may eventually want this routine should distribute the notification to all parents of the element.
		//          Unfortunately, doing so now is too problematic (since layouts cannot handle receiving the same
		//          notification multiple times). (brendand)
		//        pNodeBranch = pNodeBranch->NextBranch())
		pNodeBranch=NULL)
	{
		pNode = pNodeBranch->Parent();

		while(pNode && !pnf->IsFlagSet(NFLAGS_SENDENDED))
		{
			CElement* pElement = pNode->Element();

			if(ElementWantsNotification(pElement, pnf))
			{
				pElement->AddRef();

				NotifyElement(pElement, pnf);

				if(!pElement->IsInMarkup())
				{
					pNode = NULL;
				}

				pElement->Release();
			}

			if(pNode)
			{
				pNode = pNode->Parent();
			}
		}
	}
}

void CMarkup::NotifyDescendents(CNotification* pnf)
{
	CStackPtrAry<CElement*, 32> aryElements;

	Assert(pnf->SendTo(NFLAGS_DESCENDENTS));

	CTreePos* ptp;
	CTreePos* ptpEnd;
	DEBUG_ONLY(CTreeNode* pNodeEnd);

	BOOL fZParentsOnly = pnf->IsFlagSet( NFLAGS_ZPARENTSONLY );
	BOOL fSCAvail = pnf->IsSecondChanceAvailable();

	Assert(pnf);
	Assert(pnf->Element() || pnf->IsRangeValid());
	Assert(!pnf->Element() || pnf->Element()->IsInMarkup());
	Assert(!pnf->Element() || pnf->Element()->GetMarkup()==this);

	// Determine the range
	if(pnf->Element())
	{
		pnf->Element()->GetTreeExtent(&ptp, &ptpEnd);

		Assert(ptp);

		ptp = ptp->NextTreePos();

		if(pnf->Element()==Root() && !fSCAvail)
		{
			aryElements.EnsureSize(NumElems());
		}
	}
	else
	{
		long cpStart, cpEnd, ich;

		Assert(pnf->IsRangeValid());

		cpStart = pnf->Cp(0);
		cpEnd = cpStart + pnf->Cch(LONG_MAX);

		ptp = TreePosAtCp(cpStart, &ich);
		ptpEnd = TreePosAtCp(cpEnd, &ich);
	}

	Assert(ptp);
	Assert(ptpEnd);
	Assert(ptp->InternalCompare(ptpEnd) <= 0);

	DEBUG_ONLY(pNodeEnd = CTreePosGap(ptpEnd, TPG_LEFT).Branch());

	// Build a list of target elements
	// (This allows the tree to change shape as the notification is delivered to each target)
	while(ptp && ptp!=ptpEnd)
	{
		if(ptp->IsBeginElementScope())
		{
			CTreeNode* pNode = ptp->Branch();
			CElement* pElement = pNode->Element();

			// Remember the element if it may want the notification
			if(ElementWantsNotification(pElement, pnf))
			{
				if(fSCAvail)
				{
					if(!pnf->IsFlagSet( NFLAGS_SENDENDED))
					{
						NotifyElement(pElement, pnf);
					}

					if(pnf->IsSecondChanceRequested())
					{
						HRESULT hr;

						hr = aryElements.Append(pElement);
						if(hr)
						{
							goto Cleanup;
						}

						pElement->AddRef();

						pnf->ClearSecondChanceRequested();
					}
				}
				else
				{
					HRESULT hr;

					hr = aryElements.Append(pElement);
					if(hr)
					{
						goto Cleanup;
					}

					pElement->AddRef();
				}
			}

			// Skip over z-parents (if requested)
			if(fZParentsOnly)
			{
				BOOL fSkipElement;

				fSkipElement = pNode->IsZParent()
					|| (pnf->IsFlagSet(NFLAGS_AUTOONLY)
					&& pElement->HasLayout()
					&& !pElement->GetCurLayout()->_fAutoBelow);

				// If the element should be skipped, then advance past it
				// (If the element has layout that overlaps with another layout that was not
				//  skipped, then this routine will only skip the portion contained in the
				//  non-skipped layout)
				if(fSkipElement)
				{
					Assert(!pNodeEnd->SearchBranchToRootForScope(pElement));

					for(;;)
					{
						CElement* pElementInner;
						// Get the ending treepos,
						// stop if it is for the end of the element
						ptp = pNode->GetEndPos();

						Assert(ptp->IsEndNode());

						if(ptp->IsEdgeScope())
						{
							break;
						}
						// There is overlap, locate the beginning of the next section
						// (Tunnel through the inclusions stopping at either the next begin
						//  node for the current element, the end of the range, or if an
						//  element with layout ends)
						do
						{
							ptp = ptp->NextTreePos();
							pElementInner = ptp->Branch()->Element();
							Assert(!ptp->IsEdgeScope() || ptp->IsEndNode());
						} while(pElementInner!=pElement && ptp!=ptpEnd
							&& !(ptp->IsEdgeScope() && pElementInner->HasLayout()));

						// If range ended, stop all searching
						if(ptp == ptpEnd)
						{
							goto Cleanup;
						}

						// If the end of an element with layout was encountered,
						// treat that as the end of the skipped over element
						Assert(!ptp->IsEdgeScope() || (ptp->IsEndNode() && pElementInner->HasLayout()));

						if(ptp->IsEdgeScope())
						{
							break;
						}

						// Otherwise, reset the treenode and skip over the section of the element
						Assert(ptp->IsBeginNode() && !ptp->IsEdgeScope());

						pNode = ptp->Branch();

						Assert(pNode->Element() == pElement);
					}
				}
			}
		}

		ptp = ptp->NextTreePos();
	}

	// Deliver the notification
Cleanup:
	if(aryElements.Size())
	{
		CElement**		ppElement;
		int				c;
		CNotification	nfSc;
		CNotification*	pnfSend;

		if(fSCAvail)
		{
			nfSc.InitializeSc(pnf);
			pnfSend = &nfSc;
		}
		else
		{
			pnfSend = pnf;
		}

		for(c=aryElements.Size(),ppElement=&(aryElements[0]);
			c>0;
			c--,ppElement++)
		{
			Assert(ppElement && *ppElement);

			if(!pnf->IsFlagSet( NFLAGS_SENDENDED))
			{
				NotifyElement(*ppElement, pnfSend);
			}

			(*ppElement)->Release();
		}
	}
}

void CMarkup::NotifyTreeLevel(CNotification* pnf)
{
	Assert(!pnf->IsFlagSet(NFLAGS_SENDENDED));

	if(pnf->Type() == NTYPE_DISPLAY_CHANGE)
	{
		CTreeNode* pNode = pnf->Node();
		while(pNode)
		{
			CElement* pElement = pNode->Element();
			if(pElement->HasFlag(TAGDESC_LIST))
			{
				break;
			}
			pNode = pNode->Parent();
		}
	}
	else
	{
		// Notify the view of all layout changes
		if(!pnf->IsTextChange() && (pnf->IsLayoutChange() || pnf->IsForLayouts()))
		{
			Doc()->GetView()->Notify(pnf);
		}

		//  Forward the notification to the master markup (if it exists)
		//

		//  BUGBUG: Two changes are needed here. First, the notification should be altered to contain
		//          the "master" element prior to forwarding. Second, category flags should be added
		//          to efficiently filter notifications that need forwarding. (It's also reasonable
		//          to imagine conversion flags that say how some forwarded notifications should change
		//          prior to forwarding - e.g., some changes become a resize.) Unfortunately, these
		//          changes are too de-stabilizing to do at this time. (brendand)
		if(Master() && Master()->IsInMarkup()
			&& !pnf->IsType(NTYPE_VIEW_ATTACHELEMENT)
			&& !pnf->IsType(NTYPE_VIEW_DETACHELEMENT))
		{
			Master()->GetUpdatedLayout()->Notify(pnf);
		}
	}
}

CTreeNode* CMarkup::FindMyListContainer(CTreeNode* pNodeStartHere)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode = pNodeStartHere;

	// If we get to the CTxtSite immediately, stop searching!
	if(pNode->HasFlowLayout())
	{
		return NULL;
	}

	for(;;)
	{
		CElement* pElementScope;

		pNode = pNode->Parent();

		if(!pNode)
		{
			return NULL;
		}

		pElementScope = pNode->Element();

		if(pNode->HasFlowLayout())
		{
			return NULL;
		}

		if(pElementScope->IsFlagAndBlock(TAGDESC_LIST))
		{
			return pNode;
		}

		if(pElementScope->IsFlagAndBlock(TAGDESC_LISTITEM))
		{
			return NULL;
		}
	}
}

CTreeNode* CMarkup::SearchBranchForChildOfScope(CTreeNode* pNodeStartHere, CElement* pElementFindChildOfMyScope)
{
	Assert(pNodeStartHere);

	CTreeNode *pNode, *pNodeChild=NULL;

	for(pNode=pNodeStartHere; ; pNode=pNode->Parent())
	{
		if(!pNode)
		{
			return NULL;
		}

		CElement* pElementScope = pNode->Element();

		if(pElementFindChildOfMyScope == pElementScope)
		{
			return pNodeChild;
		}

		if(pNode->HasFlowLayout())
		{
			return NULL;
		}

		pNodeChild = pNode;
	}
}

CTreeNode* CMarkup::SearchBranchForChildOfScopeInStory(CTreeNode* pNodeStartHere, CElement* pElementFindChildOfMyScope)
{
	Assert(pNodeStartHere);

	CTreeNode *pNode, *pNodeChild=NULL;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		CElement* pElementScope = pNode->Element();

		if(pElementFindChildOfMyScope == pElementScope)
		{
			return pNodeChild;
		}

		pNodeChild = pNode;
	}

	return NULL;
}

CTreeNode* CMarkup::SearchBranchForScopeInStory(CTreeNode* pNodeStartHere, CElement* pElementFindMyScope)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		CElement* pElementScope = pNode->Element();

		if(pElementFindMyScope == pElementScope)
		{
			return pNode;
		}
	}

	return NULL;
}

CTreeNode* CMarkup::SearchBranchForScope(CTreeNode* pNodeStartHere, CElement* pElementFindMyScope)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		CElement* pElementScope = pNode->Element();

		if(pElementFindMyScope == pElementScope)
		{
			return pNode;
		}
		if(pNode->HasFlowLayout())
		{
			return NULL;
		}
	}

	Assert(!pNode);

	return NULL;
}

CTreeNode* CMarkup::SearchBranchForNode(CTreeNode* pNodeStartHere, CTreeNode* pNodeFindMe)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; ; pNode=pNode->Parent())
	{
		if(!pNode)
		{
			return NULL;
		}

		if(pNodeFindMe == pNode)
		{
			return pNode;
		}

		if(pNode->HasFlowLayout())
		{
			return NULL;
		}
	}
}

CTreeNode* CMarkup::SearchBranchForNodeInStory(CTreeNode* pNodeStartHere, CTreeNode* pNodeFindMe)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		if(pNodeFindMe == pNode)
		{
			return pNode;
		}
	}

	return NULL;
}

CTreeNode* CMarkup::SearchBranchForTag(CTreeNode* pNodeStartHere, ELEMENT_TAG etag)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode&&etag!=pNode->Tag(); pNode=pNode->Parent())
	{
		if(pNode->HasFlowLayout())
		{
			return NULL;
		}
	}

	return pNode;
}

CTreeNode* CMarkup::SearchBranchForTagInStory(CTreeNode* pNodeStartHere, ELEMENT_TAG etag)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		if(etag == pNode->Tag())
		{
			return pNode;
		}
	}

	return NULL;
}

CTreeNode* CMarkup::SearchBranchForBlockElement(CTreeNode* pNodeStartHere, CFlowLayout* pFLContext)
{
	CTreeNode* pNode;

	Assert(pNodeStartHere);

	if(!pFLContext && GetElementClient())
	{
		pFLContext = GetElementClient()->GetFlowLayout();
	}

	if(!pFLContext)
	{
		return NULL;
	}

	for(pNode=pNodeStartHere; ; pNode=pNode->Parent())
	{
		if(!pNode)
		{
			return NULL;
		}

		CElement* pElementScope = pNode->Element();

		if(pFLContext->IsElementBlockInContext(pElementScope))
		{
			return pNode;
		}

		if(pElementScope == pFLContext->ElementContent())
		{
			return NULL;
		}
	}
}

CTreeNode* CMarkup::SearchBranchForNonBlockElement(CTreeNode* pNodeStartHere, CFlowLayout* pFLContext)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	if(!pFLContext)
	{
		pFLContext = GetElementClient()->GetFlowLayout();
	}

	for(pNode=pNodeStartHere; ; pNode=pNode->Parent())
	{
		if(!pNode)
		{
			return NULL;
		}

		CElement* pElementScope = pNode->Element();

		if(!pFLContext->IsElementBlockInContext(pElementScope))
		{
			return pNode;
		}

		if(pElementScope == pFLContext->ElementOwner())
		{
			return NULL;
		}
	}
}

CTreeNode* CMarkup::SearchBranchForAnchor(CTreeNode* pNodeStartHere)
{
	Assert(pNodeStartHere);

	CTreeNode* pNode = pNodeStartHere;

	for(; ;)
	{
		if(!pNode || pNode->HasFlowLayout())
		{
			return NULL;
		}
		if(pNode->Tag() == ETAG_A)
		{
			break;
		}
		pNode = pNode->Parent();
	}

	return pNode;
}

CTreeNode* CMarkup::SearchBranchForCriteria(
	CTreeNode* pNodeStartHere, BOOL (*pfnSearchCriteria)(CTreeNode*))
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; ; pNode=pNode->Parent())
	{
		if(!pNode)
		{
			return NULL;
		}

		if(pfnSearchCriteria(pNode))
		{
			return pNode;
		}

		if(pNode->HasFlowLayout())
		{
			return NULL;
		}
	}
}

CTreeNode* CMarkup::SearchBranchForCriteriaInStory(
	CTreeNode* pNodeStartHere, BOOL (*pfnSearchCriteria)(CTreeNode*))
{
	Assert(pNodeStartHere);

	CTreeNode* pNode;

	for(pNode=pNodeStartHere; pNode; pNode=pNode->Parent())
	{
		if(pfnSearchCriteria(pNode))
		{
			return pNode;
		}
	}

	return NULL;
}

//+---------------------------------------------------------------
//
//  Member:         CMarkup::EnsureFormats
//
//---------------------------------------------------------------
void CMarkup::EnsureFormats()
{
    // Walk the tree and make sure all formats are computed
    CTreePos* ptpCurr = FirstTreePos();

    while(ptpCurr)
    {
        if(ptpCurr->IsBeginNode())
        {
            ptpCurr->Branch()->EnsureFormats();
        }
        ptpCurr = ptpCurr->NextTreePos();
    }
}

//+---------------------------------------------------------------------------
//
//  Text Frag Service
//
//----------------------------------------------------------------------------
CMarkupTextFragContext* CMarkup::EnsureTextFragContext()
{
    CMarkupTextFragContext* ptfc;

    if(HasTextFragContext())
    {
        return GetTextFragContext();
    }

    ptfc = new CMarkupTextFragContext;
    if(!ptfc)
    {
        return NULL;
    }

    if(SetTextFragContext(ptfc))
    {
        delete ptfc;
        return NULL;
    }

    return ptfc;
}

CMarkupTopElemCache* CMarkup::EnsureTopElemCache()
{
    CMarkupTopElemCache* ptec;

    if(HasTopElemCache())
    {
        return GetTopElemCache();
    }

    ptec = new CMarkupTopElemCache;
    if(!ptec)
    {
        return NULL;
    }

    if(SetTopElemCache(ptec))
    {
        delete ptec;
        return NULL;
    }

    return ptec;
}

//+==============================================================================
// 
// Method: GetSelectionChunksForLayout
//
// Synopsis: Get the 'chunks" for a given Flow Layout, as well as the Max and Min Cp's of the chunks
//              
//            A 'chunk' is a part of a SelectionSegment, broken by FlowLayout
//
//-------------------------------------------------------------------------------
VOID CMarkup::GetSelectionChunksForLayout(
         CFlowLayout* pFlowLayout,
         CPtrAry<HighlightSegment*>* paryHighlight,
         int* piCpMin,
         int* piCpMax)
{
    if(_pSelRenSvcProvider)
    {
        _pSelRenSvcProvider->GetSelectionChunksForLayout(pFlowLayout, paryHighlight, piCpMin, piCpMax);
    }
    else
    {
        Assert(piCpMax);
        *piCpMax = -1;
    }
}

HRESULT CMarkup::EnsureSelRenSvc()
{
    if(!_pSelRenSvcProvider)
    {
        _pSelRenSvcProvider = new CSelectionRenderingServiceProvider(Doc());
    }
    if(!_pSelRenSvcProvider)
    {
        return E_OUTOFMEMORY;
    }
    return S_OK;
}

VOID CMarkup::HideSelection()
{
    if(_pSelRenSvcProvider)
    {
        _pSelRenSvcProvider->HideSelection();
    }
}

VOID CMarkup::ShowSelection()
{
    if(_pSelRenSvcProvider)
    {
        _pSelRenSvcProvider->ShowSelection();
    }
}

VOID CMarkup::InvalidateSelection(BOOL fFireOM)
{
    if(_pSelRenSvcProvider)
    {
        _pSelRenSvcProvider->InvalidateSelection(TRUE, fFireOM);
    }
}

//+----------------------------------------------------------------------------
//
// Member:      EnsureCollectionCache
//
// Synopsis:    Ensures that the collection cache is built
//              NOTE: Ensures the _cache_, not the _collections_.
//
//+----------------------------------------------------------------------------
HRESULT CMarkup::EnsureCollectionCache(long lCollectionIndex)
{
    HRESULT hr;

    hr = InitCollections();
    if(hr)
    {
        goto Cleanup;
    }

    hr = CollectionCache()->EnsureAry(lCollectionIndex);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     AddToCollections
//
//  Synopsis:   add this element to the collection cache
//
//  lNumNestedTables The number of TABLE tags we're nested underneath - speeds
//                  up IMG tag handling
//-------------------------------------------------------------------------
// DEVNOTE rgardner
// This code is tighly couples with CElement::OnEnterExitInvalidateCollections and needs
// to be kept in sync with any changes in that function
HRESULT CMarkup::AddToCollections(CElement* pElement, CCollectionBuildContext* pMonsterWalk)
{
    int         i;
    HRESULT     hr = S_OK;
    CTreeNode*  pNodeForm = NULL;
    LPCTSTR     pszName = NULL;
    LPCTSTR     pszID = NULL;
    CCollectionCache* pCollectionCache;

    // Note here that the outer code has mapped the FRAMES_COLLECTION onto the same Index as
    // the WINDOW_COLLECTION, so both collections get built
    if(!pElement)
    {
        goto Cleanup;
    }

    // Names & ID's are needed by :
    // NAVDOCUMENT_COLLECTION | ANCHORS_COLLECTION
    // This flag is initialized in EnsureCollections()

    // _fIsNamed is always up to date
    if(pMonsterWalk->_fNeedNameID && pElement->IsNamed())
    {
        pElement->FindString(STDPROPID_XOBJ_NAME, &pszName);
        pElement->FindString(DISPID_CElement_id, &pszID);
    }

    // Retrieve the Collection Cache
    pCollectionCache = CollectionCache();

    // Next check to see if element belongs in the window collection
    // Only elements which don't lie inside of forms (and are not forms) belong in here.
    // Things which are not inserted into the form's element collection
    // are also put into the window collection.  E.g. anchors.
    if(pMonsterWalk->_fNeedForm)
    {
        pNodeForm = InFormCollection(pElement->GetFirstBranch());
    }

    if(pMonsterWalk->_lCollection==WINDOW_COLLECTION  && !pNodeForm && pElement->IsNamed())
    {
        LPCTSTR pszUniqueName = NULL;

        if(!pszName && !pszID)
        {
            pElement->FindString(DISPID_CElement_uniqueName, &pszUniqueName);
        }

        if(pszName || pszID || pszUniqueName)
        {
            hr = pCollectionCache->SetIntoAry(WINDOW_COLLECTION, pElement);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    // See if its a FORM within a FORM, Nav doesn't promote names of nested forms to the doc
    if(pMonsterWalk->_fNeedForm && ETAG_FORM==pElement->Tag())
    {
        pNodeForm = pElement->GetFirstBranch()->SearchBranchToRootForTag(ETAG_FORM);
        if(SameScope(pNodeForm, pElement->GetFirstBranch()))
        {
            pNodeForm = NULL;
        }
    }

    // We use the NAVDOCUMENT_COLLECTION to resolve names on the document object
    // If you have a name you get promoted ala Netscape
    // If you have an ID you get promoted, regardless
    // If you have both, you get netscapes rules
    switch(pElement->TagType())
    {
    case ETAG_FORM:
        if(pNodeForm) // only count images, forms not in a form
        {
            break;
        }
        // fallthrough

    case ETAG_IMG:
    case ETAG_APPLET:
    case ETAG_OBJECT:
        if(pMonsterWalk->_lCollection == NAVDOCUMENT_COLLECTION)
        {
            if(pszName)
            {
                hr = pCollectionCache->SetIntoAry(NAVDOCUMENT_COLLECTION, pElement);
                if(hr)
                {
                    goto Cleanup;
                }
            }
            // IE30 compatability OBJECTs/APPLETs with IDs are prmoted to the document
            else if((pElement->Tag()==ETAG_APPLET || pElement->Tag()==ETAG_OBJECT ) && pszID)
            {
                hr = pCollectionCache->SetIntoAry(NAVDOCUMENT_COLLECTION, pElement);
                if(hr)
                {
                    goto Cleanup;
                }
            }
        }
        break;

    case ETAG_INPUT:
        break;
    }

    // See if this element is a region, which means its "container" attribute
    // is "moveable", "in-flow", or "positioned". If so then put it in the
    // region collection, which is used to ensure proper rendering. The BODY
    // should not be added to the region collection.
    if(pMonsterWalk->_lCollection==REGION_COLLECTION &&
        pElement->Tag()!=ETAG_BODY && !pElement->IsPositionStatic())
    {
        hr = pCollectionCache->SetIntoAry(REGION_COLLECTION, pElement);
        if(hr)
        {
            goto Cleanup;
        }
    }

    switch(pElement->TagType())
    {
    case ETAG_LABEL:
        i = LABEL_COLLECTION;
        break;

    case ETAG_FORM:
        if(pNodeForm)
        {
            // nested forms don't go into the forms collection
            goto Cleanup;
        }
        i = FORMS_COLLECTION;
        break;

    case ETAG_AREA:
        AssertSz(FALSE, "must improve");
        goto Cleanup;
        break;

    case ETAG_A:
        // If the A element has an non empty name/id attribute, add it to
        // the anchors' collection
        if(pMonsterWalk->_lCollection == ANCHORS_COLLECTION)
        {
            LPCTSTR lpAnchorName = pszName;
            if(!lpAnchorName)
            {
                lpAnchorName = pszID;
            }
            if(lpAnchorName && lpAnchorName[0])
            {
                hr = pCollectionCache->SetIntoAry(ANCHORS_COLLECTION, pElement);
                if(hr)
                {
                    goto Cleanup;
                }
            }
        }

        if(pMonsterWalk->_lCollection == LINKS_COLLECTION)
        {
            AssertSz(FALSE, "must improve");
        }
        goto Cleanup;

    case ETAG_IMG:
        // The document.images collection in Nav has a quirky bug, for every TABLE
        // above one Nav adds 2^n IMG elements.
        i = IMAGES_COLLECTION;
        break;

    case ETAG_OBJECT:
    case ETAG_APPLET:
        i = APPLETS_COLLECTION;
        break;

    case ETAG_SCRIPT:
        i = SCRIPTS_COLLECTION;
        break;

    default:
        goto Cleanup;
    }

    if(i == pMonsterWalk->_lCollection)
    {
        hr = pCollectionCache->SetIntoAry(i, pElement);
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CMarkup::InitCollections(void)
{
    HRESULT hr = S_OK;
    CAllCollectionCacheItem* pAllCollection = NULL;
    CCollectionCache* pCollectionCache;

    // InitCollections should not be called more than once successfully.
    if(HasCollectionCache())
    {
        goto Cleanup;
    }

    pCollectionCache = new CCollectionCache(
        this,
        _pDoc,
        ENSURE_METHOD(CMarkup, EnsureCollections, ensurecollections));

    if(!pCollectionCache)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    // Initialize from the FORM_COLLECTION onwards,. leaving out the ELEMENT_COLLECTION
    hr = pCollectionCache->InitReservedCacheItems(NUM_DOCUMENT_COLLECTIONS, FORMS_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    pAllCollection = new CAllCollectionCacheItem();
    if(!pAllCollection)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pAllCollection->SetMarkup(this);

    hr = pCollectionCache->InitCacheItem(ELEMENT_COLLECTION, pAllCollection);
    if(hr)
    {
        goto Cleanup;
    }

    // Collection cache now owns this item & is responsible for freeing it

    // Turn off the default name promotion behaviour on collections
    // that don't support it in Nav.
    pCollectionCache->DontPromoteNames(ANCHORS_COLLECTION);
    pCollectionCache->DontPromoteNames(LINKS_COLLECTION);

    // The frames collection resolves ordinal access on the window object
    // so turn off the WINDOW_COLLECTIOn resultion of ordinals
    pCollectionCache->DontPromoteOrdinals(WINDOW_COLLECTION);


    // Because of VBScript compatability issues we create a dynamic type library
    // (See CDoc::BuildObjectTypeInfo()). The dynamic typeinfo contains
    // DISPIDs starting from DISPID_COLLECTION_MIN, & occupying half the DISPID space.


    // We either get Invokes from these DISPIDs or from the WINDOW_COLLECTION GIN/GINEX name resolution
    // OR from the FRAMES_COLLECTION. So we divide the avaliable DISPID range up among these
    // Three 'collections' - 2 real collections & one hand-cooked collection.

    // Divide up the WINDOW_COLLECTION DISPID's so we can tell where the Invoke came from

    // Give the lowest third to the dynamic type library
    // DISPID_COLLECTION_MIN .. (DISPID_COLLECTION_MIN+DISPID_COLLECTION_MAX)/3

    // Give the next third to names resolved on the WINDOW_COLLECTION
    pCollectionCache->SetDISPIDRange(WINDOW_COLLECTION,
        (DISPID_COLLECTION_MIN+DISPID_COLLECTION_MAX)/3+1,
        ((DISPID_COLLECTION_MIN+DISPID_COLLECTION_MAX)*2)/3);

    // In COmWindowProxy::Invoke the security code allows DISPIDs from the frames collection
    // through with no security check. Check that the DISPIDs reserved for this range match up
    Assert(FRAME_COLLECTION_MIN_DISPID == pCollectionCache->GetMaxDISPID(WINDOW_COLLECTION)+1);
    Assert(FRAME_COLLECTION_MAX_DISPID == DISPID_COLLECTION_MAX);

    // Give the final third to the FRAMES_COLLECTION
    pCollectionCache->SetDISPIDRange(FRAMES_COLLECTION,
        pCollectionCache->GetMaxDISPID(WINDOW_COLLECTION)+1,
        DISPID_COLLECTION_MAX);

    // Like NAV, we want to return the last frame that matches the name asked for,
    // rather than returning a collection.
    pCollectionCache->AlwaysGetLastMatchingCollectionItem(FRAMES_COLLECTION);

    // Like NAV, the images doesn't return sub-collections. Note that we might put multiple
    // entries in the IMG collection for the same IMG if the IMG is nested in a table - so
    // this also prevents us from retunring a sub-collection in this situation.
    pCollectionCache->AlwaysGetLastMatchingCollectionItem(IMAGES_COLLECTION);

    // Setup the lookaside variable
    hr = SetLookasidePtr(LOOKASIDE_COLLECTIONCACHE, pCollectionCache);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Member:     EnsureCollections
//
//  Synopsis:   rebuild the collections based on the current state of the tree
//
//-----------------------------------------------------------------------------
HRESULT CMarkup::EnsureCollections(long lIndex, long* plCollectionVersion)
{
    HRESULT hr = S_OK;
    CCollectionBuildContext collectionWalker(_pDoc);
    long l, lSize;
    CElement* pElemCurrent;
    CCollectionCache* pCollectionCache;

    pCollectionCache = CollectionCache();
    if(!pCollectionCache)
    {
        goto Cleanup;
    }

    // Optimize the use of the region collection. The doc flag is set by
    // the CSS code whenever theres a position: attribute on an element in the doc
    // This is a temporary Beta1 Hack to avoid building the regions collection
    if(lIndex==REGION_COLLECTION && !_pDoc->_fRegionCollection) //$$tomfakes - move this flag to CMarkup?
    {
        pCollectionCache->ResetAry(REGION_COLLECTION); // To be safe
        goto Cleanup;
    }


    if(lIndex == ELEMENT_COLLECTION)
    {
        goto Update; // All collection is always up to date, update the version no & bail out
    }
    else if(lIndex == REGION_COLLECTION)
    {
        // We ignore the collection _fIsValid flag for the REGION_COLLECTION
        // because it is unaffected by element name changes
        if(*plCollectionVersion == _pDoc->GetDocTreeVersion())
        {
            goto Cleanup;
        }
    }
    else
    {
        // Collections that are specificaly invalidated
        if(pCollectionCache->IsValidItem(lIndex))
        {
            // Doesn't change collection version, collections based on this one don't get rebuilt
            goto Cleanup; 
        }
    }

    collectionWalker._lCollection = lIndex;

    // Set flag to indicate whether or not we need to go get the name/ID of
    // elements during the walk
    if(lIndex==NAVDOCUMENT_COLLECTION || lIndex==ANCHORS_COLLECTION || lIndex==WINDOW_COLLECTION)
    {
        collectionWalker._fNeedNameID = TRUE;
    }
    // Set flag to indicate whether or not we need to go get the containing form of
    // elements during the walk
    if(lIndex==WINDOW_COLLECTION || lIndex==NAVDOCUMENT_COLLECTION ||
        lIndex==FORMS_COLLECTION || lIndex==FRAMES_COLLECTION)
    {
        collectionWalker._fNeedForm = TRUE;
    }

    // Here we blow away any cached information which is stored in the doc.
    // Usually elements, themselves, will blow away their own cached info
    // when they are first visited, but because the doc is not an element,
    // it will not be visited, and must blow away the cached info before we
    // start the walk.

    // every fixed collection is based on the all collection
    if(lIndex != ELEMENT_COLLECTION)
    {
        pCollectionCache->ResetAry(lIndex);
    }

    // Otherwise all collection is up to date, so use it because its faster
    lSize = pCollectionCache->SizeAry(ELEMENT_COLLECTION);
    for(l=0; l<lSize; l++)
    {
        hr = pCollectionCache->GetIntoAry(
            ELEMENT_COLLECTION,
            l,
            &pElemCurrent);
        if(hr)
        {
            goto Cleanup;
        }
        hr = AddToCollections(pElemCurrent, &collectionWalker);
        if(hr)
        {
            goto Cleanup;
        }
    }

Update:
    // Update the version on the collection
    if(lIndex==REGION_COLLECTION || lIndex==ELEMENT_COLLECTION)
    {
        // Collection derived from tree
        *plCollectionVersion = _pDoc->GetDocTreeVersion();
    }
    else
    {
        (*plCollectionVersion)++;
    }

    pCollectionCache->ValidateItem(lIndex);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetCollection
//
//  Synopsis:   return collection based on index in the collection cache
//
//-------------------------------------------------------------------------
HRESULT CMarkup::GetCollection(int iIndex, IHTMLElementCollection** ppdisp)
{
    Assert((iIndex>=0) && (iIndex<NUM_DOCUMENT_COLLECTIONS));

    HRESULT hr;

    if(!ppdisp)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ppdisp = NULL;

    hr = EnsureCollectionCache(iIndex);
    if(hr)
    {
        goto Cleanup;
    }

    hr = CollectionCache()->GetDisp(iIndex, (IDispatch**)ppdisp);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetElementByNameOrID
//
//  Synopsis:   look up an Element by its Name or ID
//
//  Returns:    S_OK, if it found the element.  *ppElement is set
//              S_FALSE, if multiple elements w/ name were found.
//                  *ppElement is set to the first element in list.
//              Other errors.
//-------------------------------------------------------------------------
HRESULT CMarkup::GetElementByNameOrID(LPTSTR szName, CElement** ppElement)
{
    HRESULT hr;
    CElement* pElemTemp;

    hr = EnsureCollectionCache(ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    hr = CollectionCache()->GetIntoAry(
        ELEMENT_COLLECTION,
        szName,
        FALSE,
        &pElemTemp);
    *ppElement = pElemTemp;

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     GetDispByNameOrID
//
//  Synopsis:   Retrieve an IDispatch by its Name or ID
//
//  Returns:    An IDispatch* to an element or a collection of elements.
//
//-------------------------------------------------------------------------
HRESULT CMarkup::GetDispByNameOrID(LPTSTR szName, IDispatch** ppDisp, BOOL fAlwaysCollection)
{
    HRESULT hr;

    hr = EnsureCollectionCache(ELEMENT_COLLECTION);
    if(hr)
    {
        goto Cleanup;
    }

    hr = CollectionCache()->GetDisp(
        ELEMENT_COLLECTION,
        szName,
        CacheType_Named,
        ppDisp,
        FALSE, // Case Insensitive
        NULL,
        fAlwaysCollection); // Always return a collection (even if empty, or has 1 elem)

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  member : InFormCollection
//
//  Synopsis : helper function for determining if an element will appear in a
//     form element's collecion of if it should be scoped to the document's all
//     collection. this is also used by buildTypeInfo for hooking up VBSCRIPT
//
//----------------------------------------------------------------------------
CTreeNode* CMarkup::InFormCollection(CTreeNode* pNode)
{
    CTreeNode* pNodeParentForm = NULL;

    if(!pNode)
    {
        goto Cleanup;
    }

    // NOTE: Forms are now promoted to the window, this is because the AddNamedItem
    //       no longer adds the form and we need to have access to the form for
    //       <SCRIPT FOR EVENT>
    switch(pNode->TagType())
    {
    case ETAG_INPUT:
    case ETAG_FIELDSET:
    case ETAG_SELECT:
    case ETAG_TEXTAREA:
    case ETAG_IMG:
    case ETAG_BUTTON:
    case ETAG_OBJECT:
        // COMMENT rgardner - A good optimization here would be to have the tree walker
        // retain the last scoping form element & pass in the pointer
        pNodeParentForm = pNode->SearchBranchToRootForTag(ETAG_FORM);
        break;
    }

Cleanup:
    return pNodeParentForm;
}

void* CMarkup::GetLookasidePtr(int iPtr)
{
	return (HasLookasidePtr(iPtr)?Doc()->GetLookasidePtr((DWORD*)this+iPtr):NULL);
}

HRESULT CMarkup::SetLookasidePtr(int iPtr, void* pvVal)
{
	Assert(!HasLookasidePtr(iPtr) && "Can't set lookaside ptr when the previous ptr is not cleared");

	HRESULT hr = Doc()->SetLookasidePtr((DWORD *)this+iPtr, pvVal);

	if(hr == S_OK)
	{
		_fHasLookasidePtr |= 1 << iPtr;
	}

	RRETURN(hr);
}

void* CMarkup::DelLookasidePtr(int iPtr)
{
	if(HasLookasidePtr(iPtr))
	{
		void* pvVal = Doc()->DelLookasidePtr((DWORD*)this+iPtr);
		_fHasLookasidePtr &= ~(1<<iPtr);
		return pvVal;
	}

	return NULL;
}

void CMarkup::ClearLookasidePtrs()
{
	delete DelCollectionCache();
	Assert(_pDoc->GetLookasidePtr((DWORD*)this+LOOKASIDE_COLLECTIONCACHE) == NULL);

	CMarkup* pParentMarkup = (CMarkup*)DelParentMarkup();

	if(pParentMarkup)
	{
		pParentMarkup->Release();
	}

	Assert(_pDoc->GetLookasidePtr((DWORD*)this+LOOKASIDE_PARENTMARKUP) == NULL);
}

//+---------------------------------------------------------------------------
//
//  Member:     CMarkup::PrivateQueryInterface
//
//----------------------------------------------------------------------------
HRESULT CMarkup::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT     hr = S_OK;
    const void* apfn = NULL;
    void*       pv = NULL;
    void*       appropdescsInVtblOrder = NULL;
    const IID* const* apIID = NULL;
    BOOL fPrimaryMarkup = IsPrimaryMarkup();

    *ppv = NULL;

    switch(iid.Data1)
    {
        QI_INHERITS((IPrivateUnknown*)this, IUnknown)
        QI_TEAROFF(this, IMarkupContainer, NULL)
        QI_TEAROFF(this, ISelectionRenderingServices, NULL)
        QI_TEAROFF2(this, ISegmentList, ISelectionRenderingServices, NULL)

    default:
        if(IsEqualIID(iid, CLSID_CMarkup))
        {
            *ppv = this;
            return S_OK;
        }
        else if(iid==IID_IHTMLDocument || iid==IID_IHTMLDocument2)
        {
            if(fPrimaryMarkup)
            {
                pv = (IHTMLDocument2*)Doc();
                apfn = *(void**)pv;
            }
            else
            {
                pv = (void*)this;
                apfn = (const void*)s_apfnpdIHTMLDocument2;
                appropdescsInVtblOrder = (void*)s_ppropdescsInVtblOrderIHTMLDocument2; 
            }
        }
        //else if(iid == IID_IHTMLDocument3)
        //{
        //    if(fPrimaryMarkup)
        //    {
        //        pv = (IHTMLDocument3*)Doc();
        //        apfn = *(void**)pv;
        //    }
        //    else
        //    {
        //        pv = (void*)this;
        //        apfn = (const void*)s_apfnpdIHTMLDocument3;
        //        appropdescsInVtblOrder = (void *)s_ppropdescsInVtblOrderIHTMLDocument3; 
        //    }
        //}
        //else if(iid == IID_IConnectionPointContainer)
        //{
        //    *((IConnectionPointContainer**)ppv) = new CConnectionPointContainer(
        //        fPrimaryMarkup?Doc():(CBase*)this,
        //        (IUnknown*)(IPrivateUnknown*)this);
        //    if(!*ppv)
        //    {
        //        RRETURN(E_OUTOFMEMORY);
        //    }
        //}
        else if(iid == IID_IMarkupServices)
        {
            pv = Doc();
            apfn = CDocument::s_apfnIMarkupServices;
        }
        else if(iid == IID_IHTMLViewServices)
        {
            pv = Doc();
            apfn = CDocument::s_apfnIHTMLViewServices;
        }
        //else if(iid == IID_IServiceProvider)
        //{
        //    pv = Doc();
        //    apfn = CDocument::s_apfnIServiceProvider;
        //}
        //else if(iid == IID_IOleWindow)
        //{
        //    pv = Doc();
        //    apfn = CDocument::s_apfnIOleInPlaceObjectWindowless;
        //}
        else if(iid == IID_IOleCommandTarget)
        {
            pv = Doc();
            apfn = CDocument::s_apfnIOleCommandTarget;
        }
        else
        {
            RRETURN(super::PrivateQueryInterface(iid, ppv));
        }

        if(pv)
        {
            Assert(apfn);
            hr = CreateTearOffThunk(
                pv, 
                apfn, 
                NULL, 
                ppv, 
                (IUnknown*)(IPrivateUnknown*)this, 
                *(void**)(IUnknown*)(IPrivateUnknown*)this,
                QI_MASK|ADDREF_MASK|RELEASE_MASK,
                apIID,
                appropdescsInVtblOrder);
            if(hr)
            {
                RRETURN(hr);
            }
        }
    }

    // any unknown interface will be handled by the default above
    Assert(*ppv);
    (*(IUnknown**)ppv)->AddRef();

    return S_OK;
}

HRESULT CMarkup::MovePointersToSegment(int iSegmentIndex, IMarkupPointer* pIStart, IMarkupPointer* pIEnd) 
{
    HRESULT hr = EnsureSelRenSvc();

    if(!hr)
    {
        hr = _pSelRenSvcProvider->MovePointersToSegment(iSegmentIndex, pIStart, pIEnd);
    }

    RRETURN(hr);
}

HRESULT CMarkup::GetSegmentCount(int* piSegmentCount, SELECTION_TYPE* peSegmentType)
{
    HRESULT hr = EnsureSelRenSvc();
    if(!hr)
    {
        hr = _pSelRenSvcProvider->GetSegmentCount(piSegmentCount, peSegmentType);
    }

    RRETURN(hr);
}

HRESULT CMarkup::AddSegment(
        IMarkupPointer* pIStart,
        IMarkupPointer* pIEnd,
        HIGHLIGHT_TYPE HighlightType,
        int* piSegmentIndex)
{
    HRESULT hr = EnsureSelRenSvc();

    if(OK(hr))
    {
        hr = _pSelRenSvcProvider->AddSegment(pIStart, pIEnd, HighlightType, piSegmentIndex);
    }

    RRETURN(hr);
}

HRESULT CMarkup::AddElementSegment(IHTMLElement* pIElement, int* piSegmentIndex)
{
    HRESULT hr = EnsureSelRenSvc();

    if(!hr)
    {
        hr =_pSelRenSvcProvider->AddElementSegment(pIElement, piSegmentIndex);        
    }

    RRETURN(hr);
}

HRESULT CMarkup::MoveSegmentToPointers( 
        int iSegmentIndex,
        IMarkupPointer* pIStart, 
        IMarkupPointer* pIEnd,
        HIGHLIGHT_TYPE HighlightType)
{
    HRESULT hr = EnsureSelRenSvc();

    if(!hr)
    {
        hr = _pSelRenSvcProvider->MoveSegmentToPointers(iSegmentIndex, pIStart, pIEnd, HighlightType);
    }

    RRETURN(hr);
}

HRESULT CMarkup::GetElementSegment(int iSegmentIndex, IHTMLElement** ppIElement) 
{
    HRESULT hr = EnsureSelRenSvc();

    if(!hr)
    {
        hr =  _pSelRenSvcProvider->GetElementSegment(iSegmentIndex, ppIElement);
    }

    RRETURN(hr);
}

HRESULT CMarkup::SetElementSegment(int iSegmentIndex, IHTMLElement* pIElement)
{
    HRESULT hr = EnsureSelRenSvc();

    if(!hr)
    {
        hr = _pSelRenSvcProvider->SetElementSegment(iSegmentIndex, pIElement);
    }

    RRETURN(hr);
}

HRESULT CMarkup::ClearSegment(int iSegmentIndex, BOOL fInvalidate)
{
    if(_pSelRenSvcProvider)
    {
        RRETURN(_pSelRenSvcProvider->ClearSegment(iSegmentIndex, fInvalidate));
    }
    else
    {
        return S_OK;
    }
}

HRESULT CMarkup::ClearSegments(BOOL fInvalidate)
{
    if(_pSelRenSvcProvider)
    {
        RRETURN(_pSelRenSvcProvider->ClearSegments(fInvalidate));
    }
    else
    {
        return S_OK;
    }
}

HRESULT CMarkup::ClearElementSegments()
{
    if(_pSelRenSvcProvider)
    {
        RRETURN(_pSelRenSvcProvider->ClearElementSegments());     
    }
    else
    {
        return S_OK;
    }
}

HRESULT CMarkup::OwningDoc(IHTMLDocument2** ppDoc)
{
    return _pDoc->_pPrimaryMarkup->QueryInterface(IID_IHTMLDocument2, (void**)ppDoc);
}

void CMarkup::ReplacePtr(CMarkup** pplhs, CMarkup* prhs)
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

void CMarkup::SetPtr(CMarkup** pplhs, CMarkup* prhs)
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

void CMarkup::StealPtrSet(CMarkup** pplhs, CMarkup* prhs)
{
	SetPtr(pplhs, prhs);

	if(pplhs && *pplhs)
	{
		(*pplhs)->Release();
	}
}

void CMarkup::StealPtrReplace(CMarkup** pplhs, CMarkup* prhs)
{
	ReplacePtr(pplhs, prhs);

	if(pplhs && *pplhs)
	{
		(*pplhs)->Release();
	}
}

void CMarkup::ClearPtr(CMarkup** pplhs)
{
	if(pplhs && *pplhs)
	{
		CMarkup* pElement = *pplhs;
		*pplhs = NULL;
		pElement->Release();
	}
}

void CMarkup::ReleasePtr(CMarkup* pMarkup)
{
	if(pMarkup)
	{
		pMarkup->Release();
	}
}

CMarkup::CLock::CLock(CMarkup* pMarkup)
{
	Assert(pMarkup);
	_pMarkup = pMarkup;
	pMarkup->AddRef();
}

CMarkup::CLock::~CLock()
{
	_pMarkup->Release();
}

//+---------------------------------------------------------------------------
//
//  Member:     CMarkup::EnsureStyleSheets
//
//  Synopsis:   Ensure the stylesheets collection exists, creates it if not..
//
//----------------------------------------------------------------------------
HRESULT CMarkup::EnsureStyleSheets()
{
    AssertSz(FALSE, "must improve");
    return S_OK;
}

HRESULT CMarkup::ApplyStyleSheets(
          CStyleInfo*   pStyleInfo,
          ApplyPassType passType,
          BOOL*         pfContainsImportant)
{
    HRESULT hr = S_OK;
    //AssertSz(FALSE, "must improve");
    RRETURN(hr);
}
