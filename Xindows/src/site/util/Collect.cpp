
#include "stdafx.h"
#include "Collect.h"

#define _cxx_
#include "../gen/collect.hdl"

//====================================================================
//
//  Class CElementCollectionBase, CElementCollection methods
//
//===================================================================

//+------------------------------------------------------------------------
//
//  Member:     s_classdesc
//
//  Synopsis:   class descriptor
//
//-------------------------------------------------------------------------
const CBase::CLASSDESC CElementCollectionBase::s_classdesc =
{
    0,                              // _pclsid
    NULL,                           // _pcpi
    0,                              // _dwFlags
    NULL,                           // _piidDispinterface
    &s_apHdlDescs,                  // _apHdlDesc
};

//+------------------------------------------------------------------------
//
//  Member:     s_classdesc
//
//  Synopsis:   class descriptor
//
//-------------------------------------------------------------------------
const CBase::CLASSDESC CElementCollection::s_classdesc =
{
    0,                              // _pclsid
    NULL,                           // _pcpi
    0,                              // _dwFlags
    &IID_IHTMLElementCollection,    // _piidDispinterface
    &s_apHdlDescs,                  // _apHdlDesc
};


//+------------------------------------------------------------------------
//
//  Member:     CElementCollectionBase
//
//  Synopsis:   constructor
//
//-------------------------------------------------------------------------
CElementCollectionBase::CElementCollectionBase(CCollectionCache* pCollectionCache, long lIndex)
: super(), _pCollectionCache(pCollectionCache), _lIndex(lIndex)
{
    // Tell the base to live longer than us
    _pCollectionCache->GetBase()->SubAddRef();
}

//+------------------------------------------------------------------------
//
//  Member:     ~CElementCollectionBase
//
//  Synopsis:   destructor
//
//-------------------------------------------------------------------------
CElementCollectionBase::~CElementCollectionBase()
{
    // release subobject count
    _pCollectionCache->GetBase()->SubRelease();
}

//+------------------------------------------------------------------------
//
//  Member:     ~CElementCollection
//
//  Synopsis:   destructor
//
//-------------------------------------------------------------------------
CElementCollection::~CElementCollection()
{
    _pCollectionCache->ClearDisp(_lIndex);
}

//+------------------------------------------------------------------------
//
//  Member:     PrivateQueryInterface
//
//  Synopsis:   vanilla implementation
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::PrivateQueryInterface(REFIID iid, void** ppv)
{
    *ppv = NULL;
    switch(iid.Data1)
    {
        QI_INHERITS((IHTMLElementCollection*)this, IUnknown)
        QI_TEAROFF(this, IHTMLElementCollection2, NULL)

    default:
        if(iid == CLSID_CElementCollection)
        {
            *ppv = this;
            return S_OK;
        }

        if(iid == IID_IHTMLElementCollection)
        {
            *ppv = (IHTMLElementCollection*)this;
        }
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    ((IUnknown*)*ppv)->AddRef();

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CloseErrorInfo
//
//  Synopsis:   defer to the collection cache
//
//-------------------------------------------------------------------------
HRESULT CElementCollectionBase::CloseErrorInfo(HRESULT hr)
{
    return _pCollectionCache->CloseErrorInfo(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     get_length
//
//  Synopsis:   collection object model, defers to Cache Helper
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::get_length(long* plSize)
{
    RRETURN(SetErrorInfo(_pCollectionCache->GetLength(_lIndex, plSize)));
}

//+------------------------------------------------------------------------
//
//  Member:     put_length
//
//  Synopsis:   collection object model, defers to Cache Helper
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::put_length(long lSize)
{
    RRETURN(SetErrorInfo(E_NOTIMPL));
}

//+------------------------------------------------------------------------
//
//  Member:     item
//
//  Synopsis:   collection object model
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::item(VARIANTARG var1, VARIANTARG var2, IDispatch** ppResult)
{
    RRETURN(SetErrorInfo(_pCollectionCache->Item(_lIndex, var1, var2, ppResult)));
}

//+------------------------------------------------------------------------
//
//  Member:     tags
//
//  Synopsis:   collection object model, this always returns a collection
//              and is named based on the tag, and searched based on tagname
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::tags(VARIANT var1, IDispatch** ppdisp)
{
    RRETURN(SetErrorInfo(_pCollectionCache->Tags(_lIndex, var1, ppdisp)));
}

HRESULT CElementCollection::Tags(LPCTSTR szTagName, IDispatch** ppdisp)
{
    RRETURN(_pCollectionCache->Tags(_lIndex, szTagName, ppdisp));
}

//+------------------------------------------------------------------------
//
//  Member:     urns
//
//  Synopsis:   collection object model, this always returns a collection
//              and is named based on the urn, and searched based on urn
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::urns(VARIANT var1, IDispatch** ppdisp)
{
    RRETURN(SetErrorInfo(_pCollectionCache->Urns(_lIndex, var1, ppdisp)));
}

//+------------------------------------------------------------------------
//
//  Member:     Get_newEnum
//
//  Synopsis:   collection object model
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::get__newEnum(IUnknown** ppEnum)
{
    RRETURN(SetErrorInfo(_pCollectionCache->GetNewEnum(_lIndex, ppEnum)));
}

//+------------------------------------------------------------------------
//
//  Member:     toString
//
//  Synopsis:   This is impplemented on all objects
//
//-------------------------------------------------------------------------
HRESULT CElementCollection::toString(BSTR* String)
{
    RRETURN(super::toString(String));
};

//====================================================================
//
//  Class CCollectionCache Methods:
//
//====================================================================

//+------------------------------------------------------------------------
//
//  Member:     ~CCollectionCache
//
//  Synopsis:   Constructor
//
//-------------------------------------------------------------------------
CCollectionCache::CCollectionCache(
           CBase* pBase,
           CDocument* pDoc,
           PFN_CVOID_ENSURE pfnEnsure/*=NULL*/,
           PFN_CVOID_CREATECOL pfnCreation/*=NULL*/,
           PFN_CVOID_REMOVEOBJECT pfnRemove/*=NULL*/,    
           PFN_CVOID_ADDNEWOBJECT pfnAddNewObject/*=NULL*/) :
_pBase(pBase), 
_pfnEnsure(pfnEnsure),     
_pfnRemoveObject(pfnRemove), 
_lReservedSize(0),
_pfnAddNewObject(pfnAddNewObject),
_pDoc(pDoc),   
_pfnCreateCollection(pfnCreation)
{
    Assert(pBase);  // Required.
}

//+------------------------------------------------------------------------
//
//  Member:     ~CCollectionCache
//
//  Synopsis:   Destructor
//
//-------------------------------------------------------------------------
CCollectionCache::~CCollectionCache()
{
    CacheItem* pce = _aryItems;
    UINT cSize = _aryItems.Size();

    for(; cSize--; ++pce)
    {
        delete pce->_pCacheItem;
    }
}

//+------------------------------------------------------------------------
//
//  Member:     Init
//
//  Synopsis:   Setup the cache.  This call is required if part of the
//              cache is to be reserved.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::InitReservedCacheItems(long lReservedSize, long lFromIndex/*=0*/, long lIdentityIndex/*=-1*/)
{
    HRESULT hr = E_INVALIDARG;
    CacheItem* pce = NULL;
    long l;

    Assert(lReservedSize > 0);
    Assert(lFromIndex<=lReservedSize && lFromIndex>=0);
    Assert(lIdentityIndex==-1 || (lIdentityIndex>=0 && lIdentityIndex<_lReservedSize));

    // Clear the reserved part of the cache.
    hr = _aryItems.EnsureSize(lReservedSize);
    if(hr)
    {
        goto Cleanup;
    }

    pce = _aryItems;

    memset(_aryItems, 0, lReservedSize*sizeof(CacheItem));
    _aryItems.SetSize(lReservedSize);

    // Reserved items always use the CCollectionCacheItem
    for(l=lFromIndex,pce=_aryItems+lFromIndex; l<lReservedSize; l++,++pce)
    {
        pce->Init();
        pce->_pCacheItem = new CElementAryCacheItem();
        if(pce->_pCacheItem == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }
        MemSetName((pce->_pCacheItem, "CacheItem"));
    }

    if(lIdentityIndex != -1) 
    {
        _aryItems[lIdentityIndex].fIdentity = TRUE;
    }

Cleanup:
    if(hr)
    {
        for(l=_aryItems.Size(); l>=0 ; l--)
        {
            delete pce->_pCacheItem;
        }
        _aryItems.SetSize(0);
    }
    else
    {
        _lReservedSize = lReservedSize;
    }
    RRETURN(hr);
}

HRESULT CCollectionCache::InitCacheItem(long lCacheIndex, CCollectionCacheItem* pCacheItem)
{
    Assert(pCacheItem);
    Assert(lCacheIndex>=0 && lCacheIndex<_aryItems.Size());
    Assert(!_aryItems[lCacheIndex]._pCacheItem); // better not be initialized all ready

    _aryItems[lCacheIndex].Init();
    _aryItems[lCacheIndex]._pCacheItem = pCacheItem;

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     UnloadContents
//
//  Synopsis : called when the contents of the cache need to be freed up
//      but the cache itself needs to stick around
//
//+------------------------------------------------------------------------
HRESULT CCollectionCache::UnloadContents()
{
    long l;

    for(l=_aryItems.Size()-1; l>=0; l--)
    {
        _aryItems[l]._lCollectionVersion = 0;

        if(_aryItems[l]._pCacheItem)
        {
            _aryItems[l]._pCacheItem->DeleteContents();
        }
    }

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CreateCollectionHelper
//
//  Synopsis:   Wrapper for private member, if no CreateCollection function is
//      provided, default to creating a CElementCollectionBase. If a create function
//      IS provided, then the base object has a derived collection that it wants
//
//      This returns IDispatch because that is the form that the callers need
//---------------------------------------------------------------------------
HRESULT CCollectionCache::CreateCollectionHelper(IDispatch** ppIEC, long lIndex)
{
    HRESULT hr = S_OK;

    *ppIEC = NULL;

    if(_pfnCreateCollection)
    {
        hr = CALL_METHOD((CVoid*)(void*)_pBase, _pfnCreateCollection, (ppIEC, lIndex));
    }
    else
    {
        CElementCollection* pobj;

        pobj = new CElementCollection(this, lIndex);
        if(!pobj)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = pobj->QueryInterface(IID_IDispatch, (void**)ppIEC);
        pobj->Release();
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
//  Member:     GetIntoAry
//
//  Synopsis:   Return the element at the specified index.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetIntoAry(long lIndex, long lIndexElement, CElement** ppElem)
{
    if(lIndexElement >= _aryItems[lIndex]._pCacheItem->Length())
    {
        RRETURN(DISP_E_MEMBERNOTFOUND);
    }

    if(lIndexElement < 0)
    {
        RRETURN(E_INVALIDARG);
    }

    *ppElem = _aryItems[lIndex]._pCacheItem->GetAt(lIndexElement);

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     GetIntoAry
//
//  Synopsis:   Return the index of the specified element.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetIntoAry(long lIndex, CElement* pElement, long* plIndexElement)
{
    CCollectionCacheItem* pCacheItem = _aryItems[lIndex]._pCacheItem;
    CElement* pElem;
    long lElemIndex = 0;

    Assert(plIndexElement);

    *plIndexElement = -1;

    // Using GetNext() should be more efficient than using GetAt();
    pCacheItem->MoveTo(0);
    do
    {
        pElem = pCacheItem->GetNext();
        if(pElem && pElem==pElement)
        {
            *plIndexElement = lElemIndex;
            break;
        }
        lElemIndex++;
    } while(pElem);


    RRETURN(*plIndexElement==-1 ? DISP_E_MEMBERNOTFOUND : S_OK);
}

//+------------------------------------------------------------------------
//
//  Member:     GetIntoAry
//
//  Synopsis:   Return the element with a given name
//
//  Returns:    S_OK, if it found the element.  *ppNode is set
//              S_FALSE, if multiple elements w/ name were found.
//                  *ppNode is set to the first element in list.
//              Other errors.
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetIntoAry(
         long lCollectionIndex,
         LPCTSTR Name,
         BOOL fTagName,
         CElement** ppElem,
         long iStartFrom/*=0*/,
         BOOL fCaseSensitive/*=FALSE*/)
{
    HRESULT                 hr = S_OK;
    CElement*               pElemFirstMatched = NULL;
    CElement*               pElemLastMatched = NULL;
    CElement*               pElem;
    CCollectionCacheItem*   pCacheItem = _aryItems[lCollectionIndex]._pCacheItem;
    BOOL                    fLastOne = FALSE;

    Assert(ppElem);
    *ppElem = NULL;

    if(iStartFrom == -1)
    {
        iStartFrom = 0;
        fLastOne = TRUE;
    }

    for(pCacheItem->MoveTo(iStartFrom); ;)
    {
        pElem = pCacheItem->GetNext();
        if(!pElem)
        {
            break;
        }
        if(CompareName(pElem, Name, fTagName, fCaseSensitive))
        {
            if(!pElemFirstMatched)
            {
                pElemFirstMatched = pElem;
            }
            pElemLastMatched = pElem;
        }
    }

    if(!pElemLastMatched)
    {
        hr = DISP_E_MEMBERNOTFOUND;

        if(pCacheItem->IsRangeSyntaxSupported())
        {
            CCellRangeParser cellRangeParser(Name);
            if(!cellRangeParser.Failed())
            {
                hr = S_FALSE; // for allowing expando and properties/methods on the collection (TABLE_CELL_COLLECTION)
            }
        }

        goto Cleanup;
    }

    // A collection can be marked to always return the last matching name,
    // rather than the default first matching name.
    if(DoGetLastMatchedName(lCollectionIndex))
    {
        *ppElem = pElemLastMatched;
    }
    else
    {
        // The iStartFrom has higher precedence on which element we really
        // return first or last in the collection.
        *ppElem = fLastOne ? pElemLastMatched : pElemFirstMatched;
    }

    // return S_FALSE if we have more than one element that matced
    hr = (pElemFirstMatched==pElemLastMatched ) ? S_OK : S_FALSE;

Cleanup:
    RRETURN1(hr, S_FALSE);
}

// Invalid smart collections - one using the _fIsValid technique
void CCollectionCache::InvalidateAllSmartCollections(void)
{
    long l;
    for(l=_aryItems.Size()-1; l>=0; l--)
    {
        _aryItems[l]._fIsValid = FALSE;
    }
}

// Invalidate "dump" collections - ones using old-style version management
void CCollectionCache::Invalidate(void)
{
    long l;
    for(l=_aryItems.Size()-1; l>=0 ; l--)
    {
        _aryItems[l]._lCollectionVersion = 0;
    }
}

//+------------------------------------------------------------------------
//
//  Member:     EnsureAry
//
//  Synopsis:   Make sure this index is ready for access.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::EnsureAry(long lIndex)
{
    HRESULT hr = S_OK;

    if(_pfnEnsure && lIndex<_lReservedSize)
    {
        // Ensure the reserved part of the collection
        hr = CALL_METHOD((CVoid*)(void*)_pBase, _pfnEnsure, (lIndex, &_aryItems[lIndex]._lCollectionVersion));
        if(hr)
        {
            goto Cleanup;
        }
    }

    // Ensuring a dynamic collection, need to make sure all its dependent collections
    // are ensured
    if(lIndex >= _lReservedSize)
    {
        // Ensure the collection we're based upon
        // note that this is a recursove call
        hr = EnsureAry(_aryItems[lIndex].sIndex);
        if(hr)
        {
            goto Cleanup;
        }

        // If we're a different version than the collection we're based upon, rebuild ourselves now
        if(_aryItems[lIndex]._lCollectionVersion != _aryItems[_aryItems[lIndex].sIndex]._lCollectionVersion)
        {
            switch(_aryItems[lIndex].Type)
            {
            case CacheType_Tag:
                // Rebuild based on name
                hr = BuildNamedArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].cstrName,
                    TRUE,
                    _aryItems[lIndex]._pCacheItem,
                    0,
                    _aryItems[lIndex].fIsCaseSensitive);
                break;

            case CacheType_Named:
                // Rebuild based on tag name
                hr = BuildNamedArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].cstrName,
                    FALSE,
                    _aryItems[lIndex]._pCacheItem,
                    0,
                    _aryItems[lIndex].fIsCaseSensitive);
                break;

            case CacheType_Children:
            case CacheType_DOMChildNodes:
                // Rebuild all children
                hr = BuildChildArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].pElementBase,
                    _aryItems[lIndex]._pCacheItem,
                    FALSE);
                break;

            case CacheType_AllChildren:
                // Rebuild all children
                hr = BuildChildArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].pElementBase,
                    _aryItems[lIndex]._pCacheItem,
                    TRUE);
                break;

            case CacheType_CellRange:
                // Rebuild cells acollection
                hr = BuildCellRangeArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].cstrName,
                    &(_aryItems[lIndex].rectCellRange),
                    _aryItems[lIndex]._pCacheItem);
                break;

            case CacheType_FreeEntry:
                // Free collection waiting to be reused
                break;

            case CacheType_Urn:
                hr = BuildNamedArray(
                    _aryItems[lIndex].sIndex,
                    _aryItems[lIndex].cstrName,
                    TRUE,
                    _aryItems[lIndex]._pCacheItem,
                    0,
                    _aryItems[lIndex].fIsCaseSensitive,
                    TRUE);
                break;

            default:
                Assert(0);
                break;
            }
        }
        _aryItems[lIndex]._lCollectionVersion = _aryItems[_aryItems[lIndex].sIndex]._lCollectionVersion;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     ClearDisp
//
//  Synopsis:   Clear the reference to this dispatch ptr out of the cache.
//
//              At this point,  any reference that the disp (collection)
//                may have (due to being a named collection) on other collections
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::ClearDisp(long lIndex)
{
    Assert(lIndex>=0 && lIndex<_aryItems.Size());

    // Find the collection in the cache & clear it down
    _aryItems[lIndex].pdisp = NULL;
    // Mark it as free
    _aryItems[lIndex].Type = CacheType_FreeEntry;

    if(lIndex >= _lReservedSize)
    {
        if(_aryItems[lIndex]._pCacheItem)
        {
            short sDepend;

            delete _aryItems[lIndex]._pCacheItem;
            _aryItems[lIndex]._pCacheItem = NULL;

            sDepend = _aryItems[lIndex].sIndex;
            if(sDepend >= _lReservedSize)
            {
                _aryItems[sDepend].pdisp->Release();
            }
        }
        _aryItems[lIndex].cstrName.Free();
    }
    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     GetDisp
//
//  Synopsis:   Get a dispatch ptr from the RESERVED part of the cache.
//              If fIdentity is not set, then everything works as planned.
//              If set, then we QI the base Object and return that
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetDisp(long lIndex, IDispatch** ppdisp)
{
    CacheItem* pce;

    Assert((lIndex>=0) && (lIndex<_aryItems.Size()));

    HRESULT hr = S_OK;

    *ppdisp = NULL;

    pce = &_aryItems[lIndex];

    // if not identity and there is a colllection, addref and return it
    if(!pce->fIdentity)
    {
        if(pce->pdisp)
        {
            pce->pdisp->AddRef();
        }
        else
        {
            IDispatch* pdisp;

            hr = CreateCollectionHelper(&pdisp, lIndex);
            if(hr)
            {
                goto Cleanup;
            }

            //We're not dependant on any other collection
            pce = &_aryItems[lIndex];
            pce->pdisp = pdisp;
            pce->sIndex = -1;
        }

        *ppdisp = pce->pdisp;
    }
    else
    {
        hr = DYNCAST(CElement, _pBase)->QueryInterface(IID_IDispatch, (void**)ppdisp);
    }

Cleanup:
    RRETURN(hr);
}

// Find pElement in the lIndex base Collection
HRESULT CCollectionCache::CreateChildrenCollection(
    long        lCollectionIndex, 
    CElement*   pElement, 
    IDispatch** ppDisp,
    BOOL        fAllChildren,
    BOOL        fDOMCollection)
{
    CacheItem*      pce;
    long            lSize = _aryItems.Size(), l;
    HRESULT         hr = S_OK;
    CollCacheType   Type = fAllChildren ? CacheType_AllChildren :
        (fDOMCollection)?CacheType_DOMChildNodes:CacheType_Children;

    Assert(ppDisp);
    *ppDisp = NULL;

    hr = EnsureAry(lCollectionIndex);
    if(hr)
    {
        goto Cleanup;
    }

    // Try and locate an exiting collection
    pce = &_aryItems[_lReservedSize];

    // Return this named collection if it already exists.
    for(l=_lReservedSize; l<lSize; ++l,++pce)
    {
        if(pce->Type==Type && pElement==pce->pElementBase)
        {
            pce->pdisp->AddRef();
            *ppDisp = pce->pdisp;
            goto Cleanup;
        }
    }

    // Didn't find it, create a new collection
    hr = GetFreeIndex(&l); // always returns Idx from non-reserved part of cache
    if(hr)
    {
        goto Cleanup;
    }

    hr = CreateCollectionHelper(ppDisp, l);
    if(hr)
    {
        goto Cleanup;
    }

    pce = &_aryItems[l];
    pce->Init();

    Assert(!pce->_pCacheItem);
    pce->_pCacheItem = new CElementAryCacheItem();
    if(!pce->_pCacheItem)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = BuildChildArray(lCollectionIndex, pElement, pce->_pCacheItem, fAllChildren);
    if(hr)
    {
        goto Cleanup;
    }

    pce->pElementBase = pElement;
    pce->pdisp = *ppDisp;

    pce->sIndex = lCollectionIndex; // Remember the index we depend on.
    pce->Type = Type;

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetDisp
//
//  Synopsis:   Get a dispatch ptr from the NON-RESERVED part of the cache.
//              N.B. non-reserved can never be identity collecitons,
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetDisp(
          long          lIndex,
          LPCTSTR       Name,
          CollCacheType CacheType,
          IDispatch**   ppdisp,
          BOOL          fCaseSensitive/*=FALSE*/,
          RECT*         pRect/*=NULL*/,
          BOOL          fAlwaysCollection/*=FALSE*/)
{
    long        lSize = _aryItems.Size();
    long        l;
    HRESULT     hr = S_OK;
    CacheItem*  pce;
    CRect       rectCellRange(CRect::CRECT_EMPTY);

    // named arrays are always built into an AryCacheItem
    CElementAryCacheItem aryItem;

    Assert(CacheType==CacheType_Tag || CacheType==CacheType_Named || 
        CacheType==CacheType_CellRange || CacheType==CacheType_Urn);

    typedef int (*COMPAREFN)(LPCTSTR, LPCTSTR);
    COMPAREFN CompareFn = fCaseSensitive ? FormsStringCmp : FormsStringICmp;

    *ppdisp = NULL;

    pce = &_aryItems[_lReservedSize];

    // Return this named collection if it already exists.
    for(l=_lReservedSize; l<lSize; ++l,++pce)
    {
        if(pce->Type==CacheType && lIndex==pce->sIndex && 
            pce->fIsCaseSensitive==(unsigned)fCaseSensitive &&
            !CompareFn(Name, (BSTR)pce->cstrName))
        {
            pce->pdisp->AddRef();
            *ppdisp = pce->pdisp;
            goto Cleanup;
        }
    }

    // Build the list
    if(CacheType != CacheType_CellRange)
    {
        hr=BuildNamedArray(lIndex, Name, CacheType==CacheType_Tag, 
            &aryItem, 0, fCaseSensitive, CacheType==CacheType_Urn);
    }
    else
    {           
        if(!pRect)
        {
            // Mark the rect as empty
            rectCellRange.right = -1;
        }
        else
        {
            // Use the passed in rect
            rectCellRange = *pRect;
        }
        hr = BuildCellRangeArray(lIndex, Name, &rectCellRange, &aryItem);
    }

    if(hr)
    {
        goto Cleanup;
    }

    // Return based on what the list of named elements looks like.
    if(!aryItem.Length() && !((CacheType==CacheType_Tag) || (CacheType==CacheType_Urn) || fAlwaysCollection))
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Error;
    }
    // The tags method ALWAYS should return a collection (else case)
    else if(aryItem.Length()==1 && !((CacheType==CacheType_Tag) || (CacheType==CacheType_Urn) || fAlwaysCollection))
    {
        CElement* pElem = aryItem.GetAt(0);
        hr = pElem->QueryInterface(IID_IDispatch, (void**)ppdisp);
        // Keep the ppdisp around we'll return that and just release the array.
        goto Cleanup2;
    }
    else
    {
        CElementAryCacheItem* pAryItem;

        hr = GetFreeIndex(&l); // always returns Idx from non-reserved part of cache
        if(hr)
        {
            goto Error;
        }

        hr = CreateCollectionHelper(ppdisp, l);
        if(hr)
        {
            goto Error;
        }

        pce = &_aryItems[l];
        pce->Init();

        hr = pce->cstrName.Set(Name);
        if(hr)
        {
            goto Error;
        }

        Assert(!pce->_pCacheItem);
        pce->_pCacheItem = new CElementAryCacheItem();
        if(!pce->_pCacheItem)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        // Copy the array. 
        // For perf reasons assume that the destination collection is a 
        // ary cache - which it is for now, by design
        pAryItem = DYNCAST(CElementAryCacheItem, pce->_pCacheItem);
        pAryItem->CopyAry(&aryItem); // this just copies the ptrarray _pv across

        pce->pdisp = *ppdisp;
        pce->sIndex = lIndex; // Remember the index we depend on.
        pce->Type = CacheType;
        pce->fIsCaseSensitive = fCaseSensitive;

        // Save the range for the cell range type cache so we do not need to parse
        // the name later
        if(CacheType == CacheType_CellRange)
        {
            pce->rectCellRange = rectCellRange;
        }

        // The collection this named collection was built from is now
        // used to rebuild (ensure) this collection. so we need to
        // put a reference on it so that it will not go away and its
        // location re-assigned by another call to GetFreeIndex.
        // The matching Release() will be done in the dtor
        // although it is not necessary to addref the reserved collections
        //  it is done anyhow, simply for consistency.  This addref
        // only needs to be done for non-reserved collections
        if(lIndex >= _lReservedSize)
        {
            _aryItems[lIndex].pdisp->AddRef();
        }
    }

Cleanup:
    RRETURN(hr);

Error:
    ClearInterface(ppdisp);

Cleanup2:
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Member:     GetDisp
//
//  Synopsis:   Get a dispatch ptr on an element from the cache.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetDisp(long lIndex, long lIndexElement, IDispatch** ppdisp)
{
    CElement* pElem;

    Assert((lIndex>=0) && (lIndex<_aryItems.Size()));

    *ppdisp = NULL;

    if(lIndexElement >= _aryItems[lIndex]._pCacheItem->Length())
    {
        RRETURN(DISP_E_MEMBERNOTFOUND);
    }

    if(lIndexElement < 0)
    {
        RRETURN(E_INVALIDARG);
    }

    pElem = _aryItems[lIndex]._pCacheItem->GetAt(lIndexElement);
    Assert(pElem);

    RRETURN(pElem->QueryInterface(IID_IDispatch, (void**)ppdisp));
}

//+------------------------------------------------------------------------
//
//  Member:     GetDisp
//
//  Synopsis:   Get a dispatch ptr on an element from the cache.
//      Return the nth element that mathces the name
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetDisp(
          long          lIndex,
          LPCTSTR       Name,
          long          lNthElement,
          IDispatch**   ppdisp,
          BOOL          fCaseSensitive)
{
    long                    lSize,l;
    HRESULT                 hr = DISP_E_MEMBERNOTFOUND;
    CCollectionCacheItem*   pItem;
    CElement*               pElem;

    Assert((lIndex>=0) && (lIndex<_aryItems.Size()));
    Assert(ppdisp);

    pItem = _aryItems[lIndex]._pCacheItem;

    *ppdisp = NULL;

    // if lIndexElement is too large, just pretend we
    //  didn't find it rather then erroring out
    if(lNthElement < 0)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    lSize = pItem->Length();

    if(lNthElement >= lSize)
    {
        goto Cleanup;
    }

    for(l=0; l<lSize; ++l)
    {
        pElem = pItem->GetAt(l);
        Assert(pElem);
        if(CompareName(pElem, Name, FALSE, fCaseSensitive) && !lNthElement--)
        {
            RRETURN(pElem->QueryInterface(IID_IDispatch, (void**)ppdisp));
        }
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetFreeIndex
//
//  Synopsis:   Get a free slot from the NON-RESERVED part of the cache.
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetFreeIndex(long* plIndex)
{
    CacheItem*  pce;
    long        lSize = _aryItems.Size();
    long        l;
    HRESULT     hr = S_OK;

    pce = &_aryItems[_lReservedSize];

    // Look for a free slot in the non-reserved part of the cache.
    for(l=_lReservedSize; l<lSize; ++l,++pce)
    {
        if(pce->Type == CacheType_FreeEntry)
        {
            Assert(!pce->pdisp);
            *plIndex = l;
            goto Cleanup;
        }
    }

    // If we failed to find a free slot then grow the cache by one.
    hr = _aryItems.EnsureSize(l+1);
    if(hr)
    {
        goto Cleanup;
    }
    _aryItems.SetSize(l+1);

    pce = &_aryItems[l];

    memset(pce, 0, sizeof(CacheItem));
    pce->fOKToDelete = TRUE;

    *plIndex = lSize;

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CompareName
//
//  Synopsis:   Compares a bstr with an element, either by name rules or
//              as a tagname.
//
//-------------------------------------------------------------------------
BOOL CCollectionCache::CompareName(
           CElement*    pElement,
           LPCTSTR      Name, 
           BOOL         fTagName, 
           BOOL         fCaseSensitive/*=FALSE*/)
{
    BOOL fCompare;
    typedef int (*COMPAREFN)(LPCTSTR, LPCTSTR);
    COMPAREFN CompareFn = fCaseSensitive ? FormsStringCmp : FormsStringICmp;

    if(fTagName)
    {
        fCompare = !CompareFn(Name, pElement->TagName());
    }
    else if(pElement->IsNamed())
    {
        BOOL fHasName;
        LPCTSTR pchId = pElement->GetAAname();

        // do we have a name
        fHasName = !!(pchId);

        fCompare = pchId ? !CompareFn(Name, pchId) : FALSE;
        if(!fCompare)
        {
            pchId = pElement->GetAAid();
            fCompare = pchId ? !CompareFn(Name, pchId) : FALSE;
        }

        if(!fCompare)
        {
            pchId = pElement->GetAAuniqueName();
            fCompare = pchId ? !CompareFn(Name, pchId) : FALSE;
        }
    }
    else
    {
        fCompare = FALSE;
    }

    return fCompare;
}

HRESULT CCollectionCache::BuildChildArray(
          long      lCollectionIndex, // Index of collection on which this child collection is based
          CElement* pRootElement,
          CCollectionCacheItem* pIntoCacheItem,
          BOOL      fAll)
{
    long                    lSize;
    HRESULT                 hr = S_OK;
    long                    lSourceIndex;
    CCollectionCacheItem*   pFromCacheItem;
    CElement*               pElem;
    CTreeNode*              pNode;

    Assert(lCollectionIndex>=0 && lCollectionIndex<_aryItems.Size());
    Assert(pIntoCacheItem);
    Assert(pRootElement);

    pFromCacheItem = _aryItems[lCollectionIndex]._pCacheItem;

    // BUGBUG rgardner - about this fn & the "assert(lCollectionIndex==0)"
    // As a result this fn is not very generic. This assert also assumes that
    // we're item 0 in the all collection. However, this is currently the only situation
    // this fn is called, and making the assumption optimizes the code
    Assert(lCollectionIndex == 0);

    pIntoCacheItem->ResetContents();

    // Didn't find it, create a new collection
    lSourceIndex = pRootElement->GetSourceIndex();
    // If we are outside the tree return
    if(lSourceIndex < 0)
    {
        goto Cleanup;
    }

    lSize = pFromCacheItem->Length();

    if(lSourceIndex >= lSize)
    {
        // This should never happen
        // No match - Return error 
        Assert(0);
        hr = E_UNEXPECTED;
        goto Cleanup;
    }

    // Now locate all the immediate children of the element and add them to the array
    for(;;)
    {
        pElem = pFromCacheItem->GetAt(++lSourceIndex) ;
        if(!pElem)
        {
            break;
        }
        pNode = pElem->GetFirstBranch();
        Assert(pNode);
        // optimize search to spot when we go outside scope of element
        if(!pNode->SearchBranchToRootForScope(pRootElement))
        {
            // outside scope of element
            break;
        }
        // If the fall flag is on it means all direct descendants
        // Otherwise it means only immediate children
        if(fAll || (pNode->Parent() && pNode->Parent()->Element()==pRootElement))
        {
            pIntoCacheItem->AppendElement(pElem);
        }
    }
Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     BuildNamedArray
//
//  Synopsis:   Fills in an array based on names found in the array of the
//              given index.
//              If we are building a named array on the ELEMENT_COLLECTION
//              we check if we have already created a collection of all
//              named elements, if not, we build it, if yes, we use this
//              collection to access all named elements
//
//  Result:     S_OK
//              E_OUTOFMEMORY
//
//  Note: It is now the semantics of this function to always return a
//        named array (albeit with a size of 0) instead of returning
//        DISPID_MEMBERNOTFOUND.  This allows for Tags to return an empty
//        collection.
//-------------------------------------------------------------------------
HRESULT CCollectionCache::BuildNamedArray(
          long      lCollectionIndex,
          LPCTSTR   Name,
          BOOL      fTagName,
          CCollectionCacheItem* pIntoCacheItem,
          long      iStartFrom,
          BOOL      fCaseSensitive,
          BOOL      fUrn /*=FALSE*/)
{
    HRESULT                 hr = S_OK;
    CElement*               pElem;
    CCollectionCacheItem*   pFromCacheItem;
    BOOL                    fAddElement;

    Assert(lCollectionIndex>=0 && lCollectionIndex<_aryItems.Size());
    Assert(pIntoCacheItem);

    pFromCacheItem = _aryItems[lCollectionIndex]._pCacheItem;
    pIntoCacheItem->ResetContents(); 

    for(pFromCacheItem->MoveTo(iStartFrom); ;)
    {
        pElem = pFromCacheItem->GetNext();
        if(!pElem)
        {
            break;
        }

        if(fUrn)
        {
            // 'Name' is the Urn we are looking for.  Check if this element has the requested Urn
            /*fAddElement = pElem->HasPeerWithUrn(Name); wlw note*/
        }
        else if(CompareName(pElem, Name, fTagName, fCaseSensitive))
        {
            fAddElement = TRUE;
        }
        else
        {
            fAddElement = FALSE;
        }

        if(fAddElement)
        {
            hr = pIntoCacheItem->AppendElement(pElem);
            if(hr)
            {
                goto Cleanup;
            }
        }
    } 

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     BuildCellRangeArray
//
//  Synopsis:   Fills in an array based on range of cells found in the array of the
//              given index.
//              We first check if we have already created a collection of all
//              cells, if not, we build it, if yes, we use this collection to access 
//              all cells
//
//  Result:     S_OK
//              E_OUTOFMEMORY
//
//  Note: It is now the semantics of this function to always return a
//        named array (albeit with a size of 0) instead of returning
//        DISPID_MEMBERNOTFOUND.  This allows us to return an empty
//        collection.
//-------------------------------------------------------------------------
HRESULT CCollectionCache::BuildCellRangeArray(
    long    lCollectionIndex,
    LPCTSTR szRange,
    RECT*   pRect,
    CCollectionCacheItem* pIntoCacheItem)
{
    CCollectionCacheItem*   pFromCacheItem;
    HRESULT                 hr = S_OK;
//    CTableCell*             pCell;
//    CElement*               pElem;
//
//    Assert(pRect);
//    Assert(lCollectionIndex>=0 && lCollectionIndex<_aryItems.Size());
//    Assert(pIntoCacheItem);
//
//    pFromCacheItem = _aryItems[lCollectionIndex]._pCacheItem;
//
//    if(pRect->right == -1)
//    {
//        // The rect is empty, parse it from the string
//        CCellRangeParser cellRangeParser(szRange);
//        if(cellRangeParser.Failed())
//        {
//            hr = E_INVALIDARG;
//            goto Cleanup;
//        }
//        cellRangeParser.GetRangeRect(pRect);
//    }
//
//    pIntoCacheItem->ResetContents();
//
//    // Build a list of cells in the specified range
//    for(pFromCacheItem->MoveTo(0); ;)
//    {
//        pElem = pFromCacheItem->GetNext();
//        if(!pElem)
//        {
//            break;
//        }
//        if(pElem->Tag()!=ETAG_TD && pElem->Tag()!=ETAG_TH)
//        {
//            break;
//        }
//        pCell = DYNCAST(CTableCell, pElem);
//        if(pCell->IsInRange(pRect))
//        {
//            hr = pIntoCacheItem->AppendElement(pElem);
//            if(hr)
//            {
//                goto Cleanup;
//            }
//        }
//    }
//Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetLength
//
//  Synopsis:   collection object model helper
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetLength(long lCollection, long* plSize)
{
    HRESULT hr;

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    // Get its current size.
    *plSize = SizeAry(lCollection);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     Item
//
//  Synopsis:   collection object model
//
//              we handle the following parameter cases:
//                  0 params            : by index = 0
//                  1 params bstr       : by name, index = 0
//                  1 params #          : by index
//                  2 params bstr, #    : by name, index
//                  2 params #, bstr    : by index, ignoring bstr
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::Item(long lCollection, VARIANTARG var1, VARIANTARG var2, IDispatch** ppResult)
{
    VARIANT*    pvarName = NULL;
    VARIANT*    pvarOne  = NULL;
    VARIANT*    pvarIndex = NULL;
    long        lIndex = 0;
    HRESULT     hr=E_INVALIDARG;

    if(!ppResult)
    {
        goto Cleanup;
    }

    *ppResult = NULL;

    pvarOne = (V_VT(&var1) == (VT_BYREF|VT_VARIANT)) ? V_VARIANTREF(&var1) : &var1;

    if((V_VT(pvarOne)==VT_BSTR) || V_VT(pvarOne)==(VT_BYREF|VT_BSTR))
    {
        pvarName = (V_VT(pvarOne)&VT_BYREF) ? V_VARIANTREF(pvarOne) : pvarOne;

        if((V_VT(&var2)!=VT_ERROR) && (V_VT(&var2)!=VT_EMPTY))
        {
            pvarIndex = &var2;
        }
    }
    else if((V_VT(&var1)!=VT_ERROR )&& (V_VT(&var1)!=VT_EMPTY))
    {
        pvarIndex = &var1;
    }

    if(pvarIndex)
    {
        VARIANT varNum;

        VariantInit(&varNum);
        hr = VariantChangeTypeSpecial(&varNum, pvarIndex, VT_I4);
        if(hr)
        {
            goto Cleanup;
        }

        lIndex = V_I4(&varNum);
    }

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    // Get a collection or element of the specified object.
    if(pvarName)
    {
        BSTR Name = V_BSTR(pvarName);

        if(pvarIndex)
        {
            hr = GetDisp(
                lCollection,
                Name,
                lIndex,
                ppResult,
                FALSE); // BUBUG rgardner - shouldn't ignore case
            if(hr)
            {
                goto Cleanup;
            }
        }
        else
        {
            hr = GetDisp(
                lCollection,
                Name,
                CacheType_Named,
                ppResult,
                FALSE); // BUBUG rgardner - shouldn't ignore case
            if(FAILED(hr))
            {
                HRESULT hrSave = hr; // save error code, and see if it a cell range
                hr = GetDisp(
                    lCollection,
                    Name,
                    CacheType_CellRange,
                    ppResult,
                    FALSE); // BUBUG rgardner - shouldn't ignore case
                if(hr)
                {
                    hr = hrSave; // restore error code
                    goto Cleanup;
                }
            }
        }
    }
    else
    {
        hr = GetDisp(lCollection, lIndex, ppResult);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    // If we didn't find anything, make sure to just return NULL.
    if(hr == DISP_E_MEMBERNOTFOUND)
    {
        hr = S_OK;
    }

    RRETURN(hr);
}

HRESULT CCollectionCache::GetElemAt(long lCollection, long* plCurrIndex, IDispatch** ppCurrNode)
{
    HRESULT hr;

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    *ppCurrNode = NULL;

    if(*plCurrIndex < 0)
    {
        *plCurrIndex = 0;
        goto Cleanup;
    }
    else if(*plCurrIndex > SizeAry(lCollection)-1)
    {
        *plCurrIndex = SizeAry(lCollection) - 1;
        goto Cleanup;
    }

    hr = GetDisp(lCollection, *plCurrIndex, ppCurrNode);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     Tags
//
//  Synopsis:   collection object model, this always returns a collection
//              and is named based on the tag, and searched based on tagname
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::Tags(long lCollection, VARIANT var1, IDispatch** ppdisp)
{
    VARIANT* pvarName = NULL;
    HRESULT hr = E_INVALIDARG;

    if(!ppdisp)
    {
        goto Cleanup;
    }

    *ppdisp = NULL;

    pvarName = (V_VT(&var1)==(VT_BYREF|VT_VARIANT)) ? V_VARIANTREF(&var1) : &var1;

    if((V_VT(pvarName)==VT_BSTR) || V_VT(pvarName)==(VT_BYREF|VT_BSTR))
    {
        pvarName = (V_VT(pvarName)&VT_BYREF) ? V_VARIANTREF(pvarName) : pvarName;
    }
    else
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    hr = Tags(lCollection, V_BSTR(pvarName), ppdisp);

Cleanup:
    RRETURN(hr);
}

HRESULT CCollectionCache::Tags(long lCollection, LPCTSTR szTagName, IDispatch** ppdisp)
{
    HRESULT hr;

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    // Get a collection of the specified tags.
    hr = GetDisp(
        lCollection,
        szTagName,
        CacheType_Tag,
        (IDispatch**)ppdisp,
        FALSE); // Case sensitivity ignored for TagName
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     Urns
//
//  Synopsis:   collection object model, this always returns a collection
//              and is named based on the urn, and searched based on urn
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::Urns(long lCollection, VARIANT var1, IDispatch** ppdisp)
{
    VARIANT* pvarName = NULL;
    HRESULT hr = E_INVALIDARG;

    if(!ppdisp)
    {
        goto Cleanup;
    }

    *ppdisp = NULL;

    pvarName = (V_VT(&var1)==(VT_BYREF|VT_VARIANT)) ? V_VARIANTREF(&var1) : &var1;

    if((V_VT(pvarName)==VT_BSTR) || V_VT(pvarName)==(VT_BYREF|VT_BSTR))
    {
        pvarName = (V_VT(pvarName)&VT_BYREF) ? V_VARIANTREF(pvarName) : pvarName;
    }
    else
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    // Get a collection of the elements with the specified urn
    hr = GetDisp(
        lCollection,
        V_BSTR(pvarName),
        CacheType_Urn,
        (IDispatch**)ppdisp,
        FALSE); // Case sensitivity ignored for Urn
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     GetNewEnum
//
//  Synopsis:   collection object model
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::GetNewEnum(long lCollection, IUnknown** ppEnum)
{
    CPtrAry<LPUNKNOWN>* pary = NULL;
    long                lSize;
    long                l;
    HRESULT             hr = E_INVALIDARG;

    if(!ppEnum)
    {
        goto Cleanup;
    }

    *ppEnum = NULL;

    // Make sure our collection is up-to-date.
    hr = EnsureAry(lCollection);
    if(hr)
    {
        goto Cleanup;
    }

    pary = new CPtrAry<LPUNKNOWN>;
    if(!pary)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    lSize = SizeAry(lCollection);

    hr = pary->EnsureSize(lSize);
    if(hr)
    {
        goto Error;
    }

    // Now make a snapshot of our collection.
    for(l=0; l<lSize; ++l)
    {
        IDispatch* pdisp;

        hr = GetDisp(lCollection, l, &pdisp);
        if(hr)
        {
            goto Error;
        }

        Verify(!pary->Append(pdisp));
    }

    // Turn the snapshot into an enumerator.
    hr = pary->EnumVARIANT(VT_DISPATCH, (IEnumVARIANT**) ppEnum, FALSE, TRUE);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    RRETURN(hr);

Error:
    pary->ReleaseAll();
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Member:     Remove
//
//  Synopsis:   remove the item in the collection at the given index
//
//-------------------------------------------------------------------------
HRESULT CCollectionCache::Remove(long lCollection, long lItemIndex)
{
    HRESULT hr;

    if((lItemIndex<0) || (lItemIndex>=SizeAry(lCollection)))
    {
        hr = E_INVALIDARG;
    }
    else
    {
        if(_pfnRemoveObject)
        {
            hr = CALL_METHOD((CVoid*)(void*)_pBase, _pfnRemoveObject, (lCollection, lItemIndex));
        }
        else
        {
            hr = CTL_E_METHODNOTAPPLICABLE;
        }
    }

    RRETURN(hr);
}

// returns the correct offset of given dispid in given collection and the 
// case sensetivity flag (if requested)
long CCollectionCache::GetNamedMemberOffset(long lCollectionIndex, DISPID dispid, BOOL* pfCaseSensitive/*=NULL*/)
{
    LONG lOffset;
    BOOL fSensitive;

    Assert(IsNamedCollectionMember(lCollectionIndex, dispid));

    // Check to see wich half of the dispid space the value goes
    if(IsSensitiveNamedCollectionMember(lCollectionIndex, dispid))
    {
        lOffset = GetSensitiveNamedMemberMin(lCollectionIndex);
        fSensitive = TRUE;
    }
    else
    {
        lOffset = GetNotSensitiveNamedMemberMin(lCollectionIndex);
        fSensitive = FALSE;
    }

    // return the sensitivity flag if required
    if(pfCaseSensitive != NULL)
    {
        *pfCaseSensitive = fSensitive;
    }

    return lOffset;
}

CAtomTable* CCollectionCache::GetAtomTable(BOOL* pfExpando)
{   
    Assert(_pDoc);

    if(pfExpando) 
    {
        *pfExpando = _pDoc->_fExpando; 
    }

    return _pDoc->GetAtomTable(); 
}