
#include "stdafx.h"
#include "DataCache.h"

#define celGrow     8

#ifdef _DEBUG
#define CheckFreeChain()    CheckFreeChainFn()
#else
#define CheckFreeChain()
#endif

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::constructor
//
//  Synopsis:
//
//  Returns:
//
//-------------------------------------------------------------------------
CDataCacheBase::CDataCacheBase()
{
    _pael = NULL;
    _cel = 0;
    _ielFirstFree = -1;
#ifdef _DEBUG
    _celsInCache = 0;
    _cMaxEls = 0;
#endif
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::Add
//
//  Synopsis:   Add a new DATA to the cache.
//
//  Arguments:  pvData - DATA to add to the cache
//              piel  - return index of DATA in the cache
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//  Note:       This DO NOT addref the new data in the cache
//
//-------------------------------------------------------------------------
HRESULT CDataCacheBase::Add(const void* pvData, LONG* piel, BOOL fClone)
{
    HRESULT hr = S_OK;
    CDataCacheElem* pel;
    LONG iel;
    LONG ielRet = NULL;
    CDataCacheElem* pelRet;

    if(_ielFirstFree >= 0) // Return first element of free list
    {
        ielRet = _ielFirstFree;

        hr = InitData(Elem(ielRet), pvData, fClone);
        if(hr)
        {
            goto Cleanup;
        }

        pelRet = Elem(ielRet);
        _ielFirstFree = pelRet->_ielNextFree;
    }
    else // All lower positions taken: need to add another celGrow elements
    {
        pel = _pael;

        hr = MemRealloc((void**)&pel, (_cel+celGrow)*sizeof(CDataCacheElem));
        if(hr)
        {
            goto Cleanup;
        }

        MemSetName((pel, "CDataCacheBase data - %d elements", _cel+celGrow));

        _pael = pel;
        ielRet = _cel;
        pelRet = pel + ielRet;
        _cel += celGrow;

#ifdef _DEBUG
        pelRet->_pvData = NULL;
#endif

        hr = InitData(pelRet, pvData, fClone);
        if(hr)
        {
            // Put all added elements in free list
            iel = ielRet;
        }
        else
        {
            // Use first added element
            // Put next element and subsequent ones added in free list
            iel = ielRet + 1;
        }

        // Add non yet used elements to free list
        _ielFirstFree = iel;

        for(pel=Elem(iel); ++iel<_cel; pel++)
        {
            pel->_pvData = NULL;
            pel->_ielNextFree = iel;
        }

        // Mark last element in free list
        pel->_pvData = NULL;
        pel->_ielNextFree = -1;
    }

    if(!hr)
    {
        Assert(pelRet);
        Assert(pelRet->_pvData);
        pelRet->ielRef._wCrc = ComputeDataCrc(pelRet->_pvData);
        pelRet->ielRef._wPad = 0;
        pelRet->_cRef = 0;

        if(piel)
        {
            *piel = ielRet;
        }
    }

#ifdef _DEBUG
    _cMaxEls = max(_cMaxEls, ++_celsInCache);
#endif

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::Free(iel)
//
//  Synopsis:   Free up a DATA in the cache by moving it to the
//              free list
//
//  Arguments:  iel  - index of DATA to free in the cache
//
//  Returns:    none
//
//-------------------------------------------------------------------------
void CDataCacheBase::Free(LONG iel)
{
    CDataCacheElem* _pElem = Elem(iel);

    Assert(_pElem->_pvData);

    // Passivate data
    PassivateData(_pElem);

    // Add it to free list
    _pElem->_ielNextFree = _ielFirstFree;
    _ielFirstFree = iel;

    // Flag it's freed
    _pElem->_pvData = NULL;

#ifdef _DEBUG
    _celsInCache--;
    Assert(_celsInCache >=0);
#endif
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::Free()
//
//  Synopsis:   Free up entire cache and deallocate memory
//
//-------------------------------------------------------------------------
void CDataCacheBase::Free()
{
#ifdef _DEBUG
    LONG iel;
    static BOOL fAssertDone = FALSE;

    if(!fAssertDone)
    {
        for(iel=0; iel<_cel; iel++)
        {
            if(Elem(iel)->_pvData != NULL)
            {
                AssertSz(FALSE, "CDataCacheBase::Free() - one or more cells not Empty");
                // Don't put up more than one assert
                fAssertDone = TRUE;
                break;
            }
        }
    }
#endif

    MemFree(_pael);

    _pael = NULL;
    _cel = 0;
    _ielFirstFree = -1;
#ifdef _DEBUG
    _celsInCache = 0;
#endif
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::Find(iel)
//
//  Synopsis:   Find given DATA in the cache and returns its index
//
//  Arguments:  pvData  - data to lookup
//
//  Returns:    index of DATA in the cache, -1 if not found
//
//-------------------------------------------------------------------------
LONG CDataCacheBase::Find(const void* pvData) const
{
    LONG iel;
    WORD wCrc;

    CheckFreeChain();

    wCrc = ComputeDataCrc(pvData);

    for(iel=0; iel<_cel; iel++)
    {
        CDataCacheElem* pElem = Elem(iel);

        if(pElem->_pvData && pElem->ielRef._wCrc==wCrc)
        {
            if(CompareData(pvData, pElem->_pvData))
            {
                return iel;
            }
        }
    }

    return -1;
}


//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::CheckFreeChain(iel)
//
//  Synopsis:   Check validity of the free chain
//
//  Arguments:  iel  - index of DATA to free in the cache
//
//  Returns:    none
//
//-------------------------------------------------------------------------
#ifdef _DEBUG
void CDataCacheBase::CheckFreeChainFn() const
{
    LONG cel = 0;
    LONG iel = _ielFirstFree;

    while(iel > 0)
    {
        AssertSz(iel==-1||iel<=_cel, "CDataCacheBase::CheckFreeChain() - Elem points to out of range elem");

        iel = Elem(iel)->_ielNextFree;

        if(++cel > _cel)
        {
            AssertSz(FALSE, "CDataCacheBase::CheckFreeChain() - Free chain seems to contain an infinite loop");
            return;
        }
    }
}
#endif

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::ReleaseData(iel)
//
//  Synopsis:   Decrement the ref count of DATA of given index in the cache
//              Free it if ref count goes to 0
//
//  Arguments:  iel  - index of DATA to release in the cache
//
//  Returns:    none
//
//-------------------------------------------------------------------------
void CDataCacheBase::ReleaseData(LONG iel)
{
    CheckFreeChain();

    if(iel >= 0) // Ignore default iCF
    {
        Assert(Elem(iel)->_pvData);
        Assert(Elem(iel)->_cRef > 0);

        if(--(Elem(iel)->_cRef) == 0) // Entry no longer referenced
        {
            Free (iel); // Add it to the free chain
        }
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::AddRefData(iel)
//
//  Synopsis:   Addref DATA of given index in the cache
//              Free it if ref count goes to 0
//
//  Arguments:  iel  - index of DATA to addref in the cache
//
//  Returns:    none
//
//-------------------------------------------------------------------------
void CDataCacheBase::AddRefData(LONG iel)
{
    CheckFreeChain();

    if(iel >= 0)
    {
        Assert(Elem(iel)->_pvData);
        // BUGBUG: should allocate a new element when ref count goes
        // above maximum WORD value
        Assert(Elem(iel)->_cRef < MAXWORD);
        Elem(iel)->_cRef++;
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::CacheData(*pvData, *piel, *pfDelete, fClone)
//
//  Synopsis:   Cache new DATA. This looks up the DATA and if found returns
//              its index and addref it, otherwise adds it to the cache.
//
//  Arguments:  pvData - DATA to add to the cache
//              piel  - return index of DATA in the cache
//              pfDelete - returns whether or not to delete pvData on success
//              fClone - tells whether to clone or copy pointer
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//-------------------------------------------------------------------------
HRESULT CDataCacheBase::CacheData(const void* pvData, LONG* piel, BOOL* pfDelete, BOOL fClone)
{
    HRESULT hr = S_OK;

    LONG iel = Find(pvData);

    if(pfDelete)
    {
        *pfDelete = FALSE;
    }

    if(iel >= 0)
    {
        Assert(Elem(iel)->_pvData);
        if(pfDelete)
        {
            *pfDelete = TRUE;
        }
    }
    else
    {
        hr = Add(pvData, &iel, fClone);
        if(hr)
        {
            goto Cleanup;
        }
    }

    AddRefData(iel);

    CheckFreeChain();

    if(piel)
    {
        *piel = iel;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CDataCacheBase::CacheDataPointer(**ppvData, *piel)
//
//  Synopsis:   Caches new data via CacheData, but does not clone.
//              On success, it takes care of memory management of input data.
//
//  Arguments:  ppvData - DATA to add to the cache
//              piel  - return index of DATA in the cache
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//-------------------------------------------------------------------------
HRESULT CDataCacheBase::CacheDataPointer(void** ppvData, LONG* piel)
{
    BOOL fDelete;
    HRESULT hr;

    Assert(ppvData);
    hr = CacheData(*ppvData, piel, &fDelete, FALSE);
    if(!hr)
    {
        if(fDelete)
        {
            DeletePtr(*ppvData);
        }
        *ppvData = NULL;
    }
    RRETURN(hr);
}