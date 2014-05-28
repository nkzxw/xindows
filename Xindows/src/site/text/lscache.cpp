
#include "stdafx.h"
#include "lscache.h"

#include "LineServices.h"

//+----------------------------------------------------------------------------
//
//  Function:   InitLSCache
//
//  Synopsis:   Allocate the LS context cache
//
//  Arguments:  pts - THREADSTATE for current thread
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//-----------------------------------------------------------------------------
HRESULT InitLSCache(THREADSTATE* pts)
{
    Assert(pts);
    pts->_pLSCache = new CLSCache;
    if(!pts->_pLSCache)
    {
        RRETURN(E_OUTOFMEMORY);
    }
    RRETURN(S_OK);
}

//+----------------------------------------------------------------------------
//
//  Function:   DeinitLSCache
//
//  Synopsis:   Delete the LS context cache
//
//  Arguments:  pts - THREADSTATE for current thread
//
//-----------------------------------------------------------------------------
void DeinitLSCache(THREADSTATE* pts)
{
    Assert(pts);

    delete pts->_pLSCache;
}

//-------------------------------------------------------------------------
//
// Member:   GetFreeEntry
//
// Synopsis: Finds a free entry inside LS cache. If one does not exist, creates
//           a new and adds it to the LS cache.
//           called internally by the constructors.
//
// Params:   None.
//
// Retruns:  A CLineServices object.
//
//-------------------------------------------------------------------------
CLineServices* CLSCache::GetFreeEntry(CMarkup* pMarkup, BOOL fStartUpLSDLL)
{
    CLSCacheEntry* pLSEntry;
    CLSCacheEntry  LSEntry;
    CLineServices* pLS = NULL;
    HRESULT        hr = S_OK;
    LONG           i;

    for(i=_aLSEntries.Size()-1; i>=0; i--)
    {
        pLSEntry = &_aLSEntries[i];
        if(!pLSEntry->_fUsed)
        {
            pLS = pLSEntry->_pLS;

            // If the LS context has not been intialized and initialization
            // has been requested, then start it up.
            if(fStartUpLSDLL)
            {
                hr = StartUpLSDLL(pLS, pMarkup);
                if(hr)
                {
                    pLS = NULL;
                    goto Error;
                }
                Assert(pLS->_plsc);
            }

            pLSEntry->_fUsed = TRUE;
            goto Cleanup;
        }
    }

    hr = InitLineServices(pMarkup, fStartUpLSDLL, &LSEntry._pLS);
    if(hr)
    {
        goto Cleanup;
    }

    LSEntry._fUsed = TRUE;

    hr = _aLSEntries.AppendIndirect(&LSEntry);
    if(hr)
    {
        goto Cleanup;
    }

    pLS = LSEntry._pLS;

Cleanup:
    if(pLS)
    {
        Assert(!fStartUpLSDLL || pLS->_plsc);
        _cUsed++;
    }

Error:    
    return pLS;
}

//-------------------------------------------------------------------------
//
// Member:   ReleaseEntry
//
// Synopsis: Releases an entry which was retrieved via GetFreeEntry(). It
//           puts this CLineServices object back into the cache for it to
//           be used by subsequent GetFreeEntry() calls.
//
// Params:   [pLS]: The entry to be released.
//
// Retruns:  None.
//
//-------------------------------------------------------------------------
void CLSCache::ReleaseEntry(CLineServices* pLS)
{
    CLSCacheEntry* pLSEntry;
    LONG i;

    Assert(pLS);
    for(i=0; i<_aLSEntries.Size(); i++)
    {
        pLSEntry = &_aLSEntries[i];
        if(pLSEntry->_pLS == pLS)
        {
            Assert(pLSEntry->_fUsed == TRUE);
            pLSEntry->_fUsed = FALSE;
            goto Cleanup;
        }
    }

    AssertSz(0, "Invalid cache entry passed to release");

Cleanup:
    _cUsed--;
    if(!_cUsed)
    {
        Dispose(FALSE);
    }
    return;
}

//-------------------------------------------------------------------------
//
// Member:   Dispose
//
// Synopsis: Appropriately destroys the complete cache of entries.
//
// Params:   fDiposeAll - Dispose all the entries.
//
// Retruns:  None.
//
//-------------------------------------------------------------------------
void CLSCache::Dispose(BOOL fDisposeAll)
{
    CLSCacheEntry* pLSEntry;
    LONG i;

    // Find the number of entries to free from the cache.
    LONG nUnusedEntriesNeedToFree = _aLSEntries.Size() - (fDisposeAll?0:N_CACHE_MINSIZE);

    // Do we need to free any entries?
    if(nUnusedEntriesNeedToFree > 0)
    {
        // Run thru the array finding free entries and disposing them till we
        // have disposed as many unused entries as we needed to.
        for(i=_aLSEntries.Size()-1; i>=0; i--)
        {
            Assert(nUnusedEntriesNeedToFree > 0);
            pLSEntry = &_aLSEntries[i];
            if(!pLSEntry->_fUsed)
            {
                Assert(pLSEntry->_pLS);
                DeinitLineServices(pLSEntry->_pLS);
                _aLSEntries.Delete(i);
                if(--nUnusedEntriesNeedToFree == 0)
                {
                    break;
                }
            }
        }
    }
}