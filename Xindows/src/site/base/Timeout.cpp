
#include "stdafx.h"
#include "Timeout.h"

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::GetFirstTimeoutEvent
//
//  Synopsis:   Returns the first timeout event if given event is found in the
//                 list. Im most cases the returned event will be the one
//                 with given id, but if WM_TIMER messages came out of order
//                 it can be another one with smaller target time.
//              Removes the event from the list before returning.
//
//              Return value is S_FALSE if given event is not in the list
//--------------------------------------------------------------------------
HRESULT CTimeoutEventList::GetFirstTimeoutEvent(UINT uTimerID, TIMEOUTEVENTINFO** ppTimeout)
{
    HRESULT hr = S_OK;
    int     nNumEvents = _aryTimeouts.Size();
    int     i;

    Assert(ppTimeout != NULL);

    // Find the event first
    for(i=nNumEvents-1; i>=0; i--)
    {
        if(_aryTimeouts[i]->_uTimerID == uTimerID)
        {
            break;
        }
    }

    if(i < 0)
    {
        // The event is no longer active, or there is an error
        *ppTimeout = NULL;
        hr = S_FALSE;
        goto Cleanup;
    }

    // Elements are sorted and given event is in the list.
    // As long as given element is in the list we can return the
    //      last element without further checks
    *ppTimeout = _aryTimeouts[nNumEvents-1];

    // Win16: Use GetTimeout(pTimeout->_uTimerID, dummy pTimeout) to delete this.
    // Remove it from the array
    _aryTimeouts.Delete(nNumEvents-1);

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::GetTimeout
//
//  Synopsis:   Gets timeout event with given timer id and removes it from the list
//
//              Return value is S_FALSE if given event is not in the list
//--------------------------------------------------------------------------
HRESULT CTimeoutEventList::GetTimeout(UINT uTimerID, TIMEOUTEVENTINFO** ppTimeout)
{
    int     i;
    HRESULT hr;

    for(i=_aryTimeouts.Size()-1; i>=0; i--)
    {
        if(_aryTimeouts[i]->_uTimerID == uTimerID)
        {
            break;
        }
    }

    if(i >= 0)
    {
        *ppTimeout = _aryTimeouts[i];
        // Remove the pointer
        _aryTimeouts.Delete(i);
        hr = S_OK;
    }
    else
    {
        *ppTimeout = NULL;
        hr = S_FALSE;
    }

    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::InsertIntoTimeoutList
//
//  Synopsis:   Saves given timeout info pointer in the list
//
//              Returns the ID associated with timeout entry
//--------------------------------------------------------------------------
HRESULT CTimeoutEventList::InsertIntoTimeoutList(TIMEOUTEVENTINFO* pTimeoutToInsert, UINT* puTimerID, BOOL fNewID)
{
    HRESULT hr = S_OK;
    int     i;
    int     nNumEvents = _aryTimeouts.Size();

    Assert(pTimeoutToInsert != NULL);

    // Fill the timer ID field with the next unused timer ID
    // We add this to make its appearance random
    if(fNewID)
    {
        pTimeoutToInsert->_uTimerID = _uNextTimerID++ + (DWORD)(DWORD_PTR)this;
    }

    // Find the appropriate position. Current implementation keeps the elements
    // sorted by target time, with the one having min target time near the top
    for(i=0; i<nNumEvents; i++)
    {
        if(pTimeoutToInsert->_dwTargetTime >= _aryTimeouts[i]->_dwTargetTime)
        {
            // Insert before current element
            hr = _aryTimeouts.Insert(i, pTimeoutToInsert);
            if(hr)
            {
                goto Cleanup;
            }
            break;
        }
    }

    if(i == nNumEvents)
    {
        /// Append at the end
        hr = _aryTimeouts.Append(pTimeoutToInsert);
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(puTimerID)
    {
        *puTimerID = pTimeoutToInsert->_uTimerID;
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::KillAllTimers
//
//  Synopsis:   Stops all the timers in the list and removes events from the list
//
//--------------------------------------------------------------------------
void CTimeoutEventList::KillAllTimers(void* pvObject)
{
    int i;

    for(i=_aryTimeouts.Size()-1; i>=0; i--)
    {
        Verify(FormsKillTimer(pvObject, _aryTimeouts[i]->_uTimerID) == S_OK);
        delete _aryTimeouts[i];
    }
    _aryTimeouts.DeleteAll();

    for(i=_aryPendingTimeouts.Size()-1; i>=0; i--)
    {
        delete _aryTimeouts[i];
    }
    _aryPendingTimeouts.DeleteAll();
    _aryPendingClears.DeleteAll();
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::GetPendingTimeout
//
//  Synopsis:   Gets a pending Timeout, and removes it from the list
//
//--------------------------------------------------------------------------
BOOL CTimeoutEventList::GetPendingTimeout(TIMEOUTEVENTINFO** ppTimeout)
{
    int i;
    Assert(ppTimeout);
    if((i=_aryPendingTimeouts.Size()-1) < 0)
    {
        *ppTimeout = NULL;
        return FALSE;
    }

    *ppTimeout = _aryPendingTimeouts[i];
    _aryPendingTimeouts.Delete(i);
    return TRUE;
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::ClearPendingTimeout
//
//  Synopsis:   Removes a timer from the pending list and returns TRUE.
//              If timer with ID not found, returns FALSE
//
//--------------------------------------------------------------------------
BOOL CTimeoutEventList::ClearPendingTimeout(UINT uTimerID)
{
    BOOL fFound = FALSE;

    for(int i=_aryPendingTimeouts.Size()-1; i>=0; i--)
    {
        if(_aryPendingTimeouts[i]->_uTimerID == uTimerID)
        {
            delete _aryPendingTimeouts[i];
            _aryPendingTimeouts.Delete(i);
            fFound = TRUE;
            break;
        }
    }
    return fFound;
}

//+-------------------------------------------------------------------------
//
//  Method:     CTimeoutEventList::GetPendingClear
//
//  Synopsis:   Returns TRUE and an ID of a timer that was cleared during
//              timer processing. If there are none left, it returns FALSE.
//
//--------------------------------------------------------------------------
BOOL CTimeoutEventList::GetPendingClear(LONG* plTimerID)
{
    int i;
    if((i=_aryPendingClears.Size()-1) < 0)
    {
        *plTimerID = 0;
        return FALSE;
    }

    *plTimerID = (LONG)_aryPendingClears[i];
    _aryPendingClears.Delete(i);
    return TRUE;
}