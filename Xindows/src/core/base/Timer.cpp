
#include "stdafx.h"
#include "Timer.h"

//+-------------------------------------------------------------------------
//
//  Method:     ResetTimer
//
//  Synopsis:   Resets the timer identified by dwCookie.  ResetTimer
//              results in the timer setting being changed for the
//              given timer without allocating a new, unique timer id.
//
//--------------------------------------------------------------------------
static HRESULT ResetTimer(void* pvObject, UINT idTimer, UINT uTimeout)
{
    TIMERENTRY*     pte;
    int             c;
    THREADSTATE*    pts = GetThreadState();

    // Windows NT rounds the time up to 10.  If time is less 
    // than 10, NT spews to the debugger.  Work around
    // this problem by rounding up to 10.
    if(uTimeout < 10)
    {
        uTimeout = 10;
    }

    for(c=pts->gwnd.paryTimers->Size(),pte=*pts->gwnd.paryTimers; c>0; c--,pte++)
    {
        if((pte->pvObject==pvObject) && (pte->idTimer==idTimer))
        {
            if(SetTimer(pts->gwnd.hwndGlobalWindow, pte->idTimerAlias, uTimeout, NULL) == 0)
            {
                RRETURN(E_FAIL);
            }

            return S_OK;
        }
    }

    return S_FALSE;
}

//+-------------------------------------------------------------------------
//
//  Method:     GetUniqueID
//
//  Synopsis:   Fetches a unique timer ID by checking our sorted array
//              of used IDs for a new value.
//
//--------------------------------------------------------------------------
static UINT GetUniqueID()
{
    THREADSTATE*    pts;
    int             c;
    TIMERENTRY*     pte;
    BOOL            fDone = FALSE;

    pts = GetThreadState();
    while(fDone == FALSE)
    {
        pts->gwnd.uID++;

        // Note: We don't use the range 0x0000 through 0x1FFF.  This is
        // reserved for hard-coded timer identifiers.  Compuserve incorrectly
        // intercepts timer id 0x000F, which is yet another reason for doing
        // this.
        if(pts->gwnd.uID < 0x2000)
        {
            pts->gwnd.uID = 0x2000;
        }

        fDone = TRUE;

        for(c=(*(pts->gwnd.paryTimers)).Size(),pte = *(pts->gwnd.paryTimers); c>0; c--,pte++)
        {
            if(pts->gwnd.uID == pte->idTimerAlias)
            {
                fDone = FALSE;
                break;
            }
        }
    }

    return pts->gwnd.uID;
}

//+-------------------------------------------------------------------------
//
//  Function:   FormsSetTimer
//
//  Synopsis:   Sets a timer using the forms global window.
//
//  Arguments:  [pGWS]      Pointer to a timer sink
//              [idTimer]   Caller-specified id that will be passed back
//                          on a timer event.
//              [uTimeout]  Elapsed time between timer events
//
//--------------------------------------------------------------------------
HRESULT FormsSetTimer(void* pvObject, PFN_VOID_ONTICK pfnOnTick, UINT idTimer, UINT uTimeout)
{
    THREADSTATE*    pts;
    UINT            idTimerAlias;
    HRESULT         hr;

    Assert(pvObject);

    pts = GetThreadState();

    if(pts->gwnd.hwndGlobalWindow == NULL)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    // Attempt to reset the timer.  If this fails, ie, no matching
    // timer exists, continue on and set the timer.
    hr = ResetTimer(pvObject, idTimer, uTimeout);
    if(hr == S_OK)
    {
        goto Cleanup;
    }

    hr = (*(pts->gwnd.paryTimers)).EnsureSize((*(pts->gwnd.paryTimers)).Size()+1);
    if(hr)
    {
        goto Cleanup;
    }

    idTimerAlias = GetUniqueID();

    if(uTimeout < 10)
    {
        uTimeout = 10;
    }

    if(SetTimer(pts->gwnd.hwndGlobalWindow, idTimerAlias, uTimeout, NULL) == 0)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    (*(pts->gwnd.paryTimers))[(*(pts->gwnd.paryTimers)).Size()].pvObject = pvObject;
    (*(pts->gwnd.paryTimers))[(*(pts->gwnd.paryTimers)).Size()].pfnOnTick = pfnOnTick;
    (*(pts->gwnd.paryTimers))[(*(pts->gwnd.paryTimers)).Size()].idTimerAlias = idTimerAlias;
    (*(pts->gwnd.paryTimers))[(*(pts->gwnd.paryTimers)).Size()].idTimer = idTimer;
    (*(pts->gwnd.paryTimers)).SetSize((*(pts->gwnd.paryTimers)).Size()+1);

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Function:   FormsKillTimer
//
//  Synopsis:   Kills a forms timer.
//
//  Arguments:  [dwCookie]  Cookie identifying the timer.  Obtained from
//                          FormsSetTimer.
//
//--------------------------------------------------------------------------
HRESULT FormsKillTimer(void* pvObject, UINT idTimer)
{
    THREADSTATE*    pts;
    TIMERENTRY*     pte;
    int             i, c;
    UINT            idTimerAlias;

    // Note: We do not use pts = GetThreadState() function because
    // this function is called from the DLL process detach
    // code after TlsGetValue() has ceased to function correctly.
    // This scenario is probably a result of a bug in Windows '95.
    pts = (THREADSTATE*)(TlsGetValue(_afxGlobalData._dwTls));
    if(!pts || pts->gwnd.paryTimers==NULL)
    {
        return S_FALSE;
    }

    for(c=(*(pts->gwnd.paryTimers)).Size(),i=0,pte=(*(pts->gwnd.paryTimers));
        c>0; c--,i++,pte++)
    {
        if((pte->pvObject==pvObject) && (pte->idTimer==idTimer))
        {
            idTimerAlias = pte->idTimerAlias;
            (*(pts->gwnd.paryTimers)).Delete(i);
            KillTimer(TLS(gwnd.hwndGlobalWindow), idTimerAlias);
            return S_OK;
        }
    }
    return S_FALSE;
}
