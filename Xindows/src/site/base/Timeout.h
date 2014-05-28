
#ifndef __XINDOWS_SITE_BASE_TIMEOUT_H__
#define __XINDOWS_SITE_BASE_TIMEOUT_H__

struct TIMEOUTEVENTINFO
{
    IDispatch*  _pCode;         // PCODE code to execute if set don't use _code.
    UINT        _uTimerID;      // timer id
    DWORD       _dwTargetTime;  // System time (in  milliseconds) when
                                // the timer is going to time out
    DWORD       _dwInterval;    // interval for repeating timer. set to
                                // zero if called by setTimeout
public:
    TIMEOUTEVENTINFO() : _pCode(NULL) {}
    ~TIMEOUTEVENTINFO() { ReleaseInterface(_pCode); }
};

class CTimeoutEventList
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CTimeoutEventList()
    {
        _uNextTimerID = 1;
    }

    // Gets the first event that has min target time. If uTimerID not found
    //      returns S_FALSE
    // Removes the event from the list before returning
    HRESULT GetFirstTimeoutEvent(UINT uTimerID, TIMEOUTEVENTINFO** pTimeout);

    // Returns the timer event with given timer ID, or returns S_FALSE
    //      if timer not found
    HRESULT GetTimeout(UINT uTimerID, TIMEOUTEVENTINFO** ppTimeout);

    // Inserts given timer event object into the list  and returns timer ID
    HRESULT InsertIntoTimeoutList(TIMEOUTEVENTINFO* pTimeoutToInsert, UINT* puTimerID=NULL, BOOL fNewID=TRUE);

    // Kill all the script times for given object
    void    KillAllTimers(void* pvObject);

    // Add an Interval timeout to a pending list to be requeued after script processing
    void    AddPendingTimeout(TIMEOUTEVENTINFO* pTimeout) { _aryPendingTimeouts.Append(pTimeout); }

    // Returns timer event on pending queue. Removes timer event from the list
    BOOL    GetPendingTimeout(TIMEOUTEVENTINFO** ppTimeout);

    // If an interval timeout is cleared while in script, remove it from the pending list
    BOOL    ClearPendingTimeout(UINT uTimerID);

    // a clear was called during timer script processing, clear after all processing done.
    void    AddPendingClear(LONG lTimerID) { _aryPendingClears.Append(lTimerID); }

    // Returns the timer ID of the timer to clear
    BOOL    GetPendingClear(LONG* plTimerID);

private:
    DECLARE_CPtrAry(CAryTimeouts, TIMEOUTEVENTINFO*)
    DECLARE_CPtrAry(CAryPendingTimeouts, TIMEOUTEVENTINFO*)
    DECLARE_CPtrAry(CAryPendingClears, LONG_PTR)

    CAryTimeouts        _aryTimeouts;
    CAryPendingTimeouts _aryPendingTimeouts;
    CAryPendingClears   _aryPendingClears;
    UINT                _uNextTimerID;
};

#endif //__XINDOWS_SITE_BASE_TIMEOUT_H__