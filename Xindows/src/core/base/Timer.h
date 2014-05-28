
#ifndef __XINDOWS_CORE_BASE_TIMER_H__
#define __XINDOWS_CORE_BASE_TIMER_H__

struct TIMERENTRY
{
    void*           pvObject;
    PFN_VOID_ONTICK pfnOnTick;
    UINT            idTimer;
    UINT            idTimerAlias;
};

DECLARE_CDataAry(CAryTimers, TIMERENTRY);

HRESULT FormsSetTimer(void* pvObject, PFN_VOID_ONTICK pfnOnTick, UINT idTimer, UINT uTimeout);
HRESULT FormsKillTimer(void* pvObject, UINT idTimer);

#endif //__XINDOWS_CORE_BASE_TIMER_H__