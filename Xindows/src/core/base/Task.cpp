
#include "stdafx.h"
#include "Task.h"

#define TASKMAN_TIMERID 1000
#define TIMER_INFINITE  0x7FFFFFFF

//+---------------------------------------------------------------------------
//
//  Member:     InitTaskManager
//
//  Synopsis:   Initializes the lightweight task manager
//
//----------------------------------------------------------------------------
HRESULT InitTaskManager(THREADSTATE* pts)
{
    pts->task.dwTickRun   = 200;
    pts->task.dwTickTimer = TIMER_INFINITE;

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     DeinitTaskManager
//
//----------------------------------------------------------------------------
void DeinitTaskManager(THREADSTATE* pts)
{
    CTask* ptask;

    Assert(pts->task.ptaskHead==NULL && "Active tasks remain at shutdown");
    Assert(pts->task.ptaskCur == NULL);

    while((ptask=pts->task.ptaskHead) != NULL)
    {
        Assert(!ptask->TestFlag(TASKF_INRUN));
        ptask->Terminate();
        ptask->Release();
    }

    FormsKillTimer(InitTaskManager, TASKMAN_TIMERID);
}

//+---------------------------------------------------------------------------
//
//  Member:     TaskmanResetTimer
//
//  Synopsis:   Sets the timer to the minimum interval of all unblocked
//              tasks.
//
//----------------------------------------------------------------------------
void CTask::TaskmanResetTimer()
{
    THREADSTATE* pts = GetThreadState();
    CTask*  ptask;
    DWORD   dwMin;
    DWORD   dwCurTick = GetTickCount();
    DWORD   dwFromLastTick;

    if(pts->task.fSuspended)
    {
        dwMin = TIMER_INFINITE;
    }
    else if(pts->task.cInterval > 0)
    {
        dwMin = TIMER_INFINITE;

        for(ptask=pts->task.ptaskHead; ptask; ptask=ptask->_ptaskNext)
        {
            if(!ptask->TestFlag(TASKF_BLOCKED|TASKF_TERMINATED))
            {
                // Check for smallest interal
                if(dwMin > ptask->_dwTickInterval)
                {
                    dwMin = ptask->_dwTickInterval;
                }

                dwFromLastTick = dwCurTick - ptask->_dwTickLast;

                // Check if a task needs to run earlier then that
                if(dwMin > dwFromLastTick)
                {
                    dwMin = dwFromLastTick;
                }
            }
        }
    }
    else if(pts->task.cUnblocked > 0)
    {
        dwMin = 0;
    }
    else
    {
        dwMin = TIMER_INFINITE;
    }

    TaskmanSetTimer(dwMin);
}

//+---------------------------------------------------------------------------
//
//  Member:     TaskmanSetTimer
//
//  Synopsis:   Sets timer if new interval is different then existing one
//
//----------------------------------------------------------------------------
void CTask::TaskmanSetTimer(DWORD dwTimeOut)
{
    THREADSTATE* pts = GetThreadState();

    if(dwTimeOut != pts->task.dwTickTimer)
    {
        FormsSetTimer(InitTaskManager,
            ONTICK_METHOD(CTask, OnTaskTick, ontasktick), TASKMAN_TIMERID, dwTimeOut);

        pts->task.dwTickTimer = dwTimeOut;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     TaskmanEnqueue
//
//  Synopsis:   Adds a task to the list of tasks
//
//----------------------------------------------------------------------------
void CTask::TaskmanEnqueue(CTask* ptask)
{
    THREADSTATE* pts = GetThreadState();
    CTask **pptask, *ptaskT;

    Assert(!ptask->TestFlag(TASKF_ENQUEUED));
    Assert(!ptask->TestFlag(TASKF_TERMINATED));

    for(pptask=&pts->task.ptaskHead; (ptaskT=*pptask)!=NULL; pptask=&ptaskT->_ptaskNext) ;
    *pptask = ptask;
    ptask->_ptaskNext = NULL;
    ptask->SetFlag(TASKF_ENQUEUED);

    if(!ptask->TestFlag(TASKF_BLOCKED))
    {
        pts->task.cUnblocked += 1;
    }

    if(ptask->_dwTickInterval > 0)
    {
        pts->task.cInterval += 1;
    }

    TaskmanResetTimer();
}

//+---------------------------------------------------------------------------
//
//  Member:     CTaskManager::Dequeue
//
//  Synopsis:   Removes a task from the list of tasks
//
//----------------------------------------------------------------------------
void CTask::TaskmanDequeue(CTask* ptask)
{
    THREADSTATE* pts;
    CTask **pptask, *ptaskT;

    if(!ptask->TestFlag(TASKF_ENQUEUED))
    {
        return;
    }

    pts = GetThreadState();

    for(pptask=&pts->task.ptaskHead; (ptaskT=*pptask)!=NULL; pptask=&ptaskT->_ptaskNext)
    {
        if(ptaskT == ptask)
        {
            if(pts->task.ptaskNext == ptask)
            {
                pts->task.ptaskNext = ptask->_ptaskNext;
            }

            if(pts->task.ptaskNextInLoop == ptask)
            {
                pts->task.ptaskNextInLoop = ptask->_ptaskNext;
            }

            *pptask = ptask->_ptaskNext;

            ptask->ClearFlag(TASKF_ENQUEUED);

            if(!ptask->TestFlag(TASKF_BLOCKED))
            {
                pts->task.cUnblocked -= 1;
            }

            if(ptask->_dwTickInterval > 0)
            {
                pts->task.cInterval -= 1;
            }

            TaskmanResetTimer();

            return;
        }
    }

    AssertSz(FALSE, "Task not found on task manager queue");
}

//+---------------------------------------------------------------------------
//
//  Member:     TaskmanRunTask
//
//  Synopsis:   Calls the OnRun method of the given task
//
//----------------------------------------------------------------------------
void CTask::TaskmanRunTask(THREADSTATE* pts, DWORD dwTick, CTask* ptask)
{
    // Don't allow more than one task to run at the same time
    if(pts->task.ptaskCur)
    {
        return;
    }

    ptask->SetFlag(TASKF_INRUN);
    pts->task.ptaskCur = ptask;
    ptask->_dwTickLast = dwTick;
    ptask->OnRun(dwTick+pts->task.dwTickRun);

    if(ptask->TestFlag(TASKF_DOTERM))
    {
        ptask->OnTerminate();
        TaskmanDequeue(ptask);
    }

    pts->task.ptaskCur = NULL;
    ptask->ClearFlag(TASKF_INRUN);

    if(pts->task.fSuspended)
    {
        pts->task.fSuspended = FALSE;
        TaskmanResetTimer();
    }

    if(ptask->TestFlag(TASKF_DODESTRUCT))
    {
        delete ptask;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     TaskmanRun
//
//  Synopsis:   Finds the next available task and calls it's OnRun method
//
//----------------------------------------------------------------------------
void CTask::TaskmanRun()
{
    THREADSTATE* pts = GetThreadState();
    CTask*  ptask;
    DWORD   dwTick;
    DWORD   dwNewInterval = 0;
    DWORD   dwFromLast;

    if(!pts->task.ptaskHead)
    {
        return;
    }

    if(pts->task.ptaskCur)
    {
        if(!pts->task.fSuspended)
        {
            pts->task.fSuspended = TRUE;
            TaskmanResetTimer();
        }
        return;
    }

    if(pts->task.ptaskNext == NULL)
    {
        pts->task.ptaskNext = pts->task.ptaskHead;
    }

    ptask  = pts->task.ptaskNext;
    dwTick = GetTickCount();

    // Run one 0-interval  task first
    for(;;)
    {
        if(!ptask->TestFlag(TASKF_BLOCKED|TASKF_TERMINATED) && 0==ptask->_dwTickInterval)
        {
            pts->task.ptaskNext = ptask->_ptaskNext;

            TaskmanRunTask(pts, dwTick, ptask);

            break;
        }

        ptask = ptask->_ptaskNext;

        if(ptask == NULL)
        {
            ptask = pts->task.ptaskHead;
        }

        if(ptask == pts->task.ptaskNext)
        {
            break;
        }
    }

    if(pts->task.cInterval)
    {
        // Now run all interval tasks that need to run
        ptask = pts->task.ptaskHead;

        // We cache the next task in the thread task, so we can fix it
        // when tasks are destroyed as result of being run.
        //pptaskNext = &pts->task.ptaskNextInLoop;
        for(;;)
        {
            // Before we run the task we should get the next task, since TaskmanRunTask
            // might delete any task on the chain including ptask. If the next task is deleted
            // as result of calling TaskmanRunTask, pts->task.ptaskNextInLoop will be updated.
            pts->task.ptaskNextInLoop = ptask->_ptaskNext;//*pptaskNext = ptask->_ptaskNext;

            // skip 0 tasks
            if(!ptask->TestFlag(TASKF_BLOCKED|TASKF_TERMINATED) && 0!=ptask->_dwTickInterval)
            {

                if(0 == ptask->_dwTickLast)
                {
                    // Force first run
                    dwFromLast = ptask->_dwTickInterval;
                }
                else
                {
                    dwFromLast = dwTick - ptask->_dwTickLast;
                }

                // If we are 10 ticks or less from running this task run it now
                if(dwFromLast+10 >= ptask->_dwTickInterval)
                {
                    TaskmanRunTask(pts, dwTick, ptask);
                }
                else
                {
                    if(dwFromLast < dwNewInterval)
                    {
                        dwNewInterval = dwFromLast;
                    }
                }
            }

            // grab the next task
            ptask = pts->task.ptaskNextInLoop;

            if(ptask == NULL)
            {
                break;
            }
        }

        if(dwNewInterval)
        {
            TaskmanSetTimer(dwNewInterval);
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::CTask
//
//  Synopsis:   Constructor.
//
//----------------------------------------------------------------------------
CTask::CTask(BOOL fBlocked)
{
#ifdef _DEBUG
    _dwThreadId = GetCurrentThreadId();
#endif

    _ulRefs = 1;

    if(fBlocked)
    {
        SetFlag(TASKF_BLOCKED);
    }

    TaskmanEnqueue(this);
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::~CTask
//
//  Synopsis:   Destructor.
//
//----------------------------------------------------------------------------
CTask::~CTask()
{
    Assert(_dwThreadId == GetCurrentThreadId());

    TaskmanDequeue(this);
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::AddRef
//
//  Synopsis:   Per IUnknown
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CTask::AddRef()
{
    Assert(_dwThreadId == GetCurrentThreadId());

    return (++_ulRefs);
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::Release
//
//  Synopsis:   Per IUnknown
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CTask::Release()
{
    Assert(_dwThreadId == GetCurrentThreadId());

    if(--_ulRefs == 0)
    {
        Terminate();

        if(TestFlag(TASKF_INRUN))
        {
            SetFlag(TASKF_DODESTRUCT);
        }
        else
        {
            delete this;
        }

        return 0;
    }

    return _ulRefs;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::SetInterval
//
//  Synopsis:   Sets the periodic interval for this task
//
//----------------------------------------------------------------------------
void CTask::SetInterval(DWORD dwTick)
{
    Assert(_dwThreadId == GetCurrentThreadId());

    if(dwTick != _dwTickInterval)
    {
        if(!TestFlag(TASKF_BLOCKED) && !!dwTick != !!_dwTickInterval)
        {
            THREADSTATE* pts = GetThreadState();

            if(dwTick == 0)
            {
                pts->task.cInterval -= 1;
            }
            else
            {
                pts->task.cInterval += 1;
            }
        }

        _dwTickInterval = dwTick;

        TaskmanResetTimer();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::SetBlocked
//
//  Synopsis:   Blocks or unblocks the task
//
//----------------------------------------------------------------------------
void CTask::SetBlocked(BOOL fBlocked)
{
    Assert(_dwThreadId == GetCurrentThreadId());

    if(!!fBlocked != !!TestFlag(TASKF_BLOCKED))
    {
        THREADSTATE* pts = GetThreadState();

        if(fBlocked)
        {
            SetFlag(TASKF_BLOCKED);
            pts->task.cUnblocked -= 1;

            if(_dwTickInterval > 0)
            {
                pts->task.cInterval -= 1;
            }
        }
        else
        {
            ClearFlag(TASKF_BLOCKED);
            pts->task.cUnblocked += 1;

            if(_dwTickInterval > 0)
            {
                pts->task.cInterval += 1;
            }
        }

        TaskmanResetTimer();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::Terminate
//
//  Synopsis:   Requests that task be terminated.  As soon as possible,
//              the task's OnTerminate method will be called (typically
//              before this method returns).  Actual object destruction
//              doesn't happen until all references are released.
//
//----------------------------------------------------------------------------
void CTask::Terminate()
{
    Assert(_dwThreadId == GetCurrentThreadId());

    if(!TestFlag(TASKF_TERMINATED))
    {
        if(TestFlag(TASKF_INRUN))
        {
            SetFlag(TASKF_DOTERM);
        }
        else
        {
            OnTerminate();
            TaskmanDequeue(this);
        }

        SetFlag(TASKF_TERMINATED);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::OnTerminate
//
//  Synopsis:   Base implementation.  Does nothing.
//
//----------------------------------------------------------------------------
void CTask::OnTerminate()
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CTask::OnTaskTick
//
//  Synopsis:   Callback timer procedure.
//
//----------------------------------------------------------------------------
HRESULT CTask::OnTaskTick(UINT idTimer)
{
    TaskmanRun();
    return(S_OK);
}
