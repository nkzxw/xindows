
#ifndef __XINDOWS_CORE_BASE_TASK_H__
#define __XINDOWS_CORE_BASE_TASK_H__

class CTask;

enum
{
    TASKF_BLOCKED       = 0x00000001,   // The task is explicitly blocked
    TASKF_TERMINATED    = 0x00000002,   // The Term function has been called
    TASKF_INRUN         = 0x00000004,   // The OnRun function is being called
    TASKF_DOTERM        = 0x00000008,   // Call OnTerm when OnRun returns
    TASKF_DODESTRUCT    = 0x00000010,   // Call destructor when OnRun returns
    TASKF_ENQUEUED      = 0x00000020    // Task is enqueued to task manager
};

// Functions ------------------------------------------------------------------
void TaskmanRunTask(THREADSTATE* pts, DWORD dwTick, CTask* ptask);


class NOVTABLE CTask : public CVoid
{
private:
    DECLARE_MEMCLEAR_NEW_DELETE()

public:
    static void TaskmanRun();

protected:
    static void TaskmanResetTimer();
    static void TaskmanSetTimer(DWORD dwTimeOut);
    static void TaskmanEnqueue(CTask* ptask);
    static void TaskmanDequeue(CTask* ptask);
    static void TaskmanRunTask(THREADSTATE* pts, DWORD dwTick, CTask* ptask);

public:
    CTask(BOOL fBlocked = FALSE);
    virtual ~CTask();

    virtual void OnRun(DWORD dwTimeout) = 0;
    virtual void OnTerminate() = 0;

    STDMETHOD_(ULONG,AddRef)();
    STDMETHOD_(ULONG,Release)();

    BOOL TestFlag(DWORD dwFlag)     { return(!!(_dwFlags & dwFlag)); }
    void SetFlag(DWORD dwFlag)      { _dwFlags |= dwFlag; }
    void ClearFlag(DWORD dwFlag)    { _dwFlags &= ~dwFlag; }
    void SetInterval(DWORD dwTick);
    void SetBlocked(BOOL fBlocked);
    void Terminate();

    NV_DECLARE_ONTICK_METHOD(OnTaskTick, ontasktick, (UINT idTimer));

protected:
    CTask*          _ptaskNext;         // Linked list of tasks
    ULONG           _ulRefs;            // Task reference count
    DWORD           _dwTickLast;        // Time task was last run
    DWORD           _dwTickInterval;    // Periodic interval for task
    DWORD           _dwFlags;           // See TASKF_* flags

#ifdef _DEBUG
    DWORD           _dwThreadId;        // Thread which created the task
#endif
};

#endif //__XINDOWS_CORE_BASE_TASK_H__