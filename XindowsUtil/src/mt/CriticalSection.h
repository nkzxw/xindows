
#ifndef __XINDOWSUTIL_MT_CRITICALSECTION_H__
#define __XINDOWSUTIL_MT_CRITICALSECTION_H__

class XINDOWS_PUBLIC CCriticalSection
{
public:
    CCriticalSection()  { InitializeCriticalSection(&_cs); }
    ~CCriticalSection() { DeleteCriticalSection(&_cs); }
    void Enter()        { EnterCriticalSection(&_cs); }
    void Leave()        { LeaveCriticalSection(&_cs); }
    CRITICAL_SECTION* GetPcs() { return(&_cs); }
private:
    CRITICAL_SECTION _cs;
};

class XINDOWS_PUBLIC CCriticalSectionLock
{
public:
    CCriticalSectionLock(CCriticalSection& cs) { (_pcs=&cs)->Enter(); }
    ~CCriticalSectionLock() { _pcs->Leave(); }
private:
    CCriticalSection* _pcs;
};

#define LOCK_SECTION(cs)    CCriticalSectionLock cs##_lock(cs)

#endif //__XINDOWSUTIL_MT_CRITICALSECTION_H__