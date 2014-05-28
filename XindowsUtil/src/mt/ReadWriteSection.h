
#ifndef __XINDOWSUTIL_MT_READWRITESECTION_H__
#define __XINDOWSUTIL_MT_READWRITESECTION_H__

class XINDOWS_PUBLIC CReadWriteSection
{
public:
    CReadWriteSection()
    {
        InitializeCriticalSection(&_csEnterRead);
        InitializeCriticalSection(&_csWriter);
        _lReaders = -1;
    }
    ~CReadWriteSection()
    {
        DeleteCriticalSection(&_csEnterRead);
        DeleteCriticalSection(&_csWriter);
    }
    void EnterRead()
    {
        EnterCriticalSection(&_csEnterRead);
        if(!InterlockedIncrement(&_lReaders))
        {
            EnterCriticalSection(&_csWriter);
        }
        LeaveCriticalSection(&_csEnterRead);
    }
    void LeaveRead()
    {
        if(InterlockedDecrement(&_lReaders) < 0)
        {
            LeaveCriticalSection(&_csWriter);
        }
    }
    void EnterWrite()
    {
        EnterCriticalSection(&_csEnterRead);
        EnterCriticalSection(&_csWriter);
    }
    void LeaveWrite()
    {
        LeaveCriticalSection(&_csWriter);
        LeaveCriticalSection(&_csEnterRead);
    }

    // Note that readers will continue reading underneath this critical section
    CRITICAL_SECTION* GetPcs()
    {
        return &_csEnterRead;
    }

private:
    CRITICAL_SECTION    _csEnterRead;
    CRITICAL_SECTION    _csWriter;
    LONG                _lReaders;
};

class XINDOWS_PUBLIC CReadLock
{
public:
    CReadLock(CReadWriteSection& crws) { (_prws=&crws)->EnterRead(); }
    ~CReadLock() { _prws->LeaveRead(); }
private:
    CReadWriteSection* _prws;
};

class XINDOWS_PUBLIC CWriteLock
{
public:
    CWriteLock(CReadWriteSection& crws) { (_prws=&crws)->EnterWrite(); }
    ~CWriteLock() { _prws->LeaveWrite(); }
private:
    CReadWriteSection* _prws;
};

#define LOCK_READ(crws)     CReadLock  crws##_readlock(crws)
#define LOCK_WRITE(crws)    CWriteLock crws##_writelock(crws)

#endif //__XINDOWSUTIL_MT_READWRITESECTION_H__