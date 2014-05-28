
#ifndef __XINDOWSUTIL_MT_GLOBALLOCK_H__
#define __XINDOWSUTIL_MT_GLOBALLOCK_H__

//+----------------------------------------------------------------------------
//
//  Class:      CGlobalLock
//
//  Synopsis:   Smart object to lock/unlock access to all global variables
//              Declare an instance of this class within the scope that needs
//              guarded access. Class instances may be nested.
//
//  Usage:      Lock variables by using the LOCK_GLOBALS marco
//              Simply include this macro within the appropriate scope (as
//              small as possible) to protect access. For example:
//
//-----------------------------------------------------------------------------
class XINDOWS_PUBLIC CGlobalLock
{
public:
    CGlobalLock();
    ~CGlobalLock();

#ifdef _DEBUG
    static BOOL IsThreadLocked();
#endif //_DEBUG

    // Process attach/detach routines
    static void Init();
    static void Deinit();
};

#define LOCK_GLOBALS    CGlobalLock glock

#endif //__XINDOWSUTIL_MT_GLOBALLOCK_H__