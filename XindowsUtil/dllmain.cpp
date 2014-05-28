// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"

#ifdef _DEBUG
extern void     _DbgDllProcessAttach();
extern void     _DbgDllProcessDetach();
extern void     _DbgDllThreadDetach();

#define DbgDllProcessAttach()   _DbgDllProcessAttach()
#define DbgDllProcessDetach()   _DbgDllProcessDetach()
#define DbgDllThreadDetach()    _DbgDllThreadDetach()
#else //_DEBUG
#define DbgDllProcessAttach()
#define DbgDllProcessDetach()
#define DbgDllThreadDetach()
#endif //!_DEBUG

HANDLE g_hProcessHeap = NULL;


//+---------------------------------------------------------------------------
//
//  Class:      CUnloadLibraries
//
//  Purpose:    Special class with a destructor that will be called after
//              all other static destructors are called. This ensures that
//              we don't free any of the libraries we loaded until we're
//              completely cleaned up. Win95 apparently likes to clean up
//              DLLs a little too aggressively when you call FreeLibrary
//              inside DLL_PROCESS_DETACH handlers.
//
//----------------------------------------------------------------------------
class CUnloadLibraries
{
public:
    ~CUnloadLibraries();
};

CUnloadLibraries g_CUnloadLibs;

//+---------------------------------------------------------------------------
//
//  Member:     CUnloadLibraries::~CUnloadLibraries, public
//
//  Synopsis:   class dtor
//
//  Notes:      The init_seg pragma ensures this dtor is called after all
//              others.
//
//----------------------------------------------------------------------------
CUnloadLibraries::~CUnloadLibraries()
{
    void DeinitDynamicLibraries();
    DeinitDynamicLibraries();
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
        // Prepare global variable's critical section
        CGlobalLock::Init();

        g_hProcessHeap = GetProcessHeap();
        DbgDllProcessAttach();
        break;
	case DLL_THREAD_ATTACH:
        break;
	case DLL_THREAD_DETACH:
        DbgDllThreadDetach();
        break;
	case DLL_PROCESS_DETACH:
        DbgDllProcessDetach();

        // Delete global variable's critical section
        CGlobalLock::Deinit();
		break;
	}
	return TRUE;
}

