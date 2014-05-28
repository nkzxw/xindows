
#ifndef __XINDOWSUTIL_DLL_DYNAMICLIBRARY_H__
#define __XINDOWSUTIL_DLL_DYNAMICLIBRARY_H__

struct DYNLIB
{
    HINSTANCE   hinst;
    DYNLIB*     pdynlibNext;
    char        achName[];
};

struct DYNPROC
{
    void*       pfn;
    DYNLIB*     pdynlib;
    char*       achName;
};

XINDOWS_PUBLIC HRESULT LoadProcedure(DYNPROC* pdynproc);
XINDOWS_PUBLIC HRESULT FreeDynlib(DYNLIB* pdynlib);

#endif //__XINDOWSUTIL_DLL_DYNAMICLIBRARY_H__