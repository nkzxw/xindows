
//+---------------------------------------------------------------------------
//
//  File:       DynamicLibrary.cpp
//
//  Contents:   Utility for dynamically loaded procedures.
//
//----------------------------------------------------------------------------

#include "stdafx.h"
#include "DynamicLibrary.h"

static DYNLIB* s_pdynlibHead;

//+---------------------------------------------------------------------------
//
//  Function:   DeinitDynamicLibraries
//
//  Synopsis:   Undoes the work of LoadProcedure.
//
//----------------------------------------------------------------------------
void DeinitDynamicLibraries()
{
    DYNLIB* pdynlib;

    for(pdynlib=s_pdynlibHead; pdynlib; pdynlib=pdynlib->pdynlibNext)
    {
        Assert(pdynlib->hinst);
        FreeLibrary(pdynlib->hinst);
        pdynlib->hinst = NULL;
    }
    s_pdynlibHead = NULL;
}

//+---------------------------------------------------------------------------
//
//  Function:   LoadProcedure
//
//  Synopsis:   Load library and get address of procedure.
//
//              Declare DYNLIB and DYNPROC globals describing the procedure.
//              Note that several DYNPROC structures can point to a single
//              DYNLIB structure.
//
//                  DYNLIB g_dynlibOLEDLG = { NULL, "OLEDLG.DLL" };
//                  DYNPROC g_dynprocOleUIInsertObjectA =
//                          { NULL, &g_dynlibOLEDLG, "OleUIInsertObjectA" };
//                  DYNPROC g_dynprocOleUIPasteSpecialA =
//                          { NULL, &g_dynlibOLEDLG, "OleUIPasteSpecialA" };
//
//              Call LoadProcedure to load the library and get the procedure
//              address.  LoadProcedure returns immediatly if the procedure
//              has already been loaded.
//
//                  hr = LoadProcedure(&g_dynprocOLEUIInsertObjectA);
//                  if (hr)
//                      goto Error;
//
//                  uiResult = (*(UINT (__stdcall *)(LPOLEUIINSERTOBJECTA))
//                      g_dynprocOLEUIInsertObjectA.pfn)(&ouiio);
//
//              Release the library at shutdown.
//
//                  void DllProcessDetach()
//                  {
//                      DeinitDynamicLibraries();
//                  }
//
//  Arguments:  pdynproc  Descrition of library and procedure to load.
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT LoadProcedure(DYNPROC* pdynproc)
{
    HINSTANCE   hinst;
    DYNLIB*     pdynlib = pdynproc->pdynlib;
    DWORD       dwError;

    if(pdynproc->pfn && pdynlib->hinst)
    {
        return S_OK;
    }

    if(!pdynlib->hinst)
    {
        USES_CONVERSION;
        Trace1("Loading library %s.\n", A2T(pdynlib->achName));

        // Try to load the library using the normal mechanism.
        hinst = LoadLibraryA(pdynlib->achName);

        // If that failed because the module was not be found,
        // then try to find the module in the directory we were
        // loaded from.
        if(!hinst)
        {
            dwError = GetLastError();

            if(dwError==ERROR_MOD_NOT_FOUND || dwError==ERROR_DLL_NOT_FOUND)
            {
                char achBuf1[MAX_PATH];
                char achBuf2[MAX_PATH];
                char* pch;

                // Get path name of this module.
                if(GetModuleFileNameA(NULL, achBuf1, ARRAYSIZE(achBuf1)) == 0)
                {
                    goto Error;
                }

                // Find where the file name starts in the module path.
                if(GetFullPathNameA(achBuf1, ARRAYSIZE(achBuf2), achBuf2, &pch) == 0)
                {
                    goto Error;
                }

                // Chop off the file name to get a directory name.
                *pch = 0;

                // See if there's a dll with the given name in the directory.
                if(SearchPathA(
                    achBuf2,
                    pdynlib->achName,
                    NULL,
                    ARRAYSIZE(achBuf1),
                    achBuf1,
                    NULL) != 0)
                {
                    // Yes, there's a dll. Load it.
                    hinst = LoadLibraryExA(achBuf1, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
                }
            }
        }
        if(!hinst)
        {
            goto Error;
        }

        // Link into list for DeinitDynamicLibraries
        {
            LOCK_GLOBALS;

            if(pdynlib->hinst)
            {
                FreeLibrary(hinst);
            }
            else
            {
                pdynlib->hinst = hinst;
                pdynlib->pdynlibNext = s_pdynlibHead;
                s_pdynlibHead = pdynlib;
            }
        }
    }

    pdynproc->pfn = GetProcAddress(pdynlib->hinst, pdynproc->achName);
    if(!pdynproc->pfn)
    {
        goto Error;
    }

    return S_OK;

Error:
    RRETURN(GetLastWin32Error());
}

//+---------------------------------------------------------------------------
//
//  Function:   FreeDynlib
//
//  Synopsis:   Free a solitary dynlib entry from the link list of dynlibs
//
//  Arguments:  Pointer to DYNLIB to be freed
//
//  Returns:    S_OK
//
//----------------------------------------------------------------------------
HRESULT FreeDynlib(DYNLIB* pdynlib)
{
    DYNLIB** ppdynlibPrev;

    LOCK_GLOBALS;

    if(pdynlib == s_pdynlibHead)
    {
        ppdynlibPrev = &s_pdynlibHead;
    }
    else
    {
        DYNLIB* pdynlibPrev;

        for(pdynlibPrev=s_pdynlibHead;
            pdynlibPrev&&pdynlibPrev->pdynlibNext!=pdynlib;
            pdynlibPrev=pdynlibPrev->pdynlibNext) ;

        ppdynlibPrev = pdynlibPrev ? &pdynlibPrev->pdynlibNext : NULL;
    }

    if(ppdynlibPrev)
    {
        if(pdynlib->hinst)
        {
            FreeLibrary(pdynlib->hinst);
            pdynlib->hinst = NULL;
        }

        *ppdynlibPrev = pdynlib->pdynlibNext;
    }

    return S_OK;
}