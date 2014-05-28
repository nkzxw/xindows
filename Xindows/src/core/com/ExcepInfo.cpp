
#include "stdafx.h"
#include "ExcepInfo.h"

//+---------------------------------------------------------------------------
//
//  Function:   FreeEXCEPINFO
//
//  Synopsis:   Frees resources in an excepinfo.  Does not reinitialize
//              these fields.
//
//----------------------------------------------------------------------------
void FreeEXCEPINFO(EXCEPINFO* pEI)
{
    if(pEI)
    {
        FormsFreeString(pEI->bstrSource);
        FormsFreeString(pEI->bstrDescription);
        FormsFreeString(pEI->bstrHelpFile);
    }
}
