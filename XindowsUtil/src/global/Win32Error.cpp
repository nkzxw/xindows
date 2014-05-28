
#include "stdafx.h"
#include "Win32Error.h"

//+------------------------------------------------------------------------
//
//  Function:   GetLastWin32Error
//
//  Synopsis:   Returns the last Win32 error, converted to an HRESULT.
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT GetLastWin32Error( )
{
    DWORD dw = GetLastError();
    return dw?HRESULT_FROM_WIN32(dw):E_FAIL;
}
