
#include "stdafx.h"
#include "Himetric.h"

//+---------------------------------------------------------------
//
//  Function:   HimetricFromHPix
//
//  Synopsis:   Converts horizontal pixel units to himetric units.
//
//----------------------------------------------------------------
long HimetricFromHPix(int iPix)
{
    Assert(_afxGlobalData._sizePixelsPerInch.cx);

    return MulDivQuick(iPix, HIMETRIC_PER_INCH, _afxGlobalData._sizePixelsPerInch.cx);
}

//+---------------------------------------------------------------
//
//  Function:   HimetricFromVPix
//
//  Synopsis:   Converts vertical pixel units to himetric units.
//
//----------------------------------------------------------------
long HimetricFromVPix(int iPix)
{
    Assert(_afxGlobalData._sizePixelsPerInch.cy);

    return MulDivQuick(iPix, HIMETRIC_PER_INCH, _afxGlobalData._sizePixelsPerInch.cy);
}

//+---------------------------------------------------------------
//
//  Function:   HPixFromHimetric
//
//  Synopsis:   Converts himetric units to horizontal pixel units.
//
//----------------------------------------------------------------
int HPixFromHimetric(long lHi)
{
    Assert(_afxGlobalData._sizePixelsPerInch.cx);

    return MulDivQuick(_afxGlobalData._sizePixelsPerInch.cx, lHi, HIMETRIC_PER_INCH);
}

//+---------------------------------------------------------------
//
//  Function:   VPixFromHimetric
//
//  Synopsis:   Converts himetric units to vertical pixel units.
//
//----------------------------------------------------------------
int VPixFromHimetric(long lHi)
{
    Assert(_afxGlobalData._sizePixelsPerInch.cy);

    return MulDivQuick(_afxGlobalData._sizePixelsPerInch.cy, lHi, HIMETRIC_PER_INCH);
}

//+-------------------------------------------------------------------------
//
//  Function:   UserFromHimetric
//
//  Synopsis:   Converts a himetric long value to point size float.
//
//--------------------------------------------------------------------------
float UserFromHimetric(long lValue)
{
    return ((float)MulDivQuick(lValue, 72*20, 2540))/20;
}

//+-------------------------------------------------------------------------
//
//  Function:   HimetricFromUser
//
//  Synopsis:   Converts a point size double to himetric long.  Rounds
//              to the nearest himetric value.
//
//--------------------------------------------------------------------------
long HimetricFromUser(float xf)
{
    long lResult;

    xf = xf * (((float)2540) / 72);

    if(xf > LONG_MAX)
    {
        lResult = LONG_MAX;
    }
    else if(xf > .0)
    {
        lResult = long(xf + .5);
    }
    else if(xf < LONG_MIN)
    {
        lResult = LONG_MIN;
    }
    else
    {
        lResult = long(xf - .5);
    }

    return lResult;
}