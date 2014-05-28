
#ifndef __XINDOWS_SITE_UTIL_GRABUTIL_H__
#define __XINDOWS_SITE_UTIL_GRABUTIL_H__

//+---------------------------------------------------------------------------
//
//  Enumeration: HTC
//
//  Synopsis:    Success values for hit testing.
//
//----------------------------------------------------------------------------
enum HTC
{
    HTC_NO                = 0,
    HTC_MAYBE             = 1,
    HTC_YES               = 2,

    HTC_HSCROLLBAR        = 3,
    HTC_VSCROLLBAR        = 4,

    HTC_LEFTGRID          = 5,
    HTC_TOPGRID           = 6,
    HTC_RIGHTGRID         = 7,
    HTC_BOTTOMGRID        = 8,

    HTC_NONCLIENT         = 9,

    HTC_EDGE              = 10,

    HTC_GRPTOPBORDER      = 11,
    HTC_GRPLEFTBORDER     = 12,
    HTC_GRPBOTTOMBORDER   = 13,
    HTC_GRPRIGHTBORDER    = 14,
    HTC_GRPTOPLEFTHANDLE  = 15,
    HTC_GRPLEFTHANDLE     = 16,
    HTC_GRPTOPHANDLE      = 17,
    HTC_GRPBOTTOMLEFTHANDLE = 18,
    HTC_GRPTOPRIGHTHANDLE = 19,
    HTC_GRPBOTTOMHANDLE   = 20,
    HTC_GRPRIGHTHANDLE    = 21,
    HTC_GRPBOTTOMRIGHTHANDLE = 22,

    HTC_TOPBORDER         = 23,
    HTC_LEFTBORDER        = 24,
    HTC_BOTTOMBORDER      = 25,
    HTC_RIGHTBORDER       = 26,

    HTC_TOPLEFTHANDLE     = 27,
    HTC_LEFTHANDLE        = 28,
    HTC_TOPHANDLE         = 29,
    HTC_BOTTOMLEFTHANDLE  = 30,
    HTC_TOPRIGHTHANDLE    = 31,
    HTC_BOTTOMHANDLE      = 32,
    HTC_RIGHTHANDLE       = 33,
    HTC_BOTTOMRIGHTHANDLE = 34,
    HTC_ADORNMENT         = 35
};

// More ROP codes
#define DST_PAT_NOT_OR  0x00af0229
#define DST_PAT_AND     0x00a000c9
#define DST_PAT_OR      0x00fa0089

DWORD ColorDiff(COLORREF c1, COLORREF c2);

void PatBltRect(HDC hDC, RECT* prc, int cThick, DWORD dwRop);

#endif //__XINDOWS_SITE_UTIL_GRABUTIL_H__