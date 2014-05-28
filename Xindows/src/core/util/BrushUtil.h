
#ifndef __XINDOWS_CORE_UTIL_BRUSHUTIL_H__
#define __XINDOWS_CORE_UTIL_BRUSHUTIL_H__

HBRUSH  GetCachedBrush(COLORREF color);
void    ReleaseCachedBrush(HBRUSH hbr);
void    SelectCachedBrush(HDC hdc, COLORREF crNew, HBRUSH* phbrNew, HBRUSH* phbrOld, COLORREF* pcrNow);

HBRUSH  GetCachedBmpBrush(int resId);

void    PatBltBrush(HDC hdc, LONG x, LONG y, LONG xWid, LONG yHei, DWORD dwRop, COLORREF cr);
void    PatBltBrush(HDC hdc, RECT* prc, DWORD dwRop, COLORREF cr);

#endif //__XINDOWS_CORE_UTIL_BRUSHUTIL_H__