
#include "stdafx.h"
#include "GrabUtil.h"

//+------------------------------------------------------------------------
//
//  Function:   ColorDiff
//
//  Synopsis:   Computes the color difference amongst two colors
//
//-------------------------------------------------------------------------
DWORD ColorDiff(COLORREF c1, COLORREF c2)
{
#define __squareit(n)((DWORD)((n)*(n)))
    return (__squareit((INT)GetRValue(c1)-(INT)GetRValue(c2)) +
        __squareit((INT)GetGValue(c1)-(INT)GetGValue(c2)) +
        __squareit((INT)GetBValue(c1)-(INT)GetBValue(c2)));
#undef __squareit
}

//+------------------------------------------------------------------------
//
//  Function:   PatBltRectH & PatBltRectV
//
//  Synopsis:   PatBlts the top/bottom and left/right.
//
//-------------------------------------------------------------------------
static void PatBltRectH(HDC hDC, RECT* prc, int cThick, DWORD dwRop)
{
    PatBlt(hDC, prc->left, prc->top, prc->right-prc->left, cThick, dwRop);
    PatBlt(hDC, prc->left, prc->bottom-cThick, prc->right-prc->left, cThick, dwRop);
}

static void PatBltRectV(HDC hDC, RECT* prc, int cThick, DWORD dwRop)
{
    PatBlt(hDC, prc->left, prc->top+cThick, cThick, (prc->bottom-prc->top)-(2*cThick), dwRop);
    PatBlt(hDC, prc->right-cThick, prc->top+cThick, cThick, (prc->bottom-prc->top)-(2*cThick), dwRop);
}

//+--------------------------------------------------------------------
//
//  Global helpers.
//
//---------------------------------------------------------------------
void PatBltRect(HDC hDC, RECT* prc, int cThick, DWORD dwRop)
{
    PatBltRectH(hDC, prc, cThick, dwRop);
    PatBltRectV(hDC, prc, cThick, dwRop);
}
