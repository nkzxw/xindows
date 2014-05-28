
#include "stdafx.h"
#include "ColorInfo.h"

CColorInfo::CColorInfo() :
_dwDrawAspect(DVASPECT_CONTENT),
_lindex(-1),
_pvAspect(NULL),
_ptd(NULL),
_hicTargetDev(NULL),
_cColors(0),
_cColorsMax(256)
{
}

CColorInfo::CColorInfo(DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hicTargetDev, unsigned int cColorsMax) :
_dwDrawAspect(dwDrawAspect),
_lindex(lindex),
_pvAspect(pvAspect),
_ptd(ptd),
_hicTargetDev(hicTargetDev),
_cColors(0)
{
    _cColorsMax = max(256, cColorsMax);
}

HRESULT CColorInfo::AddColors(HPALETTE hpal)
{
    unsigned int cColors = GetPaletteEntries(hpal, 0, 0, NULL);
    cColors = min(_cColorsMax-_cColors, cColors);

    if(cColors)
    {
        GetPaletteEntries(hpal, 0, cColors, &_aColors[_cColors]);
        _cColors += cColors;
    }
    RRETURN(S_OK);
}

HRESULT CColorInfo::AddColors(LOGPALETTE* pLogPal)
{
    RRETURN(AddColors(pLogPal->palNumEntries, pLogPal->palPalEntry));
}

HRESULT CColorInfo::AddColors(unsigned int cColors, PALETTEENTRY* pColors)
{
    // The check for system colors assumes they occur in the same order as
    // the real system colors.  This is the simplest possible fix for now.

    // Remove any system colors at the end (do this first so that cColors makes sense)
    Assert(g_lpHalftone.wCnt == 256);
    Assert(sizeof(PALETTEENTRY) == sizeof(DWORD));

    DWORD* pSystem = (DWORD*)&g_lpHalftone.ape[g_lpHalftone.wCnt-1];
    DWORD* pInput = (DWORD*)&pColors[cColors-1];
    while(cColors && (*pSystem==*pInput))
    {
        pSystem--;
        pInput--;
        cColors--;
    }

    // Remove any system colors at the beginning
    pSystem = (DWORD*)(g_lpHalftone.ape);
    pInput = (DWORD*)pColors;
    while(cColors && (*pSystem==*pInput))
    {
        pSystem++;
        pInput++;
        cColors--;
    }
    pColors = (PALETTEENTRY*)pInput;

    cColors = min(_cColorsMax-_cColors, cColors);
    if(cColors)
    {
        memcpy(&_aColors[_cColors], pColors, cColors*sizeof(PALETTEENTRY));
        _cColors += cColors;
    }

    RRETURN(S_OK);
}

HRESULT CColorInfo::AddColors(unsigned int cColors, RGBQUAD* pColors)
{
    cColors = min(_cColorsMax-_cColors, cColors);
    if(cColors)
    {
        CopyPaletteEntriesFromColors(&_aColors[_cColors], pColors, cColors);
        _cColors += cColors;
    }

    RRETURN(S_OK);
}

HRESULT CColorInfo::AddColors(unsigned int cColors, COLORREF* pColors)
{
    cColors = min(_cColorsMax-_cColors, cColors);
    while(cColors)
    {
        _aColors[_cColors].peRed    = GetRValue(*pColors);
        _aColors[_cColors].peGreen  = GetGValue(*pColors);
        _aColors[_cColors].peBlue   = GetBValue(*pColors);
        _aColors[_cColors].peFlags  = 0;
        pColors++;
    }

    RRETURN(S_OK);
}

HRESULT CColorInfo::AddColors(IViewObject* pVO)
{
    LPLOGPALETTE pColors = NULL;
    HRESULT hr = pVO->GetColorSet(_dwDrawAspect, _lindex, _pvAspect, _ptd, _hicTargetDev, &pColors);
    if(FAILED(hr))
    {
        hr = E_NOTIMPL;
    }
    else if(hr==S_OK && pColors)
    {
        AddColors(pColors);
        CoTaskMemFree(pColors);
    }

    RRETURN1(hr, S_FALSE);
}

BOOL CColorInfo::IsFull()
{
    Assert(_cColors <= _cColorsMax);
    return _cColors>=_cColorsMax;
}

HRESULT CColorInfo::GetColorSet(LPLOGPALETTE* ppColors)
{
    *ppColors = 0;
    if(_cColors == 0)
    {
        RRETURN1(S_FALSE, S_FALSE);
    }

    LOGPALETTE* pColors;

    // It's just easier to allocate 256 colors instead of messing about.
    *ppColors = pColors = (LOGPALETTE*)CoTaskMemAlloc(sizeof(LOGPALETTE)+255*sizeof(PALETTEENTRY));

    if(!pColors)
    {
        RRETURN(E_OUTOFMEMORY);
    }

    // This will ensure that we have a reasonable set of colors, including
    // the system colors to start
    memcpy(pColors, &g_lpHalftone, sizeof(g_lpHalftone));

    unsigned int cColors = min(236, _cColors);

    // Notice that we avoid overwriting the beginning of the _aColors array.
    // The assumption is that the colors are in some kind of order.
    memcpy(pColors->palPalEntry+10, _aColors, cColors*sizeof(PALETTEENTRY));

    for(unsigned int i=10 ; i<(cColors+10); i++)
    {
        pColors->palPalEntry[i].peFlags = PC_NOCOLLAPSE;
    }

    RRETURN(S_OK);
}