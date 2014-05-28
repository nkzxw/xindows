
#ifndef __XINDOWS_CORE_UTIL_COLORINFO_H__
#define __XINDOWS_CORE_UTIL_COLORINFO_H__

class CColorInfo
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CColorInfo();
    CColorInfo(DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hicTargetDev, unsigned int cColorMax=256);
    BOOL AddColor(BYTE red, BYTE green, BYTE blue)
    {
        if(_cColors < _cColorsMax)
        {
            _aColors[_cColors].peRed = red;
            _aColors[_cColors].peGreen = green;
            _aColors[_cColors].peBlue = blue;
            _cColors++;

            return TRUE;
        }
        return FALSE;
    }

    HRESULT AddColors(unsigned int cColors, COLORREF* pColors);
    HRESULT AddColors(unsigned int cColors, RGBQUAD* pColors);
    HRESULT AddColors(unsigned int cColors, PALETTEENTRY* pColors);
    HRESULT AddColors(HPALETTE hpal);
    HRESULT AddColors(LOGPALETTE* pLogPalette);
    HRESULT AddColors(IViewObject* pVO);
    HRESULT GetColorSet(LPLOGPALETTE* ppColorSet);
    unsigned GetCount() { return _cColors; };
    BOOL IsFull();

private:
    DWORD           _dwDrawAspect;
    LONG            _lindex;
    void*           _pvAspect;
    DVTARGETDEVICE* _ptd;
    HDC             _hicTargetDev;
    unsigned int    _cColorsMax;

    // Don't move the following three items around
    // We use them as a LOGPALETTE
    WORD            _palVersion;
    unsigned        _cColors;
    PALETTEENTRY    _aColors[256];
};

#endif //__XINDOWS_CORE_UTIL_COLORINFO_H__