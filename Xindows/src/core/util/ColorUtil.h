
#ifndef __XINDOWS_CORE_UTIL_COLORUTIL_H__
#define __XINDOWS_CORE_UTIL_COLORUTIL_H__

BOOL        IsOleColorValid(OLE_COLOR clr);
COLORREF    ColorRefFromOleColor(OLE_COLOR clr);
HPALETTE    GetDefaultPalette(HDC hdc=NULL);
BOOL        IsHalftonePalette(HPALETTE hpal);
void        CopyColorsFromPaletteEntries(RGBQUAD* prgb, const PALETTEENTRY* ppe, UINT uCount);
void        CopyPaletteEntriesFromColors(PALETTEENTRY* ppe, const RGBQUAD* prgb, UINT uCount);
HDC         GetMemoryDC();
void        ReleaseMemoryDC(HDC hdc);
COLORREF    GetSysColorQuick(int i);

inline unsigned GetPaletteSize(LOGPALETTE* pColors)
{
    Assert(pColors);
    Assert(pColors->palVersion == 0x300);

    return (sizeof(LOGPALETTE)-sizeof(PALETTEENTRY)+pColors->palNumEntries*sizeof(PALETTEENTRY));
}

inline int ComparePalettes(LOGPALETTE* pColorsLeft, LOGPALETTE* pColorsRight)
{
    // This counts on the fact that the second member of LOGPALETTE is the size
    // so if the sizes don't match, we'll stop long before either one ends.  If
    // the sizes are equal then GetPaletteSize(pColorsLeft) is the maximum size
    return memcmp(pColorsLeft, pColorsRight, GetPaletteSize(pColorsLeft));
}

struct LOGPAL256
{
    WORD            wVer;
    WORD            wCnt;
    PALETTEENTRY    ape[256];
};

extern HPALETTE     g_hpalHalftone;
extern LOGPAL256    g_lpHalftone;
extern RGBQUAD      g_rgbHalftone[];

class CYIQ
{
public:
    CYIQ(COLORREF c)
    {
        RGB2YIQ(c);
    }

    inline void RGB2YIQ(COLORREF c);
    inline void YIQ2RGB(COLORREF* pc);

    // Data members
    int _Y, _I, _Q;

    // Constants
    enum { MAX_LUM=255, MIN_LUM=0 };
};

void ContrastColors(COLORREF& c1, COLORREF& c2, int LumDiff);

#endif //__XINDOWS_CORE_UTIL_COLORUTIL_H__