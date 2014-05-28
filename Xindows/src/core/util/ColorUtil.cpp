
#include "stdafx.h"
#include "ColorUtil.h"

COLORREF    g_acrSysColor[25];
BOOL        g_fSysColorInit = FALSE;
COLORREF    g_crPaletteRelative = 0;
HPALETTE    g_hpalHalftone = NULL;
LOGPAL256   g_lpHalftone;
RGBQUAD     g_rgbHalftone[256];
HDC         g_hdcMem1 = NULL;
HDC         g_hdcMem2 = NULL;
HDC         g_hdcMem3 = NULL;

// Inverse color table
BYTE*       g_pInvCMAP = NULL;

#ifdef _DEBUG
int         g_cGetMemDc = 0;
#endif

STDAPI FormsTranslateColor(OLE_COLOR clr, HPALETTE hpal, COLORREF* pcr)
{
    int syscolor;

    switch(clr & 0xff000000)
    {
    case 0x80000000:
        syscolor = clr & 0x00ffffff;
        if(syscolor >= ARRAYSIZE(g_acrSysColor))
        {
            RRETURN(E_INVALIDARG);
        }
        clr = GetSysColorQuick(syscolor);
        break;
    case 0x01000000:
        // check validity of index
        if(hpal)
        {
            PALETTEENTRY pe;
            // try to get palette entry, if it fails we assume index is invalid
            if(!GetPaletteEntries(hpal, (UINT)(clr&0xffff), 1, &pe))
            {
                RRETURN(E_INVALIDARG); // BUGBUG? : CDK uses CTL_E_OVERFLOW
            }
        }
        break;
    case 0x02000000:
        break;
    case 0:
        if(hpal != NULL)
        {
            clr |= 0x02000000;
        }
        break;

    default :
        RRETURN(E_INVALIDARG);
    }

    if(pcr)
    {
        *pcr = clr;
    }

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Function:   IsOleColorValid
//
//  Synopsis:   Return true if the given OLE_COLOR is valid.
//
//----------------------------------------------------------------------------
BOOL IsOleColorValid(OLE_COLOR clr)
{
    return OK(FormsTranslateColor(clr, NULL, NULL));
}

//+---------------------------------------------------------------------------
//
//  Function:   ColorRefFromOleColor
//
//  Synopsis:   Map OLE_COLORs to COLORREFs.  This function does not contain
//              any error checking.   Callers should validate the color with
//              IsOleColorValid() before calling this function.
//
//----------------------------------------------------------------------------
COLORREF ColorRefFromOleColor(OLE_COLOR clr)
{
    Assert(IsOleColorValid(clr));
    Assert((clr&0x02000000) == 0);

    if((long)clr < 0)
    {
        return GetSysColorQuick(clr&0xff);
    }
    else
    {
        return clr|g_crPaletteRelative;
    }
}

HRESULT InitPalette()
{
    HDC hdc = GetDC(NULL);
    HRESULT hr = S_OK;

    g_hpalHalftone = SHCreateShellPalette(NULL);
    if(g_hpalHalftone)
    {
        g_lpHalftone.wCnt = (WORD)GetPaletteEntries(g_hpalHalftone, 0, 256, g_lpHalftone.ape);
        g_lpHalftone.wVer = 0x0300;
        CopyColorsFromPaletteEntries(g_rgbHalftone, g_lpHalftone.ape, g_lpHalftone.wCnt);
    }
    else
    {
        return GetLastWin32Error();
    }

    // Get the dithering table
    SHGetInverseCMAP((BYTE*)&g_pInvCMAP, sizeof(BYTE*));
    if(g_pInvCMAP == NULL)
    {
        return E_OUTOFMEMORY;
    }

    if(GetDeviceCaps(hdc, RASTERCAPS) & RC_PALETTE)
    {
#ifdef _DEBUG
        PALETTEENTRY ape[20], *pe = g_lpHalftone.ape;

        GetSystemPaletteEntries(hdc, 0, 10, ape);
        GetSystemPaletteEntries(hdc, 246, 10, ape+10);

        if(memcmp(ape, pe, 10*sizeof(PALETTEENTRY))!=0
            || memcmp(ape+10, pe+246, 10*sizeof(PALETTEENTRY))!=0)
        {
            Trace0("InitPalette: Unexpected system colors detected\n");
        }
#endif
        // Generate the inverse mapping table
        g_crPaletteRelative = 0x02000000;
    }
    else
    {
        g_crPaletteRelative = 0;
    }

    ReleaseDC(NULL, hdc);

    RRETURN(hr);
}

void DeinitPalette()
{
    if(g_hdcMem1)
    {
        Trace1("GetMemoryDC, about to delete dc, count=%d\n", --++g_cGetMemDc);
        DeleteDC(g_hdcMem1);
    }
    if(g_hdcMem2)
    {
        Trace1("GetMemoryDC, about to delete dc, count=%d\n", --++g_cGetMemDc);
        DeleteDC(g_hdcMem2);
    }
    if(g_hdcMem3)
    {
        Trace1("GetMemoryDC, about to delete dc, count=%d\n", --++g_cGetMemDc);
        DeleteDC(g_hdcMem3);
    }

    if(g_hpalHalftone)
    {
        DeleteObject(g_hpalHalftone);
        g_hpalHalftone = NULL;
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   GetDefaultPalette
//
//  Synopsis:   Returns a generic (halftone) palette, optionally selecting it
//              into the DC.
//
//              Where possible, use CDoc::GetPalette instead.
//
//----------------------------------------------------------------------------
HPALETTE GetDefaultPalette(HDC hdc)
{
    HPALETTE hpal = NULL;

    if(g_crPaletteRelative)
    {
        hpal = g_hpalHalftone;
        if(hdc)
        {
            SelectPalette(hdc, hpal, TRUE);
            RealizePalette(hdc);
        }
    }
    return hpal;
}

//+---------------------------------------------------------------------------
//
//  Function:   IsHalftonePalette
//
//  Synopsis:   Returns TRUE if the passed palette exactly matches the
//              halftone palette.
//
//----------------------------------------------------------------------------
BOOL IsHalftonePalette(HPALETTE hpal)
{
    int cColors;
    PALETTEENTRY ape[256];

    if(!hpal)
    {
        return FALSE;
    }

    if(hpal == g_hpalHalftone)
    {
        return TRUE;
    }

    cColors = GetPaletteEntries(hpal, 0, 0, NULL);

    // Right number of colors?
    if(cColors != ARRAYSIZE(ape))
    {
        return FALSE;
    }

    GetPaletteEntries(hpal, 0, cColors, ape);

    return (!memcmp(ape, g_lpHalftone.ape, sizeof(PALETTEENTRY)*ARRAYSIZE(ape)));
}

void CopyColorsFromPaletteEntries(RGBQUAD* prgb, const PALETTEENTRY* ppe, UINT uCount)
{
    while(uCount--)
    {
        prgb->rgbRed   = ppe->peRed;
        prgb->rgbGreen = ppe->peGreen;
        prgb->rgbBlue  = ppe->peBlue;
        prgb->rgbReserved = 0;

        prgb++;
        ppe++;
    }
}

void CopyPaletteEntriesFromColors(PALETTEENTRY* ppe, const RGBQUAD* prgb, UINT uCount)
{
    while(uCount--)
    {
        ppe->peRed   = prgb->rgbRed;
        ppe->peGreen = prgb->rgbGreen;
        ppe->peBlue  = prgb->rgbBlue;
        ppe->peFlags = 0;

        prgb++;
        ppe++;
    }
}

HDC GetMemoryDC()
{
    HDC hdcMem;

    if(_afxGlobalData._fNoDisplayChange)
    {
        hdcMem = (HDC)InterlockedExchangePointer((void**)&g_hdcMem1, NULL);
        if(hdcMem)
        {
            Trace1("GetMemoryDC, about to return dc #1, count still %d\n", g_cGetMemDc);
            return hdcMem;
        }

        hdcMem = (HDC)InterlockedExchangePointer((void**)&g_hdcMem2, NULL);
        if(hdcMem)
        {
            Trace1("GetMemoryDC, about to return dc #2, count still %d\n", g_cGetMemDc);
            return hdcMem;
        }

        hdcMem = (HDC)InterlockedExchangePointer((void**)&g_hdcMem3, NULL);
        if(hdcMem)
        {
            Trace1("GetMemoryDC, about to return dc #3, count still %d\n", g_cGetMemDc);
            return hdcMem;
        }

        Trace1("GetMemoryDC, about to return new dc, count=%d\n", ++g_cGetMemDc);
    }

    hdcMem = CreateCompatibleDC(NULL);

    if(hdcMem)
    {
        SetStretchBltMode(hdcMem, COLORONCOLOR);
        GetDefaultPalette(hdcMem);
    }

    return(hdcMem);
}

void ReleaseMemoryDC(HDC hdcIn)
{
    HDC hdc = hdcIn;

    if(_afxGlobalData._fNoDisplayChange)
    {
        hdc = (HDC)InterlockedExchangePointer((void**)&g_hdcMem1, hdc);
        if(hdc == NULL)
        {
            Trace1("GetMemoryDC, stored a dc in #1, count still %d", g_cGetMemDc);
            return;
        }

        hdc = (HDC)InterlockedExchangePointer((void**)&g_hdcMem2, hdc);
        if(hdc == NULL)
        {
            Trace1("GetMemoryDC, stored a dc in #2, count still %d", g_cGetMemDc);
            return;
        }

        hdc = (HDC)InterlockedExchangePointer((void**)&g_hdcMem3, hdc);
        if(hdc == NULL)
        {
            Trace1("GetMemoryDC, stored a dc in #3, count still %d", g_cGetMemDc);
            return;
        }

        Trace1("GetMemoryDC, about to delete dc, count=%d", --++g_cGetMemDc);
    }

    DeleteDC(hdc);
}

//+---------------------------------------------------------------------------
//
//  Function:   InitColorTranslation
//
//  Synopsis:   Initialize system color translation table.
//
//----------------------------------------------------------------------------
void InitColorTranslation()
{
    for(int i=0; i<ARRAYSIZE(g_acrSysColor); i++)
    {
        g_acrSysColor[i] = GetSysColor(i);
    }

    g_fSysColorInit = TRUE;
}

//+---------------------------------------------------------------------------
//
//  Function:   GetSysColorQuick
//
//  Synopsis:   Looks up system colors from a cache.
//
//----------------------------------------------------------------------------
COLORREF GetSysColorQuick(int i)
{
    if(!g_fSysColorInit)
    {
        InitColorTranslation();
    }

    return g_acrSysColor[i];
}


// Function: RGB2YIQ
//
// Parameter: c -- the color in RGB.
//
// Note: Converts RGB to YIQ. The standard conversion matrix obtained
//       from Foley Van Dam -- 2nd Edn.
//
// Returns: Nothing
void CYIQ::RGB2YIQ(COLORREF c)
{
    int R = GetRValue(c);
    int G = GetGValue(c);
    int B = GetBValue(c);

    _Y = (30*R + 59*G + 11*B + 50) / 100;
    _I = (60*R - 27*G - 32*B + 50) / 100;
    _Q = (21*R - 52*G + 31*B + 50) / 100;
}

// Function: YIQ2RGB
//
// Parameter: pc [o] -- the color in RGB.
//
// Note: Converts YIQ to RGB. The standard conversion matrix obtained
//       from Foley Van Dam -- 2nd Edn.
//
// Returns: Nothing
void CYIQ::YIQ2RGB(COLORREF* pc)
{
    int R, G, B;

    R = (100*_Y +  96*_I +  62*_Q + 50) / 100;
    G = (100*_Y -  27*_I -  64*_Q + 50) / 100;
    B = (100*_Y - 111*_I + 170*_Q + 50) / 100;

    // Needed because of round-off errors
    if(R < 0) R = 0; else if(R > 255) R = 255;
    if(G < 0) G = 0; else if(G > 255) G = 255;
    if(B < 0) B = 0; else if(B > 255) B = 255;

    *pc = RGB(R, G, B);
}

// Function: Luminance
//
// Parameter: c -- The color whose luminance is to be returned
//
// Returns: The luminance value [0,255]
static inline int Luminance(COLORREF c)
{
    return CYIQ(c)._Y;
}

// Function: MoveColorBy
//
// Parameters: pcColor [i,o] The color to be moved
//             nLums         Move so that new color is this bright
//                           or darker than the original: a signed
//                           number whose abs value is betn 0 and 255
//
// Returns: Nothing
static void MoveColorBy(COLORREF* pcColor, int nLums)
{
    CYIQ yiq(*pcColor);
    int Y;

    Y = yiq._Y;

    if(Y+nLums > CYIQ::MAX_LUM)
    {
        // Cannot move more in the lighter direction.
        *pcColor = RGB(255, 255, 255);
    }
    else if(Y+nLums < CYIQ::MIN_LUM)
    {
        // Cannot move more in the darker direction.
        *pcColor = RGB(0, 0, 0);
    }
    else
    {
        Y += nLums;
        if(Y < 0)
        {
            Y = 0;
        }
        if(Y > 255)
        {
            Y = 255;
        }

        yiq._Y = Y;
        yiq.YIQ2RGB(pcColor);
    }
}

// Function: ContrastColors
//
// Parameters: c1,c2 [i,o]: Colors to be contrasted
//             Luminance:   Contrast so that diff is luminance is atleast
//                          this much: A number in the range [0,255]
//
// IMPORTANT: If you change this function, make sure, not to change the order
//            of the colors because some callers depend it.  For example if
//            both colors are white, we need to guarantee that only the
//            first color (c1) is darkens and never the second (c2).
//
// Returns: Nothing
void ContrastColors(COLORREF& c1, COLORREF& c2, int LumDiff)
{
    COLORREF *pcLighter, *pcDarker;
    int l1, l2, lLighter, lDarker;
    int lDiff, lPullApartBy;

    Assert((LumDiff>=CYIQ::MIN_LUM) && (LumDiff<=CYIQ::MAX_LUM));

    l1 = Luminance(c1);
    l2 = Luminance(c2);

    // If both the colors are black, make one slightly bright so
    // things work OK below ...
    if((l1==0) && (l2==0))
    {
        c1 = RGB(1, 1, 1);
        l1 = Luminance(c1);
    }

    // Get their absolute difference
    lDiff = l1<l2 ? l2-l1 : l1-l2;

    // Are they different enuf? If yes get out
    if(lDiff >= LumDiff)
    {
        return;
    }

    // Attention:  Don't change the order of the two colors as some callers
    // depend on this order. In case both colors are the same they need
    // to know which color is made darker and which is made lighter.
    if(l1 > l2)
    {
        // c1 is lighter than c2
        pcLighter = &c1;
        pcDarker = &c2;
        lDarker = l2;
    }
    else
    {
        // c1 is darker than c2
        pcLighter = &c2;
        pcDarker = &c1;
        lDarker = l1;
    }

    // STEP 1: Move lighter color
    //
    // Each color needs to be pulled apart by this much
    lPullApartBy = (LumDiff - lDiff + 1) / 2;
    Assert(lPullApartBy > 0);
    // First pull apart the lighter color -- in +ve direction
    MoveColorBy(pcLighter, lPullApartBy);

    // STEP 2: Move darker color
    //
    // Need to move darker color in the darker direction.
    // Note: Since the lighter color may not have moved enuf
    // we compute the distance the darker color should move
    // by recomuting the luminance of the lighter color and
    // using that to move the darker color.
    lLighter = Luminance(*pcLighter);
    lPullApartBy = lLighter - LumDiff - lDarker;
    // Be sure that we are moving in the darker direction
    Assert(lPullApartBy <= 0);
    MoveColorBy(pcDarker, lPullApartBy);

    // STEP 3: If necessary, move lighter color again
    //
    // Now did we have enuf space to move in darker region, if not,
    // then move in the lighter region again
    lPullApartBy = Luminance(*pcDarker) + LumDiff - lLighter;
    if(lPullApartBy > 0)
    {
        MoveColorBy(pcLighter, lPullApartBy);
    } 

    return;
}