
#include "stdafx.h"
#include "Download.h"

#include "Image.h"
#include "Dither.h"

#define whSlop  1       // amount of difference between element width or height and image
                        // width or height within which we will blt the image in its native
                        // proportions. For now, this applies only to GIF images.

int         g_colorModeDefault = 0;

WORD        g_wIdxTrans;
WORD        g_wIdxBgColor;
WORD        g_wIdxFgColor;
COLORREF    g_crBgColor;
COLORREF    g_crFgColor;
RGBQUAD     g_rgbBgColor;
RGBQUAD     g_rgbFgColor;

PALETTEENTRY g_peVga[16] =
{
    { 0x00, 0x00, 0x00, 0x00 }, // Black
    { 0x80, 0x00, 0x00, 0x00 }, // Dark red
    { 0x00, 0x80, 0x00, 0x00 }, // Dark green
    { 0x80, 0x80, 0x00, 0x00 }, // Dark yellow
    { 0x00, 0x00, 0x80, 0x00 }, // Dark blue
    { 0x80, 0x00, 0x80, 0x00 }, // Dark purple
    { 0x00, 0x80, 0x80, 0x00 }, // Dark aqua
    { 0xC0, 0xC0, 0xC0, 0x00 }, // Light grey
    { 0x80, 0x80, 0x80, 0x00 }, // Dark grey
    { 0xFF, 0x00, 0x00, 0x00 }, // Light red
    { 0x00, 0xFF, 0x00, 0x00 }, // Light green
    { 0xFF, 0xFF, 0x00, 0x00 }, // Light yellow
    { 0x00, 0x00, 0xFF, 0x00 }, // Light blue
    { 0xFF, 0x00, 0xFF, 0x00 }, // Light purple
    { 0x00, 0xFF, 0xFF, 0x00 }, // Light aqua
    { 0xFF, 0xFF, 0xFF, 0x00 }  // White
};

#define MASK565_0   0x0000F800
#define MASK565_1   0x000007E0
#define MASK565_2   0x0000001F


void* pCreateDitherData(int xsize)
{
    UINT cdw = (xsize + 1);
    DWORD* pdw = (DWORD*)MemAlloc(cdw*sizeof(DWORD));

    if(pdw)
    {
        pdw += cdw;
        while(cdw-- > 0) *--pdw = 0x909090;
    }

    return pdw;
}

// Mask off the high byte for comparing PALETTEENTRIES, RGBQUADS, etc.
#define RGBMASK(pe) (*((DWORD*)&(pe))&0x00FFFFFF)

int x_ComputeConstrainMap(int cEntries, PALETTEENTRY* pcolors, int transparent, int* pmapconstrained)
{
    int i;
    int nDifferent = 0;

    for(i=0; i<cEntries; i++)
    {
        if(i != transparent)
        {
            pmapconstrained[i] = RGB2Index(pcolors[i].peRed, pcolors[i].peGreen, pcolors[i].peBlue);

            if(RGBMASK(pcolors[i]) != RGBMASK(g_lpHalftone.ape[pmapconstrained[i]]))
            {
                ++nDifferent;
            }
        }
    }

    // Turns out the transparent index can be outside the color set.  In this
    // case we still want to map the transparent index correctly.
    if(transparent>=0 && transparent<=255)
    {
        pmapconstrained[transparent] = g_wIdxTrans;
    }

    return nDifferent;
}

/*
constrains colors to 6X6X6 cube we use
*/
void x_ColorConstrain(unsigned char HUGEP* psrc, unsigned char HUGEP* pdst, int* pmapconstrained, long xsize)
{
    int x;

    for(x=0; x<xsize; x++)
    {
        *pdst++ = (BYTE)pmapconstrained[*psrc++];
    }
}

void x_DitherRelative(BYTE* pbSrc, BYTE* pbDst, PALETTEENTRY* pe,
                      int xsize, int ysize, int transparent, int* v_rgb_mem,
                      int yfirst, int ylast)
{
    RGBQUAD argb[256];
    int cbScan;

    cbScan = (xsize + 3) & ~3;
    pbSrc  = pbSrc + cbScan * (ysize - yfirst - 1);
    pbDst  = pbDst + cbScan * (ysize - yfirst - 1);

    CopyColorsFromPaletteEntries(argb, pe, 256);

    DitherTo8( pbDst, -cbScan, 
        pbSrc, -cbScan, BFID_RGB_8, 
        g_rgbHalftone, argb,
        g_pInvCMAP,
        0, yfirst, xsize, ylast - yfirst + 1,
        g_wIdxTrans, transparent);
}

HRESULT x_Dither(unsigned char* pdata, PALETTEENTRY* pe, int xsize, int ysize, int transparent)
{
    x_DitherRelative(pdata, pdata, pe, xsize, ysize, transparent, NULL, 
        0, ysize-1);

    return S_OK;
}

// This function differentiates between "555" and "565" 16bpp color modes, returning 15 and 16, resp.
int GetRealColorMode(HDC hdc)
{
    struct
    {
        BITMAPINFOHEADER bmih;
        union
        {
            RGBQUAD argb[256];
            DWORD   dwMasks[3];
        } u;
    } bmi;
    HBITMAP hbm;

    hbm = CreateCompatibleBitmap(hdc, 1, 1);
    if(hbm == NULL)
    {
        return 0;
    }

    // NOTE: The two calls to GetDIBits are INTENTIONAL.  Don't muck with this!
    bmi.bmih.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmih.biBitCount = 0;
    GetDIBits(hdc, hbm, 0, 1, NULL, (BITMAPINFO*)&bmi, DIB_RGB_COLORS);
    GetDIBits(hdc, hbm, 0, 1, NULL, (BITMAPINFO*)&bmi, DIB_RGB_COLORS);

    DeleteObject(hbm);

    if(bmi.bmih.biBitCount != 16)
    {
        return bmi.bmih.biBitCount;
    }

    if(bmi.bmih.biCompression != BI_BITFIELDS)
    {
        return 15;
    }

    if(bmi.u.dwMasks[0]==MASK565_0 
        && bmi.u.dwMasks[1]==MASK565_1 
        && bmi.u.dwMasks[2]==MASK565_2)
    {
        return 16;
    }
    else
    {
        return 15;
    }
}

static int CALLBACK VgaPenCallback(void* pvLogPen, LPARAM lParam)
{
    LOGPEN* pLogPen = (LOGPEN*)pvLogPen;

    if(pLogPen->lopnStyle == PS_SOLID)
    {
        PALETTEENTRY** pppe = (PALETTEENTRY**)lParam;
        PALETTEENTRY* ppe = (*pppe)++;
        COLORREF cr = pLogPen->lopnColor;

        if(cr != *(DWORD*)ppe)
        {
            Trace2("Updating VGA color %d to %08lX", ppe-g_peVga, cr);
            *(DWORD*)ppe = cr;
        }

        return (ppe < &g_peVga[15]);
    }

    return 1;
}

BOOL InitImageUtil()
{
    /* Snoop around system and determine its capabilities.
    * Initialize data structures accordingly.
    */
    HDC hdc;

    HPALETTE hPal;

    hdc = GetDC(NULL);

    g_colorModeDefault = GetRealColorMode(hdc);

    g_crBgColor = GetSysColorQuick(COLOR_WINDOW) & 0xFFFFFF;
    g_crFgColor = GetSysColorQuick(COLOR_WINDOWTEXT) & 0xFFFFFF;

    g_rgbBgColor.rgbRed   = GetRValue(g_crBgColor);
    g_rgbBgColor.rgbGreen = GetGValue(g_crBgColor);
    g_rgbBgColor.rgbBlue  = GetBValue(g_crBgColor);
    g_rgbFgColor.rgbRed   = GetRValue(g_crFgColor);
    g_rgbFgColor.rgbGreen = GetGValue(g_crFgColor);
    g_rgbFgColor.rgbBlue  = GetBValue(g_crFgColor);

    if (GetDeviceCaps(hdc, RASTERCAPS) & RC_PALETTE)
    {
        hPal = g_hpalHalftone;

        /*
        Now the extra colors
        */
        g_wIdxBgColor = (WORD)GetNearestPaletteIndex(hPal, g_crBgColor);
        g_wIdxFgColor = (WORD)GetNearestPaletteIndex(hPal, g_crFgColor);

        /*
        Choose a transparent color that lies outside of 6x6x6 cube - we will
        replace the actual color for this right before drawing.  We are going
        to use one of the "magic colors" in the static 20 for the transparent
        color. 
        */
        g_wIdxTrans = 246;  
    }

    if(g_colorModeDefault == 4)
    {
        PALETTEENTRY* ppe = g_peVga;
        EnumObjects(hdc, OBJ_PEN, VgaPenCallback, (LONG_PTR)&ppe);
        Assert(ppe == &g_peVga[16]);
    }

    ReleaseDC(NULL, hdc);

    return TRUE;
}

int GetDefaultColorMode()
{
    if(g_colorModeDefault == 0)
    {
        InitImageUtil();
    }

    return g_colorModeDefault;
}

void FreeGifAnimData(GIFANIMDATA* pgad, CImgBitsDIB* pibd)
{
    GIFFRAME *pgf, *pgfNext;

    if(pgad == NULL)
    {
        return;
    }

    for(pgf=pgad->pgf; pgf!=NULL; pgf=pgfNext)
    {
        if(pgf->pibd != pibd)
        {
            delete pgf->pibd;
        }
        if(pgf->hrgnVis)
        {
            Verify(DeleteRgn(pgf->hrgnVis));
        }
        pgfNext = pgf->pgfNext;
        MemFree(pgf);
    }
    pgad->pgf = NULL;
}

void CalcStretchRect(RECT* prectStretch, LONG wImage, LONG hImage, LONG wDisplayedImage, LONG hDisplayedImage, GIFFRAME* pgf)
{
    // set ourselves up for a stretch if the element width doesn't match that of the image
    if((wDisplayedImage>=pgf->width-whSlop) &&
        (wDisplayedImage<=pgf->width+whSlop))
    {
        wDisplayedImage = pgf->width;
    }

    if((hDisplayedImage>=pgf->height-whSlop) &&
        (hDisplayedImage<=pgf->height+whSlop))
    {
        hDisplayedImage = pgf->height;
    }

    if(wImage != 0)
    {
        prectStretch->left = MulDivQuick(pgf->left, wDisplayedImage, wImage);
        prectStretch->right = prectStretch->left +
            MulDivQuick(pgf->width, wDisplayedImage, wImage);
    }
    else
    {
        prectStretch->left = prectStretch->right = pgf->left;
    }

    if(hImage != 0)
    {
        prectStretch->top = MulDivQuick(pgf->top, hDisplayedImage, hImage);
        prectStretch->bottom = prectStretch->top +
            MulDivQuick(pgf->height, hDisplayedImage, hImage);
    }
    else
    {
        prectStretch->top = prectStretch->bottom = pgf->top;
    }
}

void getPassInfo(int logicalRowX, int height, int* pPassX, int* pRowX, int* pBandX)
{
    int passLow, passHigh, passBand;
    int pass = 0;
    int step = 8;
    int ypos = 0;

    if(logicalRowX >= height)
    {
        logicalRowX = height - 1;
    }
    passBand = 8;
    passLow = 0;
    while(step > 1)
    {
        if(pass == 3)
        {
            passHigh = height - 1;
        }
        else
        {
            passHigh = (height-1-ypos)/step + passLow;
        }
        if(logicalRowX>=passLow && logicalRowX<=passHigh)
        {
            *pPassX = pass;
            *pRowX = ypos + (logicalRowX - passLow) * step;
            *pBandX = passBand;
            return;
        }
        if(pass++ > 0)
        {
            step /= 2;
        }
        ypos = step / 2;
        passBand /= 2;
        passLow = passHigh + 1;
    }
}

CImgBits* GetPlaceHolderBitmap(BOOL fMissing)
{
    CImgBits** ppImgBits;
    CImgBitsDIB* pibd = NULL;
    HBITMAP hbm = NULL;

    ppImgBits = fMissing ? &g_pImgBitsMissing : &g_pImgBitsNotLoaded;

    if(*ppImgBits == NULL)
    {
        LOCK_GLOBALS;

        if(*ppImgBits == NULL)
        {
            hbm = (HBITMAP)LoadImage(
                _afxGlobalData._hInstResource,
                (LPCWSTR)(DWORD_PTR)(fMissing?IDB_MISSING:IDB_NOTLOADED),
                IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
            
            // NOTE (lmollico): bitmaps are in mshtml.dll
            pibd = new CImgBitsDIB();
            if(!pibd)
            {
                goto Cleanup;
            }

            Verify(!pibd->AllocCopyBitmap(hbm, FALSE, -1));

            Verify(DeleteObject(hbm));
            *ppImgBits = pibd;
            pibd = NULL;

Cleanup:
            if(pibd)
            {
                delete pibd;
            }
        }

    }

    return *ppImgBits;
}

void GetPlaceHolderBitmapSize(BOOL fMissing, SIZE* pSize)
{
    CImgBits* pib;

    pib = GetPlaceHolderBitmap(fMissing);

    pSize->cx = pib->Width();
    pSize->cy = pib->Height();
}

int DrawTextInCodePage(UINT uCP, HDC hDC, LPCWSTR lpString, int nCount, LPRECT lpRect, UINT uFormat)
{
    return DrawTextW(hDC, (LPWSTR)lpString, nCount, lpRect, uFormat);
}

//+---------------------------------------------------------------------------
//
//  DrawPlaceHolder
//
//  Synopsis:   Draw the Place holder, the ALT string and the bitmap
//
//----------------------------------------------------------------------------
void DrawPlaceHolder(
         HDC hdc, 
         RECT rectImg,
         TCHAR* lpString,
         CODEPAGE codepage,
         LCID lcid,
         SHORT sBaselineFont,
         SIZE* psizeGrab,
         BOOL fMissing,
         COLORREF fgColor,
         COLORREF bgColor,
         SIZE* psizePrint,
         BOOL fRTL,
         DWORD dwFlags)
{
    LONG     xDstWid;
    LONG     yDstHei;
    LONG     xSrcWid;
    LONG     ySrcHei;
    CImgBits*pib;
    RECT    rcDst;
    RECT    rcSrc; 
    BOOL    fDrawBlackRectForPrinting = psizePrint != NULL;

    bgColor &= 0x00FFFFFF;

    pib = GetPlaceHolderBitmap(fMissing);

    xDstWid = psizePrint ? psizePrint->cx : pib->Width();
    yDstHei = psizePrint ? psizePrint->cy : pib->Height();
    xSrcWid = pib->Width();
    ySrcHei = pib->Height();

    if((rectImg.right-rectImg.left>=10) && (rectImg.bottom-rectImg.top>=10))
    {
        if(fDrawBlackRectForPrinting)
        {
            COLORREF crBlack = 0; // zero is black
            HBRUSH   hBrush = 0;

            hBrush = CreateSolidBrush(crBlack);
            if(hBrush)
            {
                FrameRect(hdc, &rectImg, hBrush);
                DeleteObject(hBrush);
            }
            else
            {
                fDrawBlackRectForPrinting = FALSE;
            }
        }

        if(!fDrawBlackRectForPrinting)
        {
            if((bgColor==0x00ffffff) || (bgColor==0x00000000))
            {
                DrawEdge(hdc, &rectImg, BDR_SUNKENOUTER, BF_TOPLEFT);
                DrawEdge(hdc, &rectImg, BDR_SUNKENINNER, BF_BOTTOMRIGHT);
            }
            else
            {
                DrawEdge(hdc, &rectImg, BDR_SUNKENOUTER, BF_RECT);
            }
        }
    }

    if(lpString != NULL)
    {
        RECT rc;
        BOOL fGlyph = FALSE;
        UINT cch = _tcslen(lpString);

        rc.left = rectImg.left + xDstWid + 2*psizeGrab->cx;
        rc.right = rectImg.right - psizeGrab->cx;
        rc.top = rectImg.top + psizeGrab->cy;
        rc.bottom = rectImg.bottom - psizeGrab->cy;

        CIntlFont intlFont(hdc, codepage, lcid, sBaselineFont, lpString);
        SetTextColor(hdc, fgColor);

        if(!fRTL)
        {
            for(UINT i=0; i<cch; i++)
            {
                WCHAR ch = lpString[i];
                if(ch>=0x300 && IsGlyphableChar(ch))
                {
                    fGlyph = TRUE;
                    break;
                }
            }
        }

        // send complex text or text layed out right-to-left
        // to be drawn through Uniscribe
        if(fGlyph || fRTL)
        {
            HRESULT hr;
            UINT taOld = 0;
            UINT fuOptions = ETO_CLIPPED;

            if(fRTL)
            {
                taOld = GetTextAlign(hdc);
                SetTextAlign(hdc, TA_RTLREADING|TA_RIGHT);
                fuOptions |= ETO_RTLREADING;
            }

            extern HRESULT LSUniscribeTextOut(
                HDC hdc, 
                int iX, 
                int iY, 
                UINT uOptions, 
                CONST RECT* prc, 
                LPCTSTR pString, 
                UINT cch,
                int* piDx);

            hr = LSUniscribeTextOut(
                hdc,
                !fRTL?rc.left:rc.right, 
                rc.top,
                fuOptions,
                &rc,
                lpString,
                cch,
                NULL);

            if(fRTL)
            {
                SetTextAlign(hdc, taOld);
            }
        }
        else
        {
            DrawTextInCodePage(WindowsCodePageFromCodePage(codepage),
                hdc, lpString, -1, &rc, DT_LEFT|DT_WORDBREAK|DT_NOPREFIX);
        }
    }

    if(((rectImg.right-rectImg.left)<=2*psizeGrab->cx) ||
        ((rectImg.bottom-rectImg.top)<=2*psizeGrab->cy))
    {
        return;
    }

    InflateRect(&rectImg, -psizeGrab->cx, -psizeGrab->cy);

    if(xDstWid > rectImg.right-rectImg.left)
    {
        xSrcWid = MulDivQuick(xSrcWid, rectImg.right-rectImg.left, xDstWid);
        xDstWid = rectImg.right - rectImg.left;
    }
    if(yDstHei > rectImg.bottom-rectImg.top)
    {
        ySrcHei = MulDivQuick(ySrcHei, rectImg.bottom-rectImg.top, yDstHei);
        yDstHei = rectImg.bottom - rectImg.top;
    }

    rcDst.left = rectImg.left;
    rcDst.top = rectImg.top;
    rcDst.right = rcDst.left + xDstWid;
    rcDst.bottom = rcDst.top + yDstHei;

    rcSrc.left = 0;
    rcSrc.top = 0;
    rcSrc.right = xSrcWid;
    rcSrc.bottom = ySrcHei;

    pib->StretchBlt(hdc, &rcDst, &rcSrc, SRCCOPY, dwFlags);
}

int Union(int _yTop, int _yBottom, BOOL fInvalidateAll, int yBottom)
{
    if((_yTop!=-1)
        && (fInvalidateAll
        || ((_yBottom>=_yTop)
        && (yBottom>=_yTop)
        && (yBottom<=_yBottom))
        || ((_yBottom<_yTop)
        && ((yBottom>=_yTop)
        || (yBottom<=_yBottom)))))
    {
        return -1;
    }
    return _yTop;
}

void ComputeFrameVisibility(IMGANIMSTATE* pImgAnimState, LONG xWidth, LONG yHeight, LONG xDispWid, LONG yDispHei)
{
    GIFFRAME*   pgf;
    GIFFRAME*   pgfClip;
    GIFFRAME*   pgfDraw = pImgAnimState->pgfDraw;
    GIFFRAME*   pgfDrawNext = pgfDraw->pgfNext;
    RECT        rectCur;

    // determine which frames are visible or partially visible at this time
    for(pgf=pImgAnimState->pgfFirst; pgf!=pgfDrawNext; pgf=pgf->pgfNext)
    {
        if(pgf->hrgnVis != NULL)
        {
            DeleteRgn(pgf->hrgnVis);
            pgf->hrgnVis = NULL;
            pgf->bRgnKind = NULLREGION;
        }

        // This is kinda complicated.
        // We only want to subtract out this frame from its predecessors under certain
        // conditions.
        // If it's the current frame, then all we care about is transparency.
        // If its a preceding frame, then any bits from frames that preceded should
        // be clipped out if it wasn't transparent, but also wasn't to be replaced by
        // previous pixels.
        if(((pgf==pgfDraw) && !pgf->bTransFlags) ||
            ((pgf!=pgfDraw) && !pgf->bTransFlags && (pgf->bDisposalMethod!=gifRestorePrev)) ||
            ((pgf!=pgfDraw) && (pgf->bDisposalMethod==gifRestoreBkgnd)))
        {
            // clip this rgn out of those that came before us if it's not trasparent,
            // or if it leaves a background-colored hole and is not the current frame.
            // The current frame, being current, hasn't left a background-colored hole yet.
            for(pgfClip=pImgAnimState->pgfFirst; pgfClip!=pgf; pgfClip=pgfClip->pgfNext)
            {
                if(pgfClip->hrgnVis != NULL)
                {
                    if(pgf->hrgnVis == NULL)
                    {
                        // Since we'll use these regions to clip when drawing, we need them mapped
                        // for destination stretching.
                        CalcStretchRect(&rectCur, xWidth, yHeight, xDispWid, yDispHei, pgf);

                        pgf->hrgnVis = CreateRectRgnIndirect(&rectCur);
                        pgf->bRgnKind = SIMPLEREGION;
                    }

                    pgfClip->bRgnKind = (BYTE)CombineRgn(pgfClip->hrgnVis, pgfClip->hrgnVis, pgf->hrgnVis, RGN_DIFF);
                }
            }
        } // if we need to clip this frame out of its predecessor(s)

        // If this is a replace with background frame preceding the current draw frame,
        // then it is not visible at all, so set the visibility traits so it won't be drawn.
        if((pgf!=pgfDraw) && (pgf->bDisposalMethod>=gifRestoreBkgnd))
        {
            if(pgf->hrgnVis != NULL)
            {
                DeleteRgn(pgf->hrgnVis);
                pgf->hrgnVis = NULL;
                pgf->bRgnKind = NULLREGION;
            }
        }
        else if(pgf->hrgnVis == NULL)
        {
            // Since we'll use these regions to clip when drawing, we need them mapped
            // for destination stretching.
            CalcStretchRect(&rectCur, xWidth, yHeight, xDispWid, yDispHei, pgf);

            pgf->hrgnVis = CreateRectRgnIndirect(&rectCur);
            pgf->bRgnKind = SIMPLEREGION;
        }

    } // for check each frame's visibility
}




DYNLIB g_dynlibIMGUTIL = { NULL, NULL, "imgutil.dll" };

void APIENTRY InitImgUtil(void)
{
    static DYNPROC s_dynprocInitImgUtil = { NULL, &g_dynlibIMGUTIL, "InitImgUtil" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocInitImgUtil);
    if(hr)
    {
        return ;
    }

    (*(void (APIENTRY*)(void))s_dynprocInitImgUtil.pfn)();
}

STDAPI DecodeImage(IStream* pStream, IMapMIMEToCLSID* pMap, IUnknown* pEventSink)
{
    static DYNPROC s_dynprocDecodeImage = { NULL, &g_dynlibIMGUTIL, "DecodeImage" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocDecodeImage);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(IStream*, IMapMIMEToCLSID*, IUnknown*))s_dynprocDecodeImage.pfn)
        (pStream, pMap, pEventSink);
}

STDAPI CreateMIMEMap(IMapMIMEToCLSID** ppMap)
{
    static DYNPROC s_dynprocCreateMIMEMap = { NULL, &g_dynlibIMGUTIL, "CreateMIMEMap" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocCreateMIMEMap);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(IMapMIMEToCLSID**))s_dynprocCreateMIMEMap.pfn)
        (ppMap);
}

STDAPI GetMaxMIMEIDBytes(ULONG* pnMaxBytes)
{
    static DYNPROC s_dynprocGetMaxMIMEIDBytes = { NULL, &g_dynlibIMGUTIL, "GetMaxMIMEIDBytes" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocGetMaxMIMEIDBytes);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(ULONG*))s_dynprocGetMaxMIMEIDBytes.pfn)
        (pnMaxBytes);
}

STDAPI IdentifyMIMEType(const BYTE* pbBytes, ULONG nBytes, UINT* pnFormat)
{
    static DYNPROC s_dynprocIdentifyMIMEType = { NULL, &g_dynlibIMGUTIL, "IdentifyMIMEType" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocIdentifyMIMEType);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(const BYTE*, ULONG, UINT*))s_dynprocIdentifyMIMEType.pfn)
        (pbBytes, nBytes, pnFormat);
}

STDAPI SniffStream(IStream* pInStream, UINT* pnFormat, IStream** ppOutStream)
{
    static DYNPROC s_dynprocSniffStream = { NULL, &g_dynlibIMGUTIL, "SniffStream" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocSniffStream);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(IStream*, UINT*, IStream**))s_dynprocSniffStream.pfn)
        (pInStream, pnFormat, ppOutStream);
}

STDAPI ComputeInvCMAP(const RGBQUAD* pRGBColors, ULONG nColors, BYTE* pInvTable, ULONG cbTable)
{
    static DYNPROC s_dynprocComputeInvCMAP = { NULL, &g_dynlibIMGUTIL, "ComputeInvCMAP" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocComputeInvCMAP);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(const RGBQUAD*, ULONG, BYTE*, ULONG))s_dynprocComputeInvCMAP.pfn)
        (pRGBColors, nColors, pInvTable, cbTable);
}

STDAPI CreateDDrawSurfaceOnDIB(HBITMAP hbmDib, IDirectDrawSurface** ppSurface)
{
    static DYNPROC s_dynprocCreateDDrawSurfaceOnDIB = { NULL, &g_dynlibIMGUTIL, "CreateDDrawSurfaceOnDIB" };

    HRESULT hr;

    hr = LoadProcedure(&s_dynprocCreateDDrawSurfaceOnDIB);
    if(hr)
    {
        return NULL;
    }

    return (*(HRESULT (APIENTRY*)(HBITMAP, IDirectDrawSurface**))s_dynprocCreateDDrawSurfaceOnDIB.pfn)
        (hbmDib, ppSurface);
}