
#include "stdafx.h"
#include "Image.h"

typedef struct tagSRECT
{
    short    left;
    short    top;
    short    right;
    short    bottom;
} SRECT;

typedef struct
{
    DWORD   key;
    WORD    hmf;
    SRECT   bbox;
    WORD    inch;
    DWORD   reserved;
    WORD    checksum;
}ALDUSMFHEADER;

class CImgTaskWmf : public CImgTask
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    typedef CImgTask super;

    // CImgTask methods
    virtual void Decode(BOOL* pfNonProgressive);

    // CImgTaskWmf methods
    void ReadImage();
};

void CImgTaskWmf::ReadImage()
{
    METAHEADER      mh;
    LPBYTE          pbBuf = NULL;
    CImgBitsDIB*    pibd = NULL;
    HDC             hdcDst = NULL;
    HMETAFILE       hmf = NULL;
    HBITMAP         hbmSav = NULL;
    ULONG           ulSize;
    HRESULT         hr;
    RGBQUAD         argb[256];

    // Get the metafile header so we know how big it is
    if(!Read((unsigned char*)&mh, sizeof(mh)))
    {
        return;
    }

    // allocate a buffer to hold it
    ulSize = mh.mtSize * sizeof(WORD);
    pbBuf = (LPBYTE)MemAlloc(ulSize);
    if(!pbBuf)
    {
        return;
    }

    // copy the header into the front of the buffer
    memcpy(pbBuf, &mh, sizeof(METAHEADER));

    // read the metafile into memory after the header
    if(!Read(pbBuf+sizeof(METAHEADER), ulSize-sizeof(METAHEADER)))
    {
        goto Cleanup;
    }

    // convert the buffer into a metafile handle
    hmf = SetMetaFileBitsEx(ulSize, pbBuf);
    if(!hmf)
    {
        goto Cleanup;
    }

    // Free the metafile buffer
    MemFree(pbBuf);
    pbBuf = NULL;

    // Use the halftone palette for the color table
    CopyColorsFromPaletteEntries(argb, g_lpHalftone.ape, 256);
    memcpy(_ape, g_lpHalftone.ape, sizeof(PALETTEENTRY)*256);

    pibd = new CImgBitsDIB();
    if(!pibd)
    {
        goto Cleanup;
    }

    hr = pibd->AllocDIBSection(8, _xWid, _yHei, argb, 256, 255);
    if(hr)
    {
        goto Cleanup;
    }

    _lTrans = 255;

    Assert(pibd->GetBits() && pibd->GetHbm());

    memset(pibd->GetBits(), (BYTE)_lTrans, pibd->CbLine()*_yHei);

    // Render the metafile into the bitmap
    hdcDst = GetMemoryDC();
    if(!hdcDst)
    {
        goto Cleanup;
    }

    hbmSav = (HBITMAP)SelectObject(hdcDst, pibd->GetHbm());

    SaveDC(hdcDst);

    SetMapMode(hdcDst, MM_ANISOTROPIC);
    SetViewportExtEx(hdcDst, _xWid, _yHei, (SIZE*)NULL);
    PlayMetaFile(hdcDst, hmf);

    RestoreDC(hdcDst, -1);

    _pImgBits = pibd;
    pibd = NULL;

    _ySrcBot = -1;

Cleanup:
    if(hbmSav)
    {
        SelectObject(hdcDst, hbmSav);
    }
    if(hdcDst)
    {
        ReleaseMemoryDC(hdcDst);
    }
    if(pibd)
    {
        delete pibd;
    }
    if(hmf)
    {
        DeleteMetaFile(hmf);
    }
    if(pbBuf)
    {
        MemFree(pbBuf);
    }

    return;
}

void CImgTaskWmf::Decode(BOOL* pfNonProgressive)
{
    ALDUSMFHEADER amfh;

    *pfNonProgressive = TRUE;

    // KENSY: What do I do about the color table?
    // KENSY: scale according to DPI of screen
    if(!Read((unsigned char*)&amfh, sizeof(amfh)))
    {
        goto Cleanup;
    }

    _xWid = abs(MulDiv(amfh.bbox.right-amfh.bbox.left, 96, amfh.inch));
    _yHei = abs(MulDiv(amfh.bbox.bottom-amfh.bbox.top, 96, amfh.inch));

    // Post WHKNOWN
    OnSize(_xWid, _yHei, -1);

    ReadImage();

Cleanup:
    return;
}

CImgTask* NewImgTaskWmf()
{
    return new CImgTaskWmf;
}