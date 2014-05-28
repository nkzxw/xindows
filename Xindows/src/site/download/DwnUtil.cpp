
#include "stdafx.h"
#include "Download.h"

#include "Image.h"

// MIMEINFO -------------------------------------------------------------------
NEWIMGTASKFN NewImgTaskGif;
NEWIMGTASKFN NewImgTaskJpg;
NEWIMGTASKFN NewImgTaskBmp;
NEWIMGTASKFN NewImgTaskArt;
NEWIMGTASKFN NewImgTaskXbm;
NEWIMGTASKFN NewImgTaskWmf;
NEWIMGTASKFN NewImgTaskEmf;
NEWIMGTASKFN NewImgTaskPlug;
NEWIMGTASKFN NewImgTaskIco;

MIMEINFO g_rgMimeInfo[] =
{
    { 0, CFSTR_MIME_HTML,       0,               0 },
    { 0, CFSTR_MIME_TEXT,       0,               0 },
    { 0, TEXT("text/x-component"),0,             0 },
    { 0, CFSTR_MIME_GIF,        NewImgTaskGif,   0/*IDS_SAVEPICTUREAS_GIF wlw note*/ },
    { 0, CFSTR_MIME_JPEG,       NewImgTaskJpg,   0/*IDS_SAVEPICTUREAS_JPG wlw note*/ },
    { 0, CFSTR_MIME_PJPEG,      NewImgTaskJpg,   0/*IDS_SAVEPICTUREAS_JPG wlw note*/ },
    { 0, CFSTR_MIME_BMP,        NewImgTaskBmp,   0/*IDS_SAVEPICTUREAS_BMP wlw note*/ },
    { 0, CFSTR_MIME_X_ART,      NewImgTaskArt,   0/*IDS_SAVEPICTUREAS_ART wlw note*/ },
    { 0, TEXT("image/x-art"),   NewImgTaskArt,   0/*IDS_SAVEPICTUREAS_ART wlw note*/ },
    { 0, CFSTR_MIME_XBM,        NewImgTaskXbm,   0/*IDS_SAVEPICTUREAS_XBM wlw note*/ },
    { 0, CFSTR_MIME_X_BITMAP,   NewImgTaskXbm,   0/*IDS_SAVEPICTUREAS_XBM wlw note*/ },
    { 0, CFSTR_MIME_X_WMF,      NewImgTaskWmf,   0/*IDS_SAVEPICTUREAS_WMF wlw note*/ },
    { 0, CFSTR_MIME_X_EMF,      NewImgTaskEmf,   0/*IDS_SAVEPICTUREAS_EMF wlw note*/ },
    { 0, CFSTR_MIME_AVI,        0,               0/*IDS_SAVEPICTUREAS_AVI wlw note*/ },
    { 0, CFSTR_MIME_X_MSVIDEO,  0,               0/*IDS_SAVEPICTUREAS_AVI wlw note*/ },
    { 0, CFSTR_MIME_MPEG,       0,               0/*IDS_SAVEPICTUREAS_MPG wlw note*/ },
    { 0, CFSTR_MIME_QUICKTIME,  0,               0/*IDS_SAVEPICTUREAS_MOV wlw note*/ },
    { 0, CFSTR_MIME_X_PNG,      NewImgTaskPlug,  0/*IDS_SAVEPICTUREAS_PNG wlw note*/ },
    { 0, TEXT("image/png"),     NewImgTaskPlug,  0/*IDS_SAVEPICTUREAS_PNG wlw note*/ },
    { 0, TEXT("image/x-icon"),  NewImgTaskIco,   0/*IDS_SAVEPICTUREAS_BMP wlw note*/ },
};

LPCTSTR g_pchWebviewMimeWorkaround = _T("text/webviewhtml");

#define MIME_TYPE_COUNT     ARRAYSIZE(g_rgMimeInfo)

MIMEINFO g_miImagePlug =
{
    0, _T("image/x-ms-plug"), NewImgTaskPlug,  0
};

const LPCSTR g_rgpchMimeType[MIME_TYPE_COUNT] =
{
    "text/html",
    "text/plain",
    "text/x-component",
    "image/gif",
    "image/jpeg",
    "image/pjpeg",
    "image/bmp",
    "image/x-jg",
    "image/x-art",
    "image/xbm",
    "image/x-xbitmap",
    "image/x-wmf",
    "image/x-emf",
    "video/avi",
    "video/x-msvideo",
    "video/mpeg",
    "video/quicktime",
    "image/x-png",
    "image/png",
    "image/x-icon",
};

MIMEINFO* g_pmiTextHtml         = &g_rgMimeInfo[0];
MIMEINFO* g_pmiTextPlain        = &g_rgMimeInfo[1];
MIMEINFO* g_pmiTextComponent    = &g_rgMimeInfo[2];

MIMEINFO* g_pmiImagePlug        = &g_miImagePlug;

BOOL g_fInitMimeInfo = FALSE;

void InitMimeInfo()
{
    MIMEINFO*       pmi  = g_rgMimeInfo;
    const LPCSTR*   ppch = g_rgpchMimeType;
    int             c    = MIME_TYPE_COUNT;

    Assert(ARRAYSIZE(g_rgMimeInfo) == ARRAYSIZE(g_rgpchMimeType));

    for(; --c>=0; ++pmi,++ppch)
    {
        Assert(*ppch);
        pmi->cf = (CLIPFORMAT)RegisterClipboardFormatA(*ppch);
    }

    g_fInitMimeInfo = TRUE;
}

MIMEINFO* GetMimeInfoFromClipFormat(CLIPFORMAT cf)
{
    MIMEINFO* pmi;
    UINT c;

    if(!g_fInitMimeInfo)
    {
        InitMimeInfo();
    }

    for(c=MIME_TYPE_COUNT,pmi=g_rgMimeInfo; c>0; --c,++pmi)
    {
        if(pmi->cf == cf)
        {
            return pmi;
        }
    }

    return NULL;
}

MIMEINFO* GetMimeInfoFromMimeType(const TCHAR* pchMime)
{
    MIMEINFO* pmi;
    UINT c;

    if(!g_fInitMimeInfo)
    {
        InitMimeInfo();
    }

    for(c=MIME_TYPE_COUNT,pmi=g_rgMimeInfo; c>0; --c,++pmi)
    {
        if(StrCmpIC(pmi->pch, pchMime) == 0)
        {
            return pmi;
        }
    }

    // BUGBUG: the following works around urlmon bug NT 175191:
    // Mime filters do not change the mime type correctly.
    // Remove this exception when that bug is fixed.
    if(StrCmpIC(g_pchWebviewMimeWorkaround, pchMime) == 0)
    {
        return g_pmiTextHtml;
    }

    return NULL;
}

MIMEINFO* GetMimeInfoFromData(void* pvData, ULONG cbData, const TCHAR* pchProposed)
{
    HRESULT     hr;
    MIMEINFO*   pmi = NULL;
    TCHAR*      pchMimeType = NULL;

    hr = FindMimeFromData(
        NULL,             // bind context - can be NULL                                     
        NULL,             // url - can be null
        pvData,           // buffer with data to sniff - can be null (pwzUrl must be valid) 
        cbData,           // size of buffer                                                 
        pchProposed,      // proposed mime if - can be null                                 
        0,                // will be defined                                                
        &pchMimeType,     // the suggested mime                                             
        0);
    if(!hr)
    {
        pmi = GetMimeInfoFromMimeType(pchMimeType);
        if(pmi)
        {
            goto Cleanup;
        }
    }

    if(cbData && IsPluginImgFormat((BYTE*)pvData, cbData))
    {
        pmi = g_pmiImagePlug;
    }

Cleanup:
    CoTaskMemFree(pchMimeType);
    return(pmi);
}

// Shutdown -------------------------------------------------------------------
void DeinitDownload()
{
    if(g_pImgBitsNotLoaded)
    {
        delete g_pImgBitsNotLoaded;
    }

    if(g_pImgBitsMissing)
    {
        delete g_pImgBitsMissing;
    }

    DwnCacheDeinit();
}
