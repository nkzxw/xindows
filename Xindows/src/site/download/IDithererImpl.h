
#ifndef __XINDOWS_SITE_DOWNLOAD_IDITHERERIMPL_H__
#define __XINDOWS_SITE_DOWNLOAD_IDITHERERIMPL_H__

DEFINE_GUID(CLSID_CoDitherToRGB8,   0xA860CE50,0x3910,0x11d0,0x86,0xFC,0x00,0xA0,0xC9,0x13,0xF7,0x50);
DEFINE_GUID(IID_IDithererImpl,      0x7C48E840,0x3910,0x11D0,0x86,0xFC,0x00,0xA0,0xC9,0x13,0xF7,0x50);

interface IDithererImpl : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetDestColorTable(
        ULONG nColors, const RGBQUAD* prgbColors) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetEventSink(
        IImageDecodeEventSink* pEventSink) = 0;
};

#endif //__XINDOWS_SITE_DOWNLOAD_IDITHERERIMPL_H__