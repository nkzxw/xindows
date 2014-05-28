#ifndef __img_h__
#define __img_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "csite.h"

#ifndef __IHTMLImgElement_FWD_DEFINED__
#define __IHTMLImgElement_FWD_DEFINED__
typedef interface IHTMLImgElement IHTMLImgElement;
#endif     /* __IHTMLImgElement_FWD_DEFINED__ */

#ifndef __IHTMLImageElementFactory_FWD_DEFINED__
#define __IHTMLImageElementFactory_FWD_DEFINED__
typedef interface IHTMLImageElementFactory IHTMLImageElementFactory;
#endif     /* __IHTMLImageElementFactory_FWD_DEFINED__ */

#ifndef __IHTMLImgElement_INTERFACE_DEFINED__

#define __IHTMLImgElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLImgElement;


MIDL_INTERFACE("3050f240-98b5-11cf-bb82-00aa00bdce0b")
IHTMLImgElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_isMap(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_isMap(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_mimeType(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fileSize(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fileCreatedDate(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fileModifiedDate(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fileUpdatedDate(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_protocol(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_href(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_nameProp(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_border(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_border(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_vspace(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_vspace(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_hspace(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_hspace(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_alt(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_alt(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_src(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_src(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_lowsrc(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_lowsrc(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_dynsrc(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_dynsrc(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_readyState(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_complete(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_loop(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_loop(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_align(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_align(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onload(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onload(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onerror(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onerror(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onabort(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onabort(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_name(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_name(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_width(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_width(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_height(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_height(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_start(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_start(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLImgElement_INTERFACE_DEFINED__ */


#ifndef __IHTMLImageElementFactory_INTERFACE_DEFINED__

#define __IHTMLImageElementFactory_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLImageElementFactory;


MIDL_INTERFACE("3050f38e-98b5-11cf-bb82-00aa00bdce0b")
IHTMLImageElementFactory : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE create(
            /* [in] */ VARIANT width,/* [in] */ VARIANT height,/* [out] */ IHTMLImgElement** ) = 0;

};

#endif     /* __IHTMLImageElementFactory_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLImg;



EXTERN_C const GUID DIID_DispHTMLImg;


#ifndef _CImgElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementalign;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementalt;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementsrc;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementheight;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementwidth;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementborder;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementvspace;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementhspace;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementlowsrc;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementdynsrc;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementreadyState;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementcomplete;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCImgElementloop;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementstart;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementonload;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementonerror;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCImgElementonabort;
EXTERN_C const PROPERTYDESC_CSTR_GETSET s_propdescCImgElementname;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementisMap;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementmimeType;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementfileSize;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementfileCreatedDate;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementfileModifiedDate;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementfileUpdatedDate;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementprotocol;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementhref;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCImgElementnameProp;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCImgElementcache;

#endif


EXTERN_C const GUID GUID_HTMLImageElementFactory;


#ifndef _CImageElementFactory_PROPDESCS_
EXTERN_C const PROPERTYDESC_METHOD s_methdescCImageElementFactorycreate;

#endif


#endif /*__img_h__*/

