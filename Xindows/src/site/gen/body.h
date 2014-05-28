#ifndef __body_h__
#define __body_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "txtedit.h"

#ifndef __IHTMLBodyElement_FWD_DEFINED__
#define __IHTMLBodyElement_FWD_DEFINED__
typedef interface IHTMLBodyElement IHTMLBodyElement;
#endif     /* __IHTMLBodyElement_FWD_DEFINED__ */

typedef enum _bodyScroll
{
    bodyScrollyes = 1,
    bodyScrollno = 2,
    bodyScrollauto = 4,
    bodyScrolldefault = 3,
    bodyScroll_Max = 2147483647L
} bodyScroll;


EXTERN_C const ENUMDESC s_enumdescbodyScroll;


#ifndef __IHTMLBodyElement_INTERFACE_DEFINED__

#define __IHTMLBodyElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLBodyElement;


MIDL_INTERFACE("3050f1d8-98b5-11cf-bb82-00aa00bdce0b")
IHTMLBodyElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_background(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_background(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_bgProperties(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_bgProperties(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_leftMargin(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_leftMargin(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_topMargin(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_topMargin(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_rightMargin(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_rightMargin(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_bottomMargin(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_bottomMargin(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_noWrap(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_noWrap(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_bgColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_bgColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_text(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_text(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_link(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_link(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_vLink(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_vLink(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_aLink(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_aLink(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onload(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onload(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onunload(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onunload(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_scroll(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_scroll(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onselect(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onselect(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforeunload(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforeunload(
         /* [out] */ VARIANT * p) = 0;

};

#endif     /* __IHTMLBodyElement_INTERFACE_DEFINED__ */


EXTERN_C const GUID GUID_HTMLBody;



EXTERN_C const GUID DIID_DispHTMLBody;


#ifndef _CBodyElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementbgColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementbackground;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementbgProperties;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementtopMargin;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementleftMargin;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementrightMargin;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementbottomMargin;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementnoWrap;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementtext;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementlink;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementaLink;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementvLink;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementonload;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementonunload;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBodyElementscroll;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementonselect;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCBodyElementonbeforeunload;

#endif


#endif /*__body_h__*/

