#ifndef __csite_h__
#define __csite_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

#ifndef __IHTMLControlElement_FWD_DEFINED__
#define __IHTMLControlElement_FWD_DEFINED__
typedef interface IHTMLControlElement IHTMLControlElement;
#endif     /* __IHTMLControlElement_FWD_DEFINED__ */
typedef enum _htmlStart
{
    htmlStartfileopen = 0,
    htmlStartmouseover = 1,
    htmlStart_Max = 2147483647L
} htmlStart;


EXTERN_C const ENUMDESC s_enumdeschtmlStart;


#ifndef __IHTMLControlElement_INTERFACE_DEFINED__

#define __IHTMLControlElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLControlElement;


MIDL_INTERFACE("3050f4e9-98b5-11cf-bb82-00aa00bdce0b")
IHTMLControlElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_tabIndex(
         /* [in] */ short v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_tabIndex(
         /* [out] */ short * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE focus(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_accessKey(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_accessKey(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onblur(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onblur(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onfocus(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onfocus(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onresize(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onresize(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE blur(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clientHeight(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clientWidth(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clientTop(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clientLeft(
         /* [out] */ long * p) = 0;

};

#endif     /* __IHTMLControlElement_INTERFACE_DEFINED__ */


#ifndef _CSite_PROPDESCS_
EXTERN_C const PROPERTYDESC_METHOD s_methdescCSitefocus;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCSiteblur;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCSiteonblur;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCSiteonfocus;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCSiteaccessKey;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCSiteonresize;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCSiteclientHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCSiteclientWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCSiteclientTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCSiteclientLeft;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCSitetabIndex;

#endif


#endif /*__csite_h__*/

