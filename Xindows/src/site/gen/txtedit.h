#ifndef __txtedit_h__
#define __txtedit_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "csite.h"

#ifndef __IHTMLTextContainer_FWD_DEFINED__
#define __IHTMLTextContainer_FWD_DEFINED__
typedef interface IHTMLTextContainer IHTMLTextContainer;
#endif     /* __IHTMLTextContainer_FWD_DEFINED__ */

#ifndef __IHTMLTextContainer_INTERFACE_DEFINED__

#define __IHTMLTextContainer_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLTextContainer;


MIDL_INTERFACE("3050f230-98b5-11cf-bb82-00aa00bdce0b")
IHTMLTextContainer : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE get_scrollHeight(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_scrollWidth(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_scrollTop(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_scrollTop(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_scrollLeft(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_scrollLeft(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onscroll(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onscroll(
         /* [out] */ VARIANT * p) = 0;

};

#endif     /* __IHTMLTextContainer_INTERFACE_DEFINED__ */


#ifndef _CTxtSite_PROPDESCS_
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCTxtSitescrollHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCTxtSitescrollWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCTxtSitescrollTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCTxtSitescrollLeft;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCTxtSiteonscroll;

#endif


#endif /*__txtedit_h__*/

