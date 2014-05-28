#ifndef __textarea_h__
#define __textarea_h__

/* Forward Declarations */

struct ENUMDESC;

#ifndef __IHTMLTextAreaElement_FWD_DEFINED__
#define __IHTMLTextAreaElement_FWD_DEFINED__
typedef interface IHTMLTextAreaElement IHTMLTextAreaElement;
#endif     /* __IHTMLTextAreaElement_FWD_DEFINED__ */

#ifndef __IHTMLTextAreaElement_INTERFACE_DEFINED__

#define __IHTMLTextAreaElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLTextAreaElement;


MIDL_INTERFACE("3050f2aa-98b5-11cf-bb82-00aa00bdce0b")
IHTMLTextAreaElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE get_type(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_value(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_value(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_name(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_name(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_status(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_status(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_disabled(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_disabled(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_defaultValue(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_defaultValue(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE select(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onchange(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onchange(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onselect(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onselect(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_readOnly(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_readOnly(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_rows(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_rows(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_cols(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_cols(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_wrap(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_wrap(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLTextAreaElement_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLTextAreaElement;



EXTERN_C const GUID DIID_DispHTMLTextAreaElement;


#ifndef _CTextArea_PROPDESCS_

#endif


EXTERN_C const GUID GUID_HTMLRichtextElement;



EXTERN_C const GUID DIID_DispHTMLRichtextElement;


#ifndef _CRichtext_PROPDESCS_
EXTERN_C const PROPERTYDESC_CSTR_GETSET s_propdescCRichtextvalue;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCRichtexttype;
EXTERN_C const PROPERTYDESC_CSTR_GETSET s_propdescCRichtextname;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCRichtextstatus;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCRichtextdefaultValue;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCRichtextonchange;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCRichtextonselect;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCRichtextrows;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCRichtextcols;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCRichtextwrap;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCRichtextreadOnly;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCRichtextselect;

#endif


#endif /*__textarea_h__*/

