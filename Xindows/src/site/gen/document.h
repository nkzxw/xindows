#ifndef __document_h__
#define __document_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "types.h"

/* header files for imported files */
#include "body.h"

/* header files for imported files */
#include "collect.h"

#ifndef __IHTMLDocument_FWD_DEFINED__
#define __IHTMLDocument_FWD_DEFINED__
typedef interface IHTMLDocument IHTMLDocument;
#endif     /* __IHTMLDocument_FWD_DEFINED__ */

#ifndef __IHTMLDocument2_FWD_DEFINED__
#define __IHTMLDocument2_FWD_DEFINED__
typedef interface IHTMLDocument2 IHTMLDocument2;
#endif     /* __IHTMLDocument2_FWD_DEFINED__ */

#ifndef __IHTMLDocument3_FWD_DEFINED__
#define __IHTMLDocument3_FWD_DEFINED__
typedef interface IHTMLDocument3 IHTMLDocument3;
#endif     /* __IHTMLDocument3_FWD_DEFINED__ */

#ifndef __IHTMLDocument_INTERFACE_DEFINED__

#define __IHTMLDocument_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLDocument;


MIDL_INTERFACE("626FC520-A41E-11cf-A731-00A0C9082637")
IHTMLDocument : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE get_Script(
         /* [out] */ IDispatch* * p) = 0;

};

#endif     /* __IHTMLDocument_INTERFACE_DEFINED__ */


#ifndef __IHTMLDocument2_INTERFACE_DEFINED__

#define __IHTMLDocument2_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLDocument2;


MIDL_INTERFACE("332c4425-26cb-11d0-b483-00c04fd90119")
IHTMLDocument2 : public IHTMLDocument
{
public:
    virtual HRESULT STDMETHODCALLTYPE get_all(
         /* [out] */ IHTMLElementCollection* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_body(
         /* [out] */ IHTMLElement* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_activeElement(
         /* [out] */ IHTMLElement* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_title(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_title(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_readyState(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_alinkColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_alinkColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_bgColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_bgColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_fgColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fgColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_linkColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_linkColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_vlinkColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_vlinkColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_expando(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_expando(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_nameProp(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandSupported(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandEnabled(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandState(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandIndeterm(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandText(
            /* [in] */ BSTR cmdID,/* [out] */ BSTR* pcmdText) = 0;

    virtual HRESULT STDMETHODCALLTYPE queryCommandValue(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT* pcmdValue) = 0;

    virtual HRESULT STDMETHODCALLTYPE execCommand(
            /* [in] */ BSTR cmdID,/* [in] */ VARIANT_BOOL showUI,/* [in] */ VARIANT value,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE execCommandShowHelp(
            /* [in] */ BSTR cmdID,/* [out] */ VARIANT_BOOL* pfRet) = 0;

    virtual HRESULT STDMETHODCALLTYPE createElement(
            /* [in] */ BSTR eTag,/* [out] */ IHTMLElement** newElem) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onhelp(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onhelp(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onclick(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onclick(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondblclick(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondblclick(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onkeyup(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeyup(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onkeydown(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeydown(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onkeypress(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeypress(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseup(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseup(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmousedown(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmousedown(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmousemove(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmousemove(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseout(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseout(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseover(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseover(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onreadystatechange(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onreadystatechange(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onafterupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onafterupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragstart(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragstart(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onselectstart(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onselectstart(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE elementFromPoint(
            /* [in] */ long x,/* [in] */ long y,/* [out] */ IHTMLElement** elementHit) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_parentWindow(
         /* [out] */ IHTMLWindow2* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforeupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforeupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onerrorupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onerrorupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE toString(
            /* [out] */ BSTR* String) = 0;

};

#endif     /* __IHTMLDocument2_INTERFACE_DEFINED__ */


#ifndef __IHTMLDocument3_INTERFACE_DEFINED__

#define __IHTMLDocument3_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLDocument3;


MIDL_INTERFACE("3050f485-98b5-11cf-bb82-00aa00bdce0b")
IHTMLDocument3 : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE releaseCapture(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_documentElement(
         /* [out] */ IHTMLElement* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_uniqueID(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE attachEvent(
            /* [in] */ BSTR event,/* [in] */ IDispatch* pDisp,/* [out] */ VARIANT_BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE detachEvent(
            /* [in] */ BSTR event,/* [in] */ IDispatch* pDisp) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onpropertychange(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onpropertychange(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_dir(
        /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_dir(
        /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_oncontextmenu(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_oncontextmenu(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onstop(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onstop(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_parentDocument(
         /* [out] */ IHTMLDocument2* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_enableDownload(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_enableDownload(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_baseUrl(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_baseUrl(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_inheritStyleSheets(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_inheritStyleSheets(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforeeditfocus(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforeeditfocus(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE getElementsByName(
            /* [in] */ BSTR v,/* [out] */ IHTMLElementCollection** pelColl) = 0;

    virtual HRESULT STDMETHODCALLTYPE getElementById(
            /* [in] */ BSTR v,/* [out] */ IHTMLElement** pel) = 0;

    virtual HRESULT STDMETHODCALLTYPE getElementsByTagName(
            /* [in] */ BSTR v,/* [out] */ IHTMLElementCollection** pelColl) = 0;

};

#endif     /* __IHTMLDocument3_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLDocument;



EXTERN_C const GUID DIID_DispHTMLDocument;


#ifndef _CDoc_PROPDESCS_
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocScript;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocall;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocbody;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocactiveElement;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocanchors;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocapplets;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDoclinks;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocforms;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocimages;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDoctitle;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocscripts;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocembeds;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocreadyState;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocframes;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocplugins;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocbgColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocfgColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocalinkColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocvlinkColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoclinkColor;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocexpando;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDoccharset;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocdefaultCharset;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocparentWindow;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCDocdir;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconhelp;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconclick;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocondblclick;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconkeyup;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconkeydown;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconkeypress;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconmouseup;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconmousedown;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconmousemove;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconmouseout;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconmouseover;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconreadystatechange;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconafterupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconbeforeupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDocondragstart;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconselectstart;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconerrorupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconpropertychange;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconcontextmenu;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconstop;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCDoconbeforeeditfocus;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocnameProp;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandSupported;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandEnabled;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandState;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandIndeterm;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandText;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocqueryCommandValue;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocexecCommand;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocexecCommandShowHelp;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDoccreateElement;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocelementFromPoint;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDoctoString;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocstyleSheets;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDoccreateStyleSheet;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocreleaseCapture;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocrecalc;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDoccreateTextNode;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocdocumentElement;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocuniqueID;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocattachEvent;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocdetachEvent;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocparentDocument;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocenableDownload;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocbaseUrl;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDocinheritStyleSheets;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocgetElementsByName;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocgetElementsByTagName;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDocgetElementById;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCDoczoom;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDoczoomNumerator;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCDoczoomDenominator;

#endif


#endif /*__document_h__*/

