#ifndef __element_h__
#define __element_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "types.h"

/* header files for imported files */
#include "style.h"

#ifndef __IHTMLElementCollection_FWD_DEFINED__
#define __IHTMLElementCollection_FWD_DEFINED__
typedef interface IHTMLElementCollection IHTMLElementCollection;
#endif     /* __IHTMLElementCollection_FWD_DEFINED__ */

#ifndef __IHTMLElement_FWD_DEFINED__
#define __IHTMLElement_FWD_DEFINED__
typedef interface IHTMLElement IHTMLElement;
#endif     /* __IHTMLElement_FWD_DEFINED__ */

#ifndef __IHTMLElement2_FWD_DEFINED__
#define __IHTMLElement2_FWD_DEFINED__
typedef interface IHTMLElement2 IHTMLElement2;
#endif     /* __IHTMLElement2_FWD_DEFINED__ */

typedef enum _htmlListType
{
    htmlListTypeNotSet = 0,
    htmlListTypeLargeAlpha = 1,
    htmlListTypeSmallAlpha = 2,
    htmlListTypeLargeRoman = 3,
    htmlListTypeSmallRoman = 4,
    htmlListTypeNumbers = 5,
    htmlListTypeDisc = 6,
    htmlListTypeCircle = 7,
    htmlListTypeSquare = 8,
    htmlListType_Max = 2147483647L
} htmlListType;


EXTERN_C const ENUMDESC s_enumdeschtmlListType;

typedef enum _htmlMethod
{
    htmlMethodNotSet = 0,
    htmlMethodGet = 1,
    htmlMethodPost = 2,
    htmlMethod_Max = 2147483647L
} htmlMethod;


EXTERN_C const ENUMDESC s_enumdeschtmlMethod;

typedef enum _htmlWrap
{
    htmlWrapOff = 1,
    htmlWrapSoft = 2,
    htmlWrapHard = 3,
    htmlWrap_Max = 2147483647L
} htmlWrap;


EXTERN_C const ENUMDESC s_enumdeschtmlWrap;

typedef enum _htmlDir
{
    htmlDirNotSet = 0,
    htmlDirLeftToRight = 1,
    htmlDirRightToLeft = 2,
    htmlDir_Max = 2147483647L
} htmlDir;


EXTERN_C const ENUMDESC s_enumdeschtmlDir;

typedef enum _htmlInput
{
    htmlInputNotSet = 0,
    htmlInputButton = 1,
    htmlInputCheckbox = 2,
    htmlInputFile = 3,
    htmlInputHidden = 4,
    htmlInputImage = 5,
    htmlInputPassword = 6,
    htmlInputRadio = 7,
    htmlInputReset = 8,
    htmlInputSelectOne = 9,
    htmlInputSelectMultiple = 10,
    htmlInputSubmit = 11,
    htmlInputText = 12,
    htmlInputTextarea = 13,
    htmlInputRichtext = 14,
    htmlInput_Max = 2147483647L
} htmlInput;


EXTERN_C const ENUMDESC s_enumdeschtmlInput;

typedef enum _htmlEncoding
{
    htmlEncodingURL = 0,
    htmlEncodingMultipart = 1,
    htmlEncodingText = 2,
    htmlEncoding_Max = 2147483647L
} htmlEncoding;


EXTERN_C const ENUMDESC s_enumdeschtmlEncoding;

typedef enum _htmlAdjacency
{
    htmlAdjacencyBeforeBegin = 1,
    htmlAdjacencyAfterBegin = 2,
    htmlAdjacencyBeforeEnd = 3,
    htmlAdjacencyAfterEnd = 4,
    htmlAdjacency_Max = 2147483647L
} htmlAdjacency;


EXTERN_C const ENUMDESC s_enumdeschtmlAdjacency;

typedef enum _htmlTabIndex
{
    htmlTabIndexNotSet = -32768,
    htmlTabIndex_Max = 2147483647L
} htmlTabIndex;


EXTERN_C const ENUMDESC s_enumdeschtmlTabIndex;

typedef enum _htmlComponent
{
    htmlComponentClient = 0,
    htmlComponentSbLeft = 1,
    htmlComponentSbPageLeft = 2,
    htmlComponentSbHThumb = 3,
    htmlComponentSbPageRight = 4,
    htmlComponentSbRight = 5,
    htmlComponentSbUp = 6,
    htmlComponentSbPageUp = 7,
    htmlComponentSbVThumb = 8,
    htmlComponentSbPageDown = 9,
    htmlComponentSbDown = 10,
    htmlComponentSbLeft2 = 11,
    htmlComponentSbPageLeft2 = 12,
    htmlComponentSbRight2 = 13,
    htmlComponentSbPageRight2 = 14,
    htmlComponentSbUp2 = 15,
    htmlComponentSbPageUp2 = 16,
    htmlComponentSbDown2 = 17,
    htmlComponentSbPageDown2 = 18,
    htmlComponentSbTop = 19,
    htmlComponentSbBottom = 20,
    htmlComponentOutside = 21,
    htmlComponentGHTopLeft = 22,
    htmlComponentGHLeft = 23,
    htmlComponentGHTop = 24,
    htmlComponentGHBottomLeft = 25,
    htmlComponentGHTopRight = 26,
    htmlComponentGHBottom = 27,
    htmlComponentGHRight = 28,
    htmlComponentGHBottomRight = 29,
    htmlComponent_Max = 2147483647L
} htmlComponent;


EXTERN_C const ENUMDESC s_enumdeschtmlComponent;

typedef enum _htmlApplyLocation
{
    htmlApplyLocationInside = 0,
    htmlApplyLocationOutside = 1,
    htmlApplyLocation_Max = 2147483647L
} htmlApplyLocation;


EXTERN_C const ENUMDESC s_enumdeschtmlApplyLocation;


#ifndef __IHTMLElement_INTERFACE_DEFINED__

#define __IHTMLElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLElement;


MIDL_INTERFACE("3050f1ff-98b5-11cf-bb82-00aa00bdce0b")
IHTMLElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE setAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ VARIANT AttributeValue,/* [in] */ LONG lFlags) = 0;

    virtual HRESULT STDMETHODCALLTYPE getAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ LONG lFlags,/* [out] */ VARIANT* AttributeValue) = 0;

    virtual HRESULT STDMETHODCALLTYPE removeAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ LONG lFlags,/* [out] */ VARIANT_BOOL* pfSuccess) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_className(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_className(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_id(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_id(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_tagName(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_parentElement(
         /* [out] */ IHTMLElement* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_style(
         /* [out] */ IHTMLStyle* * p) = 0;

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

    virtual HRESULT STDMETHODCALLTYPE put_onkeydown(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeydown(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onkeyup(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeyup(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onkeypress(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onkeypress(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseout(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseout(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseover(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseover(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmousemove(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmousemove(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmousedown(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmousedown(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmouseup(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmouseup(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_document(
         /* [out] */ IDispatch* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_title(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_title(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_language(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_language(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onselectstart(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onselectstart(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE scrollIntoView(
            /* [in] */ VARIANT varargStart) = 0;

    virtual HRESULT STDMETHODCALLTYPE contains(
            /* [in] */ IHTMLElement* pChild,/* [out] */ VARIANT_BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_sourceIndex(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_lang(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_lang(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_offsetLeft(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_offsetTop(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_offsetWidth(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_offsetHeight(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_offsetParent(
         /* [out] */ IHTMLElement* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_innerText(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_innerText(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_outerText(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_outerText(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE click(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragstart(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragstart(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE toString(
            /* [out] */ BSTR* String) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforeupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforeupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onafterupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onafterupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onerrorupdate(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onerrorupdate(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onrowexit(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onrowexit(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onrowenter(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onrowenter(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_children(
         /* [out] */ IDispatch* * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_all(
         /* [out] */ IDispatch* * p) = 0;

};

#endif     /* __IHTMLElement_INTERFACE_DEFINED__ */


#ifndef __IHTMLElement2_INTERFACE_DEFINED__

#define __IHTMLElement2_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLElement2;


MIDL_INTERFACE("3050f434-98b5-11cf-bb82-00aa00bdce0b")
IHTMLElement2 : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE setCapture(
            /* [in] */ VARIANT_BOOL containerCapture) = 0;

    virtual HRESULT STDMETHODCALLTYPE releaseCapture(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onlosecapture(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onlosecapture(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE componentFromPoint(
            /* [in] */ long x,/* [in] */ long y,/* [out] */ BSTR* component) = 0;

    virtual HRESULT STDMETHODCALLTYPE doScroll(
            /* [in] */ VARIANT component) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onscroll(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onscroll(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondrag(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondrag(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragend(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragend(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragenter(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragenter(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragover(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragover(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondragleave(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondragleave(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_ondrop(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_ondrop(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforecut(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforecut(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_oncut(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_oncut(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforecopy(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforecopy(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_oncopy(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_oncopy(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforepaste(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforepaste(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onpaste(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onpaste(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onpropertychange(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onpropertychange(
         /* [out] */ VARIANT * p) = 0;

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

    virtual HRESULT STDMETHODCALLTYPE attachEvent(
            /* [in] */ BSTR event,/* [in] */ IDispatch* pDisp,/* [out] */ VARIANT_BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE detachEvent(
            /* [in] */ BSTR event,/* [in] */ IDispatch* pDisp) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_readyState(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onreadystatechange(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onreadystatechange(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_dir(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_dir(
         /* [out] */ BSTR * p) = 0;

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

    virtual HRESULT STDMETHODCALLTYPE clearAttributes(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_oncontextmenu(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_oncontextmenu(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_canHaveChildren(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onbeforeeditfocus(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onbeforeeditfocus(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_readyStateValue(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE getElementsByTagName(
            /* [in] */ BSTR v,/* [out] */ IHTMLElementCollection** pelColl) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_onmousehover(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_onmousehover(
         /* [out] */ VARIANT * p) = 0;

};

#endif     /* __IHTMLElement2_INTERFACE_DEFINED__ */


#ifndef _CElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementclassName;
EXTERN_C const PROPERTYDESC_CSTR_GETSET s_propdescCElementid;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementstyle_Str;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementtagName;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementparentElement;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementstyle;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementuniqueName;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementsubmitName;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementpropdescname;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonhelp;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonclick;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondblclick;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmouseout;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmouseover;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmouseup;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmousedown;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmousemove;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonkeyup;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonkeydown;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonkeypress;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonselectstart;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondragstart;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonlosecapture;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonscroll;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondrag;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondragend;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondragenter;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondragover;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondragleave;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementondrop;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonbeforecut;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementoncut;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonbeforecopy;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementoncopy;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonbeforepaste;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonpaste;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementdocument;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementlanguage;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementoffsetLeft;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementoffsetTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementoffsetWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementoffsetHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementoffsetParent;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementscrollIntoView;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementcontains;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCElementdir;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementtitle;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementsourceIndex;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementlang;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementinnerText;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementouterText;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementclick;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementtoString;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonpropertychange;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonbeforeupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonafterupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonerrorupdate;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementoncontextmenu;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonbeforeeditfocus;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementchildren;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementall;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementsetCapture;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementreleaseCapture;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementcomponentFromPoint;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementdoScroll;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCElementtabIndex;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCElementdisabled;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementfocus;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementblur;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonblur;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonfocus;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementaccessKey;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonresize;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementclientHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementclientWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementclientTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementclientLeft;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementattributes;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementuniqueID;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementreadyState;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonreadystatechange;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementscrollHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementscrollWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementscrollTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementscrollLeft;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementclearAttributes;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementcanHaveChildren;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementreadyStateValue;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementgetElementsByTagName;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCElementonmousehover;

#endif


#endif /*__element_h__*/

