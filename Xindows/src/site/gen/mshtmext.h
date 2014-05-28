#ifndef __mshtmext_h__
#define __mshtmext_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

/* header files for imported files */
#include "document.h"


#ifndef __IMarkupPointer_FWD_DEFINED__
#define __IMarkupPointer_FWD_DEFINED__
typedef interface IMarkupPointer IMarkupPointer;
#endif     /* __IMarkupPointer_FWD_DEFINED__ */

#ifndef __IMarkupContainer_FWD_DEFINED__
#define __IMarkupContainer_FWD_DEFINED__
typedef interface IMarkupContainer IMarkupContainer;
#endif     /* __IMarkupContainer_FWD_DEFINED__ */

#ifndef __IHTMLDocument2_FWD_DEFINED__
#define __IHTMLDocument2_FWD_DEFINED__
typedef interface IHTMLDocument2 IHTMLDocument2;
#endif     /* __IHTMLDocument2_FWD_DEFINED__ */

#ifndef __IMarkupServices_FWD_DEFINED__
#define __IMarkupServices_FWD_DEFINED__
typedef interface IMarkupServices IMarkupServices;
#endif     /* __IMarkupServices_FWD_DEFINED__ */

#ifndef __IMarkupContainer_FWD_DEFINED__
#define __IMarkupContainer_FWD_DEFINED__
typedef interface IMarkupContainer IMarkupContainer;
#endif     /* __IMarkupContainer_FWD_DEFINED__ */

#ifndef __IMarkupPointer_FWD_DEFINED__
#define __IMarkupPointer_FWD_DEFINED__
typedef interface IMarkupPointer IMarkupPointer;
#endif     /* __IMarkupPointer_FWD_DEFINED__ */

typedef enum _SECUREURLHOSTVALIDATE_FLAGS
{
    SUHV_PROMPTBEFORENO = 0x00000001,
    SUHV_SILENTYES = 0x00000002,
    SUHV_UNSECURESOURCE = 0x00000004,
    SECUREURLHOSTVALIDATE_FLAGS_Max = 2147483647L
} SECUREURLHOSTVALIDATE_FLAGS;


EXTERN_C const ENUMDESC s_enumdescSECUREURLHOSTVALIDATE_FLAGS;

typedef enum _POINTER_GRAVITY
{
    POINTER_GRAVITY_Left = 0,
    POINTER_GRAVITY_Right = 1,
    POINTER_GRAVITY_Max = 2147483647L
} POINTER_GRAVITY;


EXTERN_C const ENUMDESC s_enumdescPOINTER_GRAVITY;

typedef enum _ELEMENT_ADJACENCY
{
    ELEM_ADJ_BeforeBegin = 0,
    ELEM_ADJ_AfterBegin = 1,
    ELEM_ADJ_BeforeEnd = 2,
    ELEM_ADJ_AfterEnd = 3,
    ELEMENT_ADJACENCY_Max = 2147483647L
} ELEMENT_ADJACENCY;


EXTERN_C const ENUMDESC s_enumdescELEMENT_ADJACENCY;

typedef enum _MARKUP_CONTEXT_TYPE
{
    CONTEXT_TYPE_None = 0,
    CONTEXT_TYPE_Text = 1,
    CONTEXT_TYPE_EnterScope = 2,
    CONTEXT_TYPE_ExitScope = 3,
    CONTEXT_TYPE_NoScope = 4,
    MARKUP_CONTEXT_TYPE_Max = 2147483647L
} MARKUP_CONTEXT_TYPE;


EXTERN_C const ENUMDESC s_enumdescMARKUP_CONTEXT_TYPE;

typedef enum _FINDTEXT_FLAGS
{
    FINDTEXT_BACKWARDS = 0x00000001,
    FINDTEXT_WHOLEWORD = 0x00000002,
    FINDTEXT_MATCHCASE = 0x00000004,
    FINDTEXT_RAW = 0x00020000,
    FINDTEXT_MATCHDIAC = 0x20000000,
    FINDTEXT_MATCHKASHIDA = 0x40000000,
    FINDTEXT_MATCHALEFHAMZA = 0x80000000,
    FINDTEXT_FLAGS_Max = 2147483647L
} FINDTEXT_FLAGS;


EXTERN_C const ENUMDESC s_enumdescFINDTEXT_FLAGS;

typedef enum _MOVEUNIT_ACTION
{
    MOVEUNIT_PREVCHAR = 0,
    MOVEUNIT_NEXTCHAR = 1,
    MOVEUNIT_PREVCLUSTERBEGIN = 2,
    MOVEUNIT_NEXTCLUSTERBEGIN = 3,
    MOVEUNIT_PREVCLUSTEREND = 4,
    MOVEUNIT_NEXTCLUSTEREND = 5,
    MOVEUNIT_PREVWORDBEGIN = 6,
    MOVEUNIT_NEXTWORDBEGIN = 7,
    MOVEUNIT_PREVWORDEND = 8,
    MOVEUNIT_NEXTWORDEND = 9,
    MOVEUNIT_PREVPROOFWORD = 10,
    MOVEUNIT_NEXTPROOFWORD = 11,
    MOVEUNIT_NEXTURLBEGIN = 12,
    MOVEUNIT_PREVURLBEGIN = 13,
    MOVEUNIT_NEXTURLEND = 14,
    MOVEUNIT_PREVURLEND = 15,
    MOVEUNIT_PREVSENTENCE = 16,
    MOVEUNIT_NEXTSENTENCE = 17,
    MOVEUNIT_PREVBLOCK = 18,
    MOVEUNIT_NEXTBLOCK = 19,
    MOVEUNIT_ACTION_Max = 2147483647L
} MOVEUNIT_ACTION;


EXTERN_C const ENUMDESC s_enumdescMOVEUNIT_ACTION;

typedef enum _PARSE_FLAGS
{
    PARSE_ABSOLUTIFYIE40URLS = 0x00000001,
    PARSE_FLAGS_Max = 2147483647L
} PARSE_FLAGS;


EXTERN_C const ENUMDESC s_enumdescPARSE_FLAGS;

typedef enum _ELEMENT_TAG_ID
{
    TAGID_NULL = 0,
    TAGID_UNKNOWN = 1,
    TAGID_A = 2,
    TAGID_ACRONYM = 3,
    TAGID_ADDRESS = 4,
    TAGID_APPLET = 5,
    TAGID_AREA = 6,
    TAGID_B = 7,
    TAGID_BASE = 8,
    TAGID_BASEFONT = 9,
    TAGID_BDO = 10,
    TAGID_BGSOUND = 11,
    TAGID_BIG = 12,
    TAGID_BLINK = 13,
    TAGID_BLOCKQUOTE = 14,
    TAGID_BODY = 15,
    TAGID_BR = 16,
    TAGID_BUTTON = 17,
    TAGID_CAPTION = 18,
    TAGID_CENTER = 19,
    TAGID_CITE = 20,
    TAGID_CODE = 21,
    TAGID_COL = 22,
    TAGID_COLGROUP = 23,
    TAGID_COMMENT = 24,
    TAGID_COMMENT_RAW = 25,
    TAGID_DD = 26,
    TAGID_DEL = 27,
    TAGID_DFN = 28,
    TAGID_DIR = 29,
    TAGID_DIV = 30,
    TAGID_DL = 31,
    TAGID_DT = 32,
    TAGID_EM = 33,
    TAGID_EMBED = 34,
    TAGID_FIELDSET = 35,
    TAGID_FONT = 36,
    TAGID_FORM = 37,
    TAGID_FRAME = 38,
    TAGID_FRAMESET = 39,
    TAGID_GENERIC = 40,
    TAGID_H1 = 41,
    TAGID_H2 = 42,
    TAGID_H3 = 43,
    TAGID_H4 = 44,
    TAGID_H5 = 45,
    TAGID_H6 = 46,
    TAGID_HEAD = 47,
    TAGID_HR = 48,
    TAGID_HTML = 49,
    TAGID_I = 50,
    TAGID_IFRAME = 51,
    TAGID_IMG = 52,
    TAGID_INPUT = 53,
    TAGID_INS = 54,
    TAGID_KBD = 55,
    TAGID_LABEL = 56,
    TAGID_LEGEND = 57,
    TAGID_LI = 58,
    TAGID_LINK = 59,
    TAGID_LISTING = 60,
    TAGID_MAP = 61,
    TAGID_MARQUEE = 62,
    TAGID_MENU = 63,
    TAGID_META = 64,
    TAGID_NEXTID = 65,
    TAGID_NOBR = 66,
    TAGID_NOEMBED = 67,
    TAGID_NOFRAMES = 68,
    TAGID_NOSCRIPT = 69,
    TAGID_OBJECT = 70,
    TAGID_OL = 71,
    TAGID_OPTION = 72,
    TAGID_P = 73,
    TAGID_PARAM = 74,
    TAGID_PLAINTEXT = 75,
    TAGID_PRE = 76,
    TAGID_Q = 77,
    TAGID_RP = 78,
    TAGID_RT = 79,
    TAGID_RUBY = 80,
    TAGID_S = 81,
    TAGID_SAMP = 82,
    TAGID_SCRIPT = 83,
    TAGID_SELECT = 84,
    TAGID_SMALL = 85,
    TAGID_SPAN = 86,
    TAGID_STRIKE = 87,
    TAGID_STRONG = 88,
    TAGID_STYLE = 89,
    TAGID_SUB = 90,
    TAGID_SUP = 91,
    TAGID_TABLE = 92,
    TAGID_TBODY = 93,
    TAGID_TC = 94,
    TAGID_TD = 95,
    TAGID_TEXTAREA = 96,
    TAGID_TFOOT = 97,
    TAGID_TH = 98,
    TAGID_THEAD = 99,
    TAGID_TITLE = 100,
    TAGID_TR = 101,
    TAGID_TT = 102,
    TAGID_U = 103,
    TAGID_UL = 104,
    TAGID_VAR = 105,
    TAGID_WBR = 106,
    TAGID_XMP = 107,
    TAGID_COUNT = 108,
    TAGID_LAST_PREDEFINED = 10000,
    ELEMENT_TAG_ID_Max = 2147483647L
} ELEMENT_TAG_ID;


EXTERN_C const ENUMDESC s_enumdescELEMENT_TAG_ID;


#ifndef __IMarkupServices_INTERFACE_DEFINED__

#define __IMarkupServices_INTERFACE_DEFINED__

EXTERN_C const IID IID_IMarkupServices;


MIDL_INTERFACE("3050f4a0-98b5-11cf-bb82-00aa00bdce0b")
IMarkupServices : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE CreateMarkupPointer(
            /* [out] */ IMarkupPointer** ppPointer) = 0;

    virtual HRESULT STDMETHODCALLTYPE CreateMarkupContainer(
            /* [out] */ IMarkupContainer** ppMarkupContainer) = 0;

    virtual HRESULT STDMETHODCALLTYPE CreateElement(
            /* [in] */ ELEMENT_TAG_ID tagID,/* [in] */ OLECHAR* pchAttributes,/* [out] */ IHTMLElement** ppElement) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertElement(
            /* [in] */ IHTMLElement* pElementInsert,/* [in] */ IMarkupPointer* pPointerStart,/* [in] */ IMarkupPointer* pPointerFinish) = 0;

    virtual HRESULT STDMETHODCALLTYPE RemoveElement(
            /* [in] */ IHTMLElement* pElementRemove) = 0;

    virtual HRESULT STDMETHODCALLTYPE Remove(
            /* [in] */ IMarkupPointer* pPointerStart,/* [in] */ IMarkupPointer* pPointerFinish) = 0;

    virtual HRESULT STDMETHODCALLTYPE Copy(
            /* [in] */ IMarkupPointer* pPointerSourceStart,/* [in] */ IMarkupPointer* pPointerSourceFinish,/* [in] */ IMarkupPointer* pPointerTarget) = 0;

    virtual HRESULT STDMETHODCALLTYPE Move(
            /* [in] */ IMarkupPointer* pPointerSourceStart,/* [in] */ IMarkupPointer* pPointerSourceFinish,/* [in] */ IMarkupPointer* pPointerTarget) = 0;

    virtual HRESULT STDMETHODCALLTYPE InsertText(
            /* [in] */ OLECHAR* pchText,/* [in] */ long cch,/* [in] */ IMarkupPointer* pPointerTarget) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsScopedElement(
            /* [in] */ IHTMLElement* pElement,/* [out] */ BOOL* pfScoped) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetElementTagId(
            /* [in] */ IHTMLElement* pElement,/* [out] */ ELEMENT_TAG_ID* ptagId) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetTagIDForName(
            /* [in] */ BSTR bstrName,/* [out] */ ELEMENT_TAG_ID* ptagId) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetNameForTagID(
            /* [in] */ ELEMENT_TAG_ID tagId,/* [out] */ BSTR* pbstrName) = 0;

    virtual HRESULT STDMETHODCALLTYPE BeginUndoUnit(
            /* [in] */ OLECHAR* pchTitle) = 0;

    virtual HRESULT STDMETHODCALLTYPE EndUndoUnit(
            ) = 0;

};

#endif     /* __IMarkupServices_INTERFACE_DEFINED__ */


#ifndef __IMarkupContainer_INTERFACE_DEFINED__

#define __IMarkupContainer_INTERFACE_DEFINED__

EXTERN_C const IID IID_IMarkupContainer;


MIDL_INTERFACE("3050f5f9-98b5-11cf-bb82-00aa00bdce0b")
IMarkupContainer : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE OwningDoc(
            /* [out] */ IHTMLDocument2** ppDoc) = 0;

};

#endif     /* __IMarkupContainer_INTERFACE_DEFINED__ */


#ifndef __IMarkupPointer_INTERFACE_DEFINED__

#define __IMarkupPointer_INTERFACE_DEFINED__

EXTERN_C const IID IID_IMarkupPointer;


MIDL_INTERFACE("3050f49f-98b5-11cf-bb82-00aa00bdce0b")
IMarkupPointer : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE OwningDoc(
            /* [out] */ IHTMLDocument2** ppDoc) = 0;

    virtual HRESULT STDMETHODCALLTYPE Gravity(
            /* [out] */ POINTER_GRAVITY* pGravity) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetGravity(
            /* [in] */ POINTER_GRAVITY Gravity) = 0;

    virtual HRESULT STDMETHODCALLTYPE Cling(
            /* [out] */ BOOL* pfCling) = 0;

    virtual HRESULT STDMETHODCALLTYPE SetCling(
            /* [in] */ BOOL fCLing) = 0;

    virtual HRESULT STDMETHODCALLTYPE Unposition(
            ) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsPositioned(
            /* [out] */ BOOL* pfPositioned) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetContainer(
            /* [out] */ IMarkupContainer** ppContainer) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveAdjacentToElement(
            /* [in] */ IHTMLElement* pElement,/* [in] */ ELEMENT_ADJACENCY eAdj) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveToPointer(
            /* [in] */ IMarkupPointer* pPointer) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveToContainer(
            /* [in] */ IMarkupContainer* pContainer,/* [in] */ BOOL fAtStart) = 0;

    virtual HRESULT STDMETHODCALLTYPE Left(
            /* [in] */ BOOL fMove,/* [out] */ MARKUP_CONTEXT_TYPE* pContext,/* [out] */ IHTMLElement** ppElement,/* [in, out] */ long* pcch,/* [out] */ OLECHAR* pchText) = 0;

    virtual HRESULT STDMETHODCALLTYPE Right(
            /* [in] */ BOOL fMove,/* [out] */ MARKUP_CONTEXT_TYPE* pContext,/* [out] */ IHTMLElement** ppElement,/* [in, out] */ long* pcch,/* [out] */ OLECHAR* pchText) = 0;

    virtual HRESULT STDMETHODCALLTYPE CurrentScope(
            /* [out] */ IHTMLElement** ppElemCurrent) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsLeftOf(
            /* [in] */ IMarkupPointer* pPointerThat,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsLeftOfOrEqualTo(
            /* [in] */ IMarkupPointer* pPointerThat,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsRightOf(
            /* [in] */ IMarkupPointer* pPointerThat,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsRightOfOrEqualTo(
            /* [in] */ IMarkupPointer* pPointerThat,/* [out] */ BOOL* pfResult) = 0;

    virtual HRESULT STDMETHODCALLTYPE IsEqualTo(
            /* [in] */ IMarkupPointer* pPointerThat,/* [out] */ BOOL* pfAreEqual) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveUnit(
            /* [in] */ MOVEUNIT_ACTION muAction) = 0;

    virtual HRESULT STDMETHODCALLTYPE FindText(
            /* [in] */ OLECHAR* pchFindText,/* [in] */ DWORD dwFlags,/* [in] */ IMarkupPointer* pIEndMatch,/* [in] */ IMarkupPointer* pIEndSearch) = 0;

};

#endif     /* __IMarkupPointer_INTERFACE_DEFINED__ */



#endif /*__mshtmext_h__*/

