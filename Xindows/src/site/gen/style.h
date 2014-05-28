#ifndef __style_h__
#define __style_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "types.h"

#ifndef __IHTMLStyle_FWD_DEFINED__
#define __IHTMLStyle_FWD_DEFINED__
typedef interface IHTMLStyle IHTMLStyle;
#endif     /* __IHTMLStyle_FWD_DEFINED__ */

#ifndef __IHTMLStyle2_FWD_DEFINED__
#define __IHTMLStyle2_FWD_DEFINED__
typedef interface IHTMLStyle2 IHTMLStyle2;
#endif     /* __IHTMLStyle2_FWD_DEFINED__ */
typedef enum _styleTextTransform
{
    styleTextTransformNotSet = 0,
    styleTextTransformCapitalize = 1,
    styleTextTransformLowercase = 2,
    styleTextTransformUppercase = 3,
    styleTextTransformNone = 4,
    styleTextTransform_Max = 2147483647L
} styleTextTransform;


EXTERN_C const ENUMDESC s_enumdescstyleTextTransform;

typedef enum _styleDataRepeat
{
    styleDataRepeatNone = 0,
    styleDataRepeatInner = 1,
    styleDataRepeat_Max = 2147483647L
} styleDataRepeat;


EXTERN_C const ENUMDESC s_enumdescstyleDataRepeat;

typedef enum _styleOverflow
{
    styleOverflowNotSet = 0,
    styleOverflowAuto = 1,
    styleOverflowHidden = 2,
    styleOverflowVisible = 3,
    styleOverflowScroll = 4,
    styleOverflow_Max = 2147483647L
} styleOverflow;


EXTERN_C const ENUMDESC s_enumdescstyleOverflow;

typedef enum _styleTableLayout
{
    styleTableLayoutNotSet = 0,
    styleTableLayoutAuto = 1,
    styleTableLayoutFixed = 2,
    styleTableLayout_Max = 2147483647L
} styleTableLayout;


EXTERN_C const ENUMDESC s_enumdescstyleTableLayout;

typedef enum _styleBorderCollapse
{
    styleBorderCollapseNotSet = 0,
    styleBorderCollapseSeparate = 1,
    styleBorderCollapseCollapse = 2,
    styleBorderCollapse_Max = 2147483647L
} styleBorderCollapse;


EXTERN_C const ENUMDESC s_enumdescstyleBorderCollapse;

typedef enum _styleFontStyle
{
    styleFontStyleNotSet = 0,
    styleFontStyleItalic = 1,
    styleFontStyleOblique = 2,
    styleFontStyleNormal = 3,
    styleFontStyle_Max = 2147483647L
} styleFontStyle;


EXTERN_C const ENUMDESC s_enumdescstyleFontStyle;

typedef enum _styleFontVariant
{
    styleFontVariantNotSet = 0,
    styleFontVariantSmallCaps = 1,
    styleFontVariantNormal = 2,
    styleFontVariant_Max = 2147483647L
} styleFontVariant;


EXTERN_C const ENUMDESC s_enumdescstyleFontVariant;

typedef enum _styleBackgroundRepeat
{
    styleBackgroundRepeatRepeat = 0,
    styleBackgroundRepeatRepeatX = 1,
    styleBackgroundRepeatRepeatY = 2,
    styleBackgroundRepeatNoRepeat = 3,
    styleBackgroundRepeatNotSet = 4,
    styleBackgroundRepeat_Max = 2147483647L
} styleBackgroundRepeat;


EXTERN_C const ENUMDESC s_enumdescstyleBackgroundRepeat;

typedef enum _styleBackgroundAttachment
{
    styleBackgroundAttachmentFixed = 0,
    styleBackgroundAttachmentScroll = 1,
    styleBackgroundAttachmentNotSet = 2,
    styleBackgroundAttachment_Max = 2147483647L
} styleBackgroundAttachment;


EXTERN_C const ENUMDESC s_enumdescstyleBackgroundAttachment;

typedef enum _styleVerticalAlign
{
    styleVerticalAlignBaseline = 0,
    styleVerticalAlignSub = 1,
    styleVerticalAlignSuper = 2,
    styleVerticalAlignTop = 3,
    styleVerticalAlignTextTop = 4,
    styleVerticalAlignMiddle = 5,
    styleVerticalAlignBottom = 6,
    styleVerticalAlignTextBottom = 7,
    styleVerticalAlignPercentage = 8,
    styleVerticalAlignNotSet = 9,
    styleVerticalAlign_Max = 2147483647L
} styleVerticalAlign;


EXTERN_C const ENUMDESC s_enumdescstyleVerticalAlign;

typedef enum _styleFontWeight
{
    styleFontWeightNotSet = 0,
    styleFontWeight100 = 1,
    styleFontWeight200 = 2,
    styleFontWeight300 = 3,
    styleFontWeight400 = 4,
    styleFontWeight500 = 5,
    styleFontWeight600 = 6,
    styleFontWeight700 = 7,
    styleFontWeight800 = 8,
    styleFontWeight900 = 9,
    styleFontWeightNormal = 10,
    styleFontWeightBold = 11,
    styleFontWeightBolder = 12,
    styleFontWeightLighter = 13,
    styleFontWeight_Max = 2147483647L
} styleFontWeight;


EXTERN_C const ENUMDESC s_enumdescstyleFontWeight;

typedef enum _styleBackgroundPositionX
{
    styleBackgroundPositionXNotSet = 0,
    styleBackgroundPositionXLeft = 1,
    styleBackgroundPositionXCenter = 2,
    styleBackgroundPositionXRight = 3,
    styleBackgroundPositionX_Max = 2147483647L
} styleBackgroundPositionX;


EXTERN_C const ENUMDESC s_enumdescstyleBackgroundPositionX;

typedef enum _styleBackgroundPositionY
{
    styleBackgroundPositionYNotSet = 0,
    styleBackgroundPositionYTop = 1,
    styleBackgroundPositionYCenter = 2,
    styleBackgroundPositionYBottom = 3,
    styleBackgroundPositionY_Max = 2147483647L
} styleBackgroundPositionY;


EXTERN_C const ENUMDESC s_enumdescstyleBackgroundPositionY;

typedef enum _styleFontSize
{
    styleFontSizeXXSmall = 0,
    styleFontSizeXSmall = 1,
    styleFontSizeSmall = 2,
    styleFontSizeMedium = 3,
    styleFontSizeLarge = 4,
    styleFontSizeXLarge = 5,
    styleFontSizeXXLarge = 6,
    styleFontSizeSmaller = 7,
    styleFontSizeLarger = 8,
    styleFontSize_Max = 2147483647L
} styleFontSize;


EXTERN_C const ENUMDESC s_enumdescstyleFontSize;

typedef enum _styleAuto
{
    styleAutoAuto = 0,
    styleAuto_Max = 2147483647L
} styleAuto;


EXTERN_C const ENUMDESC s_enumdescstyleAuto;

typedef enum _styleNormal
{
    styleNormalNormal = 0,
    styleNormal_Max = 2147483647L
} styleNormal;


EXTERN_C const ENUMDESC s_enumdescstyleNormal;

typedef enum _styleBorderWidth
{
    styleBorderWidthThin = 0,
    styleBorderWidthMedium = 1,
    styleBorderWidthThick = 2,
    styleBorderWidth_Max = 2147483647L
} styleBorderWidth;


EXTERN_C const ENUMDESC s_enumdescstyleBorderWidth;

typedef enum _stylePosition
{
    stylePositionNotSet = 0,
    stylePositionstatic = 1,
    stylePositionrelative = 2,
    stylePositionabsolute = 3,
    stylePositionfixed = 4,
    stylePosition_Max = 2147483647L
} stylePosition;


EXTERN_C const ENUMDESC s_enumdescstylePosition;

typedef enum _styleBorderStyle
{
    styleBorderStyleNotSet = 0,
    styleBorderStyleDotted = 1,
    styleBorderStyleDashed = 2,
    styleBorderStyleSolid = 3,
    styleBorderStyleDouble = 4,
    styleBorderStyleGroove = 5,
    styleBorderStyleRidge = 6,
    styleBorderStyleInset = 7,
    styleBorderStyleOutset = 8,
    styleBorderStyleNone = 9,
    styleBorderStyle_Max = 2147483647L
} styleBorderStyle;


EXTERN_C const ENUMDESC s_enumdescstyleBorderStyle;

typedef enum _styleStyleFloat
{
    styleStyleFloatNotSet = 0,
    styleStyleFloatLeft = 1,
    styleStyleFloatRight = 2,
    styleStyleFloatNone = 3,
    styleStyleFloat_Max = 2147483647L
} styleStyleFloat;


EXTERN_C const ENUMDESC s_enumdescstyleStyleFloat;

typedef enum _styleDisplay
{
    styleDisplayNotSet = 0,
    styleDisplayBlock = 1,
    styleDisplayInline = 2,
    styleDisplayListItem = 3,
    styleDisplayNone = 4,
    styleDisplayTableHeaderGroup = 5,
    styleDisplayTableFooterGroup = 6,
    styleDisplay_Max = 2147483647L
} styleDisplay;


EXTERN_C const ENUMDESC s_enumdescstyleDisplay;

typedef enum _styleVisibility
{
    styleVisibilityNotSet = 0,
    styleVisibilityInherit = 1,
    styleVisibilityVisible = 2,
    styleVisibilityHidden = 3,
    styleVisibility_Max = 2147483647L
} styleVisibility;


EXTERN_C const ENUMDESC s_enumdescstyleVisibility;

typedef enum _styleListStyleType
{
    styleListStyleTypeNotSet = 0,
    styleListStyleTypeDisc = 1,
    styleListStyleTypeCircle = 2,
    styleListStyleTypeSquare = 3,
    styleListStyleTypeDecimal = 4,
    styleListStyleTypeLowerRoman = 5,
    styleListStyleTypeUpperRoman = 6,
    styleListStyleTypeLowerAlpha = 7,
    styleListStyleTypeUpperAlpha = 8,
    styleListStyleTypeNone = 9,
    styleListStyleType_Max = 2147483647L
} styleListStyleType;


EXTERN_C const ENUMDESC s_enumdescstyleListStyleType;

typedef enum _styleListStylePosition
{
    styleListStylePositionNotSet = 0,
    styleListStylePositionInside = 1,
    styleListStylePositionOutSide = 2,
    styleListStylePosition_Max = 2147483647L
} styleListStylePosition;


EXTERN_C const ENUMDESC s_enumdescstyleListStylePosition;

typedef enum _styleWhiteSpace
{
    styleWhiteSpaceNotSet = 0,
    styleWhiteSpaceNormal = 1,
    styleWhiteSpacePre = 2,
    styleWhiteSpaceNowrap = 3,
    styleWhiteSpace_Max = 2147483647L
} styleWhiteSpace;


EXTERN_C const ENUMDESC s_enumdescstyleWhiteSpace;

typedef enum _stylePageBreak
{
    stylePageBreakNotSet = 0,
    stylePageBreakAuto = 1,
    stylePageBreakAlways = 2,
    stylePageBreakLeft = 3,
    stylePageBreakRight = 4,
    stylePageBreak_Max = 2147483647L
} stylePageBreak;


EXTERN_C const ENUMDESC s_enumdescstylePageBreak;

typedef enum _styleCursor
{
    styleCursorAuto = 0,
    styleCursorCrosshair = 1,
    styleCursorDefault = 2,
    styleCursorHand = 3,
    styleCursorMove = 4,
    styleCursorE_resize = 5,
    styleCursorNe_resize = 6,
    styleCursorNw_resize = 7,
    styleCursorN_resize = 8,
    styleCursorSe_resize = 9,
    styleCursorSw_resize = 10,
    styleCursorS_resize = 11,
    styleCursorW_resize = 12,
    styleCursorText = 13,
    styleCursorWait = 14,
    styleCursorHelp = 15,
    styleCursorNotSet = 16,
    styleCursor_Max = 2147483647L
} styleCursor;


EXTERN_C const ENUMDESC s_enumdescstyleCursor;

typedef enum _styleDir
{
    styleDirNotSet = 0,
    styleDirLeftToRight = 1,
    styleDirRightToLeft = 2,
    styleDirInherit = 3,
    styleDir_Max = 2147483647L
} styleDir;


EXTERN_C const ENUMDESC s_enumdescstyleDir;

typedef enum _styleBidi
{
    styleBidiNotSet = 0,
    styleBidiNormal = 1,
    styleBidiEmbed = 2,
    styleBidiOverride = 3,
    styleBidiInherit = 4,
    styleBidi_Max = 2147483647L
} styleBidi;


EXTERN_C const ENUMDESC s_enumdescstyleBidi;

typedef enum _styleImeMode
{
    styleImeModeAuto = 0,
    styleImeModeActive = 1,
    styleImeModeInactive = 2,
    styleImeModeDisabled = 3,
    styleImeModeNotSet = 4,
    styleImeMode_Max = 2147483647L
} styleImeMode;


EXTERN_C const ENUMDESC s_enumdescstyleImeMode;

typedef enum _styleRubyAlign
{
    styleRubyAlignNotSet = 0,
    styleRubyAlignAuto = 1,
    styleRubyAlignLeft = 2,
    styleRubyAlignCenter = 3,
    styleRubyAlignRight = 4,
    styleRubyAlignDistributeLetter = 5,
    styleRubyAlignDistributeSpace = 6,
    styleRubyAlignLineEdge = 7,
    styleRubyAlign_Max = 2147483647L
} styleRubyAlign;


EXTERN_C const ENUMDESC s_enumdescstyleRubyAlign;

typedef enum _styleRubyPosition
{
    styleRubyPositionNotSet = 0,
    styleRubyPositionAbove = 1,
    styleRubyPositionInline = 2,
    styleRubyPosition_Max = 2147483647L
} styleRubyPosition;


EXTERN_C const ENUMDESC s_enumdescstyleRubyPosition;

typedef enum _styleRubyOverhang
{
    styleRubyOverhangNotSet = 0,
    styleRubyOverhangAuto = 1,
    styleRubyOverhangWhitespace = 2,
    styleRubyOverhangNone = 3,
    styleRubyOverhang_Max = 2147483647L
} styleRubyOverhang;


EXTERN_C const ENUMDESC s_enumdescstyleRubyOverhang;

typedef enum _styleLayoutGridChar
{
    styleLayoutGridCharNotSet = 0,
    styleLayoutGridCharAuto = 1,
    styleLayoutGridCharNone = 2,
    styleLayoutGridChar_Max = 2147483647L
} styleLayoutGridChar;


EXTERN_C const ENUMDESC s_enumdescstyleLayoutGridChar;

typedef enum _styleLayoutGridLine
{
    styleLayoutGridLineNotSet = 0,
    styleLayoutGridLineAuto = 1,
    styleLayoutGridLineNone = 2,
    styleLayoutGridLine_Max = 2147483647L
} styleLayoutGridLine;


EXTERN_C const ENUMDESC s_enumdescstyleLayoutGridLine;

typedef enum _styleLayoutGridMode
{
    styleLayoutGridModeNotSet = 0,
    styleLayoutGridModeChar = 1,
    styleLayoutGridModeLine = 2,
    styleLayoutGridModeBoth = 3,
    styleLayoutGridModeNone = 4,
    styleLayoutGridMode_Max = 2147483647L
} styleLayoutGridMode;


EXTERN_C const ENUMDESC s_enumdescstyleLayoutGridMode;

typedef enum _styleLayoutGridType
{
    styleLayoutGridTypeNotSet = 0,
    styleLayoutGridTypeLoose = 1,
    styleLayoutGridTypeStrict = 2,
    styleLayoutGridTypeFixed = 3,
    styleLayoutGridType_Max = 2147483647L
} styleLayoutGridType;


EXTERN_C const ENUMDESC s_enumdescstyleLayoutGridType;

typedef enum _styleLineBreak
{
    styleLineBreakNotSet = 0,
    styleLineBreakNormal = 1,
    styleLineBreakStrict = 2,
    styleLineBreak_Max = 2147483647L
} styleLineBreak;


EXTERN_C const ENUMDESC s_enumdescstyleLineBreak;

typedef enum _styleWordBreak
{
    styleWordBreakNotSet = 0,
    styleWordBreakNormal = 1,
    styleWordBreakBreakAll = 2,
    styleWordBreakKeepAll = 3,
    styleWordBreak_Max = 2147483647L
} styleWordBreak;


EXTERN_C const ENUMDESC s_enumdescstyleWordBreak;

typedef enum _styleTextJustify
{
    styleTextJustifyNotSet = 0,
    styleTextJustifyInterWord = 1,
    styleTextJustifyNewspaper = 2,
    styleTextJustifyDistribute = 3,
    styleTextJustifyDistributeAllLines = 4,
    styleTextJustifyInterIdeograph = 5,
    styleTextJustifyAuto = 6,
    styleTextJustify_Max = 2147483647L
} styleTextJustify;


EXTERN_C const ENUMDESC s_enumdescstyleTextJustify;

typedef enum _styleTextJustifyTrim
{
    styleTextJustifyTrimNotSet = 0,
    styleTextJustifyTrimNone = 1,
    styleTextJustifyTrimPunctuation = 2,
    styleTextJustifyTrimPunctAndKana = 3,
    styleTextJustifyTrim_Max = 2147483647L
} styleTextJustifyTrim;


EXTERN_C const ENUMDESC s_enumdescstyleTextJustifyTrim;

typedef enum _styleAccelerator
{
    styleAcceleratorFalse = 0,
    styleAcceleratorTrue = 1,
    styleAccelerator_Max = 2147483647L
} styleAccelerator;


EXTERN_C const ENUMDESC s_enumdescstyleAccelerator;


#ifndef __IHTMLStyle_INTERFACE_DEFINED__

#define __IHTMLStyle_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLStyle;


MIDL_INTERFACE("3050f25e-98b5-11cf-bb82-00aa00bdce0b")
IHTMLStyle : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_fontFamily(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fontFamily(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_fontStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fontStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_fontVariant(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fontVariant(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_fontWeight(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fontWeight(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_fontSize(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_fontSize(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_font(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_font(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_color(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_color(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_background(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_background(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundImage(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundImage(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundRepeat(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundRepeat(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundAttachment(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundAttachment(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundPosition(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundPosition(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundPositionX(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundPositionX(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_backgroundPositionY(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_backgroundPositionY(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_wordSpacing(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_wordSpacing(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_letterSpacing(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_letterSpacing(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecoration(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecoration(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecorationNone(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecorationNone(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecorationUnderline(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecorationUnderline(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecorationOverline(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecorationOverline(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecorationLineThrough(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecorationLineThrough(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textDecorationBlink(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textDecorationBlink(
         /* [out] */ VARIANT_BOOL * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_verticalAlign(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_verticalAlign(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textTransform(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textTransform(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textAlign(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textAlign(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textIndent(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textIndent(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_lineHeight(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_lineHeight(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_marginTop(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_marginTop(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_marginRight(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_marginRight(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_marginBottom(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_marginBottom(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_marginLeft(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_marginLeft(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_margin(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_margin(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_paddingTop(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_paddingTop(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_paddingRight(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_paddingRight(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_paddingBottom(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_paddingBottom(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_paddingLeft(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_paddingLeft(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_padding(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_padding(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_border(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_border(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderTop(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderTop(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderRight(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderRight(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderBottom(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderBottom(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderLeft(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderLeft(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderColor(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderColor(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderTopColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderTopColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderRightColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderRightColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderBottomColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderBottomColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderLeftColor(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderLeftColor(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderWidth(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderWidth(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderTopWidth(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderTopWidth(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderRightWidth(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderRightWidth(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderBottomWidth(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderBottomWidth(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderLeftWidth(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderLeftWidth(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderTopStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderTopStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderRightStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderRightStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderBottomStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderBottomStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderLeftStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderLeftStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_width(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_width(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_height(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_height(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_styleFloat(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_styleFloat(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_clear(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clear(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_display(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_display(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_visibility(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_visibility(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_listStyleType(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_listStyleType(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_listStylePosition(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_listStylePosition(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_listStyleImage(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_listStyleImage(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_listStyle(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_listStyle(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_whiteSpace(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_whiteSpace(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_top(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_top(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_left(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_left(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_position(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_zIndex(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_zIndex(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_overflow(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_overflow(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pageBreakBefore(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pageBreakBefore(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pageBreakAfter(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pageBreakAfter(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_cssText(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_cssText(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelTop(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelTop(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelLeft(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelLeft(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelWidth(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelWidth(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelHeight(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelHeight(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posTop(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posTop(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posLeft(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posLeft(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posWidth(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posWidth(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posHeight(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posHeight(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_cursor(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_cursor(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_clip(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clip(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE setAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ VARIANT AttributeValue,/* [in] */ LONG lFlags) = 0;

    virtual HRESULT STDMETHODCALLTYPE getAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ LONG lFlags,/* [out] */ VARIANT* AttributeValue) = 0;

    virtual HRESULT STDMETHODCALLTYPE removeAttribute(
            /* [in] */ BSTR strAttributeName,/* [in] */ LONG lFlags,/* [out] */ VARIANT_BOOL* pfSuccess) = 0;

    virtual HRESULT STDMETHODCALLTYPE toString(
            /* [out] */ BSTR* String) = 0;

};

#endif     /* __IHTMLStyle_INTERFACE_DEFINED__ */


#ifndef __IHTMLStyle2_INTERFACE_DEFINED__

#define __IHTMLStyle2_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLStyle2;


MIDL_INTERFACE("3050f4a2-98b5-11cf-bb82-00aa00bdce0b")
IHTMLStyle2 : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_tableLayout(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_tableLayout(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_borderCollapse(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_borderCollapse(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_direction(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_direction(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_behavior(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_behavior(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_position(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_position(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_unicodeBidi(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_unicodeBidi(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_bottom(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_bottom(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_right(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_right(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelBottom(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelBottom(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_pixelRight(
         /* [in] */ long v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_pixelRight(
         /* [out] */ long * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posBottom(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posBottom(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_posRight(
         /* [in] */ float v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_posRight(
         /* [out] */ float * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_imeMode(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_imeMode(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_rubyAlign(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_rubyAlign(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_rubyPosition(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_rubyPosition(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_rubyOverhang(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_rubyOverhang(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_layoutGridChar(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_layoutGridChar(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_layoutGridLine(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_layoutGridLine(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_layoutGridMode(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_layoutGridMode(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_layoutGridType(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_layoutGridType(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_layoutGrid(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_layoutGrid(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_wordBreak(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_wordBreak(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_lineBreak(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_lineBreak(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textJustify(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textJustify(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textJustifyTrim(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textJustifyTrim(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textKashida(
         /* [in] */ VARIANT v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textKashida(
         /* [out] */ VARIANT * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_textAutospace(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_textAutospace(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_overflowX(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_overflowX(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_overflowY(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_overflowY(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_accelerator(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_accelerator(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLStyle2_INTERFACE_DEFINED__ */


EXTERN_C const GUID GUID_HTMLStyle;



EXTERN_C const GUID DIID_DispHTMLStyle;


#ifndef _CStyle_PROPDESCS_
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylefontFamily;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylefontStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylefontVariant;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylefontWeight;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylefontSize;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylefont;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylecolor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylebackground;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylebackgroundColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylebackgroundImage;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylebackgroundRepeat;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylebackgroundAttachment;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylebackgroundPosition;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylebackgroundPositionX;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylebackgroundPositionY;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylewordSpacing;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleletterSpacing;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecoration;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecorationNone;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecorationUnderline;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecorationOverline;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecorationLineThrough;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextDecorationBlink;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleaccelerator;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleverticalAlign;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletextTransform;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletextAlign;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletextIndent;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylelineHeight;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylemarginTop;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylemarginRight;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylemarginBottom;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylemarginLeft;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylemargin;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepaddingTop;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepaddingRight;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepaddingBottom;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepaddingLeft;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylepadding;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborder;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderTop;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderRight;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderBottom;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderLeft;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderTopColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderRightColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderBottomColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderLeftColor;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderWidth;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleborderTopWidth;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleborderRightWidth;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleborderBottomWidth;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleborderLeftWidth;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleborderStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleborderTopStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleborderRightStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleborderBottomStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleborderLeftStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylestyleFloat;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleclear;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyledisplay;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylevisibility;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylelistStyleType;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylelistStylePosition;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylelistStyleImage;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylelistStyle;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylewhiteSpace;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyletop;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylebottom;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleleft;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleright;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylewidth;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleheight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelLeft;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposTop;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposLeft;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposWidth;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposHeight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelBottom;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStylepixelRight;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposBottom;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCStyleposRight;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleoverflow;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepageBreakBefore;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylepageBreakAfter;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylecssText;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleposition2;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylezIndex;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyleclip;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleclipTop;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleclipRight;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleclipBottom;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyleclipLeft;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylecursor;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletableLayout;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleborderCollapse;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCStyletoString;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyledirection;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleunicodeBidi;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylebehavior;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleposition;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleimeMode;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylerubyAlign;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylerubyPosition;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylerubyOverhang;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStyletextAutospace;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylelineBreak;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylewordBreak;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylelayoutGridChar;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStylelayoutGridLine;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylelayoutGridMode;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStylelayoutGridType;
EXTERN_C const PROPERTYDESC_BASIC s_propdescCStylelayoutGrid;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletextJustify;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyletextJustifyTrim;
EXTERN_C const PROPERTYDESC_NUMPROP_ENUMREF s_propdescCStyletextKashida;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleoverflowX;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCStyleoverflowY;

#endif


EXTERN_C const GUID GUID_HTMLRuleStyle;



EXTERN_C const GUID DIID_DispHTMLRuleStyle;


#ifndef _CRuleStyle_PROPDESCS_

#endif


#endif /*__style_h__*/

