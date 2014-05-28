
#ifndef __XINDOWS_SITE_H__
#define __XINDOWS_SITE_H__

enum ELEMENT_TAG
{
    ETAG_NULL = 0,
    ETAG_UNKNOWN = 1,
    ETAG_A = 2,
    ETAG_ACRONYM = 3,
    ETAG_ADDRESS = 4,
    ETAG_APPLET = 5,
    ETAG_AREA = 6,
    ETAG_B = 7,
    ETAG_BASE = 8,
    ETAG_BASEFONT = 9,
    ETAG_BDO = 10,
    ETAG_BGSOUND = 11,
    ETAG_BIG = 12,
    ETAG_BLINK = 13,
    ETAG_BLOCKQUOTE = 14,
    ETAG_BODY = 15,
    ETAG_BR = 16,
    ETAG_BUTTON = 17,
    ETAG_CAPTION = 18,
    ETAG_CENTER = 19,
    ETAG_CITE = 20,
    ETAG_CODE = 21,
    ETAG_COL = 22,
    ETAG_COLGROUP = 23,
    ETAG_COMMENT = 24,
    ETAG_DD = 25,
    ETAG_DEFAULT = 26,
    ETAG_DEL = 27,
    ETAG_DFN = 28,
    ETAG_DIR = 29,
    ETAG_DIV = 30,
    ETAG_DL = 31,
    ETAG_DT = 32,
    ETAG_EM = 33,
    ETAG_EMBED = 34,
    ETAG_FIELDSET = 35,
    ETAG_FONT = 36,
    ETAG_FORM = 37,
    ETAG_FRAME = 38,
    ETAG_FRAMESET = 39,
    ETAG_H1 = 40,
    ETAG_H2 = 41,
    ETAG_H3 = 42,
    ETAG_H4 = 43,
    ETAG_H5 = 44,
    ETAG_H6 = 45,
    ETAG_HEAD = 46,
    ETAG_HR = 47,
    ETAG_HTML = 48,
    ETAG_I = 49,
    ETAG_IFRAME = 50,
    ETAG_IMG = 51,
    ETAG_INPUT = 52,
    ETAG_TXTSLAVE = 53,
    ETAG_INS = 54,
    ETAG_ISINDEX = 55,
    ETAG_KBD = 56,
    ETAG_LABEL = 57,
    ETAG_LEGEND = 58,
    ETAG_LI = 59,
    ETAG_LINK = 60,
    ETAG_LISTING = 61,
    ETAG_MARQUEE = 63,
    ETAG_MENU = 64,
    ETAG_NEXTID = 66,
    ETAG_NOBR = 67,
    ETAG_NOEMBED = 68,
    ETAG_NOEMBED_OFF = 69,
    ETAG_NOFRAMES = 70,
    ETAG_NOFRAMES_OFF = 71,
    ETAG_NOSCRIPT = 72,
    ETAG_NOSCRIPT_OFF = 73,
    ETAG_OBJECT = 74,
    ETAG_OL = 75,
    ETAG_OPTION = 76,
    ETAG_P = 77,
    ETAG_PARAM = 78,
    ETAG_PLAINTEXT = 79,
    ETAG_PRE = 80,
    ETAG_Q = 81,
    ETAG_ROOT = 82,
    ETAG_RP = 83,
    ETAG_RT = 84,
    ETAG_RUBY = 85,
    ETAG_S = 86,
    ETAG_SAMP = 87,
    ETAG_SCRIPT = 88,
    ETAG_SELECT = 89,
    ETAG_SMALL = 90,
    ETAG_SPAN = 91,
    ETAG_STRIKE = 92,
    ETAG_STRONG = 93,
    ETAG_STYLE = 94,
    ETAG_SUB = 95,
    ETAG_SUP = 96,
    ETAG_TABLE = 97,
    ETAG_TBODY = 98,
    ETAG_TC = 99,
    ETAG_TD = 100,
    ETAG_TEXTAREA = 101,
    ETAG_TFOOT = 102,
    ETAG_TH = 103,
    ETAG_THEAD = 104,
    ETAG_TITLE_TAG = 105,
    ETAG_TR = 106,
    ETAG_TT = 107,
    ETAG_U = 108,
    ETAG_UL = 109,
    ETAG_VAR = 110,
    ETAG_WBR = 111,
    ETAG_XMP = 112,
    ETAG_GENERIC = 113,
    ETAG_GENERIC_LITERAL = 114,
    ETAG_GENERIC_BUILTIN = 115,
    ETAG_TITLE_ELEMENT = 116,
    ETAG_RAW_COMMENT = 117,
    ETAG_RAW_SOURCE = 118,
    ETAG_RAW_TEXT = 119,
    ETAG_RAW_EOF = 120,
    ETAG_RAW_CODEPAGE = 121,
    ETAG_RAW_DOCSIZE = 122,
    ETAG_RAW_BEGINFRAG = 123,
    ETAG_RAW_ENDFRAG = 124,
    ETAG_RAW_BEGINSEL = 125,
    ETAG_RAW_ENDSEL = 126,
    ETAG_ifdef = 127,
    ETAG_RAW_FACTORY = 128,
    ETAG_RAW_TAGSOURCE = 129,
    ETAG_endif = 130,
    ETAG_RAW_TEXTFRAG = 131,
    ETAG_RAW_INCLUDE = 132,
    ETAG_LAST = 133,
    ETAG_IMAGE = 51,
    ELEMENT_TAG_FORCE_LONG = LONG_MAX
};

enum TAGDESC_FLAGS
{
    TAGDESC_TEXTLESS                = 0x00000001,    // No text in a textrun should point to this
    TAGDESC_BLOCKELEMENT            = 0x00000002,    // causes a line break
    TAGDESC_LIST                    = 0x00000004,    // LI, DD, etc
    TAGDESC_LISTITEM                = 0x00000008,    // possibly preceded by a bullet/numeral/etc
    TAGDESC_HEADER                  = 0x00000010,    // H1, H2, etc
    TAGDESC_EDITREMOVABLE           = 0x00000020,    // Removed by "apply normal format"
    TAGDESC_BLKSTYLEDD              = 0x00000040,    // Block style in drop down combo
    TAGDESC_ALIGN                   = 0x00000080,    // element has alignment
    TAGDESC_CONDITIONALBLOCKELEMENT = 0x00000100,    // block element unless width specified
    TAGDESC_WONTBREAKLINE           = 0x00000200,    // block element's that dont cause the line breaker break a line (TABLE)
    TAGDESC_OWNLINE                 = 0x00000400,    // These tags have their own line.
    TAGDESC_ACCEPTHTML              = 0x00000800,    // Autodetect URLs when your ped has this.
    TAGDESC_UNUSED_0                = 0x00001000,    // - unused -
    TAGDESC_LITCTX                  = 0x00002000,    // Don't notify parent context; just create context and eat text (TITLE)
    TAGDESC_LITERALTAG              = 0x00004000,    // Taglike markup inside the scope of this tag is text
    TAGDESC_LITERALENT              = 0x00008000,    // Entitylike markup inside the scope of this tag is text
    TAGDESC_SPECIALTOKEN            = 0x00010000,    // Call SpecialToken when tag is tokenized
    TAGDESC_XSMBEGIN                = 0x00020000,    // Need to know XSM state before outputting begin tag
    TAGDESC_SAVEINDENT              = 0x00040000,    // Causes an indent while saving
    TAGDESC_SAVENBSPIFEMPTY         = 0x00080000,    // Element requires an nbsp if empty
    TAGDESC_SAVEALWAYSEND           = 0x00100000,    // Save an end tag even if _fExplicitEndTag is not set
    TAGDESC_SAVETAGOWNLINE          = 0x00200000,    // Save the tag on its own line
    TAGDESC_NEVERSAVEEND            = 0x00400000,    // Never save the end tag (BASE)
    TAGDESC_SPLITBLOCKINLIST        = 0x00800000,    // When editing split if this block is in a list
    TAGDESC_LOGICALINVISUAL         = 0x01000000,    // Element has logical text for visual page
    TAGDESC_CONTAINER               = 0x02000000,    // This flag is set for container tags: BODY, TEXTAREA, BUTTON, MARQUEE, RICHTEXT
    TAGDESC_WAITATSTART             = 0x04000000,    // Stop postparser at start of tag
    TAGDESC_WAITATEND               = 0x08000000,    // Stop postparser at end of tag
    TAGDESC_DONTWAITFORINPLACE      = 0x10000000,    // Stop but don't wait for in-place activation
    TAGDESC_SLOWPROCESS             = 0x20000000,    // Postparser needs to carefully examine this tag
    TAGDESC_ENTER_TREE_IMMEDIATELY  = 0x40000000,    // the tag needs to enter tree immediately
};

#include <MsHtmcid.h>
#include <MLang.h>
#include <ActivScp.h>

#include "gen/dwnnot.h"
#include "gen/prgsnk.h"

#include "util/TomConst.h"
#include "util/CellRangeParser.h"
#include "util/DataCache.h"
#include "util/GrabUtil.h"
#include "util/DrawInfo.h"
#include "util/Shape.h"
#include "util/ImgBits.h"
#include "util/Collect.h"

#include "text/mflags.h"
#include "text/txtdefs.h"
#include "text/usp.h"
#include "text/unisid.h"
#include "text/unipart.h"
#include "text/wchdefs.h"
#include "text/intl.h"
#include "text/unidir.h"

#include "util/FontInfo.h"

#include "text/Font.h"

#include "style/style.h"

#include "util/CfPf.h"
#include "util/PropertyBag.h"

#include "display/DispDefs.h"
#include "display/Flags.h"
#include "display/DispFlags.h"
#include "display/DispExtras.h"
#include "display/DispInfo.h"
#include "display/DispContext.h"
#include "display/SaveDispContext.h"
#include "display/DispSurface.h"
#include "display/DispFilter.h"
#include "display/DispClient.h"
#include "display/DispNode.h"
#include "display/DispLeaf.h"
#include "display/DispItemPlus.h"
#include "display/DispInterior.h"
#include "display/DispBalance.h"
#include "display/DispContent.h"
#include "display/DispContainer.h"
#include "display/DispContainerPlus.h"
#include "display/DispScroller.h"
#include "display/DispScrollerPlus.h"
#include "display/DispRoot.h"

#include "base/Settings.h"
#include "base/FormatInfo.h"
#include "base/EvtParam.h"
#include "base/Message.h"
#include "base/Notify.h"
#include "base/BorderInfo.h"
#include "base/TreePos.h"
#include "base/TreeNode.h"
#include "base/Element.h"
#include "base/TreePosGap.h"
#include "base/Timeout.h"
#include "base/Doc.h"

#include "base/Markup.h"
#include "base/MarkupPointer.h"
#include "base/Site.h"
#include "base/HtmTags.h"
#include "base/DragDrop.h"
#include "base/DropTarget.h"
#include "base/xBag.h"
#include "base/Breaker.h"
#include "base/Saver.h"
#include "base/ElementFactory.h"
#include "base/ProgSink.h"

#include "text/Disp.h"

#include "layout/Layout.h"
#include "layout/FlowLayout.h"
#include "layout/BodyLayout.h"
#include "layout/ImageLayout.h"
#include "layout/TxtLayout.h"

#include "display/DispSurface.h"

#include "text/TxtSite.h"
#include "text/BodyElement.h"
#include "text/Caret.h"
#include "text/fontlnk.h"
#include "text/TxtSlave.h"
#include "text/selrensv.h"
#include "text/lscache.h"
#include "text/_dxfrobj.h"
#include "text/SelDrag.h"
#include "text/_txtsave.h"
#include "text/LineServices.h"

#include "base/RootElement.h"

#include "miscelem/IEMediaPlayer.h"
#include "miscelem/ParaElement.h"
#include "miscelem/DivElement.h"

#include "builtin/TextArea.h"
#include "builtin/ImgHelper.h"
#include "builtin/ImgElement.h"

#include "edit/XindowsEditor.h"
#include "edit/edtrack.h"
#include "edit/blockcmd.h"

#include "stream/StreamWriteBuff.h"

#include "download/Download.h"
#include "download/Dwn.h"
#include "download/ImgAnim.h"
#include "download/ArtPlayer.h"

#endif //__XINDOWS_SITE_H__