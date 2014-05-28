
#include "stdafx.h"
#include "HtmTags.h"

#define _cxx_ 
#include "../gen/entity.h"

//+------------------------------------------------------------------------
//
//  Member:     HpcFromEtag
//
//  Synopsis:   returns the parse class for the etag
//
//  Arguments:  tag (ELEMENT_TAG)
//
//  Returns:    CHtmlParseClass *
//
//-------------------------------------------------------------------------
CHtmlParseClass* HpcFromEtag(ELEMENT_TAG tag)
{
    const CTagDesc* ptd = TagDescFromEtag(tag);

    Assert(ptd);

    return ptd?ptd->_pParseClass:0;
}

//+------------------------------------------------------------------------
//
//  Member:     TagProhibitedContainer(ELEMENT_TAG tag1, tag2)
//
//  Synopsis:   returns TRUE if tag2 is a prohibited container of tag1.
//
//  Arguments:  tag1, tag2 (ELEMENT_TAG)
//
//  Returns:    BOOL
//
//-------------------------------------------------------------------------
BOOL TagProhibitedContainer(ELEMENT_TAG tag1, ELEMENT_TAG tag2)
{
    CHtmlParseClass* phpc1 = HpcFromEtag(tag1);

    return (phpc1 && phpc1->_atagProhibitedContainers &&
        IsEtagInSet(tag2, phpc1->_atagProhibitedContainers));
}



CAssoc g_tagascA2			= {     2, 0x00000041, _T("A") };
CAssoc g_tagascACRONYM3			= {     3, 0x237a7d76, _T("ACRONYM") };
CAssoc g_tagascADDRESS4			= {     4, 0x3756949c, _T("ADDRESS") };
CAssoc g_tagascAPPLET5			= {     5, 0xab32855c, _T("APPLET") };
CAssoc g_tagascAREA6			= {     6, 0x8b4a0841, _T("AREA") };
CAssoc g_tagascB7			= {     7, 0x00000042, _T("B") };
CAssoc g_tagascBASE8			= {     8, 0xa7061045, _T("BASE") };
CAssoc g_tagascBASEFONT9			= {     9, 0x0d9f34af, _T("BASEFONT") };
CAssoc g_tagascBDO10			= {    10, 0x8908004f, _T("BDO") };
CAssoc g_tagascBGSOUND11			= {    11, 0x8dd67d7d, _T("BGSOUND") };
CAssoc g_tagascBIG12			= {    12, 0x93080047, _T("BIG") };
CAssoc g_tagascBLINK13			= {    13, 0x9d26646b, _T("BLINK") };
CAssoc g_tagascBLOCKQUOTE14			= {    14, 0x1a2679e3, _T("BLOCKQUOTE") };
CAssoc g_tagascBODY15			= {    15, 0x893e1059, _T("BODY") };
CAssoc g_tagascBR16			= {    16, 0x84000052, _T("BR") };
CAssoc g_tagascBUTTON17			= {    17, 0xdf52a5a6, _T("BUTTON") };
CAssoc g_tagascCAPTION18			= {    18, 0xcfe6a556, _T("CAPTION") };
CAssoc g_tagascCENTER19			= {    19, 0xeb5274aa, _T("CENTER") };
CAssoc g_tagascCITE20			= {    20, 0xa9261845, _T("CITE") };
CAssoc g_tagascCODE21			= {    21, 0x893e1845, _T("CODE") };
CAssoc g_tagascCOL22			= {    22, 0x9f0c004c, _T("COL") };
CAssoc g_tagascCOLGROUP23			= {    23, 0x3f2014ca, _T("COLGROUP") };
CAssoc g_tagascCOMMENT24			= {    24, 0x8dd66d2e, _T("COMMENT") };
CAssoc g_tagascDD25			= {    25, 0x88000044, _T("DD") };
CAssoc g_tagascDEL27			= {    27, 0x8b10004c, _T("DEL") };
CAssoc g_tagascDFN28			= {    28, 0x8d10004e, _T("DFN") };
CAssoc g_tagascDIR29			= {    29, 0x93100052, _T("DIR") };
CAssoc g_tagascDIV30			= {    30, 0x93100056, _T("DIV") };
CAssoc g_tagascDL31			= {    31, 0x8800004c, _T("DL") };
CAssoc g_tagascDT32			= {    32, 0x88000054, _T("DT") };
CAssoc g_tagascEM33			= {    33, 0x8a00004d, _T("EM") };
CAssoc g_tagascEMBED34			= {    34, 0x8b0a6c94, _T("EMBED") };
CAssoc g_tagascFIELDSET35			= {    35, 0x3db1251d, _T("FIELDSET") };
CAssoc g_tagascFONT36			= {    36, 0x9d3e3054, _T("FONT") };
CAssoc g_tagascFORM37			= {    37, 0xa53e304d, _T("FORM") };
CAssoc g_tagascFRAME38			= {    38, 0x9b0694a5, _T("FRAME") };
CAssoc g_tagascFRAMESET39			= {    39, 0xbff12d2c, _T("FRAMESET") };
CAssoc g_tagascH140			= {    40, 0x90000011, _T("H1") };
CAssoc g_tagascH241			= {    41, 0x90000012, _T("H2") };
CAssoc g_tagascH342			= {    42, 0x90000013, _T("H3") };
CAssoc g_tagascH443			= {    43, 0x90000014, _T("H4") };
CAssoc g_tagascH544			= {    44, 0x90000015, _T("H5") };
CAssoc g_tagascH645			= {    45, 0x90000016, _T("H6") };
CAssoc g_tagascHEAD46			= {    46, 0x83164044, _T("HEAD") };
CAssoc g_tagascHR47			= {    47, 0x90000052, _T("HR") };
CAssoc g_tagascHTML48			= {    48, 0x9b52404c, _T("HTML") };
CAssoc g_tagascI49			= {    49, 0x00000049, _T("I") };
CAssoc g_tagascIFRAME50			= {    50, 0xbb0694ae, _T("IFRAME") };
CAssoc g_tagascIMG51			= {    51, 0x9b240047, _T("IMG") };
CAssoc g_tagascINPUT52			= {    52, 0xab4274e4, _T("INPUT") };
CAssoc g_tagascINS54			= {    54, 0x9d240053, _T("INS") };
CAssoc g_tagascISINDEX55			= {    55, 0xfd5274f2, _T("ISINDEX") };
CAssoc g_tagascKBD56			= {    56, 0x852c0044, _T("KBD") };
CAssoc g_tagascLABEL57			= {    57, 0x8b0a0d0c, _T("LABEL") };
CAssoc g_tagascLEGEND58			= {    58, 0x1d163c9e, _T("LEGEND") };
CAssoc g_tagascLI59			= {    59, 0x98000049, _T("LI") };
CAssoc g_tagascLINK60			= {    60, 0x9d26604b, _T("LINK") };
CAssoc g_tagascLISTING61			= {    61, 0xd026a580, _T("LISTING") };
CAssoc g_tagascMAP62			= {    62, 0x83340050, _T("MAP") };
CAssoc g_tagascMARQUEE63			= {    63, 0xbe968d6d, _T("MARQUEE") };
CAssoc g_tagascMENU64			= {    64, 0x9d166855, _T("MENU") };
CAssoc g_tagascMETA65			= {    65, 0xa9166841, _T("META") };
CAssoc g_tagascNEXTID66			= {    66, 0x5352c49e, _T("NEXTID") };
CAssoc g_tagascNOBR67			= {    67, 0x853e7052, _T("NOBR") };
CAssoc g_tagascNOEMBED68			= {    68, 0x7e8a6c9e, _T("NOEMBED") };
CAssoc g_tagascNOFRAMES70			= {    70, 0x5f1d0d7c, _T("NOFRAMES") };
CAssoc g_tagascNOSCRIPT72			= {    72, 0x150d948f, _T("NOSCRIPT") };
CAssoc g_tagascOBJECT74			= {    74, 0x6716547e, _T("OBJECT") };
CAssoc g_tagascOL75			= {    75, 0x9e00004c, _T("OL") };
CAssoc g_tagascOPTION76			= {    76, 0x7f26a558, _T("OPTION") };
CAssoc g_tagascP77			= {    77, 0x00000050, _T("P") };
CAssoc g_tagascPARAM78			= {    78, 0x834a0d4d, _T("PARAM") };
CAssoc g_tagascPLAINTEXT79			= {    79, 0xe17cf53d, _T("PLAINTEXT") };
CAssoc g_tagascPRE80			= {    80, 0xa5400045, _T("PRE") };
CAssoc g_tagascQ81			= {    81, 0x00000051, _T("Q") };
CAssoc g_tagascRP83			= {    83, 0xa4000050, _T("RP") };
CAssoc g_tagascRT84			= {    84, 0xa4000054, _T("RT") };
CAssoc g_tagascRUBY85			= {    85, 0x85569059, _T("RUBY") };
CAssoc g_tagascS86			= {    86, 0x00000053, _T("S") };
CAssoc g_tagascSAMP87			= {    87, 0x9b069850, _T("SAMP") };
CAssoc g_tagascSCRIPT88			= {    88, 0x0126948f, _T("SCRIPT") };
CAssoc g_tagascSELECT89			= {    89, 0xe71664ae, _T("SELECT") };
CAssoc g_tagascSMALL90			= {    90, 0x99066d7c, _T("SMALL") };
CAssoc g_tagascSPAN91			= {    91, 0x8342984e, _T("SPAN") };
CAssoc g_tagascSTRIKE92			= {    92, 0xf726958f, _T("STRIKE") };
CAssoc g_tagascSTRONG93			= {    93, 0xfd3e9591, _T("STRONG") };
CAssoc g_tagascSTYLE94			= {    94, 0x9966a575, _T("STYLE") };
CAssoc g_tagascSUB95			= {    95, 0xab4c0042, _T("SUB") };
CAssoc g_tagascSUP96			= {    96, 0xab4c0050, _T("SUP") };
CAssoc g_tagascTABLE97			= {    97, 0x990a0d85, _T("TABLE") };
CAssoc g_tagascTBODY98			= {    98, 0x893e1599, _T("TBODY") };
CAssoc g_tagascTD100			= {   100, 0xa8000044, _T("TD") };
CAssoc g_tagascTEXTAREA101			= {   101, 0x9cb40d8c, _T("TEXTAREA") };
CAssoc g_tagascTFOOT102			= {   102, 0x9f3e3594, _T("TFOOT") };
CAssoc g_tagascTH103			= {   103, 0xa8000048, _T("TH") };
CAssoc g_tagascTHEAD104			= {   104, 0x83164584, _T("THEAD") };
CAssoc g_tagascTITLE105			= {   105, 0x99524d85, _T("TITLE") };
CAssoc g_tagascTR106			= {   106, 0xa8000052, _T("TR") };
CAssoc g_tagascTT107			= {   107, 0xa8000054, _T("TT") };
CAssoc g_tagascU108			= {   108, 0x00000055, _T("U") };
CAssoc g_tagascUL109			= {   109, 0xaa00004c, _T("UL") };
CAssoc g_tagascVAR110			= {   110, 0x83580052, _T("VAR") };
CAssoc g_tagascWBR111			= {   111, 0x855c0052, _T("WBR") };
CAssoc g_tagascXMP112			= {   112, 0x9b600050, _T("XMP") };
CAssoc g_tagascifdef127			= {   127, 0xeb1234d6, _T("#ifdef") };
CAssoc g_tagascendif130			= {   130, 0xf3127496, _T("#endif") };
CAssoc g_tagascIMAGE51			= {    51, 0x8f066cd5, _T("IMAGE") };

CAssoc* g_tagasc_HashTable[227] =
{
    /* 1 */ &g_tagascBASEFONT9,
    /* 1 */ &g_tagascTR106,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 2 */ &g_tagascifdef127,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascLISTING61,
    NULL,
    /* 2 */ &g_tagascWBR111,
    NULL,
    /* 1 */ &g_tagascIFRAME50,
    /* 0 */ &g_tagascCOMMENT24,
    NULL,
    /* 0 */ &g_tagascPLAINTEXT79,
    NULL,
    /* 0 */ &g_tagascNEXTID66,
    /* 0 */ &g_tagascPRE80,
    NULL,
    NULL,
    /* 1 */ &g_tagascCODE21,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascEM33,
    /* 0 */ &g_tagascISINDEX55,
    NULL,
    /* 1 */ &g_tagascBR16,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascOBJECT74,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascSUB95,
    /* 0 */ &g_tagascBLINK13,
    /* 0 */ &g_tagascOPTION76,
    /* 2 */ &g_tagascRUBY85,
    NULL,
    /* 0 */ &g_tagascDD25,
    /* 0 */ &g_tagascendif130,
    /* 3 */ &g_tagascBASE8,
    /* 0 */ &g_tagascUL109,
    NULL,
    NULL,
    NULL,
    /* 1 */ &g_tagascRP83,
    /* 0 */ &g_tagascDL31,
    /* 0 */ &g_tagascSUP96,
    NULL,
    /* 0 */ &g_tagascSTRIKE92,
    NULL,
    /* 0 */ &g_tagascNOEMBED68,
    NULL,
    NULL,
    /* 0 */ &g_tagascDT32,
    /* 0 */ &g_tagascA2,
    /* 0 */ &g_tagascB7,
    NULL,
    /* 0 */ &g_tagascXMP112,
    NULL,
    /* 0 */ &g_tagascTD100,
    NULL,
    NULL,
    /* 0 */ &g_tagascI49,
    /* 0 */ &g_tagascTH103,
    /* 0 */ &g_tagascMENU64,
    NULL,
    /* 0 */ &g_tagascBDO10,
    NULL,
    NULL,
    /* 0 */ &g_tagascP77,
    /* 0 */ &g_tagascQ81,
    NULL,
    /* 0 */ &g_tagascFONT36,
    /* 1 */ &g_tagascSTYLE94,
    /* 0 */ &g_tagascU108,
    /* 0 */ &g_tagascTT107,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascDIR29,
    NULL,
    NULL,
    /* 0 */ &g_tagascTBODY98,
    /* 0 */ &g_tagascDIV30,
    NULL,
    /* 1 */ &g_tagascCOLGROUP23,
    NULL,
    /* 1 */ &g_tagascH544,
    NULL,
    /* 0 */ &g_tagascADDRESS4,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascBGSOUND11,
    NULL,
    NULL,
    /* 0 */ &g_tagascFIELDSET35,
    /* 0 */ &g_tagascBODY15,
    NULL,
    /* 0 */ &g_tagascNOFRAMES70,
    /* 2 */ &g_tagascBUTTON17,
    NULL,
    /* 0 */ &g_tagascH140,
    /* 0 */ &g_tagascH241,
    /* 0 */ &g_tagascH342,
    /* 0 */ &g_tagascH443,
    /* 0 */ &g_tagascMETA65,
    /* 0 */ &g_tagascH645,
    NULL,
    /* 0 */ &g_tagascLINK60,
    NULL,
    /* 1 */ &g_tagascSMALL90,
    NULL,
    /* 1 */ &g_tagascAPPLET5,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 1 */ &g_tagascIMAGE51,
    NULL,
    /* 0 */ &g_tagascHTML48,
    /* 1 */ &g_tagascTITLE105,
    /* 1 */ &g_tagascKBD56,
    /* 0 */ &g_tagascCITE20,
    NULL,
    /* 0 */ &g_tagascTHEAD104,
    /* 0 */ &g_tagascRT84,
    /* 2 */ &g_tagascSCRIPT88,
    /* 2 */ &g_tagascS86,
    NULL,
    /* 4 */ &g_tagascVAR110,
    NULL,
    /* 0 */ &g_tagascLEGEND58,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 1 */ &g_tagascTABLE97,
    /* 0 */ &g_tagascINPUT52,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascHEAD46,
    /* 0 */ &g_tagascBLOCKQUOTE14,
    /* 0 */ &g_tagascAREA6,
    NULL,
    /* 0 */ &g_tagascBIG12,
    /* 0 */ &g_tagascMAP62,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascEMBED34,
    /* 0 */ &g_tagascDFN28,
    NULL,
    NULL,
    NULL,
    /* 1 */ &g_tagascFRAMESET39,
    NULL,
    /* 0 */ &g_tagascCAPTION18,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascHR47,
    NULL,
    NULL,
    /* 0 */ &g_tagascSELECT89,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascSTRONG93,
    NULL,
    /* 0 */ &g_tagascCOL22,
    NULL,
    /* 0 */ &g_tagascDEL27,
    NULL,
    NULL,
    /* 0 */ &g_tagascFORM37,
    NULL,
    /* 0 */ &g_tagascPARAM78,
    /* 0 */ &g_tagascTEXTAREA101,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascCENTER19,
    NULL,
    /* 0 */ &g_tagascACRONYM3,
    /* 0 */ &g_tagascINS54,
    NULL,
    NULL,
    NULL,
    /* 0 */ &g_tagascMARQUEE63,
    /* 1 */ &g_tagascSAMP87,
    NULL,
    NULL,
    /* 0 */ &g_tagascOL75,
    NULL,
    NULL,
    /* 0 */ &g_tagascNOSCRIPT72,
    /* 1 */ &g_tagascLI59,
    /* 0 */ &g_tagascNOBR67,
    NULL,
    /* 0 */ &g_tagascIMG51,
    NULL,
    /* 1 */ &g_tagascFRAME38,
    /* 2 */ &g_tagascTFOOT102,
    NULL,
    /* 0 */ &g_tagascSPAN91,
    /* 0 */ &g_tagascLABEL57,
};

CAssocArray g_tagasc =
{
    g_tagasc_HashTable,
    107,
    227,
    127,
    107,
    5,
    TRUE,
};
CPtrBagCi<ELEMENT_TAG> g_bKnownTags(&g_tagasc);


ELEMENT_TAG s_atagNull[] = {ETAG_NULL};
ELEMENT_TAG s_atagHtml[] = {ETAG_HTML, ETAG_NULL};
ELEMENT_TAG s_atagHead[] = {ETAG_HEAD, ETAG_NULL};
ELEMENT_TAG s_atagBody[] = {ETAG_BODY, ETAG_NULL};
ELEMENT_TAG s_atagBodyGeneric[] = {ETAG_BODY, ETAG_GENERIC, ETAG_NULL};
ELEMENT_TAG s_atagBodyHead[] = {ETAG_BODY, ETAG_HEAD, ETAG_NULL};
ELEMENT_TAG s_atagOverlapBoundary[] = {ETAG_BUTTON, ETAG_CAPTION, ETAG_HTML, ETAG_MARQUEE, /* ETAG_HTMLAREA, */ ETAG_TD, ETAG_TH, ETAG_NULL};
ELEMENT_TAG s_atagOverlapRequired[] = {ETAG_BODY, ETAG_BUTTON, ETAG_CAPTION, ETAG_GENERIC, ETAG_MARQUEE, /* ETAG_HTMLAREA, */ ETAG_TD, ETAG_TH, ETAG_NULL};
ELEMENT_TAG s_atagPBoundary[] = {ETAG_BLOCKQUOTE, ETAG_BODY, ETAG_BUTTON, ETAG_CAPTION, ETAG_DIR, ETAG_DL, ETAG_LISTING, ETAG_MARQUEE, ETAG_MENU, ETAG_OL, ETAG_PRE, ETAG_TABLE, ETAG_TBODY, ETAG_TC, ETAG_TD, ETAG_TFOOT, ETAG_TH, ETAG_THEAD, ETAG_TR, ETAG_UL, ETAG_NULL};
ELEMENT_TAG s_atagPBoundaryGeneric[] = {ETAG_BLOCKQUOTE, ETAG_BODY, ETAG_BUTTON, ETAG_CAPTION, ETAG_DIR, ETAG_DL, ETAG_GENERIC, ETAG_LISTING, ETAG_MARQUEE, ETAG_MENU, ETAG_OL, ETAG_PRE, ETAG_TABLE, ETAG_TBODY, ETAG_TC, ETAG_TD, ETAG_TFOOT, ETAG_TH, ETAG_THEAD, ETAG_TR, ETAG_UL, ETAG_NULL};
ELEMENT_TAG s_atagNestBoundary[] = {ETAG_BODY, ETAG_BUTTON, ETAG_CAPTION, ETAG_HEAD, ETAG_MARQUEE, /* ETAG_HTMLAREA, */ ETAG_TABLE, ETAG_TBODY, ETAG_TC, ETAG_TD, ETAG_TFOOT, ETAG_TH, ETAG_THEAD, ETAG_TR, ETAG_NULL};
ELEMENT_TAG s_atagButton[] = {ETAG_BUTTON, ETAG_NULL};
ELEMENT_TAG s_atagA[] = {ETAG_A, ETAG_NULL};
ELEMENT_TAG s_atagP[] = {ETAG_P, ETAG_NULL};
ELEMENT_TAG s_atagLiP[] = {ETAG_LI, ETAG_P, ETAG_NULL};
ELEMENT_TAG s_atagDdDtP[] = {ETAG_DD, ETAG_DT, ETAG_P, ETAG_NULL};
ELEMENT_TAG s_atagSelect[] = {ETAG_SELECT, ETAG_NULL};
ELEMENT_TAG s_atagOption[] = {ETAG_OPTION, ETAG_NULL};
ELEMENT_TAG s_atagIframe[] = {ETAG_IFRAME, ETAG_NULL};
ELEMENT_TAG s_atagForm[] = {ETAG_FORM, ETAG_NULL};
ELEMENT_TAG s_atagAppletObject[] = {ETAG_APPLET, ETAG_OBJECT, ETAG_NULL};
ELEMENT_TAG s_atagEOFProhibited[] = {ETAG_APPLET, ETAG_COMMENT, ETAG_OBJECT, ETAG_SCRIPT, ETAG_STYLE, ETAG_TITLE_ELEMENT, ETAG_GENERIC_LITERAL, ETAG_NOSCRIPT, ETAG_NOFRAMES, ETAG_NOEMBED, ETAG_IFRAME, ETAG_NULL };

static CHtmlParseClass s_hpcBody =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagHtml,                     // _atagEndContainers           ; BODY is strictly contained by HTML
    s_atagHtml,                     // _atagBeginContainers         ; BODY tag parse does not search above containing HTML
    s_atagBody,                     // _atagMaskingContainers
    s_atagHead,                     // _atagProhibitedContainers    ; closes any open HEAD
    s_atagHtml,                     // _atagRequiredContainers      ; BODY tag requires an HTML tag
    ETAG_HTML,                      // _etagDefaultContainer        ; BODY implies HTML if one isn't present
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope                   ; BODY can contain text
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,//CreateHtmBodyParseCtx,          // _pfnHpxCreator
    TRUE,                           // _fMerge                      ; BODY tag in a BODY is merged
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// BGSOUND, LINK, META, NEXTID, PARAM
static CHtmlParseClass s_hpcEmpty =
{
    SCOPE_EMPTY,                    // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagNull,                     // _atagEndContainers           ; end tags are irrelevant
    s_atagHtml,                     // _atagBeginContainers         ; when requiring BODY/HEAD, stop searching at HTML
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD
    ETAG_HEAD,                      // _etagDefaultContainer        ; if neither is present, imply a HEAD
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// BR, EMBED, IMG, INPUTTXT, ISINDEX, WBR
static CHtmlParseClass s_hpcEmptyText =
{
    SCOPE_EMPTY,                    // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; these empty tags are textlike
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers         ; when requiring BODY/HEAD, stop searching at HTML
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD (note that "textlike" will normally ensure BODY already)
    ETAG_BODY,                      // _etagDefaultContainer        ; if neither is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// BUTTON:
// BUTTONs are text edits...
ELEMENT_TAG s_atagButtonProhibitedContainers[] = {ETAG_BUTTON, ETAG_NULL};

static CHtmlParseClass s_hpcButton =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNestBoundary,             // _atagEndContainers
    s_atagNestBoundary,             // _atagBeginContainers         ; TABLE-like container boundaries
    NULL,                           // _atagMaskingContainers
    s_atagButton,                   // _atagProhibitedContainers    ; close previously open BUTTON
    NULL,                           // _atagRequiredContainers
    ETAG_NULL,                      // _etagDefaultContainer
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope                   ; exclude text
    ETAG_NULL,                      // _etagTextSubcontainer        ; wrap contained text in a TC
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// ACRONYM, B, BDO, BIG, BLINK, CITE, CODE, DEL, DFN, EM, FONT,
// I, INS, KBD, LABEL, Q, S, SAMP, SMALL, STRIKE, STRONG, SUB,
// SUP, TT, U, VAR
static CHtmlParseClass s_hpcOverlap =
{
    SCOPE_OVERLAP,                  // _scope
    TEXTTYPE_QUERY,                 // _texttype
    s_atagOverlapBoundary,          // _atagEndContainers           ; don't match end tags beyond TD-like boundaries
    s_atagOverlapBoundary,          // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagOverlapRequired,          // _atagRequiredContainers      ; must appear inside a BODY or TD-like container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, queue for a BODY
    TRUE,                           // _fQueueForRequired           ; queue
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// FORM:
// FORMs can overlap anything, but they mask themselves
static CHtmlParseClass s_hpcForm =
{
    SCOPE_OVERLAP,                  // _scope
    TEXTTYPE_QUERY,                 // _texttype                    ; forms are text-like when pasting
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    s_atagForm,                     // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD
    ETAG_HEAD,                      // _etagDefaultContainer        ; if neither is present, imply a HEAD
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

static CHtmlParseClass s_hpcHr =
{
    SCOPE_EMPTY,                    // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNull,                     // _atagEndContainers           ; end tags are irrelevant
    s_atagPBoundary,                // _atagBeginContainers         ; when closing Ps stop searching at TABLE-like boundaries
    NULL,                           // _atagMaskingContainers
    s_atagP,                        // _atagProhibitedContainers    ; closes P
    s_atagPBoundary,                // _atagRequiredContainers      ; must appear inside a BODY or TABLE-like container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// INPUT:
// INPUTs are textbox, buttons, checkbox ...
static CHtmlParseClass s_hpcInput =
{
    SCOPE_EMPTY,                    // _scope
    TEXTTYPE_QUERY,                 // _texttype                    ; textlike only if not hidden; see CHtmTopParseCtx::QueryTextlike
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers         ; when requiring BODY/HEAD, stop searching at HTML
    s_atagButton,                   // _atagMaskingContainers       ; hidden by BUTTON
    s_atagSelect,                   // _atagProhibitedContainers    ; close previously open SELECT
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD (note that "textlike" will normally ensure BODY already)
    ETAG_HEAD,                      // _etagDefaultContainer        ; if neither is present, imply a HEAD
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// LISTS:
// LIs close LIs; DDs and DTs close DDs and DTs
static CHtmlParseClass s_hpcLI =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNestBoundary,             // _atagEndContainers           ; close at TABLE-like boundaries
    s_atagPBoundary,                // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagLiP,                      // _atagProhibitedContainers    ; close any previously open LI or P
    s_atagPBoundary,                // _atagRequiredContainers      ; must appear inside a BODY or P-boundary container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// DD, DT
static CHtmlParseClass s_hpcDDDT =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNestBoundary,             // _atagEndContainers           ; close at TABLE-like boundaries
    s_atagPBoundary,                // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagDdDtP,                    // _atagProhibitedContainers    ; close any previously open DD, DT, or P
    s_atagPBoundary,                // _atagRequiredContainers      ; must appear inside a BODY or P-boundary container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// SELECT:
// SELECTs specify a SELECT context
static CHtmlParseClass s_hpcSelect =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNull,                     // _atagEndContainers           ; irrelevant
    s_atagHtml,                     // _atagBeginContainers
    s_atagButton,                   // _atagMaskingContainers
    s_atagSelect,                   // _atagProhibitedContainers    ; close any previously open SELECT
    s_atagBody,                     // _atagRequiredContainers      ; must appear inside a BODY
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL/*CreateHtmSelectParseCtx wlw note*/,        // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};


// OPTION:
// OPTIONs accept text, and should only appear inside a SELECT
static CHtmlParseClass s_hpcOption =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagSelect,                   // _atagEndContainers           ; if a SELECT container ends, the OPTION ends
    s_atagSelect,                   // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagOption,                   // _atagProhibitedContainers    ; close any previously open SELECT
    s_atagSelect,                   // _atagRequiredContainers      ; must appear inside a SELECT
    ETAG_NULL,                      // _etagDefaultContainer        ; if none is present, unknownify
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// OBJECT/APPLET:
// OBJECTs and APPLETs have their own context which accepts params
static CHtmlParseClass s_hpcObjectApplet =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_QUERY,                 // _texttype                    ; textlike unless in HEAD in cases; see CHtmHeadParseCtx
    s_atagHtml,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagAppletObject,             // _atagProhibitedContainers    ; close any previously open APPLET or OBJECT
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD (note textlike will usually imply BODY)
    ETAG_HEAD,                      // _etagDefaultContainer        ; if none is present, imply HEAD
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL/*CreateHtmObjectParseCtx wlw note*/,        // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// Paragraph tags (P, ADDRESS,...) end <P>s unless nested inside <UL>s etc
// An unmatched close tag implies a begin tag right before it
// ADDRESS, BLOCKQUOTE, DIR, FIELDSET, P, FIELDSET
static CHtmlParseClass s_hpcParagraph =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_QUERY,                 // _texttype                    ; textlike; force BODY etc (now forced by _atagRequired
    s_atagNestBoundary,             // _atagEndContainers           ; close at TABLE-like boundaries
    s_atagPBoundary,                // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagP,                        // _atagProhibitedContainers    ; close any previously open P
    s_atagPBoundaryGeneric,         // _atagRequiredContainers      ; must appear inside a BODY or P boundary container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply a BODY
    TRUE,                           // _fQueueForRequired           ; yes, queue when in HEAD
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_IMPLICIT_BEGIN,            // _etagUnmatchedSubstitute     ; if only the end tag appears, imply the begin tag
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// CENTER tag:
// Closes paragraphs (see IE4 bug 29566) been requested that it shouldn't (IE5 bug 1379) but we haven't changed the behavior
// if we have an unmatched end CENTER, add an implicit begin CENTER
static CHtmlParseClass s_hpcCenter =
{
    SCOPE_OVERLAP,                  // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagOverlapBoundary,          // _atagEndContainers           ; close at TD-like boundaries
    s_atagPBoundary,                // _atagBeginContainers         ; prohibited+required container is P-like
    NULL,                           // _atagMaskingContainers
    s_atagP,                        // _atagProhibitedContainers    ; closes P
    s_atagPBoundaryGeneric,         // _atagRequiredContainers      ; match _atagBeginContainers
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, queue for BODY
    TRUE,                           // _fQueueForRequired           ; yes, queue
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_IMPLICIT_BEGIN,            // _etagUnmatchedSubstitute     ; if only the end tag appears, imply the begin tag
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// Hx:
// Header tags are similar to paragraph tags, but they have overlapping scope.
// H1, H2, H3, H4, H5, H6
ELEMENT_TAG s_atagHeaders[] = {ETAG_H1, ETAG_H2, ETAG_H3, ETAG_H4, ETAG_H5, ETAG_H6, ETAG_NULL};

static CHtmlParseClass s_hpcHeader =
{
    SCOPE_OVERLAP,                  // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagOverlapBoundary,          // _atagEndContainers           ; close at TD-like boundaries
    s_atagPBoundary,                // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    s_atagP,                        // _atagProhibitedContainers    ; close any previously open P
    s_atagPBoundaryGeneric,         // _atagRequiredContainers      ; must appear inside a BODY or P boundary container
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, queue for a BODY
    TRUE,                           // _fQueueForRequired           ; yes, queue
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    s_atagHeaders,                  // _atagMatch                   ; any </Hx> can match any other <Hy>
    ETAG_IMPLICIT_BEGIN,            // _etagUnmatchedSubstitute     ; if only the end tag appears, imply the begin tag
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// The ROOT element is never created - it exists in every tree.
// It excludes text, and implies a BODY element if text is injected.
static CHtmlParseClass s_hpcRoot =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagNull,                     // _atagEndContainers
    s_atagNull,                     // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    NULL,                           // _atagRequiredContainers
    ETAG_NULL,                      // _etagDefaultContainer
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_EXCLUDE,              // _textscope
    ETAG_BODY,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// The HTML element can be explicitly or implicitly created (it is required by the HEAD and BODY).
// It immediately implies the HEAD subcontainer
// HTML excludes text (text implies a BODY)
static CHtmlParseClass s_hpcHtml =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagHtml,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    s_atagHtml,                     // _atagMaskingContainers       ; HTML in another HTML is masked
    NULL,                           // _atagProhibitedContainers
    NULL,                           // _atagRequiredContainers
    ETAG_NULL,                      // _etagDefaultContainer
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_EXCLUDE,              // _textscope                   ; exclude text
    ETAG_BODY,                      // _etagTextSubcontainer        ; if text appears inside, it implies a BODY
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    TRUE,                           // _fMerge                      ; IE5 bug 43801: merge HTML tags
    ETAG_HEAD,                      // _etagImplicitChild           ; HTML always has a HEAD child
    FALSE,                          // _fCloseImplicitChild         ; the child is left open
};

// MARQUEE:
static CHtmlParseClass s_hpcMarquee =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagHtml,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBody,                     // _atagRequiredContainers      ; must appear inside a BODY
    ETAG_BODY,                      // _etagDefaultContainer        ; if none is present, imply BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// STYLE:
// STYLEs are literal, and specify a STYLE context
static CHtmlParseClass s_hpcStyle =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_NEVER,                 // _texttype
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD
    ETAG_HEAD,                      // _etagDefaultContainer        ; if neither is present, imply a HEAD
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope                   ; able to contain text
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL/*CreateHtmStyleParseCtx wlw note*/,         // _pfnHpxCreator               ; create a StyleParseCtx
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// TEXTAREA:
// TEXTAREAs are literal, and specify a TEXTAREA context
static CHtmlParseClass s_hpcTextarea =
{
    SCOPE_NESTED,                   // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers
    s_atagButton,                   // _atagMaskingContainers       ; hidden by BUTTON container
    s_atagSelect,                   // _atagProhibitedContainers    ; close previously open SELECT
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD
    ETAG_BODY,                      // _etagDefaultContainer        ; if neither is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_INCLUDE,              // _textscope                   ; able to contain text
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL/*CreateHtmPreParseCtx wlw note*/,           // _pfnHpxCreator               ; create a StyleParseCtx
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// HYPERLINKS:
// A tags end <A> tags
static CHtmlParseClass s_hpcAnchor =
{
    SCOPE_OVERLAP,                  // _scope
    TEXTTYPE_QUERY,                 // _texttype
    s_atagOverlapBoundary,          // _atagEndContainers           ; close at TD-like boundaries
    s_atagOverlapBoundary,          // _atagBeginContainers
    s_atagButton,                   // _atagMaskingContainers       ; hide when inside BUTTON
    s_atagA,                        // _atagProhibitedContainers    ; close any previously open A
    s_atagOverlapRequired,          // _atagRequiredContainers      ; must appear inside a BODY or HEAD
    ETAG_BODY,                      // _etagDefaultContainer        ; if neither is present, queue for a BODY
    TRUE,                           // _fQueueForRequired           ; queue
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_NULL,                      // _etagUnmatchedSubstitute
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

// BR:
// end-BR tags become BR tags
static CHtmlParseClass s_hpcBr =
{
    SCOPE_EMPTY,                    // _scope
    TEXTTYPE_ALWAYS,                // _texttype                    ; textlike; force BODY etc
    s_atagNull,                     // _atagEndContainers
    s_atagHtml,                     // _atagBeginContainers         ; when requiring BODY/HEAD, stop searching at HTML
    NULL,                           // _atagMaskingContainers
    NULL,                           // _atagProhibitedContainers
    s_atagBodyHead,                 // _atagRequiredContainers      ; must appear inside a BODY or HEAD (note that "textlike" will normally ensure BODY already)
    ETAG_BODY,                      // _etagDefaultContainer        ; if neither is present, imply a BODY
    FALSE,                          // _fQueueForRequired
    TEXTSCOPE_NEUTRAL,              // _textscope
    ETAG_NULL,                      // _etagTextSubcontainer
    NULL,                           // _atagMatch
    ETAG_IMPLICIT_BEGIN,            // _etagUnmatchedSubstitute     ; </BR> becomes <BR>
    NULL,                           // _pfnHpxCreator
    FALSE,                          // _fMerge
    ETAG_NULL,                      // _etagImplicitChild
    FALSE,                          // _fCloseImplicitChild
};

const CTagDesc g_atagdesc[] =
{
    {}, {},
    {},
    {},
    { SZTAG_ADDRESS,        &s_hpcParagraph,    CBlockElement::CreateElement,
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_BLKSTYLEDD              },
    {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, 
    { SZTAG_BODY,           &s_hpcBody,         CBodyElement::CreateElement,
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_SAVEALWAYSEND           |
    TAGDESC_ACCEPTHTML              |
    TAGDESC_SPECIALTOKEN            |
    TAGDESC_CONTAINER               |
    TAGDESC_SLOWPROCESS             },
    {},
    {},
    {},
    { SZTAG_CENTER,         &s_hpcCenter,       CBlockElement::CreateElement,
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         },
    {}, {}, {}, {}, {},
    {},
    {}, {}, {}, {},
    { SZTAG_DIV,            &s_hpcParagraph,    CDivElement::CreateElement,
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_ALIGN                   |
    TAGDESC_SPLITBLOCKINLIST        },
    {},
    {},
    {},
    {},
    {},
    {},
    {},
    {}, {},
    {},
    {},
    {},
    {},
    {},
    {},
    {},
    {},
    {},
    {}, {},
    { SZTAG_IMG,            &s_hpcEmptyText,    CImgElement::CreateElement,
    TAGDESC_TEXTLESS                |
    TAGDESC_SPECIALTOKEN            },
    {},
    { _T(""),/*TXTSLAVE*/   &s_hpcEmpty,        CTxtSlave::CreateElement,
    TAGDESC_LITERALTAG              |
    TAGDESC_SAVEALWAYSEND           |
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_SAVEINDENT              |
    TAGDESC_XSMBEGIN                |
    TAGDESC_LOGICALINVISUAL         |
    TAGDESC_CONTAINER               },
    {}, {}, {},
    {},
    {}, 
    {},
    {},
    {}, {},
    {},
    {}, {}, {}, {}, {}, {}, {}, {}, {}, {},
    {},
    {}, 
    {},
    { SZTAG_P,              &s_hpcParagraph,    CParaElement::CreateElement,
    TAGDESC_BLOCKELEMENT            |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_BLKSTYLEDD              |
    TAGDESC_ALIGN                   |
    TAGDESC_SAVENBSPIFEMPTY         },
    {}, {}, {}, {},
    {},
    {}, {}, {}, {}, {}, {},
    {},
    {},
    {},
    {},
    {},
    {},
    {}, {}, {}, {}, {}, {},
    { SZTAG_TEXTAREA,       &s_hpcTextarea,     CTextArea::CreateElement,
    TAGDESC_LITERALTAG              |
    TAGDESC_SAVEALWAYSEND           |
    TAGDESC_SAVENBSPIFEMPTY         |
    TAGDESC_SAVEINDENT              |
    TAGDESC_LOGICALINVISUAL         |
    TAGDESC_CONTAINER               },
    {}, {}, {}, {}, {}, {}, {},
    {},
};


//+------------------------------------------------------------------------
//
//  Member:     CHtmTag::AttrFromName
//
//  Synopsis:   name->CAttr*
//
//-------------------------------------------------------------------------
CHtmTag::CAttr* CHtmTag::AttrFromName(const TCHAR* pchName)
{
    Assert(pchName);

    int i = _cAttr;

    // optimize for zero attrs
    if(i && pchName)
    {
        CAttr* pattr = _aAttr;
        for(; i--; pattr++)
        {
            if(!StrCmpIC(pattr->_pchName, pchName))
            {
                return pattr;
            }
        }
    }
    return NULL;
}

//+------------------------------------------------------------------------
//
//  Member:     CHtmTag::ValFromName
//
//  Synopsis:   name->val
//              Returns TRUE if attribute named by pchName is present.
//              Returns pointer to value string in *ppchVal (NULL if no
//              value is present)
//
//-------------------------------------------------------------------------
BOOL CHtmTag::ValFromName(const TCHAR* pchName, TCHAR** ppchVal)
{
    CAttr* pattr = AttrFromName(pchName);
    if(pattr)
    {
        *ppchVal = pattr->_pchVal;
        return TRUE;
    }
    else
    {
        *ppchVal = NULL;
        return FALSE;
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CHtmTag::GetXmlNamespace
//
//  Synopsis:   enumerator of "xmlns:foo" declarations
//
//-------------------------------------------------------------------------
LPTSTR CHtmTag::GetXmlNamespace(int* pIdx)
{
    Assert(pIdx);

    if(_cAttr)
    {
        CAttr* pAttr;

        for(pAttr=&(_aAttr[*pIdx]); (*pIdx)<_cAttr; (*pIdx)++,pAttr++)
        {
            if(_tcsnipre(_T("xmlns:"), 6, pAttr->_pchName, -1))
            {
                return pAttr->_pchName + 6;
            }
        }
    }
    return NULL;
}



//+------------------------------------------------------------------------
//
//  Member:     TagProhibitedContainer(ELEMENT_TAG tag1, tag2)
//
//  Synopsis:   returns TRUE if tag1 is an end container of tag2.
//
//  Arguments:  tag1, tag2 (ELEMENT_TAG)
//
//  Returns:    BOOL
//
//-------------------------------------------------------------------------
BOOL TagEndContainer(ELEMENT_TAG tag1, ELEMENT_TAG tag2)
{
    CHtmlParseClass* phpc1 = HpcFromEtag(tag1);

    return (phpc1 && phpc1->_atagEndContainers &&
        IsEtagInSet(tag2, phpc1->_atagEndContainers));
}

//+------------------------------------------------------------------------
//
//  Member:     TagHasNoEndTag(ELEMENT_TAG tag)
//
//  Synopsis:   returns TRUE if tag should not have an end tag
//
//  Arguments:  tag (ELEMENT_TAG)
//
//  Returns:    BOOL
//
//-------------------------------------------------------------------------
BOOL TagHasNoEndTag(ELEMENT_TAG tag)
{
    CHtmlParseClass* phpc = HpcFromEtag(tag);
    const CTagDesc* ptd = TagDescFromEtag(tag);

    return  (!phpc || phpc->_scope==SCOPE_EMPTY 
        || !ptd || ptd->HasFlag(TAGDESC_NEVERSAVEEND));
}


//+------------------------------------------------------------------------
//
//  Member:     CreateElement
//
//  Synopsis:   Creates an element of type etag parented to pElementParent
//
//              Inits an empty AttrBag if asked to.
//
//-------------------------------------------------------------------------
HRESULT CreateElement(
          CHtmTag*      pht,
          CElement**    ppElementResult,
          CDocument*    pDoc,
          CMarkup*      pMarkup,
          BOOL*         pfDie,
          DWORD         dwFlags)
{
    CElement*       pElement = NULL;
    const CTagDesc* ptd;
    HRESULT         hr;

    if(!pfDie)
    {
        pfDie = (BOOL*)&_afxGlobalData._Zero;
    }

    ptd = TagDescFromEtag(pht->GetTag());
    if(!ptd)
    {
        return E_FAIL;
    }

    hr = ptd->_pfnElementCreator(pht, pDoc, &pElement);
    if(hr)
    {
        goto Cleanup;
    }

    if(*pfDie)
    {
        goto Die;
    }

    hr = pElement->Init();
    if(hr)
    {
        goto Cleanup;
    }

    if(*pfDie)
    {
        goto Die;
    }

    hr = pElement->InitAttrBag(pht);
    if(hr)
    {
        goto Cleanup;
    }

    if(*pfDie)
    {
        goto Die;
    }

    {
        CElement::CInit2Context context(pht, pMarkup, dwFlags);

        hr = pElement->Init2(&context);
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(*pfDie)
    {
        goto Die;
    }

Cleanup:
    if(hr && pElement)
    {
        CElement::ClearPtr(&pElement);
    }

    *ppElementResult = pElement;
    RRETURN(hr);

Die:
    hr = E_ABORT;
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Function:   ScanNodeList
//
//  Synopsis:   Like FindContainer; used by the ValidateNodeList function
//
//-------------------------------------------------------------------------
CTreeNode** ScanNodeList(
                         CTreeNode**    apNodeStack,
                         long           cNodeStack,
                         ELEMENT_TAG*   pSetMatch,
                         ELEMENT_TAG*   pSetStop)
{
    CTreeNode** ppNode;
    long c;

    for(ppNode=apNodeStack+cNodeStack-1,c=cNodeStack; c; ppNode-=1,c-=1)
    {
        if(IsEtagInSet((*ppNode)->Tag(), pSetMatch))
        {
            return ppNode;
        }

        if(IsEtagInSet((*ppNode)->Tag(), pSetStop))
        {
            return NULL;
        }
    }

    return NULL;
}

//+------------------------------------------------------------------------
//
//  Function:   MergableTags
//
//  Synopsis:   Returns TRUE if tags can be merged with MergeTag
//
//-------------------------------------------------------------------------
ELEMENT_TAG s_atagTitles[] = { ETAG_TITLE_ELEMENT, ETAG_TITLE_TAG, ETAG_NULL };
ELEMENT_TAG* s_aatagMergable[] = { s_atagTitles, NULL };

BOOL MergableTags(ELEMENT_TAG etag1, ELEMENT_TAG etag2)
{
    if(etag1 == etag2)
    {
        return TRUE;
    }

    ELEMENT_TAG** patag;

    for(patag = s_aatagMergable; *patag; patag += 1)
    {
        if(IsEtagInSet(etag1, *patag) && IsEtagInSet(etag2, *patag))
        {
            return TRUE;
        }
    }

    return FALSE;
}