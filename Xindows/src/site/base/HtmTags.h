
#ifndef __XINDOWS_SITE_BASE_HTMLTAGS_H__
#define __XINDOWS_SITE_BASE_HTMLTAGS_H__

#include "../gen/entity.h"

class CHtmTagStm;

class CHtmTag
{
public:
    // Sub-types
    struct CAttr
    {
        TCHAR*  _pchName;       // Pointer to attribute name (NULL terminated)
        ULONG   _cchName;       // Count of attribute name characters
        TCHAR*  _pchVal;        // Pointer to attribute value (NULL terminated)
        ULONG   _cchVal;        // Count of attribute value characters
        ULONG   _ulLine;        // Source line where attribute begins
        ULONG   _ulOffset;      // Source offset where attribute begins
    };

    // Methods

    void            Reset()                     { *(DWORD*)this = 0; }

    void            SetTag(ELEMENT_TAG etag)    { _etag = (BYTE)etag; }
    ELEMENT_TAG     GetTag()                    { return ((ELEMENT_TAG)_etag); }
    BOOL            Is(ELEMENT_TAG etag)        { return (etag == (ELEMENT_TAG)_etag); }
    BOOL            IsBegin(ELEMENT_TAG etag)   { return (Is(etag) && !IsEnd()); }
    BOOL            IsEnd(ELEMENT_TAG etag)     { return (Is(etag) && IsEnd()); }

    void            SetNextBuf()                { _bFlags |= TAGF_NEXTBUF; }
    BOOL            IsNextBuf()                 { return (!!(_bFlags&TAGF_NEXTBUF)); }
    void            SetRestart()                { _bFlags |= TAGF_RESTART; }
    BOOL            IsRestart()                 { return (!!(_bFlags&TAGF_RESTART)); }
    void            SetEmpty()                  { _bFlags |= TAGF_EMPTY; }
    BOOL            IsEmpty()                   { return (!!(_bFlags&TAGF_EMPTY)); }
    void            SetEnd()                    { _bFlags |= TAGF_END; }
    BOOL            IsEnd()                     { return (!!(_bFlags&TAGF_END)); }
    void            SetTiny()                   { _bFlags |= TAGF_TINY; }
    BOOL            IsTiny()                    { return (!!(_bFlags&TAGF_TINY)); }
    void            SetAscii()                  { _bFlags |= TAGF_ASCII; }
    BOOL            IsAscii()                   { return (!!(_bFlags&TAGF_ASCII)); }
    void            SetScriptCreated()          { _bFlags |= TAGF_SCRIPTED; }
    BOOL            IsScriptCreated()           { return (!!(_bFlags&TAGF_SCRIPTED)); }
    void            SetDefer()                  { _bFlags |= TAGF_DEFER; }
    BOOL            IsDefer()                   { return (!!(_bFlags&TAGF_DEFER)); }
    BOOL            IsSource()                  { return (GetTag()==ETAG_RAW_SOURCE); }

    TCHAR*          GetPch()                    { Assert(!IsTiny() && !IsSource()); return(_pch); }
    void            SetPch(TCHAR* pch)          { Assert(!IsTiny() && !IsSource()); _pch = pch; }
    ULONG           GetCch()                    { Assert(!IsTiny() && !IsSource()); return(_cch); }
    void            SetCch(ULONG cch)           { Assert(!IsTiny() && !IsSource()); _cch = cch; }
    CHtmTagStm*     GetHtmTagStm()              { Assert(!IsTiny() && IsSource()); return(_pHTS); }
    void            SetHtmTagStm(CHtmTagStm* p) { Assert(!IsTiny() && IsSource()); _pHTS = p; }
    ULONG           GetSourceCch()              { Assert(!IsTiny() && IsSource()); return(_cch); }
    void            SetSourceCch(ULONG cch)     { Assert(!IsTiny() && IsSource()); _cch = cch; }
    ULONG           GetLine()                   { Assert(!IsTiny()); return(_ulLine); }
    void            SetLine(ULONG ulLine)       { Assert(!IsTiny()); _ulLine = ulLine; }
    ULONG           GetOffset()                 { Assert(!IsTiny()); return(_ulOffset); }
    void            SetOffset(ULONG ulOffset)   { Assert(!IsTiny()); _ulOffset = ulOffset; }
    CODEPAGE        GetCodepage()               { Assert(!IsTiny()); return(_cp); }
    void            SetCodepage(CODEPAGE cp)    { Assert(!IsTiny()); _cp = cp; }
    ULONG           GetDocSize()                { Assert(!IsTiny()); return(_ulSize); }
    void            SetDocSize(ULONG ulSize)    { Assert(!IsTiny()); _ulSize = ulSize; }

    UINT            GetAttrCount()              { return(_cAttr); }
    void            SetAttrCount(UINT i)        { _cAttr = (WORD)i; }
    CAttr*          GetAttr(UINT i)             { Assert(i < _cAttr); return(&_aAttr[i]); }

    BOOL            ValFromName(const TCHAR* pchName, TCHAR** ppchVal);
    CAttr*          AttrFromName(const TCHAR* pchName);

    LPTSTR          GetXmlNamespace(int* pIdx);

    static UINT     ComputeSize(BOOL fTiny, UINT cAttr) { return (offsetof(CHtmTag, _pch)+((!fTiny)*(offsetof(CHtmTag, _aAttr)-offsetof(CHtmTag, _pch)))+cAttr*sizeof(CAttr)); }
    UINT            ComputeSize() { return (ComputeSize(!!(_bFlags&TAGF_TINY), _cAttr)); }

private:
    DECLARE_MEMALLOC_NEW_DELETE()

    enum
    {
        TAGF_END        = 0x01,     // This is a closing tag (</FOO>)
        TAGF_EMPTY      = 0x02,     // This is an XML noscope tag (<FOO/>)
        TAGF_RESTART    = 0x04,     // True if parsing should restart (ETAG_RAW_CODEPAGE only)
        TAGF_ASCII      = 0x08,     // True if text contains only chars 0x00<ch<0x80 (ETAG_RAW_TEXT only)
        TAGF_NEXTBUF    = 0x10,     // This token should advance the text buffer (CHtmTagStm use only)
        TAGF_TINY       = 0x20,     // This token is only 4 bytes long (CHtmTagStm use only)
        TAGF_SCRIPTED   = 0x40,     // Element created through script (i.e. not through parser)
        TAGF_DEFER      = 0x80,     // True if script is known to be defered and preparser didn't block (</SCRIPT> only)
    };

    // Data Members

    BYTE            _bFlags;        // Flags (TGF_*)
    BYTE            _etag;          // Tag code (ETAG_*)
    WORD            _cAttr;         // Count of attributes in _aAttr[]

    union
    {
        union
        {
            TCHAR*  _pch;           // ETAG_RAW_TEXT        - Pointer to text (not NULL terminated)
                                    // ETAG_RAW_COMMENT     - Pointer to comment text (not NULL terminated)
                                    // Unknown tag          - Pointer to tag name (NULL terminated)
            CHtmTagStm* _pHTS;      // ETAG_RAW_SOURCE      - Pointer to CHtmTagStm for retrieving echoed source
        };
        ULONG       _ulLine;        // ETAG_SCRIPT          - Line of closing {>} of <SCRIPT> tag
        CODEPAGE    _cp;            // ETAG_RAW_CODEPAGE    - The codepage to switch to
        ULONG       _ulSize;        // ETAG_RAW_DOCSIZE     - The current document source size
    };

    union
    {
        ULONG       _cch;           // ETAG_RAW_TEXT        - Count of text characters
                                    // ETAG_RAW_COMMENT     - Count of comment characters
                                    // Unknown tag          - Count of tag name characters
        ULONG       _ulOffset;      // ETAG_SCRIPT          - Offset to closing {>} of <SCRIPT> tag
    };

    CAttr _aAttr[4];                // The attributs (real size is _cAttr)
};

class CHtmParseCtx
{
public:
    CHtmParseCtx(CHtmParseCtx* phpxParent)
    {
        _phpxParent = phpxParent;
        _phpxEmbed = phpxParent->GetHpxEmbed();
    }
    CHtmParseCtx(FLOAT f) {}
    virtual ~CHtmParseCtx() {}

    virtual HRESULT Init() { return S_OK; }
    virtual HRESULT Prepare() { return S_OK; }
    virtual HRESULT Commit() { return S_OK; }
    virtual HRESULT Finish() { return S_OK; }
    virtual HRESULT Execute() { return S_OK; }

    virtual HRESULT BeginElement(CTreeNode** ppNodeNew, CElement* pel, CTreeNode* pNodeCur, BOOL fEmpty)
    {
        RRETURN(_phpxEmbed->BeginElement(ppNodeNew, pel, pNodeCur, fEmpty));
    }
    virtual HRESULT EndElement(CTreeNode** ppNodeNew, CTreeNode* pNodeCur, CTreeNode* pNodeEnd)
    {
        RRETURN(_phpxEmbed->EndElement(ppNodeNew, pNodeCur, pNodeEnd));
    }
    virtual HRESULT AddText(CTreeNode* pNode, TCHAR* pch, ULONG cch, BOOL fAscii)
    {
        return S_OK;
    }
    virtual HRESULT AddTag(CHtmTag*pht)
    {
        return S_OK;
    }
    virtual HRESULT AddSource(CHtmTag* pht);

    virtual BOOL QueryTextlike(ELEMENT_TAG etag, CHtmTag* pht)
    {
        return _phpxParent->QueryTextlike(etag, pht);
    }
    virtual CElement* GetMergeElement()
    {
        return NULL;
    }
    virtual HRESULT InsertLPointer(CTreePos** pptp, CTreeNode* pNodeCur)
    {
        RRETURN(_phpxParent->InsertLPointer(pptp, pNodeCur));
    }
    virtual HRESULT InsertRPointer(CTreePos** pptp, CTreeNode* pNodeCur)
    {
        RRETURN(_phpxParent->InsertRPointer(pptp, pNodeCur));
    }

    virtual CHtmParseCtx* GetHpxRoot()
    {
        CHtmParseCtx* phpx = this;
        while(phpx->_phpxParent)
        {
            phpx = phpx->_phpxParent;
        }
        return phpx;
    }
    virtual CHtmParseCtx* GetHpxEmbed()
    {
        return _phpxEmbed;
    }

    ELEMENT_TAG* _atagReject;       // reject these tokens from normal processing
    ELEMENT_TAG* _atagAccept;       // allow these tokens for normal processing (tokens taken are !R || A)
    ELEMENT_TAG* _atagTag;          // tags which should be fed via AddTag (subset of R)
    ELEMENT_TAG* _atagAlwaysEnd;    // transform begin tag to end tag (subset of R)
    ELEMENT_TAG* _atagIgnoreEnd;    // throw away End tag, and don't turn into unknown (subset of R)
    BOOL         _fNeedExecute;     // call Execute (from parser-reentrant location) after Finish
    BOOL         _fDropUnknownTags; // give the context unknown tags (via AddUnknownTag)
    BOOL         _fExecuteOnEof;    // even execute me on EOF (no end tag required)
    BOOL         _fIgnoreSubsequent;// ignore the rest of the HTML file after I'm finished
    CHtmParseCtx* _phpxParent;      // immediate parent
    CHtmParseCtx* _phpxEmbed;       // context which requires notification of begin/end elements
};

#define ETAG_IMPLICIT_BEGIN     ETAG_UNKNOWN

enum PARSESCOPE
{
    SCOPE_EMPTY,        // Elements with no scope like <IMG>
    SCOPE_OVERLAP,      // Overlapping elements like <B>
    SCOPE_NESTED,       // Nesting elements like <P>
    PARSESCOPE_Last_Enum
};

enum PARSETEXTSCOPE
{
    TEXTSCOPE_NEUTRAL,  // Ability to contain text depends on parent (<FORM>)
    TEXTSCOPE_INCLUDE,  // Able to contain text (<BODY>, <TD>, <OPTION>)
    TEXTSCOPE_EXCLUDE,  // Not able to contain text (<TABLE>, <SELECT>)
    PARSETEXTSCOPE_Last_Enum
};

enum PARSETEXTTYPE
{
    TEXTTYPE_NEVER = 0, // Not textlike
    TEXTTYPE_ALWAYS,    // textlike
    TEXTTYPE_QUERY,     // Must query context
    PARSETEXTTYPE_Last_Enum
};

class CHtmlParseClass
{
public:
    // Empty, overlapping, or nested.
    PARSESCOPE      _scope;

    // Textlike tags must be treated as text with respect to including
    // or excluding them from text-sensitive containers.
    PARSETEXTTYPE _texttype;

    // EndContainers hide end tags:
    // An end tag cannot "see" a corresponding begin tag through a root
    // container. For example, in <TD><TABLE></TD></TABLE>, the </TD>
    // cannot see the <TD> because TABLE is an end container.
    // Also a tag cannot "extend" beyond the end of an end container:
    // when the end container ends, so does the contained element.
    ELEMENT_TAG* _atagEndContainers;

    // BeginContainers hide begin tags:
    // A tag cannot "see" prohibited, required, or masking containers
    // outside its begin containers. For example, in <LI><UL><LI>,
    // the second <LI> cannot see the first LI (which would be
    // prohibited) because UL is a start container for LI.
    ELEMENT_TAG* _atagBeginContainers;

    // MaskingContainers supress parsing:
    // If a masking container is on the stack while parsing a begin
    // tag, the begin tag is treated as an unknown tag.
    ELEMENT_TAG* _atagMaskingContainers;

    // ProhibitedContainers specifies end tags implied by begin tags:
    // If an element is parented by a prohibited container, the prohibited
    // container is closed before inserting the element.
    ELEMENT_TAG* _atagProhibitedContainers;

    // Required/DefaultContainer specifies begin tags implied by begin tags:
    // If an element is not parented by one of its required containers,
    // the element specified by _etagDefaultContainer is inserted if
    // possible. (if not possible, the tag is treated as an unknown tag.)
    // Also, if _fQueue is TRUE, instead of implying _etagDefaultContainer
    // immediately, the tag is queued and replayed after _etagDefaultContainer
    // appears.
    ELEMENT_TAG* _atagRequiredContainers;
    ELEMENT_TAG _etagDefaultContainer;
    BOOL _fQueueForRequired;


    // Text include/exclude specifies whether this element can contain
    // text. If TEXTSCOPE_EXCLUDE, it cannot contain text, and it must
    // specify a default subcontainer which can be inserted which can
    // contain text. Unlike the rest of the parsing DTD, these rules
    // apply downward because they are context-sensitive. (The default
    // container for text depends on which parent is TEXTSCOPE_EXCLUDE.)
    PARSETEXTSCOPE _textscope;
    ELEMENT_TAG _etagTextSubcontainer;

    // Alternate matching begin tags:
    // If non-null, an end tag of this type will match any begin tag
    // in the specified set.
    ELEMENT_TAG* _atagMatch;

    // Begin-tag substitute for an unmatched end tag
    // If an end tag is encountered which does have a corresponding
    // begin tag, the end tag is replaced by the specified begin tag.
    // E.g., an unmatched </P> is replaced by a <P>
    ELEMENT_TAG _etagUnmatchedSubstitute;

    // Context creator
    HRESULT (*_pfnHpxCreator)(CHtmParseCtx** pphpx, CElement* pelTop, CHtmParseCtx* phpxParent);

    BOOL _fMerge;

    // Implicit child
    // When an element is created, its required child is immediately created underneath it
    ELEMENT_TAG _etagImplicitChild;

    // Close required child
    // If TRUE, the implicit child is implicitly closed; otherwise, it is left open
    BOOL _fCloseImplicitChild;
};

extern CHtmlParseClass s_hpcUnknown;

class CTagDesc
{
public:
    TCHAR*              _pchTagName;
    CHtmlParseClass*    _pParseClass;
    HRESULT (*_pfnElementCreator)(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);
    DWORD               _dwTagDescFlags;

    inline BOOL         HasFlag(TAGDESC_FLAGS) const;
};

extern const CTagDesc g_atagdesc[];
extern CPtrBagCi<ELEMENT_TAG> g_bKnownTags;

inline const CTagDesc* TagDescFromEtag(ELEMENT_TAG tag)
{
    Assert(tag < ETAG_LAST);
    return &g_atagdesc[tag];
}

inline const TCHAR* NameFromEtag(ELEMENT_TAG tag)
{
    Assert(tag < ETAG_LAST);
    return(g_atagdesc[tag]._pchTagName);
}

inline ELEMENT_TAG EtagFromName(const TCHAR* pch, int nLen)
{
    return g_bKnownTags.GetCi(pch, nLen);
}

inline BOOL CTagDesc::HasFlag(TAGDESC_FLAGS flag) const
{
    return !!((_dwTagDescFlags&flag));
}

inline BOOL IsGenericTag(ELEMENT_TAG etag)
{
    return (ETAG_GENERIC==etag || ETAG_GENERIC_BUILTIN==etag || ETAG_GENERIC_LITERAL==etag);
}

//+------------------------------------------------------------------------
//
//  Function:   IsEtagInSet
//
//  Synopsis:   Determines if an element is within a set.
//
//-------------------------------------------------------------------------
inline BOOL IsEtagInSet(ELEMENT_TAG etag, ELEMENT_TAG* petagSet)
{
    Assert(petagSet);

    while(*petagSet)
    {
        if(etag == *petagSet)
        {
            return TRUE;
        }
        petagSet++;
    }
    return FALSE;
}

BOOL TagEndContainer(ELEMENT_TAG tag1, ELEMENT_TAG tag2);
BOOL TagHasNoEndTag(ELEMENT_TAG tag);
CHtmlParseClass* HpcFromEtag(ELEMENT_TAG tag);
BOOL TagProhibitedContainer(ELEMENT_TAG tag1, ELEMENT_TAG tag2);


// CreateElement function prototype.
HRESULT CreateElement(CHtmTag*      pht, 
                      CElement**    ppElementResult, 
                      CDocument*    pDoc, 
                      CMarkup *     pMarkup, 
                      BOOL *        pfDie,
                      DWORD         dwFlags=0);

extern CAssoc g_tagascA2;
extern CAssoc g_tagascACRONYM3;
extern CAssoc g_tagascADDRESS4;
extern CAssoc g_tagascAPPLET5;
extern CAssoc g_tagascAREA6;
extern CAssoc g_tagascB7;
extern CAssoc g_tagascBASE8;
extern CAssoc g_tagascBASEFONT9;
extern CAssoc g_tagascBDO10;
extern CAssoc g_tagascBGSOUND11;
extern CAssoc g_tagascBIG12;
extern CAssoc g_tagascBLINK13;
extern CAssoc g_tagascBLOCKQUOTE14;
extern CAssoc g_tagascBODY15;
extern CAssoc g_tagascBR16;
extern CAssoc g_tagascBUTTON17;
extern CAssoc g_tagascCAPTION18;
extern CAssoc g_tagascCENTER19;
extern CAssoc g_tagascCITE20;
extern CAssoc g_tagascCODE21;
extern CAssoc g_tagascCOL22;
extern CAssoc g_tagascCOLGROUP23;
extern CAssoc g_tagascCOMMENT24;
extern CAssoc g_tagascDD25;
extern CAssoc g_tagascDEL27;
extern CAssoc g_tagascDFN28;
extern CAssoc g_tagascDIR29;
extern CAssoc g_tagascDIV30;
extern CAssoc g_tagascDL31;
extern CAssoc g_tagascDT32;
extern CAssoc g_tagascEM33;
extern CAssoc g_tagascEMBED34;
extern CAssoc g_tagascFIELDSET35;
extern CAssoc g_tagascFONT36;
extern CAssoc g_tagascFORM37;
extern CAssoc g_tagascFRAME38;
extern CAssoc g_tagascFRAMESET39;
extern CAssoc g_tagascH140;
extern CAssoc g_tagascH241;
extern CAssoc g_tagascH342;
extern CAssoc g_tagascH443;
extern CAssoc g_tagascH544;
extern CAssoc g_tagascH645;
extern CAssoc g_tagascHEAD46;
extern CAssoc g_tagascHR47;
extern CAssoc g_tagascHTML48;
extern CAssoc g_tagascI49;
extern CAssoc g_tagascIFRAME50;
extern CAssoc g_tagascIMG51;
extern CAssoc g_tagascINPUT52;
extern CAssoc g_tagascINS54;
extern CAssoc g_tagascISINDEX55;
extern CAssoc g_tagascKBD56;
extern CAssoc g_tagascLABEL57;
extern CAssoc g_tagascLEGEND58;
extern CAssoc g_tagascLI59;
extern CAssoc g_tagascLINK60;
extern CAssoc g_tagascLISTING61;
extern CAssoc g_tagascMAP62;
extern CAssoc g_tagascMARQUEE63;
extern CAssoc g_tagascMENU64;
extern CAssoc g_tagascMETA65;
extern CAssoc g_tagascNEXTID66;
extern CAssoc g_tagascNOBR67;
extern CAssoc g_tagascNOEMBED68;
extern CAssoc g_tagascNOFRAMES70;
extern CAssoc g_tagascNOSCRIPT72;
extern CAssoc g_tagascOBJECT74;
extern CAssoc g_tagascOL75;
extern CAssoc g_tagascOPTION76;
extern CAssoc g_tagascP77;
extern CAssoc g_tagascPARAM78;
extern CAssoc g_tagascPLAINTEXT79;
extern CAssoc g_tagascPRE80;
extern CAssoc g_tagascQ81;
extern CAssoc g_tagascRP83;
extern CAssoc g_tagascRT84;
extern CAssoc g_tagascRUBY85;
extern CAssoc g_tagascS86;
extern CAssoc g_tagascSAMP87;
extern CAssoc g_tagascSCRIPT88;
extern CAssoc g_tagascSELECT89;
extern CAssoc g_tagascSMALL90;
extern CAssoc g_tagascSPAN91;
extern CAssoc g_tagascSTRIKE92;
extern CAssoc g_tagascSTRONG93;
extern CAssoc g_tagascSTYLE94;
extern CAssoc g_tagascSUB95;
extern CAssoc g_tagascSUP96;
extern CAssoc g_tagascTABLE97;
extern CAssoc g_tagascTBODY98;
extern CAssoc g_tagascTD100;
extern CAssoc g_tagascTEXTAREA101;
extern CAssoc g_tagascTFOOT102;
extern CAssoc g_tagascTH103;
extern CAssoc g_tagascTHEAD104;
extern CAssoc g_tagascTITLE105;
extern CAssoc g_tagascTR106;
extern CAssoc g_tagascTT107;
extern CAssoc g_tagascU108;
extern CAssoc g_tagascUL109;
extern CAssoc g_tagascVAR110;
extern CAssoc g_tagascWBR111;
extern CAssoc g_tagascXMP112;
extern CAssoc g_tagascifdef127;
extern CAssoc g_tagascendif130;
extern CAssoc g_tagascIMAGE51;


#define SZTAG_A (g_tagascA2._ach)
#define SZTAG_ACRONYM (g_tagascACRONYM3._ach)
#define SZTAG_ADDRESS (g_tagascADDRESS4._ach)
#define SZTAG_APPLET (g_tagascAPPLET5._ach)
#define SZTAG_AREA (g_tagascAREA6._ach)
#define SZTAG_B (g_tagascB7._ach)
#define SZTAG_BASE (g_tagascBASE8._ach)
#define SZTAG_BASEFONT (g_tagascBASEFONT9._ach)
#define SZTAG_BDO (g_tagascBDO10._ach)
#define SZTAG_BGSOUND (g_tagascBGSOUND11._ach)
#define SZTAG_BIG (g_tagascBIG12._ach)
#define SZTAG_BLINK (g_tagascBLINK13._ach)
#define SZTAG_BLOCKQUOTE (g_tagascBLOCKQUOTE14._ach)
#define SZTAG_BODY (g_tagascBODY15._ach)
#define SZTAG_BR (g_tagascBR16._ach)
#define SZTAG_BUTTON (g_tagascBUTTON17._ach)
#define SZTAG_CAPTION (g_tagascCAPTION18._ach)
#define SZTAG_CENTER (g_tagascCENTER19._ach)
#define SZTAG_CITE (g_tagascCITE20._ach)
#define SZTAG_CODE (g_tagascCODE21._ach)
#define SZTAG_COL (g_tagascCOL22._ach)
#define SZTAG_COLGROUP (g_tagascCOLGROUP23._ach)
#define SZTAG_COMMENT (g_tagascCOMMENT24._ach)
#define SZTAG_DD (g_tagascDD25._ach)
#define SZTAG_DEL (g_tagascDEL27._ach)
#define SZTAG_DFN (g_tagascDFN28._ach)
#define SZTAG_DIR (g_tagascDIR29._ach)
#define SZTAG_DIV (g_tagascDIV30._ach)
#define SZTAG_DL (g_tagascDL31._ach)
#define SZTAG_DT (g_tagascDT32._ach)
#define SZTAG_EM (g_tagascEM33._ach)
#define SZTAG_EMBED (g_tagascEMBED34._ach)
#define SZTAG_FIELDSET (g_tagascFIELDSET35._ach)
#define SZTAG_FONT (g_tagascFONT36._ach)
#define SZTAG_FORM (g_tagascFORM37._ach)
#define SZTAG_FRAME (g_tagascFRAME38._ach)
#define SZTAG_FRAMESET (g_tagascFRAMESET39._ach)
#define SZTAG_H1 (g_tagascH140._ach)
#define SZTAG_H2 (g_tagascH241._ach)
#define SZTAG_H3 (g_tagascH342._ach)
#define SZTAG_H4 (g_tagascH443._ach)
#define SZTAG_H5 (g_tagascH544._ach)
#define SZTAG_H6 (g_tagascH645._ach)
#define SZTAG_HEAD (g_tagascHEAD46._ach)
#define SZTAG_HR (g_tagascHR47._ach)
#define SZTAG_HTML (g_tagascHTML48._ach)
#define SZTAG_I (g_tagascI49._ach)
#define SZTAG_IFRAME (g_tagascIFRAME50._ach)
#define SZTAG_IMG (g_tagascIMG51._ach)
#define SZTAG_INPUT (g_tagascINPUT52._ach)
#define SZTAG_INS (g_tagascINS54._ach)
#define SZTAG_ISINDEX (g_tagascISINDEX55._ach)
#define SZTAG_KBD (g_tagascKBD56._ach)
#define SZTAG_LABEL (g_tagascLABEL57._ach)
#define SZTAG_LEGEND (g_tagascLEGEND58._ach)
#define SZTAG_LI (g_tagascLI59._ach)
#define SZTAG_LINK (g_tagascLINK60._ach)
#define SZTAG_LISTING (g_tagascLISTING61._ach)
#define SZTAG_MAP (g_tagascMAP62._ach)
#define SZTAG_MARQUEE (g_tagascMARQUEE63._ach)
#define SZTAG_MENU (g_tagascMENU64._ach)
#define SZTAG_META (g_tagascMETA65._ach)
#define SZTAG_NEXTID (g_tagascNEXTID66._ach)
#define SZTAG_NOBR (g_tagascNOBR67._ach)
#define SZTAG_NOEMBED (g_tagascNOEMBED68._ach)
#define SZTAG_NOFRAMES (g_tagascNOFRAMES70._ach)
#define SZTAG_NOSCRIPT (g_tagascNOSCRIPT72._ach)
#define SZTAG_OBJECT (g_tagascOBJECT74._ach)
#define SZTAG_OL (g_tagascOL75._ach)
#define SZTAG_OPTION (g_tagascOPTION76._ach)
#define SZTAG_P (g_tagascP77._ach)
#define SZTAG_PARAM (g_tagascPARAM78._ach)
#define SZTAG_PLAINTEXT (g_tagascPLAINTEXT79._ach)
#define SZTAG_PRE (g_tagascPRE80._ach)
#define SZTAG_Q (g_tagascQ81._ach)
#define SZTAG_RP (g_tagascRP83._ach)
#define SZTAG_RT (g_tagascRT84._ach)
#define SZTAG_RUBY (g_tagascRUBY85._ach)
#define SZTAG_S (g_tagascS86._ach)
#define SZTAG_SAMP (g_tagascSAMP87._ach)
#define SZTAG_SCRIPT (g_tagascSCRIPT88._ach)
#define SZTAG_SELECT (g_tagascSELECT89._ach)
#define SZTAG_SMALL (g_tagascSMALL90._ach)
#define SZTAG_SPAN (g_tagascSPAN91._ach)
#define SZTAG_STRIKE (g_tagascSTRIKE92._ach)
#define SZTAG_STRONG (g_tagascSTRONG93._ach)
#define SZTAG_STYLE (g_tagascSTYLE94._ach)
#define SZTAG_SUB (g_tagascSUB95._ach)
#define SZTAG_SUP (g_tagascSUP96._ach)
#define SZTAG_TABLE (g_tagascTABLE97._ach)
#define SZTAG_TBODY (g_tagascTBODY98._ach)
#define SZTAG_TD (g_tagascTD100._ach)
#define SZTAG_TEXTAREA (g_tagascTEXTAREA101._ach)
#define SZTAG_TFOOT (g_tagascTFOOT102._ach)
#define SZTAG_TH (g_tagascTH103._ach)
#define SZTAG_THEAD (g_tagascTHEAD104._ach)
#define SZTAG_TITLE (g_tagascTITLE105._ach)
#define SZTAG_TR (g_tagascTR106._ach)
#define SZTAG_TT (g_tagascTT107._ach)
#define SZTAG_U (g_tagascU108._ach)
#define SZTAG_UL (g_tagascUL109._ach)
#define SZTAG_VAR (g_tagascVAR110._ach)
#define SZTAG_WBR (g_tagascWBR111._ach)
#define SZTAG_XMP (g_tagascXMP112._ach)
#define SZTAG_RAW_COMMENT (_T("!"))
#define SZTAG_ifdef (g_tagascifdef127._ach)
#define SZTAG_endif (g_tagascendif130._ach)
#define SZTAG_IMAGE (g_tagascIMAGE51._ach)

BOOL MergableTags(ELEMENT_TAG etag1, ELEMENT_TAG etag2);

#endif //__XINDOWS_SITE_BASE_HTMLTAGS_H__