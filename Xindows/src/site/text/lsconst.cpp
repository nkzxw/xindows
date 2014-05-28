
#include "stdafx.h"
#include "LineServices.h"

#include "lsobj.h"

#include "msls/tatenak.h"
#include "msls/hih.h"
#include "msls/warichu.h"
#include "msls/robj.h"

const CLineServices::LSCBK CLineServices::s_lscbk =
{
    &CLineServices::NewPtr,
    &CLineServices::DisposePtr,
    &CLineServices::ReallocPtr,
    &CLineServices::FetchRun,
    &CLineServices::GetAutoNumberInfo,
    &CLineServices::GetNumericSeparators,
    &CLineServices::CheckForDigit,
    &CLineServices::FetchPap,
    &CLineServices::FetchTabs,
    &CLineServices::GetBreakThroughTab,
    &CLineServices::FGetLastLineJustification,
    &CLineServices::CheckParaBoundaries,
    &CLineServices::GetRunCharWidths,
    &CLineServices::CheckRunKernability,
    &CLineServices::GetRunCharKerning,
    &CLineServices::GetRunTextMetrics,
    &CLineServices::GetRunUnderlineInfo,
    &CLineServices::GetRunStrikethroughInfo,
    &CLineServices::GetBorderInfo,
    &CLineServices::ReleaseRun,
    &CLineServices::Hyphenate,
    &CLineServices::GetHyphenInfo,
    &CLineServices::DrawUnderline,
    &CLineServices::DrawStrikethrough,
    &CLineServices::DrawBorder,
    &CLineServices::DrawUnderlineAsText,
    &CLineServices::FInterruptUnderline,
    &CLineServices::FInterruptShade,
    &CLineServices::FInterruptBorder,
    &CLineServices::ShadeRectangle,
    &CLineServices::DrawTextRun,
    &CLineServices::DrawSplatLine,
    &CLineServices::FInterruptShaping,
    &CLineServices::GetGlyphs,
    &CLineServices::GetGlyphPositions,
    &CLineServices::ResetRunContents,
    &CLineServices::DrawGlyphs,
    &CLineServices::GetGlyphExpansionInfo,
    &CLineServices::GetGlyphExpansionInkInfo,
    &CLineServices::GetEms,
    &CLineServices::PunctStartLine,
    &CLineServices::ModWidthOnRun,
    &CLineServices::ModWidthSpace,
    &CLineServices::CompOnRun,
    &CLineServices::CompWidthSpace,
    &CLineServices::ExpOnRun,
    &CLineServices::ExpWidthSpace,
    &CLineServices::GetModWidthClasses,
    &CLineServices::GetBreakingClasses,
    &CLineServices::FTruncateBefore,
    &CLineServices::CanBreakBeforeChar,
    &CLineServices::CanBreakAfterChar,
    &CLineServices::FHangingPunct,
    &CLineServices::GetSnapGrid,
    &CLineServices::DrawEffects,
    &CLineServices::FCancelHangingPunct,
    &CLineServices::ModifyCompAtLastChar,
    &CLineServices::EnumText,
    &CLineServices::EnumTab,
    &CLineServices::EnumPen,
    &CLineServices::GetObjectHandlerInfo,
    &CLineServices::AssertFailed
};

const struct lstxtcfg CLineServices::s_lstxtcfg =
{
    LS_AVG_CHARS_PER_LINE, // cEstimatedCharsPerLine; modify as necessary
    WCH_UNDEF,
    WCH_NULL,
    WCH_SPACE,
    WCH_HYPHEN,
    WCH_TAB,
    WCH_ENDPARA1,
    WCH_ENDPARA2,
    WCH_ALTENDPARA,
    WCH_SYNTHETICLINEBREAK,
    WCH_COLUMNBREAK,
    WCH_SECTIONBREAK,
    WCH_PAGEBREAK,
    WCH_UNDEF,                          // NB (cthrash) Don't let LS default-handle NONBREAKSPACE
    WCH_NONBREAKHYPHEN,
    WCH_NONREQHYPHEN,
    WCH_EMDASH,
    WCH_ENDASH,
    WCH_EMSPACE,
    WCH_ENSPACE,
    WCH_NARROWSPACE,
    WCH_OPTBREAK,
    WCH_NOBREAK,
    WCH_FESPACE,
    WCH_ZWJ,
    WCH_ZWNJ,
    WCH_TOREPLACE,
    WCH_REPLACE,
    WCH_VISINULL,
    WCH_VISIALTENDPARA,
    WCH_VISIENDLINEINPARA,
    WCH_VISIENDPARA,
    WCH_VISISPACE,
    WCH_VISINONBREAKSPACE,
    WCH_VISINONBREAKHYPHEN,
    WCH_VISINONREQHYPHEN,
    WCH_VISITAB,
    WCH_VISIEMSPACE,
    WCH_VISIENSPACE,
    WCH_VISINARROWSPACE,
    WCH_VISIOPTBREAK,
    WCH_VISINOBREAK,
    WCH_VISIFESPACE,
    WCH_ESCANMRUN,
};

CLineServices::LSIMETHODS CLineServices::s_rgLsiMethods[CLineServices::LSOBJID_COUNT] =
{
    // The order of these is determined by the order of 
    // enum LSOBJID

    // LSOBJID_EMBEDDED
    {
        &CLineServices::CreateILSObj,
        &CEmbeddedILSObj::DestroyILSObj,
        &CEmbeddedILSObj::SetDoc,
        &CEmbeddedILSObj::CreateLNObj,
        &CEmbeddedILSObj::DestroyLNObj,
        (LSERR (WINAPI CILSObjBase::*)(PCFMTIN, FMTRES*))&CEmbeddedILSObj::Fmt,
        (LSERR (WINAPI CILSObjBase::*)(BREAKREC*, DWORD, PCFMTIN, FMTRES*))&CEmbeddedILSObj::FmtResume,
        &CEmbeddedDobj::GetModWidthPrecedingChar,
        &CEmbeddedDobj::GetModWidthFollowingChar,
        &CLineServices::TruncateChunk,
        &CEmbeddedDobj::FindPrevBreakChunk,
        &CEmbeddedDobj::FindNextBreakChunk,
        &CLineServices::ForceBreakChunk,
        &CEmbeddedDobj::SetBreak,  
        &CEmbeddedDobj::GetSpecialEffectsInside,
        &CEmbeddedDobj::FExpandWithPrecedingChar,
        &CEmbeddedDobj::FExpandWithFollowingChar,
        &CEmbeddedDobj::CalcPresentation,  
        &CEmbeddedDobj::QueryPointPcp,
        &CEmbeddedDobj::QueryCpPpoint,
        &CEmbeddedDobj::Enum,
        &CEmbeddedDobj::Display,
        &CEmbeddedDobj::DestroyDObj
    },

    // LSOBJID_NOBR
    {
        &CLineServices::CreateILSObj,
        &CNobrILSObj::DestroyILSObj,
        &CNobrILSObj::SetDoc,
        &CNobrILSObj::CreateLNObj,
        &CNobrILSObj::DestroyLNObj,
        (LSERR (WINAPI CILSObjBase::*)(PCFMTIN, FMTRES*))&CNobrILSObj::Fmt,
        (LSERR (WINAPI CILSObjBase::*)(BREAKREC*, DWORD, PCFMTIN, FMTRES*))&CNobrILSObj::FmtResume,
        &CNobrDobj::GetModWidthPrecedingChar,
        &CNobrDobj::GetModWidthFollowingChar,
        &CLineServices::TruncateChunk,
        &CNobrDobj::FindPrevBreakChunk,
        &CNobrDobj::FindNextBreakChunk,
        &CLineServices::ForceBreakChunk,
        &CNobrDobj::SetBreak,  
        &CNobrDobj::GetSpecialEffectsInside,
        &CNobrDobj::FExpandWithPrecedingChar,
        &CNobrDobj::FExpandWithFollowingChar,
        &CNobrDobj::CalcPresentation,  
        &CNobrDobj::QueryPointPcp,  
        &CNobrDobj::QueryCpPpoint,
        &CNobrDobj::Enum,
        &CNobrDobj::Display,
        &CNobrDobj::DestroyDObj
    },

    // The remainder is populated by LineServices.
};

const CLineServices::RUBYINIT CLineServices::s_rubyinit =
{
    RUBY_VERSION,
    RubyMainLineFirst,
    WCH_ESCRUBY,
    WCH_ESCMAIN,
    WCH_NULL,
    WCH_NULL,
    &CLineServices::FetchRubyPosition,
    &CLineServices::FetchRubyWidthAdjust,
    &CLineServices::RubyEnum,
};

const CLineServices::TATENAKAYOKOINIT CLineServices::s_tatenakayokoinit =
{
    TATENAKAYOKO_VERSION,
    WCH_ENDTATENAKAYOKO,
    WCH_NULL,
    WCH_NULL,
    WCH_NULL,
    &CLineServices::GetTatenakayokoLinePosition,
    &CLineServices::TatenakayokoEnum,
};

const CLineServices::HIHINIT CLineServices::s_hihinit =
{
    HIH_VERSION,
    WCH_ENDHIH,
    WCH_NULL,
    WCH_NULL,
    WCH_NULL,
    &CLineServices::HihEnum,
};

const CLineServices::WARICHUINIT CLineServices::s_warichuinit =
{
    WARICHU_VERSION,
    WCH_ENDFIRSTBRACKET,
    WCH_ENDTEXT,
    WCH_ENDWARICHU,
    WCH_NULL,
    &CLineServices::GetWarichuInfo,
    &CLineServices::FetchWarichuWidthAdjust,
    &CLineServices::WarichuEnum,
    FALSE
};

const CLineServices::REVERSEINIT CLineServices::s_reverseinit =
{
    REVERSE_VERSION,
    WCH_ENDREVERSE,
    WCH_NULL,
    WCH_NULL,
    WCH_NULL,
    &CLineServices::ReverseEnum,
};

const WCHAR CLineServices::s_achTabLeader[tomLines] =
{
    WCH_NULL,
    WCH_DOT,
    WCH_HYPHEN
};

const CLineServices::SYNTHDATA CLineServices::s_aSynthData[SYNTHTYPE_COUNT] =
{
    // wch              idObj               typeEndObj                          fObjStart   fObjEnd,    fHidden     fLSCPStop   pszSynthName
    { WCH_UNDEF,        idObjTextChp,       SYNTHTYPE_NONE,                     FALSE,      FALSE,      TRUE,       FALSE,      DEBUG_ONLY(_T("[none]")) },
    { WCH_SECTIONBREAK, idObjTextChp,       SYNTHTYPE_NONE,                     FALSE,      FALSE,      FALSE,      FALSE,      DEBUG_ONLY(_T("[sectionbreak]")) },
    { WCH_REVERSE,      LSOBJID_REVERSE,    SYNTHTYPE_ENDREVERSE,               TRUE,       FALSE,      FALSE,      FALSE,      DEBUG_ONLY(_T("[reverse]")) },
    { WCH_ENDREVERSE,   LSOBJID_REVERSE,    SYNTHTYPE_NONE,                     FALSE,      TRUE,       FALSE,      FALSE,      DEBUG_ONLY(_T("[endreverse]")) },
    { WCH_NOBRBLOCK,    LSOBJID_NOBR,       SYNTHTYPE_ENDNOBR,                  TRUE,       FALSE,      FALSE,      FALSE,      DEBUG_ONLY(_T("[nobr]")) },
    { WCH_NOBRBLOCK,    LSOBJID_NOBR,       SYNTHTYPE_NONE,                     FALSE,      TRUE,       FALSE,      FALSE,      DEBUG_ONLY(_T("[endnobr]")) },
    { WCH_ENDPARA1,     idObjTextChp,       SYNTHTYPE_NONE,                     FALSE,      FALSE,      FALSE,      TRUE,       DEBUG_ONLY(_T("[endpara1]")) },
    { WCH_ALTENDPARA,   idObjTextChp,       SYNTHTYPE_NONE,                     FALSE,      FALSE,      FALSE,      FALSE,      DEBUG_ONLY(_T("[altendpara]")) },
    { WCH_UNDEF,     	LSOBJID_RUBY,       SYNTHTYPE_ENDRUBYTEXT,              TRUE,       FALSE,      FALSE,      FALSE,      DEBUG_ONLY(_T("[rubymain]")) },
    { WCH_ESCMAIN,     	LSOBJID_RUBY,       SYNTHTYPE_NONE,                  	TRUE,       TRUE,       FALSE,      FALSE,      DEBUG_ONLY(_T("[endrubymain]")) },
    { WCH_ESCRUBY,     	LSOBJID_RUBY,       SYNTHTYPE_NONE,                  	FALSE,      TRUE,       FALSE,      FALSE,      DEBUG_ONLY(_T("[endrubytext]")) },
    { WCH_SYNTHETICLINEBREAK, idObjTextChp, SYNTHTYPE_NONE,                     FALSE,      FALSE,      FALSE,      TRUE,       DEBUG_ONLY(_T("[linebreak]")) },
};


// This is a line-services structure defining an escape
// sequence.  In this case, it is the escape sequence ending
// a nobr block.  It is one character long, and that character
// must be in the range wch_nobrblock through wch_nobrblock
// (i.e. must be exactly wch_nobrblock).
const LSESC CNobrILSObj::s_lsescEndNOBR[NBREAKCHARS] = 
{
    { WCH_NOBRBLOCK,  WCH_NOBRBLOCK },
};
