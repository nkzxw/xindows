
#ifndef __XINDOWS_SITE_TEXT_UNISID_H__
#define __XINDOWS_SITE_TEXT_UNISID_H__

typedef BYTE    CHAR_CLASS;
typedef BYTE    SCRIPT_ID;
typedef __int64 SCRIPT_IDS;

const SCRIPT_ID  sidNil = SCRIPT_ID(-1);
const SCRIPT_IDS sidsNotSet = SCRIPT_IDS(0);
const SCRIPT_IDS sidsAll = SCRIPT_IDS(0x7FFFFFFFFFFFFFFF);

extern const SCRIPT_ID asidAscii[128];

SCRIPT_ID ScriptIDFromCharClass(CHAR_CLASS cc);

// Trident internal pseudo-script_id
const SCRIPT_ID sidSurrogateA = sidLim;
const SCRIPT_ID sidSurrogateB = sidSurrogateA + 1;
const SCRIPT_ID sidAmbiguous = sidSurrogateB + 1;
const SCRIPT_ID sidEUDC = sidAmbiguous + 1;
const SCRIPT_ID sidHalfWidthKana = sidEUDC + 1;
const SCRIPT_ID sidTridentLim = sidHalfWidthKana + 1;

const SCRIPT_ID sidLimPlusSurrogates = sidSurrogateB + 1; // for fontlinking (OPTIONSETTINGS)

// SCRIPT_IDS
inline SCRIPT_IDS ScriptBit(SCRIPT_ID sid)
{
    return SCRIPT_IDS(1)<<sid;
}

DEBUG_ONLY(const TCHAR* SidName(SCRIPT_ID sid));
DEBUG_ONLY(void DumpSids(SCRIPT_IDS sids));

SCRIPT_IDS ScriptIDsFromFont(HDC hdc, HFONT hfont, BOOL fTrueType);

inline SCRIPT_ID FoldScriptIDs(SCRIPT_ID sidL, SCRIPT_ID sidR)
{
    // This is not a bidirectional merge -- it is only intended for merging
    // a sid with another preceeding it.
    return (sidR==sidMerge || (sidR==sidAsciiSym && sidL==sidAsciiLatin))?sidL:sidR;
}

// Fontlinking support
SCRIPT_ID UnUnifyHan(CODEPAGE cpDoc, LCID lcid);
CODEPAGE DefaultCodePageFromScript(SCRIPT_ID* psid, CODEPAGE cpDoc, LCID lcid);
BYTE DefaultCharSetFromScriptAndCharset(SCRIPT_ID sid, BYTE bCharSetCF);
BYTE DefaultCharSetFromScriptAndCodePage(SCRIPT_ID sid, UINT uiFamilyCodePage);

extern const SCRIPT_IDS s_sidsTable[];
extern const BYTE s_abScriptIDsIndex[256];
inline SCRIPT_IDS ScriptIDsFromCharSet(BYTE bCharSet)
{
    return s_sidsTable[s_abScriptIDsIndex[bCharSet]];
}

SCRIPT_ID DefaultSidForCodePage(UINT uiFamilyCodePage);
SCRIPT_ID RegistryAppropriateSidFromSid(SCRIPT_ID sid);

// Context digit helper
SCRIPT_ID DefaultScriptIDFromLang(LANGID lang);

#endif //__XINDOWS_SITE_TEXT_UNISID_H__