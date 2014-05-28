
#ifndef __XINDOWS_SITE_TEXT_UNIWBK_H__
#define __XINDOWS_SITE_TEXT_UNIWBK_H__

typedef BYTE CHAR_CLASS;
typedef BYTE WBKCLS;

enum
{
    wbkclsPunctSymb,     // 0
    wbkclsKanaFollow,    // 1
    wbkclsKatakanaW,     // 2
    wbkclsHiragana,      // 3
    wbkclsTab,           // 4
    wbkclsKanaDelim,     // 5
    wbkclsPrefix,        // 6
    wbkclsPostfix,       // 7
    wbkclsSpaceA,        // 8
    wbkclsAlpha,         // 9
    wbkclsIdeoW,         // 10
    wbkclsSuperSub,      // 11
    wbkclsDigitsN,       // 12
    wbkclsPunctInText,   // 13
    wbkclsDigitsW,       // 14
    wbkclsKatakanaN,     // 15
    wbkclsHangul,        // 16
    wbkclsLatinW,        // 17
    wbkclsLim
};

WBKCLS WordBreakClassFromCharClass(CHAR_CLASS cc);
BOOL   IsWordBreakBoundaryDefault(WCHAR, WCHAR);
BOOL   IsProofWordBreakBoundary(WCHAR, WCHAR);

#endif //__XINDOWS_SITE_TEXT_UNIWBK_H__