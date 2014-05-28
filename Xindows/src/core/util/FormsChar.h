
#ifndef __XINDOWS_CORE_UTIL_FORMSCHAR_H__
#define __XINDOWS_CORE_UTIL_FORMSCHAR_H__

#undef IsCharAlpha
#undef IsCharSpace
#undef IsCharAlphaNumeric

#define IsCharAlpha         FormsIsCharAlpha
#define IsCharSpace         FormsIsCharSpace
#define IsCharPunct         FormsIsCharPunct
#define IsCharDigit         FormsIsCharDigit
#define IsCharCntrl         FormsIsCharCntrl
#define IsCharXDigit        FormsIsCharXDigit
#define IsCharBlank         FormsIsCharBlank
#define IsCharBlank         FormsIsCharBlank
#define IsCharAlphaNumeric  FormsIsCharAlphaNumeric

extern "C" DWORD adwData[218];
extern "C" BYTE abIndex[98][8];
extern "C" BYTE abType1Alpha[256];
extern "C" BYTE abType1Punct[256];
extern "C" BYTE abType1Digit[256];

#define __BIT_SHIFT     0
#define __INDEX_SHIFT   5
#define __PAGE_SHIFT    8

#define __BIT_MASK      31
#define __INDEX_MASK    7

// straight lookup functions are inlined.
#define ISCHARFUNC(type, wch) \
{ \
    return (adwData[abIndex[abType1##type[wch>>__PAGE_SHIFT]] \
        [(wch>>__INDEX_SHIFT)&__INDEX_MASK]]>>(wch&__BIT_MASK)) & 1; \
}
    
extern "C" { inline BOOL FormsIsCharAlpha(WCHAR wch) ISCHARFUNC(Alpha, wch) }
extern "C" { inline BOOL FormsIsCharPunct(WCHAR wch) ISCHARFUNC(Punct, wch) }
extern "C" { inline BOOL FormsIsCharDigit(WCHAR wch) ISCHARFUNC(Digit, wch) }

// switched lookup functions are not inlined.
extern "C" BOOL FormsIsCharCntrl(WCHAR wch);
extern "C" BOOL FormsIsCharXDigit(WCHAR wch);
extern "C" BOOL FormsIsCharSpace(WCHAR wch);
extern "C" BOOL FormsIsCharAlphaNumeric(WCHAR wch);
extern "C" BOOL FormsIsCharBlank(WCHAR wch);

#endif //__XINDOWS_CORE_UTIL_FORMSCHAR_H__