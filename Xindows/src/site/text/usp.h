
#ifndef __XINDOWS_SITE_TEXT_USP_H__
#define __XINDOWS_SITE_TEXT_USP_H__

#include <usp10.h>

#define STACK_ALLOC_RUNS    16
#define USP10_NOT_FOUND     0x80070002

HRESULT WINAPI ScriptItemize(
    PCWSTR                  pwcInChars,     // In   Unicode string to be itemized
    int                     cInChars,       // In   Character count to itemize
    int                     cItemGrow,      // In   Items to grow by if paryItems is too small
    const SCRIPT_CONTROL*   psControl,      // In   Analysis control (optional)
    const SCRIPT_STATE*     psState,        // In   Initial bidi algorithm state (optional)
    CDataAry<SCRIPT_ITEM>*  paryItems,      // Out  Array to receive itemization
    PINT                    pcItems);       // Out  Count of items processed    HRESULT hr;

const SCRIPT_PROPERTIES* GetScriptProperties(WORD eScript);
const WORD GetNumericScript(DWORD lang);

inline BOOL IsComplexScript(WORD eScript)
{
    return GetScriptProperties(eScript)->fComplex;
}

inline BOOL IsNumericScript(WORD eScript)
{
    return GetScriptProperties(eScript)->fNumeric;
}

inline BOOL NeedWordBreak(WORD eScript)
{
    return GetScriptProperties(eScript)->fNeedsWordBreaking;
}

inline BOOL GetScriptCharSet(WORD eScript)
{
    return GetScriptProperties(eScript)->bCharSet;
}

#endif //__XINDOWS_SITE_TEXT_USP_H__