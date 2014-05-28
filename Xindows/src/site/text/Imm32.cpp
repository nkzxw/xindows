
#include "stdafx.h"

static const WORD s_CodePageTable[] =
{
// CodePage       PLID  primary language
// -------------------------------------
       0,       // 00 - undefined
    1256,       // 01 - Arabic
    1251,       // 02 - Bulgarian
    1252,       // 03 - Catalan
     950,       // 04 - Taiwan, Hong Kong (PRC and Singapore are 936)
    1250,       // 05 - Czech
    1252,       // 06 - Danish
    1252,       // 07 - German
    1253,       // 08 - Greek
    1252,       // 09 - English
    1252,       // 0a - Spanish
    1252,       // 0b - Finnish
    1252,       // 0c - French
    1255,       // 0d - Hebrew
    1250,       // 0e - Hungarian
    1252,       // 0f - Icelandic
    1252,       // 10 - Italian
     932,       // 11 - Japan
     949,       // 12 - Korea
    1252,       // 13 - Dutch
    1252,       // 14 - Norwegian
    1250,       // 15 - Polish
    1252,       // 16 - Portuguese
       0,       // 17 - Rhaeto-Romanic
    1250,       // 18 - Romanian
    1251,       // 19 - Russian
    1250,       // 1a - Croatian
    1250,       // 1b - Slovak
    1250,       // 1c - Albanian
    1252,       // 1d - Swedish
     874,       // 1e - Thai
    1254,       // 1f - Turkish
       0,       // 20 - Urdu
    1252,       // 21 - Indonesian
    1251,       // 22 - Ukranian
    1251,       // 23 - Byelorussian
    1250,       // 24 - Slovenian
    1257,       // 25 - Estonia
    1257,       // 26 - Latvian
    1257,       // 27 - Lithuanian
       0,       // 28 - undefined
    1256,       // 29 - Farsi
    1258,       // 2a - Vietnanese
       0,       // 2b - undefined
       0,       // 2c - undefined
    1252,       // 2d - Basque
       0,       // 2e - Sorbian
    1251,       // 2f - Macedonian
    1252,       // 30 - Sutu        *** use 1252 for the following ***
    1252,       // 31 - Tsonga
    1252,       // 32 - Tswana
    1252,       // 33 - Venda
    1252,       // 34 - Xhosa
    1252,       // 35 - Zulu
    1252,       // 36 - Africaans (uses 1252)
    1252,       // 38 - Faerose
    1252,       // 39 - Hindi
    1252,       // 3a - Maltese
    1252,       // 3b - Sami
    1252,       // 3c - Gaelic
    1252,       // 3e - Malaysian
       0,       // 3f -
       0,       // 40 -
    1252,       // 41 - Swahili
       0,       // 42 - 
       0,       // 43 - 
       0,       // 44 - 
    1252,       // 45 - Bengali *** The following languages use 1252, but
    1252,       // 46 - Gurmuki     actually have only Unicode characters ***
    1252,       // 47 - Gujarait
    1252,       // 48 - Oriya
    1252,       // 49 - Tamil
    1252,       // 4a - Telugu
    1252,       // 4b - Kannada
    1252,       // 4c - Malayalam
       0,       // 4d - 
       0,       // 4e - 
       0,       // 4f - 
    1252,       // 50 - Mongolian
    1252,       // 51 - Tibetan
       0,       // 52 - 
    1252,       // 53 - Khmer
    1252,       // 54 - Lao
    1252,       // 55 - Burmese
       0,       // 56 - LANG_MAX
};

#define ns_CodePageTable    (sizeof(s_CodePageTable)/sizeof(s_CodePageTable[0]))

#if !defined(lidSerbianCyrillic)
  #define lidSerbianCyrillic 0xc1a
#else
  #if lidSerbianCyrillic != 0xc1a
    #error "lidSerbianCyrillic macro value has changed"
  #endif // lidSerbianCyrillic
#endif

/*
 *  ConvertLanguageIDtoCodePage (lid)
 *
 *  @mfunc      Maps a language ID to a Code Page
 *
 *  @rdesc      returns Code Page
 *
 *  @devnote:
 *      This routine takes advantage of the fact that except for Chinese,
 *      the code page is determined uniquely by the primary language ID,
 *      which is given by the low-order 10 bits of the lcid.
 *
 *      The WORD s_CodePageTable could be replaced by a BYTE with the addition
 *      of a couple of if's and the BYTE table replaced by a nibble table
 *      with the addition of a shift and a mask.  Since the table is only
 *      92 bytes long, it seems that the simplicity of using actual code page
 *      values is worth the extra bytes.
 */
UINT ConvertLanguageIDtoCodePage(WORD langid)
{
    WORD langidT = PRIMARYLANGID(langid);       // langidT = primary language (Plangid)
    CODEPAGE cp;

    if(langidT >= LANG_CROATIAN)                // Plangid = 0x1a
    {
        if(langid == lidSerbianCyrillic)        // Special case for langid = 0xc1a
        {
            return 1251;                        // Use Cyrillic code page
        }
        if(langidT >= ns_CodePageTable)         // Africans Plangid = 0x36, which
        {
            return CP_ACP;                      // is outside table
        }
    }

    cp = s_CodePageTable[langidT];              // Translate Plangid to code page

    if(cp!=950 || (langid&0x400))               // All but Singapore, PRC, and Serbian
    {
        return cp!=0?cp:CP_ACP;                 // Remember there are holes in the array
                                                // that may not always be there.
    }
    return 936;                                 // Singapore and PRC
}

/*
 *  GetKeyboardCodePage ()
 *
 *  @mfunc      Gets Code Page for keyboard active on current thread
 *
 *  @rdesc      returns Code Page
 */
UINT GetKeyboardCodePage()
{
    return ConvertLanguageIDtoCodePage((WORD)(DWORD)GetKeyboardLayout(0 /*idThread*/));
}
