
#ifndef __XINDOWS_CORE_UTIL_COLORVALUE_H__
#define __XINDOWS_CORE_UTIL_COLORVALUE_H__

// CColorValue related structure/functions/definitions
struct COLORVALUE_PAIR
{
    const TCHAR* szName;
    DWORD dwValue;
};

const struct COLORVALUE_PAIR* FindColorByValue(DWORD dwValue);
const struct COLORVALUE_PAIR* FindColorByColor(DWORD lColor);
const struct COLORVALUE_PAIR* FindColorByName(const TCHAR* szString);

class CColorValue : public CVoid
{
public:
    DWORD _dwValue;

    // What a royal mess
    //
    // Some types are directly extractable simply by masking with MASK_FLAG.  These are:
    //
    //      TYPE_UNDEF          0xff
    //      TYPE_POUND1         0xfe
    //      TYPE_POUND2         0xfd
    //      TYPE_POUND3         0xfc
    //      TYPE_POUND4         0xfb
    //      TYPE_POUND5         0xfa
    //      TYPE_RGB            0xf9
    //      TYPE_TRANSPARENT    0xf8
    //      TYPE_UNIXSYSCOL     0x04
    //      TYPE_POUND6         0x00
    //
    // Some types pack some extra info (the color index) into the high byte (yum).  These are:
    //
    //      TYPE_SYSNAME    0xA0 - 0xBF
    //      TYPE_SYSINDEX   0xC0 - 0xDF
    //      TYPE_NAME       0x11 - 0x9C
    //
    // Believe it or not, there is actually room for more stuff the range 0xE0 - F8 is apparently not used!
    //
    // This scheme will blow when the number of named colors exceeds 143 or when the number of system colors
    // exceeds 31
    //
    // At the cost of an extra indirection, named colors could be handled is a much simpler manner.  A single
    // flag value for each type of named color and then the lower 3 bytes are the index into the table.
    enum TYPE
    {
        TYPE_UNDEF     = 0xFF000000,    // color is undefined
        TYPE_POUND1    = 0xFE000000,
        TYPE_POUND2    = 0xFD000000,
        TYPE_POUND3    = 0xFC000000,
        TYPE_POUND4    = 0xFB000000,
        TYPE_POUND5    = 0xFA000000,
        TYPE_RGB       = 0xF9000000,    // color is rgb(r,g,b) or rgb(r%,g%,b%) (or mix)
        TYPE_TRANSPARENT = 0xF8000000,  // color is "transparent"
        TYPE_UNIXSYSCOL  = 0x04000000,  // color is unix system colormap entry index
        TYPE_NAME      = 0x11000000,    // color is named
        TYPE_SYSINDEX  = 0xC0000000,    // color is a system color (not named)
        TYPE_SYSNAME   = 0xA0000000,    // color is a named system color needs 29 in high byte
        TYPE_POUND6    = 0x00000000     // color is #rrggbb
    };

    enum MASK
    {
        MASK_COLOR      = 0x00ffffff,    // mask to extract color
        MASK_SYSCOLOR   = 0x1f000000,    // mask to extract the system color
        MASK_FLAG       = 0xff000000     // mask to extract type
    };

#define VALUE_UNDEF     0xFFFFFFFF

    int Index(const DWORD dwValue) const;

    static HRESULT FormatAsPound6Str(BSTR* pszColor, DWORD dwColor);
    HRESULT FormatBuffer(LPTSTR szBuffer, UINT uMaxLen, BOOL fReturnAsHex=FALSE) const;

    CColorValue() { _dwValue = VALUE_UNDEF; }
    CColorValue(DWORD dwValue) { _dwValue = dwValue; }
    CColorValue(VARIANT*);

    // NB: by setting fLookupName to TRUE, we will preferentially
    // map to color names, not numbers.
    long SetValue(long lColor, BOOL fLookupName, TYPE type);
    long SetValue(long lColor, BOOL fLookupName) { return SetValue(lColor, fLookupName, TYPE_POUND6); }
    long SetValue(const struct COLORVALUE_PAIR* pColorName);
    long SetRawValue(DWORD dwValue);
    long SetFromRGB(DWORD dwRGB) { return SetRawValue(((dwRGB&0xffL)<<16)|(dwRGB&0xff00L)|((dwRGB&0xff0000L)>>16)); }
    long SetSysColor(int nIndex);

    OLE_COLOR GetOleColor() const;
    COLORREF GetColorRef() const;
    DWORD   GetIntoRGB(void) const;
    TYPE    GetType() const;
    BOOL    IsSysColor() const { return ((_dwValue&MASK_FLAG)>=TYPE_SYSNAME && (_dwValue&MASK_FLAG)<TYPE_TRANSPARENT); }
    long    GetRawValue() const { return _dwValue; }

    HRESULT FromString(LPCTSTR pch, BOOL fValidOnly=FALSE, int iStrLen=-1);
    HRESULT RgbColor(LPCTSTR pch, int iStrLen);
    HRESULT HexColor(LPCTSTR pch, int iStrLen, BOOL fValidOnly);
    HRESULT NameColor(LPCTSTR pch);

    HRESULT Persist(IStream* pStream) const;
    HRESULT IsValid() const;

    // These are the 2 primary methods for interfacing
    BOOL IsDefined(void) const
    {
        return ((VALUE_UNDEF!=_dwValue) && ((_dwValue&MASK_FLAG)!=TYPE_TRANSPARENT));
    }
    BOOL IsNull(void) const
    {
        return VALUE_UNDEF==_dwValue;
    }
    void Undefine() { _dwValue = VALUE_UNDEF; }

    CColorValue& operator=(DWORD dwValue) { _dwValue = dwValue; return *this; }
    CColorValue(const CColorValue& other) { _dwValue = other._dwValue; }
};

#endif //__XINDOWS_CORE_UTIL_COLORVALUE_H__