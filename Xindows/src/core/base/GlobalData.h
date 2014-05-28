
#ifndef __XINDOWS_CORE_BASE_GLOBALDATA_H__
#define __XINDOWS_CORE_BASE_GLOBALDATA_H__

union ZERO_STRUCTS
{
    POINT       pt;
    POINTL      ptl;
    SIZE        size;
    SIZEL       sizel;
    RECTL       rcl;
    RECT        rc;
    TCHAR       ach[1];
    OLECHAR     achole[1];
    CHAR        acha[1];
    char        rgch[1];
    GUID        guid;
    IID         iid;
    CLSID       clsid;
    DISPPARAMS  dispparams;
    BYTE        ab[256];
};

// hold for number shaping used by system
typedef enum NUMSHAPE
{
    NUMSHAPE_CONTEXT = 0,
    NUMSHAPE_NONE    = 1,
    NUMSHAPE_NATIVE  = 2,
};

struct THREADSTATE;
struct GlobalData
{
    ZERO_STRUCTS    _Zero;
    DWORD           _dwTls;                     // TLS index associated with THREADSTATEs
	THREADSTATE*    _pts;                       // Head of THREADSTATE chain

    HINSTANCE	    _hInstThisModule;           // Module Instance
    TCHAR           _szModulePath[MAX_PATH];
    TCHAR           _szModuleFolder[MAX_PATH];
	HINSTANCE	    _hInstResource;             // Resource Instance

    LCID            _lcidUserDefault;           // Locale Information
    UINT            _cpDefault;
    BOOL            _fFarEastWinNT;
    BOOL            _fBidiSupport;              // COMPLEXSCRIPT
    BOOL            _fComplexScriptInput;

    NUMSHAPE        _iNumShape;
    DWORD           _uLangNationalDigits;

    BOOL            _fHighContrastMode;

    BOOL            _fSystemFontsNeedRefreshing;
    BOOL            _fNoDisplayChange;

    SIZE            _sizeScrollbar;
    SIZEL           _sizelScrollbar;
    SIZE            _sizePixelsPerInch;
};

extern GlobalData _afxGlobalData;

#endif //__XINDOWS_CORE_BASE_GLOBALDATA_H__