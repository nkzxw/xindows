
#ifndef __XINDOWS_CORE_UTIL_WINDOWSUTIL_H__
#define __XINDOWS_CORE_UTIL_WINDOWSUTIL_H__

HRESULT RegisterWindowClass(
            TCHAR*      pstrClass,
            LRESULT     (CALLBACK *pfnWndProc)(HWND, UINT, WPARAM, LPARAM),
            UINT        style,
            TCHAR*      pstrBase,
            WNDPROC*    ppfnBase,
            ATOM*       patom,
            HICON       hIconSm=NULL);

void    UpdateChildTree(HWND hWnd);
HCURSOR SetCursorIDC(LPCTSTR idr);

//+------------------------------------------------------------------------
//
//  Class:      CCurs (Curs)
//
//  Purpose:    System cursor stack wrapper class.  Creating one of these
//              objects pushes a new system cursor (the wait cursor, the
//              arrow, etc) on top of a stack; destroying the object
//              restores the old cursor.
//
//  Interface:  Constructor     Pushes a new cursor
//              Destructor      Pops the old cursor back
//
//-------------------------------------------------------------------------
class CCurs
{
public:
    CCurs(LPCTSTR idr);
    ~CCurs();

private:
    HCURSOR _hcrs;
    HCURSOR _hcrsOld;
};

//+---------------------------------------------------------------------
//
//  VB helpers
//
//----------------------------------------------------------------------
#define VB_LBUTTON  1
#define VB_RBUTTON  2
#define VB_MBUTTON  4

#define VB_SHIFT    1
#define VB_CONTROL  2
#define VB_ALT      4

short VBButtonState(WPARAM);
short VBShiftState();
short VBShiftState(DWORD grfKeyState);

#endif //__XINDOWS_CORE_UTIL_WINDOWSUTIL_H__