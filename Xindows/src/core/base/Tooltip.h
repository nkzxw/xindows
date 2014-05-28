
#ifndef __XINDOWS_CORE_BASE_TOOLTIP_H__
#define __XINDOWS_CORE_BASE_TOOLTIP_H__

void    FormsShowTooltip(TCHAR* szText, HWND hwnd, MSG msg, RECT* prc,
                         DWORD_PTR dwCookie, BOOL fRTL);
void    FormsHideTooltip(BOOL fReset);
BOOL    FormsTooltipMessage(UINT msg, WPARAM wParam, LPARAM lParam);

//+-------------------------------------------------------------------------
//
//  Class:      CTooltip
//
//  Synopsis:   Tooltip window
//
//--------------------------------------------------------------------------
class CTooltip
{
    friend void DeinitTooltip(THREADSTATE* pts);

public:
    DECLARE_MEMCLEAR_NEW_DELETE();

    CTooltip();
    ~CTooltip();

    HRESULT Init(HWND hwnd);

    void    Show(TCHAR* szText, HWND hwnd, MSG msg,  RECT* prc, DWORD_PTR dwCookie, BOOL fRTL);
    void    Hide();

private:
    HWND        _hwnd;
    HWND        _hwndOwner;
    DWORD_PTR   _dwCookie;
    CString     _cstrText;

    TOOLINFO    _ti;
};

#endif //__XINDOWS_CORE_BASE_TOOLTIP_H__