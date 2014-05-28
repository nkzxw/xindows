
#ifndef __XINDOWS_CORE_UTIL_OFFSCREENCONTEXT_H__
#define __XINDOWS_CORE_UTIL_OFFSCREENCONTEXT_H__

interface IDirectDrawSurface;

//+------------------------------------------------------------------------
//
//  Class:      COffScreenContext
//
//  Synopsis:   Manages off-screen drawing context.
//
//-------------------------------------------------------------------------
class COffScreenContext
{
public:
	COffScreenContext(HDC hdcWnd, long width, long height, HPALETTE hpal, DWORD dwFlags);
	~COffScreenContext();

	HDC GetDC(RECT* prc);
	HDC ReleaseDC(HWND hwnd, BOOL fDraw=TRUE);


private:
	BOOL CreateDDB(long width, long height);
	BOOL GetDDB(HPALETTE hpal);
	void ReleaseDDB();
	BOOL CreateDDSurface(long width, long height, HPALETTE hpal);
	BOOL GetDDSurface(HPALETTE hpal);
	void ReleaseDDSurface();

	// actual dimensions
	long    _widthActual;
	long    _heightActual;

	RECT	_rc;		// requested position information (for viewport)
	HDC		_hdcMem;
	HDC		_hdcWnd;
	BOOL	_fOffScreen;
	long	_cBitsPixel;
	int		_nSavedDC;
	HBITMAP	_hbmMem;
	HBITMAP	_hbmOld;
	BOOL	_fCaret;
	BOOL	_fUseDD;
	BOOL	_fUse3D;
	IDirectDrawSurface* _pDDSurface;
};

enum
{
	OFFSCR_SURFACE	= 0x80000000,	// Use DD surface for offscreen buffer
	OFFSCR_3DSURFACE= 0x40000000,	// Use 3D DD surface for offscreen buffer
	OFFSCR_CARET	= 0x20000000,	// Manage caret around BitBlt's
	OFFSCR_BPP		= 0x000000FF	// bits-per-pixel mask
};

HRESULT InitSurface();

#endif //__XINDOWS_CORE_UTIL_OFFSCREENCONTEXT_H__