
#include "stdafx.h"
#include "OffScreenContext.h"

#include <ddraw.h>
#include "../../include/ddrawex.h"

const GUID CLSID_DirectDrawFactory = 
{ 0x4fd2a832, 0x86c8, 0x11d0, { 0x8f, 0xca, 0x0, 0xc0, 0x4f, 0xd9, 0x18, 0x9d }};

const GUID IID_IDirectDrawFactory = 
{ 0x4fd2a833, 0x86c8, 0x11d0, { 0x8f, 0xca, 0x0, 0xc0, 0x4f, 0xd9, 0x18, 0x9d }};


IDirectDraw*    g_pDirectDraw	= NULL;
HRESULT         g_hrDirectDraw	= S_OK;

PALETTEENTRY g_pal16[] = // maps to 16 colors for standard 4-bpp palette
{
	// Red  Green Blue  Flags
	{ 0x00, 0x00, 0x00, 0x00 },
	{ 0x80, 0x00, 0x00, 0x00 },
	{ 0x00, 0x80, 0x00, 0x00 },
	{ 0x80, 0x80, 0x00, 0x00 },
	{ 0x00, 0x00, 0x80, 0x00 },
	{ 0x80, 0x00, 0x80, 0x00 },
	{ 0x00, 0x80, 0x80, 0x00 },
	{ 0xC0, 0xC0, 0xC0, 0x00 },
	{ 0x80, 0x80, 0x80, 0x00 },
	{ 0xFF, 0x00, 0x00, 0x00 },
	{ 0x00, 0xFF, 0x00, 0x00 },
	{ 0xFF, 0xFF, 0x00, 0x00 },
	{ 0x00, 0x00, 0xFF, 0x00 },
	{ 0xFF, 0x00, 0xFF, 0x00 },
	{ 0x00, 0xFF, 0xFF, 0x00 },
	{ 0xFF, 0xFF, 0xFF, 0x00 }
};

PALETTEENTRY g_pal2[] = // maps to 2 colors for standard 1-bpp palette
{
	// Red  Green Blue  Flags
	{ 0x00, 0x00, 0x00, 0x00 },
	{ 0xFF, 0xFF, 0xFF, 0x00 }
};

CCriticalSection g_csOscCache;

#define DEFAULT_HEIGHT  150

struct
{
	SIZE		_sizeTgt;
	long		_areaTgt;
	SIZE		_size;
	long		_cHits;
	long		_cMisses;
	HDC			_hdcMem;
	HPALETTE	_hpal;
	long		_cBitsPixel;
	BOOL		_fInUse;
	HBITMAP		_hbmMem;
	HBITMAP		_hbmOld;
	BOOL		_fUseDD;
	BOOL		_fUse3D;
	IDirectDrawSurface* _pDDSurface;
} g_OscCache = { 0, DEFAULT_HEIGHT };

long GetCacheAreaTgt()
{
	if(g_OscCache._sizeTgt.cx == 0)
	{
		g_OscCache._sizeTgt.cx = GetSystemMetrics(SM_CXSCREEN);
		g_OscCache._areaTgt = g_OscCache._sizeTgt.cx * g_OscCache._sizeTgt.cy;
	}
	return g_OscCache._areaTgt;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::COffScreenContext
//
//  Synopsis:   Setup a DD surface or DDB for offscreen drawing.
//
//-------------------------------------------------------------------------
COffScreenContext::COffScreenContext(HDC hdcWnd, long width, long height, HPALETTE hpal, DWORD dwFlags)
{
	_fUseDD = !!(dwFlags & (OFFSCR_SURFACE|OFFSCR_3DSURFACE|OFFSCR_BPP));
	_fUse3D = !!(dwFlags & OFFSCR_3DSURFACE);

	GetCacheAreaTgt();

	Assert(hdcWnd);
	_hdcWnd = hdcWnd;
	_hdcMem = NULL;

	_nSavedDC = 0;
	_fOffScreen = FALSE;
	_hbmMem = NULL;
	_pDDSurface = NULL; 

	// zero width means use default cache dimensions
	if(width == 0)
	{
		_widthActual  = g_OscCache._sizeTgt.cx;
		_heightActual = g_OscCache._sizeTgt.cy;
	}
	else
	{
		_widthActual  = width;
		_heightActual = height;
	}

	_fCaret = !!(dwFlags & OFFSCR_CARET);
	_cBitsPixel = dwFlags & OFFSCR_BPP;

	switch(_cBitsPixel)
	{
	case 1:
	case 4:
	case 8:
	case 16:
	case 24:
	case 32:
		break;
	default:
		_cBitsPixel = GetDeviceCaps(_hdcWnd, PLANES) * GetDeviceCaps(_hdcWnd, BITSPIXEL);
		break;
	}

	Trace((_T("OffScreen construct - DD: %d width: %d height: %d BitsPixel: %d\n"),
		_fUseDD, _widthActual, _heightActual, _cBitsPixel));

	if(_cBitsPixel == 8)
	{
		if(hpal == NULL)
		{
			Assert(g_hpalHalftone);
			hpal = g_hpalHalftone;
		}
	}
	else
	{
		hpal = NULL;
	}

	// we either need to use one or we already have one (frankman)
	if(_fUseDD || g_OscCache._pDDSurface)
	{
		_fUseDD = TRUE; 
		if(!GetDDSurface(hpal))
		{
			return;
		}
	}
	else
	{
		if(!GetDDB(hpal))
		{
			return;
		}
	}

	// We have succesfully created the offscreen context.
	_fOffScreen = TRUE;

	// set palette
	if(hpal)
	{
		SelectPalette(_hdcMem, hpal, TRUE);
		RealizePalette(_hdcMem);
	}

	Trace((_T("OffScreen Cache - hits: %d misses: %d widthTgt: %d heightTgt: %d\n"),
		g_OscCache._cHits, g_OscCache._cMisses,
		g_OscCache._sizeTgt.cx, g_OscCache._sizeTgt.cy));

	return;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::GetDC
//
//  Synopsis:   Get the DC and establish view parameters.
//
//-------------------------------------------------------------------------
HDC COffScreenContext::GetDC(RECT* prc)
{
	if(_fOffScreen)
	{
		_rc = *prc;
		_nSavedDC = SaveDC(_hdcMem);
		Assert(_nSavedDC);
		SetViewportOrgEx(_hdcMem, -_rc.left, -_rc.top, (POINT*)NULL);
		return _hdcMem;
	}
	return _hdcWnd;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::ReleaseDC
//
//  Synopsis:   If painting offscreen, blt the bits to the screen.
//
//-------------------------------------------------------------------------
HDC COffScreenContext::ReleaseDC(HWND hwnd, BOOL fDraw/*=TRUE*/)
{
	if(_fOffScreen)
	{
		if(fDraw)
		{
			if(_fCaret)
			{
				::HideCaret(hwnd);
			}

			BitBlt(_hdcWnd,
				_rc.left, _rc.top,
				_rc.right-_rc.left,
				_rc.bottom-_rc.top,
				_hdcMem,
				_rc.left, _rc.top,
				SRCCOPY);

			if(_fCaret)
			{
				::ShowCaret(hwnd);
			}
		}

		if(_nSavedDC)
		{
			Verify(RestoreDC(_hdcMem, _nSavedDC));
			_nSavedDC = 0;
		}
	}
	return _hdcWnd;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::~COffScreenContext()
//
//  Synopsis:   Release resources used by this class.
//
//-------------------------------------------------------------------------
COffScreenContext::~COffScreenContext()
{
	_fOffScreen = FALSE;
	if(_nSavedDC)
	{
		// restore the DC for the cached case
		Verify(RestoreDC(_hdcMem, _nSavedDC));
		_nSavedDC = 0;
	}

	// Put back a reasonable palette so ours can be deleted safely
	SelectPalette(_hdcMem, (HPALETTE)GetStockObject(DEFAULT_PALETTE), TRUE);

	if(_fUseDD)
	{
		ReleaseDDSurface();
	}
	else
	{
		ReleaseDDB();
	}
}

//+------------------------------------------------------------------------
//
//  Function:   ClearSurfaceCache()
//
//  Synopsis:   Free the cached DD surface, if any.
//
//-------------------------------------------------------------------------
void ClearSurfaceCache()
{
	// Quick test without entering critical section, then retest afterwards in case it
	// changes.
	if(!g_OscCache._fInUse && g_OscCache._fUseDD && g_OscCache._pDDSurface!=NULL)
	{
		LOCK_SECTION(g_csOscCache);

		if(!g_OscCache._fInUse && g_OscCache._fUseDD && g_OscCache._pDDSurface!=NULL)
		{
			Trace0("surface cache deleted\n");
			if(g_OscCache._hdcMem)
			{
				SelectPalette(g_OscCache._hdcMem, (HPALETTE)GetStockObject(DEFAULT_PALETTE), TRUE);
				Verify(g_OscCache._pDDSurface->ReleaseDC(g_OscCache._hdcMem) == DD_OK);
				g_OscCache._hdcMem = NULL;
			}
			ClearInterface(&g_OscCache._pDDSurface);
			g_OscCache._fUseDD = FALSE;
		}
	}
}

//+------------------------------------------------------------------------
//
//  Function:   ClearDDBCache()
//
//  Synopsis:   Free the cached DDD, if any.
//
//-------------------------------------------------------------------------
void ClearDDBCache()
{
	if(!g_OscCache._fInUse && !g_OscCache._fUseDD && g_OscCache._hbmMem!=NULL)
	{
		Trace0("DDB cache deleted\n");
		if(g_OscCache._hdcMem)
		{
			if(g_OscCache._hbmOld)
			{
				SelectObject(g_OscCache._hdcMem, g_OscCache._hbmOld);
			}
			g_OscCache._hbmOld = NULL;
			DeleteDC(g_OscCache._hdcMem);
			g_OscCache._hdcMem = NULL;
		}
		DeleteObject(g_OscCache._hbmMem);
		g_OscCache._hbmMem = NULL;
	}
}

//+----------------------------------------------------------------------------
//
//  Function:   InitSurface
//
//  Synopsis:   Allocate DirectDraw object which is used to create surfaces
//              for offscreen rendering
//
//-----------------------------------------------------------------------------
HRESULT InitSurface()
{
	if(g_hrDirectDraw || g_pDirectDraw) // DD already initialized (or failed)
	{
		return g_hrDirectDraw;
	}

	LOCK_SECTION(g_csOscCache);

	// Check again after obtaining the lock
	// (since another thread could have made the attempt while we waited)
	if(g_hrDirectDraw || g_pDirectDraw)
    {
        return g_hrDirectDraw;
    }

	IDirectDrawFactory* pDDFactory = NULL;
	// create the DDFactory used to create the DirectDraw object
	g_hrDirectDraw = CoCreateInstance(CLSID_DirectDrawFactory,
		NULL, CLSCTX_INPROC_SERVER, IID_IDirectDrawFactory, (void**)&pDDFactory);
	if(SUCCEEDED(g_hrDirectDraw))
	{
		// DirectDraw will put up annoying UI if the bpp is less than 8.
		// Low bpp video cards are simply not support.  Force DirectDraw to
		// silently quietly fail under this circumstance.
		DWORD dwMode = SetErrorMode(SEM_FAILCRITICALERRORS);

		// create the actual DirectDraw object from which we can create surfaces later
		g_hrDirectDraw = pDDFactory->CreateDirectDraw(NULL, GetDesktopWindow(),
			DDSCL_NORMAL, 0, NULL, &g_pDirectDraw);
		SetErrorMode(dwMode);

		// release the factory
		ReleaseInterface((IUnknown*)pDDFactory);
	}

	RRETURN(g_hrDirectDraw);
}

//+----------------------------------------------------------------------------
//
//  Function:   DeinitSurface
//
//  Synopsis:   Release the DirectDraw surface factory and all our cached surfaces.
//
//-----------------------------------------------------------------------------
void DeinitSurface()
{
	LOCK_SECTION(g_csOscCache);
	ClearSurfaceCache();
	ClearInterface(&g_pDirectDraw);
	ClearDDBCache();
	Trace((_T("OffScreen Cache - hits: %d misses: %d widthTgt: %d heightTgt: %d\n"),
		g_OscCache._cHits, g_OscCache._cMisses,
		g_OscCache._sizeTgt.cx, g_OscCache._sizeTgt.cy));
}

DDPIXELFORMAT aPixelFormats[] =
{
	{ sizeof(DDPIXELFORMAT), DDPF_RGB|DDPF_PALETTEINDEXED1, 0, 1, 0, 0, 0 },
	{ sizeof(DDPIXELFORMAT), DDPF_RGB|DDPF_PALETTEINDEXED4, 0, 4, 0, 0, 0 },
	{ sizeof(DDPIXELFORMAT), DDPF_RGB|DDPF_PALETTEINDEXED8, 0, 8, 0, 0, 0, 0 }, 
	{ sizeof(DDPIXELFORMAT), DDPF_RGB, 0, 16, 0x0000F800, 0x000007E0, 0x0000001F, 0 },
	{ sizeof(DDPIXELFORMAT), DDPF_RGB, 0, 24, 0x00FF0000, 0x0000FF00, 0x000000FF, 0 },
	{ sizeof(DDPIXELFORMAT), DDPF_RGB, 0, 32, 0x00FF0000, 0x0000FF00, 0x000000FF, 0 }
};

//+------------------------------------------------------------------------
//
//  Function:   PixelFormat()
//
//  Synopsis:   Return a pointer to the DDPIXELFORMAT structure compatible
//              with the specified bits-per-pixel (and HDC, if 16-bpp).
//
//-------------------------------------------------------------------------
DDPIXELFORMAT* PixelFormat(HDC hdc, long cBitsPixel)
{
	static BOOL fMaskInit = FALSE;

	DDPIXELFORMAT* pPF = NULL;
	for(int i=0; i<ARRAYSIZE(aPixelFormats); i++)
	{
		if(aPixelFormats[i].dwRGBBitCount == (DWORD)cBitsPixel)
		{
			pPF = &aPixelFormats[i];
			break;
		}
	}

	if(cBitsPixel==16 && !fMaskInit)
	{
		// for 16-bit displays we need to decide if we're 555 or 565
		// cacheing this answer assumes we're only dealing with 16-bit
		// screen displays and that the bit format won't change on the fly
		fMaskInit = TRUE;
		HBITMAP hbm;
		struct
		{
			BITMAPINFOHEADER bih;
			DWORD bf[3];
		}   bi;
		hbm = CreateCompatibleBitmap(hdc, 1, 1);
		if(!hbm)
		{
			return NULL;
		}
		memset(&bi, 0, sizeof(bi));
		bi.bih.biSize = sizeof(BITMAPINFOHEADER);

		// first call will fill in the optimal biBitCount
		GetDIBits(hdc, hbm, 0, 1, NULL, (BITMAPINFO*)&bi, DIB_RGB_COLORS);
		if(bi.bih.biBitCount == 16)
		{
			// we'll default to the preset 5-6-5 if screen is not in 16-bit mode
			// second call will get the optimal bitfields
			GetDIBits(hdc, hbm, 0, 1, NULL, (BITMAPINFO*)&bi, DIB_RGB_COLORS);
			// Win95 only supports 5-5-5 and 5-6-5
			// For NT we'll assume this covers the majority cases too
			pPF->dwRBitMask = bi.bf[0]; // red mask
			pPF->dwGBitMask = bi.bf[1]; // green mask
			pPF->dwBBitMask = bi.bf[2]; // blue mask
		}
		else
		{
			Assert(0 && "invalid biBitCount");
		}
		DeleteObject(hbm);
	}
	return pPF;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::CreateDDSurface
//
//  Synopsis:   Create a DD surface with the specified dimensions and palette.
//
//-------------------------------------------------------------------------
BOOL COffScreenContext::CreateDDSurface(long width, long height, HPALETTE hpal)
{
	HRESULT hr = InitSurface();
	if(FAILED(hr))
	{
		return FALSE;
	}

	DDPIXELFORMAT* pPF = PixelFormat(_hdcWnd, _cBitsPixel);
	if(!pPF)
	{
		return FALSE;
	}

	DDSURFACEDESC ddsd;

	ddsd.dwSize = sizeof(ddsd);
	ddsd.ddpfPixelFormat = *pPF;
	ddsd.dwFlags = DDSD_CAPS|DDSD_HEIGHT|DDSD_WIDTH|DDSD_PIXELFORMAT;
	ddsd.ddsCaps.dwCaps = DDSCAPS_DATAEXCHANGE|DDSCAPS_OWNDC;
	if(_fUse3D)
	{
		ddsd.ddsCaps.dwCaps |= DDSCAPS_3DDEVICE;
	}
	ddsd.dwWidth = width;
	ddsd.dwHeight = height;
	hr = g_pDirectDraw->CreateSurface(&ddsd, &_pDDSurface, NULL);
	if(FAILED(hr))
	{
		return FALSE;
	}

	// set color table
	if(_cBitsPixel <= 8)
	{
		IDirectDrawPalette*	pDDPal;
		PALETTEENTRY*		pPal;
		PALETTEENTRY		pal256[256];
		long				cEntries;
		DWORD				pcaps;

		if(_cBitsPixel == 8)
		{
			cEntries = GetPaletteEntries(hpal, 0, 256, pal256);
			pPal = pal256;
			pcaps = DDPCAPS_8BIT;
		}
		else if(_cBitsPixel == 4)
		{
			cEntries = 16;
			pPal = g_pal16;
			pcaps = DDPCAPS_4BIT;
		}
		else if(_cBitsPixel == 1)
		{
			cEntries = 2;
			pPal = g_pal2;
			pcaps = DDPCAPS_1BIT;
		}
		else
		{
			Assert(0 && "invalid cBitsPerPixel");
			return FALSE;
		}

		// create and initialize a new DD palette
		hr = g_pDirectDraw->CreatePalette(pcaps|DDPCAPS_INITIALIZE, pPal, &pDDPal, NULL);
		if(SUCCEEDED(hr))
		{
			// attach the DD palette to the DD surface
			hr = _pDDSurface->SetPalette(pDDPal);
			pDDPal->Release();
		}
		if(FAILED(hr))
		{
			return FALSE;
		}
	}

	hr = _pDDSurface->GetDC(&_hdcMem);
	return SUCCEEDED(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::GetDDSurface
//
//  Synopsis:   Create a suitable DD surface or get one from the cache.
//
//-------------------------------------------------------------------------
BOOL COffScreenContext::GetDDSurface(HPALETTE hpal)
{
	HRESULT hr;
	LOCK_SECTION(g_csOscCache);

	if(!g_OscCache._fInUse)
	{
		ClearDDBCache(); // don't allow both DD & DDB in the cache

		if(g_OscCache._pDDSurface != NULL)
		{
			Assert(g_OscCache._fUseDD);
			if(_widthActual>g_OscCache._size.cx
				|| _heightActual>g_OscCache._size.cy
				|| _cBitsPixel!=g_OscCache._cBitsPixel
				|| (_fUse3D&&!g_OscCache._fUse3D))
			{
				++g_OscCache._cMisses;
				ClearSurfaceCache();
			}
			else
			{
				++g_OscCache._cHits;
				g_OscCache._fInUse = TRUE;
				_pDDSurface = g_OscCache._pDDSurface;
				_pDDSurface->AddRef();
				_hdcMem = g_OscCache._hdcMem;
				_widthActual = g_OscCache._size.cx;
				_heightActual = g_OscCache._size.cy;
			}
		}
		if(g_OscCache._pDDSurface == NULL)
		{
			// use max area allowed for the cache, so we get max reuse potential
			// favor width over height for those wide text runs
			// also adjust the max size in a growing fashion based on former allocations
			g_OscCache._size.cx = max(g_OscCache._size.cx, max(min(g_OscCache._areaTgt/_heightActual, g_OscCache._sizeTgt.cx), _widthActual));
			g_OscCache._size.cy = max(g_OscCache._size.cy , max(g_OscCache._areaTgt/g_OscCache._size.cx, _heightActual));
			if(!CreateDDSurface(g_OscCache._size.cx, g_OscCache._size.cy, hpal))
			{
				return FALSE;
			}
			Trace0("surface cache created\n");
			g_OscCache._fInUse = TRUE;
			g_OscCache._fUseDD = TRUE;
			g_OscCache._fUse3D = _fUse3D;
			g_OscCache._pDDSurface = _pDDSurface;
			g_OscCache._pDDSurface->AddRef(); // addref the global surface
			g_OscCache._cBitsPixel = _cBitsPixel;
			g_OscCache._hdcMem = _hdcMem;
			g_OscCache._hpal = hpal;
			_widthActual = g_OscCache._size.cx;
			_heightActual = g_OscCache._size.cy;
		}
	}
	else
	{
		Trace0("surface cache in use\n");
	}

	if(_pDDSurface == NULL)
	{
		if(!CreateDDSurface(_widthActual, _heightActual, hpal))
		{
			return FALSE;
		}
	}

	// reset the color table when using cache
	if(_cBitsPixel==8 && _pDDSurface==g_OscCache._pDDSurface
		&& (hpal!=g_OscCache._hpal || hpal!=g_hpalHalftone))
	{
		IDirectDrawPalette* pDDPal;
		PALETTEENTRY pal256[256];
		long cEntries;

		cEntries = GetPaletteEntries(hpal, 0, 256, pal256);

		// get the DD palette and set entries
		hr = _pDDSurface->GetPalette(&pDDPal);
		if(SUCCEEDED(hr))
		{
			hr = pDDPal->SetEntries(0, 0, cEntries, pal256);
			pDDPal->Release();
			if(SUCCEEDED(hr))
			{
				g_OscCache._hpal = hpal;
			}
		}
	}

	return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::ReleaseDDSurface
//
//  Synopsis:   Release the DD surface or return it to the cache.
//
//-------------------------------------------------------------------------
void COffScreenContext::ReleaseDDSurface()
{
	if(_pDDSurface == NULL)
	{
		return;
	}

	if(_pDDSurface == g_OscCache._pDDSurface)
	{
		// keep the surface cache; mark as available
		LOCK_SECTION(g_csOscCache);

		Assert(g_OscCache._fInUse);
		g_OscCache._fInUse = FALSE; 

		// If some rogue client left the surface locked we could be locked out for a long time.
		// So, to be more robust, we attempt to lock the surface and if we fail, we discard the
		// cache so we can start fresh next time.
		DDSURFACEDESC desc;
		desc.dwSize = sizeof(desc);

		HRESULT hr = _pDDSurface->Lock(NULL, &desc, DDLOCK_SURFACEMEMORYPTR|DDLOCK_WAIT, NULL);
		if(hr == S_OK)
		{
			_pDDSurface->Unlock(NULL);
		}
		else
		{
			Assert(0 && "DD surface left locked!");
			// release the surface cache
			ClearSurfaceCache();
		}
		_pDDSurface->Release();
	}
	else
	{
		// free the non-cached surface
		if(_hdcMem)
		{
			_pDDSurface->ReleaseDC(_hdcMem);
			_hdcMem = NULL;
		}
		_pDDSurface->Release();
	}
	_pDDSurface = NULL;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::CreateDDB
//
//  Synopsis:   Create a DDB with the specified dimensions.
//
//-------------------------------------------------------------------------
BOOL COffScreenContext::CreateDDB(long width, long height)
{
	_hbmMem = CreateCompatibleBitmap(_hdcWnd, width, height);
	if(!_hbmMem)
	{
		return FALSE;
	}

	_hdcMem = CreateCompatibleDC(_hdcWnd);
	if(!_hdcMem)
	{
		return FALSE;
	}

	_hbmOld = (HBITMAP)SelectObject(_hdcMem, _hbmMem);
	if(!_hbmOld)
	{
		return FALSE;
	}

	return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::GetDDB
//
//  Synopsis:   Create a suitable DDB or get one from the cache.
//
//-------------------------------------------------------------------------
BOOL COffScreenContext::GetDDB(HPALETTE hpal)
{
	LOCK_SECTION(g_csOscCache);

	if(!g_OscCache._fInUse)
	{
		ClearSurfaceCache(); // don't allow both DD & DDB in the cache

		if(g_OscCache._hbmMem != NULL)
		{
			if(_widthActual>g_OscCache._size.cx || _heightActual>g_OscCache._size.cy
				|| _cBitsPixel!=g_OscCache._cBitsPixel)
			{
				++g_OscCache._cMisses;
				ClearDDBCache();
			}
			else
			{
				++g_OscCache._cHits;
				g_OscCache._fInUse = TRUE;
				_hbmMem = g_OscCache._hbmMem;
				_hdcMem = g_OscCache._hdcMem;
				_widthActual = g_OscCache._size.cx;
				_heightActual = g_OscCache._size.cy;
			}
		}
		if(g_OscCache._hbmMem==NULL && _widthActual*_heightActual<=g_OscCache._areaTgt)
		{
			// use max area allowed for the cache, so we get max reuse potential
			// favor width over height for those wide text runs
			g_OscCache._size.cx  = max(min(g_OscCache._areaTgt/_heightActual, g_OscCache._sizeTgt.cx), _widthActual);
			g_OscCache._size.cy = max(g_OscCache._areaTgt/g_OscCache._size.cx, _heightActual);
			if(!CreateDDB(g_OscCache._size.cx, g_OscCache._size.cy))
			{
				return FALSE;
			}
			Trace0("DDB cache created\n");
			g_OscCache._fInUse = TRUE;
			g_OscCache._hbmMem = _hbmMem;
			g_OscCache._hbmOld = _hbmOld;
			g_OscCache._cBitsPixel = _cBitsPixel;
			g_OscCache._hdcMem = _hdcMem;
			g_OscCache._hpal = hpal;
			_widthActual = g_OscCache._size.cx;
			_heightActual = g_OscCache._size.cy;
		}
	}
	else
	{
		Trace0("DDB cache in use\n");
	}

	if(_hbmMem == NULL)
	{
		if(!CreateDDB(_widthActual, _heightActual))
		{
			return FALSE;
		}
	}

	return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     COffScreenContext::ReleaseDDB
//
//  Synopsis:   Release the DDB or return it to the cache.
//
//-------------------------------------------------------------------------
void COffScreenContext::ReleaseDDB()
{
	if(_hbmMem == NULL)
	{
		return;
	}

	if(_hbmMem == g_OscCache._hbmMem)
	{
		// keep the DDB cache; mark as available
		LOCK_SECTION(g_csOscCache);
		Assert(g_OscCache._fInUse);
		g_OscCache._fInUse = FALSE;
	}
	else
	{
		// Prevent Windows from RIPing when we delete our palette later on
		SelectPalette(_hdcMem, (HPALETTE)GetStockObject(DEFAULT_PALETTE), TRUE);

		// free the non-cached DDB
		if(_hbmOld)
		{
			SelectObject(_hdcMem, _hbmOld);
		}
		_hbmOld = NULL;
		DeleteDC(_hdcMem);
		_hdcMem = NULL;
		DeleteObject(_hbmMem);
	}
	_hbmMem = NULL;
}