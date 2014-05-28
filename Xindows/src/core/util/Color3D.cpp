
#include "stdafx.h"
#include "Color3D.h"

static COLORREF DoGetColor(BOOL fUseSystem, COLORREF coColor, int lColorIndex)
{
	return fUseSystem?GetSysColorQuick(lColorIndex):coColor;
}

COLORREF ThreeDColors::BtnFace()
{
	return DoGetColor(_fUseSystem, _coBtnFace, COLOR_BTNFACE);
}

COLORREF ThreeDColors::BtnLight()
{
    return DoGetColor( _fUseSystem, _coBtnLight, COLOR_3DLIGHT);
}

COLORREF ThreeDColors::BtnShadow()
{
	return DoGetColor(_fUseSystem, _coBtnShadow, COLOR_BTNSHADOW);
}

COLORREF ThreeDColors::BtnHighLight()
{
	return DoGetColor(_fUseSystem, _coBtnHighLight, COLOR_BTNHIGHLIGHT);
}

COLORREF ThreeDColors::BtnDkShadow()
{
	return DoGetColor(_fUseSystem, _coBtnDkShadow, COLOR_3DDKSHADOW);
}

COLORREF ThreeDColors::BtnText()
{
	return DoGetColor(TRUE, 0, COLOR_BTNTEXT);
}

static inline HBRUSH DoGetBrush(BOOL fUseSystem, COLORREF coColor)
{
	return GetCachedBrush(coColor);
}

HBRUSH ThreeDColors::BrushBtnFace()
{
	return DoGetBrush(_fUseSystem, BtnFace());
}

HBRUSH ThreeDColors::BrushBtnLight()
{
	return DoGetBrush(_fUseSystem, BtnLight());
}

HBRUSH ThreeDColors::BrushBtnShadow()
{
	return DoGetBrush(_fUseSystem, BtnShadow());
}

HBRUSH ThreeDColors::BrushBtnHighLight()
{
	return DoGetBrush(_fUseSystem, BtnHighLight());
}

HBRUSH ThreeDColors::BrushBtnText()
{
	return DoGetBrush(TRUE, BtnText());
}

//+--------------------------------------------------------------------
//
//  Functions:  LightenColor & DarkenColor
//
//  Synopsis:   These take RGBs and return RGBs which are "lighter" or
//              "darker"
//
//---------------------------------------------------------------------
static unsigned Increase(unsigned x)
{
	x = x * 5 / 3;

	return x>255?255:x;
}

static unsigned Decrease(unsigned x)
{
	return x * 3 / 5;
}

static COLORREF LightenColor(COLORREF cr)
{
    return
        (Increase((cr&0x000000FF)>> 0)<< 0) |
        (Increase((cr&0x0000FF00)>> 8)<< 8) |
        (Increase((cr&0x00FF0000)>>16)<<16);
}

static COLORREF DarkenColor(COLORREF cr)
{
    return
        (Decrease((cr&0x000000FF)>> 0)<< 0) |
        (Decrease((cr&0x0000FF00)>> 8)<< 8) |
        (Decrease((cr&0x00FF0000)>>16)<<16);
}

//+--------------------------------------------------------------------
//
//  Member:     ThreeDColors::SetBaseColor
//
//  Synopsis:   This is called to reestablish the 3D colors, based on a
//              single color.
//
//---------------------------------------------------------------------
void ThreeDColors::SetBaseColor(OLE_COLOR coBaseColor)
{
	// Sentinal color (0xffffffff) means use default which is button face
	_fUseSystem = (coBaseColor&0x80ffffff) == DEFAULT_BASE_COLOR;

	if(_fUseSystem)
	{
		return;
	}

	// Ok, now synthesize some colors! 
	//
	// First, use the base color as the button face color
	_coBtnFace = ColorRefFromOleColor(coBaseColor);

	_coBtnLight = _coBtnFace;

	// Dark shadow is always black and button face
	// (or so Win95 seems to indicate)
	_coBtnDkShadow = 0;

	// Now, lighten/darken colors
	_coBtnShadow = DarkenColor(_coBtnFace);

	HDC hdc = TLS(hdcDesktop);

	if(!hdc)
	{
		return;
	}

	COLORREF coRealBtnFace = GetNearestColor(hdc, _coBtnFace);

	_coBtnHighLight = LightenColor(_coBtnFace);

	if(GetNearestColor(hdc, _coBtnHighLight) == coRealBtnFace)
	{
		_coBtnHighLight = LightenColor(_coBtnHighLight);

		if(GetNearestColor( hdc, _coBtnHighLight ) == coRealBtnFace)
		{
			_coBtnHighLight = RGB(255, 255, 255);
		}
	}

	_coBtnShadow = DarkenColor(_coBtnFace);

	if(GetNearestColor(hdc, _coBtnShadow) == coRealBtnFace)
	{
		_coBtnShadow = DarkenColor(_coBtnFace);

		if(GetNearestColor(hdc, _coBtnShadow) == coRealBtnFace)
		{
			_coBtnShadow = RGB(0, 0, 0);
		}
	}
}

void ThreeDColors::NoDither()
{
	_coBtnFace |= 0x02000000;
	_coBtnLight |= 0x02000000;
	_coBtnShadow |= 0x02000000;
	_coBtnHighLight |= 0x02000000;
	_coBtnDkShadow |= 0x02000000;
}