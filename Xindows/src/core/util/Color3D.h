
#ifndef __XINDOWS_CORE_UTIL_COLOR3D_H__
#define __XINDOWS_CORE_UTIL_COLOR3D_H__

//+---------------------------------------------------------------------------
//
//  Class:      ThreeDColors (c3d)
//
//  Purpose:    Synthesizes 3D beveling colors
//
//  History:    21-Mar-95   EricVas      Created
//
//----------------------------------------------------------------------------
#define OLECOLOR_FROM_SYSCOLOR(__c)     ((__c)|0x80000000)
#define DEFAULT_BASE_COLOR              OLECOLOR_FROM_SYSCOLOR(COLOR_BTNFACE)

class ThreeDColors
{
public:
	enum INITIALIZER { NO_COLOR };
	ThreeDColors(INITIALIZER) {}

	ThreeDColors(OLE_COLOR coInitialBaseColor=DEFAULT_BASE_COLOR)
	{
		SetBaseColor(coInitialBaseColor);
	}

	void SetBaseColor(OLE_COLOR);
	void NoDither();

	// Use these to fetch the synthesized colors
	COLORREF BtnFace();
	COLORREF BtnLight();
	COLORREF BtnShadow();
	COLORREF BtnHighLight();
	COLORREF BtnDkShadow();

	virtual COLORREF BtnText();

	// Use these to fetch cached brushes (must call ReleaseCachedBrush)
	// when finished.
	HBRUSH BrushBtnFace();
	HBRUSH BrushBtnLight();
	HBRUSH BrushBtnShadow();
	HBRUSH BrushBtnHighLight();

	virtual HBRUSH BrushBtnText();

private:
	BOOL _fUseSystem;

	COLORREF _coBtnFace;
	COLORREF _coBtnLight;
	COLORREF _coBtnShadow;
	COLORREF _coBtnHighLight;
	COLORREF _coBtnDkShadow;
};

#endif //__XINDOWS_CORE_UTIL_COLOR3D_H__