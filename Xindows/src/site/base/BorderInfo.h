
#ifndef __XINDOWS_SITE_BASE_BORDERINFO_H__
#define __XINDOWS_SITE_BASE_BORDERINFO_H__

//+---------------------------------------------------------------------------
//
//  Class:      CBorderInfo
//
//  Purpose:    Class used to hold a collection of border information.  This is
//              really just a struct with a constructor.
//
//----------------------------------------------------------------------------
class CBorderInfo
{
public:
	CBorderInfo() { Init(); }
	CBorderInfo(BOOL fDummy) { /* do nothing - work postponed */ }
	void Init()
	{
		memset(this, 0, sizeof(CBorderInfo));
		// Have to set up some default widths.
		CUnitValue cuv;
		cuv.SetValue(4/*MEDIUM*/, CUnitValue::UNIT_PIXELS);
		aiWidths[BORDER_TOP] = aiWidths[BORDER_BOTTOM] =
			cuv.GetPixelValue(NULL, CUnitValue::DIRECTION_CY, 0, 0);
		aiWidths[BORDER_RIGHT] = aiWidths[BORDER_LEFT] =
			cuv.GetPixelValue(NULL, CUnitValue::DIRECTION_CX, 0, 0);
	}

	// The collection of border data:
	BOOL	fNotDraw;		// do not draw
	// These values reflect essentials of the border
	// (These are those needed by CSite::GetClientRect)
	BYTE	abStyles[4];	// top,right,bottom,left
	int		aiWidths[4];	// top,right,bottom,left
	WORD	wEdges;			// which edges to draw

	// These values reflect all remaining details
	// (They are set only if GetBorderInfo is called with fAll==TRUE)
	RECT	rcSpace;		// top, bottom, right and left space
	SIZE	sizeCaption;	// location of the caption
	int		offsetCaption;	// caption offset on top
	int		xyFlat;			// hack to support combined 3d and flat
	COLORREF acrColors[4][3]; // top,right,bottom,left:base color, inner color, outer color
};

void DrawBorder(CFormDrawInfo* pDI, LPRECT lprc, CBorderInfo* pborderinfo);

#endif //__XINDOWS_SITE_BASE_BORDERINFO_H__