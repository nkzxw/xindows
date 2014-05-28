
#include "stdafx.h"
#include "BorderInfo.h"

//+-------------------------------------------------------------------------
//
//  Function:   DrawBorder
//
//  Synopsis:   This routine is functionally equivalent to the Win95
//              DrawEdge API
//  xyFlat:     positive means, we draw flat inside, outside if negative
//
//--------------------------------------------------------------------------
#define DRAWBORDER_OUTER    0
#define DRAWBORDER_SPACE    1
#define DRAWBORDER_INNER    2
#define DRAWBORDER_TOTAL    3
#define DRAWBORDER_INCRE    4
#define DRAWBORDER_LAST     5

void DrawBorder(CFormDrawInfo* pDI, LPRECT lprc, CBorderInfo* pborderinfo)
{
	HDC			hdc = pDI->GetDC(TRUE);
	RECT		rc;
	RECT		rcSave; // Original rectangle, maybe adjusted for flat border.
	RECT		rcOrig; // Original rectangle of the element container.
	HBRUSH		hbrOld = NULL;
	COLORREF	crNow = CLR_INVALID;
	COLORREF	crNew;
	HPEN		hPen;
	// Ordering of the next 2 arrays are top,right,bottom,left.
	BYTE		abCanUseJoin[4];
	int			aiNumBorderLines[4];
	int			aiLineWidths[4][DRAWBORDER_LAST]; // outer, spacer, inner
	UINT		wEdges = pborderinfo->wEdges;
	int			i, j;
	POINT		polygon[8];
	int			xyFlat = abs(pborderinfo->xyFlat);
	BOOL		fdrawCaption = pborderinfo->sizeCaption.cx || pborderinfo->sizeCaption.cy;
	BOOL		fContinuingPoly;
	SIZE		sizeLegend;
	SIZE		sizeRect;
	POINT		ptBrushOrg;

	// save legend size
	sizeLegend = pborderinfo->sizeCaption;

	Assert(pborderinfo);
	if(pborderinfo->fNotDraw)
	{
		return;
	}

	// offsetCaption is already transformed in GetBorderInfo
	rc.top = lprc->top + pDI->DocPixelsFromWindowY(pborderinfo->rcSpace.top)
		+ pborderinfo->offsetCaption;
	rc.bottom = lprc->bottom
		- pDI->DocPixelsFromWindowY(pborderinfo->rcSpace.bottom);
	rc.left = lprc->left
		+ pDI->DocPixelsFromWindowX(pborderinfo->rcSpace.left);
	rc.right = lprc->right
		- pDI->DocPixelsFromWindowX(pborderinfo->rcSpace.right);

	rcOrig = rc;

	Assert(rc.left <= rc.right);
	Assert(rc.top <= rc.bottom);

	if(pborderinfo->xyFlat < 0)
	{
		InflateRect(&rc, pborderinfo->xyFlat, pborderinfo->xyFlat);
	}

	rcSave = rc;

	hPen = (HPEN)GetStockObject(NULL_PEN);
	SelectObject(hdc, hPen);

	for(i=0; i<4; i++)
	{
		aiLineWidths[i][DRAWBORDER_OUTER] = pborderinfo->aiWidths[i] - xyFlat;
		aiLineWidths[i][DRAWBORDER_TOTAL] = max(aiLineWidths[i][DRAWBORDER_OUTER], 1);
	}

	sizeRect.cx = rc.right - rc.left;
	sizeRect.cy = rc.bottom - rc.top;

	// set brush origin so that dither patterns are anchored correctly
	::GetViewportOrgEx(hdc, &ptBrushOrg);
	ptBrushOrg.x += rc.left;
	ptBrushOrg.y += rc.top;

	// validate border size
	// if the broders are too big, truncate the bottom and right parts
	if(aiLineWidths[BORDER_TOP][DRAWBORDER_OUTER]+
		aiLineWidths[BORDER_BOTTOM][DRAWBORDER_OUTER] > sizeRect.cy)
	{
		aiLineWidths[BORDER_TOP][DRAWBORDER_OUTER] =
			(aiLineWidths[BORDER_TOP][DRAWBORDER_OUTER]>sizeRect.cy) ?
			sizeRect.cy : aiLineWidths[BORDER_TOP][DRAWBORDER_OUTER];
		aiLineWidths[BORDER_BOTTOM][DRAWBORDER_OUTER] =
			sizeRect.cy - aiLineWidths[BORDER_TOP][DRAWBORDER_OUTER];
	}

	if(aiLineWidths[BORDER_RIGHT][DRAWBORDER_OUTER]+
		aiLineWidths[BORDER_LEFT][DRAWBORDER_OUTER] > sizeRect.cx)
	{
		aiLineWidths[BORDER_LEFT][DRAWBORDER_OUTER] =
			(aiLineWidths[BORDER_LEFT][DRAWBORDER_OUTER]>sizeRect.cx) ?
			sizeRect.cx : aiLineWidths[BORDER_LEFT][DRAWBORDER_OUTER];
		aiLineWidths[BORDER_RIGHT][DRAWBORDER_OUTER] =
			sizeRect.cx - aiLineWidths[BORDER_LEFT][DRAWBORDER_OUTER];
	}

	for(i=0; i<4; i++)
	{
		aiLineWidths[i][DRAWBORDER_TOTAL] = max(aiLineWidths[i][DRAWBORDER_OUTER], 1);
	}

	for(i=0; i<4; i++)
	{
		switch(pborderinfo->abStyles[i])
		{
		case fmBorderStyleNone:
			aiLineWidths[i][DRAWBORDER_TOTAL] = 0;

		case fmBorderStyleSingle:
		case fmBorderStyleRaisedMono:
		case fmBorderStyleSunkenMono:
		case fmBorderStyleDotted:
		case fmBorderStyleDashed:
			aiNumBorderLines[i] = 1;
			aiLineWidths[i][DRAWBORDER_INNER] = 0;
			aiLineWidths[i][DRAWBORDER_SPACE] = 0;
			abCanUseJoin[i] = !!(aiLineWidths[i][DRAWBORDER_OUTER] > 0);
			break;

		case fmBorderStyleRaised:
		case fmBorderStyleSunken:
		case fmBorderStyleEtched:
		case fmBorderStyleBump:
			aiLineWidths[i][DRAWBORDER_SPACE] = 0; // Spacer
			aiNumBorderLines[i] = 3;
			aiLineWidths[i][DRAWBORDER_INNER] = aiLineWidths[i][DRAWBORDER_OUTER] >> 1;
			aiLineWidths[i][DRAWBORDER_OUTER] -= aiLineWidths[i][DRAWBORDER_INNER];
			abCanUseJoin[i] = !!(aiLineWidths[i][DRAWBORDER_TOTAL] > 1);
			break;

		case fmBorderStyleDouble:
			aiNumBorderLines[i] = 3;
			aiLineWidths[i][DRAWBORDER_SPACE] = aiLineWidths[i][DRAWBORDER_OUTER] / 3; // Spacer
			// If this were equal to the line above,
			aiLineWidths[i][DRAWBORDER_INNER] = (aiLineWidths[i][DRAWBORDER_OUTER]
			- aiLineWidths[i][DRAWBORDER_SPACE])/2;
			// we'd get widths of 3,1,1 instead of 2,1,2
			aiLineWidths[i][DRAWBORDER_OUTER] -= aiLineWidths[i][DRAWBORDER_SPACE]
			+ aiLineWidths[i][DRAWBORDER_INNER];
			// evaluate border join
			// we don't want to have a joint polygon borders if
			// the distribution does not match
			abCanUseJoin[i] =
				pborderinfo->acrColors[(i==0)?3:(i-1)][0]==pborderinfo->acrColors[i][0] &&
				(aiLineWidths[i][DRAWBORDER_TOTAL]>2) &&
				(aiLineWidths[(i==0)?3:(i-1)][DRAWBORDER_TOTAL]%3)==(aiLineWidths[i][DRAWBORDER_TOTAL]%3);
			break;
		}
		abCanUseJoin[i] = abCanUseJoin[i] &&
			pborderinfo->abStyles[(4+i-1)%4]==pborderinfo->abStyles[i];
		aiLineWidths[i][DRAWBORDER_INCRE] = 0;
	}

	// Loop: draw outer lines (j=0), spacer (j=1), inner lines (j=2)
	for(j=DRAWBORDER_OUTER; j<=DRAWBORDER_INNER; j++)
	{
		fContinuingPoly = FALSE;
		if(j != DRAWBORDER_SPACE) // if j==1, this line is a spacer only - don't draw lines, just deflate the rect.
		{
			i = 0;
			// We'll work around the border edges CW, starting with the right edge, in an attempt to reduce calls to polygon().
			// we must draw left first to follow Windows standard
			if(wEdges&BF_LEFT && (j<aiNumBorderLines[BORDER_LEFT]) && aiLineWidths[BORDER_LEFT][j])
			{
				// There's a left edge
				// get color brush for left side
				crNew = pborderinfo->acrColors[BORDER_LEFT][j];
				if(crNew != crNow)
				{
					HBRUSH hbrNew;
					SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
					// not supported on WINCE
					::UnrealizeObject(hbrNew);
					::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
				}

				// build left side polygon
				polygon[i].x = rc.left + aiLineWidths[BORDER_LEFT][j]; // lower right corner
				polygon[i++].y = rcSave.bottom - MulDivQuick(aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE]+aiLineWidths[BORDER_LEFT][j],
					aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL]);
				polygon[i].x = rc.left; // lower left corner
				polygon[i++].y = rcSave.bottom - MulDivQuick(aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE],
					aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL]);

				if(!(wEdges&BF_TOP) ||
					( pborderinfo->acrColors[BORDER_LEFT][j]!=pborderinfo->acrColors[BORDER_TOP][j])
					|| !abCanUseJoin[BORDER_TOP])
				{
					polygon[i].x = rc.left; // upper left corner
					polygon[i++].y = rcSave.top + MulDivQuick(aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL]);
					polygon[i].x = rc.left + aiLineWidths[BORDER_LEFT][j]; // upper right corner
					polygon[i++].y = rcSave.top + MulDivQuick(aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE]+aiLineWidths[BORDER_LEFT][j],
						aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL]);
					Polygon(hdc, polygon, i);
					i = 0;
				}
				else
				{
					fContinuingPoly = TRUE;
				}
			}
			if(wEdges&BF_TOP && (j<aiNumBorderLines[BORDER_TOP]) && aiLineWidths[BORDER_TOP][j])
			{
				// There's a top edge
				if(!fContinuingPoly)
				{
					// get color brush for top side
					crNew = pborderinfo->acrColors[BORDER_TOP][j];
					if(crNew != crNow)
					{
						HBRUSH hbrNew;
						SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
						// not supported on WINCE
						::UnrealizeObject(hbrNew);
						::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
						i = 0;
					}
				}
				// build top side polygon

				// up left
				polygon[i].x = rcSave.left + ((wEdges&BF_LEFT)?
					MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE],
					aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL]):0);
				polygon[i++].y  = rc.top;

				if(fdrawCaption)
				{
					// shrink legend
					sizeLegend.cx = sizeLegend.cx
						- aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE];
					sizeLegend.cy = sizeLegend.cy
						- aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE];

					polygon[i].x    = rc.left + sizeLegend.cx;
					polygon[i++].y  = rc.top;
					polygon[i].x    = rc.left + sizeLegend.cx;
					polygon[i++].y  = rc.top + aiLineWidths[BORDER_TOP][j];

					polygon[i].x    = rcSave.left + ((wEdges&BF_LEFT)?
						MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE]+aiLineWidths[BORDER_TOP][j],
						aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL]):0);

					polygon[i++].y  = rc.top + aiLineWidths[BORDER_TOP][j];

					Polygon(hdc, polygon, i);
					i = 0;
					polygon[i].x    = rc.left + sizeLegend.cy;
					polygon[i++].y  = rc.top + aiLineWidths[BORDER_TOP][j];
					polygon[i].x    = rc.left + sizeLegend.cy;
					polygon[i++].y  = rc.top;
				}

				// upper right
				polygon[i].x = rcSave.right - ((wEdges&BF_RIGHT)?
					MulDivQuick(aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE],
					aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL]):0);
				polygon[i++].y = rc.top;

				// lower right
				polygon[i].x = rcSave.right - ((wEdges&BF_RIGHT)?
					MulDivQuick(aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE]+aiLineWidths[BORDER_TOP][j],
					aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL]):0);
				polygon[i++].y = rc.top + aiLineWidths[BORDER_TOP][j];

				if(!fdrawCaption)
				{
					polygon[i].x = rcSave.left + MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE]+aiLineWidths[BORDER_TOP][j],
						aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL]);
					polygon[i++].y  = rc.top + aiLineWidths[BORDER_TOP][j];
				}

				Polygon(hdc, polygon, i);
                i = 0;
			}

			fContinuingPoly = FALSE;
			i = 0;
			if(wEdges&BF_RIGHT && (j<aiNumBorderLines[BORDER_RIGHT]) && aiLineWidths[BORDER_RIGHT][j])
			{
				// There's a right edge
				// get color brush for right side
				crNew = pborderinfo->acrColors[BORDER_RIGHT][j];
				if(crNew != crNow)
				{
					HBRUSH hbrNew;
					SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
					// not supported on WINCE
					::UnrealizeObject(hbrNew);
					::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
				}

				// build right side polygon
				polygon[i].x = rc.right - aiLineWidths[BORDER_RIGHT][j]; // upper left corner
				polygon[i++].y =  rcSave.top + MulDivQuick(aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE]+aiLineWidths[BORDER_RIGHT][j],
					aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL]);

				if(pborderinfo->abStyles[BORDER_RIGHT]==pborderinfo->abStyles[BORDER_TOP] &&
					aiLineWidths[BORDER_RIGHT][j]==1)
				{
					// upper right corner fix: we have to overlap one pixel to avoid holes
					polygon[i].x = rc.right - 1;
					polygon[i].y = rcSave.top + MulDivQuick(aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL]);

					polygon[i+1].x = rc.right;
					polygon[i+1].y = polygon[i].y;

					i = i + 2;
				}
				else
				{
					polygon[i].x = rc.right;
					polygon[i++].y = rcSave.top + MulDivQuick(aiLineWidths[BORDER_TOP][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL]);
				}

				if(!(wEdges&BF_BOTTOM) || (pborderinfo->acrColors[BORDER_RIGHT][j]!=pborderinfo->acrColors[BORDER_BOTTOM][j])
					|| !abCanUseJoin[BORDER_BOTTOM])
				{
					polygon[i].x = rc.right;                                     // lower right corner
					polygon[i++].y = rcSave.bottom - MulDivQuick(aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL]);

					polygon[i].x = rc.right - aiLineWidths[BORDER_RIGHT][j];     // lower left corner
					polygon[i++].y = rcSave.bottom - MulDivQuick(aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE]+aiLineWidths[BORDER_RIGHT][j],
						aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL]);
					Polygon(hdc, polygon, i);
					i = 0;
				}
				else
				{
					fContinuingPoly = TRUE;
				}
			}
			if(wEdges&BF_BOTTOM && (j<aiNumBorderLines[BORDER_BOTTOM]) && aiLineWidths[BORDER_BOTTOM][j])
			{
				// There's a bottom edge
				if(!fContinuingPoly)
				{
					// get color brush for bottom side
					crNew = pborderinfo->acrColors[BORDER_BOTTOM][j];
					if(crNew != crNow)
					{
						HBRUSH hbrNew;
						SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
						// not supported on WINCE
						::UnrealizeObject(hbrNew);
						::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
						i = 0;
					}
				}

				// build bottom side polygon
				polygon[i].x = rcSave.right - MulDivQuick(aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL]);
				polygon[i++].y = rc.bottom;

				if(pborderinfo->abStyles[BORDER_BOTTOM]==pborderinfo->abStyles[BORDER_LEFT] &&
					aiLineWidths[BORDER_RIGHT][j]==1)
				{
					// bottom left
					polygon[i].x = rcSave.left + MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL]);
					polygon[i].y = rc.bottom;

					// lower left fix, we have to overlap 1 pixel to avoid holes
					polygon[i+1].x = polygon[i].x;
					polygon[i+1].y = rc.bottom - 1;

					i = i + 2;
				}
				else
				{
					polygon[i].x = rcSave.left + MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
						aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE],
						aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL]);
					polygon[i++].y = rc.bottom;
				}

				// upper left, we have to overlap 1 pixel to avoid holes
				polygon[i].x = rcSave.left + MulDivQuick(aiLineWidths[BORDER_LEFT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE]+aiLineWidths[BORDER_BOTTOM][j],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL]);
				polygon[i++].y = rc.bottom - aiLineWidths[BORDER_BOTTOM][j];
				polygon[i].x = rcSave.right  - MulDivQuick(aiLineWidths[BORDER_RIGHT][DRAWBORDER_TOTAL],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE]+aiLineWidths[BORDER_BOTTOM][j],
					aiLineWidths[BORDER_BOTTOM][DRAWBORDER_TOTAL]);

				polygon[i++].y = rc.bottom - aiLineWidths[BORDER_BOTTOM][j];

				Polygon(hdc, polygon, i);
				i = 0;
			}

		}


		// Shrink rect for this line or spacer
		rc.top    += aiLineWidths[BORDER_TOP][j];
		rc.right  -= aiLineWidths[BORDER_RIGHT][j];
		rc.bottom -= aiLineWidths[BORDER_BOTTOM][j];
		rc.left   += aiLineWidths[BORDER_LEFT][j];

		// increment border shifts
		aiLineWidths[BORDER_RIGHT][DRAWBORDER_INCRE] += aiLineWidths[BORDER_RIGHT][j];
		aiLineWidths[BORDER_TOP][DRAWBORDER_INCRE] += aiLineWidths[BORDER_TOP][j];
		aiLineWidths[BORDER_BOTTOM][DRAWBORDER_INCRE] += aiLineWidths[BORDER_BOTTOM][j];
		aiLineWidths[BORDER_LEFT][DRAWBORDER_INCRE] += aiLineWidths[BORDER_LEFT][j];
	}

	// Okay, now let's draw the flat border if necessary.
	if(xyFlat != 0)
	{
		if(pborderinfo->xyFlat < 0)
		{
			rc = rcOrig;
		}
		rc.right++;
		rc.bottom++;
		xyFlat++;

		if(wEdges & BF_RIGHT)
		{
			crNew = pborderinfo->acrColors[BORDER_RIGHT][1];
			if(crNew != crNow)
			{
				HBRUSH hbrNew;
				SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
				// not supported on WINCE
				::UnrealizeObject(hbrNew);
				::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
			}
			Rectangle(hdc,
				rc.right - xyFlat,
				rc.top + ((wEdges&BF_TOP) ? 0 : xyFlat),
				rc.right,
				rc.bottom - ((wEdges&BF_BOTTOM) ? 0 : xyFlat)
				);
		}
		if(wEdges & BF_BOTTOM)
		{
			crNew = pborderinfo->acrColors[BORDER_BOTTOM][1];
			if(crNew != crNow)
			{
				HBRUSH hbrNew;
				SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
				// not supported on WINCE
				::UnrealizeObject(hbrNew);
				::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
			}
			Rectangle(
				hdc,
				rc.left+((wEdges&BF_LEFT)?0:xyFlat),
				rc.bottom-xyFlat,
				rc.right-((wEdges&BF_RIGHT)?0:xyFlat),
				rc.bottom);
		}

		if(wEdges & BF_TOP)
		{
			crNew = pborderinfo->acrColors[BORDER_TOP][1];
			if(crNew != crNow)
			{
				HBRUSH hbrNew;
				SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
				// not supported on WINCE
				::UnrealizeObject(hbrNew);
				::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
			}
			Rectangle(hdc,
				rc.left+((wEdges&BF_LEFT)?0:xyFlat),
				rc.top,
				rc.right-((wEdges&BF_RIGHT)?0:xyFlat),
				rc.top+xyFlat);
		}
		if(wEdges & BF_LEFT)
		{
			crNew = pborderinfo->acrColors[BORDER_LEFT][1];
			if(crNew != crNow)
			{
				HBRUSH hbrNew;
				SelectCachedBrush(hdc, crNew, &hbrNew, &hbrOld, &crNow);
				// not supported on WINCE
				::UnrealizeObject(hbrNew);
				::SetBrushOrgEx(hdc, POSITIVE_MOD(ptBrushOrg.x, 8), POSITIVE_MOD(ptBrushOrg.y, 8), NULL);
			}
			Rectangle(
				hdc,
				rc.left,
				rc.top+((wEdges&BF_TOP)?0:xyFlat),
				rc.left+xyFlat,
				rc.bottom-((wEdges&BF_BOTTOM)?0:xyFlat));
		}
	}

	if(hbrOld)
	{
		ReleaseCachedBrush((HBRUSH)SelectObject(hdc, hbrOld));
	}
}