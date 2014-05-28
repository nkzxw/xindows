
#ifndef __XINDOWS_SITE_DISPLAY_DISPINFO_H__
#define __XINDOWS_SITE_DISPLAY_DISPINFO_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispInfo
//
//  Synopsis:   Structure for returning display node information regarding
//              coordinate system offsets, clipping, and scrolling.
//
//----------------------------------------------------------------------------
class CDispInfo
{
	friend class CDispExtras;

public:
	CDispInfo(const CDispExtras* pExtras)
	{
		if(pExtras == NULL)
		{
			_pInsetOffset = (const CSize*)&_afxGlobalData._Zero.size;
			_prcBorderWidths = (const CRect*)&_afxGlobalData._Zero.rc;
		}
		else
		{
			pExtras->GetExtraInfo(this);
		}
	}
	~CDispInfo() {}

	// optional information filled in by CDispFlags::GetExtraInfo
	const CSize*    _pInsetOffset;	    // never NULL
	const CRect*    _prcBorderWidths;   // never NULL

	// information filled in by CalcDispInfo virtual method
	CRect           _rcPositionedClip;  // clip positioned content (container coords.)
	CRect           _rcContainerClip;   // clip border and scroll bars (container coords.)
	CRect           _rcBackgroundClip;  // clip background content (content coords.)
	CRect           _rcFlowClip;        // clip flow content (flow coords.)
	CSize           _borderOffset;      // offset to border
	CSize           _contentOffset;     // offset to content inside border
	CSize           _scrollOffset;      // scroll offset
	CSize           _sizeContent;       // size of content
	CSize           _sizeBackground;    // size of background (exclude border and scrollbars)

private:
	// temporary rect used to store simple border widths
	CRect           _rcTemp;
};

#endif //__XINDOWS_SITE_DISPLAY_DISPINFO_H__
