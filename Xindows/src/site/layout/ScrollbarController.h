
#ifndef __XINDOWS_SITE_LAYOUT_SCROLLBARCONTROLLER_H__
#define __XINDOWS_SITE_LAYOUT_SCROLLBARCONTROLLER_H__

#include "Scrollbar.h"
class CDispScroller;

//+---------------------------------------------------------------------------
//
//  Class:      CScrollbarController
//
//  Synopsis:   Transient object to control scroll bar during user interaction.
//
//----------------------------------------------------------------------------
class CScrollbarController : public CScrollbar, public CVoid
{
public:
	static void StartScrollbarController(
		CLayout*		pLayout,
		CDispScroller*	pDispScroller,
		CServer*		pServerHost,
		long			buttonWidth,
		CMessage*		pMessage);

	static void StopScrollbarController();

private:
	friend HRESULT InitScrollbar(THREADSTATE* pts);
	friend void DeinitScrollbar(THREADSTATE* pts);

	// constructor is private,
	// use InitScrollbar or StartScrollbarController
	CScrollbarController() { _pLayout = NULL; }
	// destructor is private, because this object
	// is deleted only by DeinitScrollbar
	~CScrollbarController()
	{
		AssertSz(_pLayout==NULL, "CScrollbarController destructed prematurely");
	}

public:
	CLayout* GetLayout() const { return _pLayout; }
	CDrawInfo* GetDrawInfo() { return &_drawInfo; }

	SCROLLBARPART GetPartPressed() const { return _partPressed; }

	static long GetRepeatDelay() { return 250; }
	static long GetRepeatRate() { return 50; }
	static long GetFocusRate() { return 500; }

private:
	NV_DECLARE_ONTICK_METHOD(OnTick, ontick, (UINT idTimer));  // For global window
	NV_DECLARE_ONMESSAGE_METHOD(OnMessage, onmessage, (UINT msg, WPARAM wParam, LPARAM lParam));

	void MouseMove(const CPoint& pt);
	void GetScrollInfo(
		LONG* pContentSize,
		LONG* pContainerSize,
		LONG* pScrollAmount) const;

	SCROLLBARPART		_partPressedStart;
	SCROLLBARPART		_partPressed;
	CLayout*			_pLayout;
	CDispScroller*		_pDispScroller;
	CServer*			_pServerHost;
	int					_direction;
	long				_buttonWidth;
	CRect				_rcScrollbar;
	CPoint				_ptMouse;
	long				_mouseInThumb;
	CDrawInfo			_drawInfo;
};

#endif //__XINDOWS_SITE_LAYOUT_SCROLLBARCONTROLLER_H__