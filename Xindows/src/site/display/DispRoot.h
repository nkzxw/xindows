
#ifndef __XINDOWS_SITE_DISPLAY_DISPROOT_H__
#define __XINDOWS_SITE_DISPLAY_DISPROOT_H__

#include "../include/IDispObserver.h"

class CDispItemPlus;
class CDispGroup;

//+---------------------------------------------------------------------------
//
//  Class:      CDispRoot
//
//  Synopsis:   A container node at the root of a display tree.
//
//----------------------------------------------------------------------------
class CDispRoot : public CDispContainer
{
	DECLARE_DISPNODE(CDispRoot, CDispContainer, DISPNODETYPE_ROOT)

	friend class CDispNode;
	friend class CDispScroller;

protected:
	IDispObserver*		_pDispObserver;
	int					_cOpen;
	BOOL				_fDrawLock;
	BOOL				_fRecalcLock;
	CDispDrawContext	_drawContext;
	CDispSurface*		_pRenderSurface;
	CDispSurface*		_pOffscreenBuffer;
	CRect				_rcContent;
	BOOL				_fCanScrollDC;
	BOOL				_fCanSmoothScroll;
	BOOL				_fDirectDraw;
	BOOL				_fTexture;
	BOOL				_fAllowOffscreen;
	BOOL				_fWantOffscreen;
	short				_cBufferDepth;
	HPALETTE			_hpal;

	void CheckReenter() const
	{
		AssertSz(!_fDrawLock, "Open/CloseDisplayTree prohibited during Draw()");
	}
	void ReleaseRenderSurface()
	{
		delete _pRenderSurface;
		_pRenderSurface = NULL;
	}
	void ReleaseOffscreenBuffer();

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

	CDispRoot(IDispObserver* pObserver, CDispClient* pDispClient);

	// call Destroy instead of delete
protected:
	~CDispRoot();
public:
	// How we paint and where.  This covers the final destination, buffering and so on.
	void SetDestination(HDC hdc, IDirectDrawSurface* pSurface);

	// lines = -1 for no buffering, 0 for full size, > 0 for number of lines
	// (in pixels) to buffer
	void SetOffscreenBufferInfo(HPALETTE hpal, short bufferDepth, BOOL fDirectDraw, BOOL fTexture, BOOL fWantOffscreen, BOOL fAllowOffscreen);
	BOOL SetupOffscreenBuffer(CDispDrawContext* pContext);

	void SetRootSize(const SIZE& size, BOOL fInvalidateAll);
	const POINT& GetRootPosition() const { return _rcContainer.TopLeft(); }
	void GetRootPosition(POINT* ppt) const { *ppt = _rcContainer.TopLeft(); }
	void SetRootPosition(const POINT& pt)
	{
		_rcContainer.MoveTo(pt);
		_rcVisBounds.MoveTo(pt);
	}

	BOOL SetContentOffset(const SIZE& size);

	// return a pointer to the root node's display context
	// NOTE: returns a non-AddRef'd pointer
	CDispDrawContext* GetDrawContext() { return &_drawContext; }

	// DrawRoot contains special optimizations for dealing with scrollbars of
	// its first child
	void DrawRoot(
        CDispDrawContext* pContext,
        void* pClientData,
        HRGN hrgnDraw=NULL,
        const RECT* prcDraw=NULL);

	// DrawNode draws the given node on the given surface without any special
	// optimizations
	void DrawNode(
		CDispNode* pNode,
		CDispSurface* pSurface,
		const POINT& ptOrg,
		HRGN rgnClip,
		void* pClientData);

	// hit testing
	BOOL HitTest(
		CPoint* pptHit,
		COORDINATE_SYSTEM coordinateSystem,
		void* pClientData,
		BOOL fHitContent,
		long cFuzzyHitTest=0);

	// quickly draw border and background
	void EraseBackground(
		CDispDrawContext* pContext,
		void* pClientData,
		HRGN hrgnDraw=NULL,
		const RECT* prcDraw=NULL);

#ifdef _DEBUG
    void OpenDisplayTree();
    void CloseDisplayTree();
#else
	void OpenDisplayTree() { _cOpen++; }
	void CloseDisplayTree()
	{
		if(_cOpen == 1)
		{
			RecalcRoot();
		}
		_cOpen--;
	}
#endif

	BOOL DisplayTreeIsOpen() const { return (_cOpen>0); }

	void RecalcRoot();

	void SetObserver(IDispObserver* pObserver) { _pDispObserver = pObserver; }
	IDispObserver* GetObserver() { return _pDispObserver; }

	void SetCanSmoothScroll(BOOL fSmooth=TRUE) { _fCanSmoothScroll = fSmooth; }
	BOOL CanSmoothScroll() const { return _fCanSmoothScroll; }

	void SetCanScrollDC(BOOL fCanScrollDC=TRUE) { _fCanScrollDC = fCanScrollDC; }
	BOOL CanScrollDC() const { return _fCanScrollDC; }

	// static creation methods
	static CDispItemPlus* CreateDispItemPlus(
		CDispClient* pDispClient,
		BOOL fHasExtraCookie,
		BOOL fHasUserClip,
		BOOL fHasInset,
		DISPNODEBORDER borderType,
		BOOL fRightToLeft);
	static CDispContainer* CreateDispContainer(
		CDispClient* pDispClient,
		BOOL fHasExtraCookie,
		BOOL fHasUserClip,
		BOOL fHasInset,
		DISPNODEBORDER borderType,
		BOOL fRightToLeft);
	static CDispContainer* CreateDispContainer(
		const CDispItemPlus* pItemPlus);
	static CDispScroller* CreateDispScroller(
		CDispClient* pDispClient,
		BOOL fHasExtraCookie,
		BOOL fHasUserClip,
		BOOL fHasInset,
		DISPNODEBORDER borderType,
		BOOL fRightToLeft);

	void InvalidateRoot(const CRegion& rgn,
		BOOL fSynchronousRedraw,
		BOOL fInvalChildWindows);
	void InvalidateRoot(const CRect& rc,
		BOOL fSynchronousRedraw,
		BOOL fInvalChildWindows);

protected:
	// CDispNode overrides
	virtual BOOL PreDraw(CDispDrawContext* pContext);
	virtual void DrawSelf(
		CDispDrawContext* pContext,
		CDispNode* pChild);
	virtual BOOL ScrollRectIntoView(
		const CRect& rc,
		CRect::SCROLLPIN spVert,
		CRect::SCROLLPIN spHorz,
		COORDINATE_SYSTEM coordinateSystem,
		BOOL fScrollBits);

	// CDispInteriorNode overrides
	virtual void PushContext(
		const CDispNode* pChild,
		CDispContextStack* pContextStack,
		CDispContext* pContext) const;

	virtual void RecalcChildren(
		BOOL fForceRecalc,
		BOOL fSuppressInval,
		CDispDrawContext* pContext);

	BOOL ScrollRect(
		const CRect& rcScroll,
		const CSize& scrollDelta,
		CDispScroller* pScroller,
		const CRegion& rgnInvalid,
		BOOL fMayScrollDC);

private:
	void DrawEntire(CDispDrawContext* pContext);
	void DrawBands(
		CDispDrawContext* pContext,
		CRegion* prgnRedraw,
		const CRegionStack& redrawRegionStack);
	void DrawBand(CDispDrawContext* pContext);
};

#endif //__XINDOWS_SITE_DISPLAY_DISPROOT_H__
