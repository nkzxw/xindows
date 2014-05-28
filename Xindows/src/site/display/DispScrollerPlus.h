
#ifndef __XINDOWS_SITE_DISPLAY_DISPSCROLLERPLUS_H__
#define __XINDOWS_SITE_DISPLAY_DISPSCROLLERPLUS_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispScrollerPlus
//
//  Synopsis:   Scrolling container with optional border and user clip.
//
//----------------------------------------------------------------------------
class CDispScrollerPlus : public CDispScroller
{
	DECLARE_DISPNODE(CDispScrollerPlus, CDispScroller, DISPNODETYPE_SCROLLERPLUS)

protected:
	LONG _extra[]; // variable length, must be last
	// 100-140 bytes (4-44 bytes + 96 bytes for CDispScroller)

	// object can be created only by CDispRoot, and destructed only from
	// special methods
protected:
	void* operator new(size_t cb) { Assert(FALSE); return NULL; } // do not use
	void* operator new(size_t cb, size_t cbExtra)
	{
		return MemAllocClear(cb+cbExtra);
	}
	void operator delete(void* pv) { MemFree(pv); }

	CDispScrollerPlus(CDispClient* pDispClient) : CDispScroller(pDispClient) {}
	~CDispScrollerPlus() {}

protected:
	// CDispNode overrides
	virtual void GetNodeTransform(
		CSize* pOffset,
		COORDINATE_SYSTEM source,
		COORDINATE_SYSTEM destination) const;
	virtual void GetNodeTransform(
		CDispContext* pContext,
		COORDINATE_SYSTEM source,
		COORDINATE_SYSTEM destination) const;
	virtual void CalcDispInfo(
		const CRect& rcClip,
		CDispInfo* pdi) const;
	virtual CDispExtras* GetExtras() const { return (CDispExtras*)_extra; }

private:
	void GetInsideBorderRect(CRect* prc, const CDispInfo& di) const;
};

#endif //__XINDOWS_SITE_DISPLAY_DISPSCROLLERPLUS_H__
