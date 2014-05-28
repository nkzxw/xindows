
#ifndef __XINDOWS_SITE_DISPLAY_DISPFILTER_H__
#define __XINDOWS_SITE_DISPLAY_DISPFILTER_H__

class CDispNode;
class CDispContext;
class CDispDrawContext;

//+---------------------------------------------------------------------------
//
//  Class:      CDispFilter
//
//  Synopsis:   Abstract base class for display tree filter objects which
//              perform filtering for arbitrary display nodes.
//
//----------------------------------------------------------------------------
class CDispFilter
{
public:
	virtual void SetSize(const SIZE& size) = 0;

	virtual void DrawFiltered(CDispDrawContext* pContext) = 0;

	virtual void Invalidate(const RECT& rc, BOOL fSynchronousRedraw) = 0;

	virtual void Invalidate(HRGN hrgn, BOOL fSynchronousRedraw) = 0;

	virtual void SetOpaque(BOOL fOpaque) = 0;
};

#endif //__XINDOWS_SITE_DISPLAY_DISPFILTER_H__
