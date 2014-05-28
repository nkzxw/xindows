
#ifndef __XINDOWS_SITE_DISPLAY_DISPLEAF_H__
#define __XINDOWS_SITE_DISPLAY_DISPLEAF_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispLeafNode
//
//  Synopsis:   Abstract base class for all leaf nodes in the display tree.
//
//----------------------------------------------------------------------------
class CDispLeafNode : public CDispNode
{
	DECLARE_DISPNODE_ABSTRACT(CDispLeafNode, CDispNode)

	friend class CDispNode;
	friend class CDispInteriorNode;

protected:
	// no additional data
	// 36 bytes (0 bytes + 36 bytes for CDispNode base class)

protected:
	// object can be created only by derived classes, and destructed only from
	// special methods
	CDispLeafNode() : CDispNode() {}
	virtual ~CDispLeafNode() {}

public:
	// size and position
	virtual void GetSize(SIZE* psize) const { _rcVisBounds.GetSize(psize); }
	virtual void SetSize(const SIZE& size, BOOL fInvalidateAll)
	{
		SetSize(size, _afxGlobalData._Zero.rc, fInvalidateAll);
	}
	virtual const CPoint& GetPosition() const { return _rcVisBounds.TopLeft(); }
	virtual void GetPositionTopRight(CPoint* pptTopRight)
	{
		_rcVisBounds.GetTopRight(pptTopRight);
	}
	virtual void SetPosition(const POINT& ptTopLeft);
	virtual void SetPositionTopRight(const POINT& ptTopRight);
	virtual void FlipBounds() { _rcVisBounds.MirrorX(); }

	virtual void* GetCookie() const { return NULL; }

protected:
	// CDispLeafNode virtuals
	virtual void NotifyInViewChange(
		CDispContext* pContext,
		BOOL fResolvedVisible,
		BOOL fWasResolvedVisible,
		BOOL fNoRedraw) = 0;

	void SetSize(
		const SIZE& size,
		const RECT& rcBorderWidths,
		BOOL fInvalidateAll);
};


#endif //__XINDOWS_SITE_DISPLAY_DISPLEAF_H__
