
#ifndef __XINDOWS_SITE_DISPLAY_DISPCONTAINERPLUS_H__
#define __XINDOWS_SITE_DISPLAY_DISPCONTAINERPLUS_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispContainerPlus
//
//  Synopsis:   A container node with optional border and user clip.
//
//----------------------------------------------------------------------------
class CDispContainerPlus : public CDispContainer
{
	DECLARE_DISPNODE(CDispContainerPlus, CDispContainer, DISPNODETYPE_CONTAINERPLUS)

protected:
	LONG _extra[]; // variable length, must be last
	// 72-112 bytes (4-44 bytes + 68 bytes for CDispContainer)

	// object can be created only by CDispRoot, and destructed only from
	// special methods
protected:
	void* operator new(size_t cb) { Assert(FALSE); return NULL; } // do not use
	void* operator new(size_t cb, size_t cbExtra)
	{
		return MemAllocClear(cb+cbExtra);
	}
	void operator delete(void* pv) { MemFree(pv); }

	CDispContainerPlus(CDispClient* pDispClient) : CDispContainer(pDispClient) {}
	CDispContainerPlus(const CDispItemPlus* pItemPlus);
	~CDispContainerPlus() {}

public:
	// CDispNode overrides
	virtual void SetSize(const SIZE& size, BOOL fInvalidateAll);

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

	// CDispInterior overrides
	virtual BOOL ComputeVisibleBounds();
};

#endif //__XINDOWS_SITE_DISPLAY_DISPCONTAINERPLUS_H__
