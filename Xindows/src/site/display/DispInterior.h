
#ifndef __XINDOWS_SITE_DISPLAY_DISPINTERIOR_H__
#define __XINDOWS_SITE_DISPLAY_DISPINTERIOR_H__

// constants used to determine when tree balancing is necessary
#define MAX_CHILDREN_BEFORE_BALANCE     100 // don't balance nodes with less
                                            // than this many children
#define MAX_CHILDREN_AFTER_BALANCE      50  // produce nodes with this many
                                            // children after balancing
#define MIN_CHILDREN_BEFORE_BALANCE     5   // only group more than this
                                            // many children during balancing

//+---------------------------------------------------------------------------
//
//  Class:      CDispInteriorNode
//
//  Synopsis:   Base class for interior (non-leaf) display nodes.
//
//----------------------------------------------------------------------------
class CDispInteriorNode : public CDispNode
{
	DECLARE_DISPNODE_ABSTRACT(CDispInteriorNode, CDispNode)

	friend class CDispRoot;
	friend class CDispNode;
	friend class CDispLeafNode;
	friend class CDispItemPlus;
	friend class CDispContainer;
	friend class CDispContainerPlus;
	friend class CDispScroller;
	friend class CDispScrollerPlus;
	friend class CDispDrawContext;
	friend class CDispContextStack;

protected:
	CDispNode*	_pFirstChildNode;	// pointer to first child
	CDispNode*	_pLastChildNode;	// pointer to last child
	LONG		_cChildren;			// children count
	// 48 bytes (12 bytes + 36 bytes for CDispNode base class)

protected:
	// object can be created only by derived classes, and destructed only from
	// special methods
	CDispInteriorNode() : CDispNode()
	{
		SetFlag(CDispFlags::s_interiorNode);
	}
	virtual ~CDispInteriorNode();

public:
	CDispNode* GetFirstChildNode() const;
	CDispNode* GetFirstChildNodeInLayer(DISPNODELAYER layer) const;

	CDispNode* GetLastChildNode() const;
	CDispNode* GetLastChildNodeInLayer(DISPNODELAYER layer) const;

	CDispNode* GetChildInFlow(LONG index) const;

	LONG CountChildren() const;

	void InsertChildInFlow(CDispNode* pNewChild);
	void InsertFirstChildInFlow(CDispNode* pNewChild);
	void InsertChildInZLayer(CDispNode* pNewChild, LONG zOrder)
	{
		if(zOrder < 0)
		{
			InsertChildInNegZ(pNewChild, zOrder);
		}
		else
		{
			InsertChildInPosZ(pNewChild, zOrder);
		}
	}

	void TakeChildrenFrom(
		CDispInteriorNode* pOldParent,
		CDispNode* pFirstChildToTake=NULL,
		CDispNode* pChildToInsertAfter=NULL);
	virtual CDispScroller* HitScrollInset(CPoint* pptHit, DWORD* pdwScrollDir);

protected:
	void InsertFirstChildNode(CDispNode* pNewChild);
	void InsertLastChildNode(CDispNode* pNewChild);
	void InsertChildInNegZ(CDispNode* pNewChild, LONG zOrder);
	void InsertChildInPosZ(CDispNode* pNewChild, LONG zOrder);

	void DestroyChildren();
	void BalanceTree();
	void CreateBalanceGroups(
		CDispNode* pFirstChild,
		CDispNode* pStopChild,
		LONG cChildrenInEachGroup);

	// Determine whether a node and its children are in-view or not.
	BOOL CalculateInView();

	void SetSubtreeFlags(const CDispFlags& flags);
	void ClearSubtreeFlags(const CDispFlags& flags);

	// CDispNode overrides
	virtual BOOL PreDraw(CDispDrawContext* pContext);
	virtual void DrawSelf(CDispDrawContext* pContext, CDispNode* pChild);
	virtual BOOL HitTestPoint(CDispHitContext* pContext) const;
	virtual void CalcDispInfo(const CRect& rcClip, CDispInfo* pdi) const;

	// CDispInteriorNode virtuals
	virtual void PushContext(
		const CDispNode* pChild,
		CDispContextStack* pContextStack,
		CDispContext* pContext) const;
	virtual BOOL ComputeVisibleBounds();
	virtual BOOL SubtractOpaqueChildren(CRegion* prgn, CDispContext* pContext);
	virtual BOOL CalculateInView(
		CDispContext* pContext,
		BOOL fPositionChanged,
		BOOL fNoRedraw);
	virtual void RecalcChildren(
		BOOL fForceRecalc,
		BOOL fSuppressInval,
		CDispDrawContext* pContext);

	void GetScrollableContentSize(CSize* psize, BOOL fAddBorder=FALSE) const;

	BOOL PreDrawChild(
		CDispNode* pChild,
		CDispDrawContext* pContext,
		const CDispContext& saveContext) const;

	void InsertChildNode(
		CDispNode* pChild,
		CDispNode* pPreviousSibling,
		CDispNode* pNextSibling);

	void SetPositionHasChanged();

#ifdef _DEBUG
public:
    virtual void VerifyTreeCorrectness() const;

protected:
    virtual size_t GetMemorySize() const;
    void VerifyChildrenCount() const;
    void VerifyFlags(const CDispFlags& mask, const CDispFlags& value, BOOL fEqual) const;
#endif
};

#endif //__XINDOWS_SITE_DISPLAY_DISPINTERIOR_H__