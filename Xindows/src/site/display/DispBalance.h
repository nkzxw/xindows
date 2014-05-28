
#ifndef __XINDOWS_SITE_DISPLAY_DISPBALANCE_H__
#define __XINDOWS_SITE_DISPLAY_DISPBALANCE_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispBalanceNode
//
//  Synopsis:   Synthetic node used to group similar nodes for tree efficiency.
//
//----------------------------------------------------------------------------
class CDispBalanceNode : public CDispInteriorNode
{
	DECLARE_DISPNODE(CDispBalanceNode, CDispInteriorNode, DISPNODETYPE_BALANCE)

	// object can be created only by CDispInteriorNode, and destructed only from
	// special methods
protected:
	friend class CDispInteriorNode;

    DECLARE_MEMCLEAR_NEW_DELETE()

	CDispBalanceNode() : CDispInteriorNode()
	{
		SetFlag(CDispFlags::s_balanceGroupInit);
	}
	~CDispBalanceNode() {}

public:
	// CDispNode overrides
	virtual LONG GetZOrder() const
	{
		Assert(FALSE); // shouldn't be here!
		return 0;
	}
};

#endif //__XINDOWS_SITE_DISPLAY_DISPBALANCE_H__
