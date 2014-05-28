
#ifndef __XINDOWS_SITE_DISPLAY_DISPCONTENT_H__
#define __XINDOWS_SITE_DISPLAY_DISPCONTENT_H__

//+---------------------------------------------------------------------------
//
//  Class:      CDispContentNode
//
//  Contents:   An abstract interior node which represents content
//              (vs. a balancing node)
//
//----------------------------------------------------------------------------
class CDispContentNode : public CDispInteriorNode
{
	DECLARE_DISPNODE_ABSTRACT(CDispContentNode, CDispInteriorNode)

	friend class CDispNode;

protected:
	CDispClient* _pDispClient;
	// 52 bytes (4 bytes + 48 bytes for CDispInteriorNode super class)

	// object can be created only by derived classes, and destructed only from
	// special methods
protected:
	CDispContentNode(CDispClient* pDispClient) : CDispInteriorNode()
	{
		Assert(pDispClient != NULL);
		_pDispClient = pDispClient;
	}
	virtual ~CDispContentNode() {}

public:
	// CDispNode overrides
	virtual LONG GetZOrder() const;
	virtual CDispClient* GetDispClient() const { return _pDispClient; }
};

#endif //__XINDOWS_SITE_DISPLAY_DISPCONTENT_H__
