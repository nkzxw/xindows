
#ifndef __XINDOWS_SITE_BASE_ROOTELEMENT_H__
#define __XINDOWS_SITE_BASE_ROOTELEMENT_H__

class CRootElement : public CTxtSite
{
    typedef CTxtSite super;

public:
    CRootElement(CDocument* pDoc) : super(ETAG_ROOT, pDoc) {}

    void Notify(CNotification* pNF);

    virtual HRESULT ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget);

    virtual HRESULT YieldCurrency(CElement* pElemNew);
    virtual void YieldUI(CElement* pElemNew);
    virtual HRESULT BecomeUIActive();

    DECLARE_MEMCLEAR_NEW_DELETE();

    DECLARE_CLASSDESC_MEMBERS;
};

#endif //__XINDOWS_SITE_BASE_ROOTELEMENT_H__