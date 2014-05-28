
#ifndef __XINDOWS_SITE_LAYOUT_BODYLAYOUT_H__
#define __XINDOWS_SITE_LAYOUT_BODYLAYOUT_H__

//+---------------------------------------------------------------------------
//
// CBodyLayout
//
//----------------------------------------------------------------------------
class CBodyLayout : public CFlowLayout
{
public:
    typedef CFlowLayout super;

    CBodyLayout(CElement* pElement) : CFlowLayout(pElement)
    {
        _fFocusRect = FALSE;
        _fContentsAffectSize = FALSE;
    }

    virtual HRESULT STDMETHODCALLTYPE HandleMessage(CMessage* pMessage);
    virtual DWORD   CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault = NULL);
    virtual void    NotifyScrollEvent(RECT* prcScroll, SIZE* psizeScrollDelta);

    void            RequestFocusRect(BOOL fOn);
    void            RedrawFocusRect();
    HRESULT         GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape);

    BOOL            GetBackgroundInfo(CFormDrawInfo* pDI, BACKGROUNDINFO* pbginfo, BOOL fAll=TRUE, BOOL fRightToLeft=FALSE);

    virtual HRESULT OnSelectionChange(void);

    void            UpdateScrollInfo(CDispNodeInfo* pdni) const;

    CBodyElement* Body() { return (CBodyElement*)ElementOwner(); }


private:
    unsigned _fFocusRect:1; // Draw a focus rect

protected:
    DECLARE_LAYOUTDESC_MEMBERS
};

#endif //__XINDOWS_SITE_LAYOUT_BODYLAYOUT_H__