
#ifndef __XINDOWS_SITE_LAYOUT_TXTLAYOUT_H__
#define __XINDOWS_SITE_LAYOUT_TXTLAYOUT_H__

class CRichtextLayout : public CFlowLayout
{
public:
    typedef CFlowLayout super;

    CRichtextLayout(CElement* pElementLayout) : super(pElementLayout)
    {
        _fCanHaveChildren = TRUE;
    }

    DECLARE_MEMCLEAR_NEW_DELETE()

    virtual HRESULT Init();
    HRESULT GetFontSize(CCalcInfo* pci, SIZE* psizeFontShort, SIZE* psizeFontLong);

    // Sizing and Positioning
    virtual DWORD CalcSizeHelper(CCalcInfo* pci, int charX, int charY, SIZE* psize);
    virtual DWORD CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault=NULL);
    virtual BOOL GetMultiLine() const { return TRUE; }
    virtual BOOL GetAutoSize()  const { return FALSE; }
    void SetWrap();
    BOOL IsWrapSet();

    virtual HRESULT OnTextChange(void);
    virtual HRESULT OnSelectionChange(void);

    virtual void DrawClient(
        const RECT*     prcBounds,
        const RECT*     prcRedraw,
        CDispSurface*   pDispSurface,
        CDispNode*      pDispNode,
        void*           cookie,
        void*           pClientData,
        DWORD           dwFlags);

protected:
    DECLARE_LAYOUTDESC_MEMBERS;
};

class CTextAreaLayout : public CRichtextLayout
{
public:
    typedef CRichtextLayout super;

    CTextAreaLayout(CElement* pElementLayout) : super(pElementLayout)
    {
        _fCanHaveChildren = FALSE;
    }

    DECLARE_MEMCLEAR_NEW_DELETE()
};

#endif //__XINDOWS_SITE_LAYOUT_TXTLAYOUT_H__