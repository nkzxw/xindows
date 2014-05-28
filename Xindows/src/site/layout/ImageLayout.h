
#ifndef __XINDOWS_SITE_LAYOUT_IMAGELAYOUT_H__
#define __XINDOWS_SITE_LAYOUT_IMAGELAYOUT_H__

#include "../builtin/ImgHelper.h"

struct IMGANIMSTATE;

class CImageLayout : public CLayout
{
public:
    typedef CLayout super;

    CImageLayout(CElement* pElementLayout) : super(pElementLayout)
    {
        _fNoUIActivateInDesign = TRUE;
    }
    DECLARE_MEMCLEAR_NEW_DELETE()

    // Sizing and Positioning
    void DrawFrame(IMGANIMSTATE* pImgAnimState) {}

    CImgHelper* GetImgHelper();
    BOOL IsOpaque() { return GetImgHelper()->IsOpaque(); }

    void HandleViewChange(
        DWORD       flags,
        const RECT* prcClient,
        const RECT* prcClip,
        CDispNode*  pDispNode);

protected:
    DECLARE_LAYOUTDESC_MEMBERS
};

class CImgElementLayout : public CImageLayout
{
public:
    typedef CImageLayout super;
    CImgElementLayout(CElement* pElementLayout) : super(pElementLayout)
    {
    }
    DECLARE_MEMCLEAR_NEW_DELETE()

    virtual void Draw(CFormDrawInfo* pDI, CDispNode*);

    // Drag & drop
    virtual HRESULT PreDrag(
        DWORD dwKeyState, IDataObject** ppDO, IDropSource** ppDS);
    virtual HRESULT PostDrag(HRESULT hr, DWORD dwEffect);
    virtual void GetMarginInfo(
        CParentInfo* ppri,
        LONG* plLeftMargin,
        LONG* plTopMargin,
        LONG* plRightMargin,
        LONG* plBottomMargin);
    virtual DWORD CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault);
    void InvalidateFrame();
};

#endif //__XINDOWS_SITE_LAYOUT_IMAGELAYOUT_H__