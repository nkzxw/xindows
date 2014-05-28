
#ifndef __XINDOWS_SITE_BASE_SITE_H__
#define __XINDOWS_SITE_BASE_SITE_H__

#define _hxx_
#include "../gen/csite.hdl"

//+---------------------------------------------------------------------------
//
//  Class:      CSite (site)
//
//  Purpose:    Base class for all site objects used by CDoc.
//
//----------------------------------------------------------------------------
class CSite : public CElement
{
    DECLARE_CLASS_TYPES(CSite, CElement)

private:
    DECLARE_MEMCLEAR_NEW_DELETE()

public:
    // Construct / Destruct
    CSite(ELEMENT_TAG etag, CDocument* pDoc);

    virtual HRESULT Init();

    // IUnknown methods
    DECLARE_PRIVATE_QI_FUNCS(CBase)

#define _CSite_
#include "../gen/csite.hdl"

    // only derived classes should call this
    virtual HRESULT CreateLayout() { Assert(FALSE); RRETURN(E_FAIL); }

    // Notify site that it is being detached from the form
    //   Children should be detached and released
    //-------------------------------------------------------------------------
    //  +override : add behavior
    //  +call super : last
    //  -call parent : no
    //  +call children : first
    //-------------------------------------------------------------------------
    virtual float Detach() { return 0.0; }
};

int CompareElementsByZIndex(const void* pv1, const void* pv2);

#define MAX_BORDER_SPACE  1000

DWORD GetBorderInfoHelper(CTreeNode* pNodeContext, CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll=FALSE);
void GetBorderColorInfoHelper(CTreeNode* pNodeContext, CDocInfo* pdci, CBorderInfo* pborderinfo);

void CalcBgImgRect(
                   CTreeNode*       pNode,
                   CFormDrawInfo*   pDI,
                   const SIZE*      psizeObj,
                   const SIZE*      psizeImg,
                   CPoint*          pptBackOrg,
                   RECT*            prcBackClip);

#endif //__XINDOWS_SITE_BASE_SITE_H__