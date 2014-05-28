
#include "stdafx.h"
#include "ImageLayout.h"

IMPLEMENT_LAYOUT_FNS(CImgElement, CImgElementLayout)

extern BOOL GetCachedImageSize(LPCTSTR pchURL, SIZE* psize);

extern HRESULT CreateImgDataObject(
           CDocument*   pDoc,
           CImgCtx*     pImgCtx,
           CBitsCtx*    pBitsCtx,
           CElement*    pElement,
           CGenDataObject** ppImgDO);


class CImgDragDropSrcInfo : public CDragDropSrcInfo
{
public:
    CImgDragDropSrcInfo() : CDragDropSrcInfo()
    {
        _srcType = DRAGDROPSRCTYPE_IMAGE;
    }

    CGenDataObject* _pImgDO; // Data object for the image being dragged
};

const CLayout::LAYOUTDESC CImageLayout::s_layoutdesc =
{
    0, // _dwFlags
};

//+-------------------------------------------------------------------------
//
//  Method:     CImageLayout::CalcSize
//
//  Synopsis   : This function adjusts the size of the image
//               it uses the image Width and Height if present
//
//--------------------------------------------------------------------------
DWORD CImgElementLayout::CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault)
{
    Assert(ElementOwner());
    CImgElement*    pImg = DYNCAST(CImgElement, ElementOwner());
    CScopeFlag      csfCalcing(this);
    CElement::CLock LockS(ElementOwner(), CElement::ELEMENTLOCK_SIZING);
    CSaveCalcInfo   sci(pci, this);
    CSize           sizeOriginal;
    DWORD           grfReturn;
    HRESULT         hr;

    Assert(pci);
    Assert(psize);

    GetSize(&sizeOriginal);

    if(_fForceLayout)
    {
        pci->_grfLayout |= LAYOUT_FORCE;
        _fForceLayout = FALSE;
    }

    grfReturn  = (pci->_grfLayout & LAYOUT_FORCE);
    _fSizeThis = (_fSizeThis || (pci->_grfLayout&LAYOUT_FORCE));

    // If sizing is needed or it is a min/max request, handle it here
    if((pci->_smMode!=SIZEMODE_SET && _fSizeThis)
        || pci->_smMode==SIZEMODE_MMWIDTH
        || pci->_smMode==SIZEMODE_MINWIDTH)
    {
        Assert(pImg->_pImage);
        pImg->_pImage->CalcSize(pci, psize);

        if(pci->_smMode==SIZEMODE_NATURAL || pci->_smMode==SIZEMODE_SET
            || pci->_smMode==SIZEMODE_FULLSIZE)
        {
            // If dirty, ensure display tree nodes exist
            if(_fSizeThis && (SUCCEEDED(hr=EnsureDispNode(pci, (grfReturn&LAYOUT_FORCE)))))
            {
                if(pImg->_pImage->_pBitsCtx)
                {
                    SetPositionAware();
                }

                if(hr == S_FALSE)
                {
                    grfReturn |= LAYOUT_HRESIZE | LAYOUT_VRESIZE;
                }
            }
            _fSizeThis    = FALSE;
            grfReturn    |= LAYOUT_THIS |
                (psize->cx!=sizeOriginal.cx?LAYOUT_HRESIZE:0) |
                (psize->cy != sizeOriginal.cy?LAYOUT_VRESIZE:0);

            // Size display nodes if size changes occurred
            if(grfReturn & (LAYOUT_FORCE|LAYOUT_HRESIZE|LAYOUT_VRESIZE))
            {
                SizeDispNode(pci, *psize, (grfReturn&LAYOUT_FORCE));
                pImg->_pImage->SetActivity();
            }
        }
        else if(pci->_smMode == SIZEMODE_MMWIDTH)
        {
            psize->cy = psize->cx;
        }
    }
    // Otherwise, defer to default handling
    else
    {
        grfReturn = super::CalcSize(pci, psize);
    }

    return grfReturn;
}

void CImgElementLayout::InvalidateFrame()
{
    CRect rcImg;

    DYNCAST(CImgElement, ElementOwner())->_pImage->GetRectImg(&rcImg);
    Invalidate(&rcImg);
}

//+----------------------------------------------------------------------------
//
// Member: GetMarginInfo
//
//  add hSpace/vSpace to the margin
//
//-----------------------------------------------------------------------------
void CImgElementLayout::GetMarginInfo(
          CParentInfo* ppri,
          LONG* plLMargin,
          LONG* plTMargin,
          LONG* plRMargin,
          LONG* plBMargin)
{
    CImgElement* pImg = DYNCAST(CImgElement, ElementOwner());

    super::GetMarginInfo(ppri, plLMargin, plTMargin, plRMargin, plBMargin);
    Assert(pImg && pImg->_pImage);
    pImg->_pImage->GetMarginInfo(ppri, plLMargin, plTMargin, plRMargin, plBMargin);
}

//+---------------------------------------------------------------------------
//
//  Member:     Draw
//
//  Synopsis:   Paint the object.
//
//----------------------------------------------------------------------------
void CImgElementLayout::Draw(CFormDrawInfo* pDI, CDispNode*)
{
    CImgElement* pImg = DYNCAST(CImgElement, ElementOwner());

    // draw the image
    pImg->_pImage->Draw(pDI);
}

HRESULT CImgElementLayout::PreDrag(DWORD dwKeyState, IDataObject** ppDO, IDropSource** ppDS)
{
    HRESULT                 hr          = S_OK;
    CImgDragDropSrcInfo*    pDragInfo   = NULL;
    CGenDataObject*         pDO         = NULL;
    CImgElement*            pImgElem    = DYNCAST(CImgElement, ElementOwner());
    CImgHelper*             pImg        = pImgElem->_pImage;

    Assert(!Doc()->_pDragDropSrcInfo);

    pDragInfo = new CImgDragDropSrcInfo;
    if(!pDragInfo)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    hr = CreateImgDataObject(Doc(), pImg->_pImgCtx, pImg->_pBitsCtx, ElementOwner(), &pDO);
    if(hr)
    {
        goto Cleanup;
    }
    Assert(pDO);
    pDO->SetBtnState(dwKeyState);

    hr = pDO->QueryInterface(IID_IDataObject, (void**)ppDO);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDO->QueryInterface(IID_IDropSource, (void**)ppDS);
    if(hr)
    {
        goto Cleanup;
    }

    pDragInfo->_pImgDO = pDO;
    Doc()->_pDragDropSrcInfo = pDragInfo;

Cleanup:
    if(hr)
    {
        if(pDragInfo)
        {
            delete pDragInfo;
        }
        if(pDO)
        {
            pDO->Release();
        }
    }
    RRETURN(hr);
}

HRESULT CImgElementLayout::PostDrag(HRESULT hrDrop, DWORD dwEffect)
{
    Assert(Doc()->_pDragDropSrcInfo);

    DYNCAST(CImgDragDropSrcInfo, Doc()->_pDragDropSrcInfo)->_pImgDO->Release();
    RRETURN(S_OK);
}



CImgHelper* CImageLayout::GetImgHelper()
{
    CElement* pElement = ElementOwner();
    CImgHelper* pImgHelper = NULL;

    Assert(pElement);
    if(pElement->Tag() == ETAG_IMG)
    {
        pImgHelper = DYNCAST(CImgElement, pElement)->_pImage;
    }

    Assert(pImgHelper);

    return pImgHelper;
}

//+---------------------------------------------------------------------------
//
//  Member:     HandleViewChange
//
//  Synopsis:   Respond to change of in view status
//
//  Arguments:  flags           flags containing state transition info
//              prcClient       client rect in global coordinates
//              prcClip         clip rect in global coordinates
//              pDispNode       node which moved
//
//----------------------------------------------------------------------------
void CImageLayout::HandleViewChange(
           DWORD        flags,
           const RECT*  prcClient,
           const RECT*  prcClip,
           CDispNode*   pDispNode)
{
    CImgHelper* pImg = GetImgHelper();

    if(!pImg->_fVideoPositioned)
    {
        pImg->_fVideoPositioned = TRUE;
        pImg->SetVideo();
    }

    if(pImg->_hwnd)
    {           
        CRect rcClip(*prcClip);
        UINT uFlags = SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOCOPYBITS;

        if(flags & VCF_INVIEWCHANGED)
        {
            uFlags |= (flags&VCF_INVIEW) ? SWP_SHOWWINDOW : SWP_HIDEWINDOW;
        }

        DeferSetWindowPos(pImg->_hwnd, (RECT*)prcClient, uFlags, NULL);

        rcClip.OffsetRect(-((const CRect*)prcClient)->TopLeft().AsSize());
        ::SetWindowRgn(pImg->_hwnd, ::CreateRectRgnIndirect(&rcClip), !(flags&VCF_NOREDRAW));

        if(pImg->_pVideoObj)
        {
            RECT rcImg;

            rcImg.top = 0;
            rcImg.left = 0;
            rcImg.bottom = prcClient->bottom - prcClient->top;
            rcImg.right = prcClient->right - prcClient->left;
            pImg->_pVideoObj->SetWindowPosition(&rcImg);
        }
    }

    return;
}