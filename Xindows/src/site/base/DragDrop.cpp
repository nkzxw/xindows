
#include "stdafx.h"
#include "DragDrop.h"

CDragDropTargetInfo::CDragDropTargetInfo(CDocument* pDoc)
{
    _pDoc = pDoc;
}

CDragDropTargetInfo::~CDragDropTargetInfo()
{
    ReleaseInterface(_pStart);
    ReleaseInterface(_pEnd);
}

HRESULT CDragDropTargetInfo::StoreSelection()
{
    HRESULT hr = S_OK;

    ISegmentList* pSegmentList = NULL;
    IMarkupServices* pMarkup = NULL;

    int ctSegment = 0;

    hr = _pDoc->GetCurrentSelectionSegmentList(&pSegmentList);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pSegmentList->GetSegmentCount(&ctSegment, &_eType);           
    if(hr)
    {
        goto Cleanup; 
    }

    if(ctSegment > 0)
    {
        hr = _pDoc->QueryInterface(IID_IMarkupServices, (void**)&pMarkup);
        if(hr)
        {
            goto Cleanup; 
        }

        hr = pMarkup->CreateMarkupPointer(&_pStart);
        if(hr)
        {
            goto Cleanup; 
        }

        hr = pMarkup->CreateMarkupPointer(&_pEnd);
        if(hr)
        {
            goto Cleanup; 
        }

        hr = pSegmentList->MovePointersToSegment(0, _pStart, _pEnd);

        _pElemCurrentAtStoreSel = _pDoc->_pElemCurrent;
    }

Cleanup:
    ReleaseInterface(pSegmentList);
    ReleaseInterface(pMarkup);
    RRETURN(hr);
}

HRESULT CDragDropTargetInfo::RestoreSelection()
{
    HRESULT hr = S_OK;

    // Ensure that the current element we began a drag & drop with is the same
    // as the currency now. This makes us not whack selection back - if our host
    // has changed currency on us in some way ( as access does over ole controls)
    if(_pDoc->_pElemCurrent==_pElemCurrentAtStoreSel &&
        _pStart && _pEnd && _eType!=SELECTION_TYPE_None)
    {
        hr = _pDoc->Select(_pStart, _pEnd, _eType);
    }

    _pElemCurrentAtStoreSel = NULL;
    RRETURN(hr);
}




CDragStartInfo::CDragStartInfo(CElement* pElementDrag, DWORD dwStateKey, IUniformResourceLocator* pUrlToDrag)
{
    _pElementDrag = pElementDrag;
    _pElementDrag->SubAddRef();
    _dwStateKey = dwStateKey;
    _pUrlToDrag = pUrlToDrag;
    if(_pUrlToDrag)
    {
        _pUrlToDrag->AddRef();
    }
    _dwEffectAllowed = DROPEFFECT_UNINITIALIZED;
}

CDragStartInfo::~CDragStartInfo()
{
    _pElementDrag->SubRelease();
    ReleaseInterface(_pUrlToDrag);
    ReleaseInterface(_pDataObj);
    ReleaseInterface(_pDropSource);
}

HRESULT CDragStartInfo::CreateDataObj()
{
    CLayout* pLayout = _pElementDrag->GetUpdatedNearestLayout();

    if(!pLayout)
    {
        CMarkup* pMarkup = _pElementDrag->GetMarkup();

        if(pMarkup->Master())
        {
            pLayout = pMarkup->Master()->GetUpdatedLayout();
        }
    }

    RRETURN(pLayout ? pLayout->DoDrag(_dwStateKey, _pUrlToDrag, TRUE) : E_FAIL);
}
