
#include "stdafx.h"
#include "CutCommand.h"

using namespace EdUtil;

//+---------------------------------------------------------------------------
//
//  CCutCommand::Exec
//
//----------------------------------------------------------------------------
HRESULT CCutCommand::PrivateExec(DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut)
{
    HRESULT             hr = S_OK;
    IHTMLElement*       pElement = NULL;
    int                 iSegmentCount;
    IMarkupPointer*     pStart = NULL;
    IMarkupPointer*     pEnd = NULL;
    SELECTION_TYPE      eSelectionType;
    IMarkupServices*    pMarkupServices = GetMarkupServices();
    ISegmentList*       pSegmentList = NULL;
    BOOL                fRet;
    BOOL                fNotRange = TRUE;
    CUndoUnit           undoUnit(GetEditor());
    CXindowsEditor*     pEditor = GetEditor();

    ((IHTMLEditor*)pEditor)->AddRef();

    // Do the prep work
    IFC(GetSegmentList(&pSegmentList));    
    IFC(pSegmentList->GetSegmentCount(&iSegmentCount, &eSelectionType));           

    // Cut is allowed iff we have a non-empty segment
    if(eSelectionType==SELECTION_TYPE_Caret || iSegmentCount==0)
    {
        goto Cleanup;
    }
    IFC(pMarkupServices->CreateMarkupPointer(&pStart));
    IFC(pMarkupServices->CreateMarkupPointer(&pEnd));   
    IFC(MovePointersToSegmentHelper(GetViewServices(), pSegmentList, 0, &pStart, &pEnd));
    IFC(pStart->IsEqualTo(pEnd, &fRet));
    if(fRet)
    {
        goto Cleanup;
    }

    // Cannot delete or cut unless the range is in the same flow layout
    if(!PointersInSameFlowLayout(pStart, pEnd, NULL, GetViewServices()))
    {
        goto Cleanup;
    }

    // Now Handle the cut 
    /*IFC(undoUnit.Begin(IDS_EDUNDOCUT)); wlw note*/

    IFC(FindCommonElement(pMarkupServices, GetViewServices(), pStart, pEnd, &pElement));

    if(!pElement)
    {
        goto Cleanup;
    }

    IFC(GetViewServices()->FireCancelableEvent( 
        pElement,
        DISPID_EVMETH_ONCUT,
        DISPID_EVPROP_ONCUT,
        _T("cut"),
        &fRet));

    if(!fRet)
    {
        goto Cleanup;
    }

    IFC(GetViewServices()->SaveSegmentsToClipboard(pSegmentList));

    fNotRange = (eSelectionType!=SELECTION_TYPE_Auto && eSelectionType!=SELECTION_TYPE_Control);
    IFC(pEditor->Delete(pStart, pEnd, fNotRange));

    if(eSelectionType == SELECTION_TYPE_Selection)
    {
        pEditor->GetSelectionManager()->EmptySelection();
    }

Cleanup:   
    ReleaseInterface((IHTMLEditor*)pEditor);
    ReleaseInterface(pSegmentList);
    ReleaseInterface(pStart);
    ReleaseInterface(pEnd);
    ReleaseInterface(pElement);
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  CCutCommand::QueryStatus
//
//----------------------------------------------------------------------------
HRESULT CCutCommand::PrivateQueryStatus(OLECMD* pCmd, OLECMDTEXT* pcmdtext)
{
    HRESULT             hr = S_OK;
    INT                 iSegmentCount;
    IMarkupPointer*     pStart = NULL;
    IMarkupPointer*     pEnd = NULL;
    IHTMLElement*       pElement = NULL;
    BOOL                fEditable;
    IMarkupServices*    pMarkupServices = GetMarkupServices();
    IHTMLViewServices*  pViewServices = GetViewServices();
    ISegmentList*       pSegmentList = NULL;
    SELECTION_TYPE      eSelectionType;
    BOOL                fRet;

    pViewServices->AddRef();

    // Status is disabled by default
    pCmd->cmdf = MSOCMDSTATE_DISABLED;

    // Get Segment list and selection type
    IFC(GetSegmentList(&pSegmentList));
    IFC(pSegmentList->GetSegmentCount(&iSegmentCount, &eSelectionType));

    // Cut is allowed iff we have a non-empty segment
    if(eSelectionType==SELECTION_TYPE_Caret || iSegmentCount==0)
    {
        goto Cleanup;
    }
    IFC(GetSegmentPointers(pSegmentList, 0, &pStart, &pEnd));
    IFC(pStart->IsEqualTo(pEnd, &fRet));
    if(fRet)
    {
        goto Cleanup;
    }

    // Fire cancelable event
    IFC(FindCommonElement(pMarkupServices, GetViewServices(), pStart, pEnd, &pElement));
    if(!pElement) 
    {
        goto Cleanup;
    }

    IFC(pViewServices->FireCancelableEvent(
        pElement, 
        DISPID_EVMETH_ONBEFORECUT,
        DISPID_EVPROP_ONBEFORECUT,
        _T("beforecut"),
        &fRet));

    if(!fRet)
    {
        pCmd->cmdf = MSOCMDSTATE_UP; 
        goto Cleanup;
    }

    if(eSelectionType!=SELECTION_TYPE_Auto && eSelectionType!=SELECTION_TYPE_Control)
    {
        IFC(pViewServices->IsEditableElement(pElement, &fEditable));
        if((!fEditable))
        {
            goto Cleanup;
        }
    }

    // Cannot delete or cut unless the range is in the same flow layout
    if(PointersInSameFlowLayout(pStart, pEnd, NULL, pViewServices))
    {
        pCmd->cmdf = MSOCMDSTATE_UP; 
    }

Cleanup:
    ReleaseInterface(pViewServices);
    ReleaseInterface(pSegmentList);
    ReleaseInterface(pStart);
    ReleaseInterface(pEnd);
    ReleaseInterface(pElement);
    RRETURN(hr);
}