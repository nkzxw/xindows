
#include "stdafx.h"

//+----------------------------------------------------------------------------
//
//  Function:   RemoveWithBreakOnEmpty
//
//  Synopsis:   Helper fcn which removes stuff between two pointers, and
//              sets the break on empty bit of the block element above
//              the left hand side of the removal.  This is needed for
//              compatibility with IE4 RemoveChars which was used for most
//              text removal operations.
//
//-----------------------------------------------------------------------------
HRESULT RemoveWithBreakOnEmpty(CMarkupPointer* pPointerStart, CMarkupPointer* pPointerFinish)
{
    CMarkup*    pMarkup = pPointerStart->Markup();
    CDocument*  pDoc = pPointerStart->Doc();
    CTreeNode*  pNodeBlock;
    HRESULT     hr = S_OK;

    // First, locate the block element above here and set the break on empty bit
    pNodeBlock = pMarkup->SearchBranchForBlockElement(pPointerStart->Branch());

    if(pNodeBlock)
    {
        pNodeBlock->Element()->_fBreakOnEmpty = TRUE;
    }

    // Now, do the remove
    hr = pDoc->Remove(pPointerStart, pPointerFinish);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}


//+----------------------------------------------------------------------------
//
//  Function:   UnoverlapPartials
//
//  Synopsis:   Assuming the contents of an element have been gutted by a
//              Remove operation, this function will take any elements which
//              partially overlapp it, and remove that partial overalpping.
//
//              For example: calling this on y where we have y overlapped
//              from both sides: "<x><y></x><z></y></z>" will produce
//              "<x><y></y></x><z></z>".
//
//-----------------------------------------------------------------------------
#ifdef _DEBUG
static void IsBefore(
          CElement* pElement1, ELEMENT_ADJACENCY eAdj1,
          CElement* pElement2, ELEMENT_ADJACENCY eAdj2)
{
    CMarkupPointer pmp1(pElement1->Doc());
    CMarkupPointer pmp2(pElement2->Doc());

    pmp1.MoveAdjacentToElement(pElement1, eAdj1);
    pmp2.MoveAdjacentToElement(pElement2, eAdj2);

    Assert(pmp1.IsLeftOf(&pmp2));
}
#endif

HRESULT UnoverlapPartials(CElement* pElement)
{
    HRESULT         hr = S_OK;
    CDocument*      pDoc = pElement->Doc();
    CMarkupPointer  pmp(pDoc);
    CElement*       pElementOverlap = NULL;
    CMarkupPointer  pmpStart(pDoc), pmpFinish(pDoc);

    Assert(pElement->IsInMarkup());

    pElement->GetMarkup()->AddRef();

    hr = pmp.MoveAdjacentToElement(pElement, ELEM_ADJ_BeforeEnd);

    if(hr)
    {
        goto Cleanup;
    }

    for(;;)
    {
        MARKUP_CONTEXT_TYPE ct;
        CTreeNode* pNodeOverlap;

        hr = pmp.Left(TRUE, &ct, &pNodeOverlap, NULL, NULL, NULL);

        if(hr)
        {
            goto Cleanup;
        }

        // All text and no-scopes better have been removed.
        if(ct!=CONTEXT_TYPE_EnterScope && ct!=CONTEXT_TYPE_ExitScope)
        {
            break;
        }

        if(ct==CONTEXT_TYPE_ExitScope && pNodeOverlap->Element()==pElement)
        {
            break;
        }

        pElementOverlap = pNodeOverlap->Element();
        pElementOverlap->AddRef();

        if(ct == CONTEXT_TYPE_EnterScope)
        {
            DEBUG_ONLY(IsBefore(pElementOverlap, ELEM_ADJ_BeforeBegin, pElement, ELEM_ADJ_BeforeBegin));

            hr = pmpStart.MoveAdjacentToElement(pElementOverlap, ELEM_ADJ_BeforeBegin);

            if(hr)
            {
                goto Cleanup;
            }

            hr = pmpFinish.MoveAdjacentToElement(pElement, ELEM_ADJ_AfterEnd);

            if(hr)
            {
                goto Cleanup;
            }
        }
        else
        {
            DEBUG_ONLY(IsBefore( pElement, ELEM_ADJ_AfterEnd, pElementOverlap, ELEM_ADJ_AfterEnd));

            hr = pmpFinish.MoveAdjacentToElement(pElementOverlap, ELEM_ADJ_AfterEnd);

            if(hr)
            {
                goto Cleanup;
            }

            hr = pmpStart.MoveAdjacentToElement(pElement, ELEM_ADJ_AfterEnd);

            if(hr)
            {
                goto Cleanup;
            }
        }

        hr = pDoc->RemoveElement(pElementOverlap);

        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->InsertElement(pElementOverlap, &pmpStart, &pmpFinish);

        if(hr)
        {
            goto Cleanup;
        }

        pElementOverlap->Release();

        pElementOverlap = NULL;
    }

Cleanup:

    pElement->GetMarkup()->Release();

    if(pElementOverlap)
    {
        pElementOverlap->Release();
    }

    RRETURN(hr);
}