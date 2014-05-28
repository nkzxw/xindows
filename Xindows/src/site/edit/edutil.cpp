//+---------------------------------------------------------------------
//
//  Copyright (C) Microsoft Corporation, 1998.
//
//  File:       EDUTIL.CXX
//
//  Contents:   Utility functions for CMsHtmled
//
//  History:    15-Jan-98   raminh  Created
//
//  Notes:      This file contains some utility functions from Trident,
//              such as LoadLibrary, which have been modified to eliminate
//              dependencies. In addition, it provides the implementation
//              for editing commands such as InsertObject etc.
//------------------------------------------------------------------------

#include "stdafx.h"
#include "edutil.h"

#include "sload.h"
#include "edcmd.h"
#include "blockcmd.h"
#include "SelectionManager.h"

using namespace EdUtil;

Direction Reverse( Direction iDir )
{
    if( iDir == LEFT )
        return RIGHT;
    else if (iDir == RIGHT)
        return LEFT;
    else
        return iDir;
}



//////////////////////////////////////////////////////////////////////////
//
//  CEditPointer's Constructor/Destructor Implementation
//
//////////////////////////////////////////////////////////////////////////

CEditPointer::CEditPointer(
    CXindowsEditor *       pEd,
    IMarkupPointer *    pPointer /* = NULL */ )
{
     _pEd = pEd;
    _pPointer = NULL;
    _pLeftBoundary = NULL;
    _pRightBoundary = NULL;
    _fBound = FALSE;
    
    if( pPointer != NULL )
    {
        _pPointer = pPointer;
        _pPointer->AddRef();
    }
    else
    {
        _pEd->GetMarkupServices()->CreateMarkupPointer( & _pPointer);
    }
}
    

CEditPointer::CEditPointer(
    const CEditPointer& lp )
{
    _pEd = lp._pEd;
    
    if ((_pPointer = lp._pPointer) != NULL)
        _pPointer->AddRef();

    if ((_pLeftBoundary = lp._pLeftBoundary) != NULL)
        _pLeftBoundary->AddRef();

    if ((_pRightBoundary = lp._pRightBoundary) != NULL)
        _pRightBoundary->AddRef();

    _fBound = lp._fBound;
}


CEditPointer::~CEditPointer()
{
    ReleaseInterface( _pPointer );
    ReleaseInterface( _pLeftBoundary );
    ReleaseInterface( _pRightBoundary );
}



//////////////////////////////////////////////////////////////////////////
//
//  CEditPointer's Method Implementations
//
//////////////////////////////////////////////////////////////////////////

HRESULT
CEditPointer::SetBoundary(
    IMarkupPointer *    pLeftBoundary,
    IMarkupPointer *    pRightBoundary )
{
    HRESULT hr = S_OK;

    ReplaceInterface( & _pLeftBoundary , pLeftBoundary );
    ReplaceInterface( & _pRightBoundary , pRightBoundary );
    _fBound = TRUE;
    
    RRETURN( hr );
}

HRESULT
CEditPointer::SetBoundaryForDirection(
    Direction       eDir,
    IMarkupPointer* pBoundary )
{
    HRESULT hr = S_OK;

    if( eDir == LEFT )
        hr = SetBoundary( pBoundary, NULL);
    else
        hr = SetBoundary( NULL, pBoundary);

    RRETURN( hr );
}


HRESULT
CEditPointer::ClearBoundary()
{
    HRESULT hr = S_OK;
    ClearInterface( & _pRightBoundary );
    ClearInterface( & _pLeftBoundary );
    _fBound = FALSE;
    RRETURN( hr );
}


BOOL
CEditPointer::IsPointerInLeftBoundary()
{
    BOOL fWithin = TRUE;

    if( _fBound && _pLeftBoundary )
    {
        IsRightOfOrEqualTo( _pLeftBoundary, & fWithin ); // we are within if we are to the right or equal to the left boundary
    }
    
    return fWithin;
}


BOOL
CEditPointer::IsPointerInRightBoundary()
{
    BOOL fWithin = TRUE;

    if( _fBound && _pRightBoundary )
    {
        IsLeftOfOrEqualTo( _pRightBoundary, & fWithin ); // we are within if we are to the left or equal to the left boundary
    }
    
    return fWithin;
}


BOOL 
CEditPointer::IsWithinBoundary()
{  
    BOOL fWithin = TRUE;

    if( _fBound )
    {
        if( _pLeftBoundary )    
        {
            IsRightOfOrEqualTo( _pLeftBoundary, & fWithin);
        }

        if( fWithin && _pRightBoundary )
        {
            IsLeftOfOrEqualTo( _pRightBoundary, & fWithin);
        }
    }
    
    return fWithin;
}


BOOL
CEditPointer::IsWithinBoundary( 
    Direction inDir )
{
    if( ! _fBound )
        return TRUE;
        
    if( inDir == LEFT )
    {
        return( IsPointerInRightBoundary() );       
    }
    else
    {
        return( IsPointerInLeftBoundary() );
    }
}


HRESULT
CEditPointer::Constrain()
{
    HRESULT hr = S_OK;

    if( _fBound )
    {
        if( ! IsPointerInLeftBoundary() )
        {
            Assert( _pLeftBoundary );
            IFC( _pPointer->MoveToPointer( _pLeftBoundary ));
        }

        if( ! IsPointerInRightBoundary() )
        {
            Assert( _pRightBoundary );
            IFC( _pPointer->MoveToPointer( _pRightBoundary ));
        }
    }
    
Cleanup:
    RRETURN( hr );
}



HRESULT
CEditPointer::Scan(
    Direction               eDir,
    DWORD                   eBreakCondition,
    DWORD *                 peBreakConditionFound,
    IHTMLElement **         ppElement,
    ELEMENT_TAG_ID *        peTagId,
    TCHAR *                 pChar,
    DWORD                   eScanOptions)
{
    HRESULT hr = S_OK;
    BOOL fDone = FALSE;
    LONG lChars;
    TCHAR tChar = 0;
    DWORD dwContextAdjustment;
    MARKUP_CONTEXT_TYPE eCtxt = CONTEXT_TYPE_None;
    DWORD eBreakFound = BREAK_CONDITION_None;
    DWORD dwTest = BREAK_CONDITION_None;
    SP_IHTMLElement spElement;
    ELEMENT_TAG_ID eTagId = TAGID_NULL;
 

    IFC( Constrain() );
    
    while( ! fDone )
    {
        lChars = 1;
        tChar = 0;
        dwContextAdjustment = 1;

        if(    ( eBreakCondition & BREAK_CONDITION_Text && ( ! CheckFlag( eScanOptions , SCAN_OPTION_ChunkifyText )))
            || ( eBreakCondition & BREAK_CONDITION_NoScopeBlock ))
        {
            hr = Move( eDir, TRUE, & eCtxt, & spElement, & lChars, &tChar);
        }
        else
        {
            hr = Move( eDir, TRUE, & eCtxt, & spElement, NULL, NULL);
        }
        
        if( hr == E_HITBOUNDARY )
        {
            fDone = TRUE;
            eBreakFound = BREAK_CONDITION_Boundary;
            hr = S_OK;
            break;
        }
        
        switch( eCtxt )
        {
            case CONTEXT_TYPE_ExitScope:
                dwContextAdjustment = 2;
                // FALL THROUGH
            case CONTEXT_TYPE_EnterScope:
            {
                IFC( _pEd->GetMarkupServices()->GetElementTagId( spElement, &eTagId ));
            
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Site ))
                {
                    BOOL fTest = FALSE;
                    
                    IFC( _pEd->GetViewServices()->IsSite( spElement, &fTest, NULL, NULL, NULL ));
                    if( fTest )
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterSite;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }
                
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_TextSite ))
                {
                    BOOL fTest = FALSE;
                    BOOL fText = FALSE;
                    
                    IFC( _pEd->GetViewServices()->IsSite( spElement, &fTest, &fText, NULL, NULL ));
                    
                    if( fTest && fText )
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterTextSite;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }
                
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Block ))
                {
                    BOOL fTest = FALSE;
                    IFC( _pEd->GetViewServices()->IsBlockElement( spElement, &fTest ));
                    
                    if( fTest )
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterBlock;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }

                if( CheckFlag( eBreakCondition , BREAK_CONDITION_BlockPhrase ))
                {
                    if( eTagId == TAGID_RT || eTagId == TAGID_RP )
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterBlockPhrase;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }

                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Control ) )
                {
                    if( IsIntrinsic( _pEd->GetMarkupServices(), spElement ))
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterControl;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }
                else if (eScanOptions & SCAN_OPTION_SkipControls)
                {
                    if( IsIntrinsic( _pEd->GetMarkupServices(), spElement ) )
                    {
                        IFC( _pPointer->MoveAdjacentToElement(spElement, (eDir == RIGHT) ? ELEM_ADJ_AfterEnd : ELEM_ADJ_BeforeBegin) );
                        continue;
                    }                    
                }
                
                            
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Anchor ))
                {
                    if( eBreakFound == BREAK_CONDITION_None && eTagId == TAGID_A ) // we didn't hit layout, block, intrinsic or no-scope
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterAnchor;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }
                
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Phrase ))
                {
                    if( eBreakFound == BREAK_CONDITION_None ) // we didn't hit layout, block, intrinsic, no-scope or anchor
                    {
                        dwTest = dwContextAdjustment * BREAK_CONDITION_EnterPhrase;
                        if( CheckFlag( eBreakCondition, dwTest ))
                        {
                            eBreakFound |= dwTest;
                            fDone = TRUE;
                        }
                    }
                }
                
                break;
            }

            case CONTEXT_TYPE_NoScope:
            {
                IFC( _pEd->GetMarkupServices()->GetElementTagId( spElement, &eTagId ));
            
                
                // this only looks a bit strange because there is no begin/end NoScope
                
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_NoScope ))
                {
                    if( eCtxt == CONTEXT_TYPE_NoScope )
                    {
                        eBreakFound |= BREAK_CONDITION_NoScope;
                        fDone = TRUE;
                    }
                }

                // Could be a noscope with layout. For example, an image.
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_NoScopeSite ))
                {
                    BOOL fTest = FALSE;
                    
                    IFC( _pEd->GetViewServices()->IsSite( spElement, &fTest, NULL, NULL, NULL ));
                    if( fTest )
                    {
                        eBreakFound |= BREAK_CONDITION_NoScopeSite;
                        fDone = TRUE;
                    }
                }
                
                if( CheckFlag( eBreakCondition , BREAK_CONDITION_NoScopeBlock ))
                {
                    BOOL fTest = FALSE;
                    IFC( _pEd->GetViewServices()->IsBlockElement( spElement, &fTest ));
                    
                    if( fTest || eTagId == TAGID_BR )
                    {
                        eBreakFound |= BREAK_CONDITION_NoScopeBlock;
                        fDone = TRUE;
                    }
                }                

                if( CheckFlag( eBreakCondition , BREAK_CONDITION_Control ))
                {
                    if( IsIntrinsic( _pEd->GetMarkupServices(), spElement ))
                    {
                        eBreakFound |= BREAK_CONDITION_Control;
                        fDone = TRUE;
                    }
                }
                break;                            
            }
            
            case CONTEXT_TYPE_Text:
            {
                if (CheckFlag( eBreakCondition , BREAK_CONDITION_Text )
                    && (   (   (eScanOptions & SCAN_OPTION_SkipWhitespace) 
                            && IsWhiteSpace(tChar))
                        || (   (eScanOptions & SCAN_OPTION_SkipNBSP) 
                            && tChar == WCH_NBSP)))
                {
                    continue; 
                }
                else if( CheckFlag( eBreakCondition , BREAK_CONDITION_Text ) && tChar != _T('\r') )
                {
                    eBreakFound |= BREAK_CONDITION_Text;
                    fDone = TRUE;
                }
                else if( CheckFlag( eBreakCondition, BREAK_CONDITION_NoScopeBlock ) && tChar == _T('\r'))
                {
                    eBreakFound |= BREAK_CONDITION_NoScopeBlock;
                    fDone = TRUE;
                }

                break;
            }

            case CONTEXT_TYPE_None:
            {
                // An error has occured
                eBreakFound |= BREAK_CONDITION_Error;
                fDone = TRUE;
                break;
            }
        }
    }
    
Cleanup:

    if( peBreakConditionFound )
        *peBreakConditionFound = eBreakFound;
        
    if( ppElement )
    {
        ReplaceInterface( ppElement, spElement.p );
    }

    if( peTagId )
        *peTagId = eTagId;

    if( pChar )
        *pChar = tChar;
    
    RRETURN( hr );
}


HRESULT 
CEditPointer::Move(                                     // Directional Wrapper for Left or Right
    Direction               inDir,                      //      [in]     Direction of travel
    BOOL                    fMove,                      //      [in]     Should we actually move the pointer
    MARKUP_CONTEXT_TYPE*    pContext,                   //      [out]    Context change
    IHTMLElement**          ppElement,                  //      [out]    Element we pass over
    long*                   pcch,                       //      [in,out] number of characters to read back
    OLECHAR*                pchText )                   //      [out]    characters
{
    HRESULT hr = E_FAIL;
    
    Assert( _pPointer );
    
    if( inDir == LEFT )
    {
        IFC( _pEd->GetViewServices()->LeftOrSlave(_pPointer, fMove, pContext, ppElement, pcch, pchText ));
    }
    else
    {
        IFC( _pEd->GetViewServices()->RightOrSlave(_pPointer, fMove, pContext, ppElement, pcch, pchText ));
    }

    if( ! IsWithinBoundary() )
    {
        Constrain();
        hr = E_HITBOUNDARY;
    }

Cleanup:
    RRETURN( hr );
}



//+===================================================================================
// Method:      MoveWord
//
// Synopsis:    Moves the pointer to the previous or next word. This method takes into
//              account block and site ends.
//
// Parameters:  
//              eDir            [in]    Direction to move
//              pfNotAtBOL      [out]   What line is pointer on after move? (optional)
//              pfAtLogcialBOL  [out]   Is pointer at lbol after move? (otional)
//+===================================================================================

HRESULT 
CEditPointer::MoveWord(
    Direction               eDir, 
    BOOL *                  pfNotAtBOL     /* =NULL */ , 
    BOOL *                  pfAtLogicalBOL /* =NULL */  )   
{
    HRESULT hr = S_OK;
    
    if( eDir == LEFT )
        hr = MoveUnit( eDir, MOVEUNIT_PREVWORDBEGIN, pfNotAtBOL, pfAtLogicalBOL );
    else
        hr = MoveUnit( eDir, MOVEUNIT_NEXTWORDBEGIN, pfNotAtBOL, pfAtLogicalBOL );

    RRETURN( hr );
}


//+===================================================================================
// Method:      MoveCharacter
//
// Synopsis:    Moves the pointer to the previous or next character. This method takes
//              into account block and site ends.
//
// Parameters:  
//              eDir            [in]    Direction to move
//              pfNotAtBOL      [out]   What line is pointer on after move? (optional)
//              pfAtLogcialBOL  [out]   Is pointer at lbol after move? (otional)
//+===================================================================================

HRESULT 
CEditPointer::MoveCharacter(
    Direction               eDir,
    BOOL *                  pfNotAtBOL      /*=NULL*/,
    BOOL *                  pfAtLogicalBOL  /*=NULL*/ )
{
    HRESULT hr = S_OK;
    BOOL fNearText = FALSE;
    CEditPointer tLooker( _pEd );
    DWORD dwBreak = BREAK_CONDITION_OMIT_PHRASE-BREAK_CONDITION_Anchor;
    DWORD dwFound = BREAK_CONDITION_None;
    IFC( tLooker->MoveToPointer( this ));
    IFC( tLooker.Scan( eDir, dwBreak, &dwFound ));

    fNearText =  CheckFlag( dwFound, BREAK_CONDITION_Text );
    
    if( eDir == LEFT )
    {
        if( fNearText )
        {
            hr = MoveUnit( eDir, MOVEUNIT_PREVCLUSTERBEGIN, pfNotAtBOL, pfAtLogicalBOL );
        }
        else
        {
            hr = MoveUnit( eDir, MOVEUNIT_PREVCLUSTEREND, pfNotAtBOL, pfAtLogicalBOL );
        }
    }
    else
    {
        if( fNearText )
        {
            hr = MoveUnit( eDir, MOVEUNIT_NEXTCLUSTEREND, pfNotAtBOL, pfAtLogicalBOL );
        }
        else
        {
            hr = MoveUnit( eDir, MOVEUNIT_NEXTCLUSTERBEGIN, pfNotAtBOL, pfAtLogicalBOL );
        }
    }
Cleanup:
    RRETURN( hr );
}


//+===================================================================================
// Method:      MoveUnit
//
// Synopsis:    Moves the pointer to the previous or next character. This method takes
//              into account block and site ends.
//
// Parameters:  
//              eDir            [in]    Direction to move
//              pfNotAtBOL      [out]   What line is pointer on after move? (optional)
//              pfAtLogcialBOL  [out]   Is pointer at lbol after move? (otional)
//+===================================================================================

HRESULT 
CEditPointer::MoveUnit(
    Direction               eDir,
    MOVEUNIT_ACTION         eUnit,
    BOOL *                  pfNotAtBOL      /*=NULL*/,
    BOOL *                  pfAtLogicalBOL  /*=NULL*/ )
{
    HRESULT hr = S_OK;
    BOOL fNotAtBOL = TRUE;
    BOOL fAtLogicalBOL = FALSE;
    BOOL fBetweenLines;
    BOOL fBeyondThisLine;
    BOOL fAtEdgeOfLine;
    BOOL fThereIsAnotherLine = FALSE;
    BOOL fBeyondNextLine = FALSE;
    BOOL fLineBreakDueToTextWrapping = FALSE;
    DWORD dwBreak = BREAK_CONDITION_Site | BREAK_CONDITION_NoScopeSite | BREAK_CONDITION_Control;
    DWORD dwFound = BREAK_CONDITION_None;
    SP_IHTMLElement spSite;
    CEditPointer epWalker( _pEd );
    CEditPointer epDestination( _pEd );
    CEditPointer epBoundary( _pEd );
    CEditPointer epNextLine( _pEd );

    Assert( pfNotAtBOL && pfAtLogicalBOL );
    if( ! pfNotAtBOL || ! pfAtLogicalBOL )
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    fNotAtBOL = *pfNotAtBOL;
    fAtLogicalBOL = *pfAtLogicalBOL;

    IFC( epDestination->MoveToPointer( this ));
    IFC( epBoundary->MoveToPointer( this ));
    IFC( epWalker->MoveToPointer( this ));
    IFC( epNextLine->MoveToPointer( this ));

    IFC( epDestination->MoveUnit( eUnit ));


    if( eDir == LEFT )
    {
        BOOL tfNotAtBOL = *pfNotAtBOL;
        BOOL tfAtLogicalBOL = *pfNotAtBOL;
        DWORD dwIgnore = BREAK_CONDITION_Phrase | BREAK_CONDITION_Anchor;
        IFC( _pEd->GetViewServices()->MoveMarkupPointer( epBoundary, LAYOUT_MOVE_UNIT_CurrentLineStart, -1, &tfNotAtBOL, &tfAtLogicalBOL ));
        
        hr = _pEd->GetViewServices()->MoveMarkupPointer( epNextLine, LAYOUT_MOVE_UNIT_PreviousLineEnd, -1, &tfNotAtBOL, &tfAtLogicalBOL );
        if( ! FAILED( hr ))
        {
            fThereIsAnotherLine = TRUE;
            IFC( epDestination->IsLeftOf( epNextLine, &fBeyondNextLine ));
        }

        IFC( epNextLine->IsEqualTo( epBoundary, &fLineBreakDueToTextWrapping )); // If the current line start and previous line end are the same point in the markup, we are breaking the line due to wrapping
        IFC( epDestination->IsLeftOf( epBoundary, &fBeyondThisLine ));
        IFC( epWalker.IsLeftOfOrEqualTo( epBoundary, dwIgnore, &fAtEdgeOfLine ));
    }
    else
    {
        BOOL tfNotAtBOL = *pfNotAtBOL;
        BOOL tfAtLogicalBOL = *pfNotAtBOL;
        DWORD dwIgnore = BREAK_CONDITION_Phrase | BREAK_CONDITION_Anchor;
        IFC( _pEd->GetViewServices()->MoveMarkupPointer( epBoundary, LAYOUT_MOVE_UNIT_CurrentLineEnd, -1, &tfNotAtBOL, &tfAtLogicalBOL ));
        hr = _pEd->GetViewServices()->MoveMarkupPointer( epNextLine, LAYOUT_MOVE_UNIT_NextLineStart, -1, &tfNotAtBOL, &tfAtLogicalBOL );
        if( ! FAILED( hr ))
        {
            fThereIsAnotherLine = TRUE;
            IFC( epDestination->IsRightOf( epNextLine, &fBeyondNextLine ));
        }

        IFC( epNextLine->IsEqualTo( epBoundary, &fLineBreakDueToTextWrapping )); // If the current line END and previous line START are the same point in the markup, we are breaking the line due to wrapping
        IFC( epDestination->IsRightOf( epBoundary, &fBeyondThisLine ));
        IFC( epWalker.IsRightOfOrEqualTo( epBoundary, dwIgnore, &fAtEdgeOfLine ));
    }

    //
    // If I'm not at the edge of the line, my destination is the edge of the line.
    //
    
    if( ! fAtEdgeOfLine && fBeyondThisLine )
    {
        IFC( epDestination->MoveToPointer( epBoundary ));
    }

    //
    // If I am at the edge of the line and there is another line, and my destination
    // is beyond that line - my destination is that line.
    //

    if( fAtEdgeOfLine && fBeyondThisLine && fBeyondNextLine && ! fLineBreakDueToTextWrapping )
    {
        // we are at the edge of the line and our destination is beyond the next line boundary
        // so move our destination to that line boundary.
        IFC( epDestination->MoveToPointer( epNextLine ));
    }

    //
    // Scan towards my destination. If I hit a site boundary, move to the other
    // side of it and be done. Otherwise, move to the next line.
    //
    
    IFC( epWalker.SetBoundaryForDirection( eDir, epDestination ));
    hr = epWalker.Scan( eDir, dwBreak, &dwFound, &spSite );

    if( CheckFlag( dwFound, BREAK_CONDITION_NoScopeSite ) ||
        CheckFlag( dwFound, BREAK_CONDITION_EnterControl ))
    {
        IFC( epWalker->MoveAdjacentToElement( spSite , eDir == LEFT ? ELEM_ADJ_BeforeBegin : ELEM_ADJ_AfterEnd ));
        goto CalcBOL;
    }
    else if( CheckFlag( dwFound, BREAK_CONDITION_ExitControl ))
    {
        // do not move at all
        goto Cleanup;
    }
    else if( CheckFlag( dwFound, BREAK_CONDITION_Site ))
    {
        // move wherever scan put us...
        if( eDir == LEFT )
        {
            fNotAtBOL = TRUE;
            fAtLogicalBOL = FALSE;
        }
        else
        {
            fNotAtBOL = FALSE;
            fAtLogicalBOL = TRUE;
        }

        goto Done;
    }
    else if( CheckFlag( dwFound, BREAK_CONDITION_Boundary ))
    {
        // No site transitions between here and our destination.

        if( fBeyondThisLine && fAtEdgeOfLine && ! fLineBreakDueToTextWrapping )
        {
            // If our destination pointer is on another line than our start pointer...
            
            IFC( epWalker->MoveToPointer( this ));
            
            if( eDir == LEFT )
            {
                IFC( _pEd->GetViewServices()->MoveMarkupPointer( epWalker, LAYOUT_MOVE_UNIT_PreviousLineEnd, -1, &fNotAtBOL, &fAtLogicalBOL ));
            }
            else
            {
                IFC( _pEd->GetViewServices()->MoveMarkupPointer( epWalker, LAYOUT_MOVE_UNIT_NextLineStart, -1, &fNotAtBOL, &fAtLogicalBOL ));
            }
            
            goto Done;
        }
        else
        {
            // We started and ended on same line with no little stops along the way - move to destination...
            
            IFC( epWalker->MoveToPointer( epDestination ));
        }
    }
    else
    {
        // we hit some sort of error, go to cleanup
        hr = E_FAIL;
        goto Cleanup;
    }

CalcBOL:

    if( pfNotAtBOL || pfAtLogicalBOL )
    {
        //
        // Fix up fNotAtBOL - if the cp we are at is between lines, we should
        // always be at the beginning of the line. One exception - if we are to the 
        // left of a layout, we should render on the previous line. If we are to the
        // right of a layout, we should render on the next line.
        //

        IFC( _pEd->GetViewServices()->IsPointerBetweenLines( epWalker, &fBetweenLines ));

        if( fBetweenLines )
        {
            CEditPointer tPointer( _pEd );
            BOOL fAtNextLineFuzzy = FALSE;
            DWORD dwBreak = BREAK_CONDITION_OMIT_PHRASE - BREAK_CONDITION_Anchor;
            DWORD dwFound = BREAK_CONDITION_None;
            IFC( tPointer.MoveToPointer( epWalker ));
            IFC( tPointer.Scan( RIGHT, dwBreak, &dwFound ));

            if( fThereIsAnotherLine )
                IFC( tPointer.IsEqualTo( epNextLine, BREAK_CONDITION_Phrase | BREAK_CONDITION_Anchor, & fAtNextLineFuzzy ));

            if( ! CheckFlag( dwFound, BREAK_CONDITION_Site ) &&
                ! fAtNextLineFuzzy )
            {
                // No site to the right of me and I'm not right next to the next line, 
                // render at the bol.
                fNotAtBOL = FALSE;
                fAtLogicalBOL = TRUE;
            }
            else
            {
                // there was a site to my right - render at the end of the line
                fNotAtBOL = TRUE;
                fAtLogicalBOL = FALSE;
            }
        }
    }

Done:
    IFC( MoveToPointer( epWalker ));

Cleanup:
    if( ! FAILED( hr ))
    {
        if( pfNotAtBOL )
            *pfNotAtBOL = fNotAtBOL;

        if( pfAtLogicalBOL )
            *pfAtLogicalBOL = fAtLogicalBOL;
    }
    
    RRETURN( hr );

}


//+====================================================================================
//
// Method: IsEqualTo
//
// Synopsis: Am I in the same place as the passed in pointer if I ignore dwIgnoreBreaks?
//
//------------------------------------------------------------------------------------

HRESULT 
CEditPointer::IsEqualTo( 
    IMarkupPointer *        pPointer,
    DWORD                   dwIgnore,
    BOOL *                  pfEqual )
{
    HRESULT hr = S_OK;
    BOOL fEqual = FALSE;
    Assert( pPointer );
    Assert( pfEqual );
    Direction dwWhichWayToPointer = SAME;

    IFC( OldCompare( this, pPointer, &dwWhichWayToPointer ));

    if( dwWhichWayToPointer == SAME )
    {
        // quick out - same exact place
        fEqual = TRUE;
    }
    else
    {
        CEditPointer ep( _pEd );
        DWORD dwSearch = BREAK_CONDITION_ANYTHING;
        DWORD dwFound = BREAK_CONDITION_None;
        dwSearch = ClearFlag( dwSearch , dwIgnore );

        IFC( ep->MoveToPointer( this ));
        IFC( ep.SetBoundaryForDirection( dwWhichWayToPointer, pPointer ));
        IFC( ep.Scan( dwWhichWayToPointer, dwSearch, & dwFound ));
        fEqual = dwFound == BREAK_CONDITION_Boundary;
    }
    
Cleanup:
    if( pfEqual )
        *pfEqual = fEqual;
        
    RRETURN( hr );
}


HRESULT
CEditPointer::IsLeftOfOrEqualTo(
    IMarkupPointer *        pPointer,
    DWORD                   dwIgnore,
    BOOL *                  pfEqual )
{
    BOOL fEqual;
    HRESULT hr = S_OK;

    IFC( this->IsLeftOfOrEqualTo( pPointer, &fEqual ));

    if( ! fEqual )
    {
        CEditPointer ep( _pEd );
        DWORD dwSearch = BREAK_CONDITION_ANYTHING;
        DWORD dwFound = BREAK_CONDITION_None;

        dwSearch = ClearFlag( dwSearch , dwIgnore );
        IFC( ep->MoveToPointer( this ));
        IFC( ep.SetBoundaryForDirection( LEFT, pPointer ));
        IFC( ep.Scan( LEFT, dwSearch, & dwFound ));
        fEqual = dwFound == BREAK_CONDITION_Boundary;
    }
    
Cleanup:
    if( pfEqual )
        *pfEqual = fEqual;

    RRETURN( hr );
}



HRESULT
CEditPointer::IsRightOfOrEqualTo(
    IMarkupPointer *        pPointer,
    DWORD                   dwIgnore,
    BOOL *                  pfEqual )
{
    BOOL fEqual;
    HRESULT hr = S_OK;

    IFC( this->IsRightOfOrEqualTo( pPointer, &fEqual ));

    if( ! fEqual )
    {
        CEditPointer ep( _pEd );
        DWORD dwSearch = BREAK_CONDITION_ANYTHING;
        DWORD dwFound = BREAK_CONDITION_None;
        
        dwSearch = ClearFlag( dwSearch , dwIgnore );
        IFC( ep->MoveToPointer( this ));
        IFC( ep.SetBoundaryForDirection( RIGHT, pPointer ));
        IFC( ep.Scan( RIGHT, dwSearch, & dwFound ));
        fEqual = dwFound == BREAK_CONDITION_Boundary;
    }
    
Cleanup:
    if( pfEqual )
        *pfEqual = fEqual;

    RRETURN( hr );
}



//+====================================================================================
//
// Method: Between
//
// Synopsis: Am I in - between the 2 given pointers ?
//
//------------------------------------------------------------------------------------

BOOL
CEditPointer::Between( 
    IMarkupPointer* pStart, 
    IMarkupPointer * pEnd )
{
    BOOL fBetween = FALSE;
    HRESULT hr;

     IFC( IsRightOfOrEqualTo( pStart, & fBetween ));
     if ( fBetween )
     {
        IFC( IsLeftOfOrEqualTo( pEnd, & fBetween ));    // CTL_E_INCOMPATIBLE will bail - but this is ok              
     }
        
Cleanup:
    return fBetween;
}

//+====================================================================================
//
// Method: IsInDifferentEditableSite
//
// Synopsis: Are we in a different site-selectable object from the edit context
//
//------------------------------------------------------------------------------------


BOOL
CEditPointer::IsInDifferentEditableSite()
{
    return ( _pEd->IsInDifferentEditableSite( _pPointer ));
}


//
// General markup services helpers
//

HRESULT
EdUtil::CopyMarkupPointer(IMarkupServices *pMarkupServices,
                          IMarkupPointer  *pSource,
                          IMarkupPointer  **ppDest )
{
    HRESULT hr;

    hr = pMarkupServices->CreateMarkupPointer(ppDest);
    if (FAILED(hr))
        goto Cleanup;

    hr = (*ppDest)->MoveToPointer(pSource);

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//  Method:     CSegmentListIter::CSegmentListIter
//  Synopsis:   ctor
//-----------------------------------------------------------------------------

CSegmentListIter::CSegmentListIter()
{
    _iSegmentCount   = 0;
    _iCurrentSegment = 0;
    _pLeft = _pRight = NULL;
    _pSegmentList    = NULL;
    _pViewServices   = NULL;
}

//+----------------------------------------------------------------------------
//  Method:     CSegmentListIter::CSegmentListIter
//  Synopsis:   dtor
//-----------------------------------------------------------------------------

CSegmentListIter::~CSegmentListIter()
{
    ReleaseInterface(_pLeft);
    ReleaseInterface(_pRight);
    ReleaseInterface(_pSegmentList);
    ReleaseInterface(_pViewServices);
}

//+----------------------------------------------------------------------------
//  Method:     CSegmentListIter::Init
//  Synopsis:   init method
//-----------------------------------------------------------------------------
HRESULT CSegmentListIter::Init(IMarkupServices *pMarkupServices, IHTMLViewServices *pViewServices, ISegmentList *pSegmentList)
{
    HRESULT hr;

    //
    // Set up pointers
    //
    ReleaseInterface(_pLeft);
    ReleaseInterface(_pRight);
    ReleaseInterface(_pSegmentList);
    ReplaceInterface(&_pViewServices, pViewServices);

    hr = pSegmentList->GetSegmentCount(&_iSegmentCount, NULL);
    if (FAILED(hr))
        goto Cleanup;

    hr = pMarkupServices->CreateMarkupPointer(&_pLeft);
    if (FAILED(hr))
        goto Cleanup;

    hr = pMarkupServices->CreateMarkupPointer(&_pRight);
    if (FAILED(hr))
        goto Cleanup;

    // Cache segment list
    _pSegmentList = pSegmentList;
    _pSegmentList->AddRef();

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//  Method:     CSegmentListIter::Next
//  Synopsis:   Move pointers to next segment.
//              Returns S_FALSE if last segment
//-----------------------------------------------------------------------------

HRESULT CSegmentListIter::Next(IMarkupPointer **ppLeft, IMarkupPointer **ppRight)
{
    HRESULT hr;

    //
    // Advance to next segment
    //

    if (_iCurrentSegment < _iSegmentCount)
    {
        hr = MovePointersToSegmentHelper(_pViewServices, _pSegmentList, _iCurrentSegment, &_pLeft, &_pRight);
        *ppLeft = _pLeft;
        *ppRight = _pRight;

        _iCurrentSegment++;
    }
    else
    {
        hr = S_FALSE;
    }

    RRETURN1(hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//  Method:     VariantCompareColor, local helper
//  Synopsis:   compares color
//-----------------------------------------------------------------------------

BOOL
EdUtil::VariantCompareColor(VARIANT * pvarColor1, VARIANT * pvarColor2)
{
    BOOL        fResult;
    CVariant    var;
    COLORREF    color1;
    COLORREF    color2;

    if (   V_VT(pvarColor1) == VT_NULL
        || V_VT(pvarColor2) == VT_NULL
       )
    {
        fResult = V_VT(pvarColor1) == V_VT(pvarColor2);
        goto Cleanup;
    }

    if (VariantChangeTypeSpecial(&var, pvarColor1,  VT_I4))
        goto Error;

    color1 = (COLORREF)V_I4(&var);

    if (VariantChangeTypeSpecial(&var, pvarColor2, VT_I4))
        goto Error;

    color2 = (COLORREF)V_I4(&var);

    fResult = color1 == color2;

Cleanup:
    return fResult;

Error:
    fResult = FALSE;
    goto Cleanup;
}

//+----------------------------------------------------------------------------
//  Method:     VariantCompareBSTRS, local helper
//  Synopsis:   compares 2 btrs
//-----------------------------------------------------------------------------

BOOL
EdUtil::VariantCompareBSTRS(VARIANT * pvar1, VARIANT * pvar2)
{
    BOOL    fResult;
    TCHAR  *pStr1;
    TCHAR  *pStr2;
    TCHAR  ach[1] = {0};

    if (V_VT(pvar1) == VT_BSTR && V_VT(pvar2) == VT_BSTR)
    {
        pStr1 = V_BSTR(pvar1) ? V_BSTR(pvar1) : ach;
        pStr2 = V_BSTR(pvar2) ? V_BSTR(pvar2) : ach;

        fResult = StrCmpC(pStr1, pStr2) == 0;
    }
    else
    {
        fResult = FALSE;
    }

    return fResult;
}

//+---------------------------------------------------------------------------
//
//  Function:   FormsAllocString
//
//  Synopsis:   Allocs a BSTR and initializes it from a string.  If the
//              initializer is NULL or the empty string, the resulting bstr is
//              NULL.
//
//  Arguments:  [pch]   -- String to initialize BSTR.
//              [pBSTR] -- The result.
//
//  Returns:    HRESULT.
//
//  Modifies:   [pBSTR]
//
//  History:    5-06-94   adams   Created
//
//----------------------------------------------------------------------------

HRESULT
EdUtil::FormsAllocStringW(LPCWSTR pch, BSTR * pBSTR)
{
    HRESULT hr;

    Assert(pBSTR);
    if (!pch || !*pch)
    {
        *pBSTR = NULL;
        return S_OK;
    }
    *pBSTR = SysAllocString(pch);
    hr = *pBSTR ? S_OK : E_OUTOFMEMORY;
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//  Method:     VariantCompareFontName, local helper
//  Synopsis:   compares font names
//-----------------------------------------------------------------------------

BOOL
EdUtil::VariantCompareFontName(VARIANT * pvarName1, VARIANT * pvarName2)
{
    return VariantCompareBSTRS(pvarName1, pvarName2);
}

//+----------------------------------------------------------------------------
//  Method:     VariantCompareFontSize, local helper
//  Synopsis:   compares font size
//-----------------------------------------------------------------------------

BOOL
EdUtil::VariantCompareFontSize(VARIANT * pvarSize1, VARIANT * pvarSize2)
{
    CVariant    convVar1;
    CVariant    convVar2;
    BOOL        fResult;

    Assert(pvarSize1);
    Assert(pvarSize2);

    if (   V_VT(pvarSize1) == VT_NULL
        || V_VT(pvarSize2) == VT_NULL
       )
    {
        fResult = V_VT(pvarSize1) == V_VT(pvarSize2);
        goto Cleanup;
    }

    if (VariantChangeTypeSpecial(&convVar1, pvarSize1, VT_I4))
        goto Error;

    if (VariantChangeTypeSpecial(&convVar2, pvarSize2, VT_I4))
        goto Error;

    fResult = V_I4(&convVar1) == V_I4(&convVar2);

Cleanup:
    return fResult;

Error:
    fResult = FALSE;
    goto Cleanup;
}

//+----------------------------------------------------------------------------
//  Method:     CBreakContainer::Add
//  Synopsis:   Add an element to the break container
//-----------------------------------------------------------------------------
VOID CBreakContainer::Set(ELEMENT_TAG_ID tagId, Mask mask)
{
    if (mask & BreakOnStart)
        bitFieldStart.Set(tagId);
    else
        bitFieldStart.Clear(tagId);

    if (mask & BreakOnEnd)
        bitFieldEnd.Set(tagId);
    else
        bitFieldEnd.Clear(tagId);
}

//+----------------------------------------------------------------------------
//  Method:     CBreakContainer::Test
//  Synopsis:   Tests an element in the break container
//-----------------------------------------------------------------------------
VOID CBreakContainer::Clear(ELEMENT_TAG_ID tagId, Mask mask)
{
    if (mask & BreakOnStart)
        bitFieldStart.Clear(tagId);

    if (mask & BreakOnEnd)
        bitFieldEnd.Clear(tagId);
}

//+----------------------------------------------------------------------------
//  Method:     CBreakContainer::Clear
//  Synopsis:   Clears an element in the break container
//-----------------------------------------------------------------------------
BOOL CBreakContainer::Test(ELEMENT_TAG_ID tagId, Mask mask)
{
    BOOL bResult = FALSE;

    switch (mask)
    {
    case BreakOnStart:
        bResult = bitFieldStart.Test(tagId);
        break;

    case BreakOnEnd:
        bResult = bitFieldEnd.Test(tagId);
        break;

    case BreakOnBoth:
        bResult = bitFieldStart.Test(tagId) && bitFieldEnd.Test(tagId) ;
        break;
    }

    return bResult;
}

HRESULT
EdUtil::ConvertOLEColorToRGB(VARIANT *pvarargIn)
{
    HRESULT hr = S_OK;
    DWORD   dwColor;

    if (V_VT(pvarargIn) != VT_BSTR)
    {
        hr = VariantChangeTypeSpecial(pvarargIn, pvarargIn, VT_I4);

        if (!hr && V_VT(pvarargIn) == VT_I4)
        {
            //
            // Note SujalP and TerryLu:
            //
            // If the color coming in is not a string type, then it is assumed
            // to be in a numeric format which is BBGGRR. However, FormatRange
            // (actually put_color()) expects a RRGGBB. For this reason, whenever
            // we get a VT_I4, we convert it to a RRGGBB. To do this we use
            // then helper on CColorValue, SetFromRGB() wto which we pass an
            // ****BBGGRR**** value. It just flips the bytes around and ends
            // up with a RRGGBB value, which we retrieve from GetColorRef().
            //
            V_VT(pvarargIn) = VT_I4;
            dwColor = V_I4(pvarargIn);

            V_I4(pvarargIn) = ((dwColor & 0xff) << 16)
                             | (dwColor & 0xff00)
                             | ((dwColor & 0xff0000) >> 16);

        }
    }

    RRETURN(hr);
}

HRESULT
EdUtil::ConvertRGBToOLEColor(VARIANT *pvarargIn)
{
    //
    // It just flips the byte order so this is a valid implementation
    //
    return ConvertOLEColorToRGB(pvarargIn);
}

BOOL
EdUtil::IsListContainer(ELEMENT_TAG_ID tagId)
{
    switch (tagId)
    {
    case TAGID_OL:
    case TAGID_UL:
    case TAGID_MENU:
    case TAGID_DIR:
    case TAGID_DL:
        return TRUE;

    default:
        return FALSE;
    }
}

BOOL
EdUtil::IsListItem(ELEMENT_TAG_ID tagId)
{
    switch (tagId)
    {
    case TAGID_LI:
    case TAGID_DD:
    case TAGID_DT:
        return TRUE;

    default:
        return FALSE;
    }
}

// Font height conversion data.  Valid HTML font sizes ares [1..7]
// NB (cthrash) These are in twips, and are in the 'smallest' font
// size.  The scaling takes place in CFontCache::GetCcs().

// BUGBUG (cthrash) We will need to get these values from the registry
// when we internationalize this product, so as to get sizing appropriate
// for the target locale.

//BUGBUG johnv: Where did these old numbers come from?  The new ones now correspond to
//TwipsFromHtmlSize[2] defined above.
//static const int aiSizesInTwips[7] = { 100, 130, 160, 200, 240, 320, 480 };

static const int aiSizesInTwips[7] = { 151, 200, 240, 271, 360, 480, 720 };

int
EdUtil::ConvertHtmlSizeToTwips(int nHtmlSize)
{
    // If the size is out of our conversion range do correction
    // Valid HTML font sizes ares [1..7]
    nHtmlSize = max( 1, min( 7, nHtmlSize ) );

    return aiSizesInTwips[ nHtmlSize - 1 ];
}

int EdUtil::ConvertTwipsToHtmlSize(int nFontSize)
{
    int nNumElem = ARRAYSIZE(aiSizesInTwips);

    // Now convert the point size to size units used by HTML
    // Valid HTML font sizes ares [1..7]
    int i;
    for(i=0; i<nNumElem; i++)
    {
        if(nFontSize <= aiSizesInTwips[i])
        {
            break;
        }
    }

    return i+1;
}

#define LF          10
#define CR          13
#define FF          TEXT('\f')
#define TAB         TEXT('\t')
#define VT          TEXT('\v')
#define PS          0x2029

BOOL
EdUtil::IsWhiteSpace(TCHAR ch)
{
    return (    ch == L' '
             || InRange( ch, TAB, CR ));
}

static HRESULT
EnsureValidScope(IHTMLViewServices *pViewServices, IMarkupPointer *pCurrent)
{
    HRESULT       hr;
    IHTMLElement  *pElement = NULL;

    hr = pViewServices->CurrentScopeOrSlave(pCurrent, &pElement);
    if (FAILED(hr))
        goto Cleanup;

    if (!pElement)
    {
        hr = pViewServices->RightOrSlave(pCurrent, TRUE, NULL, &pElement, NULL, NULL);

        if (FAILED(hr) || !pElement)
        {
            hr = pViewServices->LeftOrSlave(pCurrent, TRUE, NULL, &pElement, NULL, NULL);
            if (FAILED(hr))
                goto Cleanup;

            Assert(pElement);
        }

    }

Cleanup:
    ReleaseInterface(pElement);
    RRETURN(hr);
}


//+====================================================================================
//
// Method: PointersInSameFlowLayout
//
// Synopsis: Helper to check if pointers are in the same flow layout.
//
//------------------------------------------------------------------------------------



BOOL
EdUtil::PointersInSameFlowLayout(IMarkupPointer * pStart, 
                                         IMarkupPointer * pEnd, 
                                         IHTMLElement ** ppFlowElement,
                                         IHTMLViewServices* pViewServices )
{
    BOOL                fInSameFlow = FALSE;
    IHTMLElement    *   pElementFlowStart = NULL;
    IHTMLElement    *   pElementFlowEnd = NULL;
    IObjectIdentity *   pIdentity = NULL;
    HRESULT             hr;

    hr = pViewServices->GetFlowElement(pStart, & pElementFlowStart);
    if ( hr || (! pElementFlowStart) )
        goto Cleanup;

    hr = pViewServices->GetFlowElement(pEnd, & pElementFlowEnd);
    if ( hr || (! pElementFlowEnd) )
        goto Cleanup;

    hr = pElementFlowStart->QueryInterface( IID_IObjectIdentity, (void **) &pIdentity );
    if (hr)
        goto Cleanup;

    fInSameFlow = ( pIdentity->IsEqualObject( pElementFlowEnd ) == S_OK );

    if (ppFlowElement)
    {
        if (fInSameFlow)
        {
            *ppFlowElement = pElementFlowStart;
            pElementFlowStart->AddRef();
        }
        else
        {
            *ppFlowElement = NULL;
        }
    }
        
Cleanup:
    ReleaseInterface(pElementFlowStart);
    ReleaseInterface(pElementFlowEnd);
    ReleaseInterface(pIdentity);

    return fInSameFlow;
}

HRESULT
EdUtil::MovePointersToSegmentHelper(
    IHTMLViewServices* pViewServices,
    ISegmentList*    pSegmentList,
    INT              iSegmentIndex,
    IMarkupPointer** ppStart,
    IMarkupPointer** ppEnd,
    BOOL             fCheckPtrOrder, /* = TRUE */
    BOOL             fCheckCurScope  /* = TRUE */)
{
    HRESULT        hr;
    INT            iPosition;
    IMarkupPointer *pTemp;

    hr = pSegmentList->MovePointersToSegment(iSegmentIndex, *ppStart, *ppEnd);
    if (FAILED(hr))
        goto Cleanup;

    //
    // Make sure both pointers have a current scope
    //

    if (fCheckCurScope)
    {
        hr = EnsureValidScope(pViewServices, *ppStart);
        if (FAILED(hr))
            goto Cleanup;

        hr = EnsureValidScope(pViewServices, *ppEnd);
        if (FAILED(hr))
            goto Cleanup;
    }

    //
    // Check the left/right orientation
    //

    if (fCheckPtrOrder)
    {
        OldCompare( (*ppStart), *ppEnd, &iPosition);
        if (iPosition == LEFT)
        {
            // swap pointers
            pTemp = *ppStart;
            *ppStart = *ppEnd;
            *ppEnd = pTemp;
        }
    }

Cleanup:
    RRETURN(hr);
}


HRESULT EdUtil::FindCommonElement( IMarkupServices *pMarkupServices,
                                   IHTMLViewServices *pViewServices,
                                     IMarkupPointer  *pStart,
                                     IMarkupPointer  *pEnd,
                                     IHTMLElement    **ppElement )
{
    HRESULT         hr;
    IMarkupPointer  *pLeft;
    IMarkupPointer  *pRight;
    INT             iPosition;
    IMarkupPointer  *pCurrent = NULL;
    IHTMLElement    *pOldElement = NULL;

    // init for error case
    *ppElement = NULL;

    //
    // Find right/left pointers
    //

    hr = OldCompare( pStart, pEnd, &iPosition);
    if (FAILED(hr))
        goto Cleanup;

    if (iPosition == SAME)
    {
        hr = pViewServices->CurrentScopeOrSlave(pStart, ppElement);
        goto Cleanup;
    }
    else if (iPosition == RIGHT)
    {
        pLeft = pStart;     // weak ref
        pRight = pEnd;      // weak ref
    }
    else
    {
        pLeft = pEnd;       // weak ref
        pRight = pStart;    // weak ref
    }

    //
    // Walk the left pointer up until the right end of the element
    // is to the right of pRight
    //

    hr = pViewServices->CurrentScopeOrSlave(pLeft, ppElement);
    if (FAILED(hr))
        goto Cleanup;

    hr = pMarkupServices->CreateMarkupPointer(&pCurrent);
    if (FAILED(hr))
        goto Cleanup;

    while (*ppElement)
    {
        hr = pCurrent->MoveAdjacentToElement(*ppElement, ELEM_ADJ_AfterEnd);
        if (FAILED(hr))
            goto Cleanup;

        hr = OldCompare( pCurrent, pRight, &iPosition);
        if (FAILED(hr))
            goto Cleanup;

        if (iPosition != RIGHT)
            break; // found common element

        pOldElement = *ppElement;
        hr = pOldElement->get_parentElement(ppElement);
        pOldElement->Release();
        if (FAILED(hr))
            goto Cleanup;
        Assert(*ppElement); // we should never walk up past the root of the tree
    }

Cleanup:
    ReleaseInterface(pCurrent);
    RRETURN(hr);
}

HRESULT EdUtil::FindBlockElement( IMarkupServices  *pMarkupServices,
                                    IHTMLElement     *pElement,
                                    IHTMLElement     **ppBlockElement )
{
    HRESULT             hr = S_OK;
    IHTMLElement        *pOldElement = NULL;
    IHTMLViewServices   *pViewServices = NULL;
    BOOL                bBlockElement;

    *ppBlockElement = pElement;
    pElement->AddRef();

    hr = pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID *)&pViewServices);
    if (FAILED(hr))
        goto Cleanup;

    do
    {
        hr = pViewServices->IsBlockElement(*ppBlockElement, &bBlockElement);
        if (FAILED(hr) || bBlockElement)
            goto Cleanup;

        hr = pViewServices->IsLayoutElement(*ppBlockElement, &bBlockElement);
        if (FAILED(hr) || bBlockElement)
            goto Cleanup;

        pOldElement = *ppBlockElement;
        hr = pOldElement->get_parentElement(ppBlockElement);
        pOldElement->Release();
        if (FAILED(hr))
            goto Cleanup;
    }
    while (*ppBlockElement);

Cleanup:
    ReleaseInterface(pViewServices);
    RRETURN(hr);
}


//+---------------------------------------------------------------------------
//
//  Method:       CopyAttributes.
//
//  Synopsis:     Wrapper to IHTMLDocument2::mergeAttributes
//
//----------------------------------------------------------------------------
HRESULT
EdUtil::CopyAttributes( IHTMLViewServices *pViewSrv, IHTMLElement * pSrcElement, IHTMLElement * pDestElement, BOOL fCopyId)
{
    HRESULT hr = E_POINTER;

    if (!pViewSrv)
        goto Cleanup;

    AssertSz(FALSE, "must improve");

Cleanup:
    RRETURN( hr );
}

HRESULT EdUtil::ReplaceElement( IMarkupServices *pMarkupServices,
                                IHTMLElement    *pOldElement,
                                IHTMLElement    *pNewElement,
                                IMarkupPointer  *pUserStart,
                                IMarkupPointer  *pUserEnd)
{
    HRESULT        hr;
    IMarkupPointer *pStart = NULL;
    IMarkupPointer *pEnd = NULL;
    IHTMLViewServices *pViewSrv = NULL;
    //
    // Set up markup pointers
    //

    if (pUserStart)
    {
        pStart = pUserStart;
        pStart->AddRef();
    }
    else
    {
        hr = pMarkupServices->CreateMarkupPointer(&pStart);
        if (FAILED(hr))
            goto Cleanup;
    }

    if (pUserEnd)
    {
        pEnd = pUserEnd;
        pEnd->AddRef();
    }
    else
    {
        hr = pMarkupServices->CreateMarkupPointer(&pEnd);
        if (FAILED(hr))
            goto Cleanup;
    }

    //
    // Replace the element
    //

    hr = pEnd->MoveAdjacentToElement(pOldElement, ELEM_ADJ_AfterEnd);
    if (FAILED(hr))
        goto Cleanup;

    hr = pStart->MoveAdjacentToElement(pOldElement, ELEM_ADJ_BeforeBegin);
    if (FAILED(hr))
        goto Cleanup;
    
    hr = pMarkupServices->QueryInterface(IID_IHTMLViewServices, (void**)&pViewSrv);
    if (FAILED(hr))
        goto Cleanup;

    hr = CopyAttributes(pViewSrv, pOldElement, pNewElement, TRUE);
    if (FAILED(hr))
        goto Cleanup;
        
    hr = InsertElement(pMarkupServices, pNewElement, pStart, pEnd);
    if (FAILED(hr))
        goto Cleanup;

    hr = pMarkupServices->RemoveElement(pOldElement);
    if (FAILED(hr))
        goto Cleanup;

Cleanup:
    ReleaseInterface(pViewSrv);
    ReleaseInterface(pStart);
    ReleaseInterface(pEnd);
    RRETURN(hr);
}
//+====================================================================================
//
// Method: IsElementPositioned
//
// Synopsis: Does this element have a Relative/Absolute Position.
//
//------------------------------------------------------------------------------------

BOOL
EdUtil::IsElementPositioned(IHTMLElement* pElement)
{
    HRESULT hr = S_OK;

    IHTMLStyle * pStyle = NULL;
    BSTR bsPosition = NULL;
    
    BOOL fIsPosition = FALSE;

    if( pElement == NULL )
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    hr = pElement->get_style( & pStyle );
    if ( hr || pStyle == NULL )
        goto Cleanup;

    hr = pStyle->get_position( & bsPosition);
    if ( hr ) 
        goto Cleanup;

    if ( StrCmpIW(_T("absolute"), bsPosition ) == 0)
    {
        fIsPosition = TRUE;
    }
    else if ( StrCmpIW(_T("relative"), bsPosition ) == 0)
    {
        fIsPosition = TRUE;
    }

    
Cleanup:
    SysFreeString( bsPosition );
    ReleaseInterface( pStyle );
    
    return ( fIsPosition );
}

//+====================================================================================
//
// Method: IsElementSized
//
// Synopsis: Does this element have a width/height style.
//
//------------------------------------------------------------------------------------


BOOL
EdUtil::IsElementSized(IHTMLElement* pElement)
{

    HRESULT hr = S_OK;
    IHTMLStyle * pCurStyle = NULL;
    VARIANT varSize;
    int width = 0;
    int height = 0;
    VariantInit( & varSize );
    
    BOOL fIsSized = FALSE;


    IFC( pElement->get_style( & pCurStyle ));

    IFC( pCurStyle->get_width( & varSize ));
    width = varSize.lVal;
    
    IFC( pCurStyle->get_height( & varSize ));    
    height = varSize.lVal;
    
    if (( width != 0 ) || ( height != 0 ))
        fIsSized = TRUE;
    
Cleanup:

    ReleaseInterface( pCurStyle );
    
    return ( fIsSized );
}

BOOL 
EdUtil::IsShiftKeyDown()
{
    return (GetKeyState(VK_SHIFT) & 0x8000) ;
}

BOOL 
EdUtil::IsControlKeyDown()
{
    return (GetKeyState(VK_CONTROL) & 0x8000) ;
}

BOOL 
EdUtil::IsArrowKey( SelectionMessage* pMessage)
{
    return (( pMessage->wParam >= VK_LEFT) && ( pMessage->wParam <= VK_DOWN ) );
}
  

HRESULT EdUtil::MoveAdjacentToElementHelper(IMarkupPointer *pMarkupPointer, IHTMLElement *pElement, ELEMENT_ADJACENCY elemAdj)
{
    HRESULT hr;
    
    hr = pMarkupPointer->MoveAdjacentToElement(pElement, elemAdj);
    if (hr)
    {
        if (elemAdj == ELEM_ADJ_AfterBegin)
            hr = pMarkupPointer->MoveAdjacentToElement(pElement, ELEM_ADJ_BeforeBegin);
        else if (elemAdj == ELEM_ADJ_BeforeEnd)
            hr = pMarkupPointer->MoveAdjacentToElement(pElement, ELEM_ADJ_AfterEnd);
    }
    RRETURN(hr);
}


HRESULT
EdUtil::FindBlockLimit(
    IMarkupServices*    pMarkupServices, 
    IHTMLViewServices*  pViewServices,
    Direction           direction, 
    IMarkupPointer      *pPointer, 
    IHTMLElement        **ppElement, 
    MARKUP_CONTEXT_TYPE *pContext,
    BOOL                fExpanded,
    BOOL                fLeaveInside,
    BOOL                fCanCrossLayout)
{
    HRESULT             hr;
    MARKUP_CONTEXT_TYPE context;
    SP_IHTMLElement     spElement;
    ELEMENT_TAG_ID      tagId;
    BOOL                fFoundExitScope = FALSE;
    BOOL                bLayout = FALSE;

    if (ppElement)
        *ppElement = NULL; 

    // Find the block
    for (;;)
    {
        // Move to next scope change (note we only get enter/exit scope for blocks)
        IFR( BlockMove( pMarkupServices, pViewServices, pPointer, direction, TRUE, &context, &spElement) );

        switch (context)
        {
            case CONTEXT_TYPE_ExitScope:
                if (!spElement || !fExpanded)
                    goto FoundBlockLimit; // went too far
                
                // Make sure we didn't exit the body
                IFR( pMarkupServices->GetElementTagId(spElement, &tagId) );
                if (tagId == TAGID_BODY)
                {
                    fLeaveInside = FALSE; // force adjustment
                    goto FoundBlockLimit;
                }

                // Check for flow layout
                IFR( pViewServices->IsLayoutElement(spElement, &bLayout) );
                if (bLayout)
                    goto FoundBlockLimit;
                
                fFoundExitScope = TRUE;                
                break;

            case CONTEXT_TYPE_Text:
            case CONTEXT_TYPE_NoScope:
                if (fFoundExitScope)
                    goto FoundBlockLimit;
                break;

            case CONTEXT_TYPE_EnterScope:
                Assert(IsBlockCommandLimit( pMarkupServices, pViewServices, spElement, context)); 
                IFR( pViewServices->IsLayoutElement(spElement, &bLayout));
                goto FoundBlockLimit;
        }
    }
    
FoundBlockLimit:
    
    if (!fLeaveInside || (bLayout && !fCanCrossLayout))
        IFR( BlockMoveBack( pMarkupServices, pViewServices, pPointer, direction, TRUE, pContext, ppElement) );    

    RRETURN(hr);
}

BOOL 
EdUtil::IsBlockCommandLimit( IMarkupServices* pMarkupServices, 
                             IHTMLViewServices* pViewServices, 
                             IHTMLElement *pElement, 
                             MARKUP_CONTEXT_TYPE context) 
{
    HRESULT         hr;
    ELEMENT_TAG_ID  tagId;
    BOOL            bResult = FALSE;

    switch (context)
    {
        case CONTEXT_TYPE_NoScope:
        case CONTEXT_TYPE_Text:
            return FALSE;
    }

    //
    // Check exceptions
    //
    
    IFR( pMarkupServices->GetElementTagId(pElement, &tagId) );
    switch (tagId)     
    {
        case TAGID_BUTTON:
        case TAGID_COL:
        case TAGID_COLGROUP:
        case TAGID_TBODY:
        case TAGID_TFOOT:
        case TAGID_TH:
        case TAGID_THEAD:
        case TAGID_TR:
            return FALSE;            
    }

    //
    // Otherwise, return IsBlockElement || IsLayoutElement
    //
    
    IFR( pViewServices->IsBlockElement(pElement, &bResult) );
    
    if (!bResult)
        IFR( pViewServices->IsLayoutElement(pElement, &bResult) );
    
    return bResult;    
}

HRESULT 
EdUtil::BlockMove(
    IMarkupServices         * pMarkupServices,
    IHTMLViewServices       * pViewServices,
    IMarkupPointer          *pMarkupPointer, 
    Direction               direction, 
    BOOL                    fMove,
    MARKUP_CONTEXT_TYPE *   pContext,
    IHTMLElement * *        ppElement)
{
    HRESULT                 hr;
    MARKUP_CONTEXT_TYPE     context;
    SP_IHTMLElement         spElement;
    SP_IMarkupPointer       spPointer;

    Assert(direction == LEFT || direction == RIGHT);

    if (!fMove)
    {
        IFR( pMarkupServices->CreateMarkupPointer(&spPointer) );
        IFR( spPointer->MoveToPointer(pMarkupPointer) );
        
        pMarkupPointer = spPointer; // weak ref 
    }
    
    for (;;)
    {
        if (direction == LEFT)
            IFC( pViewServices->LeftOrSlave(pMarkupPointer, TRUE, &context, &spElement, NULL, NULL) )
        else
            IFC( pViewServices->RightOrSlave(pMarkupPointer, TRUE, &context, &spElement, NULL, NULL) );

        switch (context)
        {
            case CONTEXT_TYPE_Text:
                goto Cleanup; // done

            case CONTEXT_TYPE_EnterScope:
                if (IsIntrinsic(pMarkupServices, spElement))
                {
                    if (direction == LEFT)
                        IFC( pMarkupPointer->MoveAdjacentToElement( spElement, ELEM_ADJ_BeforeBegin ) )
                    else
                        IFC( pMarkupPointer->MoveAdjacentToElement( spElement, ELEM_ADJ_AfterEnd ) ); 

                    continue;    
                }
                // fall through
                
            case CONTEXT_TYPE_ExitScope:
            case CONTEXT_TYPE_NoScope:
                        if (!spElement || IsBlockCommandLimit( pMarkupServices, pViewServices, spElement, context))
                    goto Cleanup; // done;
                break;  

            default:
                hr = E_FAIL; // CONTEXT_TYPE_None
                goto Cleanup;
        }
    }
    
Cleanup:
    if (ppElement)
    {
        if (SUCCEEDED(hr))
        {
            *ppElement = spElement;
            if (*ppElement)
                (*ppElement)->AddRef();
        }
        else
        {
            *ppElement = NULL;
        }
    }
        
    if (pContext)
    {
        *pContext = (SUCCEEDED(hr)) ? context : CONTEXT_TYPE_None;
    }    

    RRETURN(hr);
}

HRESULT 
EdUtil::BlockMoveBack(
    IMarkupServices*        pMarkupServices,
    IHTMLViewServices           * pViewServices,    
    IMarkupPointer          *pMarkupPointer, 
    Direction               direction, 
    BOOL                    fMove,
    MARKUP_CONTEXT_TYPE *   pContext,
    IHTMLElement * *        ppElement)
{
    if (direction == RIGHT)
    {
        RRETURN(BlockMove( pMarkupServices, pViewServices, pMarkupPointer, LEFT, fMove, pContext, ppElement));
    }
    else
    {
        Assert(direction == LEFT);
        RRETURN(BlockMove(pMarkupServices, pViewServices, pMarkupPointer, RIGHT, fMove, pContext, ppElement));
    }
}

BOOL 
EdUtil::DoesSegmentContainText( IMarkupServices * pMarkupServices, IHTMLViewServices *pViewServices, IMarkupPointer *pStart, IMarkupPointer *pEnd)
{
    HRESULT                 hr;
    SP_IMarkupPointer       spCurrent;
    BOOL                    bEqual;
    MARKUP_CONTEXT_TYPE     context;
    TCHAR                   ch;
    LONG                    cch;

    IFC(CopyMarkupPointer( pMarkupServices, pStart, &spCurrent));
    
    do
    {
        IFC( spCurrent->IsEqualTo(pEnd, &bEqual) );
        if (bEqual)
            return FALSE; // no text 

        cch = 1;
        IFC( pViewServices->RightOrSlave(spCurrent, TRUE, &context, NULL, &cch, &ch) );

        if (context == CONTEXT_TYPE_Text)
        {
            if (ch == WCH_NBSP)
                continue;
                
            goto Cleanup;
        }
    } 
    while (context != CONTEXT_TYPE_NoScope);

Cleanup:
    return TRUE;
}


//+==============================================================================
//
//  Method:     MovePointerOutOfScope
//
//  Synopsis:   This routine moves the passed in markup pointer out of the 
//              influence of tags
//
//              ie - we keep moving the pointer on exit scopes, stop on enter scope.
//
//+==============================================================================

HRESULT
EdUtil::MovePointerOutOfScope( 
    IMarkupServices *   pMarkupServices,
    IHTMLViewServices * pViewServices,
    IMarkupPointer *    pOrigin, 
    Direction           eDir,
    IMarkupPointer *    pBoundary,
    BOOL * pfAdjustedPointer,
    BOOL fStopAtBlockOnExitScope /* = TRUE */,
    BOOL fStopAtBlockOnEnterScope, /*=TRUE*/
    BOOL fStopAtLayout, /*=TRUE*/
    BOOL fStopAtIntrinsic /* = TRUE*/) 
{
    HRESULT             hr = S_OK;
    MARKUP_CONTEXT_TYPE ct = CONTEXT_TYPE_None;
    BOOL                fDone = FALSE;
    BOOL fBlock = FALSE;
    BOOL fFoundBlock = FALSE;
    BOOL fLayout = FALSE;
    BOOL fIntrinsic = FALSE;
    BOOL fSameTree = FALSE;
    IMarkupContainer * pIContainer1 = NULL;
    IMarkupContainer * pIContainer2 = NULL;
    IUnknown* pUnk1 = NULL;
    IUnknown* pUnk2 = NULL;
    
    // Interface Pointers
    IHTMLElement      * pHTMLElement = NULL;
    IMarkupPointer    * pTemp = NULL;
    IMarkupPointer    * pBlock = NULL;
    
    //
    IFC( pOrigin->GetContainer( & pIContainer1 ));
    IFC( pBoundary->GetContainer( & pIContainer2 ));
    Assert( pIContainer1 && pIContainer2 );
    IFC( pIContainer1->QueryInterface( IID_IUnknown, (void**) & pUnk1 ));
    IFC( pIContainer2->QueryInterface( IID_IUnknown, (void**) & pUnk2 ));
    if ( pUnk1 == pUnk2 )
            fSameTree = TRUE;
        
    IFC( pMarkupServices->CreateMarkupPointer( & pTemp ) );
    IFC( pMarkupServices->CreateMarkupPointer( & pBlock ));
    
    IFC( pTemp->MoveToPointer( pOrigin ) );


    if ( fSameTree )
    {
        IFC( pTemp->IsEqualTo( pBoundary , & fDone ));
    }

    while ( ! fDone )
    {           
        ClearInterface(&pHTMLElement);
        if( eDir == LEFT )
            hr = pViewServices->LeftOrSlave(pTemp, TRUE, & ct, & pHTMLElement, NULL, NULL);
        else
            hr = pViewServices->RightOrSlave(pTemp, TRUE, & ct, & pHTMLElement, NULL, NULL);
        if( hr )
            goto Cleanup;

        if ( pHTMLElement )
        {
            IFC( pViewServices->IsBlockElement( pHTMLElement, & fBlock ));
            IFC( pViewServices->IsLayoutElement( pHTMLElement, & fLayout )); 
            fIntrinsic = IsIntrinsic( pMarkupServices, pHTMLElement);
        }

        if ( ( fIntrinsic && fStopAtIntrinsic ) ||
             ( fLayout && fStopAtLayout )  )
        {
            fDone = TRUE;
        }
        else
        {
            switch( ct )
            {
                case CONTEXT_TYPE_ExitScope:            
                    //
                    // we stop on exit scope of blocks iff fStopAtBlockOnExitScope is True
                    // but we will skip over other tags - so we will merrily skip over a </b>
                    //

                    if ( fBlock && fStopAtBlockOnExitScope )
                        fDone = TRUE;
                    else
                    {
                        fFoundBlock = TRUE;
                        IFC( pBlock->MoveToPointer( pTemp ));
                    }
                    break;        
                    
                case CONTEXT_TYPE_EnterScope:
                    //
                    // we stop on exit/enter scope of blocks iff fStopAtBlockOnEnterScope is True
                    // but we will skip over other tags - so we will merrily skip over a </b>
                    //

                    if ( fBlock && fStopAtBlockOnEnterScope )
                        fDone = TRUE;
                    else
                    {
                        fFoundBlock = TRUE;
                        IFC( pBlock->MoveToPointer( pTemp ));
                    }
                    break;  
                    
                case CONTEXT_TYPE_Text:
                case CONTEXT_TYPE_NoScope:                    
                    fDone = TRUE;
                    break;
                

                case CONTEXT_TYPE_None:
                    fDone = TRUE;
                    break;
            }
        }
        // Check to see if we hit the boundary of our search
        BOOL fAtBoundary = FALSE;
        if ( fSameTree )
        {
            IFC( pTemp->IsEqualTo( pBoundary , &fAtBoundary ));
        }
        fDone |= fAtBoundary;
        
        // Reset for next iteration
        ClearInterface( & pHTMLElement );
        fBlock = FALSE;
        fLayout = FALSE;
        fIntrinsic = FALSE;
    }
    
    // If we found text, move our pointer
    if( fFoundBlock )
    {
        hr = pOrigin->MoveToPointer( pBlock );
        if ( pfAdjustedPointer )
            *pfAdjustedPointer = TRUE;
    }        

Cleanup:
    ReleaseInterface( pUnk1 );
    ReleaseInterface( pUnk2 );
    ReleaseInterface( pIContainer1 );
    ReleaseInterface( pIContainer2 );
    ReleaseInterface( pHTMLElement );    
    ReleaseInterface( pTemp );
    ReleaseInterface( pBlock );
    
    RRETURN( hr );
}

BOOL
EdUtil::IsIntrinsic( IMarkupServices* pMarkupServices,
                     IHTMLElement* pIHTMLElement )
{                     
    HRESULT hr = S_OK;
    ELEMENT_TAG_ID eTag = TAGID_NULL;
    BOOL fIntrinsic = FALSE;
    
    IFC( pMarkupServices->GetElementTagId( pIHTMLElement, &eTag ));

    switch( eTag )
    {
        case TAGID_BUTTON:
        case TAGID_TEXTAREA:
        case TAGID_FIELDSET:
        case TAGID_LEGEND:
        case TAGID_MARQUEE:
        case TAGID_SELECT:
            fIntrinsic = TRUE;
            break;
    }
    
Cleanup:
    return fIntrinsic;
}   

//+==============================================================================
//
//  Method:     MovePointerToText
//
//  Synopsis:   This routine moves the passed in markup pointer until in the 
//              specified direction until it hits either text or a layout 
//              boundary.
//
//+==============================================================================

HRESULT
EdUtil::MovePointerToText( 
    IMarkupServices *   pMarkupServices,
    IHTMLViewServices * pViewServices,
    IMarkupPointer *    pOrigin, 
    Direction           eDir,
    IMarkupPointer *    pBoundary,
    BOOL *              pfHitText,
    BOOL fStopAtBlock ) 
{
    HRESULT             hr = S_OK;
    MARKUP_CONTEXT_TYPE ct;
    BOOL                fDone = FALSE;
    INT                 iResult = 0;
    ELEMENT_TAG_ID      eTag;
    LONG                cch = 1;
    
    // Interface Pointers
    IHTMLElement      * pHTMLElement = NULL;
    IMarkupPointer    * pTemp = NULL;

    *pfHitText = FALSE;

    hr = pMarkupServices->CreateMarkupPointer( & pTemp );
    if (! hr)   hr = pTemp->MoveToPointer( pOrigin );
    if (hr)
        goto Cleanup;
    
    // Rule #2 - walk in the appropriate direction looking for text or a noscope.
    // If we happen to enter the scope of another element, fine. If we try to leave
    // the scope of an element, bail.
    
    while ( ! fDone )
    {
        ClearInterface( & pHTMLElement );

        //
        // Check to see if we hit the boundary of our search
        //
        // If the pointer is beyond the boundary, fall out
        // boundary left of pointer =  -1
        // boundary equals pointer =    0
        // boundary right of pointer =  1
        //
        
        OldCompare( pBoundary, pTemp , &iResult);

        if((     eDir == LEFT && iResult > -1 ) 
            || ( eDir == RIGHT && iResult < 1 ))
        {
            goto Cleanup;    // this is okay since CR_Boundary does not move the pointer
        }
        
        //    
        // Move in the appropriate direction...
        //
        
        if( eDir == LEFT )
            hr = pViewServices->LeftOrSlave(pTemp, TRUE, & ct, & pHTMLElement, &cch, NULL);
        else
            hr = pViewServices->RightOrSlave(pTemp, TRUE, & ct, & pHTMLElement, &cch, NULL);
        if( hr )
            goto Cleanup;
            
        switch( ct )
        {
            case CONTEXT_TYPE_Text:

                //
                // Hit text - done
                //
                
                *pfHitText = TRUE;
                fDone = TRUE;
                break;
                
            case CONTEXT_TYPE_NoScope:

                //
                // Only stop if we hit a renderable (layout-ness) noscope
                // TODO : figure out if this is a glyph boundary
                //
                
                if( pHTMLElement )
                {
                    BOOL fHasLayout = FALSE;
                    BOOL fIsBlock = FALSE;
                    
                    IFC( pMarkupServices->GetElementTagId( pHTMLElement, &eTag ));
                    IFC( pViewServices->IsBlockElement( pHTMLElement , & fIsBlock ));
                    IFC( pViewServices->IsLayoutElement( pHTMLElement, & fHasLayout ));

                    if( fHasLayout || 
                        ( fIsBlock && fStopAtBlock ) || 
                        eTag == TAGID_BR )
                    {
                        *pfHitText = TRUE;
                        fDone = TRUE;
                    }
                }
                
                break;
            
            case CONTEXT_TYPE_EnterScope:
                if( pHTMLElement )
                {
                    BOOL fHasLayout = FALSE;

                    //
                    // Only stop for intrinsics, otherwise pass on through
                    // TODO: Figure out if this is a glyph boundary
                    //
                    //  Also, it seems wrong to not stop for layout boundaries
                    //
                    BOOL fIntrinsic  = IsIntrinsic( pMarkupServices, pHTMLElement);
                    if ( fIntrinsic )
                    {
                        *pfHitText = fDone;
                        fDone = TRUE;
                    }
                    else
                    {
                        IFC( pViewServices->IsLayoutElement( pHTMLElement, &fHasLayout ));
                        if (fHasLayout)
                        {
                            *pfHitText = fDone;
                            fDone = TRUE;
                        }
                    }

                }

                break;

            // TODO : Figure out if the range needs this or if we don't need this

            case CONTEXT_TYPE_ExitScope:
                if( pHTMLElement )
                {
                    BOOL fHasLayout = FALSE;
                    BOOL fIsBlock = FALSE;
                    
                    IFC( pMarkupServices->GetElementTagId( pHTMLElement, &eTag ));
                    IFC( pViewServices->IsBlockElement( pHTMLElement , & fIsBlock ));
                    IFC( pViewServices->IsLayoutElement( pHTMLElement, & fHasLayout ));

                    if( fHasLayout || 
                      ( fIsBlock && fStopAtBlock ) || 
                      eTag == TAGID_BR )
                    {
                        fDone = TRUE;
                    }
                }
                
                break;                 

            case CONTEXT_TYPE_None:
                fDone = TRUE;
                break;
        }        
    }

    //
    // If we found text, move our pointer
    //
    
    if( *pfHitText )
    {
        //
        // We have inevitably gone one move too far, back up one move
        //
    
        if( eDir == LEFT )
            hr = pViewServices->RightOrSlave(pTemp, TRUE, & ct, NULL, &cch, NULL);
        else
            hr = pViewServices->LeftOrSlave(pTemp, TRUE, & ct, NULL, &cch, NULL);
        if( hr )
            goto Cleanup;
            
        hr = pOrigin->MoveToPointer( pTemp );
    }
    
Cleanup:
    ReleaseInterface( pHTMLElement );
    ReleaseInterface( pTemp );
    RRETURN( hr );
}


//+---------------------------------------------------------------------------
//
//  Method:       EdUtil::ExpandToWord
//
//  Synopsis:     Expands empty selection to word if inside text.
//
//----------------------------------------------------------------------------

HRESULT
EdUtil::ExpandToWord(IMarkupServices * pMarkupServices, IHTMLViewServices *pViewServices, IMarkupPointer * pmpStart, IMarkupPointer * pmpEnd)
{
    SP_IMarkupPointer spmpPosition, spmpStart, spmpEnd;
    INT               iPosition;
    BOOL              fEqual, fExpand = FALSE;
    HRESULT           hr;

    Assert(pMarkupServices && pmpStart && pmpEnd);

    //
    // Markup pointers have to be at the same position in order for us to expand to a word.
    //

    hr = pmpStart->IsEqualTo(pmpEnd, &fEqual);
    if (hr || !fEqual)
        goto Cleanup;

    //
    // We have a collapsed selection (caret).  Now lets see
    // if we are inside a word of text.
    //

    IFC(CopyMarkupPointer(pMarkupServices, pmpStart, &spmpStart));
    IFC(CopyMarkupPointer(pMarkupServices, pmpEnd, &spmpEnd));
    IFC(CopyMarkupPointer(pMarkupServices, pmpStart, &spmpPosition));
    IFC(spmpEnd->MoveUnit(MOVEUNIT_NEXTWORDEND));
    IFC(spmpStart->MoveToPointer(spmpEnd));
    IFC(spmpStart->MoveUnit(MOVEUNIT_PREVWORDBEGIN));
    IFC(OldCompare( spmpStart, spmpPosition, &iPosition));

    if (iPosition != RIGHT)
        goto Cleanup;

    IFC(OldCompare( spmpPosition, spmpEnd, &iPosition));

    if (iPosition != RIGHT)
        goto Cleanup;

    {
        MARKUP_CONTEXT_TYPE mctContext;
        // BUGBUG: Due to a bug in MoveUnit(MOVEUNIT_NEXTWORDEND) we have to check
        // whether we ended up at the end of the markup.  Try removing this
        // once bug 37129 is fixed.
        IFC(pViewServices->RightOrSlave(spmpEnd, FALSE, &mctContext, NULL, NULL, NULL));
        if (mctContext == CONTEXT_TYPE_None)
            goto Cleanup;
    }

    fExpand = TRUE;

Cleanup:

    if (hr || !fExpand)
        hr = S_FALSE;
    else
    {
        pmpStart->MoveToPointer(spmpStart);
        pmpEnd->MoveToPointer(spmpEnd);
    }

    RRETURN1(hr, S_FALSE);
}


HRESULT 
EdUtil::FindListContainer( IMarkupServices *pMarkupServices,
                           IHTMLElement    *pElement,
                           IHTMLElement    **ppListContainer )
{
    HRESULT                         hr;
    SP_IHTMLElement                 spNewElement;
    SP_IHTMLElement                 spCurrentElement;
    ELEMENT_TAG_ID                  tagId;    
    CSmartPtr<IHTMLViewServices>    spViewServices;
    BOOL                            bLayout;

    *ppListContainer = NULL;

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID *)&spViewServices) );

    spCurrentElement = pElement;
    do
    {
        IFR( pMarkupServices->GetElementTagId(spCurrentElement, &tagId) ); 

        if (IsListContainer(tagId)) // found container
        {
            *ppListContainer = spCurrentElement;
            spCurrentElement->AddRef();
            break;
        }            

        IFR (spViewServices->IsLayoutElement(spCurrentElement, &bLayout) );
        if (bLayout)
            return S_OK; // done - don't cross layout
            
        IFR( spCurrentElement->get_parentElement(&spNewElement) );
        spCurrentElement = spNewElement;
    }
    while (spCurrentElement != NULL);

    RRETURN(hr);        
}

HRESULT 
EdUtil::FindContainer( IMarkupServices *pMarkupServices,
                       IHTMLElement    *pElement,
                       IHTMLElement    **ppContainer )
{
    HRESULT                         hr;
    SP_IHTMLElement                 spNewElement;
    SP_IHTMLElement                 spCurrentElement;
    CSmartPtr<IHTMLViewServices>    spViewServices;
    BOOL                            fContainer;

    *ppContainer = NULL;

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID *)&spViewServices) );

    spCurrentElement = pElement;
    do
    {
        IFR( spViewServices->IsContainerElement(spCurrentElement, &fContainer, NULL) ); 

        if (fContainer) // found container
        {
            *ppContainer = spCurrentElement;
            spCurrentElement->AddRef();
            break;
        }            

        IFR( spCurrentElement->get_parentElement(&spNewElement) );
        spCurrentElement = spNewElement;
    }
    while (spCurrentElement != NULL);

    RRETURN(hr);        
}

HRESULT 
EdUtil::FindListItem( IMarkupServices *pMarkupServices,
                      IHTMLElement    *pElement,
                      IHTMLElement    **ppListContainer )
{
    HRESULT                         hr;
    SP_IHTMLElement                 spNewElement;
    SP_IHTMLElement                 spCurrentElement;
    ELEMENT_TAG_ID                  tagId;
    BOOL                            bLayout;
    CSmartPtr<IHTMLViewServices>    spViewServices;

    *ppListContainer = NULL;

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID *)&spViewServices) );

    spCurrentElement = pElement;
    do
    {
        IFR( pMarkupServices->GetElementTagId(spCurrentElement, &tagId) ); 

        if (IsListItem(tagId)) // found container
        {
            *ppListContainer = spCurrentElement;
            spCurrentElement->AddRef();
            break;
        }            
            
        IFR (spViewServices->IsLayoutElement(spCurrentElement, &bLayout) );
        if (bLayout)
            return S_OK; // done - don't cross layout
            
        IFR( spCurrentElement->get_parentElement(&spNewElement) );
        spCurrentElement = spNewElement;
    }
    while (spCurrentElement != NULL);

    RRETURN(hr);        
}

HRESULT 
EdUtil::InsertBlockElement(
    IMarkupServices *pMarkupServices, 
    IHTMLElement    *pElement, 
    IMarkupPointer  *pStart, 
    IMarkupPointer  *pEnd, 
    IMarkupPointer  *pCaret)
{
    HRESULT             hr;
    BOOL                bAdjustStart = FALSE;
    BOOL                bAdjustEnd = FALSE;
    

    //
    // Do we need to adjust the pointer after insiertion?
    //

    IFR( pCaret->IsEqualTo(pStart, &bAdjustStart) );    
    if (!bAdjustStart)
    {
        IFR( pCaret->IsEqualTo(pEnd, &bAdjustEnd) );    
    }

    //
    // Insert the element
    //
    
    IFR( InsertElement(pMarkupServices, pElement, pStart, pEnd) );

    //
    // Fixup the pointer
    //
    
    if (bAdjustStart)
    {
        IFR( pCaret->MoveAdjacentToElement(pElement, ELEM_ADJ_AfterBegin) );
    }
    else if (bAdjustEnd)
    {
        IFR( pCaret->MoveAdjacentToElement(pElement, ELEM_ADJ_BeforeEnd) );
    }
    
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   NextEventTime
//
//  Synopsis:   Returns a value which can be use to determine when a given
//              number of milliseconds has passed.
//
//  Arguments:  [ulDelta] -- Number of milliseconds after which IsTimePassed
//                           will return TRUE.
//
//  Returns:    A timer value.  Guaranteed not to be zero.
//
//  Notes:      Due to the algorithm used in IsTimePassed, [ulDelta] cannot
//              be greater than ULONG_MAX/2.
//
//----------------------------------------------------------------------------

ULONG
EdUtil::NextEventTime(ULONG ulDelta)
{
    ULONG ulCur;
    ULONG ulRet;

    Assert(ulDelta < ULONG_MAX/2);

    ulCur = GetTickCount();

    if ((ULONG_MAX - ulCur) < ulDelta)
        ulRet = ulDelta - (ULONG_MAX - ulCur);
    else
        ulRet = ulCur + ulDelta;

    if (ulRet == 0)
        ulRet = 1;

    return ulRet;
}

//+---------------------------------------------------------------------------
//
//  Function:   IsTimePassed
//
//  Synopsis:   Returns TRUE if the current time is later than the given time.
//
//  Arguments:  [ulTime] -- Timer value returned from NextEventTime().
//
//  Returns:    TRUE if the current time is later than the given time.
//
//----------------------------------------------------------------------------

BOOL
EdUtil::IsTimePassed(ULONG ulTime)
{
    ULONG ulCur = GetTickCount();

    if ((ulCur > ulTime) && (ulCur - ulTime < ULONG_MAX/2))
        return TRUE;

    return FALSE;
}


HRESULT 
EdUtil::GetEditResourceLibrary(
    HINSTANCE   *hResourceLibrary)
{    
    //if (!g_hEditLibInstance)
    //{
    //    g_hEditLibInstance = MLLoadLibrary(_T("mshtmler.dll"), g_hInstance, ML_CROSSCODEPAGE);
    //}
    //*hResourceLibrary = g_hEditLibInstance;

    //if (!g_hEditLibInstance)
    //    return E_FAIL; // TODO: can we convert GetLastError() to an HRESULT?

    return S_OK;
}

            
BOOL
EdUtil::SameElements(
    IHTMLElement *      pElement1,
    IHTMLElement *      pElement2 )
{
    HRESULT hr = S_OK;
    BOOL fEqual = FALSE;

    IObjectIdentity * pId1 = NULL;

    if( pElement1 == NULL || pElement2 == NULL )
        goto Cleanup;
        
    IFC( pElement1->QueryInterface( IID_IObjectIdentity , (LPVOID *) & pId1 ));
    hr = pId1->IsEqualObject( pElement2 );
    fEqual = ( hr == S_OK );
    
Cleanup:
    ReleaseInterface( pId1 );
    return fEqual;
}


//+====================================================================================
//
// Method: EquivalentElements
//
// Synopsis: Test elements for 'equivalency' - ie if they are the same element type,
//           and have the same class, id , and style.
//
//------------------------------------------------------------------------------------

BOOL 
EdUtil::EquivalentElements( 
            IMarkupServices* pMarkupServices, IHTMLElement* pIElement1, IHTMLElement* pIElement2 )
{
    ELEMENT_TAG_ID eTag1 = TAGID_NULL;
    ELEMENT_TAG_ID eTag2 = TAGID_NULL;
    BOOL fEquivalent = FALSE;
    HRESULT hr = S_OK;
    IHTMLStyle * pIStyle1 = NULL;
    IHTMLStyle * pIStyle2 = NULL;
    BSTR id1 = NULL;
    BSTR id2 = NULL;
    BSTR class1 = NULL;
    BSTR class2 = NULL;
    BSTR style1 = NULL;
    BSTR style2 = NULL;
    
    IFC( pMarkupServices->GetElementTagId( pIElement1, & eTag1 ));
    IFC( pMarkupServices->GetElementTagId( pIElement2, & eTag2 ));

    //
    // Compare Tag Id's
    //
    if ( eTag1 != eTag2 )
        goto Cleanup;

    //
    // Compare Id's
    //
    IFC( pIElement1->get_id( & id1 ));
    IFC( pIElement2->get_id( & id2 ));

    if ((( id1 != NULL ) || ( id2 != NULL )) && 
        ( StrCmpIW( id1, id2) != 0))
        goto Cleanup;

    //
    // Compare Class
    //
    IFC( pIElement1->get_className( & class1 ));
    IFC( pIElement2->get_className( & class2 ));

        
    if ((( class1 != NULL ) || ( class2 != NULL )) &&
        ( StrCmpIW( class1, class2) != 0 ) )
        goto Cleanup;

    //
    // Compare Style's
    //        
    IFC( pIElement1->get_style( & pIStyle1 ));
    IFC( pIElement2->get_style( & pIStyle2 ));
    IFC( pIStyle1->toString( & style1 ));
    IFC( pIStyle2->toString( & style2 ));
       
    if ((( style1 != NULL ) || ( style2 != NULL )) &&
        ( StrCmpIW( style1, style2) != 0 ))
        goto Cleanup;

    fEquivalent = TRUE;        
Cleanup:
    SysFreeString( id1 );
    SysFreeString( id2 );
    SysFreeString( class1 );
    SysFreeString( class2 );
    SysFreeString( style1 );
    SysFreeString( style2 );
    ReleaseInterface( pIStyle1 );
    ReleaseInterface( pIStyle2 );
    
    AssertSz(!FAILED(hr), "Unexpected failure in Equivalent Elements");

    return ( fEquivalent );
}


HRESULT 
EdUtil::InsertElement(IMarkupServices *pMarkupServices, IHTMLElement *pElement, IMarkupPointer *pStart, IMarkupPointer *pEnd)
{
    HRESULT                         hr;
    CSmartPtr<IHTMLViewServices>    spViewServices;
    BOOL                            bBlock, bLayout;
    ELEMENT_TAG_ID                  tagId;
    
    IFR( pMarkupServices->InsertElement(pElement, pStart, pEnd) );

    //
    // Set additional attributes
    //

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID *)&spViewServices) );

    IFR( spViewServices->IsBlockElement(pElement, &bBlock) );
    if (bBlock)
    {
        IFR( spViewServices->IsLayoutElement(pElement, &bLayout) );
        if (!bLayout)
        {
            IFR( pMarkupServices->GetElementTagId(pElement, &tagId) );
            if (!IsListContainer(tagId) && tagId != TAGID_LI)
                IFR( spViewServices->InflateBlockElement(pElement) );
        }
    }

    return S_OK;
}

HRESULT 
EdUtil::NeedsNewBlock(IMarkupServices *pMarkupServices, IMarkupPointer *pLeft, IMarkupPointer *pRight, BOOL *pfNeedsBlock)
{
    HRESULT             hr;
    MARKUP_CONTEXT_TYPE context;
    SP_IHTMLElement     spNewElement;
    ELEMENT_TAG_ID      tagId;
    BOOL                bEqual;
    
    CSmartPtr<IHTMLViewServices>    spViewServices;

    *pfNeedsBlock = FALSE;

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID*)&spViewServices) );
    
    IFR( pLeft->IsEqualTo(pRight, &bEqual) );

    // Check left pointer
    context = CONTEXT_TYPE_None; 
    if (!bEqual)
        IFR( BlockMove(pMarkupServices, spViewServices, pLeft, RIGHT, FALSE, &context, NULL) );
        
    if (context != CONTEXT_TYPE_EnterScope)
    {
        IFR( BlockMove(pMarkupServices, spViewServices, pLeft, LEFT, FALSE, &context, &spNewElement) );
        if (context == CONTEXT_TYPE_ExitScope)
        {
            IFR( pMarkupServices->GetElementTagId(spNewElement, &tagId) );
            *pfNeedsBlock = (tagId == TAGID_BODY) || IsListContainer(tagId);
        }
        else
        {
            *pfNeedsBlock = TRUE;
        }
    } 

    // Check right pointer
    if (!(*pfNeedsBlock ))
    {
        context = CONTEXT_TYPE_None;
        if (!bEqual)
            IFR( BlockMove(pMarkupServices, spViewServices, pRight, LEFT, FALSE, &context, NULL) );

        if (context != CONTEXT_TYPE_EnterScope)
        {
            IFR( BlockMove(pMarkupServices, spViewServices, pRight, RIGHT, FALSE, &context, &spNewElement) );
            if (context == CONTEXT_TYPE_ExitScope)
            {
                IFR( pMarkupServices->GetElementTagId(spNewElement, &tagId) );            
                *pfNeedsBlock = (tagId == TAGID_BODY) || IsListContainer(tagId);
            }
            else
            {
                *pfNeedsBlock = TRUE;
            }
        } 
    }

    return S_OK;
}

HRESULT 
EdUtil::FindTagAbove( IMarkupServices *pMarkupServices,
                      IHTMLElement    *pElement,
                      ELEMENT_TAG_ID  tagIdGoal,
                      IHTMLElement    **ppElement,
                      BOOL            fStopAtLayout)
{
    HRESULT                         hr;
    SP_IHTMLElement                 spNewElement;
    SP_IHTMLElement                 spCurrentElement;
    ELEMENT_TAG_ID                  tagId;
    CSmartPtr<IHTMLViewServices>    spViewServices;
    BOOL                            fSite;
    
    *ppElement = NULL;

    IFR( pMarkupServices->QueryInterface(IID_IHTMLViewServices, (LPVOID*)&spViewServices) );

    spCurrentElement = pElement;
    do
    {
        IFR( pMarkupServices->GetElementTagId(spCurrentElement, &tagId) ); 

        if (tagId == tagIdGoal) // found container
        {
            *ppElement= spCurrentElement;
            spCurrentElement->AddRef();
            break;
        }            

        if (fStopAtLayout)
        {
            IFR( spViewServices->IsSite(spCurrentElement, &fSite, NULL, NULL, NULL) );
            if (fSite)
                break;                
        }
        
        IFR( spCurrentElement->get_parentElement(&spNewElement) );
        spCurrentElement = spNewElement;
    }
    while (spCurrentElement != NULL);

    RRETURN(hr);        
    
}


BOOL 
EdUtil::IsInWindow( HWND hwnd, POINT pt, BOOL fConvertToScreen /*=FALSE*/ )
{
    RECT windowRect ;
    POINT myPt;
    
    myPt.x = pt.x;
    myPt.y = pt.y;

    if ( fConvertToScreen )
        ::ClientToScreen( hwnd, &myPt );

    ::GetWindowRect( hwnd, & windowRect );

    return ( ::PtInRect( &windowRect, myPt ) ); 
}

//
// CUndoUnit helper
//

CUndoUnit::CUndoUnit(CXindowsEditor *pEd) 
{
    _pEd= pEd; 
    _fEndUndo = FALSE;
}

CUndoUnit::~CUndoUnit() 
{
    if (_fEndUndo) 
        _pEd->GetMarkupServices()->EndUndoUnit();
}

HRESULT 
CUndoUnit::Begin(UINT uiStringId) 
{
    HRESULT hr;
    TCHAR   *pchUndoTitle;
        
    AssertSz(_fEndUndo == FALSE, "Too many undo units opened.");
    Assert(_pEd);
    
    pchUndoTitle = _pEd->GetCachedString(uiStringId);
    if (!pchUndoTitle)
        return E_FAIL;

    IFR( _pEd->GetMarkupServices()->BeginUndoUnit(pchUndoTitle) );

    _fEndUndo = TRUE;
    return S_OK;
}

//
// CStringCache
//

CStringCache::CStringCache(UINT uiStart, UINT uiEnd)
{
    _uiStart = uiStart;
    _uiEnd = uiEnd;
    _pCache = new CacheEntry[_uiEnd - _uiStart + 1];
    
    if (_pCache)
    {
        for (UINT i = 0; i < (_uiEnd - _uiStart + 1); i++)
        {
            _pCache[i].pchString = NULL;
        }
    }
}

CStringCache::~CStringCache()
{
    for (UINT i = 0; i < (_uiEnd - _uiStart + 1); i++)
    {
        delete [] _pCache[i].pchString;
    }
    delete [] _pCache;
}

TCHAR *
CStringCache::GetString(UINT uiStringId)
{
    HRESULT     hr;
    CacheEntry  *pEntry;
    HINSTANCE   hinstEditResDLL;
    INT         iResult;
    const int   iBufferSize = 1024;
    
    Assert(_pCache);
    if (!_pCache || uiStringId < _uiStart || uiStringId > _uiEnd)
        return NULL; // error

    pEntry = &_pCache[uiStringId - _uiStart];    
    
    if (pEntry->pchString == NULL)
    {
        TCHAR pchBuffer[iBufferSize];
        
        IFC( GetEditResourceLibrary(&hinstEditResDLL) );
        pchBuffer[iBufferSize-1] = 0;  // so we are always 0 terminated
        iResult = LoadString( hinstEditResDLL, uiStringId, pchBuffer, ARRAYSIZE(pchBuffer)-1 );
        if (!iResult)
            goto Cleanup;

        pEntry->pchString = new TCHAR[_tcslen(pchBuffer)+1];
        if (pEntry->pchString)
            StrCpy(pEntry->pchString, pchBuffer);         
    }

    return pEntry->pchString;

Cleanup:
    return NULL;
}
