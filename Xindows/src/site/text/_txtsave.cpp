
#include "stdafx.h"
#include "_txtsave.h"

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver constructor
//
//  Synopsis:   Set up a range saver object.
//
//----------------------------------------------------------------------------
CRangeSaver::CRangeSaver(
         CMarkupPointer*    pLeft,
         CMarkupPointer*    pRight,
         DWORD              dwFlags,
         CStreamWriteBuff*  pswb,
         CMarkup*           pMarkup,
         CElement*          pelContainer) : CTreeSaver(pMarkup)
{
    Initialize(pLeft, pRight, dwFlags, pswb, pMarkup, pelContainer);
}

void CRangeSaver::Initialize(
        CMarkupPointer*     pLeft,
        CMarkupPointer*     pRight,
        DWORD               dwFlags,
        CStreamWriteBuff*   pswb,
        CMarkup*            pMarkup,
        CElement*           pelContainer)
{
    Assert(pLeft && pRight && pswb && pMarkup);
    Assert(pLeft->IsPositioned() && pRight->IsPositioned());
    Assert(pLeft->Markup() == pMarkup);
    Assert(pRight->Markup() == pMarkup);

    pMarkup->EmbedPointers();

    // Set CTreeSaver data
    _pswb            = pswb;
    _pMarkup         = pMarkup;
    _pelContainer    = pelContainer;
    _fSymmetrical    = TRUE;

    Verify(!_tpgStart.MoveTo(pLeft->GetEmbeddedTreePos(), TPG_RIGHT));
    _tpgStart.SetAttachPreference(TPG_RIGHT);
    Verify(!_tpgStart.MoveRight(TPG_OKNOTTOMOVE|TPG_SKIPPOINTERS));
    _fLBStartLeft = FALSE;

    Verify(!_tpgEnd.MoveTo(pRight->GetEmbeddedTreePos(), TPG_RIGHT));
    _tpgEnd.SetAttachPreference(TPG_RIGHT);
    Verify(!_tpgEnd.MoveRight(TPG_OKNOTTOMOVE|TPG_SKIPPOINTERS));
    _fLBEndLeft = FALSE;    

    Assert(!_pelContainer
        || (_tpgStart.Branch()->SearchBranchToRootForScope(pelContainer)
        && _tpgEnd.Branch()->SearchBranchToRootForScope(pelContainer)));

    // Set CRangeSaver data
    _dwFlags = dwFlags;
    memset(&_header, 0, sizeof(_header));

    if(!(_dwFlags & RSF_NO_IE4_COMPAT_SEL))
    {
        DoIE4SelectionCollapse();
    }

    // Compute fragment element
    if(_dwFlags & RSF_NO_IE4_COMPAT_FRAG)
    {
        CTreeNode* pNodeLeft = _tpgStart.Branch();
        CTreeNode* pNodeRight = _tpgEnd.Branch();
        CTreeNode* pNodeCommon;

        pNodeCommon = pNodeLeft->GetFirstCommonAncestor(pNodeRight, NULL);
        Assert(pNodeCommon);

        _pelFragment = pNodeCommon->Element();
    }
    else
    {
        ComputeIE4Fragment();
    }

    if(!(_dwFlags&RSF_NO_IE4_COMPAT_SEL) && _dwFlags&RSF_SELECTION)
    {
        ComputeIE4Selection();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::SaveSelection
//
//  Synopsis:   Called by CTreeSaver::Save to signify where the selection
//              begins and ends
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::SaveSelection(BOOL fEnd)
{
    HRESULT hr = S_OK;

    if(_dwFlags & RSF_CFHTML_HEADER)
    {
        if(fEnd)
        {
            // Record the offset for the selection end
            hr = GetStmOffset(&_header.iSelectionEnd);
            if(hr)
            {
                goto Cleanup;
            }
        }
        else        
        {
            // Record the offset for the selection beginning
            hr = GetStmOffset(&_header.iSelectionStart);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }
Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::DoIE4SelectionCollapse
//
//  Synopsis:   Moves the gaps in to simulate the range that
//              IE4 started with when doing a range save
//
//----------------------------------------------------------------------------
void CRangeSaver::DoIE4SelectionCollapse()
{
    // BUGBUG: handle a range with no content under it!

    // Move the start gap in until we hit real content
    _tpgStart.SetAttachPreference( TPG_RIGHT );

    while(!(_tpgStart.AttachedTreePos()->IsText()
        && _tpgStart.AttachedTreePos()->Cch()))
    {
        // never cross container boundries
        if(_tpgStart.AttachedTreePos()->IsNode()
            && _tpgStart.AttachedTreePos()->IsEdgeScope()
            && _tpgStart.AttachedTreePos()->Branch()->IsContainer())
        {
            break;
        }

        // If we hit a LB char and we don't have it marked
        // on our left, break
        if(!_fLBStartLeft && LineBreakChar(&_tpgStart))
        {
            break;
        }

        _fLBStartLeft = FALSE;

        // BUGBUG: is this IE4 compat?
        if(_tpgStart == _tpgEnd)
        {
            break;
        }

        Verify(!_tpgStart.MoveRight());
    }

    Verify(!_tpgStart.MoveLeft(TPG_VALIDGAP|TPG_OKNOTTOMOVE));

    // Move the end gap in until we hit real content
    _tpgEnd.SetAttachPreference(TPG_LEFT);

    while(!(_tpgEnd.AttachedTreePos()->IsText() 
        && _tpgEnd.AttachedTreePos()->Cch()) && ! _fLBEndLeft)
    {
        // never cross container boundries
        if(_tpgEnd.AttachedTreePos()->IsNode()
            && _tpgEnd.AttachedTreePos()->IsEdgeScope()
            && _tpgEnd.AttachedTreePos()->Branch()->IsContainer())
        {
            break;
        }

        // BUGBUG: is this IE4 compat?
        if(_tpgStart == _tpgEnd)
        {
            break;
        }

        Verify(!_tpgEnd.MoveLeft());

        if(LineBreakChar(&_tpgEnd))
        {
            _fLBEndLeft = TRUE;
            break;
        }
    }

    Verify(!_tpgEnd.MoveRight(TPG_VALIDGAP|TPG_OKNOTTOMOVE));
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::ComputeIE4Fragment
//
//  Synopsis:   Tries to figure out the same fragment element
//              that IE4 would have.
//
//----------------------------------------------------------------------------
void CRangeSaver::ComputeIE4Fragment()
{
    CTreeNode* pNodeLeft = _tpgStart.Branch();
    CTreeNode* pNodeRight = _tpgEnd.Branch();
    CTreePosGap tpgLeft, tpgRight;

    // Compute the first common element ancestor for the range
    pNodeLeft = pNodeLeft->GetFirstCommonAncestor(pNodeRight, NULL);
    Assert(pNodeLeft);

    pNodeRight = pNodeRight->SearchBranchToRootForScope(pNodeLeft->Element());
    Assert(pNodeRight);

    _pelFragment = pNodeRight->Element();

    // Include any phrase element which starts at this "run"
    while(!IsElementBlockInContainer(_pelFragment)
        && !ScopesLeftOfStart(_pelFragment) && _pelFragment!=_pelContainer)
    {
        // move to the next parent that covers the entire range
        pNodeLeft = pNodeLeft->Parent();
        pNodeRight = pNodeRight->Parent();

        pNodeLeft = pNodeLeft->GetFirstCommonAncestor(pNodeRight, NULL);
        Assert(pNodeLeft);

        pNodeRight = pNodeRight->SearchBranchToRootForScope(pNodeLeft->Element());
        Assert(pNodeRight);

        _pelFragment = pNodeRight->Element();
    }

    // If the fragment is selected completely, move up an element
    if(!ScopesLeftOfStart(_pelFragment)
        && ! ScopesRightOfEnd(_pelFragment)
        && !(_pelFragment->HasFlowLayout()
        && _pelFragment->GetFlowLayout()->GetContentMarkup()==_pMarkup)
        && _pelFragment!=_pelContainer 
        && !_pelFragment->IsRoot())
    {
        // move to the next parent that covers the entire range
        pNodeLeft = pNodeLeft->Parent();
        pNodeRight = pNodeRight->Parent();

        pNodeLeft = pNodeLeft->GetFirstCommonAncestor(pNodeRight, NULL);
        Assert(pNodeLeft);

        pNodeRight = pNodeRight->SearchBranchToRootForScope(pNodeLeft->Element());
        Assert(pNodeRight);

        _pelFragment = pNodeRight->Element();
    }

    Assert(_tpgStart.Branch()->SearchBranchToRootForScope(_pelFragment));
    Assert(_tpgEnd.Branch()->SearchBranchToRootForScope(_pelFragment));
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::ComputeIE4Selection
//
//  Synopsis:   Tries to figure out the same selection
//              that IE4 would have.
//
//----------------------------------------------------------------------------
void CRangeSaver::ComputeIE4Selection()
{
    CTreePosGap tpgTemp(TPG_LEFT);
    BOOL        fLBLeft;
    CTreeNode*  pNodeAbove;

    // Move the start gap in until we hit real content
    Verify(!tpgTemp.MoveTo(&_tpgStart));
    fLBLeft = _fLBStartLeft;

    pNodeAbove = tpgTemp.Branch();

    while(!(tpgTemp.AttachedTreePos()->IsText() 
        && tpgTemp.AttachedTreePos()->Cch())  && !fLBLeft)
    {
        // Never cross container boundries
        if(tpgTemp.AttachedTreePos()->IsNode()
            && tpgTemp.AttachedTreePos()->IsEdgeScope()
            && tpgTemp.AttachedTreePos()->Branch()->IsContainer())
        {
            break;
        }

        if(tpgTemp.AttachedTreePos()->IsBeginElementScope()
            && pNodeAbove->SearchBranchToRootForScope(tpgTemp.Branch()->Element()))
        {
            if(tpgTemp.AttachedTreePos()->Branch()->Element() == _pelFragment)
            {
                break;
            }

            Verify(!_tpgStart.MoveTo(tpgTemp.AttachedTreePos(), TPG_LEFT));
            _fLBStartLeft = LineBreakChar(&_tpgStart);
        }

        Verify(!tpgTemp.MoveLeft());

        if(LineBreakChar(&tpgTemp))
        {
            fLBLeft = TRUE;
        }
    }

    Verify(!_tpgStart.MoveLeft(TPG_VALIDGAP|TPG_OKNOTTOMOVE));

    // Move the end gap in until we hit real content
    Verify(!tpgTemp.MoveTo(&_tpgEnd));
    fLBLeft = _fLBEndLeft;

    pNodeAbove = tpgTemp.Branch();

    tpgTemp.SetAttachPreference(TPG_RIGHT);

    while(!(tpgTemp.AttachedTreePos()->IsText() && tpgTemp.AttachedTreePos()->Cch()))
    {
        // If we hit a LB char and we don't have it marked
        // on our left, break
        if(!fLBLeft && LineBreakChar(&tpgTemp))
        {
            break;
        }

        // Never cross container boundries
        if(tpgTemp.AttachedTreePos()->IsNode()
            && tpgTemp.AttachedTreePos()->IsEdgeScope()
            && tpgTemp.AttachedTreePos()->Branch()->IsContainer())
        {
            break;
        }

        if(tpgTemp.AttachedTreePos()->IsEndElementScope() 
            && pNodeAbove->SearchBranchToRootForScope(tpgTemp.Branch()->Element()))
        {
            if(tpgTemp.AttachedTreePos()->Branch()->Element() == _pelFragment)
            {
                break;
            }

            Verify(!_tpgEnd.MoveTo(tpgTemp.AttachedTreePos(), TPG_RIGHT));
            _fLBEndLeft = FALSE;
        }

        fLBLeft = FALSE;
        Verify(!tpgTemp.MoveRight());
    }

    Verify(!_tpgEnd.MoveRight(TPG_VALIDGAP|TPG_OKNOTTOMOVE));
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::GetStmOffset
//
//  Synopsis:   Returns the current offset in the output stream.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::GetStmOffset(LONG* plOffset)
{
    HRESULT        hr;
    ULARGE_INTEGER ulRet;
    LARGE_INTEGER  lSeekZero  = { 0, 0 };

    hr = _pswb->Seek(lSeekZero, STREAM_SEEK_CUR, &ulRet);
    if(hr)
    {
        goto Cleanup;
    }

    *plOffset = ulRet.QuadPart;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::SetStmOffset
//
//  Synopsis:   Sets the offset in the output stream.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::SetStmOffset(LONG lOffset)
{
    LARGE_INTEGER l;

    l.QuadPart = lOffset;

    RRETURN(_pswb->Seek(l, STREAM_SEEK_SET, NULL));
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::Save
//
//  Synopsis:   Save the range to the stream _pswb.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::Save()
{
    HRESULT hr = S_OK;

    if(_dwFlags & RSF_NO_ENTITIZE_UNKNOWN)
    {
        // Do not entitize unknown characters if asked not to
        _pswb->SetEntitizeUnknownChars(FALSE);
    }

    if(_dwFlags & RSF_CFHTML_HEADER)
    {
        // Write an empty cfhtml header.  We will go back later and fill
        // it out properly once we have the correct offsets.
        hr = WriteCFHTMLHeader();
        if(hr)
        {
            goto Cleanup;
        }

        hr = GetStmOffset(&_header.iHTMLStart);
        if(hr)
        {
            goto Cleanup;
        }
    }

    // Open up the context, starting with the doc header.
    {
        if(!(_dwFlags & RSF_CONTEXT))
        {
            _pswb->BeginSuppress();
        }

        _pMarkup->Doc()->WriteDocHeader(_pswb);

        hr = WriteOpenContext();
        if(hr)
        {
            goto Cleanup;
        }

        if(!(_dwFlags & RSF_CONTEXT))
        {
            _pswb->EndSuppress();
        }
    }

    // The real meat -- where we call the superclass.
    if((_dwFlags&RSF_FRAGMENT) || (_dwFlags&RSF_SELECTION))
    {
        // Call the superclass to save the body 
        hr = CTreeSaver::Save();
        if(hr)
        {
            goto Cleanup;
        }
    }

    {
        if(!(_dwFlags & RSF_CONTEXT))
        {
            _pswb->BeginSuppress();
        }

        hr = WriteCloseContext();
        if(hr)
        {
            goto Cleanup;
        }

        if(!(_dwFlags & RSF_CONTEXT))
        {
            _pswb->EndSuppress();
        }
    }

    if(_dwFlags & RSF_CFHTML_HEADER)
    {
        // Save the current position as the end of html, and go back
        // to the beginning to write the header with the offsets we
        // now know.
        hr = GetStmOffset(&_header.iHTMLEnd);
        if(hr)
        {
            goto Cleanup;
        }

        hr = SetStmOffset(0);
        if(hr)
        {
            goto Cleanup;
        }

        hr = WriteCFHTMLHeader();
        if(hr)
        {
            goto Cleanup;
        }

        hr = SetStmOffset(_header.iHTMLEnd);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CRangeSaver::WriteCFHTMLHeader
//
//  Synopsis:   Writes the CF-HTML header to _pswb.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::WriteCFHTMLHeader()
{
    HRESULT hr;
    TCHAR*  pchOut = NULL;
    DWORD   dwOldFlags = _pswb->ClearFlags(WBF_ENTITYREF);

    Assert(_dwFlags & RSF_CFHTML_HEADER);

    _pswb->SetFlags(WBF_NO_WRAP);

    static TCHAR szCfHtmlHeaderFormat[] =
        _T("Version:<0s>\r\n")
        _T("StartHTML:<1d9>\r\n")
        _T("EndHTML:<2d9>\r\n")
        _T("StartFragment:<3d9>\r\n")
        _T("EndFragment:<4d9>\r\n")
        _T("StartSelection:<5d9>\r\n")
        _T("EndSelection:<6d9>\r\n")
        _T("SourceURL:<7s>\r\n");

    Format(
        _afxGlobalData._hInstResource,
        FMT_OUT_ALLOC,
        & pchOut, 512,
        szCfHtmlHeaderFormat,
        _T("1.0"),
        (long)_header.iHTMLStart, (long)_header.iHTMLEnd, (long)_header.iFragmentStart,
        (long)_header.iFragmentEnd, (long)_header.iSelectionStart, (long)_header.iSelectionEnd,
        _pMarkup->Doc()->_cstrUrl);

    // Do not let _pswb mess with forced cr/lfs in this string
    _pswb->BeginPre();
    hr = _pswb->Write(pchOut);
    if(hr)
    {
        goto Cleanup;
    }
    _pswb->EndPre();

Cleanup:
    MemFree(pchOut);
    _pswb->RestoreFlags(dwOldFlags);
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Members:    CRangeSaver::WriteOpenContext
//
//  Synopsis:   Writes out all tags that come into of scope between
//              _pelFragment and the root.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::WriteOpenContext()
{
    HRESULT hr = S_OK;
    CStackPtrAry<CElement*, 32> aryElements;
    CTreeNode* pNodeCur;

    // First find the context of _pelFragment that is above the start.
    pNodeCur = _tpgStart.Branch();
    Assert(pNodeCur);

    pNodeCur = pNodeCur->SearchBranchToRootForScope(_pelFragment);
    Assert(pNodeCur);

    // build an array of all of the elements that we want to write out
    for(; pNodeCur->Tag()!=ETAG_ROOT; pNodeCur=pNodeCur->Parent())
    {
        aryElements.Append(pNodeCur->Element());
    }

    // Write out our array of elements
    {
        CElement** ppElement;
        int iElement;

        for(iElement=aryElements.Size(),ppElement=&(aryElements[iElement-1]);
            iElement>0; iElement--,ppElement--)
        {
            if(!*ppElement)
            {
                continue;
            }

            hr = SaveElement(*ppElement, FALSE);
            if(hr)
            {
                goto Cleanup;
            }

            if(*ppElement == _pelFragment)
            {
                DWORD dwFlags = _pswb->ClearFlags(WBF_ENTITYREF);

                hr = _pswb->Write(_T("<!--StartFragment-->"));
                if(hr)
                {
                    goto Cleanup;
                }

                _pswb->RestoreFlags(dwFlags);

                hr = GetStmOffset(&_header.iFragmentStart);
                if(hr)
                {
                    goto Cleanup;
                }

                if(_header.iSelectionStart == 0)
                {
                    _header.iSelectionStart = _header.iFragmentStart;
                }
            }

            if((*ppElement)->Tag()==ETAG_HTML && (_dwFlags&RSF_CONTEXT))
            {
                // Make sure not to save out elements we do not want in cfhtml,
                // such as meta charset tags and styles.
                CElement* pHead = NULL/*_pMarkup->GetHeadElement() wlw note*/;

                if(pHead)
                {
                    DWORD dwFlags = _pswb->SetFlags(WBF_NO_CHARSET_META);

                    pHead->Save(_pswb, FALSE);
                    if(hr)
                    {
                        goto Cleanup;
                    }

                    CTreeSaver ts(pHead, _pswb);
                    hr = ts.Save();
                    if(hr)
                    {
                        goto Cleanup;
                    }

                    pHead->Save(_pswb, TRUE);
                    if(hr)
                    {
                        goto Cleanup;
                    }

                    _pswb->RestoreFlags(dwFlags);

                    // There may be elements that start in the head and overlap
                    // over our current stack.  We don't want to write out the
                    // start for these elements twice so zero them out here
                    {
                        CTreePos* ptpCur;

                        pHead->GetTreeExtent(NULL, &ptpCur);
                        Assert(ptpCur);

                        ptpCur = ptpCur->NextTreePos();

                        while(ptpCur->IsBeginNode() && ! ptpCur->IsEdgeScope())
                        {
                            int iIndex = aryElements.Find(ptpCur->Branch()->Element());
                            if(iIndex != -1)
                            {
                                aryElements[iIndex] = NULL;
                            }

                            ptpCur = ptpCur->NextTreePos();
                        }
                    } // no duplicates
                } // pHead
            } // ETAG_HTML

            if(*ppElement && *ppElement!=_pelFragment)
            {
                hr = _pswb->NewLine();
                if(hr)
                {
                    goto Cleanup;
                }
            }

        } // for loop
    } // write out array

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Members:    CRangeSaver::WriteCloseContext
//
//  Synopsis:   Writes out all tags that go out of scope between
//              _pelFragment and the root.
//
//----------------------------------------------------------------------------
HRESULT CRangeSaver::WriteCloseContext()
{
    HRESULT hr = S_OK;
    CTreeNode* pNodeCur;

    pNodeCur = _tpgEnd.Branch();
    Assert(pNodeCur);

    pNodeCur = pNodeCur->SearchBranchToRootForScope(_pelFragment);
    Assert(pNodeCur);

    // build an array of all of the elements that we want to write out
    for(; pNodeCur->Tag()!=ETAG_ROOT; pNodeCur=pNodeCur->Parent())
    {
        CElement* pElementCur = pNodeCur->Element();

        if(pElementCur == _pelFragment)
        {
            DWORD dwFlags = _pswb->ClearFlags(WBF_ENTITYREF);

            hr = GetStmOffset(&_header.iFragmentEnd);
            if(hr)
            {
                goto Cleanup;
            }

            if(_header.iSelectionEnd == 0)
            {
                _header.iSelectionEnd = _header.iFragmentEnd;
            }

            hr = _pswb->Write(_T("<!--EndFragment-->"));
            if(hr)
            {
                goto Cleanup;
            }

            _pswb->RestoreFlags(dwFlags);
        }

        hr = SaveElement(pElementCur, TRUE);
        if(hr)
        {
            goto Cleanup;
        }

        hr = ForceClose(pElementCur);
        if(hr)
        {
            goto Cleanup;
        }

        hr = _pswb->NewLine();
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}
