
#include "stdafx.h"
#include "LineServices.h"

#include "OneRun.h"
#include "LSMeasurer.h"
#include "MarginInfo.h"
#include "lsobj.h"
#include "fontlnk.h"
#include "selrensv.h"
#include "lsrender.h"

#include "msls/lsdefs.h"
#include "msls/lskysr.h"
#include "msls/lscontxt.h"
#include "msls/tatenak.h"
#include "msls/hih.h"
#include "msls/warichu.h"
#include "msls/robj.h"
#include "msls/lssetdoc.h"
#include "msls/lsensubl.h"
#include "msls/lshyph.h"
#include "msls/lsulinfo.h"
#include "msls/lspap.h"
#include "msls/lsenum.h"
#include "msls/lscrline.h"
#include "msls/lsqline.h"
#include "msls/lsffi.h"

#define ONERUN_NORMAL       0x00
#define ONERUN_SYNTHETIC    0x01
#define ONERUN_ANTISYNTH    0x02

enum KASHIDA_PRIORITY
{
    KASHIDA_PRIORITY1,  // SCRIPT_JUSTIFY_ARABIC_TATWEEL
    KASHIDA_PRIORITY2,  // SCRIPT_JUSTIFY_ARABIC_SEEN
    KASHIDA_PRIORITY3,  // SCRIPT_JUSTIFY_ARABIC_HA
    KASHIDA_PRIORITY4,  // SCRIPT_JUSTIFY_ARABIC_ALEF
    KASHIDA_PRIORITY5,  // SCRIPT_JUSTIFY_ARABIC_RA
    KASHIDA_PRIORITY6,  // SCRIPT_JUSTIFY_ARABIC_BARA
    KASHIDA_PRIORITY7,  // SCRIPT_JUSTIFY_ARABIC_BA
    KASHIDA_PRIORITY8,  // SCRIPT_JUSTIFY_ARABIC_NORMAL
    KASHIDA_PRIORITY9,  // Max - lowest priority
};


//-----------------------------------------------------------------------------
//
//  Member:     CLineFlags::AddLineFlag
//
//  Synopsis:   Set flags for a given cp.
//
//-----------------------------------------------------------------------------
LSERR CLineFlags::AddLineFlag(LONG cp, DWORD dwlf)
{   
    int c = _aryLineFlags.Size();

    if(!c || cp>=_aryLineFlags[c-1]._cp)
    {
        CFlagEntry fe(cp, dwlf);

        return (S_OK==_aryLineFlags.AppendIndirect(&fe)
            ? lserrNone : lserrOutOfMemory);
    }

    return lserrNone;
}

LSERR CLineFlags::AddLineFlagForce(LONG cp, DWORD dwlf)
{
    CFlagEntry fe(cp, dwlf);

    _fForced = TRUE;
    return (S_OK==_aryLineFlags.AppendIndirect(&fe)
        ? lserrNone : lserrOutOfMemory);
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetLineFlags
//
//  Synopsis:   Given a cp, it computes all the flags which have been turned
//              on till that cp.
//
//-----------------------------------------------------------------------------
DWORD CLineFlags::GetLineFlags(LONG cpMax)
{
    DWORD dwlf;
    LONG i;

    dwlf = FLAG_NONE;

    for(i=0; i<_aryLineFlags.Size(); i++)
    {
        if(_aryLineFlags[i]._cp >= cpMax)
        {
            if(_fForced)
            {
                continue;
            }
            else
            {
                break;
            }
        }
        else
        {
            dwlf |= _aryLineFlags[i]._dwlf;
        }
    }

    return dwlf;
}


//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::AddLineCount
//
//  Synopsis:   Adds a particular line count at a given cp. It also checks if the
//              count has already been added at that cp. This is needed to solve
//              2 problems with maintaining line counts:
//              1) A run can be fetched multiple times. In this case we want to
//                 increment the counts just once.
//              2) LS can over fetch runs, in which case we want to disregard
//                 the counts of those runs which did not end up on the line.
//
//-----------------------------------------------------------------------------
LSERR CLineCounts::AddLineCount(LONG cp, LC_TYPE lcType, LONG count)
{
    CLineCount lc(cp, lcType, count);
    int i = _aryLineCounts.Size();

    while(i--)
    {
        if(_aryLineCounts[i]._cp != cp)
        {
            break;
        }

        if(_aryLineCounts[i]._lcType == lcType)
        {
            return lserrNone;
        }
    }

    return S_OK==_aryLineCounts.AppendIndirect(&lc)?lserrNone:lserrOutOfMemory;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetLineCount
//
//  Synopsis:   Finds a particular line count at a given cp.
//
//-----------------------------------------------------------------------------
LONG CLineCounts::GetLineCount(LONG cp, LC_TYPE lcType)
{
    LONG count = 0;

    for(LONG i=0; i<_aryLineCounts.Size(); i++)
    {
        if(_aryLineCounts[i]._lcType==lcType && _aryLineCounts[i]._cp<cp)
        {
            count += _aryLineCounts[i]._count;
        }
    }
    return count;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetMaxLineValue
//
//  Synopsis:   Finds a particular line value uptil a given cp.
//
//-----------------------------------------------------------------------------
LONG CLineCounts::GetMaxLineValue(LONG cp, LC_TYPE lcType)
{
    LONG value = LONG_MIN;

    for(LONG i=0; i<_aryLineCounts.Size(); i++)
    {
        if(_aryLineCounts[i]._lcType==lcType && _aryLineCounts[i]._cp<cp)
        {
            value = max(value, _aryLineCounts[i]._count);
        }
    }
    return value;
}


//-----------------------------------------------------------------------------
//
//  Function:   Deinit
//
//  Synopsis:   This function is called during the destruction of the
//              COneRunFreeList. It frees up any allocated COneRun objects.
//
//  Returns:    nothing
//
//-----------------------------------------------------------------------------
void COneRunFreeList::Deinit()
{
    COneRun* por;
    COneRun* porNext;

    por = _pHead;
    while(por)
    {
        Assert(por->_pCF == NULL);
        Assert(por->_pComplexRun == NULL);
        porNext = por->_pNext;
        delete por;
        por = porNext;
    }
}

//-----------------------------------------------------------------------------
//
//  Function: GetFreeOneRun
//
//  Synopsis: Gets a free one run object. If we already have some in the free
//            list, then we need to use those, else allocate off the heap.
//            If porClone is non-NULL then we will clone in that one run
//            into the newly allocated one.
//
//  Returns:  The run
//
//-----------------------------------------------------------------------------
COneRun* COneRunFreeList::GetFreeOneRun(COneRun* porClone)
{
    COneRun* por = NULL;

    if(_pHead)
    {
        por = _pHead;
        _pHead = por->_pNext;
    }
    else
    {
        por = new COneRun();
    }
    if(por)
    {
        if(porClone)
        {
            if(por != por->Clone(porClone))
            {
                SpliceIn(por);
                por = NULL;
                goto Cleanup;
            }
        }
        else
        {
            memset(por, 0, sizeof(COneRun));
            por->_bConvertMode = CVM_UNINITED;
        }
    }
Cleanup:    
    return por;
}

//-----------------------------------------------------------------------------
//
//  Function:   SpliceIn
//
//  Synopsis:   Returns runs which are no longer needed back to the free list.
//              It also uninits all the runs.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void COneRunFreeList::SpliceIn(COneRun* pFirst)
{
    Assert(pFirst);
    COneRun* por = pFirst;
    COneRun* porLast = NULL;

    // Clear out the runs when they are put into the free list.
    while(por)
    {
        porLast = por;
        por->Deinit();

        // TODO(SujalP): por->_pNext is valid after Deinit!!!! Change this so
        // that this code does not depend on this.
        por = por->_pNext;
    }

    Assert(porLast);
    porLast->_pNext = _pHead;
    _pHead = pFirst;
}

//-----------------------------------------------------------------------------
//
//  Function:   Init
//
//-----------------------------------------------------------------------------
void COneRunCurrList::Init()
{
    _pHead = _pTail = NULL;
}

//-----------------------------------------------------------------------------
//
//  Function:   Deinit
//
//-----------------------------------------------------------------------------
void COneRunCurrList::Deinit()
{
    Assert(_pHead==NULL && _pTail==NULL);
}

//-----------------------------------------------------------------------------
//
//  Function:   SpliceOut
//
//  Synopsis:   Removes a chunk of runs from pFirst to pLast from the current list.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void COneRunCurrList::SpliceOut(COneRun* pFirst, COneRun* pLast)
{
    Assert(pFirst && pLast);

    // If the first node being removed is the head node then
    // let us deal with that
    if(pFirst->_pPrev == NULL)
    {
        Assert(pFirst == _pHead);
        _pHead = pLast->_pNext;
    }
    else
    {
        pFirst->_pPrev->_pNext = pLast->_pNext;
    }

    // If the last node being removed is the tail node then
    // let us deal with that
    if(pLast->_pNext == NULL)
    {
        Assert(pLast == _pTail);
        _pTail = pFirst->_pPrev;
    }
    else
    {
        pLast->_pNext->_pPrev = pFirst->_pPrev;
    }

    // Clear the next and prev pointers in the spliced out portion
    pFirst->_pPrev = NULL;
    pLast->_pNext = NULL;
}

//-----------------------------------------------------------------------------
//
//  Function:   SpliceInAfterMe
//
//  Synopsis:   Adds pFirst into the currentlist after the position
//              indicated by pAfterMe. If pAfterMe is NULL its added to the head.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void COneRunCurrList::SpliceInAfterMe(COneRun* pAfterMe, COneRun* pFirst)
{
    COneRun** ppor;

    ppor = (pAfterMe==NULL) ? &_pHead : &pAfterMe->_pNext;
    pFirst->_pNext = *ppor;
    *ppor = pFirst;
    pFirst->_pPrev = pAfterMe;

    COneRun* pBeforeMe = pFirst->_pNext;
    ppor = pBeforeMe==NULL ? &_pTail : &pBeforeMe->_pPrev;
    *ppor = pFirst;
}

//-----------------------------------------------------------------------------
//
//  Function:   SpliceInBeforeMe
//
//  Synopsis:   Adds the onerun identified by pFirst before the run
//              identified by pBeforeMe. If pBeforeMe is NULL then it
//              adds it at the tail.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void COneRunCurrList::SpliceInBeforeMe(COneRun* pBeforeMe, COneRun* pFirst)
{
    COneRun** ppor;

    ppor = pBeforeMe==NULL ? &_pTail : &pBeforeMe->_pPrev;
    pFirst->_pPrev = *ppor;
    *ppor = pFirst;
    pFirst->_pNext = pBeforeMe;

    COneRun* pAfterMe = pFirst->_pPrev;
    ppor = pAfterMe == NULL ? &_pHead : &pAfterMe->_pNext;
    *ppor = pFirst;
}



//-----------------------------------------------------------------------------
//
//  Function:   InitLineServices (global)
//
//  Synopsis:   Instantiates a instance of the CLineServices object and makes
//              the requisite calls into the LineServices DLL.
//
//  Returns:    HRESULT
//              *ppLS - pointer to newly allocated CLineServices object
//
//-----------------------------------------------------------------------------
HRESULT InitLineServices(CMarkup* pMarkup, BOOL fStartUpLSDLL, CLineServices** ppLS)
{
    HRESULT hr = S_OK;
    CLineServices* pLS;

    // Create our Line Services interface object
    pLS = new CLineServices(pMarkup);
    if(!pLS)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    
    pLS->_plsc = NULL;
    if(fStartUpLSDLL)
    {
        hr = StartUpLSDLL(pLS, pMarkup);
        if(hr)
        {
            delete pLS;
            goto Cleanup;
        }
    }

    // Return value
    *ppLS = pLS;

Cleanup:
    RRETURN(hr);
}

HRESULT StartUpLSDLL(CLineServices* pLS, CMarkup* pMarkup)
{
    LSCONTEXTINFO lsci;
    HRESULT hr = S_OK;

    if(pMarkup != pLS->GetMarkup())
    {
        pLS->_treeInfo._tpFrontier.Reinit(pMarkup, 0);
    }

    if(pLS->_plsc)
    {
        goto Cleanup;
    }

    // Fetch the Far East object handlers
    hr = pLS->SetupObjectHandlers();
    if(hr)
    {
        goto Cleanup;
    }

    // Populate the LSCONTEXTINFO
    lsci.version = 0;
    lsci.cInstalledHandlers = CLineServices::LSOBJID_COUNT;
    *(CLineServices::LSIMETHODS**)&lsci.pInstalledHandlers = pLS->s_rgLsiMethods;
    lsci.lstxtcfg = pLS->s_lstxtcfg;
    lsci.pols = (POLS)pLS;
    *(CLineServices::LSCBK*)&lsci.lscbk = CLineServices::s_lscbk;
    lsci.fDontReleaseRuns = TRUE;

    // Call in to Line Services
    hr = HRFromLSERR(LsCreateContext(&lsci, &pLS->_plsc));
    if(hr)
    {
        goto Cleanup;
    }

    // Set Expansion/Compression tables
    hr = pLS->SetModWidthPairs();
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//-----------------------------------------------------------------------------
//
//  Function:   DeinitLineServices (global)
//
//  Synopsis:   Frees a CLineServes object.
//
//  Returns:    HRESULT
//
//-----------------------------------------------------------------------------
HRESULT DeinitLineServices(CLineServices* pLS)
{
    HRESULT hr = S_OK;

    if(pLS->_plsc)
    {
        hr = HRFromLSERR(LsDestroyContext(pLS->_plsc));
    }

    delete pLS;

    RRETURN(hr);
}

//-----------------------------------------------------------------------------
//
//  Function:   SetupObjectHandlers (member)
//
//  Synopsis:   LineServices uses object handlers for special textual
//              representation.  There are six such objects in Trident,
//              and for five of these, the callbacks are implemented by
//              LineServices itself.  The sixth object, our handle for
//              embedded/nested objects, is implemented in lsobj.cxx.
//
//  Returns:    S_OK - Success
//              E_FAIL - A LineServices error occurred
//
//-----------------------------------------------------------------------------
HRESULT CLineServices::SetupObjectHandlers()
{
    HRESULT hr = E_FAIL;
    ::LSIMETHODS* pLsiMethod;

    pLsiMethod = (::LSIMETHODS*)s_rgLsiMethods;

    if(lserrNone != LsGetRubyLsimethods(pLsiMethod+LSOBJID_RUBY))
    {
        goto Cleanup;
    }

    if(lserrNone != LsGetTatenakayokoLsimethods(pLsiMethod+LSOBJID_TATENAKAYOKO))
    {
        goto Cleanup;
    }

    if(lserrNone != LsGetHihLsimethods(pLsiMethod+LSOBJID_HIH))
    {
        goto Cleanup;
    }

    if(lserrNone != LsGetWarichuLsimethods(pLsiMethod+LSOBJID_WARICHU))
    {
        goto Cleanup;
    }

    if(lserrNone != LsGetReverseLsimethods(pLsiMethod+LSOBJID_REVERSE))
    {
        goto Cleanup;
    }

    hr = S_OK;

Cleanup:
    return hr;
}

//-----------------------------------------------------------------------------
//
//  Function:   NewPtr (member, LS callback)
//
//  Synopsis:   A client-side allocation routine for LineServices.
//
//  Returns:    Pointer to buffer allocated, or NULL if out of memory.
//
//-----------------------------------------------------------------------------
void* WINAPI CLineServices::NewPtr(DWORD cb)
{
    void* p;

    p = MemAlloc(cb);

    MemSetName((p, "CLineServices::NewPtr"));

    return p;
}

//-----------------------------------------------------------------------------
//
//  Function:   DisposePtr (member, LS callback)
//
//  Synopsis:   A client-side 'free' routine for LineServices
//
//  Returns:    Nothing.
//
//-----------------------------------------------------------------------------
void WINAPI CLineServices::DisposePtr(void* p)
{
    MemFree(p);
}

//-----------------------------------------------------------------------------
//
//  Function:   ReallocPtr (member, LS callback)
//
//  Synopsis:   A client-side reallocation routine for LineServices
//
//  Returns:    Pointer to new buffer, or NULL if out of memory
//
//-----------------------------------------------------------------------------
void* WINAPI CLineServices::ReallocPtr(void* p, DWORD cb)
{
    void* q = p;
    HRESULT hr;

    hr = MemRealloc(&q, cb);

    return hr?NULL:q;
}

LSERR WINAPI CLineServices::GleanInfoFromTheRun(COneRun* por, COneRun** pporOut)
{
    LSERR         lserr = lserrNone;
    const CCharFormat* pCF;
    BOOL          fWasTextRun = TRUE;
    SYNTHTYPE     synthCur = SYNTHTYPE_NONE;
    LONG          cp = por->Cp();
    LONG          nDirLevel;
    LONG          nSynthDirLevel;
    COneRun*      porOut = por;
    BOOL          fLastPtp = por->_ptp == _treeInfo._ptpLayoutLast;
    BOOL          fNodeRun = por->_ptp->IsNode();
    CTreeNode*    pNodeRun = fNodeRun ? por->_ptp->Branch() : NULL;

    por->_fHidden = FALSE;
    porOut->_fNoTextMetrics = FALSE;

    if(_pMeasurer->_fMeasureFromTheStart && (cp-_cpStart)<_pMeasurer->_cchPreChars)
    {
        WhiteAtBOL(cp);
        por->MakeRunAntiSynthetic();
        goto Done;
    }

    // Take care of hidden runs. We will simply anti-synth them
    if(por->_fCharsForNestedElement && por->GetCF()->IsDisplayNone())
    {
        Assert(!_fIsEditable);
        if(IsFirstNonWhiteOnLine(cp))
        {
            WhiteAtBOL(cp, por->_lscch);
        }
        _lineCounts.AddLineCount(cp, LC_HIDDEN, por->_lscch);
        por->MakeRunAntiSynthetic();
        goto Done;
    }

    // BR with clear causes a break after the line break,
    // where as clear on phrase or block elements should clear
    // before the phrase or block element comes into scope.
    // Clear on aligned elements is handled separately, so
    // do not mark the line with clear flags for aligned layouts.
    if(cp != _cpStart
        && fNodeRun
        && pNodeRun->Tag()!=ETAG_BR
        && !por->GetFF()->_fAlignedLayout
        && _pMeasurer->TestForClear(_pMarginInfo, cp-1, TRUE, por->GetFF()))
    {
        lserr = TerminateLine(por, TL_ADDEOS, &porOut);
        Assert(lserr!=lserrNone || porOut->_synthType!=SYNTHTYPE_NONE);
        goto Cleanup;
    }

    if(!por->_fCharsForNestedLayout && fNodeRun)
    {
        CElement*   pElement          = pNodeRun->Element();
        BOOL        fFirstOnLineInPre = FALSE;
        const CCharFormat* pCF        = por->GetCF();

        // This run can never be hidden, no matter what the HTML says
        Assert(!por->_fHidden);

        if(!fLastPtp && IsFirstNonWhiteOnLine(cp))
        {
            WhiteAtBOL(cp);

            // If we have a <BR> inside a PRE tag then we do not want to
            // break if the <BR> is the first thing on this line. This is
            // because there should have been a \r before the <BR> which
            // caused one line break and we should not have the <BR> break
            // another line. The only execption is when the <BR> is the
            // first thing in the paragraph. In this case we *do* want to
            // break (the exception was discovered via bug 47870).
            if(_pPFFirst->HasPre(_fInnerPFFirst)
                && !(_lsMode==LSMODE_MEASURER?_li._fFirstInPara:_pli->_fFirstInPara))
            {
                fFirstOnLineInPre = TRUE;
            }
        }

        if(por->_ptp->IsEdgeScope()
            && (_pFlowLayout->IsElementBlockInContext(pElement) || fLastPtp)
            && (!fFirstOnLineInPre
            || (por->_ptp->IsEndElementScope() && pElement->_fBreakOnEmpty && pElement->Tag()==ETAG_PRE))
            && pElement->Tag()!=ETAG_BR)
        {

            lserr = TerminateLine(por, (fLastPtp?TL_ADDEOS:TL_ADDLBREAK), &porOut);
            if(lserr!=lserrNone || !porOut)
            {
                lserr = lserrOutOfMemory;
                goto Done;
            }

            // Hide the run that contains the WCH_NODE for the end
            // edge in the layout
            por->_fHidden = TRUE;

            if(fLastPtp)
            {
                if(IsFirstNonWhiteOnLine(cp)
                    && (_fIsEditable || _pFlowLayout->GetContentTextLength()==0))
                {
                    CCcs* pccs = GetCcs(por, _pci->_hdc, _pci);
                    if(pccs)
                    {
                        RecalcLineHeight(pccs, &_li);
                    }

                    // This line is not a dummy line. It has a height. The
                    // code later will treat it as a dummy line. Prevent
                    // that from happening.
                    _fLineWithHeight = TRUE;
                }
                goto Done;
            }
            // Bug66768: If we came here at BOL without the _fHasBulletOrNum flag being set, it means
            // that the previous line had a <BR> and we will terminate this line at the </li> so that
            // all it contains is the </li>. In this case we do infact want the line to have a height
            // so users can type there.
            else if(ETAG_LI==pElement->Tag()
                && por->_ptp->IsEndNode()
                && !_li._fHasBulletOrNum
                && IsFirstNonWhiteOnLine(cp))
            {
                CCcs* pccs = GetCcs(por, _pci->_hdc, _pci);
                if(pccs)
                {
                    RecalcLineHeight(pccs, &_li);
                }
                _fLineWithHeight = TRUE;
            }
        }
        else if(por->_ptp->IsEndNode() && pElement->Tag()==ETAG_BR && !fFirstOnLineInPre)
        {
            _lineFlags.AddLineFlag(cp-1, FLAG_HAS_A_BR);
            AssertSz(por->_ptp->IsEndElementScope(), "BR's cannot be proxied!");
            Assert(por->_lscch == 1);
            Assert(por->_lscchOriginal == 1);

            lserr = TerminateLine(por, TL_ADDNONE, &porOut);
            if(lserr != lserrNone)
            {
                goto Done;
            }
            if(!porOut)
            {
                porOut = por;
            }

            por->FillSynthData(SYNTHTYPE_LINEBREAK);
            _pMeasurer->TestForClear(_pMarginInfo, cp, FALSE, por->GetFF());
            if(IsFirstNonWhiteOnLine(cp))
            {
                _fLineWithHeight = TRUE;
            }
        }
        // Handle premature Ruby end here
        // Example: <RUBY>some text</RUBY>
        else if(pElement->Tag()==ETAG_RUBY
            && por->_ptp->IsEndNode()       // this is an end ruby tag
            && _fIsRuby                     // and we currently have an open ruby
            && !_fIsRubyText                // and we have not yet closed it off
            && !IsFrozen())
        {
            COneRun* porTemp = NULL;
            Assert(por->_lscch == 1);
            Assert(por->_lscchOriginal == 1);

            // if we got here then we opened a ruby but never ended the main
            // text (i.e., with an RT).  Now we want to end everything, so this
            // involves first appending a synthetic to end the main text and then
            // another one to end the (nonexistent) pronunciation text.
            lserr = AppendILSControlChar(por, SYNTHTYPE_ENDRUBYMAIN, &porOut);
            Assert(lserr!=lserrNone || porOut->_synthType!=SYNTHTYPE_NONE);
            lserr = AppendILSControlChar(por, SYNTHTYPE_ENDRUBYTEXT, &porTemp);
            Assert(lserr!=lserrNone || porTemp->_synthType!=SYNTHTYPE_NONE);

            // We set this to FALSE because por will eventually be marked as
            // "not processed yet", which means that the above condition will trip
            // again unless we indicate that the ruby is now closed
            _fIsRuby = FALSE;
        }
        else if(_fNoBreakForMeasurer
            && por->_ptp->IsEndNode()
            && (pElement->Tag()==ETAG_NOBR
            || pElement->Tag()==ETAG_WBR)
            && !IsFrozen())
        {
            // NOTES on WBR handling: By ending the NOBR ILS obj here, we create a
            // break oportunity.  When fetchrun gets called again, the check below
            // will see (from the pcf) that the text still can't break, but that
            // it's not within an ILS pfnFmt, so it will just start a new one. Also,
            // since wbr is in the story, we can just substitute it, and don't need
            // to create a synthetic char.
            //
            // BUGBUG (mikejoch) This will cause grief if the NOBR is overlapped with
            // a reverse object. In fact there are numerous problems mixing NOBRs
            // with reverse objects; this will need to be addressed separately.
            // FUTURE (mikejoch) It would be a good idea to actually write a real word
            // breaking function for NOBRs instead of terminating and restarting the
            // object like this. Doing so would get rid of this elseif, reduce the
            // synthetic store, and generally be clearer.
            if(pElement->Tag() == ETAG_WBR)
            {
                _lineFlags.AddLineFlag(cp, FLAG_HAS_EMBED_OR_WBR);
            }
            por->FillSynthData(SYNTHTYPE_ENDNOBR);
        }
        else
        {
            // A normal phrase element start or end run. Just make
            // it antisynth so that it is hidden from LS
            por->MakeRunAntiSynthetic();
        }

        if(!por->IsAntiSyntheticRun())
        {
            // Empty lines will need some height!
            if(pElement->_fBreakOnEmpty)
            {
                if(IsFirstNonWhiteOnLine(cp))
                {
                    // We provide a line height only if something in our whitespace
                    // is not providing some visual representation of height. So if our
                    // whitespace was either aligned or abspos'd then we do NOT provide
                    // any height since these sites provide some height of their own.
                    if(_lineCounts.GetLineCount(por->Cp(), LC_ALIGNEDSITES)==0
                        && _lineCounts.GetLineCount(por->Cp(), LC_ABSOLUTESITES)==0)
                    {
                        CCcs* pccs = GetCcs(por, _pci->_hdc, _pci);
                        if(pccs)
                        {
                            RecalcLineHeight(pccs, &_li);
                        }
                        _fLineWithHeight = TRUE;
                    }
                }
            }

            // If we have already decided to give the line a height then we want
            // to get the text metrics else we do not want the text metrics. The
            // reasons are explained in the blurb below.
            if(!_fLineWithHeight)
            {
                // If we have not anti-synth'd the run, it means that we have terminated
                // the line -- either due to a block element or due to a BR element. In either
                // of these 2 cases, if the element did not have break on empty, then we
                // do not want any of them to induce a descent. If it did have a break
                // on empty then we have already computed the heights, so there is no
                // need to do so again.
                por->_fNoTextMetrics = TRUE;
            }
        }

        if(pCF->IsRelative(por->_fInnerCF))
        {
            _lineFlags.AddLineFlag(cp, FLAG_HAS_RELATIVE);
        }

        // Set the flag on the line if the current pcf has a background
        // image or color
        if(pCF->HasBgImage(por->_fInnerCF) || pCF->HasBgColor(por->_fInnerCF))
        {
            // NOTE(SujalP): If _cpStart has a background, and nothing else ends up
            // on the line, then too we want to draw the background. But since the
            // line is empty cpMost == _cpStart and hence GetLineFlags will not
            // find this flag. To make sure that it does, we subtract 1 from the cp.
            // (Bug  43714).
            (cp==_cpStart) ? _lineFlags.AddLineFlagForce(cp-1, FLAG_HAS_BACKGROUND)
                : _lineFlags.AddLineFlag(cp, FLAG_HAS_BACKGROUND);
        }

        if(pCF->_fBidiEmbed && _pBidiLine==NULL && !IsFrozen())
        {
            _pBidiLine = new CBidiLine(_treeInfo, _cpStart, _li._fRTL, _pli);
            Assert(GetDirLevel(por->_lscpBase) == 0);
        }

        // Some cases here which we need to let thru for the orther glean code
        // to look at........
        goto Done;
    }

    // Figure out CHP, the layout info.
    pCF = IsAdornment() ? _pCFLi : por->GetCF();
    Assert(pCF);

    if(pCF->IsVisibilityHidden())
    {
        _lineFlags.AddLineFlag(cp, FLAG_HAS_NOBLAST);
    }

    // If we've transitioned directions, begin or end a reverse object.
    if(_pBidiLine!=NULL &&
        (nDirLevel=_pBidiLine->GetLevel(cp))!=(nSynthDirLevel=GetDirLevel(por->_lscpBase)))
    {
        if(!IsFrozen())
        {
            // Determine the type of synthetic character to add.
            if(nDirLevel > nSynthDirLevel)
            {
                synthCur = SYNTHTYPE_REVERSE;
            }
            else
            {
                synthCur = SYNTHTYPE_ENDREVERSE;
            }

            // Add the new synthetic character.
            lserr = AppendILSControlChar(por, synthCur, &porOut);
            Assert(lserr!=lserrNone || porOut->_synthType!=SYNTHTYPE_NONE);
            goto Cleanup;
        }
    }
    else if(pCF->_bSpecialObjectFlagsVar != _bSpecialObjectFlagsVar)
    {
        if(!IsFrozen())
        {
            if(CheckForSpecialObjectBoundaries(por, &porOut))
            {
                goto Cleanup;
            }
        }
    }

    CHPFromCF( por, pCF );

    por->_brkopt = (pCF->_fLineBreakStrict?fBrkStrict:0) | (pCF->_fNarrow?0:fCscWide);

    // Note the relative stuff
    if(!por->_fCharsForNestedLayout && pCF->IsRelative(por->_fInnerCF))
    {
        _lineFlags.AddLineFlag(cp, FLAG_HAS_RELATIVE);
    }

    // Set the flag on the line if the current pcf has a background
    // image or color
    if(pCF->HasBgImage(por->_fInnerCF) || pCF->HasBgColor(por->_fInnerCF))
    {
        _lineFlags.AddLineFlag(cp, FLAG_HAS_BACKGROUND);
    }

    if(!por->_fCharsForNestedLayout)
    {
        const TCHAR chFirst = por->_pchBase[0];

        // Currently the only nested elements we have other than layouts are hidden
        // elements. These are taken care of before we get here, so we should
        // never be here with this flag on.
        Assert(!por->_fCharsForNestedElement);

        // A regular text run.

        // Arye: Might be in edit mode on last empty line and trying
        // to measure. Need to fail gracefully.
        // AssertSz( *ppwchRun, "Expected some text.");

        // Begin nasty exception code to deal with all the possible
        // special characters in the run.
        Assert(pCF != NULL);
        if(_pBidiLine==NULL && (IsRTLChar(chFirst) || pCF->_fBidiEmbed))
        {
            if(!IsFrozen())
            {
                _pBidiLine = new CBidiLine(_treeInfo, _cpStart, _li._fRTL, _pli);
                Assert(GetDirLevel(por->_lscpBase) == 0);

                if(_pBidiLine!=NULL && _pBidiLine->GetLevel(cp)>0)
                {
                    synthCur = SYNTHTYPE_REVERSE;
                    // Add the new synthetic character.
                    lserr = AppendILSControlChar(por, synthCur, &porOut);
                    Assert(lserr!=lserrNone || porOut->_synthType!=SYNTHTYPE_NONE);
                    goto Cleanup;
                }
            }
        }

        // Don't blast disabled lines.
        if(pCF->_fDisabled)
        {
            _lineFlags.AddLineFlag(cp, FLAG_HAS_NOBLAST);
        }

        // NB (cthrash) If an NBSP character is at the beginning of a
        // text run, LS will convert that to a space before calling
        // GetRunCharWidths.  This implies that we will fail to recognize
        // the presence of NBSP chars if we only check at GRCW. So an
        // additional check is required here.  Similarly, if our 
        if(chFirst == WCH_NBSP)
        {
            _lineFlags.AddLineFlag(cp, FLAG_HAS_NBSP);
        }

        lserr = ChunkifyTextRun(por, &porOut);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        if(porOut != por)
        {
            goto Cleanup;
        }
    }
    else
    {
        // It has to be some layout other than the layout we are measuring
        Assert(por->Branch()->GetUpdatedLayout()
            && por->Branch()->GetUpdatedLayout()!=_pFlowLayout);

#ifdef _DEBUG
        CElement* pElementLayout = por->Branch()->GetUpdatedLayout()->ElementOwner();
        long cpElemStart = pElementLayout->GetFirstCp();

        // Count the characters in this site, so LS can skip over them on this line.
        Assert(por->_lscch == GetNestedElementCch(pElementLayout));
#endif

        _fHasSites = _fMinMaxPass;

        // We check if this site belongs on its own line.
        // If so, we terminate this line with an EOS marker.
        if(IsOwnLineSite(por))
        {
            // This guy belongs on his own line.  But we only have to terminate the
            // current line if he's not the first thing on this line.
            if(cp != _cpStart)
            {
                Assert(!por->_fHidden);

                // We're not first on line.  Terminate this line!
                lserr = TerminateLine(por, TL_ADDEOS, &porOut);
                Assert(lserr!=lserrNone || porOut->_synthType!=SYNTHTYPE_NONE);
                goto Cleanup;
            }
            // else we are first on line, so even though this guy needs to be
            // on his own line, keep going, because he is!
        }
        // If we kept looking after a single line site in _fScanForCR and we
        // came here, it means that we have another site after the single site and
        // hence should terminate the line
        else if(_fScanForCR && _fSingleSite)
        {
            lserr = TerminateLine(por, TL_ADDEOS, &porOut);
            goto Cleanup;
        }

        // Whatever this is, it is under a different site, so we have
        // to give LS an embedded object notice, and later recurse back
        // to format this other site.  For tables and spans and such, we have
        // to count and skip over the characters in this site.
        por->_lsCharProps.idObj = LSOBJID_EMBEDDED;

        Assert(cp == cpElemStart-1);

        // ppwchRun shouldn't matter for a LSOBJID_EMBEDDED, but chunkify
        // objectrun might modify for putting in the symbols for aligned and
        // abspos'd sites
        fWasTextRun = FALSE;
        lserr = ChunkifyObjectRun(por, &porOut);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        por = porOut;
    }

    lserr = SetRenderingHighlights(por);
    if(lserr != lserrNone)
    {
        goto Cleanup;
    }

    if(fWasTextRun && !por->_fHidden)
    {
        lserr = CheckForUnderLine(por);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        if(_chPassword)
        {
            lserr = CheckForPassword(por);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }
        }
        else
        {
            lserr = CheckForTransform(por);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }
        }
    }

Cleanup:
    // Some characters we don't want contributing to the line height.
    // Non-textual runs shouldn't, don't bother for hidden runs either.
    //
    // ALSO, we want text metrics if we had a BR or block element
    // which was not break on empty
    porOut->_fNoTextMetrics |= porOut->_fHidden || porOut->_lsCharProps.idObj!=LSOBJID_TEXT;

Done:
    *pporOut = porOut;
    return lserr;
}

BOOL IsPreLikeNode(CTreeNode* pNode)
{
    ELEMENT_TAG eTag = pNode->Tag();
    return (eTag==ETAG_PRE || eTag==ETAG_XMP || eTag==ETAG_PLAINTEXT || eTag==ETAG_LISTING);
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckForSpecialObjectBoundaries
//
//  Synopsis:   This function checks to see whether special objects should
//              be opened or closed based on the CF properties
//
//  Returns:    The one run to submit to line services via the pporOut
//              parameter.  Will only change this if necessary, and will
//              return TRUE if that parameter changed.
//
//-----------------------------------------------------------------------------
BOOL WINAPI CLineServices::CheckForSpecialObjectBoundaries(COneRun* por, COneRun** pporOut)
{
    BOOL fRet = FALSE;
    LSERR lserr;
    const CCharFormat* pCF = por->GetCF();
    Assert(pCF);

    if(pCF->_fIsRuby && !_fIsRuby)
    {
        Assert(!_fIsRubyText);

        // Open up a new ruby here (we only enter here in the open ruby case,
        // the ruby object is closed when we see an /RT)
        _fIsRuby = TRUE;
#ifdef RUBY_OVERHANG
        // We set this flag here so that LS will try to do modify the width of
        // run, which will trigger a call to FetchRubyWidthAdjust
        por->_lsCharProps.fModWidthOnRun = TRUE;
#endif
        _yMaxHeightForRubyBase = 0;
        _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_NOBLAST);
        _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_RUBY);
        lserr = AppendILSControlChar(por, SYNTHTYPE_RUBYMAIN, pporOut);
        Assert(lserr!=lserrNone || (*pporOut)->_synthType!=SYNTHTYPE_NONE);
        fRet = TRUE;
    }
    else if(pCF->_fIsRubyText != _fIsRubyText)
    {
        Assert(_fIsRuby);

        // if _fIsRubyText is true, that means we have now arrived at text that
        // is no longer Ruby Text.  So, we should close the Ruby object by passing
        // ENDRUBYTEXT to Line Services
        if(_fIsRubyText)
        {
            _fIsRubyText = FALSE;
            _fIsRuby = FALSE;
            lserr = AppendILSControlChar(por, SYNTHTYPE_ENDRUBYTEXT, pporOut);
            Assert(lserr!=lserrNone || (*pporOut)->_synthType!=SYNTHTYPE_NONE);
            fRet = TRUE;
        }
        // if _fIsRubyText is false, that means that we are now entering text that
        // is Ruby text.  So, we must tell Line Services that we are no longer
        // giving it main text.
        else
        {
            _fIsRubyText = TRUE;
            lserr = AppendILSControlChar(por, SYNTHTYPE_ENDRUBYMAIN, pporOut);
            Assert(lserr!=lserrNone || (*pporOut)->_synthType!=SYNTHTYPE_NONE);
            fRet = TRUE;
        }
    }
    else
    {
        BOOL fNoBreak = pCF->HasNoBreak(por->_fInnerCF);
        if(fNoBreak != !!_fNoBreakForMeasurer)
        {
            Assert(!IsFrozen());

            // phrase elements inside PRE's which have layout will not have the HasPre bit turned
            // on and hence we will still start a NOBR object for them. The problem then is that
            // we need to terminate the line for '\r'. To do this we will have to scan the text.
            // To minimize the scanning we find out if we are really in such a situation.
            if(_pMarkup->SearchBranchForCriteriaInStory(por->Branch(), IsPreLikeNode))
            {
                _fScanForCR = TRUE;
                goto Cleanup;
            }

            if(!IsOwnLineSite(por))
            {
                // Begin or end NOBR block.
                // BUGBUG (paulpark) this won't work with nearby-nested bidi's and nobr's because
                // the por's we cache in _arySynth are not accurately postioned.
                lserr = AppendILSControlChar(por,
                    (fNoBreak?SYNTHTYPE_NOBR:SYNTHTYPE_ENDNOBR), pporOut);
                Assert(lserr!=lserrNone || (*pporOut)->_synthType!=SYNTHTYPE_NONE);
                fRet = TRUE;
            }
        }
    }

Cleanup:    
    return fRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetAutoNumberInfo (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetAutoNumberInfo(
                                 LSKALIGN* plskalAnm,    // OUT
                                 PLSCHP plsChpAnm,       // OUT
                                 PLSRUN* pplsrunAnm,     // OUT                                  
                                 WCHAR* pwchAdd,         // OUT
                                 PLSCHP plsChpWch,       // OUT                                 
                                 PLSRUN* pplsrunWch,     // OUT                                  
                                 BOOL* pfWord95Model,    // OUT
                                 long* pduaSpaceAnm,     // OUT
                                 long* pduaWidthAnm)     // OUT
{
    LSTRACE(GetAutoNumberInfo);
    LSNOTIMPL(GetAutoNumberInfo);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetNumericSeparators (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetNumericSeparators(
                                    PLSRUN plsrun,          // IN
                                    WCHAR* pwchDecimal,     // OUT
                                    WCHAR* pwchThousands)   // OUT
{
    LSTRACE(GetNumericSeparators);

    // BUGBUG (cthrash) Should set based on locale.
    *pwchDecimal = L'.';
    *pwchThousands = L',';

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckForDigit (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::CheckForDigit(
                             PLSRUN plsrun,      // IN
                             WCHAR wch,          // IN
                             BOOL* pfIsDigit)    // OUT
{
    LSTRACE(CheckForDigit);

    // BUGBUG (mikejoch) IsCharDigit() doesn't check for international numbers.

    *pfIsDigit = IsCharDigit(wch);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FetchPap (member, LS callback)
//
//  Synopsis:   Callback to fetch paragraph properties for the current line.
//
//  Returns:    lserrNone
//              lserrOutOutMemory
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FetchPap(/*[in]*/LSCP lscp, /*[out]*/PLSPAP pap)
{
    LSTRACE(FetchPap);
    LSERR       lserr = lserrNone;
    const CParaFormat* pPF;
    CTreePos*   ptp;
    CTreeNode*  pNode;
    LONG        cp;
    BOOL        fInnerPF;
    CComplexRun* pcr = NULL;
    CElement*   pElementFL = _pFlowLayout->ElementContent();

    Assert(lscp <= _treeInfo._lscpFrontier);

    if(lscp < _treeInfo._lscpFrontier)
    {
        COneRun* por = FindOneRun(lscp);
        Assert(por);
        if(!por)
        {
            goto Cleanup;
        }
        ptp = por->_ptp;
        pPF = por->_pPF;
        fInnerPF = por->_fInnerPF;
        cp = por->Cp();
        pcr = por->GetComplexRun();
    }
    else
    {
        // The problem is that we are at the end of the list
        // and hence we cannot find the interesting one-run. In this
        // case we have to use the frontier information. However,
        // the frontier information maybe exhausted, so we need to
        // refresh it by calling AdvanceTreePos() here.
        if(!_treeInfo._cchRemainingInTreePos && !_treeInfo._fHasNestedElement)
        {
            if(!_treeInfo.AdvanceTreePos())
            {
                lserr = lserrOutOfMemory;
                goto Cleanup;
            }
        }
        ptp = _treeInfo._ptpFrontier;
        pPF = _treeInfo._pPF; // should we use CLineServices::_pPFFirst here???
        fInnerPF = _treeInfo._fInnerPF;
        cp  = _treeInfo._cpFrontier;
    }

    Assert(ptp);
    Assert(pPF);

    // Set up paragraph properties
    PAPFromPF(pap, pPF, fInnerPF, pcr);

    pap->cpFirst = cp;

    // BUGBUG SLOWBRANCH: GetBranch is **way** too slow to be used here.
    pNode = ptp->GetBranch();
    if(pNode->Element() == pElementFL)
    {
        pap->cpFirstContent = _treeInfo._cpLayoutFirst;
    }
    else
    {
        CTreeNode* pNodeBlock = _treeInfo._pMarkup->SearchBranchForBlockElement(pNode, _pFlowLayout);
        if (pNodeBlock)
        {
            CElement* pElementPara = pNodeBlock->Element();
            pap->cpFirstContent = (pElementPara==pElementFL) ?
                _treeInfo._cpLayoutFirst : pElementPara->GetFirstCp();
        }
        else
        {
            pap->cpFirstContent = _treeInfo._cpLayoutFirst;
        }
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   FetchTabs (member, LS callback)
//
//  Synopsis:   Callback to return tab positions for the current line.
//
//              LineServices calls the callback when it encounters a tab in
//              the line, but does not pass the plsrun.  The cp is supposed to
//              be used to locate the paragraph.
//
//              Instead of allocating a buffer for the return value, we return
//              a table that resides on the CLineServices object.  The tab
//              values are in twips.
//
//  Returns:
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FetchTabs(
                         LSCP lscp,                      // IN
                         PLSTABS plstabs,                // OUT
                         BOOL* pfHangingTab,             // OUT
                         long* pduaHangingTab,           // OUT
                         WCHAR* pwchHangingTabLeader )   // OUT
{
    LSTRACE(FetchTabs);

    LONG cTab = _pPFFirst->GetTabCount(_fInnerPFFirst);

    // This tab might end up on the current line, so we can't blast.

    _fSawATab = TRUE;

    // Note: lDefaultTab is a constant defined in textedit.h

    plstabs->duaIncrementalTab = lDefaultTab;

    // BUGBUG (cthrash) What to do for hanging tabs? CSS? Auto?

    *pfHangingTab = FALSE;
    *pduaHangingTab = 0;
    *pwchHangingTabLeader = 0;

    AssertSz(cTab>=0&&cTab<=MAX_TAB_STOPS, "illegal tab count");

    if(!_pPFFirst->HasTabStops(_fInnerPFFirst) && cTab<2)
    {
        if(cTab == 1)
        {
            plstabs->duaIncrementalTab = _pPFFirst->GetTabPos(_pPFFirst->_rgxTabs[0]);
        }

        plstabs->iTabUserDefMac = 0;
        plstabs->pTab = NULL;
    }
    else
    {
        LSTBD* plstbd = _alstbd + cTab - 1;

        while(cTab)
        {
            long uaPos;
            long lAlign;
            long lLeader;

            _pPFFirst->GetTab(--cTab, &uaPos, &lAlign, &lLeader);

            Assert(lAlign>=0 && lAlign<tomAlignBar && lLeader>=0 && lLeader<tomLines );

            // NB (cthrash) To ensure that the LSKTAB cast is safe, we
            // verify that that define's haven't changed values in
            // CLineServices::InitTimeSanityCheck().
            plstbd->lskt = LSKTAB(lAlign);
            plstbd->ua = uaPos;
            plstbd->wchTabLeader = s_achTabLeader[lLeader];
            plstbd--;
        }

        plstabs->iTabUserDefMac = cTab;
        plstabs->pTab = _alstbd;
    }

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetBreakThroughTab (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetBreakThroughTab(
                                  long uaRightMargin,         // IN
                                  long uaTabPos,              // IN
                                  long* puaRightMarginNew)    // OUT
{
    LSTRACE(GetBreakThroughTab);
    LSNOTIMPL(GetBreakThroughTab);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckParaBoundaries (member, LS callback)
//
//  Synopsis:   Callback to determine whether two cp's reside in different
//              paragraphs (block elements in HTML terms).
//
//  Returns:    lserrNone
//              *pfChanged - TRUE if cp's are in different block elements.
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::CheckParaBoundaries(
                                   LSCP lscpOld,       // IN
                                   LSCP lscpNew,       // IN
                                   BOOL* pfChanged)    // OUT
{
    LSTRACE(CheckParaBoundaries);
    *pfChanged = FALSE;
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetRunCharWidths (member, LS callback)
//
//  Synopsis:   Callback to return character widths of text in the current run,
//              represented by plsrun.
//
//  Returns:    lserrNone
//              rgDu - array of character widths
//              pduDu - sum of widths in rgDu, upto *plimDu characters
//              plimDu - character count in rgDu
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetRunCharWidths(
                                PLSRUN plsrun,              // IN
                                LSDEVICE lsDeviceID,        // IN
                                LPCWSTR pwchRun,            // IN
                                DWORD cwchRun,              // IN
                                long du,                    // IN
                                LSTFLOW kTFlow,             // IN
                                int* rgDu,                  // OUT
                                long* pduDu,                // OUT
                                long* plimDu)               // OUT
{
    LSTRACE(GetRunCharWidths);

    LSERR lserr = lserrNone;
    pwchRun = plsrun->_fMakeItASpace ? _T(" ") : pwchRun;
    const WCHAR* pch = pwchRun;
    int* pdu = rgDu;
    int* pduEnd = rgDu + cwchRun;
    long duCumulative = 0;
    LONG cpCurr = plsrun->Cp();
    const CCharFormat* pCF = plsrun->GetCF();
    CCcs* pccs;
    CBaseCcs* pBaseCcs;
    HDC hdc;
    LONG duChar = 0;

    Assert( cwchRun );

    pccs = GetCcs(plsrun, _pci->_hdc, _pci);
    if(!pccs)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }
    hdc = pccs->GetHDC();
    pBaseCcs = pccs->GetBaseCcs();

    // CSS
    Assert(IsAdornment() || !plsrun->_fCharsForNestedElement);

    // Add up the widths
    while(pdu < pduEnd)
    {
        // The tight inner loop.
        TCHAR ch = *pch++;

        ProcessOneChar(cpCurr, pdu, rgDu, ch);

        Verify(pBaseCcs->Include(hdc, ch, duChar));

        duCumulative += *pdu++ = duChar;

        if(duCumulative > du)
        {
            break;
        }
    }

    if(GetLetterSpacing(pCF) || pCF->HasCharGrid(TRUE))
    {
        lserr = GetRunCharWidthsEx(plsrun, pwchRun, cwchRun,
            du, rgDu, pduDu, plimDu);
    }
    else
    {
        *pduDu = duCumulative;
        *plimDu = pdu - rgDu;
    }

Cleanup:
    return lserr;
}

LSERR CLineServices::GetRunCharWidthsEx(
                                  PLSRUN plsrun,
                                  LPCWSTR pwchRun,
                                  DWORD cwchRun,
                                  LONG du,
                                  int* rgDu,
                                  long* pduDu,
                                  long* plimDu)
{
    LONG lserr = lserrNone;
    LONG duCumulative = 0;
    LONG cpCurr = plsrun->Cp();
    int* pdu = rgDu;
    int* pduEnd = rgDu + cwchRun;
    const CCharFormat* pCF = plsrun->GetCF();
    LONG xLetterSpacing = GetLetterSpacing(pCF);
    CCcs* pccs;

    pccs = GetCcs(plsrun, _pci->_hdc, _pci);
    if(!pccs)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }

    _lineFlags.AddLineFlagForce(cpCurr, FLAG_HAS_NOBLAST);

    if(pCF->HasCharGrid(TRUE))
    {
        long lGridSize = GetCharGridSize();

        switch(pCF->GetLayoutGridType(TRUE))
        {
        case styleLayoutGridTypeStrict:
        case styleLayoutGridTypeFixed:
            // we have strict or fixed layout grid for this run
            {
                if(plsrun->IsOneCharPerGridCell())
                {
                    for(pdu=rgDu; pdu<pduEnd;)
                    {
                        if(xLetterSpacing < 0)
                        {
                            MeasureCharacter(pwchRun[pdu-rgDu], cpCurr, rgDu, pccs, pdu);
                        }
                        *pdu += xLetterSpacing;
                        *pdu = GetClosestGridMultiple(lGridSize, *pdu);
                        duCumulative += *pdu++;
                        if(duCumulative > du)
                        {
                            break;
                        }
                    }
                }
                else
                {
                    BOOL fLastInChain = FALSE;
                    LONG duOuter = 0;

                    // Get width of this run
                    for(pdu=rgDu; pdu<pduEnd;)
                    {
                        if(xLetterSpacing < 0)
                        {
                            MeasureCharacter(pwchRun[pdu-rgDu], cpCurr, rgDu, pccs, pdu);
                        }
                        *pdu += xLetterSpacing;
                        duCumulative += *pdu++;
                        if(duCumulative > du)
                        {
                            duOuter = *(pdu-1);
                            duCumulative -= duOuter;
                            fLastInChain = TRUE;
                            break;
                        }
                    }

                    if(!fLastInChain)
                    {
                        // Fetch next run
                        COneRun* pNextRun = 0;
                        LPCWSTR  lpcwstrJunk;
                        DWORD    dwJunk;
                        BOOL     fJunk;
                        LSCHP    lschpJunk;
                        FetchRun(plsrun->_lscpBase + plsrun->_lscch, &lpcwstrJunk, &dwJunk, &fJunk, &lschpJunk, 
                            &pNextRun);
                        fLastInChain = !plsrun->IsSameGridClass(pNextRun);
                    }
                    if(!fLastInChain)
                    {   // we are inside a runs chain with same grid properties
                        plsrun->_xWidth = duCumulative;
                    }
                    else
                    {   // we are at the end of runs chain with same grid properties

                        // Get width of runs chain
                        LONG duChainWidth = duCumulative;
                        COneRun* pFirstRun = plsrun;
                        while(pFirstRun->_pPrev)
                        {
                            while(pFirstRun->_pPrev && !pFirstRun->_pPrev->_ptp->IsText())
                            {
                                pFirstRun = pFirstRun->_pPrev;
                            }
                            if(plsrun->IsSameGridClass(pFirstRun->_pPrev))
                            {
                                duChainWidth += pFirstRun->_pPrev->_xWidth;
                                pFirstRun = pFirstRun->_pPrev;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // Update width and draw offset of the chain
                        LONG duGrid = GetClosestGridMultiple(lGridSize, duChainWidth);
                        LONG lOffset = (duGrid - duChainWidth) / 2;
                        plsrun->_xWidth = duCumulative + duGrid - duChainWidth - lOffset;
                        plsrun->_xDrawOffset = lOffset;
                        COneRun* pPrevRun = plsrun->_pPrev;
                        while(pPrevRun)
                        {
                            while(pPrevRun && !pPrevRun->_ptp->IsText())
                            {
                                pPrevRun = pPrevRun->_pPrev;
                            }
                            if(plsrun->IsSameGridClass(pPrevRun))
                            {
                                pPrevRun->_xDrawOffset = lOffset;
                                pPrevRun = pPrevRun->_pPrev;
                            }
                            else
                            {
                                break;
                            }
                        }
                        duCumulative += duOuter + duGrid - duChainWidth;
                    }
                }
            }
            break;

        case styleLayoutGridTypeLoose:
        default:
            // we have loose layout grid or layout grid type is not set for this run
            for(pdu=rgDu; pdu<pduEnd;)
            {
                if(xLetterSpacing < 0)
                {
                    MeasureCharacter(pwchRun[pdu-rgDu], cpCurr, rgDu, pccs, pdu);
                }
                *pdu += xLetterSpacing;
                *pdu += LooseTypeWidthIncrement(pwchRun[pdu-rgDu], (plsrun->_brkopt==fCscWide), lGridSize);
                duCumulative += *pdu++;
                if(duCumulative > du)
                {
                    break;
                }
            }
        }
    }
    else
    {
        // But we must be careful to deal with negative letter spacing,
        // since that will cause us to have to actually get some character
        // widths since the line length will be extended.  If the
        // letterspacing is positive, then we know all the letters we're
        // gonna use have already been measured.
        for(pdu=rgDu; pdu<pduEnd;)
        {
            if(xLetterSpacing < 0)
            {
                MeasureCharacter(pwchRun[pdu-rgDu], cpCurr, rgDu, pccs, pdu);
            }
            *pdu += xLetterSpacing;
            duCumulative += *pdu++;
            if(duCumulative > du)
            {
                break;
            }
        }

        if(xLetterSpacing && pdu>=pduEnd && !_pCFLi)
        {
            LPCWSTR  lpcwstrJunk;
            DWORD    dwJunk;
            BOOL     fJunk;
            LSCHP    lschpJunk;
            COneRun* por;
            int      xLetterSpacingNextRun = 0;

            lserr = FetchRun(plsrun->_lscpBase+plsrun->_lscch,
                &lpcwstrJunk, &dwJunk, &fJunk, &lschpJunk, &por);

            if(lserr==lserrNone && por->_ptp!=_treeInfo._ptpLayoutLast)
            {
                xLetterSpacingNextRun = GetLetterSpacing(por->GetCF());
            }
            if(xLetterSpacing != xLetterSpacingNextRun)
            {
                *(pdu-1) -= xLetterSpacing;
                duCumulative -= xLetterSpacing;
            }
        }
    }

    *pduDu = duCumulative;
    *plimDu = pdu - rgDu;

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetLineOrCharGridSize()
//
//  Synopsis:   This function finds out what the height or width of a grid cell 
//              is in pixels, making conversions if necessary.  If this value has 
//              already been calculated, the calculated value is immediately returned
//
//  Returns:    long value of the width of a grid cell in pixels
//
//-----------------------------------------------------------------------------
long WINAPI CLineServices::GetLineOrCharGridSize(BOOL fGetCharGridSize)
{
    const CUnitValue* pcuv;
    CUnitValue::DIRECTION dir;
    long* lGridSize = fGetCharGridSize ? &_lCharGridSize : &_lLineGridSize;

    // If we already have a cached value, return that, or if we haven't set
    // set up the ccs yet return zero
    if(*lGridSize!=0 || !_pPFFirst)
    {
        goto Cleanup;
    }

    pcuv = fGetCharGridSize ? &(_pPFFirst->GetCharGridSize(TRUE)) : &(_pPFFirst->GetLineGridSize(TRUE));
    // The uv should have some value, otherwise we shouldn't even be
    // here making calculations for a non-existent grid.
    switch(pcuv->GetUnitType())
    {
    case CUnitValue::UNIT_NULLVALUE:
        break;

    case CUnitValue::UNIT_ENUM:
        // need to handle "auto" here
        if(pcuv->GetUnitValue() == styleLayoutGridCharAuto && _pccsCache) 
        {
            *lGridSize = fGetCharGridSize ? _pccsCache->GetBaseCcs()->_xMaxCharWidth :
                _pccsCache->GetBaseCcs()->_yHeight;
        }
        else
        {
            *lGridSize = 0;
        }
        break;

    case CUnitValue::UNIT_PERCENT:
        *lGridSize = ((fGetCharGridSize?_xWrappingWidth:_pci->_sizeParent.cy) * pcuv->GetUnitValue()) / 10000;
        break;

    default:
        dir = fGetCharGridSize ? CUnitValue::DIRECTION_CX : CUnitValue::DIRECTION_CY;
        *lGridSize = pcuv->GetPixelValue(NULL, dir, fGetCharGridSize?_xWrappingWidth:_pci->_sizeParent.cy);
        break;
    }

    if(pcuv->IsPercent() && _pFlowLayout && !fGetCharGridSize)
    {
        _pFlowLayout->GetDisplay()->SetVertPercentAttrInfo();
    }
Cleanup:
    return *lGridSize;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetClosestGridMultiple
//
//  Synopsis:   This function just calculates the width of an object in lGridSize
//              multiples.  For example, if lGridSize is 12 and lObjSize is 16,
//              this function would return 24.  If lObjSize is 0, this function
//              will return 0.
//              
//  Returns:    long value of the width in pixels
//
//-----------------------------------------------------------------------------
long WINAPI CLineServices::GetClosestGridMultiple(long lGridSize, long lObjSize)
{
    long lReturnWidth = lObjSize;
    long lRemainder;
    if(lObjSize==0 || lGridSize==0)
    {
        goto Cleanup;
    }

    lRemainder = lObjSize % lGridSize;
    lReturnWidth = lObjSize + lGridSize - (lRemainder?lRemainder:lGridSize);

Cleanup:
    return lReturnWidth;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckRunKernability (member, LS callback)
//
//  Synopsis:   Callback to test whether current runs should be kerned.
//
//              We do not support kerning at this time.
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::CheckRunKernability(
                                   PLSRUN plsrunLeft,  // IN
                                   PLSRUN plsrunRight, // IN
                                   BOOL* pfKernable)   // OUT
{
    LSTRACE(CheckRunKernability);

    *pfKernable = FALSE;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetRunCharKerning (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetRunCharKerning(
                                 PLSRUN plsrun,              // IN
                                 LSDEVICE lsDeviceID,        // IN
                                 LPCWSTR pwchRun,            // IN
                                 DWORD cwchRun,              // IN
                                 LSTFLOW kTFlow,             // IN
                                 int* rgDu)                  // OUT
{
    LSTRACE(GetRunCharKerning);

    DWORD iwch = cwchRun;
    int* pDu = rgDu;

    while(iwch--)
    {
        *pDu++ = 0;
    }

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetRunTextMetrics (member, LS callback)
//
//  Synopsis:   Callback to return text metrics of the current run
//
//  Returns:    lserrNone
//              plsTxMet - LineServices textmetric structure (lstxm.h)
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetRunTextMetrics(
                                 PLSRUN   plsrun,            // IN
                                 LSDEVICE lsDeviceID,        // IN
                                 LSTFLOW  kTFlow,            // IN
                                 PLSTXM   plsTxMet)          // OUT
{
    LSTRACE(GetRunTextMetrics);
    const CCharFormat* pCF;
    CCcs* pccs;
    CBaseCcs* pBaseCcs;
    LONG lLineHeight;
    LSERR lserr = lserrNone;

    Assert(plsrun);

    if(plsrun->_fNoTextMetrics)
    {
        ZeroMemory(plsTxMet, sizeof(LSTXM));
        goto Cleanup;
    }

    pCF = plsrun->GetCF();
    pccs = GetCcs(plsrun, _pci->_hdc, _pci);
    if(!pccs)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }
    pBaseCcs = pccs->GetBaseCcs();

    // Cache this in case we do width modification.
    plsTxMet->fMonospaced = pBaseCcs->_fFixPitchFont ? TRUE : FALSE;
    _fHasOverhang |= ((plsrun->_xOverhang=pBaseCcs->_xOverhang) != 0);

    // Keep track of the line heights specified in all the
    // runs so that we can adjust the line height at the end.
    // Note that we don't want to include break character NEVER
    // count towards height.
    lLineHeight = RememberLineHeight(plsrun->Cp(), pCF, pBaseCcs);

    if(_fHasSites)
    {
        Assert(_fMinMaxPass);
        plsTxMet->dvAscent = 1;
        plsTxMet->dvDescent = 0;
        plsTxMet->dvMultiLineHeight = 1;
    }
    else
    {
        long dvAscent, dvDescent;

        dvDescent = long(pBaseCcs->_yDescent) - pBaseCcs->_yOffset;
        dvAscent  = pBaseCcs->_yHeight - dvDescent;

        plsTxMet->dvAscent = dvAscent;
        plsTxMet->dvDescent = dvDescent;
        plsTxMet->dvMultiLineHeight = lLineHeight;
    }

    if(pBaseCcs->_xOverhangAdjust)
    {
        _lineFlags.AddLineFlag(plsrun->Cp(), FLAG_HAS_NOBLAST);
    }

Cleanup:    
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetRunUnderlineInfo (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//  Note(SujalP): Lineservices is a bit wierd. It will *always* want to try to
//  merge runs and that to based on its own algorithm. That algorith however is
//  not the one which IE40 implements. For underlining, IE 40 always has a
//  single underline when we have mixed font sizes. The problem however is that
//  this underline is too far away from the smaller pt text in the mixed size
//  line (however within the dimensions of the line). When we give this to LS,
//  it thinks that the UL is outside the rect of the small character and deems
//  it incorrect and does not call us for a callback. To overcome this problem
//  we tell LS that the UL is a 0 (baseline) but remember the distance ourselves
//  in the PLSRUN.
//
//  Also, since color of the underline can change from run-to-run, we
//  return different underline types to LS so as to prevent it from
//  merging such runs. This also helps avoid merging when we are drawing overlines.
//  Overlines are drawn at different heigths (unlinke underlines) from pt
//  size to pt size. (This probably is a bug -- but this is what IE40 does
//  so lets just go with that for now).
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetRunUnderlineInfo(
                                   PLSRUN plsrun,          // IN
                                   PCHEIGHTS heightsPres,  // IN
                                   LSTFLOW kTFlow,         // IN
                                   PLSULINFO plsUlInfo)    // OUT
{
    LSTRACE(GetRunUnderlineInfo);
    BYTE bUnderlineType;
    static BOOL s_fToggleSwitch = FALSE;

    if (!plsrun->_dwImeHighlight)
    {
        bUnderlineType = CFU_CF1UNDERLINE;
    }
    else
    {
        // BUGBUG (cthrash) We need to switch between dotted and solid
        // underlining when the text goes from unconverted to converted.
        bUnderlineType = plsrun->_fUnderlineForIME ? CFU_UNDERLINEDOTTED : 0;
    }

    plsUlInfo->kulbase = bUnderlineType | (s_fToggleSwitch?CFU_SWITCHSTYLE:0);
    s_fToggleSwitch = !s_fToggleSwitch;

    plsUlInfo->cNumberOfLines = 1;
    plsUlInfo->dvpUnderlineOriginOffset = 0;
    plsUlInfo->dvpFirstUnderlineOffset = 0;
    plsUlInfo->dvpFirstUnderlineSize = 1;
    plsUlInfo->dvpGapBetweenLines = 0;
    plsUlInfo->dvpSecondUnderlineSize = 0;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetRunStrikethroughInfo (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetRunStrikethroughInfo(
                                       PLSRUN plsrun,          // IN
                                       PCHEIGHTS heightPres,   // IN
                                       LSTFLOW kTFlow,         // IN
                                       PLSSTINFO plsStInfo)    // OUT
{
    LSTRACE(GetRunStrikethroughInfo);
    LSNOTIMPL(GetRunStrikethroughInfo);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetBorderInfo (member, LS callback)
//
//  Synopsis:   Not implemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetBorderInfo(
                             PLSRUN plsrun,      // IN
                             LSTFLOW ktFlow,     // IN
                             long* pdurBorder,   // OUT
                             long* pdupBorder)   // OUT
{
    LSTRACE(GetBorderInfo);

    // This should only ever be called if we set the fBorder flag in lschp.
    LSNOTIMPL(GetBorderInfo);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   ReleaseRun (member, LS callback)
//
//  Synopsis:   Callback to release plsrun object, which we don't do.  We have
//              a cache of COneRuns which we keep (and grow) for the lifetime
//              of the CLineServices object.
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::ReleaseRun(/*[in]*/PLSRUN plsrun)
{
    LSTRACE(ReleaseRun);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   Hyphenate (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::Hyphenate(
                         PCLSHYPH plsHyphLast,   // IN
                         LSCP cpBeginWord,       // IN
                         LSCP cpExceed,          // IN
                         PLSHYPH plsHyph)        // OUT
{
    LSTRACE(Hyphenate);
    // FUTURE (mikejoch) Need to adjust cp values if kysr != kysrNil.
    plsHyph->kysr = kysrNil;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetHyphenInfo (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetHyphenInfo(
                             PLSRUN plsrun,      // IN
                             DWORD* pkysr,       // OUT
                             WCHAR* pwchKysr)    // OUT
{
    LSTRACE(GetHyphenInfo);

    *pkysr = kysrNil;
    *pwchKysr = WCH_NULL;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FInterruptUnderline (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FInterruptUnderline(
                                   PLSRUN plsrunFirst,         // IN
                                   LSCP cpLastFirst,           // IN
                                   PLSRUN plsrunSecond,        // IN
                                   LSCP cpStartSecond,         // IN
                                   BOOL* pfInterruptUnderline) // OUT
{
    LSTRACE(FInterruptUnderline);
    // FUTURE (mikejoch) Need to adjust cp values if we ever interrupt underlining.

    *pfInterruptUnderline = FALSE;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FInterruptShade (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FInterruptShade(
                               PLSRUN plsrunFirst,         // IN
                               PLSRUN plsrunSecond,        // IN
                               BOOL* pfInterruptShade)     // OUT
{
    LSTRACE(FInterruptShade);

    *pfInterruptShade = TRUE;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FInterruptBorder (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FInterruptBorder(
                                PLSRUN plsrunFirst,         // IN
                                PLSRUN plsrunSecond,        // IN
                                BOOL* pfInterruptBorder)    // OUT
{
    LSTRACE(FInterruptBorder);

    *pfInterruptBorder = FALSE;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FInterruptShaping (member, LS callback)
//
//  Synopsis:   We compare CF between the runs to see if they are different
//              enough to cause an interrup in shaping between the runs
//
//  Arguments:  kTFlow              text direction and orientation
//              plsrunFirst         run pointer for the previous run
//              plsrunSecond        run pointer for the current run
//              pfInterruptShaping  TRUE - Interrupt shaping
//                                  FALSE - Don't interrupt shaping, merge runs
//
//  Returns:    LSERR               lserrNone if succesful
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FInterruptShaping(
                                 LSTFLOW kTFlow,                 // IN
                                 PLSRUN plsrunFirst,             // IN
                                 PLSRUN plsrunSecond,            // IN
                                 BOOL* pfInterruptShaping)       // OUT
{
    LSTRACE(FInterruptShaping);

    Assert(pfInterruptShaping!=NULL && plsrunFirst!=NULL && plsrunSecond!=NULL);

    CComplexRun* pcr1 = plsrunFirst->GetComplexRun();
    CComplexRun* pcr2 = plsrunSecond->GetComplexRun();

    Assert(pcr1!=NULL && pcr2!=NULL);

    *pfInterruptShaping = (pcr1->GetScript() != pcr2->GetScript());

    if(!*pfInterruptShaping)
    {
        const CCharFormat* pCF1 = plsrunFirst->GetCF();
        const CCharFormat* pCF2 = plsrunSecond->GetCF();

        Assert(pCF1!=NULL && pCF2!=NULL);

        // We want to break the shaping if the formats are not similar format
        *pfInterruptShaping = !pCF1->CompareForLikeFormat(pCF2);
    }

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetGlyphs (member, LS callback)
//
//  Synopsis:   Returns an index of glyph ids for the run passed in
//
//  Arguments:  plsrun              pointer to the run
//              pwch                string of character codes
//              cwch                number of characters in pwch
//              kTFlow              text direction and orientation
//              rgGmap              map of glyph info parallel to pwch
//              prgGindex           array of output glyph indices
//              prgGprop            array of output glyph properties
//              pcgindex            number of glyph indices
//
//  Returns:    LSERR               lserrNone if succesful
//                                  lserrInvalidRun if failure
//                                  lserrOutOfMemory if memory alloc fails
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetGlyphs(
                         PLSRUN plsrun,          // IN
                         LPCWSTR pwch,           // IN
                         DWORD cwch,             // IN
                         LSTFLOW kTFlow,         // IN
                         PGMAP rgGmap,           // OUT
                         PGINDEX* prgGindex,     // OUT
                         PGPROP* prgGprop,       // OUT
                         DWORD* pcgindex)        // OUT
{
    LSTRACE(GetGlyphs);

    LSERR lserr = lserrNone;
    HRESULT hr = S_OK;
    CComplexRun* pcr;
    DWORD cMaxGly;
    HDC hdc = _pci->_hdc;
    HDC hdcShape = NULL;
    HFONT hfont;
    HFONT hfontOld = NULL;
    SCRIPT_CACHE* psc;
    WORD* pGlyphBuffer = NULL;
    WORD* pGlyph = NULL;
    SCRIPT_VISATTR* pVisAttr = NULL;
    CCcs* pccs = NULL;
    CBaseCcs* pBaseCcs = NULL;
    CString strTransformedText;
    BOOL fTriedFontLink = FALSE;

    pcr = plsrun->GetComplexRun();
    if(pcr == NULL)
    {
        Assert(FALSE);
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }

    pccs = GetCcs(plsrun, hdc, _pci);
    if(pccs == NULL)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }

    pBaseCcs = pccs->GetBaseCcs();

    hfont = pBaseCcs->_hfont;
    Assert(hfont != NULL);
    psc = pBaseCcs->GetUniscribeCache();
    Assert(psc != NULL);

    // In some fonts in some locales, NBSPs aren't rendered like spaces.
    // Under these circumstances, we need to convert NBSPs to spaces
    // before calling ScriptShape.
    // BUGBUG: Due to a bug in Bidi Win9x GDI, we can't detect that
    // old bidi fonts lack an NBSP (IE5 bug 68214). We hack around this by
    // simply always swapping the space character for the NBSP. Since this
    // only happens for glyphed runs, US perf is not impacted.
    if(_lineFlags.GetLineFlags(plsrun->Cp()+cwch) & FLAG_HAS_NBSP)
    {
        const WCHAR* pwchStop;
        WCHAR* pwch2;

        HRESULT hr = strTransformedText.Set(pwch, cwch);
        if(hr == S_OK)
        {
            pwch = strTransformedText;

            pwch2 = (WCHAR*) pwch;
            pwchStop = pwch + cwch;

            while(pwch2 < pwchStop)
            {
                if(*pwch2 == WCH_NBSP)
                {
                    *pwch2 = L' ';
                }

                pwch2++;
            }
        }
    }

    // Inflate the number of max glyphs to generate
    // A good high end guess is the number of chars plus 6% or 10,
    // whichever is greater.
    cMaxGly = cwch + max(10, (int)cwch>>4);

    Assert(cMaxGly > 0);
    pGlyphBuffer = (WORD*)NewPtr(cMaxGly*(sizeof(WORD)+sizeof(SCRIPT_VISATTR)));
    // Our memory alloc failed. No point in going on.
    if(pGlyphBuffer == NULL)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }
    pGlyph = (WORD*)pGlyphBuffer;
    pVisAttr = (SCRIPT_VISATTR*)(pGlyphBuffer + cMaxGly);

    // Repeat the shaping process until it succeds, or fails for a reason different
    // from insufficient memory, a cache fault, or failure to glyph a character using
    // the current font
    do
    {
        // If a prior ::ScriptShape() call failed because it needed the font
        // selected into the hdc, then select the font into the hdc and try
        // again.
        if(hr == E_PENDING)
        {
            // If we have a valid hdcShape, then ScriptShape() failed for an
            // unknown reason. Bail out.
            if(hdcShape != NULL)
            {
                AssertSz(FALSE, "ScriptShape() failed for an unknown reason");
                lserr = LSERRFromHR(hr);
                goto Cleanup;
            }

            // Select the current font into the hdc and set hdcShape to hdc.
            hfontOld = SelectFontEx(hdc, hfont);
            hdcShape = hdc;
        }
        // If a prior ::ScriptShape() call failed because it was unable to
        // glyph a character with the current font, swap the font around and
        // try it again.
        else if(hr == USP_E_SCRIPT_NOT_IN_FONT)
        {
            if(!fTriedFontLink)
            {
                // Unable to find the glyphs in the font. Font link to try an
                // alternate font which might work.
                fTriedFontLink = TRUE;

                // Set the sid for the complex run to match the text (instead
                // of sidDefault.
                Assert(plsrun->_ptp->IsText());
                plsrun->_sid = plsrun->_ptp->Sid();

                // Deselect the font if we selected it.
                if(hdcShape != NULL)
                {
                    Assert(hfontOld != NULL);
                    SelectFontEx(hdc, hfontOld);
                    hdcShape = NULL;
                    hfontOld = NULL;
                }

                // Get the font using the normal sid from the text to fontlink.
                pccs = GetCcs(plsrun, hdc, _pci);
                if(pccs == NULL)
                {
                    lserr = lserrOutOfMemory;
                    goto Cleanup;
                }

                pBaseCcs = pccs->GetBaseCcs();

                // Reset the hfont and psc using the new pccs.
                hfont = pBaseCcs->_hfont;
                Assert(hfont != NULL);
                psc = pBaseCcs->GetUniscribeCache();
                Assert(psc != NULL);
            }
            else
            {
                // We tried to font link but we still couldn't make it work.
                // Blow the SCRIPT_ANALYSIS away and just let GDI deal with it.
                pcr->NukeAnalysis();
            }
        }
        // If ScriptShape() failed because of insufficient buffer space,
        // resize the buffer
        else if(hr == E_OUTOFMEMORY)
        {
            WORD* pGlyphBufferT = NULL;

            // enlarge the glyph count by another 6% of run or 10, whichever is larger.
            cMaxGly += max(10, (int)cwch>>4);

            Assert(cMaxGly > 0);
            pGlyphBufferT = (WORD*)ReallocPtr(pGlyphBuffer,
                cMaxGly*(sizeof(WORD)+sizeof(SCRIPT_VISATTR)));
            if(pGlyphBufferT != NULL)
            {
                pGlyphBuffer = pGlyphBufferT;
                pGlyph = (WORD*)pGlyphBuffer;
                pVisAttr = (SCRIPT_VISATTR*)(pGlyphBuffer + cMaxGly);
            }
            else
            {
                // Memory alloc failure.
                lserr = lserrOutOfMemory;
                goto Cleanup;
            }
        }

        // Try to shape the script again
        hr = ::ScriptShape(hdcShape, psc, pwch, cwch, cMaxGly, pcr->GetAnalysis(),
            pGlyph, rgGmap, pVisAttr, (int*)pcgindex);
    } while(hr==E_PENDING || hr==USP_E_SCRIPT_NOT_IN_FONT || hr==E_OUTOFMEMORY);

    // NB (mikejoch) We shouldn't ever fail except for the OOM case. USP should
    // always be loaded, since we wouldn't get a valid eScript otherwise.
    Assert(hr==S_OK || hr==E_OUTOFMEMORY);
    lserr = LSERRFromHR(hr);

Cleanup:
    // Restore the font if we selected it
    if(hfontOld != NULL)
    {
        Assert(hdcShape != NULL);
        SelectFontEx(hdc, hfontOld);
    }

    // If LS passed us a string which was an aggregate of several runs (which
    // happens if we returned FALSE from FInterruptShaping()) then we need
    // to make sure that the same _sid is stored in each por covered by the
    // aggregate string. Normally this isn't a problem, but if we changed
    // por->_sid for font linking then it becomes necessary. We determine
    // if a change occurred by comparing plsrun->_sid to sidDefault, which
    // is the value plsrun->_sid is always set to for a glyphed run (in
    // ChunkifyTextRuns()).
    if(plsrun->_sid!=sidDefault && plsrun->_lscch<(LONG)cwch)
    {
        DWORD sidAlt = plsrun->_sid;
        COneRun* por = plsrun->_pNext;
        LONG lscchMergedRuns = cwch - plsrun->_lscch;

        while(lscchMergedRuns>0 && por)
        {
            if(por->IsNormalRun() || por->IsSyntheticRun())
            {
                por->_sid = sidAlt;
                lscchMergedRuns -= por->_lscch;
            }
            por = por->_pNext;
        }
    }

    if(lserr == lserrNone)
    {
        // Move the values from the working buffer to the output arguments
        pcr->CopyPtr(pGlyphBuffer);
        *prgGindex = pGlyph;
        *prgGprop = (WORD*)pVisAttr;
    }
    else
    {
        // free up the allocated memory on failure
        if(pGlyphBuffer != NULL)
        {
            DisposePtr(pGlyphBuffer);
        }
        *prgGindex = NULL;
        *prgGprop = NULL;
        *pcgindex = 0;
    }

    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetGlyphPositions (member, LS callback)
//
//  Synopsis:   Returns an index of glyph ids for the run passed in
//
//  Arguments:  plsrun              pointer to the run
//              lsDevice            presentation or reference
//              pwch                string of character codes
//              rgGmap              map of glyphs
//              cwch                number of characters in pwch
//              prgGindex           array of output glyph indices
//              prgGprop            array of output glyph properties
//              pcgindex            number of glyph indices
//              kTFlow              text direction and orientation
//              rgDu                array of glyph widths
//              rgGoffset           array of glyph offsets
//
//  Returns:    LSERR               lserrNone if succesful
//                                  lserrModWidthPairsNotSet if failure
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetGlyphPositions(
                                 PLSRUN plsrun,          // IN
                                 LSDEVICE lsDevice,      // IN
                                 LPWSTR pwch,            // IN
                                 PCGMAP pgmap,           // IN
                                 DWORD cwch,             // IN
                                 PCGINDEX rgGindex,      // IN
                                 PCGPROP rgGprop,        // IN
                                 DWORD cgindex,          // IN
                                 LSTFLOW kTFlow,         // IN
                                 int* rgDu,              // OUT
                                 PGOFFSET rgGoffset)     // OUT
{
    LSTRACE(GetGlyphPositions);

    LSERR lserr = lserrNone;
    HRESULT hr = S_OK;
    CComplexRun* pcr;
    HDC hdc = _pci->_hdc;
    HDC hdcPlace = NULL;
    HFONT hfont;
    HFONT hfontOld = NULL;
    SCRIPT_CACHE* psc;
    CCcs* pccs = NULL;
    CBaseCcs* pBaseCcs = NULL;

    pcr = plsrun->GetComplexRun();
    if(pcr == NULL)
    {
        Assert(FALSE);
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }

    pccs = GetCcs(plsrun, hdc, _pci);
    if(pccs == NULL)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }
    pBaseCcs = pccs->GetBaseCcs();

    hfont = pBaseCcs->_hfont;
    Assert(hfont != NULL);
    psc = pBaseCcs->GetUniscribeCache();
    Assert(psc != NULL);


    // Try to place the glyphs
    hr = ::ScriptPlace(hdcPlace, psc, rgGindex, cgindex, (SCRIPT_VISATTR*)rgGprop,
        pcr->GetAnalysis(), rgDu, rgGoffset, NULL);

    // Handle failure
    if(hr == E_PENDING)
    {
        Assert(hdcPlace == NULL);

        // Select the current font into the hdc and set hdcShape to hdc.
        hfontOld = SelectFontEx(hdc, hfont);
        hdcPlace = hdc;

        // Try again
        hr = ::ScriptPlace(hdcPlace, psc, rgGindex, cgindex, (SCRIPT_VISATTR*)rgGprop,
            pcr->GetAnalysis(), rgDu, rgGoffset, NULL);

    }

    // NB (mikejoch) We shouldn't ever fail except for the OOM case (if USP is
    // allocating the cache). USP should always be loaded, since we wouldn't
    // get a valid eScript otherwise.
    Assert(hr==S_OK || hr==E_OUTOFMEMORY);
    lserr = LSERRFromHR(hr);

Cleanup:
    // Restore the font if we selected it
    if(hfontOld != NULL)
    {
        SelectFontEx(hdc, hfontOld);
    }

    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   ResetRunContents (member, LS callback)
//
//  Synopsis:   This callback seems to be more informational.
//
//  Arguments:  plsrun              pointer to the run
//              cpFirstOld          cpFirst before shaping
//              dcpOld              dcp before shaping
//              cpFirstNew          cpFirst after shaping
//              dcpNew              dcp after shaping
//
//  Returns:    LSERR               lserrNone if succesful
//                                  lserrMismatchLineContext if failure
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::ResetRunContents(
                                PLSRUN plsrun,      // IN
                                LSCP cpFirstOld,    // IN
                                LSDCP dcpOld,       // IN
                                LSCP cpFirstNew,    // IN
                                LSDCP dcpNew)       // IN
{
    LSTRACE(ResetRunContents);
    // FUTURE (mikejoch) Need to adjust cp values if we ever implement this.
    // FUTURE (paulnel) this doesn't appear to be needed for IE. Clarification
    // is being obtained from Line Services
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetGlyphExpansionInfo (member, LS callback)
//
//  Synopsis:   This callback is used for glyph based justification
//              1. For Arabic, it handles kashida insetion
//              2. For cluster characters, (thai vietnamese) it keeps tone 
//                 marks on their base glyphs
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetGlyphExpansionInfo(
                                     PLSRUN plsrun,                  // IN
                                     LSDEVICE lsDeviceID,            // IN
                                     LPCWSTR pwch,                   // IN
                                     PCGMAP rggmap,                  // IN
                                     DWORD cwch,                     // IN
                                     PCGINDEX rgglyph,               // IN
                                     PCGPROP rgProp,                 // IN
                                     DWORD cglyph,                   // IN
                                     LSTFLOW kTFlow,                 // IN
                                     BOOL fLastTextChunkOnLine,      // IN
                                     PEXPTYPE rgExpType,             // OUT
                                     LSEXPINFO* rgexpinfo)           // OUT
{
    LSTRACE(GetGlyphExpansionInfo);

    LSERR lserr = lserrNone;
    SCRIPT_VISATTR* psva = (SCRIPT_VISATTR*)&rgProp[0];
    CComplexRun* pcr;
    const SCRIPT_PROPERTIES* psp = NULL;
    BOOL fKashida = FALSE;
    int iKashidaWidth = 1;  // width of a kashida
    int iPropL = 0;         // index to the connecting glyph left
    int iPropR = 0;         // index to the connecting glyph right
    int iBestPr = -1;       // address of the best priority in a word so far
    int iPrevBestPr = -1;   
    int iKashidaLevel = 0;
    int iBestKashidaLevel = 10;
    BYTE bBestPr = SCRIPT_JUSTIFY_NONE;
    BYTE bCurrentPr = SCRIPT_JUSTIFY_NONE;
    BYTE expType = exptNone;

    pcr = plsrun->GetComplexRun();
    UINT justifyStyle = plsrun->_pPF->_uTextJustify;

    if(pcr == NULL)
    {
        Assert(FALSE);
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }

    // 1. From the script analysis we can tell what language we are dealing with
    //    a. if we are Arabic block languages, we need to do the kashida insertion
    //    b. if we are Thai or Vietnamese, we need to set the expansion type for diacritics 
    //       to none so they remain on their base glyph.
    psp = GetScriptProperties(pcr->GetAnalysis()->eScript);

    // Check to see if we are an Arabic block language
    fKashida = IsInArabicBlock(psp->langid);

    // if we are going to do kashida insertion we need to get the kashida width information
    if(fKashida)
    {
        lserr = GetKashidaWidth(plsrun, &iKashidaWidth);

        if(lserr != lserrNone)
        {
            goto Cleanup;
        }
    }

    //initialize rgExpType
    expType = exptNone;
    memset(rgExpType, expType, sizeof(EXPTYPE)*cglyph);

    // initialize rgexpinfo to zeros
    memset(rgexpinfo, 0, sizeof(LSEXPINFO)*cglyph);

    // Loop through the glyph attributes. We assume logical order here
    while(iPropL < (int)cglyph)
    {
        bCurrentPr = psva[iPropL].uJustification;

        switch(bCurrentPr)
        {
        case SCRIPT_JUSTIFY_NONE:
            rgExpType[iPropL] = exptNone;
            break;

        case SCRIPT_JUSTIFY_ARABIC_BLANK:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                rgExpType[iPropL] = exptAddWhiteSpace;

                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    break;
                case styleTextJustifyNewspaper:
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    break;
                case styleTextJustifyDistribute:
                    rgexpinfo[iPropL].prior = 2;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    break;
                default:
                    rgExpType[iPropL] = exptNone;
                    if(iBestPr >=0)
                    {
                        iPrevBestPr = iPropL;
                        rgExpType[iBestPr] = exptAddInkContinuous;
                        rgexpinfo[iBestPr].prior = 1;
                        rgexpinfo[iBestPr].duMax = lsexpinfInfinity;
                        rgexpinfo[iBestPr].u.AddInkContinuous.duMin = iKashidaWidth;
                        iBestPr = -1;
                        bBestPr = SCRIPT_JUSTIFY_NONE;
                        iBestKashidaLevel = 10;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;

                if(justifyStyle<styleTextJustifyInterWord ||
                    justifyStyle>styleTextJustifyDistributeAllLines)
                {
                    if(iBestPr >=0)
                    {
                        iPrevBestPr = iPropL;
                        rgExpType[iBestPr] = exptAddInkContinuous;
                        rgexpinfo[iBestPr].prior = 1;
                        rgexpinfo[iBestPr].duMax = lsexpinfInfinity;
                        rgexpinfo[iBestPr].u.AddInkContinuous.duMin = iKashidaWidth;
                        iBestPr = -1;
                        bBestPr = SCRIPT_JUSTIFY_NONE;
                        iBestKashidaLevel = 10;
                    }
                }
            }
            break;

        case SCRIPT_JUSTIFY_CHARACTER:
            // this value was prefilled to exptAddWhiteSpace
            // we should not be here for Arabic type text
            Assert(!fKashida);
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                rgExpType[iPropL] = exptAddWhiteSpace;
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgexpinfo[iPropL].prior = 0;
                    break;
                case styleTextJustifyNewspaper:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                case styleTextJustifyDistribute:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                }
                rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
            }
            // Final character on the line should not get justification
            // flag set.
            else
                rgExpType[iPropL] = exptNone;
            break;
        case SCRIPT_JUSTIFY_BLANK:
            // this value was prefilled to exptAddWhiteSpace
            // we should not be here for Arabic type text
            Assert(!fKashida);
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                rgExpType[iPropL] = exptAddWhiteSpace;
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                case styleTextJustifyNewspaper:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                case styleTextJustifyDistribute:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgexpinfo[iPropL].prior = 1;
                    break;
                }
                rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed after kashida and seen characters
        case SCRIPT_JUSTIFY_ARABIC_KASHIDA:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY1;
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

        case SCRIPT_JUSTIFY_ARABIC_SEEN:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL] = exptAddInkContinuous;
                    rgexpinfo[iPropL].prior = 1;
                    rgexpinfo[iPropL].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY2;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;
                        // these types of kashida go after this glyph visually
                        for(iPropR=iPropL+1; iPropR<(int)cglyph&&psva[iPropR].fDiacritic; iPropR++);
                        if(iPropR != iPropL)
                        {
                            iPropR--;
                        }

                        iBestPr = iPropR;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed before the ha and alef
        case SCRIPT_JUSTIFY_ARABIC_HA:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY3;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;
                        iBestPr = iPropL - 1;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

        case SCRIPT_JUSTIFY_ARABIC_ALEF:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY4;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;
                        iBestPr = iPropL - 1;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed before prior character if prior char
            // is a medial ba type
        case SCRIPT_JUSTIFY_ARABIC_RA:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY5;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;

                        if(psva[iPropL-1].uJustification == SCRIPT_JUSTIFY_ARABIC_BA)
                        {
                            iBestPr = iPropL - 2;
                        }
                        else
                        {
                            iBestPr = iPropL - 1;
                        }
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed before last normal type in word
        case SCRIPT_JUSTIFY_ARABIC_BARA:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY6;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;

                        // these types of kashida go before this glyph visually
                        iBestPr = iPropL - 1;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed before last normal type in word
        case SCRIPT_JUSTIFY_ARABIC_BA:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY7;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        if(iKashidaLevel <= iBestKashidaLevel)
                        {
                            iBestKashidaLevel = iKashidaLevel;

                            // these types of kashida go before this glyph visually
                            iBestPr = iPropL - 1;
                        }
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

            // kashida is placed before last normal type in word
        case SCRIPT_JUSTIFY_ARABIC_NORMAL:
            if(!fLastTextChunkOnLine || (DWORD)iPropL!=(cglyph-1))
            {
                switch(justifyStyle)
                {
                case styleTextJustifyInterWord:
                    rgExpType[iPropL] = exptNone;
                    break;
                case styleTextJustifyNewspaper:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistribute:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                case styleTextJustifyDistributeAllLines:
                    rgExpType[iPropL-1] = exptAddInkContinuous;
                    rgexpinfo[iPropL-1].prior = 1;
                    rgexpinfo[iPropL-1].duMax = lsexpinfInfinity;
                    rgexpinfo[iPropL-1].fCanBeUsedForResidual = TRUE;
                    rgexpinfo[iPropL-1].u.AddInkContinuous.duMin = iKashidaWidth;
                    break;
                default:
                    iKashidaLevel = KASHIDA_PRIORITY8;
                    if(iKashidaLevel <= iBestKashidaLevel)
                    {
                        iBestKashidaLevel = iKashidaLevel;

                        // these types of kashida go before this glyph visually
                        iBestPr = iPropL - 1;
                    }
                    break;
                }
            }
            // Final character on the line should not get justification
            // flag set.
            else
            {
                rgExpType[iPropL] = exptNone;
            }
            break;

        default:
            AssertSz( 0, "We have a new SCRIPT_JUSTIFY type");
        }

        iPropL++;
    }

    if(justifyStyle<styleTextJustifyInterWord || justifyStyle>styleTextJustifyDistributeAllLines)
    {
        if(iBestPr >= 0)
        {
            rgexpinfo[iBestPr].prior = 1;
            rgexpinfo[iBestPr].duMax = lsexpinfInfinity;
            if(fKashida)
            {
                rgexpinfo[iBestPr].u.AddInkContinuous.duMin = iKashidaWidth;
            }
        }
    }

Cleanup:

    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetGlyphExpansionInkInfo (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetGlyphExpansionInkInfo(
                                        PLSRUN plsrun,              // IN
                                        LSDEVICE lsDeviceID,        // IN
                                        GINDEX gindex,              // IN
                                        GPROP gprop,                // IN
                                        LSTFLOW kTFlow,             // IN
                                        DWORD cAddInkDiscrete,      // IN
                                        long* rgDu)                 // OUT
{
    LSTRACE(GetGlyphExpansionInkInfo);
    LSNOTIMPL(GetGlyphExpansionInkInfo);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FTruncateBefore (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FTruncateBefore(
                               PLSRUN plsrunCur,       // IN
                               LSCP cpCur,             // IN
                               WCHAR wchCur,           // IN
                               long durCur,            // IN
                               PLSRUN plsrunPrev,      // IN
                               LSCP cpPrev,            // IN
                               WCHAR wchPrev,          // IN
                               long durPrev,           // IN
                               long durCut,            // IN
                               BOOL* pfTruncateBefore) // OUT
{
    LSTRACE(FTruncateBefore);
    // FUTURE (mikejoch) Need to adjust cp values if we ever implement this.
    LSNOTIMPL(FTruncateBefore);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FHangingPunct (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FHangingPunct(
                             PLSRUN plsrun,
                             MWCLS mwcls,
                             WCHAR wch,
                             BOOL* pfHangingPunct)
{
    LSTRACE(FHangingPunct);

    *pfHangingPunct = FALSE;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetSnapGrid (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetSnapGrid(
                           WCHAR* rgwch,           // IN
                           PLSRUN* rgplsrun,       // IN
                           LSCP* rgcp,             // IN
                           DWORD cwch,             // IN
                           BOOL* rgfSnap,          // OUT
                           DWORD* pwGridNumber)    // OUT
{
    LSTRACE(GetSnapGrid);
    // FUTURE (mikejoch) Need to adjust cp values if we ever do grid justification.

    *pwGridNumber = 0;

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   FCancelHangingPunct (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FCancelHangingPunct(
                                   LSCP cpLim,                // IN
                                   LSCP cpLastAdjustable,      // IN
                                   WCHAR wch,                  // IN
                                   MWCLS mwcls,                // IN
                                   BOOL* pfCancelHangingPunct) // OUT
{
    LSTRACE(FCancelHangingPunct);
    // FUTURE (mikejoch) Need to adjust cp values if we ever implement this.
    LSNOTIMPL(FCancelHangingPunct);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   ModifyCompAtLastChar (member, LS callback)
//
//  Synopsis:   Unimplemented LineServices callback
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::ModifyCompAtLastChar(
                                    LSCP cpLim,             // IN
                                    LSCP cpLastAdjustable,  // IN
                                    WCHAR wchLast,          // IN
                                    MWCLS mwcls,            // IN
                                    long durCompLastRight,  // IN
                                    long durCompLastLeft,   // IN
                                    long* pdurChangeComp)   // OUT
{
    LSTRACE(ModifyCompAtLastChar);
    // FUTURE (mikejoch) Need to adjust cp values if we ever implement this.
    LSNOTIMPL(ModifyCompAtLastChar);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   EnumText (member, LS callback)
//
//  Synopsis:   Enumeration function, currently unimplemented
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::EnumText(
                        PLSRUN plsrun,           // IN
                        LSCP cpFirst,            // IN
                        LSDCP dcp,               // IN
                        LPCWSTR rgwch,           // IN
                        DWORD cwch,              // IN
                        LSTFLOW lstflow,         // IN
                        BOOL fReverseOrder,      // IN
                        BOOL fGeometryProvided,  // IN
                        const POINT* pptStart,   // IN
                        PCHEIGHTS pheightsPres,  // IN
                        long dupRun,             // IN
                        BOOL fCharWidthProvided, // IN
                        long* rgdup)             // IN
{
    LSTRACE(EnumText);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   EnumTab (member, LS callback)
//
//  Synopsis:   Enumeration function, currently unimplemented
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::EnumTab(
                       PLSRUN plsrun,          // IN
                       LSCP cpFirst,           // IN
                       WCHAR* rgwch,           // IN                   
                       WCHAR wchTabLeader,     // IN
                       LSTFLOW lstflow,        // IN
                       BOOL fReversedOrder,    // IN
                       BOOL fGeometryProvided, // IN
                       const POINT* pptStart,  // IN
                       PCHEIGHTS pheightsPres, // IN
                       long dupRun)
{
    LSTRACE(EnumTab);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   EnumPen (member, LS callback)
//
//  Synopsis:   Enumeration function, currently unimplemented
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::EnumPen(
                       BOOL fBorder,           // IN
                       LSTFLOW lstflow,        // IN
                       BOOL fReverseOrder,     // IN
                       BOOL fGeometryProvided, // IN
                       const POINT* pptStart,  // IN
                       long dup,               // IN
                       long dvp)               // IN
{
    LSTRACE(EnumPen);

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   GetObjectHandlerInfo (member, LS callback)
//
//  Synopsis:   Returns an object handler for the client-side functionality
//              of objects which are handled primarily by LineServices.
//
//  Returns:    lserrNone
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetObjectHandlerInfo(
                                    DWORD idObj,        // IN
                                    void* pObjectInfo)  // OUT
{
    LSTRACE(GetObjectHandlerInfo);

    switch(idObj)
    {
    case LSOBJID_RUBY:
        Assert(sizeof(RUBYINIT) == sizeof(::RUBYINIT));
        *(RUBYINIT*)pObjectInfo = s_rubyinit;
        break;

    case LSOBJID_TATENAKAYOKO:
        Assert( sizeof(TATENAKAYOKOINIT) == sizeof(::TATENAKAYOKOINIT));
        *(TATENAKAYOKOINIT*)pObjectInfo = s_tatenakayokoinit;
        break;

    case LSOBJID_HIH:
        Assert( sizeof(HIHINIT) == sizeof(::HIHINIT));
        *(HIHINIT*)pObjectInfo = s_hihinit;
        break;

    case LSOBJID_WARICHU:
        Assert( sizeof(WARICHUINIT) == sizeof(::WARICHUINIT));
        *(WARICHUINIT*)pObjectInfo = s_warichuinit;
        break;

    case LSOBJID_REVERSE:
        Assert( sizeof(REVERSEINIT) == sizeof(::REVERSEINIT));
        *(REVERSEINIT*)pObjectInfo = s_reverseinit;
        break;

    default:
        AssertSz(0, "Unknown LS object ID.");
        break;
    }

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   AssertFailed (member, LS callback)
//
//  Synopsis:   Assert callback for LineServices
//
//  Returns:    Nothing.
//
//-----------------------------------------------------------------------------
void WINAPI CLineServices::AssertFailed(char* szMessage, char* szFile, int iLine)
{
    LSTRACE(AssertFailed);
}

//-----------------------------------------------------------------------------
//
//  Function:   ChunkifyTextRun
//
//  Synopsis:   Break up a text run if necessary.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::ChunkifyTextRun(COneRun* por, COneRun** pporOut)
{
    LONG    cchRun;
    LPCWSTR pwchRun;
    BOOL    fHasInclEOLWhite = por->_pPF->HasInclEOLWhite(por->_fInnerPF);
    const DWORD cpCurr = por->Cp();
    LSERR   lserr = lserrNone;

    *pporOut = por;

    // 1) If there is a whitespace at the beginning of line, we
    //    do not want to show the whitespace (0 width -- the right
    //    way to do it is to say that the run is hidden).
    if(IsFirstNonWhiteOnLine(cpCurr))
    {
        const TCHAR* pwchRun = por->_pchBase;
        DWORD cp = cpCurr;
        cchRun  = por->_lscch;

        if(!fHasInclEOLWhite)
        {
            while(cchRun)
            {
                const TCHAR ch = *pwchRun++;

                if(!IsWhite(ch))
                {
                    break;
                }

                // Note a whitespace at BOL
                WhiteAtBOL(cp);
                //_lineFlags.AddLineFlag(cp, FLAG_HAS_NOBLAST);
                cp++;

                // Goto next character
                cchRun--;
            }
        }


        // Did we find any whitespaces at BOL? If we did, then
        // create a chunk with those whitespace and mark them
        // as hidden.
        if(cchRun != por->_lscch)
        {
            por->_lscch -= cchRun;
            por->_fHidden = TRUE;
            goto Cleanup;
        }
    }

    // 2. Fold whitespace after an aligned or abspos'd site if there
    //    was a whitespace *before* the aligned site. The way we do this
    //    folding is by saying that the present space is hidden.
    {
        const TCHAR chFirst = *por->_pchBase;
        if(!fHasInclEOLWhite && !_fIsEditable)
        {
            if(IsWhite(chFirst) && NeedToEatThisSpace(por))
            {
                _lineFlags.AddLineFlag(cpCurr, FLAG_HAS_NOBLAST);
                por->_lscch = 1;
                por->_fHidden = TRUE;
                goto Cleanup;
            }
        }

        if(WCH_NONREQHYPHEN == chFirst)
        {
            _lineFlags.AddLineFlag(cpCurr, FLAG_HAS_NOBLAST);
        }
    }

    if(!fHasInclEOLWhite && IsFirstNonWhiteOnLine(cpCurr))
    {
        // 3. Note any \n\r's
        const TCHAR ch = *por->_pchBase;

        if(ch==TEXT('\r') || ch==TEXT('\n'))
        {
            WhiteAtBOL(cpCurr);
            _lineFlags.AddLineFlag(cpCurr, FLAG_HAS_NOBLAST);
        }
    }

    if(_fScanForCR)
    {
        cchRun = por->_lscch;
        pwchRun = por->_pchBase;

        if(WCH_CR == *pwchRun)
        {
            // If all we have on the line are sites, then we do not want the \r to
            // contribute to height of the line and hence we set _fNoTextMetrics to TRUE.
            if(LineHasOnlySites(por->Cp()))
            {
                por->_fNoTextMetrics = TRUE;
            }

            lserr = TerminateLine(por, TL_ADDNONE, pporOut);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }
            if(*pporOut == NULL)
            {
                // All lines ending in a carriage return have the BR flag set on them.
                _lineFlags.AddLineFlag(cpCurr, FLAG_HAS_A_BR);
                por->FillSynthData(SYNTHTYPE_LINEBREAK);
                *pporOut = por;
            }
            goto Cleanup;
        }
        // if we came here after a single site in _fScanForCR mode, it means that we
        // have some text after the single site and hence it should be on the next
        // line. Hence terminate the line here. (If it were followed by a \r, then
        // it would have fallen in the above case which would have consumed that \r)
        else if(_fSingleSite)
        {
            lserr = TerminateLine(por, TL_ADDEOS, pporOut);
            goto Cleanup;
        }
        else
        {
            LONG cch = 0;
            while(cch != cchRun)
            {
                if(WCH_CR == *pwchRun)
                {
                    break;
                }
                cch++;
                pwchRun++;
            }
            por->_lscch = cch;
        }
    }

    Assert(por->_ptp && por->_ptp->IsText() && por->_sid==DWORD(por->_ptp->Sid()));

    // BUGBUG(CThrash): Please improve this sid check!
    if(_pBidiLine!=NULL
        || sidAsciiLatin!=por->_sid
        || _afxGlobalData._iNumShape!=NUMSHAPE_NONE
        || _li._fLookaheadForGlyphing)
    {
        BOOL fGlyph = FALSE;
        BOOL fForceGlyphing = FALSE;
        BOOL fNeedBidiLine = (_pBidiLine != NULL);
        BOOL fRTL = FALSE;
        DWORD uLangDigits = LANG_NEUTRAL;
        WCHAR chNext = WCH_NULL;

        // 4. Note any glyphable etc chars
        cchRun = por->_lscch;
        pwchRun = por->_pchBase;
        while(cchRun-- && !(fGlyph && fNeedBidiLine))
        {
            const TCHAR ch = *pwchRun++;

            fGlyph |= IsGlyphableChar(ch);
            fNeedBidiLine |= IsRTLChar(ch);
        }

        // 5. Break run based on the text direction.
        if(fNeedBidiLine && _pBidiLine==NULL)
        {
            _pBidiLine = new CBidiLine(_treeInfo, _cpStart, _li._fRTL, _pli);
        }
        if(_pBidiLine != NULL)
        {
            por->_lscch = _pBidiLine->GetRunCchRemaining(por->Cp(), por->_lscch);
            // FUTURE (mikejoch) We are handling symmetric swapping by forced
            // glyphing of RTL runs. We may be able to do this faster by
            // swapping symmetric characters in CLSRenderer::TextOut().
            fRTL = fForceGlyphing = _pBidiLine->GetLevel(por->Cp()) & 1;
        }

        // 6. Break run based on the digits.
        if(_afxGlobalData._iNumShape != NUMSHAPE_NONE)
        {
            cchRun = por->_lscch;
            pwchRun = por->_pchBase;
            while(cchRun && !InRange(*pwchRun, ',', '9'))
            {
                pwchRun++;
                cchRun--;
            }
            if(cchRun)
            {
                if(_afxGlobalData._iNumShape == NUMSHAPE_NATIVE)
                {
                    uLangDigits = _afxGlobalData._uLangNationalDigits;
                    fGlyph = TRUE;
                }
                else
                {
                    COneRun* porContext;

                    Assert(_afxGlobalData._iNumShape==NUMSHAPE_CONTEXT && InRange(*pwchRun, ',', '9'));

                    // Scan back for first strong text.
                    cchRun = pwchRun - por->_pchBase;
                    pwchRun--;
                    while(cchRun!=0 && (!IsStrongClass(DirClassFromCh(*pwchRun)) || InRange(*pwchRun, WCH_LRM, WCH_RLM)))
                    {
                        cchRun--;
                        pwchRun--;
                    }
                    porContext = _listCurrent._pTail;
                    if(porContext == por)
                    {
                        porContext = porContext->_pPrev;
                    }
                    while(cchRun==0 && porContext!=NULL)
                    {
                        if(porContext->IsNormalRun() && porContext->_pchBase!=NULL)
                        {
                            cchRun = porContext->_lscch;
                            pwchRun = porContext->_pchBase + cchRun - 1;
                            while(cchRun!=0 && (!IsStrongClass(DirClassFromCh(*pwchRun)) || InRange(*pwchRun, WCH_LRM, WCH_RLM)))
                            {
                                cchRun--;
                                pwchRun--;
                            }
                        }
                        porContext = porContext->_pPrev;
                    }

                    if(cchRun != 0)
                    {
                        if(ScriptIDFromCh(*pwchRun) == DefaultScriptIDFromLang(_afxGlobalData._uLangNationalDigits))
                        {
                            uLangDigits = _afxGlobalData._uLangNationalDigits;
                            fGlyph = TRUE;
                        }
                    }
                    else if(IsRTLLang(_afxGlobalData._uLangNationalDigits) && _li._fRTL)
                    {
                        uLangDigits = _afxGlobalData._uLangNationalDigits;
                        fGlyph = TRUE;
                    }
                }
            }
        }

        // 7. Check if we should have glyphed the prior run (for combining
        //    Latin diacritics; esp. Vietnamese)
        if(_lsMode == LSMODE_MEASURER)
        {
            if(fGlyph && !_li._fLookaheadForGlyphing && IsCombiningMark(*(por->_pchBase)))
            {
                // We want to break the shaping if the formats are not similar
                COneRun* porPrev = _listCurrent._pTail;

                while(porPrev!=NULL && !porPrev->IsNormalRun())
                {
                    porPrev = porPrev->_pPrev;
                }

                if(porPrev!=NULL && !porPrev->_lsCharProps.fGlyphBased)
                {
                    const CCharFormat* pCF1 = por->GetCF();
                    const CCharFormat* pCF2 = porPrev->GetCF();

                    Assert(pCF1!=NULL && pCF2!=NULL);
                    _li._fLookaheadForGlyphing = pCF1->CompareForLikeFormat(pCF2);
                }
            }
        }
        else
        {
            if(_li._fLookaheadForGlyphing)
            {
                Assert(por->_lscch >= 1);

                CTxtPtr txtptr(_pMarkup, por->Cp()+por->_lscch-1);

                // N.B. (mikejoch) This is an extremely non-kosher lookup to do
                // here. It is quite possible that chNext will be from an
                // entirely different site. If that happens, though, it will
                // only cause the unnecessary glyphing of this run, which
                // doesn't actually affect the visual appearence.
                while((chNext=txtptr.NextChar()) == WCH_NODE);
                if(IsCombiningMark(chNext))
                {
                    // Good chance this run needs to be glyphed with the next one.
                    // Turn glyphing on.
                    fGlyph = fForceGlyphing = TRUE;
                }
            }
        }

        // 8. Break run based on the script.
        if(fGlyph || fRTL)
        {
            CComplexRun* pcr = por->GetNewComplexRun();

            if(pcr == NULL)
            {
                return lserrOutOfMemory;
            }

            pcr->ComputeAnalysis(_pFlowLayout, fRTL, fForceGlyphing, chNext,
                _chPassword, por, _listCurrent._pTail, uLangDigits);

            // Something on the line needs glyphing.
            if(por->_lsCharProps.fGlyphBased)
            {
                _fGlyphOnLine = TRUE;

#ifndef NO_UTF16
                if(por->_sid!=sidSurrogateA && por->_sid!=sidSurrogateB)
#endif
                {
                    por->_sid = sidDefault;
                }
            }
        }
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   NeedToEatThisSpace
//
//  Synopsis:   Decide if the current space needs to be eaten. We eat any space
//              after a abspos or aligned site *IF* that site was preceeded by
//              a space too.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
BOOL CLineServices::NeedToEatThisSpace(COneRun* porIn)
{
    BOOL fMightFold = FALSE;
    BOOL fFold      = FALSE;
    CElement* pElementLayout;
    CTreeNode* pNode;
    COneRun* por;

    por = FindOneRun(porIn->_lscpBase);
    if(por==NULL && porIn->_lscpBase>=_listCurrent._pTail->_lscpBase)
    {
        por = _listCurrent._pTail;
    }

    // BUGBUG(SujalP): We will not fold across hidden stuff... need to fix this...
    while(por)
    {
        if(por->_fCharsForNestedLayout)
        {
            pNode = por->Branch();
            Assert(pNode->NeedsLayout());
            pElementLayout = pNode->GetUpdatedLayout()->ElementOwner();
            if(!pElementLayout->IsInlinedElement())
            {
                fMightFold = TRUE;
            }
            else
            {
                break;
            }
        }
        else
        {
            break;
        }
        por = por->_pPrev;
    }

    if(fMightFold)
    {
        if(!por)
        {
            fFold = TRUE;
        }
        else if(!por->_fCharsForNestedElement)
        {
            Assert(por->_pchBase);
            TCHAR ch = por->_pchBase[por->_lscch-1];
            if(ch==_T(' ') || ch==_T('\t'))
            {
                fFold = TRUE;
            }
        }
    }

    return fFold;
}

//-----------------------------------------------------------------------------
//
//  Function:   ChunkifyObjectRun
//
//  Synopsis:   Breakup a object run if necessary.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::ChunkifyObjectRun(COneRun* por, COneRun** pporOut)
{
    CElement*   pElementLayout;
    CLayout*    pLayout;
    CTreeNode*  pNode;
    BOOL        fInlinedElement;
    BOOL        fIsAbsolute = FALSE;
    LSERR       lserr  = lserrNone;
    LONG        cp     = por->Cp();
    COneRun*    porOut = por;

    Assert(por->_lsCharProps.idObj == LSOBJID_EMBEDDED);
    pNode = por->Branch();
    Assert(pNode);
    pLayout = pNode->GetUpdatedLayout();
    Assert(pLayout);
    pElementLayout = pLayout->ElementOwner();
    Assert(pElementLayout);
    fInlinedElement = pElementLayout->IsInlinedElement();

    // Setup all the various flags and counts
    if(fInlinedElement)
    {
        _lineCounts.AddLineCount(cp, LC_INLINEDSITES, por->_lscch);
    }
    else
    {
        fIsAbsolute = pNode->IsAbsolute((stylePosition)por->GetFF()->_bPositionType);

        if(!fIsAbsolute)
        {
            _lineCounts.AddLineCount(cp, LC_ALIGNEDSITES, por->_lscch);
            _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_ALIGNED);
        }
        else
        {
            // This is the only opportunity for us to measure abspos'd sites
            if(_lsMode == LSMODE_MEASURER)
            {
                _lineCounts.AddLineCount(cp, LC_ABSOLUTESITES, por->_lscch);
                pLayout->SetXProposed(0);
                pLayout->SetYProposed(0);
            }
            _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_ABSOLUTE_ELT);
        }
    }
    _lineFlags.AddLineFlag(cp, FLAG_HAS_EMBED_OR_WBR);

    if(por->_fCharsForNestedRunOwner)
    {
        _lineFlags.AddLineFlag(cp, FLAG_HAS_NESTED_RO);
        _pMeasurer->_pRunOwner = pLayout;
    }

    if(IsOwnLineSite(por))
    {
        _fSingleSite = TRUE;
        _li._fSingleSite = TRUE;
        _li._fHasEOP = TRUE;
    }

    if(!fInlinedElement)
    {
        // Since we are not showing this run, lets just anti-synth it!
        por->MakeRunAntiSynthetic();

        // And remember that these chars are white at BOL
        if(IsFirstNonWhiteOnLine(cp))
        {
            AlignedAtBOL(cp, por->_lscch);
        }
    }

    *pporOut = porOut;
    return lserr;
}

//+==============================================================================
//
//  Method: GetRenderingHighlights
//
//  Synopsis: Get the type of highlights between these two cp's by going through
//            the array of HighlightSegments on
//
//  A 'Highlight' denotes a "non-content-based" way of changing the rendering
// of something. Currently the only highlight is for selection. This may change.
//
//-------------------------------------------------------------------------------
// BUGBUG ( marka ) - this routine has been made generic enough so it could work
// for any type of highlighting. However we'll assume there's just Selection
// as then we can take 'quick outs' and bail out of looping through the entire array all the time
DWORD CLineServices::GetRenderingHighlights(int cpLineMin, int cpLineMax)
{
    int i;
    HighlightSegment** ppHighlight;
    DWORD dwHighlight = HIGHLIGHT_TYPE_None;

    CLSRenderer* pRenderer = GetRenderer();

    if(pRenderer->_cpSelMax != - 1)
    {
        for(i=pRenderer->_aryHighlight.Size(),ppHighlight=pRenderer->_aryHighlight;
            i>0; i--,ppHighlight++)
        {
            if(((*ppHighlight)->_cpStart<=cpLineMin) && ((*ppHighlight)->_cpEnd>=cpLineMax))
            {
                dwHighlight |= (*ppHighlight)->_dwHighlightType ;
            }
        }
    }
    return dwHighlight;
}

//+====================================================================================
//
// Method: SetRenderingMarkup
//
// Synopsis: Set any 'markup' on this COneRun
//
// BUGBUG: marka - currently the only type of Markup is Selection. We see if the given
//         run is selected, and if so we mark-it.
//
//         This used to be called ChunkfiyForSelection. However, under due to the new selection
//         model's use of TreePos's the Run is automagically chunkified for us
//
//------------------------------------------------------------------------------------
LSERR CLineServices::SetRenderingHighlights(COneRun* por)
{
    DWORD dwHighlight;
    DWORD dwImeHighlight;    
    int cpMin = por->Cp()  ;
    int cpMax = por->Cp() + por->_lscch ;

    // If we are not rendering we should do nothing
    if(_lsMode != LSMODE_RENDERER)
    {
        goto Cleanup;
    }

    // We will not show selection if it is hidden or its an object.
    if(por->_fHidden || por->_lsCharProps.idObj==LSOBJID_EMBEDDED)
    {
        goto Cleanup;
    }

    dwHighlight = GetRenderingHighlights(cpMin, cpMax);

#ifndef NO_IME
    dwImeHighlight = (dwHighlight/HIGHLIGHT_TYPE_ImeInput) & 7;

    if(dwImeHighlight)
    {
        por->ImeHighlight(_pFlowLayout, dwImeHighlight);
    }
#endif // NO_IME

    if((dwHighlight&HIGHLIGHT_TYPE_Selected) != 0)
    {
        por->Selected(_pFlowLayout, TRUE);
    }

Cleanup:
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckForUnderLine
//
//  Synopsis:   Check if the current run is underlined/overlined.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::CheckForUnderLine(COneRun* por)
{
    LSERR lserr = lserrNone;

    if(!por->_dwImeHighlight)
    {
        const CCharFormat* pCF = por->GetCF();

        por->_lsCharProps.fUnderline = (_fIsEditable || !pCF->IsVisibilityHidden())
            && (pCF->_fUnderline || pCF->_fOverline || pCF->_fStrikeOut);
    }
    else
    {
        por->_lsCharProps.fUnderline = por->_fUnderlineForIME;
    }

    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckForTransform
//
//  Synopsis:   If text transformation is necessary for the run, then do it here.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::CheckForTransform(COneRun* por)
{
    LSERR lserr = lserrNone;
    const CCharFormat* pCF = por->GetCF();
    if(pCF->IsTextTransformNeeded())
    {
        TCHAR chPrev = WCH_NULL;

        _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_NOBLAST);

        if(pCF->_bTextTransform == styleTextTransformCapitalize)
        {
            COneRun* porTemp = FindOneRun(por->_lscpBase-1);

            // If no run found on this line before this run, then its the first
            // character on this line, and hence needs to be captilaized. This is
            // done implicitly by initializing chPrev to WCH_NULL.
            while(porTemp)
            {
                // If it is anti-synth, then its not shown at all or
                // is not in the flow, hence we ignore it for the purposes
                // of transformations. If it is synthetic then we will need
                // to look at runs before the synthetic to determine the
                // prev character. Hence, we only need to investigate normal runs
                if(porTemp->IsNormalRun())
                {
                    // If the previous run had a nested layout, then we will ignore it.
                    // The rule says that if there is a layout in the middle of a word
                    // then the character after the layout is *NOT* capitalized. If the
                    // layout is at the beginning of the word then the character needs
                    // to be capitalized. In essence, we completely ignore layouts when
                    // deciding whether a char should be capitalized, i.e. if we had
                    // the combination:
                    // <charX><layout><charY>, then capitalization of <charY> depends
                    // only on what <charX> -- ignoring the fact that there is a layout
                    // (It does not matter if the layout is hidden, aligned, abspos'd
                    // or relatively positioned).
                    if(!porTemp->_fCharsForNestedLayout && porTemp->_synthType==SYNTHTYPE_NONE)
                    {
                        Assert(porTemp->_ptp->IsText());
                        chPrev = porTemp->_pchBase[porTemp->_lscch-1];
                        break;
                    }
                }
                porTemp = porTemp->_pPrev;
            }
        }

        por->_pchBase = (TCHAR*)TransformText(por->_cstrRunChars, por->_pchBase, por->_lscch,
            pCF->_bTextTransform, chPrev);
        if(por->_pchBase == NULL)
        {
            lserr = lserrOutOfMemory;
        }
    }
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckForPassword
//
//  Synopsis:   If text transformation is necessary for the run, then do it here.
//
//  Returns:    lserr.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::CheckForPassword(COneRun* por)
{
    LSERR lserr = lserrNone;
    CString strPassword;
    HRESULT hr;

    Assert(_chPassword);

    for(LONG i=0; i<por->_lscch; i++)
    {
        hr = strPassword.Append(&_chPassword, 1);
        if(hr != S_OK)
        {
            lserr = lserrOutOfMemory;
            goto Cleanup;
        }
    }
    por->_pchBase = por->SetString(strPassword);
    if(por->_pchBase == NULL)
    {
        lserr = lserrOutOfMemory;
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   AdjustForNavBRBug (member)
//
//  Synopsis:   Navigator will break between a space and a BR character at
//              the end of a line if the BR character, when it's width is
//              treated like that of a space character, will cause line to
//              overflow.  This is idiotic, but necessary for compat.
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLineServices::AdjustCpLimForNavBRBug(
                                      LONG xWrapWidth,        // IN
                                      LSLINFO* plslinfo)      // IN/OUT
{
    LSERR lserr = lserrNone;

    // check 1: (a) We must not be in a PRE and (b) the line
    // must have at least three chars.
    if(!_pPFFirst->HasPre(_fInnerPFFirst) && _cpLim-plslinfo->cpFirstVis>=3)
    {
        COneRun* por = FindOneRun(_lscpLim-1);
        if(!por)
        {
            goto Cleanup;
        }

        // check 2: Line must have ended in a BR
        if(por->_ptp->IsEndNode() && por->Branch()->Tag()==ETAG_BR)
        {
            // check 3: The BR char must be preceeded by a space char.

            // Go to the begin BR
            por = por->_pPrev;
            if(!(por
                && por->IsAntiSyntheticRun()
                && por->_ptp->IsBeginNode()
                && por->Branch()->Tag()==ETAG_BR))
            {
                goto Cleanup;
            }

            // Now go one more beyond that to check for the space
            do
            {
                por = por->_pPrev;
            } while(por && por->IsAntiSyntheticRun() && !por->_fCharsForNestedLayout);
            
            if(!por)
            {
                goto Cleanup;
            }

            // BUGBUG(SujalP + CThrash): This will not work if the space was
            // in, say, a reverse object. But then this *is* a NAV bug. If
            // somebody complains vehemently, then we will fix it...
            if(por->IsNormalRun() && por->_ptp->IsText()
                && WCH_SPACE==por->_pchBase[por->_lscch-1])
            {
                if(_fMinMaxPass)
                {
                    ((CDisplay*)_pMeasurer->_pdp)->SetNavHackPossible();
                }

                // check 4: must have overflowed, because the width of a BR
                // character is included in _xWidth
                if(!_fMinMaxPass && _li._xWidth>xWrapWidth)
                {
                    // check 5:  The BR alone cannot be overflowing.  We must
                    // have at least one pixel more to break before the BR.
                    HRESULT hr;
                    LSTEXTCELL lsTextCell;

                    hr = QueryLineCpPpoint(_lscpLim, FALSE, NULL, &lsTextCell);

                    if(S_OK==hr && (_li._xWidth - lsTextCell.dupCell)>xWrapWidth)
                    {
                        // Break between the space and the BR.  Yuck! Also 2
                        // here because one for open BR and one for close BR
                        _cpLim -= 2;

                        // The char for open BR was antisynth'd in the lscp
                        // space, so just reduce by one char for the close BR.
                        _lscpLim--;
                    }
                }
            }
        }
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   AdjustForRelElementAtEnd (member)
//
//  Synopsis: In our quest to put as much on the line as possible we will end up
//      putting the begin splay for a relatively positioned element on this
//      line(the first character in this element will be on the next line)
//      This is not really acceptable for positioning purposes and hence
//      we detect that this happened and chop off the relative element
//      begin splay (and any chars anti-synth'd after it) so that they will
//      go to the next line. (Look at bug 54162).
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
VOID CLineServices::AdjustForRelElementAtEnd()
{
    // By looking for _lscpLim - 1, we find the last but-one-run. After
    // this run, there will possibly be antisynthetic runs (all at the same
    // lscp -- but different cp's -- as the last char of this run) which
    // have ended up on this line. It is these set of antisynthetic runs
    // that we have to investigate to find any ones which begin a relatively
    // positioned element. If we do find one, we will stop and modify cpLim
    // so as to not include the begin splay and any antisynth's after it.
    COneRun* por = FindOneRun(_lscpLim - 1);

    Assert(por);

    // Go to the start of the antisynth chunk (if any).
    por = por->_pNext;

    // Walk while we have antisynth runs
    while(por && por->IsAntiSyntheticRun()
        && _lscpLim >= por->_lscpBase+por->_lscch)
    {
        // If it begins a relatively positioned element, then we want to
        // stop and modify the cpLim
        if(por->_ptp->IsBeginElementScope() && por->GetCF()->IsRelative(por->_fInnerCF))
        {
            _cpLim = por->Cp();
            break;
        }
        Assert(por->_lscch == 1);
        por = por->_pNext;
    }
}

//-----------------------------------------------------------------------------
//
//  Function:   ComputeWhiteInfo (member)
//
//  Synopsis:   A post pass for the CMeasurer to compute whitespace
//              information (_cchWhite and _xWhite) on the associated
//              CLine object (_li).
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
HRESULT CLineServices::ComputeWhiteInfo(LSLINFO* plslinfo, LONG* pxMinLineWidth, DWORD dwlf,
                                LONG durWithTrailing, LONG durWithoutTrailing)
{
    HRESULT hr = S_OK;
    BOOL  fInclEOLWhite = _pPFFirst->HasInclEOLWhite(_fInnerPFFirst);
    CMarginInfo* pMarginInfo = (CMarginInfo*)_pMarginInfo;

    // Note that cpLim is no longer an LSCP.  It's been converted by the
    // caller.
    const LONG cpLim = _cpLim;
    CTxtPtr txtptr(_pMarkup, cpLim);
    const TCHAR chBreak = txtptr.GetPrevChar();

    LONG lscpLim = _lscpLim - 1;
    COneRun* porLast = FindOneRun(lscpLim);

    Assert(_cpLim == _cpStart+_li._cch);

    // Compute all the flags for the line
    Assert(dwlf == _lineFlags.GetLineFlags(cpLim));
    _li._fHasAbsoluteElt      = (dwlf & FLAG_HAS_ABSOLUTE_ELT) ? TRUE : FALSE;
    _li._fHasAligned          = (dwlf & FLAG_HAS_ALIGNED)      ? TRUE : FALSE;
    _li._fHasEmbedOrWbr       = (dwlf & FLAG_HAS_EMBED_OR_WBR) ? TRUE : FALSE;
    _li._fHasNestedRunOwner   = (dwlf & FLAG_HAS_NESTED_RO)    ? TRUE : FALSE;
    _li._fHasBackground       = (dwlf & FLAG_HAS_BACKGROUND)   ? TRUE : FALSE;
    _li._fHasNBSPs            = (dwlf & FLAG_HAS_NBSP)         ? TRUE : FALSE;
    _fHasRelative             = (dwlf & FLAG_HAS_RELATIVE)     ? TRUE : FALSE;
    _fFoundLineHeight         = (dwlf & FLAG_HAS_LINEHEIGHT)   ? TRUE : FALSE;
    pMarginInfo->_fClearLeft |= (dwlf & FLAG_HAS_CLEARLEFT)    ? TRUE : FALSE;
    pMarginInfo->_fClearRight|= (dwlf & FLAG_HAS_CLEARRIGHT)   ? TRUE : FALSE;

    _li._fHasBreak            = (dwlf & FLAG_HAS_A_BR)         ? TRUE : FALSE;

    // Lines containing \r's also need to be marked _fHasBreak
    if(!_li._fHasBreak && _fExpectCRLF && plslinfo->endr==endrEndPara)
    {
        _li._fHasBreak = TRUE;
    }

    _pFlowLayout->_fContainsRelative |= _fHasRelative;

    // If all we have is whitespaces till here then mark it as a dummy line
    if(IsDummyLine(cpLim))
    {
        const LONG cchHidden = _lineCounts.GetLineCount(cpLim, LC_HIDDEN);
        const LONG cch = (_cpLim - _cpStart) - cchHidden;

        _li._fDummyLine = TRUE;
        _li._fForceNewLine = FALSE;

        // If this line was a dummy line because all it contained was hidden
        // characters, then we need to mark the entire line as hidden.  Also
        // if the paragraph contains nothing (except a blockbreak), we also
        // need the hide the line.  Note that LI's are an exception to this
        // rule -- even if all we have on the line is a blockbreak, we don't
        // want to hide it if it's an LI. (LI's are excluded in the dummy
        // line check).
        if(cchHidden && (cch==0 || _li._fFirstInPara))
        {
            _li._fHidden = TRUE;
            _li._yBeforeSpace = 0;
        }
    }

    // Also find out all the relevant counts
    _cInlinedSites  = _lineCounts.GetLineCount(cpLim, LC_INLINEDSITES);
    _cAbsoluteSites = _lineCounts.GetLineCount(cpLim, LC_ABSOLUTESITES);
    _cAlignedSites  = _lineCounts.GetLineCount(cpLim, LC_ALIGNEDSITES);

    // And the relevant values
    if(_fFoundLineHeight)
    {
        _lMaxLineHeight = max(plslinfo->dvpMultiLineHeight,
            _lineCounts.GetMaxLineValue(cpLim, LC_LINEHEIGHT));
    }

    // Consume trailing WCH_NOSCOPE/WCH_BLOCKBREAK characters.
    _li._cchWhite = 0;
    _li._xWhite = 0;

    if(porLast && porLast->_ptp->IsNode() && porLast->Branch()->Tag()!=ETAG_BR)
    {
        // BUGBUG (cthrash) We're potentially tacking on characters but are
        // not included their widths.  These character can/will have widths in
        // edit mode.  The problem is, LS never measured them, so we'd have
        // to measure them separately.
        CTxtPtr txtptrT(txtptr);
        long dcch;

        // If we have a site that lives on it's own line, we may have stopped
        // fetching characters prematurely.  Specifically, we may have left
        // some space characters behind.
        while(TEXT(' ') == txtptrT.GetChar())
        {
            if(!txtptrT.AdvanceCp(1))
            {
                break;
            }
        }
        dcch = txtptrT.GetCp() - txtptr.GetCp();

        _li._cchWhite += (SHORT)dcch;
        _li._cch += dcch;
    }

    // Compute _cchWhite and _xWhite of line
    if(!fInclEOLWhite)
    {
        BOOL fDone = FALSE;
        TCHAR ch;
        COneRun* por = porLast;
        LONG index = por ? lscpLim - por->_lscpBase : 0;

        while(por && !fDone)
        {
            // BUGBUG(SujalP): As noted below, this does not work with
            // aligned/abspos'd sites at EOL
            if(por->_fCharsForNestedLayout)
            {
                fDone = TRUE;
                break;
            }

            if(por->IsNormalRun())
            {
                for(LONG i=index; i>=0; i--)
                {
                    ch = por->_pchBase[i];
                    if(IsWhite(ch)
                        // If its a no scope char and we are not showing it then
                        // we treat it like a whitespace.
                        // BUGBUG(SujalP) We also need this check only
                        // if the site is aligned/abspos'd
                        )
                    {
                        _li._cchWhite++;
                        txtptr.AdvanceCp(-1);
                    }
                    else
                    {
                        fDone = TRUE;
                        break;
                    }
                }
            }
            por = por->_pPrev;
            index = por ? por->_lscch - 1 : 0;
        }

        _li._xWhite  = durWithTrailing - durWithoutTrailing;
        _li._xWidth -= _li._xWhite;

        if(porLast && chBreak==WCH_NODE && !_fScanForCR)
        {
            CTreePos* ptp = porLast->_ptp;

            if(ptp->IsEndElementScope() && ptp->Branch()->Tag()==ETAG_BR)
            {
                LONG cp = CPFromLSCP(plslinfo->cpFirstVis);
                _li._fEatMargin = LONG(txtptr.GetCp())==cp + 1;
            }
        }
    }
    else if(_fScanForCR)
    {
        HRESULT hr;
        LSTEXTCELL lsTextCell;
        CTxtPtr tp(_pMarkup, cpLim);
        TCHAR ch = tp.GetChar();
        TCHAR chPrev = tp.GetPrevChar();
        BOOL fDecWidth = FALSE;
        LONG cpJunk;

        if(chPrev==_T('\n') || chPrev==_T('\r'))
        {
            fDecWidth = TRUE;
        }
        else if(ch == WCH_NODE)
        {
            CTreePos* ptpLast = _pMarkup->TreePosAtCp(cpLim-1, &cpJunk);

            if(ptpLast->IsEndNode()
                && (ptpLast->Branch()->Tag()==ETAG_BR
                || IsPreLikeNode(ptpLast->Branch())))
            {
                fDecWidth = TRUE;
            }
        }

        if(fDecWidth)
        {
            // The width returned by LS includes the \n, which we don't want
            // included in CLine::_xWidth.
            hr = QueryLineCpPpoint(_lscpLim-1, FALSE, NULL, &lsTextCell);
            if(!hr)
            {
                _li._xWidth -= lsTextCell.dupCell; // note _xWhite is unchanged
                if(pxMinLineWidth)
                {
                    *pxMinLineWidth -= lsTextCell.dupCell;
                }
            }
        }
    }
    else
    {
        _li._cchWhite = plslinfo->cpFirstVis - _cpStart;
    }

    // If the white at the end of the line meets the white at the beginning
    // of a line, then we need to shift the BOL white to the EOL white.
    if(_cWhiteAtBOL+_li._cchWhite >= _li._cch)
    {
        _li._cchWhite = _li._cch;
    }

    // Find out if the last char on the line has any overhang, and if so set it on
    // the line.
    if(_fHasOverhang)
    {
        COneRun* por = porLast;

        while(por)
        {
            if(por->_ptp->IsText())
            {
                _li._xLineOverhang = por->_xOverhang;
                if(pxMinLineWidth)
                {
                    *pxMinLineWidth += por->_xOverhang;
                }
                break;
            }
            else if(por->_fCharsForNestedLayout)
            {
                break;
            }
            // Continue for synth and antisynth runs
            por = por->_pPrev;
        }
    }

    // Fold the S_FALSE case in -- don't propagate.
    hr = SUCCEEDED(hr) ? S_OK : hr;
    if(hr)
    {
        goto Cleanup;
    }

    DecideIfLineCanBeBlastedToScreen(_cpStart+_li._cch-_li._cchWhite, dwlf);

Cleanup:
    RRETURN(hr);
}

//-----------------------------------------------------------------------------
//
//  Function:   DecideIfLineCanBeBlastedToScreen
//
//  Synopsis:   Decides if it is possible for a line to be blasted to screen
//              in a single shot.
//
//  Params:     The last cp in the line
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void CLineServices::DecideIfLineCanBeBlastedToScreen(LONG cpEndLine, DWORD dwlf)
{
    // By default you cannot blast a line to the screen
    _li._fCanBlastToScreen = FALSE;

    // 0) If we have eaten a space char or have fontlinking, we cannot blast
    //    line to screen
    if(dwlf & FLAG_HAS_NOBLAST)
    {
        goto Cleanup;
    }

    // 1) No justification
    if(_pPFFirst->GetBlockAlign(_fInnerPFFirst) == htmlBlockAlignJustify)
    {
        goto Cleanup;
    }

    // 2) Only simple LTR
    if(_pBidiLine != NULL)
    {
        goto Cleanup;
    }

    // 3) Cannot handle passwords for blasting.
    if(_chPassword)
    {
        goto Cleanup;
    }

    // 4) If there is a glyph on the line then do not blast.
    if(_fGlyphOnLine)
    {
        goto Cleanup;
    }

    // 5) There's IME highlighting, so we can't blast.
    if(_fSawATab)
    {
        goto Cleanup;
    }

    // 6) There's IME highlighting, so we can't blast.
    if(_fHasIMEHighlight)
    {
        goto Cleanup;
    }

    // None of the restrictions apply, lets blast the sucker to the screen!
    _li._fCanBlastToScreen = TRUE;

Cleanup:
    return;
}

LONG CLineServices::RememberLineHeight(LONG cp, const CCharFormat* pCF, const CBaseCcs* pBaseCcs)
{
    long lAdjLineHeight;

    AssertSz(pCF->_cuvLineHeight.GetUnitType()!=CUnitValue::UNIT_PERCENT,
        "Percent units should have been converted to points in ApplyInnerOuterFormats!");

    // If there's no height set, get out quick.
    if(pCF->_cuvLineHeight.IsNull())
    {
        lAdjLineHeight = pBaseCcs->_yHeight + abs(pBaseCcs->_yOffset);
    }
    // Apply the CSS Attribute LINE_HEIGHT
    else
    {
        const CUnitValue* pcuvUseThis = &pCF->_cuvLineHeight;
        long lFontHeight = 1;

        if(pcuvUseThis->GetUnitType() == CUnitValue::UNIT_FLOAT)
        {
            CUnitValue cuv;
            cuv = pCF->_cuvLineHeight;
            cuv.ScaleTo(CUnitValue::UNIT_EM);
            pcuvUseThis = &cuv;
            lFontHeight = _pci->DocPixelsFromWindowY(pCF->GetHeightInTwips(_pFlowLayout->Doc()), TRUE);
        }

        lAdjLineHeight = pcuvUseThis->YGetPixelValue(_pci, 0, lFontHeight);
        NoteLineHeight(cp, lAdjLineHeight);
    }

    if(pCF->HasLineGrid(TRUE))
    {
        _lineFlags.AddLineFlag(cp, FLAG_HAS_NOBLAST);
        lAdjLineHeight = GetClosestGridMultiple(GetLineGridSize(), lAdjLineHeight);
        NoteLineHeight(cp, lAdjLineHeight);
    }

    return lAdjLineHeight;
}


//-------------------------------------------------------------------------
//
//  Member:     VerticalAlignObjects
//
//  Synopsis:   Process all vertically aligned objects and adjust the line
//              height
//
//-------------------------------------------------------------------------
void CLineServices::VerticalAlignObjects(CLSMeasurer& lsme, unsigned long cSites, long xLineShift)
{
    LONG        cch;
    LONG        yTxtAscent  = _li._yHeight - _li._yDescent;
    LONG        yTxtDescent = _li._yDescent;
    LONG        yDescent    = _li._yDescent;
    LONG        yAscent     = yTxtAscent;
    LONG        yAbsHeight  = 0;
    LONG        yTmpAscent  = 0;
    LONG        yTmpDescent = 0;
    BOOL        fMeasuring  = TRUE;
    BOOL        fPositioning= FALSE;
    CElement*   pElementFL = _pFlowLayout->ElementOwner();
    CLayout*    pRunOwner;
    CLayout*    pLayout;
    CElement*   pElementLayout;
    const CCharFormat* pCF = NULL;
    CTreeNode*  pNode;
    CTreePos*   ptpStart;
    CTreePos*   ptp;
    LONG        ich;
    LONG        cchAdvanceStart;
    LONG        cchAdvance;
    // this is the display's margins.
    CMarginInfo* pMarginInfo = (CMarginInfo*)_pMarginInfo;
    // this is site margins
    long        xLeftMargin = 0;
    long        xRightMargin = 0;

    htmlControlAlign atAbs = htmlControlAlignNotSet;
    BOOL        fRTLDisplay = _pFlowLayout->GetDisplay()->IsRTL();
    LONG        cchPreChars = lsme._fMeasureFromTheStart ? 0 : lsme._cchPreChars;

    ptpStart = _pFlowLayout->GetContentMarkup()->TreePosAtCp(lsme.GetCp()-_li._cch-cchPreChars, &ich);
    Assert(ptpStart);
    cchAdvanceStart = min(long(_li._cch+cchPreChars), ptpStart->GetCch()-ich);

    // first pass we measure the line's baseline and height
    // second pass we set the rcProposed of the site relative to the line.
    while(fMeasuring || fPositioning)
    {
        cch = _li._cch + cchPreChars;
        lsme.Advance(-cch);
        ptp = ptpStart;
        cchAdvance = cchAdvanceStart;
        pRunOwner = _pFlowLayout;

        while(cch)
        {
            if(ptp->IsBeginElementScope())
            {
                pNode = ptp->Branch();

                pLayout = pElementFL!=pNode->Element()&&pNode->NeedsLayout()
                    ? pNode->GetUpdatedLayout() : NULL;
            }
            else
            {
                pNode = ptp->GetBranch();
                pLayout = NULL;
            }

            pElementLayout = pLayout ? pNode->Element() : NULL;
            pCF = pNode->GetCharFormat();

            // If we are transisitioning from a normal chunk to relative or
            // relative to normal chunk, add a new chunk to the line. Relative
            // elements can end in the prechar's so look for transition in
            // prechar's too.
            if(fMeasuring && _fHasRelative && !pCF->IsDisplayNone()
                && (ptp->IsBeginElementScope()
                || ptp->IsText()
                || (ptp->IsEndElementScope()
                &&  cch>(_li._cch-(lsme._fMeasureFromTheStart?lsme._cchPreChars:0)))))
            {
                TestForTransitionToOrFromRelativeChunk(
                    lsme,
                    pCF->IsRelative(SameScope(pNode, pElementFL)),
                    pElementLayout);
            }

            // If the current branch is a site and not the current CTxtSite
            if(pLayout && (_fIsEditable || !pLayout->IsDisplayNone()))
            {
                htmlControlAlign    atSite;
                BOOL                fOwnLine;
                LONG                lBorderSpace;
                LONG                lVPadding = pElementLayout->TestClassFlag(CElement::ELEMENTDESC_VPADDING) ? 1 : 0;
                LONG                yObjHeight;
                ELEMENT_TAG         etag = pElementLayout->Tag();
                CTreeNode*          pNodeLayout = pElementLayout->GetFirstBranch();
                BOOL                fAbsolute = pNodeLayout->IsAbsolute();

                atSite   = pElementLayout->GetSiteAlign();
                fOwnLine = pElementLayout->HasFlag(TAGDESC_OWNLINE);

                if(fAbsolute || pElementLayout->IsInlinedElement())
                {
                    LONG yProposed = 0;
                    LONG yTopMargin, yBottomMargin;

                    pLayout->GetMarginInfo(_pci, &xLeftMargin, &yTopMargin, &xRightMargin, &yBottomMargin);

                    // Do the horizontal positioning. We can do it either during
                    // measuring or during vertical positioning. We arbitrarily
                    // chose to do it during measuring.
                    if(fMeasuring)
                    {
                        {
                            LONG xPosOfCp = 0;
                            if(!fAbsolute || pNode->GetFancyFormat()->_fAutoPositioned)
                            {
                                BOOL fRTLFlow;
                                LSCP lscp = LSCPFromCP(lsme.GetCp());

                                xPosOfCp = CalculateXPositionOfLSCP(lscp, FALSE, &fRTLFlow);

                                // Adjust xPosOfCp if we have any RTL cases
                                // Bug 45767 - check for flow not being same direction of line
                                if(fRTLDisplay || _li._fRTL || fRTLFlow)
                                {
                                    if(!fAbsolute)
                                    {
                                        CSize size;
                                        pLayout->GetSize(&size);

                                        AdjustXPosOfCp(fRTLDisplay, _li._fRTL, fRTLFlow, &xPosOfCp, size.cx);
                                    }
                                    else if(fRTLDisplay)
                                    {
                                        // To make the xPosOfCp have the correct measurement for the
                                        // right to left display tree node, we make it negative because
                                        // origin (0,0) is top/right in RTL display nodes.
                                        xPosOfCp = -xPosOfCp;
                                    }

                                }

                                if(pCF->HasCharGrid(FALSE))
                                {
                                    long xWidth;

                                    _pFlowLayout->GetSiteWidth(pLayout, _pci, FALSE, _xWidthMaxAvail, &xWidth);
                                    xPosOfCp += (GetClosestGridMultiple(GetCharGridSize(), xWidth) - xWidth)/2;
                                }

                                // absolute margins are added in CLayout::HandlePositionRequest
                                // due to reverse flow issues.
                                if(!fAbsolute)
                                {
                                    if(!fRTLDisplay)
                                    {
                                        xPosOfCp += xLeftMargin;
                                    }
                                    else
                                    {
                                        xPosOfCp += xRightMargin;
                                    }
                                }

                                pLayout->SetXProposed(xPosOfCp);
                            }

                            if(fAbsolute)
                            {
                                // BUGBUG (paulnel) What about when the left or right has been specified and the width is auto?
                                //                  The xPos needs to be adjusted to properly place this case.

                                // fix for #3214, absolutely positioned sites with width auto
                                // need to be measure once we know the line shift, since their
                                // size depends on the xPosition in the line.
                                CSaveCalcInfo sci(_pci);
                                long xWidth;

                                _pci->_fUseOffset    = TRUE;
                                _pci->_xOffsetInLine = xLineShift + xPosOfCp;

                                _pFlowLayout->GetSiteWidth(pLayout, _pci, TRUE, _xWidthMaxAvail, &xWidth);
                            }
                        }
                    }

                    if(!fAbsolute)
                    {
                        CSize size;

                        if(etag == ETAG_HR)
                        {
                            lBorderSpace = GRABSIZE;
                        }
                        else
                        {
                            // Netscape has a pixel descent and ascent to the line if one of the sites
                            // spans the entire line vertically(#20029, #25534).
                            lBorderSpace = lVPadding;
                        }

                        pLayout->GetSize(&size);
                        yObjHeight = max(0L, size.cy+yTopMargin+yBottomMargin) + (2*lBorderSpace);

                        if(pCF->_fIsRubyText)
                        {
                            RubyInfo* pRubyInfo = GetRubyInfoFromCp(lsme.GetCp());
                            if(pRubyInfo)
                            {
                                yObjHeight += pRubyInfo->yHeightRubyBase - yTxtDescent + pRubyInfo->yDescentRubyText;
                            }                        
                        }

                        switch(atSite)
                        {
                            // align to the baseline of the text
                        case htmlControlAlignNotSet:
                        case htmlControlAlignBottom:
                        case htmlControlAlignBaseline:
                            {
                                LONG lDefDescent =
                                    (pElementLayout->TestClassFlag(CElement::ELEMENTDESC_HASDEFDESCENT)) ? 4 : 0;

                                if(fMeasuring)
                                {
                                    yTmpDescent = lBorderSpace + lDefDescent;
                                    yTmpAscent  = yObjHeight - yTmpDescent;
                                }
                                else
                                {
                                    yProposed += yAscent - yObjHeight + 2*lBorderSpace + lDefDescent;
                                }
                                break;
                            }

                            // align to the top of the text
                        case htmlControlAlignTextTop:
                            if(fMeasuring)
                            {
                                yTmpAscent  = yTxtAscent + lBorderSpace;
                                yTmpDescent = yObjHeight - yTxtAscent - lBorderSpace;
                            }
                            else
                            {
                                yProposed += yAscent - yTxtAscent + lBorderSpace;
                            }
                            break;

                            // center of the image aligned to the baseline of the text
                        case htmlControlAlignMiddle:
                        case htmlControlAlignCenter:
                            if(fMeasuring)
                            {
                                yTmpAscent  = (yObjHeight + 1)/2; // + 1 for round off
                                yTmpDescent = yObjHeight/2;
                            }
                            else
                            {
                                yProposed += (yAscent + yDescent - yObjHeight) / 2 + lBorderSpace;
                            }
                            break;

                            // align to the top, absmiddle and absbottom of the line, doesn't really
                            // effect the ascent and descent directly, so we store the
                            // absolute height of the object and recompute the ascent
                            // and descent at the end.
                        case htmlControlAlignAbsMiddle:
                            if(fPositioning)
                            {
                                yProposed += (yAscent + yDescent - yObjHeight) / 2 + lBorderSpace;
                                break;
                            } // fall through when measuring and update max abs height
                        case htmlControlAlignTop:
                            if(fPositioning)
                            {
                                yProposed += lBorderSpace;
                                break;
                            } // fall through when measuring and update the max abs height
                        case htmlControlAlignAbsBottom:
                            if(fMeasuring)
                            {
                                yTmpAscent = 0;
                                yTmpDescent = 0;
                                if(yObjHeight > yAbsHeight)
                                {
                                    yAbsHeight = yObjHeight;
                                    atAbs = atSite;
                                }
                            }
                            else
                            {
                                yProposed += yAscent + yDescent - yObjHeight + lBorderSpace;
                            }
                            break;

                        default:        // we don't want to do anything for
                            if(fOwnLine)
                            {
                                if(fMeasuring)
                                {
                                    yTmpDescent = lBorderSpace;
                                    yTmpAscent  = yObjHeight - lBorderSpace;
                                }
                                else
                                {
                                    yProposed += yAscent -  yObjHeight + lBorderSpace;
                                }
                            }
                            break;      // left/center/right aligned objects
                        }

                        // Keep track of the max ascent and descent
                        if(fMeasuring)
                        {
                            if(yTmpAscent > yAscent)
                            {
                                yAscent = yTmpAscent;
                            }
                            if(yTmpDescent > yDescent)
                            {
                                yDescent = yTmpDescent;
                            }
                        }
                        else if(pCF->HasLineGrid(FALSE))
                        {
                            pLayout->SetYProposed(_li._yHeight - yObjHeight - _li._yDescent + yTopMargin);
                        }
                        else
                        {
                            pLayout->SetYProposed(yProposed + lsme._cyTopBordPad + yTopMargin);
                        }
                    }
                }

                // If positioning, add the current layout to the display tree
                if(fPositioning
                    && (_pci->_smMode==SIZEMODE_NATURAL
                    || _pci->_smMode==SIZEMODE_SET
                    || _pci->_smMode==SIZEMODE_FULLSIZE)
                    && !pElementLayout->IsAligned())
                {
                    const CFancyFormat* pFF = pElementLayout->GetFirstBranch()->GetFancyFormat();
                    long xPos;

                    if(!fRTLDisplay)
                    {
                        xPos = pMarginInfo->_xLeftMargin + _li._xLeft;
                    }
                    else
                    {
                        xPos = - pMarginInfo->_xRightMargin - _li._xRight;
                    }

                    if(!pFF->_fPositioned && !pCF->_fRelative)
                    {
                        // we are not using GetYTop for to get the offset of the line because
                        // the before space is not added to the height yet.
                        lsme._pDispNodePrev =
                            _pFlowLayout->GetDisplay()->AddLayoutDispNode(
                            _pci,
                            pElementLayout->GetFirstBranch(),
                            xPos,
                            lsme._yli+_li._yBeforeSpace+(_li._yHeight-_li._yExtent)/2,
                            lsme._pDispNodePrev);
                    }
                    else
                    {
                        // If top and bottom or left and right are "auto", position the object
                        xPos += pLayout->GetXProposed();

                        if(fAbsolute)
                        {
                            pLayout->SetYProposed(lsme._cyTopBordPad);
                        }

                        CPoint ptAuto(xPos,
                            lsme._yli+_li._yBeforeSpace+(_li._yHeight-_li._yExtent)/2+pLayout->GetYProposed());

                        pElementLayout->ZChangeElement(0, &ptAuto);
                    }
                }
            }

            if(pElementLayout)
            {
                // setup cchAdvance to skip the current layout
                cchAdvance = lsme.GetNestedElementCch(pElementLayout, &ptp);
                Assert(ptp);
            }

            cch -= cchAdvance;
            lsme.Advance(cchAdvance);
            ptp = ptp->NextTreePos();
            cchAdvance = min(cch, ptp->GetCch());
        }

        // We have just finished measuring, update the line's ascent and descent.
        if(fMeasuring)
        {
            // If we have ALIGN_TYPEABSBOTTOM or ALIGN_TYPETOP, they do not contribute
            // to ascent or descent based on the baseline
            if(yAbsHeight > yAscent + yDescent)
            {
                if(atAbs == htmlControlAlignAbsMiddle)
                {
                    LONG yDiff = yAbsHeight - yAscent - yDescent;
                    yAscent += (yDiff + 1) / 2;
                    yDescent += yDiff / 2;
                }
                else if(atAbs == htmlControlAlignAbsBottom)
                {
                    yAscent = yAbsHeight - yDescent;
                }
                else
                {
                    yDescent = yAbsHeight - yAscent;
                }
            }

            // now update the line height
            _li._yHeight = yAscent + yDescent;
            _li._yDescent = yDescent;

            if(pCF && pCF->HasLineGrid(FALSE))
            {
                LONG yNormalHeight = _li._yHeight;
                _li._yHeight = GetClosestGridMultiple(GetLineGridSize(), _li._yHeight);

                // We don't want to cook up our own line descent if the biggest thing on this 
                // line is just text... our descent has already been calculated for us in that
                // case
                if(yNormalHeight != yTxtAscent+yTxtDescent)
                {
                    _li._yDescent += (_li._yHeight - yNormalHeight)/2 + (_li._yHeight - yNormalHeight)%2;
                }
            }

            // Without this line, line heights specified through
            // styles would override the natural height of the
            // image. This would be cool, but the W3C doesn't
            // like it. Absolute & aligned sites do not affect
            // line height.
            if(_cInlinedSites)
            {
                _fFoundLineHeight = FALSE;
            }

            Assert(_li._yHeight >= 0);

            // Allow last minute adjustment to line height, we need
            // to call this here, because when positioning all the
            // site in line for the display tree, we want the correct
            // YTop.
            AdjustLineHeight();

            // And now the positioning pass
            fMeasuring = FALSE;
            fPositioning = TRUE;
        }
        else
        {
            fPositioning = FALSE;
        }
    }
}

//-----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetRubyInfoFromCp(LONG cpRubyText)
//
//  Synopsis:   Linearly searches through the list of ruby infos and
//              returns the info of the ruby object that contains the given
//              cp.  Note that this function should only be called with a
//              cp that corresponds to a position within Ruby pronunciation
//              text.
//
//              BUGBUG (t-ramar): this code does not take advantage of
//              the fact that this array is sorted by cp, but it does depend 
//              on this fact.  This may be a problem because the entries in this 
//              array are appended in the FetchRubyPosition Line Services callback.
//
//-----------------------------------------------------------------------------
RubyInfo* CLineServices::GetRubyInfoFromCp(LONG cpRubyText)
{
    RubyInfo* pRubyInfo = NULL;
    int i;

    if((RubyInfo*)_aryRubyInfo == NULL)
    {
        goto Cleanup;
    }

    for(i=_aryRubyInfo.Size(),pRubyInfo=_aryRubyInfo; i>0; i--,pRubyInfo++)
    {
        if(pRubyInfo->cp > cpRubyText)
        {
            break;
        }
    }
    pRubyInfo--;

    // if this assert fails, chances are that the cp isn't doesn't correspond
    // to a position within some ruby pronunciation text
    Assert(pRubyInfo >= (RubyInfo*)_aryRubyInfo);

Cleanup:
    return pRubyInfo;
}

//-----------------------------------------------------------------------------
//
//  Member:     CLSMeasurer::AdjustLineHeight()
//
//  Synopsis:   Adjust for space before/after and line spacing rules.
//
//-----------------------------------------------------------------------------
void CLineServices::AdjustLineHeight()
{
    // This had better be true.
    Assert(_li._yHeight >= 0);

    // Need to remember these for hit testing.
    _li._yExtent = _li._yHeight;

    // Only do this if there is a line height used somewhere.
    if(_lMaxLineHeight!=LONG_MIN && _fFoundLineHeight)
    {
        _li._yDescent += (_lMaxLineHeight - _li._yHeight) / 2;
        _li._yHeight = _lMaxLineHeight;
    }
}

//-----------------------------------------------------------------------------
//
// Member:      CLineServices::MeasureLineShift (fZeroLengthLine)
//
// Synopsis:    Computes and returns the line x shift due to alignment
//
//-----------------------------------------------------------------------------
LONG CLineServices::MeasureLineShift(LONG cp, LONG xWidthMax, BOOL fMinMax, LONG* pdxRemainder)
{
    long xShift;
    UINT uJustified;

    Assert(_li._fRTL == (unsigned)_pPFFirst->HasRTL(_fInnerPFFirst));

    xShift = ComputeLineShift(
        (htmlAlign)_pPFFirst->GetBlockAlign(_fInnerPFFirst),
        _pFlowLayout->GetDisplay()->IsRTL(),
        _li._fRTL,
        fMinMax,
        xWidthMax,
        _li._xLeft+_li._xWidth+_li._xLineOverhang+_li._xRight,
        &uJustified,
        pdxRemainder);
    _li._fJustified = uJustified;
    return xShift;
}

//-----------------------------------------------------------------------------
//
// Member:      CalculateXPositionOfLSCP
//
// Synopsis:    Calculates the X position for LSCP
//
//-----------------------------------------------------------------------------
LONG CLineServices::CalculateXPositionOfLSCP(
                                        LSCP lscp,          // LSCP to return the position of.
                                        BOOL fAfterPrevCp,  // Return the trailing point of the previous LSCP (for an ambigous bidi cp)
                                        BOOL* pfRTLFlow)    // Flow direction of LSCP.
{
    LSTEXTCELL lsTextCell;
    HRESULT hr;
    BOOL fRTLFlow = FALSE;
    BOOL fUsePrevLSCP = FALSE;
    LONG xRet;

    if(fAfterPrevCp && _pBidiLine != NULL)
    {
        LSCP lscpPrev = FindPrevLSCP(lscp, &fUsePrevLSCP);
        if(fUsePrevLSCP)
        {
            lscp = lscpPrev;
        }
    }

    hr = QueryLineCpPpoint(lscp, FALSE, NULL, &lsTextCell, &fRTLFlow);

    if(pfRTLFlow)
    {
        *pfRTLFlow = fRTLFlow;
    }

    xRet = lsTextCell.pointUvStartCell.u;

    // If we're querying for a character which cannot be measured (e.g. a
    // section break char), then LS returns the last character it could
    // measure.  To get the x-position, we add the width of this character.
    if (S_OK==hr && (lsTextCell.cpEndCell<lscp || fUsePrevLSCP))
    {
        if((unsigned)fRTLFlow == _li._fRTL)
        {
            xRet += lsTextCell.dupCell;
        }
        else
        {
            xRet -= lsTextCell.dupCell;
            // What is happening here is that we are being positioned at say pixel
            // pos 10 (xRet=10) and are asked to draw reverese a character which is
            // 11 px wide. So we would then draw at px10, at px9 ... and finally at
            // px 0 -- for at grand total of 11 px. Having drawn at 0, we would be
            // put back at -1. While the going back by 1px is correct, at the BOL
            // this will put us at -1, which is unaccepatble and hence the max with 0.
            xRet = max(0L, xRet);
        }
    }
    else if(hr==S_OK && lsTextCell.cCharsInCell>1 && lscp>lsTextCell.cpStartCell)
    {
        long lClusterAdjust = MulDivQuick(lscp-lsTextCell.cpStartCell,
            lsTextCell.dupCell, lsTextCell.cCharsInCell);
        // we have multiple cps mapped to one glyph. This simply places the caret
        // a percentage of space between beginning and end
        if((unsigned)fRTLFlow == _li._fRTL)
        {
            xRet += lClusterAdjust;
        }
        else
        {
            xRet -= lClusterAdjust;
        }
    }

    Assert(hr || xRet>=0);

    return hr?0:xRet;
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::CalcPositionsOfRangeOnLine
//
// Synopsis:    Find the position of a stretch of text starting at cpStart and
//              and running to cpEnd, inclusive. The text may be broken into
//              multiple rects if the line has reverse objects (runs with
//              mixed directionallity) in it.
//
// Returns:     The number of chunks in the range. Usually this will just be
//              one. If an error occurs it will be zero. The actual width of
//              text (in device units) is returned in paryChunks as rects from
//              the beginning of the line. The top and bottom entries of each
//              rect will be 0. No assumptions should be made about the order
//              of the rects; the first rect may or may not be the leftmost or
//              rightmost.
//
//-----------------------------------------------------------------------------
LONG CLineServices::CalcPositionsOfRangeOnLine(
    LONG cpStart,
    LONG cpEnd,
    LONG xShift,
    CDataAry<RECT>* paryChunks)
{
    CStackDataAry<LSQSUBINFO, 4> aryLsqsubinfo;
    LSTEXTCELL lsTextCell;
    LSCP lscpStart = LSCPFromCP(max(cpStart, _cpStart));
    LSCP lscpEnd = LSCPFromCP(cpEnd);
    HRESULT hr;
    BOOL fSublineReverseFlow = FALSE;
    LONG xStart;
    LONG xEnd;
    RECT rcChunk;
    LONG i;
    LSTFLOW tflow = (!_li._fRTL ? lstflowES : lstflowWS);

    Assert(paryChunks!=NULL && paryChunks->Size()==0);
    Assert(cpStart <= cpEnd);

    rcChunk.top = rcChunk.bottom = 0;

    aryLsqsubinfo.Grow(4); // Guaranteed to succeed since we're working from the stack.

    hr = QueryLineCpPpoint(lscpStart, FALSE, &aryLsqsubinfo, &lsTextCell, FALSE);

    xStart = lsTextCell.pointUvStartCell.u;

    for(i=aryLsqsubinfo.Size()-1; i>=0; i--)
    {
        const LSQSUBINFO& qsubinfo = aryLsqsubinfo[i];
        const LSQSUBINFO& qsubinfoParent = aryLsqsubinfo[max((LONG)(i-1), 0L)];

        if(lscpEnd < (LSCP)(qsubinfo.cpFirstSubline+qsubinfo.dcpSubline))
        {
            // lscpEnd is in this subline. Break out.
            break;
        }

        // If the subline and its parent are going in different directions
        // stuff the current range into the chunk array and move xStart to
        // the "end" (relative to the parent) of the current subline.
        if((qsubinfo.lstflowSubline&fUDirection) != (qsubinfoParent.lstflowSubline&fUDirection))
        {
            // Append the start of the chunk to the chunk array.
            rcChunk.left = xShift + xStart;

            fSublineReverseFlow = !!((qsubinfo.lstflowSubline^tflow) & fUDirection);

            // Append the end of the chunk to the chunk array.
            // If the subline flow doesn't match the line direction then we're
            // moving against the flow of the line and we will subtract the
            // subline width from the subline start to find the end point.
            rcChunk.right = xShift + qsubinfo.pointUvStartSubline.u +
                (fSublineReverseFlow?-qsubinfo.dupSubline:qsubinfo.dupSubline);

            // do some reverse flow cleanup before inserting rect into the array
            if(rcChunk.left > rcChunk.right)
            {
                Assert(fSublineReverseFlow);
                long temp = rcChunk.left;
                rcChunk.left = rcChunk.right + 1;
                rcChunk.right = temp + 1;
            }

            paryChunks->AppendIndirect(&rcChunk);

            xStart = qsubinfo.pointUvStartSubline.u + (fSublineReverseFlow?1:-1);
        }
    }

    aryLsqsubinfo.Grow(4); // Guaranteed to succeed since we're working from the stack.

    hr = QueryLineCpPpoint(lscpEnd, FALSE, &aryLsqsubinfo, &lsTextCell, FALSE);

    xEnd = lsTextCell.pointUvStartCell.u;
    if(lscpEnd > lsTextCell.cpEndCell)
    {
        xEnd += ((aryLsqsubinfo.Size()==0 ||
            !((aryLsqsubinfo[aryLsqsubinfo.Size()-1].lstflowSubline^tflow)&fUDirection))?lsTextCell.dupCell:-lsTextCell.dupCell);
    }

    for(i=aryLsqsubinfo.Size()-1; i>=0; i--)
    {
        const LSQSUBINFO& qsubinfo = aryLsqsubinfo[i];
        const LSQSUBINFO& qsubinfoParent = aryLsqsubinfo[max((LONG)(i-1), 0L)];

        if(lscpStart >= qsubinfo.cpFirstSubline)
        {
            // lscpStart is in this subline. Break out.
            break;
        }

        // If the subline and its parent are going in different directions
        // stuff the current range into the chunk array and move xEnd to
        // the "start" (relative to the parent) of the current subline.
        if((qsubinfo.lstflowSubline&fUDirection) != (qsubinfoParent.lstflowSubline&fUDirection))
        {
            fSublineReverseFlow = !!((qsubinfo.lstflowSubline^tflow) & fUDirection);

            if(xEnd != qsubinfo.pointUvStartSubline.u)
            {
                // Append the start of the chunk to the chunk array.
                rcChunk.left = xShift + qsubinfo.pointUvStartSubline.u;

                // Append the end of the chunk to the chunk array.
                rcChunk.right = xShift + xEnd;

                // do some reverse flow cleanup before inserting rect into the array
                if(rcChunk.left > rcChunk.right)
                {
                    Assert(fSublineReverseFlow);
                    long temp = rcChunk.left;
                    rcChunk.left = rcChunk.right + 1;
                    rcChunk.right = temp + 1;
                }

                paryChunks->AppendIndirect(&rcChunk);
            }

            // If the subline flow doesn't match the line direction then we're
            // moving against the flow of the line and we will subtract the
            // subline width from the subline start to find the end point.
            xEnd = qsubinfo.pointUvStartSubline.u +
                (fSublineReverseFlow?-(qsubinfo.dupSubline-1):(qsubinfo.dupSubline-1));
        }
    }

    rcChunk.left = xShift + xStart;
    rcChunk.right = xShift + xEnd;
    // do some reverse flow cleanup before inserting rect into the array
    if(rcChunk.left > rcChunk.right)
    {
        long temp = rcChunk.left;
        rcChunk.left = rcChunk.right + 1;
        rcChunk.right = temp + 1;
    }
    paryChunks->AppendIndirect(&rcChunk);

    return paryChunks->Size();
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::CalcRectsOfRangeOnLine
//
// Synopsis:    Find the position of a stretch of text starting at cpStart and
//              and running to cpEnd, inclusive. The text may be broken into
//              multiple runs if different font sizes or styles are used, or there
//              is mixed directionallity in it.
//
// Returns:     The number of chunks in the range. Usually this will just be
//              one. If an error occurs it will be zero. The actual width of
//              text (in device units) is returned in paryChunks as rects of
//              offsets from the beginning of the line. 
//              No assumptions should be made about the order of the chunks;
//              the first chunk may or may not be the chunk which includes
//              cpStart.
//
//-----------------------------------------------------------------------------
LONG CLineServices::CalcRectsOfRangeOnLine(
                                      LONG cpStart,
                                      LONG cpEnd,
                                      LONG xShift,
                                      LONG yPos,
                                      CDataAry<RECT>* paryChunks,
                                      DWORD dwFlags)
{
    CStackDataAry<LSQSUBINFO, 4> aryLsqsubinfo;
    HRESULT hr;
    LSTEXTCELL lsTextCell;
    LSCP lscpRunStart = LSCPFromCP(max(cpStart, _cpStart));
    LSCP lscpEnd = min(LSCPFromCP(cpEnd), _lscpLim);
    BOOL fSublineReverseFlow;
    LSTFLOW tflow = (!_li._fRTL ? lstflowES : lstflowWS);
    LONG xStart;
    LONG xEnd;
    LONG yTop = 0;
    LONG yBottom = 0;
    RECT rcChunk;
    RECT rcLast = { 0 };
    COneRun* porCurrent = _listCurrent._pHead;

    // we should never come in here with an LSCP that is in the middle of a COneRun. Those types (for selection)
    // should go through CalcPositionsOfRangeOnLine.

    // move quickly to the por that has the right lscpstart
    while(porCurrent->_lscpBase < lscpRunStart)
    {
        porCurrent = porCurrent->_pNext;

        // If we assert here, something is messed up. Please investigate
        Assert(porCurrent != NULL);
        // if we reached the end of the list we need to bail out.
        if(porCurrent == NULL)
        {
            return paryChunks->Size();
        }
    }

    // for selection we want to start highlight invalidation at beginning of the run
    // to avoid vertical line turds with RTL text.
    if(dwFlags & RFE_SELECTION)
    {
        lscpRunStart = porCurrent->_lscpBase;
    }

    Assert(paryChunks!=NULL && paryChunks->Size()==0);
    Assert(cpStart <= cpEnd);

    while(lscpRunStart < lscpEnd)
    {
        // if we reached the end of the list we need to bail out.
        if(porCurrent == NULL)
        {
            return paryChunks->Size();
        }

        switch(porCurrent->_fType)
        {
        case ONERUN_NORMAL:
            {
                aryLsqsubinfo.Grow(4); // Guaranteed to succeed since we're working from the stack.

                hr = QueryLineCpPpoint(lscpRunStart, FALSE, &aryLsqsubinfo, &lsTextCell, FALSE);

                // 1. If we failed, return what we have so far
                // 2. If LS returns less than the cell we have asked for we have a problem with LS.
                //    To solve this problem (especially for Vietnamese) we will go to the next
                //    COneRun in the list and try again.
                if(hr || lsTextCell.cpStartCell<lscpRunStart)
                {
                    lscpRunStart = porCurrent->_lscpBase + porCurrent->_lscch;
                    break;
                }

                long  nDepth = aryLsqsubinfo.Size() - 1;
                Assert(nDepth >= 0);
                const LSQSUBINFO& qsubinfo = aryLsqsubinfo[nDepth];
                long lAscent = qsubinfo.heightsPresRun.dvAscent;

                // now set the end position based on which way the subline flows
                fSublineReverseFlow = !!((qsubinfo.lstflowSubline^tflow) & fUDirection);

                xStart = lsTextCell.pointUvStartCell.u + (!_li._fRTL ? 0 : -1);

                // If we're querying for a character which cannot be measured (e.g. a
                // section break char), then LS returns the last character it could
                // measure. Therefore, if we are okay, use dupRun for the distance.
                // Otherwise, query the last character of porCurrent. This prevents us
                // having to loop when LS creates dobj's that are cch of five (5).
                if((LSCP)(porCurrent->_lscpBase+porCurrent->_lscch) <= (LSCP)(qsubinfo.cpFirstRun+qsubinfo.dcpRun))
                {
                    xEnd = xStart +
                        (!fSublineReverseFlow?qsubinfo.dupRun:-qsubinfo.dupRun);
                }
                else
                {
                    aryLsqsubinfo.Grow(4); // Guaranteed to succeed since we're working from the stack.
                    long lscpLast = min(lscpRunStart+porCurrent->_lscch, _lscpLim);

                    hr = QueryLineCpPpoint(lscpLast, FALSE, &aryLsqsubinfo, &lsTextCell, FALSE);

                    // BUGBUG: (paulnel) LineServices ignores hyphens as a part of the line...even though they say it is in the subline
                    // if we querried on a valid lscp, yet LS give us back a lesser value, add the width of the last cell given.
                    xEnd = lsTextCell.pointUvStartCell.u + (!_li._fRTL?0:-1) + 
                        (lsTextCell.cpStartCell>=lscpLast?0:lsTextCell.dupCell);
                }

                // get the top and bottom for the rect
                if(porCurrent->_fCharsForNestedLayout)
                {
                    // NOTICE: Absolutely positioned, aligned, and Bold elements are ONERUN_ANTISYNTH types. See note below
                    RECT rc;
                    long cyHeight;
                    const CCharFormat* pCF = porCurrent->GetCF();
                    CTreeNode* pNodeCur = porCurrent->_ptp->Branch();
                    CLayout* pLayout = pNodeCur->GetUpdatedLayout();

                    pLayout->GetRect(&rc);

                    cyHeight = rc.bottom - rc.top;

                    // XProposed and YProposed have been set to the amount of margin
                    // the layout has.                
                    xStart = pLayout->GetXProposed();
                    xEnd = xStart + (rc.right - rc.left);

                    yTop = yPos + _pMeasurer->_li._yBeforeSpace + pLayout->GetYProposed();
                    yBottom = yTop + cyHeight;

                    // take care of any nested relatively positioned elements
                    if(pCF->_fRelative && (dwFlags&RFE_NESTED_REL_RECTS))
                    {
                        long xRelLeft=0, yRelTop=0;

                        // get the layout's relative positioning to its parent. The parent's relative
                        // positioning would be adjusted in RegionFromElement
                        CTreeNode* pNodeParent = pNodeCur->Parent();
                        if(pNodeParent)
                        {
                            pNodeCur->GetRelTopLeft(pNodeParent->Element(), _pci, &xRelLeft, &yRelTop);
                        }

                        xStart += xRelLeft;
                        xEnd += xRelLeft;
                        yTop += yRelTop;
                        yBottom += yRelTop;
                    }

                }
                else
                {
                    const CCharFormat* pCF = porCurrent->GetCF();

                    // The current character does not have height. Throw it out.
                    if(lAscent == 0)
                    {
                        break;
                    }

                    yBottom = yPos + _pMeasurer->_li._yHeight 
                        - _pMeasurer->_li._yDescent + _pMeasurer->_li._yTxtDescent;

                    yTop = yBottom - _pMeasurer->_li._yTxtDescent - lAscent;

                    // If we are ruby text, adjust the height to the correct position above the line.
                    if(pCF->_fIsRubyText)
                    {
                        RubyInfo* pRubyInfo = GetRubyInfoFromCp(porCurrent->_lscpBase);

                        if(pRubyInfo)
                        {
                            yBottom = yPos + _pMeasurer->_li._yHeight - _pMeasurer->_li._yDescent 
                                + pRubyInfo->yDescentRubyBase - pRubyInfo->yHeightRubyBase;
                            yTop = yBottom - pRubyInfo->yDescentRubyText - lAscent;
                        }
                    }
                }

                rcChunk.left = xShift + xStart;
                rcChunk.top = yTop;
                rcChunk.bottom = yBottom;
                rcChunk.right = xShift + xEnd;

                // do some reverse flow cleanup before inserting rect into the array
                if(rcChunk.left > rcChunk.right)
                {
                    long temp = rcChunk.left;
                    rcChunk.left = rcChunk.right;
                    rcChunk.right = temp;
                }

                // In the event we have <A href="x"><B>text</B></A> we get two runs of
                // the same rect. One for the Anchor and one for the bold. These two
                // rects will xor themselves when drawing the wiggly and look like they
                // did not select. This patch resolves this issue for the time being.
                if(!(rcChunk.left==rcLast.left && 
                    rcChunk.right==rcLast.right &&
                    rcChunk.top==rcLast.top &&
                    rcChunk.bottom==rcLast.bottom))
                {
                    paryChunks->AppendIndirect(&rcChunk);
                }

                rcLast = rcChunk;

                lscpRunStart = porCurrent->_lscpBase + porCurrent->_lscch;
            }
            break;

        case ONERUN_SYNTHETIC:
            // We want to set the lscpRunStart to move to the next start position
            // when dealing with synthetics (reverse objects, etc.)
            lscpRunStart = porCurrent->_lscpBase + porCurrent->_lscch;
            break;

        case ONERUN_ANTISYNTH:
            // NOTICE:
            // this case covers absolutely positioned elements and aligned elements
            // However, per BrendanD and SujalP, this is not the correct place
            // to implement focus rects for these elements. Some
            // work needs to be done to the CAdorner to properly 
            // handle absolutely positioned elements. RegionFromElement should handle
            // frames for aligned objects.
            break;

        default:
            Assert("Missing COneRun type");
            break;
        }

        porCurrent = porCurrent->_pNext;
    }

    return paryChunks->Size();
}

//-----------------------------------------------------------------------------
//
// Member:      CLineServices::RecalcLineHeight()
//
// Synopsis:    Reset the height of the the line we are measuring if the new
//              run of text is taller than the current maximum in the line.
//
//-----------------------------------------------------------------------------
void CLineServices::RecalcLineHeight(CCcs* pccs, CLine* pli)
{
    AssertSz(pli,  "we better have a line!");
    AssertSz(pccs, "we better have a some metric's here");

    if(pccs)
    {
        SHORT yAscent;
        SHORT yDescent;

        pccs->GetBaseCcs()->GetAscentDescent(&yAscent, &yDescent);

        if(yAscent < pli->_yHeight-pli->_yDescent)
        {
            yAscent = pli->_yHeight - pli->_yDescent;
        }
        if(yDescent > pli->_yDescent)
        {
            pli->_yDescent = yDescent;
        }

        pli->_yHeight = yAscent + pli->_yDescent;

        Assert(pli->_yHeight >= 0);
    }
}

//-----------------------------------------------------------------------------
//
// Member:      TestFortransitionToOrFromRelativeChunk
//
// Synopsis:    Test if we are transitioning from a relative chunk to normal
//                chunk or viceversa
//
//-----------------------------------------------------------------------------
void CLineServices::TestForTransitionToOrFromRelativeChunk(
    CLSMeasurer& lsme,
    BOOL fRelative,
    CElement* pElementLayout)
{
    CTreeNode* pNodeRelative = NULL;
    CElement* pElementFL = _pFlowLayout->ElementOwner();

    // if the current line is relative and the chunk is not
    // relative or if the current line is not relative and the
    // the current chunk is relative then break out.
    if(fRelative)
    {
        pNodeRelative = lsme.CurrBranch()->GetCurrentRelativeNode(pElementFL);
    }
    if(DifferentScope(_pElementLastRelative, pNodeRelative))
    {
        UpdateLastChunkInfo(lsme, pNodeRelative->SafeElement(), pElementLayout);
    }
}

//-----------------------------------------------------------------------------
//
// Member:        UpdateLastChunkInfo
//
// Synopsis:    We have just transitioned from a relative chunk to normal chunk
//              or viceversa, or inbetween relative chunks, so update the last
//              chunk information.
//
//-----------------------------------------------------------------------------
void CLineServices::UpdateLastChunkInfo(CLSMeasurer& lsme, CElement* pElementRelative, CElement* pElementLayout)
{
    LONG xPosCurrChunk;

    if(long(lsme.GetCp()) > _cpLastChunk)
    {
        xPosCurrChunk = AddNewChunk(lsme.GetCp());
    }
    else
    {
        xPosCurrChunk = _xPosLastChunk;
    }

    _cpLastChunk = lsme.GetCp();
    _xPosLastChunk = xPosCurrChunk;
    _pElementLastRelative = pElementRelative;
    _fLastChunkSingleSite = pElementLayout && pElementLayout->IsOwnLineElement(_pFlowLayout);
    _fLastChunkHasBulletOrNum = pElementRelative && pElementRelative->IsTagAndBlock(ETAG_LI);
}

//-----------------------------------------------------------------------------
//
// Member:      AddNewChunk
//
// Synopsis:    Adds a new chunk of either relative or non-relative text to
//              the line.
//
//-----------------------------------------------------------------------------
LONG CLineServices::AddNewChunk(LONG cp)
{
    BOOL fRTLFlow = FALSE;
    BOOL fRTLDisplay = _pFlowLayout->GetDisplay()->IsRTL();
    CLSLineChunk* plcNew = new CLSLineChunk();
    LONG xPosCurrChunk = CalculateXPositionOfCp(cp, FALSE, &fRTLFlow);

    // BUGBUG (PaulNel) This is only a bandaid. Some serious consideration needs to be
    // given to handling relative text areas that have flows in different directions.
    // It is not too difficult to come up with a case where chunks would be discontiguous.
    // How should that be handled? The coding to take care of this situation will be
    // significantly intrusive, but should be taken care of, nonetheless. Perhaps IE6
    // timeframe will be best to address this issue.
    if(_li._fRTL != (unsigned)fRTLFlow)
    {
        AdjustXPosOfCp(fRTLDisplay, _li._fRTL, fRTLFlow, &xPosCurrChunk, _li._xWidth);
        if(_li._fRTL)
        {
            xPosCurrChunk--;
        }
        // In a RTL display, 0 is on the right. Here we don't want to put a negative
        // value into the chunk's width, so change the sign.
        if(fRTLDisplay)
        {
            xPosCurrChunk = -xPosCurrChunk;
        }
    }

    if(!plcNew)
    {
        goto Cleanup;
    }

    plcNew->_cch             = cp  - _cpLastChunk;
    plcNew->_xWidth          = xPosCurrChunk - _xPosLastChunk;
    plcNew->_fRelative       = _pElementLastRelative != NULL;
    plcNew->_fSingleSite     = _fLastChunkSingleSite;
    plcNew->_fHasBulletOrNum = _fLastChunkHasBulletOrNum;

    // append this chunk to the line
    if(_plcLastChunk)
    {
        Assert(!_plcLastChunk->_plcNext);

        _plcLastChunk->_plcNext = plcNew;
        _plcLastChunk = plcNew;
    }
    else
    {
        _plcLastChunk = _plcFirstChunk = plcNew;
    }

Cleanup:
    return xPosCurrChunk;
}

//-----------------------------------------------------------------------------
//
// Member:      WidthOfChunksSoFarCore
//
// Synopsis:    This function finds the width of all the chunks collected so far.
//              This is needed because we need to position sites in a line relative
//              to the BOL. So if a line is going to be chunked up (because of
//              relative stuff before it), then we need to subtract their total width.
//
//-----------------------------------------------------------------------------
LONG CLineServices::WidthOfChunksSoFarCore()
{
    CLSLineChunk* plc = _plcFirstChunk;
    LONG xWidth = 0;

    Assert(plc);
    while(plc)
    {
        xWidth += plc->_xWidth;
        plc = plc->_plcNext;
    }
    return xWidth;
}

//-----------------------------------------------------------------------------
//
// Member:      AdjustXPosOfCp()
//
// Synopsis:    used to adjust xPosofCp when flow is opposite direction
//
//-----------------------------------------------------------------------------
void CLineServices::AdjustXPosOfCp(BOOL fRTLDisplay, BOOL fRTLLine, BOOL fRTLFlow, long* pxPosOfCp, long xLayoutWidth)
{
    if(fRTLDisplay != fRTLLine)
    {
        *pxPosOfCp = _li._xWidth - *pxPosOfCp;
    }
    if(fRTLDisplay != fRTLFlow)
    {
        *pxPosOfCp -= xLayoutWidth - 1;
    }
}

//-----------------------------------------------------------------------------
//
// Member:      GetKashidaWidth()
//
// Synopsis:    gets the width of the kashida character (U+0640) for Arabic
//              justification
//
//-----------------------------------------------------------------------------
LSERR CLineServices::GetKashidaWidth(PLSRUN plsrun, int* piKashidaWidth)
{
    LSERR lserr = lserrNone;
    HRESULT hr = S_OK;
    HDC hdc = _pci->_hdc;
    HFONT hfontOld = NULL;
    HDC hdcFontProp = NULL;
    CCcs* pccs = NULL;
    CBaseCcs* pBaseCcs = NULL;
    SCRIPT_CACHE* psc = NULL;

    SCRIPT_FONTPROPERTIES  sfp;
    sfp.cBytes = sizeof(SCRIPT_FONTPROPERTIES);

    pccs = GetCcs(plsrun, _pci->_hdc, _pci);
    if(pccs == NULL)
    {
        lserr = lserrOutOfMemory;
        goto Cleanup;
    }
    pBaseCcs = pccs->GetBaseCcs();

    psc = pBaseCcs->GetUniscribeCache();
    Assert(psc != NULL);

    hr = ScriptGetFontProperties(hdcFontProp, psc, &sfp);

    AssertSz(hr!=E_INVALIDARG, "You might need to update USP10.DLL");

    // Handle failure
    if(hr == E_PENDING)
    {
        Assert(hdcFontProp == NULL);

        // Select the current font into the hdc and set hdcFontProp to hdc.
        hfontOld = SelectFontEx(hdc, pBaseCcs->_hfont);
        hdcFontProp = hdc;

        hr = ScriptGetFontProperties(hdcFontProp, psc, &sfp);

    }
    Assert(hr==S_OK || hr==E_OUTOFMEMORY);

    lserr = LSERRFromHR(hr);

    if(lserr == lserrNone)
    {
        *piKashidaWidth = max(sfp.iKashidaWidth, 1);
    }

Cleanup:
    // Restore the font if we selected it
    if(hfontOld != NULL)
    {
        SelectFontEx(hdc, hfontOld);
    }

    return lserr;
}

// The following tables depend on the fact that there is no more
// than a single alternate for each MWCLS and that at most one
// condition needs to be taken into account in resolving each MWCLS

// NB (cthrash) This is a packed table.  The first three elements (brkcls,
// brkclsAlt and brkopt) are indexed by CHAR_CLASS.  The fourth column we
// access (brkclsLow) we access by char value.  The fourth column is for a
// speed optimization.
#if (defined(_MSC_VER) && (_MSC_VER>=1200))
#define BRKINFO(a,b,c,d) { CLineServices::a, CLineServices::b, CLineServices::c, CLineServices::d }
#else
#define BRKINFO(a,b,c,d) { DWORD(CLineServices::a), DWORD(CLineServices::b), DWORD(CLineServices::c), DWORD(CLineServices::d) }
#endif

const CLineServices::PACKEDBRKINFO CLineServices::s_rgBrkInfo[CHAR_CLASS_MAX] =
{
    //       brkcls             brkclsAlt         brkopt     brkclsLow                 CC     (QPID)
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   0 WOB_(1)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   1 NOPP(2)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   2 NOPA(2)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   3 NOPW(2)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   4 HOP_(3)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   5 WOP_(4)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsAlpha        ), //   6 WOP5(5)  
    BRKINFO( brkclsQuote,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //   7 NOQW(6)  
    BRKINFO( brkclsQuote,       brkclsOpen,       fCscWide,  brkclsAlpha        ), //   8 AOQW(7)  
    BRKINFO( brkclsOpen,        brkclsNil,        fBrkNone,  brkclsSpaceN       ), //   9 WOQ_(8)  
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsSpaceN       ), //  10 WCB_(9)  
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  11 NCPP(10) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsSpaceN       ), //  12 NCPA(10) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsSpaceN       ), //  13 NCPW(10) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  14 HCP_(11) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  15 WCP_(12) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  16 WCP5(13) 
    BRKINFO( brkclsQuote,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  17 NCQW(14) 
    BRKINFO( brkclsQuote,       brkclsClose,      fCscWide,  brkclsAlpha        ), //  18 ACQW(15) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  19 WCQ_(16) 
    BRKINFO( brkclsQuote,       brkclsClose,      fCscWide,  brkclsAlpha        ), //  20 ARQW(17) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsAlpha        ), //  21 NCSA(18) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  22 HCO_(19) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  23 WC__(20) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  24 WCS_(20)
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  25 WC5_(21) 
    BRKINFO( brkclsClose,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  26 WC5S(21)
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  27 NKS_(22) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  28 WKSM(23) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  29 WIM_(24) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  30 NSSW(25) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  31 WSS_(26) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAsciiSpace   ), //  32 WHIM(27) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsExclaInterr  ), //  33 WKIM(28) 
    BRKINFO( brkclsIdeographic, brkclsNoStartIdeo,fBrkStrict,brkclsQuote        ), //  34 NKSL(29) 
    BRKINFO( brkclsIdeographic, brkclsNoStartIdeo,fBrkStrict,brkclsAlpha        ), //  35 WKS_(30) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsPrefix       ), //  36 WKSC(30) 
    BRKINFO( brkclsIdeographic, brkclsNoStartIdeo,fBrkStrict,brkclsPostfix      ), //  37 WHS_(31) 
    BRKINFO( brkclsExclaInterr, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  38 NQFP(32) 
    BRKINFO( brkclsExclaInterr, brkclsNil,        fBrkNone,  brkclsQuote        ), //  39 NQFA(32) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsOpen         ), //  40 WQE_(33) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsClose        ), //  41 WQE5(34) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  42 NKCC(35) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  43 WKC_(36) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsNumSeparator ), //  44 NOCP(37) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsSpaceN       ), //  45 NOCA(37) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsNumSeparator ), //  46 NOCW(37) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsSlash        ), //  47 WOC_(38) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  48 WOCS(38)
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  49 WOC5(39) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  50 WOC6(39)
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  51 AHPW(40) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsNumeral      ), //  52 NPEP(41) 
    BRKINFO( brkclsNumSeparator,brkclsNil,        fBrkNone,  brkclsNumeral      ), //  53 NPAR(41) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  54 HPE_(42) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  55 WPE_(43) 
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  56 WPES(43)
    BRKINFO( brkclsNoStartIdeo, brkclsNil,        fBrkNone,  brkclsNumeral      ), //  57 WPE5(44) 
    BRKINFO( brkclsInseparable, brkclsNil,        fBrkNone,  brkclsNumSeparator ), //  58 NISW(45) 
    BRKINFO( brkclsInseparable, brkclsNil,        fBrkNone,  brkclsNumSeparator ), //  59 AISW(46) 
    BRKINFO( brkclsGlueA,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  60 NQCS(47) 
    BRKINFO( brkclsGlueA,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  61 NQCW(47) 
    BRKINFO( brkclsGlueA,       brkclsNil,        fBrkNone,  brkclsAlpha        ), //  62 NQCC(47) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsExclaInterr  ), //  63 NPTA(48) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  64 NPNA(48) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  65 NPEW(48) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  66 NPEH(48) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  67 APNW(49) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  68 HPEW(50) 
    BRKINFO( brkclsPrefix,      brkclsNil,        fBrkNone,  brkclsAlpha        ), //  69 WPR_(51) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  70 NQEP(52) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  71 NQEW(52) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  72 NQNW(52) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  73 AQEW(53) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  74 AQNW(53) 
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  75 AQLW(53)
    BRKINFO( brkclsPostfix,     brkclsNil,        fBrkNone,  brkclsAlpha        ), //  76 WQO_(54) 
    BRKINFO( brkclsAsciiSpace,  brkclsNil,        fBrkNone,  brkclsAlpha        ), //  77 NSBL(55) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  78 WSP_(56) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  79 WHI_(57) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  80 NKA_(58) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  81 WKA_(59) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  82 ASNW(60) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  83 ASEW(60) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  84 ASRN(60) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  85 ASEN(60)
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  86 ALA_(61) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  87 AGR_(62) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), //  88 ACY_(63) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  89 WID_(64) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  90 WPUA(65) 
    BRKINFO( brkclsHangul,      brkclsNil,        fBrkNone,  brkclsOpen         ), //  91 NHG_(66) 
    BRKINFO( brkclsHangul,      brkclsNil,        fBrkNone,  brkclsPrefix       ), //  92 WHG_(67) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsClose        ), //  93 WCI_(68) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  94 NOI_(69) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), //  95 WOI_(70) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), //  96 WOIC(70) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), //  97 WOIL(70)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), //  98 WOIS(70)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), //  99 WOIT(70)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 100 NSEN(71) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 101 NSET(71) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 102 NSNW(71) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 103 ASAN(72) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 104 ASAE(72) 
    BRKINFO( brkclsNumeral,     brkclsNil,        fBrkNone,  brkclsAlpha        ), // 105 NDEA(73) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), // 106 WD__(74) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 107 NLLA(75) 
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsAlpha        ), // 108 WLA_(76) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 109 NWBL(77) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 110 NWZW(77) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 111 NPLW(78) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 112 NPZW(78) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 113 NPF_(78) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 114 NPFL(78)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 115 NPNW(78) 
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), // 116 APLW(79) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 117 APCO(79) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 118 ASYW(80) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 119 NHYP(81) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 120 NHYW(81) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 121 AHYW(82) 
    BRKINFO( brkclsQuote,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 122 NAPA(83) 
    BRKINFO( brkclsQuote,       brkclsNil,        fBrkNone,  brkclsOpen         ), // 123 NQMP(84) 
    BRKINFO( brkclsSlash,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 124 NSLS(85) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsClose        ), // 125 NSF_(86) 
    BRKINFO( brkclsSpaceN,      brkclsNil,        fBrkNone,  brkclsAlpha        ), // 126 NSBS(86) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 127 NLA_(87) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 128 NLQ_(88) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 129 NLQN(88)
    BRKINFO( brkclsAlpha,       brkclsIdeographic,fCscWide,  brkclsAlpha        ), // 130 ALQ_(89) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 131 NGR_(90) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 132 NGRN(90)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 133 NGQ_(91) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 134 NGQN(91)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 135 NCY_(92) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 136 NCYP(93) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 137 NCYC(93) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 138 NAR_(94) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 139 NAQN(95) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 140 NHB_(96) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 141 NHBC(96) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 142 NHBW(96) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 143 NHBR(96) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 144 NASR(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 145 NAAR(97) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 146 NAAC(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 147 NAAD(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 148 NAED(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 149 NANW(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 150 NAEW(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 151 NAAS(97) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 152 NHI_(98) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 153 NHIN(98) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 154 NHIC(98) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 155 NHID(98) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 156 NBE_(99) 
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsAlpha        ), // 157 NBEC(99) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 158 NBED(99) 
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsAlpha        ), // 159 NGM_(100)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsGlueA        ), // 160 NGMC(100)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 161 NGMD(100)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 162 NGJ_(101)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 163 NGJC(101)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 164 NGJD(101)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 165 NOR_(102)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 166 NORC(102)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 167 NORD(102)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 168 NTA_(103)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 169 NTAC(103)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 170 NTAD(103)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 171 NTE_(104)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 172 NTEC(104)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 173 NTED(104)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 174 NKD_(105)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 175 NKDC(105)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 176 NKDD(105)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 177 NMA_(106)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 178 NMAC(106)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 179 NMAD(106)
    BRKINFO( brkclsThaiFirst,   brkclsNil,        fBrkNone,  brkclsNil          ), // 180 NTH_(107)
    BRKINFO( brkclsThaiFirst,   brkclsNil,        fBrkNone,  brkclsNil          ), // 181 NTHC(107)
    BRKINFO( brkclsThaiFirst,   brkclsNil,        fBrkNone,  brkclsNil          ), // 182 NTHD(107)
    BRKINFO( brkclsThaiFirst,   brkclsNil,        fBrkNone,  brkclsNil          ), // 183 NTHT(107)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 184 NLO_(108)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 185 NLOC(108)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 186 NLOD(108)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 187 NTI_(109)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 188 NTIC(109)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 189 NTID(109)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 190 NGE_(110)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 191 NGEQ(111)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 192 NBO_(112)
    BRKINFO( brkclsGlueA,       brkclsNil,        fBrkNone,  brkclsNil          ), // 193 NBSP(113)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 194 NOF_(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 195 NOBS(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 196 NOEA(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 197 NONA(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 198 NONP(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 199 NOEP(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 200 NONW(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 201 NOEW(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 202 NOLW(114)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 203 NOCO(114)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 204 NOSP(114)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 205 NOEN(114)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 206 NET_(115)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 207 NCA_(116)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 208 NCH_(117)
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsNil          ), // 209 WYI_(118)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 210 NBR_(119)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 211 NRU_(120)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 212 NOG_(121)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 213 NSI_(122)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 214 NSIC(122)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 215 NTN_(123)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 216 NTNC(123)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 217 NKH_(124)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 218 NKHC(124)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 219 NKHD(124)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 220 NBU_(125)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 221 NBUC(125)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 222 NBUD(125)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 223 NSY_(126)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 224 NSYC(126)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 225 NSYW(126)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 226 NMO_(127)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 227 NMOC(127)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 228 NMOD(127)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          ), // 229 NHS_(128)
    BRKINFO( brkclsIdeographic, brkclsNil,        fBrkNone,  brkclsNil          ), // 230 WHT_(129)
    BRKINFO( brkclsCombining,   brkclsNil,        fBrkNone,  brkclsNil          ), // 231 LS__(130)
    BRKINFO( brkclsAlpha,       brkclsNil,        fBrkNone,  brkclsNil          )  // 232 XNW_(131)
};

// Break pair information for normal or strict Kinsoku
const BYTE s_rgbrkpairsKinsoku[CLineServices::brkclsLim][CLineServices::brkclsLim] =
{
    //1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21   = brclsAfter
    //  brkclsBefore:
    1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  2,  1, //   1 brkclsOpen    
    0,  1,  1,  1,  0,  0,  2,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //   2 brkclsClose   
    0,  1,  2,  1,  2,  0,  2,  4,  0,  0,  4,  2,  1,  2,  1,  4,  0,  0,  0,  2,  1, //   3 brkclsNoStartIdeo
    0,  1,  2,  1,  0,  0,  0,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //   4 brkclsExclamInt
    0,  1,  2,  1,  2,  0,  0,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //   5 brkclsInseparable
    2,  1,  2,  1,  0,  2,  0,  2,  2,  0,  3,  2,  1,  2,  1,  2,  2,  0,  0,  2,  1, //   6 brkclsPrefix  
    0,  1,  2,  1,  0,  0,  0,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //   7 brkclsPostfix 
    0,  1,  2,  1,  2,  0,  2,  4,  0,  0,  4,  2,  1,  2,  1,  4,  0,  0,  0,  2,  1, //   8 brkclsIdeoW   
    0,  1,  2,  1,  2,  0,  2,  0,  3,  0,  3,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //   9 brkclsNumeral 
    0,  1,  2,  1,  0,  0,  0,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //  10 brkclsSpaceN  
    0,  1,  2,  1,  2,  2,  2,  4,  3,  0,  3,  2,  1,  2,  1,  4,  0,  0,  0,  2,  1, //  11 brkclsAlpha   
    2,  1,  2,  1,  2,  2,  2,  2,  2,  2,  2,  2,  1,  2,  1,  2,  2,  2,  2,  2,  1, //  12 brkclsGlueA   
    0,  1,  2,  1,  2,  0,  2,  0,  2,  0,  3,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //  13 brkclsSlash   
    1,  1,  2,  1,  2,  2,  3,  2,  2,  2,  2,  2,  1,  2,  1,  2,  2,  2,  2,  2,  1, //  14 brkclsQuotation
    0,  1,  2,  1,  0,  2,  2,  0,  2,  0,  3,  2,  1,  2,  2,  0,  0,  0,  0,  2,  1, //  15 brkclsNumSeparator
    0,  1,  2,  1,  2,  0,  2,  4,  0,  0,  4,  2,  1,  2,  1,  4,  0,  0,  0,  2,  1, //  16 brkclsHangul  
    0,  1,  2,  1,  2,  0,  2,  0,  0,  0,  0,  2,  1,  2,  1,  4,  2,  2,  2,  2,  1, //  17 brkclsThaiFirst
    0,  1,  2,  1,  2,  0,  2,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  2,  2,  2,  1, //  18 brkclsThaiLast
    0,  1,  2,  1,  2,  0,  0,  0,  0,  0,  0,  2,  2,  2,  1,  0,  2,  2,  2,  2,  1, //  19 brkclsThaiAlpha
    0,  1,  2,  1,  2,  2,  2,  4,  3,  0,  3,  2,  1,  2,  1,  4,  0,  0,  0,  1,  1, //  20 brkclsCombining
    0,  1,  2,  1,  0,  0,  0,  0,  0,  0,  0,  2,  1,  2,  1,  0,  0,  0,  0,  2,  1, //  21 brkclsAsiiSpace
};

BOOL CanQuickBrkclsLookup(WCHAR ch)
{
    return (ch < BRKCLS_QUICKLOOKUPCUTOFF);
}

CLineServices::BRKCLS QuickBrkclsFromCh(WCHAR ch)
{
    Assert(ch);  // This won't work for ch==0.
    Assert(CanQuickBrkclsLookup(ch));
    return (CLineServices::BRKCLS)CLineServices::s_rgBrkInfo[ch].brkclsLow;
}

CLineServices::BRKCLS BrkclsFromCh(WCHAR ch, DWORD brkopt)
{
    Assert(!CanQuickBrkclsLookup(ch)); // Should take another code path.

    CHAR_CLASS cc = CharClassFromCh(ch);
    Assert(cc < CHAR_CLASS_MAX);

    const CLineServices::PACKEDBRKINFO* p = CLineServices::s_rgBrkInfo + cc;

    return CLineServices::BRKCLS((p->brkopt&brkopt) ? p->brkclsAlt : p->brkcls);
}

// Standard Breaking Behaviors retaining normal line break for non-FE text
static const LSBRK s_rglsbrkNormal[] = 
{
    /* 0*/ 1,1,  // always allowed
    /* 1*/ 0,0,  // always prohibited
    /* 2*/ 0,1,  // only allowed across space
    /* 3*/ 0,1,  // only allowed across space (word wrap case)
    /* 4*/ 1,1,  // always allowed (no CJK/Hangul word wrap case)
};

// Breaking Behaviors allowing FE style breaking in the middle of words (any language)
static const LSBRK s_rglsbrkBreakAll[] = 
{
    /* 0*/ 1,1,  // always allowed
    /* 1*/ 0,0,  // always prohibited
    /* 2*/ 0,1,  // only allowed across space
    /* 3*/ 1,1,  // always allowed (no word wrap case)
    /* 4*/ 1,1,  // always allowed (no CJK/Hangul word wrap case)
};

// Breaking Behaviors allowing Hangul style breaking 
static const LSBRK s_rglsbrkKeepAll[] = 
{
    /* 0*/ 1,1,  // always allowed
    /* 1*/ 0,0,  // always prohibited
    /* 2*/ 0,1,  // only allowed across space
    /* 3*/ 0,1,  // only allowed across space (word wrap case)
    /* 4*/ 0,1,  // only allowed across space (CJK/Hangul word wrap case)
};

const struct lsbrk* alsbrkTables[4] =
{                                                           
    s_rglsbrkNormal,
    s_rglsbrkNormal,
    s_rglsbrkBreakAll,
    s_rglsbrkKeepAll
};

LSERR CLineServices::CheckSetBreaking()
{
    const struct lsbrk* lsbrkCurr = alsbrkTables[_pPFFirst->_fWordBreak];

    HRESULT hr;

    // Are we in need of calling LsSetBreaking?
    if(lsbrkCurr == _lsbrkCurr)
    {
        hr = S_OK;
    }
    else
    {
        hr = HRFromLSERR(LsSetBreaking(_plsc,
            sizeof(s_rglsbrkNormal)/sizeof(LSBRK),
            lsbrkCurr,
            brkclsLim,
            (const BYTE*)s_rgbrkpairsKinsoku));

        _lsbrkCurr = (struct lsbrk*)lsbrkCurr;
    }

    RRETURN(hr);
}

LSERR WINAPI CLineServices::GetBreakingClasses(
                                  PLSRUN plsrun,          // IN
                                  LSCP lscp,              // IN
                                  WCHAR wch,              // IN
                                  BRKCLS* pbrkclsFirst,   // OUT
                                  BRKCLS* pbrkclsSecond)  // OUT
{
    LSTRACE(GetBreakingClasses);
    LSERR lserr = lserrNone;

    if(CanQuickBrkclsLookup(wch))
    {
        // prefer ASCII (true block) for performance
        Assert(IsNotThaiTypeChar(wch));
        *pbrkclsFirst = *pbrkclsSecond = QuickBrkclsFromCh(wch);
    }
    else if(IsNotThaiTypeChar(wch))  
    {
        *pbrkclsFirst = *pbrkclsSecond = BrkclsFromCh(wch, plsrun->_brkopt);
    }
    else
    {    
        CComplexRun* pcr = plsrun->GetComplexRun();

        if(pcr != NULL)
        {
            LONG cp = CPFromLSCP(lscp);

            pcr->ThaiTypeBrkcls(_pMarkup, cp, (::BRKCLS*)pbrkclsFirst, (::BRKCLS*)pbrkclsSecond);
        }
        else
        {
            // BUGFIX 14717 (a-pauln)
            // A complex run has not been created so pass this through the normal
            // Kinsoku classes for clean failure.
            *pbrkclsFirst = *pbrkclsSecond = BrkclsFromCh(wch, plsrun->_brkopt);
        }
    }

    return lserr;
}

const BRKCOND CLineServices::s_rgbrkcondBeforeChar[brkclsLim] =
{
    brkcondPlease,  // brkclsOpen
    brkcondNever,   // brkclsClose
    brkcondNever,   // brkclsNoStartIdeo
    brkcondNever,   // brkclsExclaInterr
    brkcondCan,     // brkclsInseparable
    brkcondCan,     // brkclsPrefix
    brkcondCan,     // brkclsPostfix
    brkcondPlease,  // brkclsIdeographic
    brkcondCan,     // brkclsNumeral
    brkcondCan,     // brkclsSpaceN
    brkcondCan,     // brkclsAlpha
    brkcondCan,     // brkclsGlueA
    brkcondPlease,  // brkclsSlash
    brkcondCan,     // brkclsQuote
    brkcondCan,     // brkclsNumSeparator
    brkcondCan,     // brkclsHangul
    brkcondCan,     // brkclsThaiFirst
    brkcondNever,   // brkclsThaiLast
    brkcondNever,   // brkclsThaiMiddle
    brkcondCan,     // brkclsCombining
    brkcondCan,     // brkclsAsciiSpace
};

LSERR WINAPI CLineServices::CanBreakBeforeChar(
                                  BRKCLS brkcls,          // IN
                                  BRKCOND* pbrktxtBefore) // OUT
{
    LSTRACE(CanBreakBeforeChar);

    Assert(brkcls>=0 && brkcls<brkclsLim);

    *pbrktxtBefore = s_rgbrkcondBeforeChar[brkcls];

    return lserrNone;
}

const BRKCOND CLineServices::s_rgbrkcondAfterChar[brkclsLim] =
{
    brkcondPlease,  // brkclsOpen
    brkcondCan,     // brkclsClose
    brkcondCan,     // brkclsNoStartIdeo
    brkcondCan,     // brkclsExclaInterr
    brkcondCan,     // brkclsInseparable
    brkcondCan,     // brkclsPrefix
    brkcondCan,     // brkclsPostfix
    brkcondPlease,  // brkclsIdeographic
    brkcondCan,     // brkclsNumeral
    brkcondCan,     // brkclsSpaceN
    brkcondCan,     // brkclsAlpha
    brkcondNever,   // brkclsGlueA
    brkcondPlease,  // brkclsSlash
    brkcondCan,     // brkclsQuote
    brkcondCan,     // brkclsNumSeparator
    brkcondCan,     // brkclsHangul
    brkcondNever,   // brkclsThaiFirst
    brkcondCan,     // brkclsThaiLast
    brkcondNever,   // brkclsThaiAlpha
    brkcondCan,     // brkclsCombining
    brkcondCan,     // brkclsAsciiSpace
};

LSERR WINAPI CLineServices::CanBreakAfterChar(
                                 BRKCLS brkcls,          // IN
                                 BRKCOND* pbrktxtAfter)  // OUT
{
    LSTRACE(CanBreakAfterChar);

    Assert(brkcls>=0 && brkcls<brkclsLim);

    *pbrktxtAfter = s_rgbrkcondAfterChar[ brkcls ];

    return lserrNone;
}

LSERR WINAPI CLineServices::DrawUnderline(
                             PLSRUN plsrun,          // IN
                             UINT kUlBase,           // IN
                             const POINT* pptStart,  // IN
                             DWORD dupUl,            // IN
                             DWORD dvpUl,            // IN
                             LSTFLOW kTFlow,         // IN
                             UINT kDisp,             // IN
                             const RECT* prcClip)    // IN
{
    LSTRACE(DrawUnderline);

    GetRenderer()->DrawUnderline(plsrun, kUlBase, pptStart, dupUl, dvpUl, kTFlow, kDisp, prcClip);
    return lserrNone;
}

LSERR WINAPI CLineServices::DrawStrikethrough(
                                 PLSRUN plsrun,          // IN
                                 UINT kStBase,           // IN
                                 const POINT* pptStart,  // IN
                                 DWORD dupSt,            // IN
                                 DWORD dvpSt,            // IN
                                 LSTFLOW kTFlow,         // IN
                                 UINT kDisp,             // IN
                                 const RECT* prcClip)    // IN
{
    LSTRACE(DrawStrikethrough);
    LSNOTIMPL(DrawStrikethrough);
    return lserrNone;
}

LSERR WINAPI CLineServices::DrawBorder(
                          PLSRUN plsrun,                              // IN
                          const POINT* pptStart,                      // IN
                          PCHEIGHTS pheightsLineFull,                 // IN
                          PCHEIGHTS pheightsLineWithoutAddedSpace,    // IN
                          PCHEIGHTS pheightsSubline,                  // IN
                          PCHEIGHTS pheightRuns,                      // IN
                          long dupBorder,                             // IN
                          long dupRunsInclBorders,                    // IN
                          LSTFLOW kTFlow,                             // IN
                          UINT kDisp,                                 // IN
                          const RECT* prcClip)                        // IN
{
    LSTRACE(DrawBorder);
    LSNOTIMPL(DrawBorder);
    return lserrNone;
}

LSERR WINAPI CLineServices::DrawUnderlineAsText(
                                   PLSRUN plsrun,          // IN
                                   const POINT* pptStart,  // IN
                                   long dupLine,           // IN
                                   LSTFLOW kTFlow,         // IN
                                   UINT kDisp,             // IN
                                   const RECT* prcClip)    // IN
{
    LSTRACE(DrawUnderlineAsText);
    LSNOTIMPL(DrawUnderlineAsText);
    return lserrNone;
}

LSERR WINAPI CLineServices::ShadeRectangle(
                              PLSRUN plsrun,                              // IN
                              const POINT* pptStart,                      // IN
                              PCHEIGHTS pheightsLineWithAddSpace,         // IN
                              PCHEIGHTS pheightsLineWithoutAddedSpace,    // IN
                              PCHEIGHTS pheightsSubline,                  // IN
                              PCHEIGHTS pheightsRunsExclTrail,            // IN
                              PCHEIGHTS pheightsRunsInclTrail,            // IN
                              long dupRunsExclTrail,                      // IN
                              long dupRunsInclTrail,                      // IN
                              LSTFLOW kTFlow,                             // IN
                              UINT kDisp,                                 // IN
                              const RECT* prcClip)                        // IN
{
    LSTRACE(ShadeRectangle);
    LSNOTIMPL(ShadeRectangle);
    return lserrNone;
}

LSERR WINAPI CLineServices::DrawTextRun(
                           PLSRUN plsrun,          // IN
                           BOOL fStrikeout,        // IN
                           BOOL fUnderline,        // IN
                           const POINT* pptText,   // IN
                           LPCWSTR plwchRun,       // IN
                           const int* rgDupRun,    // IN
                           DWORD cwchRun,          // IN
                           LSTFLOW kTFlow,         // IN
                           UINT kDisp,             // IN
                           const POINT* pptRun,    // IN
                           PCHEIGHTS heightPres,   // IN
                           long dupRun,            // IN
                           long dupLimUnderline,   // IN
                           const RECT* pRectClip)  // IN
{
    LSTRACE(DrawTextRun);

    TCHAR* pch = (TCHAR*)plwchRun;
    if(plsrun->_fMakeItASpace)
    {
        Assert(cwchRun == 1);
        pch = _T(" ");
    }
    GetRenderer()->TextOut(
        plsrun,   fStrikeout, fUnderline, pptText,
        pch,      rgDupRun,   cwchRun,    kTFlow,
        kDisp,    pptRun,     heightPres, dupRun,
        dupLimUnderline,      pRectClip);

    return lserrNone;
}

LSERR WINAPI CLineServices::DrawSplatLine(
                             enum lsksplat,                              // IN
                             LSCP cpSplat,                               // IN
                             const POINT* pptSplatLine,                  // IN
                             PCHEIGHTS pheightsLineFull,                 // IN
                             PCHEIGHTS pheightsLineWithoutAddedSpace,    // IN
                             PCHEIGHTS pheightsSubline,                  // IN
                             long dup,                                   // IN
                             LSTFLOW kTFlow,                             // IN
                             UINT kDisp,                                 // IN
                             const RECT* prcClip)                        // IN
{
    LSTRACE(DrawSplatLine);
    // FUTURE (mikejoch) Need to adjust cpSplat if we ever implement this.
    LSNOTIMPL(DrawSplatLine);
    return lserrNone;
}

//+---------------------------------------------------------------------------
//
//  Member:     CLineServices::DrawGlyphs
//
//  Synopsis:   Draws the glyphs which are passed in
//
//  Arguments:  plsrun              pointer to the run
//              fStrikeout          is this run struck out?
//              fUnderline          is this run underlined?
//              pglyph              array of glyph indices
//              rgDu                array of widths after justification
//              rgDuBeforeJust      array of widths before justification
//              rgGoffset           array of glyph offsets
//              rgGprop             array of glyph properties
//              rgExpType           array of glyph expansion types
//              cglyph              number of glyph indices
//              kTFlow              text direction and orientation
//              kDisp               display mode - opaque, transparent
//              pptRun              starting point of the run
//              heights             presentation height for this run
//              dupRun              presentation width of this run
//              dupLimUnderline     underline limit
//              pRectClip           clipping rectangle
//
//  Returns:    LSERR               lserrNone if succesful
//                                  lserrInvalidRun if failure
//
//----------------------------------------------------------------------------
LSERR WINAPI CLineServices::DrawGlyphs(
                          PLSRUN plsrun,                  // IN
                          BOOL fStrikeout,                // IN
                          BOOL fUnderline,                // IN
                          PCGINDEX pglyph,                // IN
                          const int* rgDu,                // IN
                          const int* rgDuBeforeJust,      // IN
                          PGOFFSET rgGoffset,             // IN
                          PGPROP rgGprop,                 // IN
                          PCEXPTYPE rgExpType,            // IN
                          DWORD cglyph,                   // IN
                          LSTFLOW kTFlow,                 // IN
                          UINT kDisp,                     // IN
                          const POINT* pptRun,            // IN
                          PCHEIGHTS heightsPres,          // IN
                          long dupRun,                    // IN
                          long dupLimUnderline,           // IN
                          const RECT* pRectClip)          // IN
{
    LSTRACE(DrawGlyphs);

    GetRenderer()->GlyphOut(
        plsrun,    fStrikeout, fUnderline, pglyph,
        rgDu,      rgDuBeforeJust,         rgGoffset,
        rgGprop,   rgExpType,  cglyph,     kTFlow,
        kDisp,     pptRun,     heightsPres,
        dupRun,    dupLimUnderline,        pRectClip);
    return lserrNone;
}

LSERR WINAPI CLineServices::DrawEffects(
                           PLSRUN plsrun,              // IN
                           UINT EffectsFlags,          // IN
                           const POINT* ppt,           // IN
                           LPCWSTR lpwchRun,           // IN
                           const int* rgDupRun,        // IN
                           const int* rgDupLeftCut,    // IN
                           DWORD cwchRun,              // IN
                           LSTFLOW kTFlow,             // IN
                           UINT kDisp,                 // IN
                           PCHEIGHTS heightPres,       // IN
                           long dupRun,                // IN
                           long dupLimUnderline,       // IN
                           const RECT* pRectClip)      // IN
{
    LSTRACE(DrawEffects);
    LSNOTIMPL(DrawEffects);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
// Ruby support
//
//-----------------------------------------------------------------------------
#define RUBY_OFFSET     2

LSERR WINAPI CLineServices::FetchRubyPosition(
                                 LSCP lscp,                          // IN
                                 LSTFLOW lstflow,                    // IN
                                 DWORD cdwMainRuns,                  // IN
                                 const PLSRUN* pplsrunMain,          // IN
                                 PCHEIGHTS pcheightsRefMain,         // IN
                                 PCHEIGHTS pcheightsPresMain,        // IN
                                 DWORD cdwRubyRuns,                  // IN
                                 const PLSRUN* pplsrunRuby,          // IN
                                 PCHEIGHTS pcheightsRefRuby,         // IN
                                 PCHEIGHTS pcheightsPresRuby,        // IN
                                 PHEIGHTS pheightsRefRubyObj,        // OUT
                                 PHEIGHTS pheightsPresRubyObj,       // OUT
                                 long* pdvrOffsetMainBaseline,       // OUT
                                 long* pdvrOffsetRubyBaseline,       // OUT
                                 long* pdvpOffsetRubyBaseline,       // OUT
                                 enum rubycharjust* prubycharjust,   // OUT
                                 BOOL* pfSpecialLineStartEnd)        // OUT
{
    LSTRACE(FetchRubyPosition);
    LONG yAscent;
    RubyInfo rubyInfo;
    long yRubyOffset = (cdwRubyRuns>0) ? RUBY_OFFSET : 0;

    *pdvrOffsetMainBaseline = 0;  // Don't want to offset the main text from the ruby object's baseline

    if(cdwMainRuns)
    {
        HandleRubyAlignStyle((COneRun*)(*pplsrunMain), prubycharjust, pfSpecialLineStartEnd);
    }

    rubyInfo.cp = CPFromLSCP(lscp);

    // Ruby offset is the sum of the ascent of the main text
    // and the descent of the pronunciation text.
    yAscent = max(_yMaxHeightForRubyBase, pcheightsRefMain->dvAscent);

    *pdvpOffsetRubyBaseline = 
        *pdvrOffsetRubyBaseline = 
        yAscent + yRubyOffset + pcheightsRefRuby->dvDescent;
    rubyInfo.yHeightRubyBase = yAscent + pcheightsRefMain->dvDescent + yRubyOffset;
    rubyInfo.yDescentRubyBase = pcheightsRefMain->dvDescent;
    rubyInfo.yDescentRubyText = pcheightsRefRuby->dvDescent;

    pheightsRefRubyObj->dvAscent = yAscent + pcheightsRefRuby->dvAscent + pcheightsRefRuby->dvDescent + yRubyOffset;
    pheightsRefRubyObj->dvDescent = pcheightsRefMain->dvDescent;
    pheightsRefRubyObj->dvMultiLineHeight = 
        pheightsRefRubyObj->dvAscent + pheightsRefRubyObj->dvDescent;
    memcpy(pheightsPresRubyObj, pheightsRefRubyObj, sizeof(*pheightsRefRubyObj));


    // code in GetRubyInfoFromCp depends on the idea that this callback 
    // is called in order of increasing cps. This of course depends on Line Services.  
    // If this is not guaranteed to be true, then we can't just blindly append the 
    // entry here, we must insert it in sorted order or hold cp ranges in the RubyInfos
    Assert(_aryRubyInfo.Size()==0 || _aryRubyInfo[_aryRubyInfo.Size()-1].cp<=rubyInfo.cp);
    if(_aryRubyInfo.FindIndirect(&rubyInfo) == -1)
    {
        _aryRubyInfo.AppendIndirect(&rubyInfo);
    }

    return lserrNone;
}


// Ruby Align Style table
// =========================
// Holds the justification values to pass to line services for each ruby alignment type.
static const enum rubycharjust s_aRubyAlignStyleValues[] =
{
    rcjCenter,   // not set
    rcjCenter,   // auto
    rcjLeft,     // left
    rcjCenter,   // center
    rcjRight,    // right
    rcj010,      // distribute-letter
    rcj121,      // distribute-space
    rcjCenter    // line-edge
};

void WINAPI CLineServices::HandleRubyAlignStyle(
                                    COneRun* porMain,                   // IN
                                    enum rubycharjust* prubycharjust,   // OUT
                                    BOOL* pfSpecialLineStartEnd)        // OUT
{
    Assert(porMain);

    CTreeNode*  pNode = porMain->Branch();  // This will call_ptp->GetBranch()
    CElement*   pElement = pNode->SafeElement();
    VARIANT     varRubyAlign;
    styleRubyAlign styAlign;

    pElement->ComputeExtraFormat(DISPID_A_RUBYALIGN, TRUE, pNode, &varRubyAlign);
    styAlign = (((CVariant*)&varRubyAlign)->IsEmpty()) 
        ? styleRubyAlignNotSet  : (styleRubyAlign)V_I4(&varRubyAlign);

    Assert(styAlign>=styleRubyAlignNotSet && styAlign<=styleRubyAlignLineEdge);

    *prubycharjust = s_aRubyAlignStyleValues[styAlign];
    *pfSpecialLineStartEnd = (styAlign == styleRubyAlignLineEdge);

    if(styAlign==styleRubyAlignNotSet || styAlign==styleRubyAlignAuto) 
    {
        // default behavior should be centered alignment for latin characters,
        // distribute-space for ideographic characters
        const SCRIPT_ID sid = porMain->_ptp->Sid();
        if(sid>=sidFEFirst && sid<=sidFELast)
        {
            *prubycharjust = rcj121;
        }
    }
}

LSERR WINAPI CLineServices::FetchRubyWidthAdjust(
                                    LSCP cp,                // IN
                                    PLSRUN plsrunForChar,   // IN
                                    WCHAR wch,              // IN
                                    MWCLS mwclsForChar,     // IN
                                    PLSRUN plsrunForRuby,   // IN
                                    rubycharloc rcl,        // IN
                                    long durMaxOverhang,    // IN
                                    long* pdurAdjustChar,   // OUT
                                    long* pdurAdjustRuby)   // OUT
{
    LSTRACE(FetchRubyWidthAdjust);
    Assert(plsrunForRuby);

    COneRun*    porRuby = (COneRun*)plsrunForRuby;
    CTreeNode*  pNode = porRuby->Branch(); // This will call_ptp->GetBranch()
    CElement*   pElement = pNode->SafeElement();
    styleRubyOverhang sty;

    {
        VARIANT varValue;

        pElement->ComputeExtraFormat(
            DISPID_A_RUBYOVERHANG, 
            TRUE, 
            pNode, 
            &varValue);

        sty = (((CVariant*)&varValue)->IsEmpty())
            ? styleRubyOverhangNotSet : (styleRubyOverhang)V_I4(&varValue);
    }

    *pdurAdjustChar = 0;
    *pdurAdjustRuby = (sty==styleRubyOverhangNone) ? 0 : -durMaxOverhang;
    return lserrNone;
}

LSERR WINAPI CLineServices::RubyEnum(
                        PLSRUN plsrun,              // IN
                        PCLSCHP plschp,             // IN
                        LSCP cp,                    // IN
                        LSDCP dcp,                  // IN
                        LSTFLOW lstflow,            // IN
                        BOOL fReverse,              // IN
                        BOOL fGeometryNeeded,       // IN
                        const POINT* pt,            // IN
                        PCHEIGHTS pcheights,        // IN
                        long dupRun,                // IN
                        const POINT* ptMain,        // IN
                        PCHEIGHTS pcheightsMain,    // IN
                        long dupMain,               // IN
                        const POINT* ptRuby,        // IN
                        PCHEIGHTS pcheightsRuby,    // IN
                        long dupRuby,               // IN
                        PLSSUBL plssublMain,        // IN
                        PLSSUBL plssublRuby)        // IN
{
    LSTRACE(RubyEnum);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
// Tatenakayoko (HIV) support
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetTatenakayokoLinePosition(
    LSCP cp,                            // IN
    LSTFLOW lstflow,                    // IN
    PLSRUN plsrun,                      // IN
    long dvr,                           // IN
    PHEIGHTS pheightsRef,               // OUT
    PHEIGHTS pheightsPres,              // OUT
    long* pdvrDescentReservedForClient) // OUT
{
    LSTRACE(GetTatenakayokoLinePosition);
    // FUTURE (mikejoch) Need to adjust cp if we ever implement this.
    LSNOTIMPL(GetTatenakayokoLinePosition);
    return lserrNone;
}

LSERR WINAPI CLineServices::TatenakayokoEnum(
                                PLSRUN plsrun,          // IN
                                PCLSCHP plschp,         // IN
                                LSCP cp,                // IN
                                LSDCP dcp,              // IN
                                LSTFLOW lstflow,        // IN
                                BOOL fReverse,          // IN
                                BOOL fGeometryNeeded,   // IN
                                const POINT* pt,        // IN
                                PCHEIGHTS pcheights,    // IN
                                long dupRun,            // IN
                                LSTFLOW lstflowT,       // IN
                                PLSSUBL plssubl)        // IN
{
    LSTRACE(TatenakayokoEnum);
    LSNOTIMPL(TatenakayokoEnum);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
// Warichu support
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::GetWarichuInfo(
                              LSCP cp,                            // IN
                              LSTFLOW lstflow,                    // IN
                              PCOBJDIM pcobjdimFirst,             // IN
                              PCOBJDIM pcobjdimSecond,            // IN
                              PHEIGHTS pheightsRef,               // OUT
                              PHEIGHTS pheightsPres,              // OUT
                              long* pdvrDescentReservedForClient) // OUT
{
    LSTRACE(GetWarichuInfo);
    // FUTURE (mikejoch) Need to adjust cp if we ever implement this.
    LSNOTIMPL(GetWarichuInfo);
    return lserrNone;
}

LSERR WINAPI CLineServices::FetchWarichuWidthAdjust(
                                       LSCP cp,                        // IN
                                       enum warichucharloc wcl,        // IN
                                       PLSRUN plsrunForChar,           // IN
                                       WCHAR wch,                      // IN
                                       MWCLS mwclsForChar,             // IN
                                       PLSRUN plsrunWarichuBracket,    // IN
                                       long* pdurAdjustChar,           // OUT
                                       long* pdurAdjustBracket)        // OUT
{
    LSTRACE(FetchWarichuWidthAdjust);
    // FUTURE (mikejoch) Need to adjust cp if we ever implement this.
    LSNOTIMPL(FetchWarichuWidthAdjust);
    return lserrNone;
}

LSERR WINAPI CLineServices::WarichuEnum(
                           PLSRUN plsrun,                  // IN: plsrun for the entire Warichu Object
                           PCLSCHP plschp,                 // IN: lschp for lead character of Warichu Object
                           LSCP cp,                        // IN: cp of first character of Warichu Object
                           LSDCP dcp,                      // IN: number of characters in Warichu Object
                           LSTFLOW lstflow,                // IN: text flow at Warichu Object
                           BOOL fReverse,                  // IN: whether text should be reversed for visual order
                           BOOL fGeometryNeeded,           // IN: whether Geometry should be returned
                           const POINT* pt,                // IN: starting position, iff fGeometryNeeded
                           PCHEIGHTS pcheights,            // IN: height of Warichu object, iff fGeometryNeeded
                           long dupRun,                    // IN: length of Warichu Object, iff fGeometryNeeded
                           const POINT* ptLeadBracket,     // IN: point for second line iff fGeometryNeeded and plssublSecond not NULL
                           PCHEIGHTS pcheightsLeadBracket, // IN: height for ruby line iff fGeometryNeeded 
                           long dupLeadBracket,            // IN: length of Ruby line iff fGeometryNeeded and plssublSecond not NULL
                           const POINT* ptTrailBracket,    // IN: point for second line iff fGeometryNeeded and plssublSecond not NULL
                           PCHEIGHTS pcheightsTrailBracket,// IN: height for ruby line iff fGeometryNeeded 
                           long dupTrailBracket,           // IN: length of Ruby line iff fGeometryNeeded and plssublSecond not NULL
                           const POINT* ptFirst,           // IN: starting point for main line iff fGeometryNeeded
                           PCHEIGHTS pcheightsFirst,       // IN: height of main line iff fGeometryNeeded
                           long dupFirst,                  // IN: length of main line iff fGeometryNeeded 
                           const POINT* ptSecond,          // IN: point for second line iff fGeometryNeeded and plssublSecond not NULL
                           PCHEIGHTS pcheightsSecond,      // IN: height for ruby line iff fGeometryNeeded and plssublSecond not NULL
                           long dupSecond,                 // IN: length of Ruby line iff fGeometryNeeded and plssublSecond not NULL
                           PLSSUBL plssublLeadBracket,     // IN: subline for lead bracket
                           PLSSUBL plssublTrailBracket,    // IN: subline for trail bracket
                           PLSSUBL plssublFirst,           // IN: first subline in Warichu object
                           PLSSUBL plssublSecond)          // IN: second subline in Warichu object
{
    LSTRACE(WarichuEnum);
    LSNOTIMPL(WarichuEnum);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
// HIH support
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::HihEnum(
                       PLSRUN plsrun,          // IN
                       PCLSCHP plschp,         // IN
                       LSCP cp,                // IN
                       LSDCP dcp,              // IN
                       LSTFLOW lstflow,        // IN
                       BOOL fReverse,          // IN
                       BOOL fGeometryNeeded,   // IN
                       const POINT* pt,        // IN
                       PCHEIGHTS pcheights,    // IN
                       long dupRun,            // IN
                       PLSSUBL plssubl)        // IN
{
    LSTRACE(HihEnum);
    LSNOTIMPL(HihEnum);
    return lserrNone;
}

//-----------------------------------------------------------------------------
//
// Reverse Object support
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::ReverseEnum(
                           PLSRUN plsrun,          // IN
                           PCLSCHP plschp,         // IN
                           LSCP cp,                // IN
                           LSDCP dcp,              // IN
                           LSTFLOW lstflow,        // IN
                           BOOL fReverse,          // IN
                           BOOL fGeometryNeeded,   // IN
                           const POINT* ppt,       // IN
                           PCHEIGHTS pcheights,    // IN
                           long dupRun,            // IN
                           LSTFLOW lstflowSubline, // IN
                           PLSSUBL plssubl)        // IN
{
    LSTRACE(ReverseEnum);

    return LsEnumSubline(plssubl, fReverse, fGeometryNeeded, ppt);
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::SetRenderer
//
// Synopsis:    Sets up the line services object to indicate that we will be
//              using for rendering.
//
//+----------------------------------------------------------------------------
void CLineServices::SetRenderer(CLSRenderer* pRenderer, BOOL fWrapLongLines, const CCharFormat* pCFLi)
{
    _pMeasurer = pRenderer;
    _pCFLi     = pCFLi;
    _lsMode    = LSMODE_RENDERER;
    _fWrapLongLines = fWrapLongLines;
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::SetMeasurer
//
// Synopsis:    Sets up the line services object to indicate that we will be
//              using for measuring / hittesting.
//
//+----------------------------------------------------------------------------
void CLineServices::SetMeasurer(CLSMeasurer* pMeasurer, LSMODE lsMode, BOOL fWrapLongLines)
{
    _pMeasurer = pMeasurer;
    Assert(lsMode!=LSMODE_NONE && lsMode!=LSMODE_RENDERER);
    _lsMode = lsMode;
    _fWrapLongLines = fWrapLongLines;
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::GetRenderer
//
//+----------------------------------------------------------------------------
CLSRenderer* CLineServices::GetRenderer()
{
    return ((_pMeasurer&&_lsMode==LSMODE_RENDERER) ? DYNCAST(CLSRenderer, _pMeasurer) : NULL);
}

//+----------------------------------------------------------------------------
//
// Member:      CLineServices::GetMeasurer
//
//+----------------------------------------------------------------------------
CLSMeasurer* CLineServices::GetMeasurer()
{
    return ((_pMeasurer&&_lsMode==LSMODE_MEASURER) ? _pMeasurer : NULL);
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::PAPFromPF
//
//  Synopsis:   Construct a PAP from PF
//
//-----------------------------------------------------------------------------
static const BYTE s_ablskjustMap[] =
{
    lskjFullInterWord,          // NotSet (default)
    lskjFullInterWord,          // InterWord
    lskjFullInterLetterAligned, // Newspaper
    lskjFullInterLetterAligned, // Distribute
    lskjFullInterLetterAligned, // DistributeAllLines
    lskjFullScaled,             // InterIdeograph
    lskjFullInterWord           // Auto (?)
};

void CLineServices::PAPFromPF(PLSPAP pap, const CParaFormat* pPF, BOOL fInnerPF, CComplexRun* pcr)
{
    _fExpectCRLF = pPF->HasPre(fInnerPF) || _fIsTextLess;
    const BOOL fJustified = pPF->GetBlockAlign(fInnerPF)==htmlBlockAlignJustify
        && !pPF->HasInclEOLWhite(fInnerPF) && !_fMinMaxPass;

    // line services format flags (lsffi.h)
    pap->grpf = fFmiApplyBreakingRules  // Use our breaking tables
        | fFmiSpacesInfluenceHeight;    // Whitespace contributes to extent

    if(_fWrapLongLines)
    {
        pap->grpf |= fFmiWrapAllSpaces;
    }
    else
    {
        pap->grpf |= fFmiForceBreakAsNext;   // No emergency breaks
    }

    pap->uaLeft = 0;
    pap->uaRightBreak = 0;
    pap->uaRightJustify = 0;
    pap->duaIndent = 0;
    pap->duaHyphenationZone = 0;

    // Justification type
    pap->lsbrj = lsbrjBreakJustify;

    // Justification
    if(fJustified)
    {
        _li._fJustified = JUSTIFY_FULL;

        // A. If we are a complex script set lskj to do glyphing
        if(pcr!=NULL || _pMarkup->Doc()->GetCodePage()==CP_THAI)
        {
            pap->lskj = lskjFullGlyphs;


            if(!pPF->_cuvTextKashida.IsNull())
            {
                Assert(_xWidthMaxAvail > 0);
                pap->lsbrj = lsbrjBreakThenExpand;

                // set the amount of the right break
                long xKashidaPercent = pPF->_cuvTextKashida.GetPercentValue(CUnitValue::DIRECTION_CX, _xWidthMaxAvail);
                _xWrappingWidth = _xWidthMaxAvail - xKashidaPercent;

                // we need to set this amount as twips
                Assert(_pci);
                pap->uaRightBreak = _pci->TwipsFromDeviceCX(xKashidaPercent);

                Assert(_xWrappingWidth >= 0);
            }
        }
        else
        {
            pap->lskj = LSKJUST(s_ablskjustMap[ pPF->_uTextJustify ]);
        }
        _fExpansionOrCompression = pPF->_uTextJustify > styleTextJustifyInterWord;
    }
    else
    {
        _fExpansionOrCompression = FALSE;
        pap->lskj = lskjNone;
    }

    // Alignment
    pap->lskal = lskalLeft;

    // Autonumbering
    pap->duaAutoDecimalTab = 0;

    // kind of paragraph ending
    pap->lskeop = _fExpectCRLF ? lskeopEndPara1 : lskeopEndParaAlt;

    // Main text flow direction
    Assert(pPF->HasRTL(fInnerPF) == (BOOL) _li._fRTL);
    if(!_li._fRTL)
    {
        pap->lstflow = lstflowES;
    }
    else
    {
        if(_pBidiLine == NULL)
        {
            _pBidiLine = new CBidiLine(_treeInfo, _cpStart, _li._fRTL, _pli);
        }
        pap->lstflow = lstflowWS;
    }
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::CHPFromCF
//
//  Synopsis:   Construct a CHP from CF
//
//-----------------------------------------------------------------------------
void CLineServices::CHPFromCF(COneRun* por, const CCharFormat* pCF)
{
    PLSCHP pchp = &por->_lsCharProps;

    // The structure has already been zero'd out in fetch run, which sets almost
    // everything we care about to the correct value (0).
    if (pCF->_fTextAutospace)
    {
        _lineFlags.AddLineFlag(por->Cp(), FLAG_HAS_NOBLAST);

        if(pCF->_fTextAutospace & TEXTAUTOSPACE_ALPHA)
        {
            pchp->fModWidthOnRun = TRUE;
            por->_csco |= cscoAutospacingAlpha;
        }
        if(pCF->_fTextAutospace & TEXTAUTOSPACE_NUMERIC)
        {
            pchp->fModWidthOnRun = TRUE;
            por->_csco |= cscoAutospacingDigit;
        }
        if(pCF->_fTextAutospace & TEXTAUTOSPACE_SPACE)
        {
            pchp->fModWidthSpace = TRUE;
            por->_csco |= cscoAutospacingAlpha;
        }
        if(pCF->_fTextAutospace & TEXTAUTOSPACE_PARENTHESIS)
        {
            pchp->fModWidthOnRun = TRUE;
            por->_csco |= cscoAutospacingParen;
        }
    }

    if(_fExpansionOrCompression)
    {
        pchp->fCompressOnRun = TRUE;
        pchp->fCompressSpace = TRUE; 
        pchp->fCompressTable = TRUE;
        pchp->fExpandOnRun = 0 == (pCF->_bPitchAndFamily&FF_SCRIPT);
        pchp->fExpandSpace = TRUE;
        pchp->fExpandTable = TRUE;
    }

    pchp->idObj = LSOBJID_TEXT;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::SetPOLS
//
//  Synopsis:   We call this function when we assign a CLineServices
//              to a CLSMeasurer.
//
//-----------------------------------------------------------------------------
void CLineServices::SetPOLS(CFlowLayout* pFlowLayout, CCalcInfo* pci)
{
    CTreePos *ptpStart, *ptpLast;
    CElement* pElementOwner = pFlowLayout->ElementOwner();

    _pFlowLayout = pFlowLayout;
    _fIsEditable = _pFlowLayout->IsEditable();
    _fIsTextLess = pElementOwner->HasFlag(TAGDESC_TEXTLESS);
    _fIsTD = pElementOwner->Tag() == ETAG_TD;
    _pMarkup = _pFlowLayout->GetContentMarkup();
    _fHasSites = FALSE;
    _pci = pci;
    _plsline = NULL;
    _chPassword = _pFlowLayout->GetPasswordCh();

    // We have special wrapping rules inside TDs with width specified.
    // Make note so the ILS breaking routines can break correctly.
    _xTDWidth = MAX_MEASURE_WIDTH;
    if(_fIsTD)
    {
        const LONG iUserWidth = 0;/*DYNCAST(CTableCellLayout, pFlowLayout)->GetSpecifiedPixelWidth(pci); wlw note*/

        if(iUserWidth)
        {
            _xTDWidth = iUserWidth;
        }
    }

    ClearLinePropertyFlags();

    _pFlowLayout->GetContentTreeExtent(&ptpStart, &ptpLast);
    _treeInfo._cpLayoutFirst = ptpStart->GetCp() + 1;
    _treeInfo._cpLayoutLast  = ptpLast->GetCp();
    _treeInfo._ptpLayoutLast = ptpLast;
    _treeInfo._tpFrontier.BindToCp(0);

    InitChunkInfo(_treeInfo._cpLayoutFirst);

    _pPFFirst = NULL;
    DEBUG_ONLY(_cpStart = -1);
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::ClearPOLS
//
//  Synopsis:   We call this function when we have finished using the measurer.
//
//-----------------------------------------------------------------------------
void CLineServices::ClearPOLS()
{
    // This assert will fire if we are leaking lsline's.  This happens
    // if somebody calls LSDoCreateLine without ever calling DiscardLine.
    Assert(_plsline == NULL);
    _pMarginInfo = NULL;
    if(_plcFirstChunk)
    {
        DeleteChunks();
    }
    DiscardOneRuns();

    if(_pccsCache)
    {
        _pccsCache->Release();
        _pccsCache = NULL;
        _pCFCache = NULL;
    }
    if(_pccsAltCache)
    {
        _pccsAltCache->Release();
        _pccsAltCache = NULL;
        _pCFAltCache = NULL;
    }
}

static CCharFormat s_cfBullet;

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetCFSymbol
//
//  Synopsis:   Get the a CF for the symbol passed in and put it in the COneRun.
//
//-----------------------------------------------------------------------------
void CLineServices::GetCFSymbol(COneRun* por, TCHAR chSymbol, const CCharFormat* pcfIn)
{
    static BOOL s_fBullet = FALSE;

    CCharFormat* pcfOut = por->GetOtherCF();

    Assert(pcfIn && pcfOut);
    if(pcfIn==NULL || pcfOut==NULL)
    {
        goto Cleanup;
    }

    if(!s_fBullet)
    {
        // N.B. (johnv) For some reason, Win95 does not render the Windings font properly
        //  for certain characters at less than 7 points.  Do not go below that size!
        s_cfBullet.SetHeightInTwips(TWIPS_FROM_POINTS(7));
        s_cfBullet._bCharSet = SYMBOL_CHARSET;
        s_cfBullet._fNarrow = FALSE;
        s_cfBullet._bPitchAndFamily = (BYTE)FF_DONTCARE;
        s_cfBullet._latmFaceName= fc().GetAtomWingdings();
        s_cfBullet._bCrcFont = s_cfBullet.ComputeFontCrc();

        s_fBullet = TRUE;
    }

    // Use bullet char format
    *pcfOut = s_cfBullet;

    pcfOut->_ccvTextColor = pcfIn->_ccvTextColor;

    // Important - CVM_SYMBOL is a special mode where out WC chars are actually
    // zero-extended MB chars.  This allows us to have a codepage-independent
    // call to ExTextOutA. (cthrash)
    por->SetConvertMode(CVM_SYMBOL);

Cleanup:
    return;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::GetCFNumber
//
//  Synopsis:   Get the a CF for the number passed in and put it in the COneRun.
//
//-----------------------------------------------------------------------------
void CLineServices::GetCFNumber(COneRun* por, const CCharFormat* pcfIn)
{
    CCharFormat* pcfOut = por->GetOtherCF();

    *pcfOut = *pcfIn;
    pcfOut->_fSubscript = pcfOut->_fSuperscript = FALSE;
    pcfOut->_bCrcFont   = pcfOut->ComputeFontCrc();
}

LONG CLineServices::GetDirLevel(LSCP lscp)
{
    LONG nLevel;
    COneRun* pHead;

    nLevel = _li._fRTL;
    for(pHead=_listCurrent._pHead; pHead; pHead=pHead->_pNext)
    {
        if(lscp >= pHead->_lscpBase)
        {
            if(pHead->IsSyntheticRun())
            {
                SYNTHTYPE synthtype = pHead->_synthType;
                // Since SYNTHTYPE_REVERSE preceeds SYNTHTYPE_ENDREVERSE and only
                // differs in the last bit, we can compute nLevel with bit magic. We
                // have to be sure this condition really exists of course, so we
                // Assert() it above.
                if(IN_RANGE(SYNTHTYPE_DIRECTION_FIRST, synthtype, SYNTHTYPE_DIRECTION_LAST))
                {
                    nLevel -= (((synthtype&1)<<1) - 1);
                }
            }
        }
        else
        {
            break;
        }
    }

    return nLevel;
}

#define MIN_FOR_LS 1

HRESULT CLineServices::Setup(
                     LONG xWidthMaxAvail,
                     LONG cp,
                     CTreePos* ptp,
                     const CMarginInfo* pMarginInfo,
                     const CLine* pli,
                     BOOL fMinMaxPass)
{
    const CParaFormat* pPF;
    BOOL fWrapLongLines = _fWrapLongLines;
    HRESULT hr = S_OK;

    Assert(_pMeasurer);

    if(!_treeInfo._fInited || cp!=long(_pMeasurer->GetCp()))
    {
        DiscardOneRuns();
        hr = _treeInfo.InitializeTreeInfo(_pFlowLayout, _fIsEditable, cp, ptp);
        if(hr != S_OK)
        {
            goto Cleanup;
        }
    }

    _lineFlags.InitLineFlags();
    _cpStart     = _cpAccountedTill = cp;
    _pMarginInfo = pMarginInfo;
    _cWhiteAtBOL = 0;
    _cAlignedSitesAtBOL = 0;
    _cInlinedSites = 0;
    _cAbsoluteSites = 0;
    _cAlignedSites = 0;
    _lCharGridSize = 0;
    _lLineGridSize = 0;
    _pNodeForAfterSpace = NULL;
    if(_lsMode == LSMODE_MEASURER)
    {
        _pli = NULL;
    } 
    else
    {
        Assert(_lsMode==LSMODE_HITTEST || _lsMode==LSMODE_RENDERER);
        _pli = pli;
        _li._fLookaheadForGlyphing = (_pli ? _pli->_fLookaheadForGlyphing : FALSE);
    }

    ClearLinePropertyFlags(); // zero out all flags
    _fWrapLongLines = fWrapLongLines; // preserve this flag when we 0 _dwProps

    // We're getting max, so start really small.
    _lMaxLineHeight = LONG_MIN;
    _fFoundLineHeight = FALSE;

    _xWrappingWidth = -1;

    pPF = _treeInfo._pPF;
    _fInnerPFFirst = _treeInfo._fInnerPF;

    // BUGBUG (a-pauln) some elements are getting assigned a PF from the
    //        wrong branch in InitializeTreeInfo() above. This hack is a
    //        temporary correction of the problem's manifestation until
    //        we determine how to correct it.
    if(!_treeInfo._fHasNestedElement || !ptp)
    {
        _li._fRTL = pPF->HasRTL(_fInnerPFFirst);
    }
    else
    {
        pPF = ptp->Branch()->GetParaFormat();
        _li._fRTL = pPF->HasRTL(_fInnerPFFirst);
    }

    if(!_pFlowLayout->GetDisplay()->GetWordWrap()
        || (pPF&&pPF->HasPre(_fInnerPFFirst)&&!_fWrapLongLines))
    {
        _xWrappingWidth = xWidthMaxAvail;
        xWidthMaxAvail = MAX_MEASURE_WIDTH;
    }
    else if(xWidthMaxAvail <= MIN_FOR_LS)
    {
        //BUGBUG(SujalP): Remove hack when LS gets their in-efficient calc bug fixed.
        xWidthMaxAvail = 0;
    }

    _xWidthMaxAvail = xWidthMaxAvail;
    if(_xWrappingWidth == -1)
    {
        _xWrappingWidth = _xWidthMaxAvail;
    }

    _fMinMaxPass = fMinMaxPass;
    _pPFFirst = pPF;
    DeleteChunks();
    InitChunkInfo(cp - (_pMeasurer->_fMeasureFromTheStart?0:_pMeasurer->_cchPreChars));

Cleanup:
    RRETURN(hr);
}

//-----------------------------------------------------------------------------
//
//  Function:   CLineServices::GetMinDurBreaks (member)
//
//  Synopsis:   Determine the minimum width of the line.  Also compute any
//              adjustments to the maximum width of the line.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::GetMinDurBreaks(LONG* pdvMin, LONG* pdvMaxDelta)
{
    LSERR lserr;

    Assert(_plsline);
    Assert(_fMinMaxPass);
    Assert(pdvMin);
    Assert(pdvMaxDelta);

    // First we call LsGetMinDurBreaks.  This call does the right thing only
    // for text runs, not ILS objects.
    if(!_fScanForCR)
    {
        LONG dvDummy;

        lserr = LsGetMinDurBreaks(GetContext(), _plsline, &dvDummy, pdvMin);
        if(lserr)
        {
            goto Cleanup;
        }
    }

    // Now we need to go and compute the true max width.  The current max
    // width is incorrect by the difference in the min and max widths of
    // dobjs for which these values are not the same (e.g. tables).  We've
    // cached the difference in the dobj's, so we need to enumerate these
    // and add them up.  The enumeration callback adjusts the value in
    // CLineServices::dvMaxDelta;
    _dvMaxDelta = 0;

    lserr = LsEnumLine(
        _plsline,
        FALSE,                      // fReverseOrder
        FALSE,                      // fGeometryNeeded
        &_afxGlobalData._Zero.pt);  // pptOrg

    *pdvMaxDelta = _dvMaxDelta;

    if(_fScanForCR)
    {
        *pdvMin = _li._xWidth + _dvMaxDelta;
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   CLineServices destructor
//
//  Synopsis:   Free the COneRun cache
//
//-----------------------------------------------------------------------------
CLineServices::~CLineServices()
{
    if(_pccsCache)
    {
        _pccsCache->Release();
    }

    if(_pccsAltCache)
    {
        _pccsAltCache->Release();
    }
}

//-----------------------------------------------------------------------------
//
//  Function:   CLineServices::DiscardLine (member)
//
//  Synopsis:   The lifetime of a CLineServices object is that of it's
//              containing CDoc.  This function is used to clear state
//              after each use of the object, as opposed to the destructor.
//              This function needs to be called after each measured/rendered
//              line (~CLine.)
//
//-----------------------------------------------------------------------------
void CLineServices::DiscardLine()
{
    if(_plsline)
    {
        LsDestroyLine(_plsc, _plsline);
        _lineFlags.DeleteAll();
        _lineCounts.DeleteAll();
        _plsline = NULL;
    }

    // For now just do this simple thing here. Eventually we will do more
    // complex things like holding onto tree state.
    _treeInfo._fInited = FALSE;

    if(_pBidiLine != NULL)
    {
        delete _pBidiLine;
        _pBidiLine = NULL;
    }

    _aryRubyInfo.DeleteAll();
}


//+----------------------------------------------------------------------------
//
// Member:      InitChunkInfo
//
// Synopsis:    Initializest the chunk info store.
//
//-----------------------------------------------------------------------------
void CLineServices::InitChunkInfo(LONG cp)
{
    _cpLastChunk = cp;
    _xPosLastChunk = 0;
    _plcFirstChunk = _plcLastChunk = NULL;
    _pElementLastRelative = NULL;
    _fLastChunkSingleSite = FALSE;
}

//+----------------------------------------------------------------------------
//
// Member:      DeleteChunks
//
// Synopsis:    Delete chunk related information in the line
//
//-----------------------------------------------------------------------------
void CLineServices::DeleteChunks()
{
    while(_plcFirstChunk)
    {
        CLSLineChunk* plc = _plcFirstChunk;

        _plcFirstChunk = _plcFirstChunk->_plcNext;
        delete plc;
    }
    _plcLastChunk = NULL;
}

//-----------------------------------------------------------------------------
//
//  Function:   QueryLinePointPcp (member)
//
//  Synopsis:   Wrapper for LsQueryLinePointPcp
//
//  Returns:    S_OK    - Success
//              S_FALSE - depth was zero
//              E_FAIL  - error
//
//-----------------------------------------------------------------------------
HRESULT CLineServices::QueryLinePointPcp(
                                 LONG u,                     // IN
                                 LONG v,                     // IN
                                 LSTFLOW* pktFlow,           // OUT
                                 PLSTEXTCELL plstextcell)    // OUT
{
    POINTUV uvPoint;
    CStackDataAry<LSQSUBINFO, 4> aryLsqsubinfo;
    HRESULT hr;
    DWORD nDepthIn = 4;

    uvPoint.u = u;
    uvPoint.v = v;

    Assert(_plsline);

#define NDEPTH_MAX 32

    for(;;)
    {
        DWORD nDepth;
        LSERR lserr = LsQueryLinePointPcp(
            _plsline,
            &uvPoint,
            nDepthIn,
            aryLsqsubinfo,
            &nDepth,
            plstextcell);

        if(lserr == lserrNone)
        {
            hr = S_OK;
            // get the flow direction for proper x+/- manipulation
            if(nDepth > 0)
            {
                if(aryLsqsubinfo[nDepth-1].idobj!=LSOBJID_TEXT
                    && aryLsqsubinfo[nDepth-1].idobj!=LSOBJID_EMBEDDED)
                {
                    LSQSUBINFO &qsubinfo = aryLsqsubinfo[nDepth-1];
                    plstextcell->dupCell = qsubinfo.dupObj;
                    plstextcell->pointUvStartCell = qsubinfo.pointUvStartSubline;
                    plstextcell->cCharsInCell = 0;
                    plstextcell->cpStartCell = qsubinfo.cpFirstSubline;
                    plstextcell->cpEndCell = qsubinfo.cpFirstSubline + qsubinfo.dcpSubline;
                }
                else
                {
                    plstextcell->cCharsInCell = plstextcell->cpEndCell - plstextcell->cpStartCell + 1;
                }
                *pktFlow = aryLsqsubinfo[nDepth-1].lstflowSubline;
            }
            else if(nDepth == 0)
            {
                // HACK ALERT(MikeJoch):
                // See hack alert by SujalP below. We can run into this case
                // when the line is terminated by a [section break] type
                // character. We should take it upon ourselves to fill in
                // plstextcell and pktFlow when this happens.
                LONG duIgnore;

                plstextcell->cpStartCell = _lscpLim - 1;
                plstextcell->cpEndCell = _lscpLim;
                plstextcell->dupCell = 0;
                plstextcell->cCharsInCell = 1;

                hr = GetLineWidth(&plstextcell->pointUvStartCell.u, &duIgnore);

                // If we don't have a level, assume that the flow is in the line direction.
                if(pktFlow)
                {
                    *pktFlow = _li._fRTL ? fUDirection : 0;
                }

            }
            else
            {
                hr = E_FAIL;
            }
            break;
        }
        else if(lserr == lserrInsufficientQueryDepth) 
        {
            if(nDepthIn > NDEPTH_MAX)
            {
                hr = E_FAIL;
                break;
            }

            nDepthIn *= 2;
            Assert(nDepthIn <= NDEPTH_MAX); // That would be rediculous

            hr = aryLsqsubinfo.Grow(nDepthIn);
            if(hr)
            {
                break;
            }
            // Loop back.
        }
        else
        {
            hr = E_FAIL;
            break;
        }
    }

    RRETURN1(hr, S_FALSE);
}

//-----------------------------------------------------------------------------
//
//  Function:   QueryLineCpPpoint (member)
//
//  Synopsis:   Wrapper for LsQueryLineCpPpoint
//
//  Returns:    S_OK    - Success
//              S_FALSE - depth was zero
//              E_FAIL  - error
//
//-----------------------------------------------------------------------------
HRESULT CLineServices::QueryLineCpPpoint(
                                 LSCP lscp,                              // IN
                                 BOOL fFromGetLineWidth,                 // IN
                                 CDataAry<LSQSUBINFO>* paryLsqsubinfo,   // IN/OUT
                                 PLSTEXTCELL plstextcell,                // OUT
                                 BOOL* pfRTLFlow)                        // OUT
{
    CStackDataAry<LSQSUBINFO, 4> aryLsqsubinfo;
    HRESULT hr;
    DWORD nDepthIn;
    LSTFLOW ktFlow;
    DWORD nDepth;

    Assert(_plsline);

    if(paryLsqsubinfo == NULL)
    {
        aryLsqsubinfo.Grow(4); // Guaranteed to succeed since we're working from the stack.
        paryLsqsubinfo = &aryLsqsubinfo;
    }
    nDepthIn = paryLsqsubinfo->Size();

#define NDEPTH_MAX 32

    for(;;)
    {
        LSERR lserr = LsQueryLineCpPpoint(
            _plsline,
            lscp,
            nDepthIn,
            *paryLsqsubinfo,
            &nDepth,
            plstextcell);

        if(lserr == lserrNone)
        {
            // HACK ALERT(SujalP):
            // Consider the case where the line contains just 3 characters:
            // A[space][blockbreak] at cp 0, 1 and 2 respectively. If we query
            // LS at cp 2, if would expect lsTextCell to point to the zero
            // width cell containing [blockbreak], meaning that
            // lsTextCell.pointUvStartCell.u would be the width of the line
            // (including whitespaces). However, upon being queried at cp=2
            // LS returns a nDepth of ***0*** because it thinks this is some
            // splat crap. This problem breaks a lot of our callers, hence
            // we attemp to hack around this problem.
            // NOTE: In case LS fixes their problem, be sure that hittest.htm
            // renders all its text properly for it exhibits a problem because
            // of this problem.
            // ORIGINAL CODE: hr = nDepth ? S_OK : S_FALSE;
            if(nDepth == 0)
            {
                LONG duIgnore;

                if(!fFromGetLineWidth && lscp>=_lscpLim-1)
                {
                    plstextcell->cpStartCell = _lscpLim - 1;
                    plstextcell->cpEndCell = _lscpLim;
                    plstextcell->dupCell = 0;

                    hr = GetLineWidth(&plstextcell->pointUvStartCell.u, &duIgnore);

                }
                else
                {
                    hr = S_FALSE;
                }

                // If we don't have a level, assume that the flow is in the line direction.
                if(pfRTLFlow)
                {
                    *pfRTLFlow = _li._fRTL;
                }
            }
            else
            {
                hr = S_OK;
                LSQSUBINFO& qsubinfo = (*paryLsqsubinfo)[nDepth-1];

                if(qsubinfo.idobj!=LSOBJID_TEXT
                    && qsubinfo.idobj!=LSOBJID_EMBEDDED)
                {
                    plstextcell->dupCell = qsubinfo.dupObj;
                    plstextcell->pointUvStartCell = qsubinfo.pointUvStartObj;
                    plstextcell->cCharsInCell = 0;
                    plstextcell->cpStartCell = qsubinfo.cpFirstSubline;
                    plstextcell->cpEndCell = qsubinfo.cpFirstSubline + qsubinfo.dcpSubline;
                }

                // if we are going in the opposite direction of the line we
                // will need to compensate for proper xPos
                ktFlow = qsubinfo.lstflowSubline;
                if(pfRTLFlow)
                {
                    *pfRTLFlow = (ktFlow == lstflowWS);
                }
            }
            break;
        }
        else if(lserr == lserrInsufficientQueryDepth) 
        {
            if(nDepthIn > NDEPTH_MAX)
            {
                hr = E_FAIL;
                break;
            }

            nDepthIn *= 2;
            Assert(nDepthIn <= NDEPTH_MAX); // That would be ridiculous

            hr = paryLsqsubinfo->Grow(nDepthIn);
            if(hr)
            {
                break;
            }
            // Loop back.
        }
        else
        {
            hr = E_FAIL;
            break;
        }
    }

    if(hr == S_OK)
    {
        Assert((LONG)nDepth <= paryLsqsubinfo->Size());
        paryLsqsubinfo->SetSize(nDepth);
    }

    RRETURN1(hr, S_FALSE);
}

HRESULT CLineServices::GetLineWidth(LONG* pdurWithTrailing, LONG* pdurWithoutTrailing)
{
    LSERR lserr;
    LONG  duIgnore;

    lserr = LsQueryLineDup(
        _plsline, &duIgnore, &duIgnore, &duIgnore,
        pdurWithoutTrailing, pdurWithTrailing);
    if(lserr != lserrNone)
    {
        goto Cleanup;
    }

Cleanup:
    return HRFromLSERR(lserr);
}

BOOL CLineServices::IsOwnLineSite(COneRun* por)
{
    return (por->_fCharsForNestedLayout?
        por->_ptp->Branch()->Element()->IsOwnLineElement(_pFlowLayout):FALSE);
}

//+----------------------------------------------------------------------------
//
//  Function:   CLineServices::UnUnifyHan
//
//  Synopsis:   Pick of one the Far East script id's for an sidHan range.
//
//              Most ideographic characters fall in sidHan (~34k out of all
//              of ucs-2), but not all fonts support every Han character.
//              This function attempts to pick one of the four FE sids.
//
//  Returns:    sidKana     for Japanese
//              sidHangul   for Korean
//              sidBopomofo for Traditional Chinese
//              sidHan      for Simplified Chinese
//
//-----------------------------------------------------------------------------
SCRIPT_ID CLineServices::UnUnifyHan(
                          HDC hdc,
                          UINT uiFamilyCodePage,
                          LCID lcid,
                          COneRun* por)
{
    SCRIPT_ID sid = sidTridentLim;

    if(!lcid)
    {
        switch(uiFamilyCodePage)
        {
        case CP_CHN_GB:     sid = sidHan;       break;
        case CP_KOR_5601:   sid = sidHangul;    break;
        case CP_TWN:        sid = sidBopomofo;  break;
        case CP_JPN_SJ:     sid = sidKana;      break;
        }
    }
    else
    {
        LANGID lid = LANGIDFROMLCID(lcid);
        WORD plid = PRIMARYLANGID(lid);

        if(plid == LANG_CHINESE)
        {
            if(SUBLANGID(lid) == SUBLANG_CHINESE_TRADITIONAL)
            {
                sid = sidBopomofo;
            }
            else
            {
                sid = sidHan;
            }
        }
        else if(plid == LANG_KOREAN)
        {
            sid = sidHangul;
        }
        else if(plid == LANG_JAPANESE)
        {
            sid = sidKana;
        }
    }

    if(sid == sidTridentLim)
    {
        long cp = por->Cp();
        long cch = GetScriptTextBlock(por->_ptp, &cp);
        DWORD dwPriorityCodePages = fc().GetSupportedCodePageInfo(hdc) & (FS_JOHAB|FS_CHINESETRAD|FS_WANSUNG|FS_CHINESESIMP|FS_JISJAPAN);
        DWORD dwCodePages = GetTextCodePages(dwPriorityCodePages, cp, cch);

        dwCodePages &= dwPriorityCodePages;

        if(dwCodePages & FS_JISJAPAN)
        {
            sid = sidKana;
        }
        else if(!dwCodePages || dwCodePages&FS_CHINESETRAD)
        {
            sid = sidBopomofo;
        }
        else if(dwCodePages & FS_WANSUNG)
        {
            sid = sidHangul;
        }
        else
        {
            sid = sidHan;
        }
    }

    return sid;
}

//+----------------------------------------------------------------------------
//
//  Function:   CLineServices::DisambiguateScriptId
//
//  Synopsis:   Some characters have sporadic coverage in Far East fonts.  To
//              make matters worse, these characters often have good glyphs in
//              Latin fonts.  We call IMLangCodePages to check the coverage
//              of codepages for the codepoints of interest, and then pick
//              an appropriate script id.
//
//              The first priority is to not switch the font.  If it appears
//              that the prevailing font has coverage, we will preferentially
//              pick this font.  Otherwise, we go through a somewhat arbitrary
//              priority list to pick a better guess.
//
//  Returns:    TRUE if the default font should be used.
//
//              Script ID, iff the default font isn't used.
//
//-----------------------------------------------------------------------------
const DWORD s_adwCodePagesMap[] =
{
    FS_JOHAB,           // 0
    FS_CHINESETRAD,     // 1
    FS_WANSUNG,         // 2
    FS_CHINESESIMP,     // 3
    FS_JISJAPAN,        // 4
    FS_VIETNAMESE,      // 5
    FS_BALTIC,          // 6
    FS_ARABIC,          // 7
    FS_HEBREW,          // 8
    FS_TURKISH,         // 9
    FS_GREEK,           // 10
    FS_CYRILLIC,        // 11
    FS_LATIN2           // 12
};

const SCRIPT_ID s_asidCodePagesMap[] =
{
    sidHangul,          // 0  FS_JOHAB
    sidBopomofo,        // 1  FS_CHINESETRAD
    sidHangul,          // 2  FS_WANSUNG
    sidHan,             // 3  FS_CHINESESIMP
    sidKana,            // 4  FS_JISJAPAN
    sidLatin,           // 5  FS_VIETNAMESE
    sidLatin,           // 6  FS_BALTIC
    sidArabic,          // 7  FS_ARABIC
    sidHebrew,          // 8  FS_HEBREW
    sidLatin,           // 9  FS_TURKISH
    sidGreek,           // 10 FS_GREEK
    sidCyrillic,        // 11 FS_CYRILLIC
    sidLatin            // 12 FS_LATIN2
};

const BYTE s_abCodePagesMap[] =
{
    JOHAB_CHARSET,      // 0  FS_JOHAB
    CHINESEBIG5_CHARSET,// 1  FS_CHINESETRAD
    HANGEUL_CHARSET,    // 2  FS_WANSUNG
    GB2312_CHARSET,     // 3  FS_CHINESESIMP
    SHIFTJIS_CHARSET,   // 4  FS_JISJAPAN
    VIETNAMESE_CHARSET, // 5  FS_VIETNAMESE
    BALTIC_CHARSET,     // 6  FS_BALTIC
    ARABIC_CHARSET,     // 7  FS_ARABIC
    HEBREW_CHARSET,     // 8  FS_HEBREW
    TURKISH_CHARSET,    // 9  FS_TURKISH
    GREEK_CHARSET,      // 10 FS_GREEK
    RUSSIAN_CHARSET,    // 11 FS_CYRILLIC
    EASTEUROPE_CHARSET  // 12 FS_LATIN2
};

BOOL CLineServices::DisambiguateScriptId(
                                    HDC hdc,                    // IN
                                    CBaseCcs* pBaseCcs,         // IN
                                    COneRun* por,               // IN
                                    SCRIPT_ID* psidAlt,         // OUT
                                    BYTE* pbCharSetAlt )        // OUT                                    
{
    long cp = por->Cp();
    long cch = GetScriptTextBlock(por->_ptp, &cp);
    DWORD dwCodePages = GetTextCodePages(0, cp ,cch);
    BOOL fCheckAltFont;
    SCRIPT_ID sid = sidDefault;
    BYTE bCharSetCurrent = pBaseCcs->_bCharSet;
    BOOL fFECharSet;

    if(dwCodePages)
    {
        // (1) Check if the current font will cover the glyphs
        if(bCharSetCurrent == ANSI_CHARSET)
        {
            fCheckAltFont = 0==(dwCodePages&FS_LATIN1) || 0==(pBaseCcs->_sids&ScriptBit(sidLatin));
            fFECharSet = FALSE;
        }
        else
        {
            int i = ARRAYSIZE(s_adwCodePagesMap);

            fCheckAltFont = TRUE;
            fFECharSet = IsFECharset(bCharSetCurrent);

            while(i--)
            {
                if(bCharSetCurrent == s_abCodePagesMap[i])
                {
                    fCheckAltFont = 0 == (dwCodePages&s_adwCodePagesMap[i]);
                    break;
                }
            }
        }

        if(fCheckAltFont)
        {
            SCRIPT_IDS sidsFace = fc().EnsureScriptIDsForFont(hdc, pBaseCcs, FALSE);

            if(fFECharSet)
            {
                sidsFace &= (ScriptBit(sidKana)|ScriptBit(sidHangul)|ScriptBit(sidBopomofo)|ScriptBit(sidHan));
            }            

            // (2) Check if the current fontface will cover the glyphs
            // (3) Pick an appropriate sid for the glyphs
            if(dwCodePages & FS_LATIN1)
            {
                if(sidsFace & ScriptBit(sidLatin))
                {
                    sid = sidAmbiguous;
                    fCheckAltFont = 0 == (pBaseCcs->_sids&ScriptBit(sidLatin));
                }
                else
                {
                    sid = sidLatin;
                }

                *pbCharSetAlt = ANSI_CHARSET;
            }
            else
            {
                int i = ARRAYSIZE(s_adwCodePagesMap);
                int iAlt = -1;

                dwCodePages &= fc().GetSupportedCodePageInfo(hdc);

                if(fFECharSet)
                {
                    // NOTE (cthrash) If are native font doesn't support it, don't switch to
                    // another FE charset because the font creation will simply fail and give
                    // us an undesirable font.
                    dwCodePages &= ~(FS_JOHAB|FS_CHINESETRAD|FS_WANSUNG|FS_CHINESESIMP|FS_JISJAPAN);
                }

                if(dwCodePages)
                {
                    while(i--)
                    {
                        if(dwCodePages & s_adwCodePagesMap[i])
                        {
                            const SCRIPT_IDS sids = ScriptBit(s_asidCodePagesMap[i]);

                            if(sidsFace & sids)
                            {
                                fCheckAltFont = 0 == (pBaseCcs->_sids&sids);
                                break;          // (2)
                            }
                            else if(-1 == iAlt)
                            {
                                iAlt = i;       // Candidate for (3), provided nothing meet requirement for (2)
                            }
                        }
                    }

                    if(-1 != i)
                    {
                        sid = sidAmbiguous;
                        *pbCharSetAlt = s_abCodePagesMap[i];
                    }
                    else if(-1 != iAlt)
                    {
                        sid = s_asidCodePagesMap[iAlt];
                        *pbCharSetAlt = s_abCodePagesMap[iAlt];
                    }
                    else
                    {
                        fCheckAltFont = FALSE;
                    }
                }
            }
        }
    }
    else
    {
        fCheckAltFont = FALSE;
    }

    *psidAlt = sid;

    return fCheckAltFont;
}

//+----------------------------------------------------------------------------
//
//  Function:   CLineServices::GetScriptTextBlock
//
//  Synopsis:   Get the cp and cch of the current text block, which may be
//              artificially chopped up due to the presence of pointer nodes.
//
//  Returns:    cch and cp.
//
//-----------------------------------------------------------------------------
#define FONTLINK_SCAN_MAX   32L

long CLineServices::GetScriptTextBlock(/*[in]*/CTreePos* ptp, /*[out]*/long* pcp)
{
    const long ich = *pcp - ptp->GetCp(); // chars to the left of the cp in the ptp
    const SCRIPT_ID sid = ptp->Sid();
    long cchBack;
    long cchForward;
    CTreePos* ptpStart = ptp;

    // Find the 'true' start, backing up over pointers.
    // Look for the beginning of the sid block.
    for(cchBack=ich; cchBack<=FONTLINK_SCAN_MAX;)
    {
        CTreePos* ptpPrev = ptpStart->PreviousTreePos();

        if(!ptpPrev || ptpPrev->IsNode()
            || (ptpPrev->IsText() && ptpPrev->Sid()!=sid))
        {
            break;
        }

        ptpStart = ptpPrev;

        cchBack += ptpPrev->GetCch();
    }

    // Find the 'true' end.
    for(cchForward=ptp->GetCch()-ich; cchForward<=FONTLINK_SCAN_MAX;)
    {
        CTreePos* ptpNext = ptp->NextTreePos();

        if(!ptpNext || ptpNext->IsNode()
            || (ptpNext->IsText() && ptpNext->Sid()!=sid))
        {
            break;
        }

        ptp = ptpNext;

        cchForward += ptp->GetCch();
    }

    // Determine the cp at which to start, and the cch to scan.
    if(cchBack <= FONTLINK_SCAN_MAX)
    {
        *pcp = ptpStart->GetCp();
    }
    else
    {
        *pcp = ptpStart->GetCp() + cchBack - FONTLINK_SCAN_MAX;
    }

    return min(cchBack, FONTLINK_SCAN_MAX)+min(cchForward, FONTLINK_SCAN_MAX);
}

//+----------------------------------------------------------------------------
//
//  Function:   CLineServices::GetTextCodePages
//
//  Synopsis:   Determine the codepage coverage of cch chars of text starting
//              at cp using IMLangCodePages.
//
//  Returns:    Codepage coverage as a bitfield (FS_*, see wingdi.h).
//
//-----------------------------------------------------------------------------
#define GETSTRCODEPAGES_CHUNK   10L

DWORD CLineServices::GetTextCodePages(DWORD dwPriorityCodePages, long cp, long cch)
{
    HRESULT hr;
    extern IMultiLanguage* g_pMultiLanguage;
    IMLangCodePages* pMLangCodePages = NULL;
    DWORD dwCodePages;
    LONG cchCodePages;
    long cchValid;
    CTxtPtr tp(_pMarkup, cp);
    const TCHAR* pch = tp.GetPch(cchValid);
    long cchInTp = cchValid;
    DWORD dwMask;

    // PERF (cthrash) In the following HTML, we will have a sidAmbiguous
    // run with a single space char in in (e.g. amazon.com): <B>&middot;</B> X
    // The middot is sidAmbiguous, the space is sidMerge, and thus they merge.
    // The X of course is sidAsciiLatin.  When called for a single space char,
    // we don't want to bother calling MLANG.  Just give the caller a stock
    // answer so as to avoid the unnecessary busy work.
    if(cch==1 && *pch==WCH_SPACE)
    {
        return dwPriorityCodePages?dwPriorityCodePages:DWORD(-1);
    }

    dwCodePages = dwPriorityCodePages;
    dwMask = DWORD(-1);

    hr = EnsureMultiLanguage();
    if(hr)
    {
        goto Cleanup;
    }

    hr = g_pMultiLanguage->QueryInterface(IID_IMLangCodePages, (void**)&pMLangCodePages);
    if(hr)
    {
        goto Cleanup;
    }

    // Strip out the leading WCH_NBSP chars
    while(cch)
    {
        long cchToCheck = min(GETSTRCODEPAGES_CHUNK, min(cch, cchValid));
        const TCHAR* pchStart = pch;
        const TCHAR* pchNext;
        DWORD dwRet = 0;

        while(*pch == WCH_NBSP)
        {
            pch++;
            cchToCheck--;
            if(!cchToCheck)
            {
                break;
            }
        }

        pchNext = pch + cchToCheck;

        if(cchToCheck)
        {
            hr = pMLangCodePages->GetStrCodePages(pch, cchToCheck, dwCodePages, &dwRet, &cchCodePages);
            if(hr)
            {
                goto Cleanup;
            }

            dwCodePages = dwMask & dwRet;

            // No match, not much we can do.
            if(!dwCodePages)
            {
                break;
            }

            // This is an error condition of sorts.  The whole run couldn't be
            // covered by a single codepage.  Just bail and give the user what
            // we have so far.
            if(cchToCheck > cchCodePages)
            {
                if(pch[cchCodePages] == WCH_NBSP)
                {
                    pchNext = pch + cchCodePages + 1; // resume after NBSP char
                }
                else
                {
                    break;
                }
            }
        }

        // Decrement our counts, advance our pointers.
        cchValid -= (long)(pchNext - pchStart);
        cch -= (long)(pchNext - pchStart);

        if(cch)
        {
            DWORD dw = dwCodePages;
            int i, j;

            // First check if we have only one codepage represented
            // If we do, we're done.  Note the largest bit we care
            // about is FS_JOHAB (0x00200000) so we look at 22 bits.
            for(i=22,j=0; i--&&j<2; dw>>=1)
            {
                j += dw & 1;
            }

            if(j == 1)
            {
                break;
            }

            if(cchValid)
            {
                pch = pchNext;
            }
            else
            {
                tp.AdvanceCp(cchInTp);
                pch = tp.GetPch(cchInTp);
                cchValid = cchInTp;
            }

            dwMask = dwCodePages;
        }
    }

Cleanup:
    ReleaseInterface(pMLangCodePages);
    return dwCodePages;
}

//-----------------------------------------------------------------------------
//
//  Function:   CLineServices::ShouldSwitchFontsForPUA
//
//  Synopsis:   Decide if we should switch fonts for Unicode Private Use Area
//              characters.  This is a heuristic approach.
//
//  Returns:    BOOL    - true if we should switch fonts
//              *psid   - the SCRIPT_ID to which we should switch
//
//-----------------------------------------------------------------------------
BOOL CLineServices::ShouldSwitchFontsForPUA(
                                       HDC hdc,
                                       UINT uiFamilyCodePage,
                                       CBaseCcs* pBaseCcs,
                                       const CCharFormat* pcf,
                                       SCRIPT_ID* psid)
{
    BOOL fShouldSwitch;

    if(pcf->_fExplicitFace)
    {
        // Author specified a face name -- don't switch fonts on them.

        fShouldSwitch = FALSE;
    }
    else
    {
        SCRIPT_IDS sidsFace = fc().EnsureScriptIDsForFont(hdc, pBaseCcs, FALSE);
        SCRIPT_ID sid = sidDefault;

        if(pcf->_lcid)
        {
            // Author specified a LANG attribute -- see if the current font
            // covers this script
            const LANGID langid = LANGIDFROMLCID(pcf->_lcid);

            if(langid == LANG_CHINESE)
            {
                sid = SUBLANGID(langid)==SUBLANG_CHINESE_TRADITIONAL ? sidBopomofo : sidHan;
            }
            else
            {
                sid = DefaultScriptIDFromLang(langid);
            }
        }
        else
        {
            // Check the document codepage, then the system codepage
            switch (uiFamilyCodePage)
            {
            case CP_CHN_GB:     sid = sidHan;       break;
            case CP_KOR_5601:   sid = sidHangul;    break;
            case CP_TWN:        sid = sidBopomofo;  break;
            case CP_JPN_SJ:     sid = sidKana;      break;
            default:
                {
                    switch(_afxGlobalData._cpDefault)
                    {
                    case CP_CHN_GB:     sid = sidHan;       break;
                    case CP_KOR_5601:   sid = sidHangul;    break;
                    case CP_TWN:        sid = sidBopomofo;  break;
                    case CP_JPN_SJ:     sid = sidKana;      break;
                    default:            sid = sidDefault;   break;                        
                    }
                }
            }
        }

        if(sid != sidDefault)
        {
            fShouldSwitch = (sidsFace&ScriptBit(sid)) == sidsNotSet;
            *psid = sid;
        }
        else
        {
            fShouldSwitch = FALSE;
        }
    }

    return fShouldSwitch;            
}

//-----------------------------------------------------------------------------
//
//  Function:   CLineServices::GetCcs
//
//  Synopsis:   Gets the suitable font (CCcs) for the given COneRun.
//
//  Returns:    CCcs
//
//-----------------------------------------------------------------------------
CCcs* CLineServices::GetCcs(COneRun* por, HDC hdc, CDocInfo* pdi)
{
    CCcs* pccs;
    const CCharFormat* pCF = por->GetCF();
    const BOOL fDontFontLink = !por->_ptp
        || !por->_ptp->IsText()
        || _chPassword
        || pCF->_bCharSet==SYMBOL_CHARSET
        || pCF->_fDownloadedFont
        || pdi->_pDoc->GetCodePage()==50000;
    BOOL fCheckAltFont;   // TRUE if _pccsCache does not have glyphs needed for sidText
    BOOL fPickNewAltFont; // TRUE if _pccsAltCache needs to be created anew
    SCRIPT_ID sidAlt = 0;
    BYTE bCharSetAlt = 0;
    SCRIPT_ID sidText;

    // NB (cthrash) Although generally it will, CCharFormat::_latmFaceName need
    // not necessarily match CBaseCcs::_latmLFFaceName.  This is particularly true
    // when a generic CSS font-family is used (Serif, Fantasy, etc.)  We won't
    // know the actual font properties until we call CCcs::MakeFont.  For this
    // reason, we always compute the _pccsCache first, even though it's
    // possible (not likely, but possible) that we'll never use the font.

    // If we have a different pCF then what _pccsCache is based on,
    if(pCF!=_pCFCache
        // *or* If we dont have a cached _pccsCache
        || !_pccsCache)
    {
        pccs = fc().GetCcs(hdc, pdi, pCF);
        if(pccs)
        {
            if(CVM_UNINITED != por->_bConvertMode)
            {
                pccs->GetBaseCcs()->_bConvertMode = por->_bConvertMode; // BUGBUG (cthrash) Is this right?
            }

            _pCFCache = (!por->_fMustDeletePcf ? pCF : NULL);

            if(_pccsCache)
            {
                _pccsCache->Release();
            }

            _pccsCache = pccs;
        }
        else
        {
            AssertSz(0, "CCcs failed to be created.");
            goto Cleanup;
        }
    }

    Assert(pCF == _pCFCache || por->_fMustDeletePcf);
    pccs = _pccsCache;

    if(fDontFontLink)
    {
        goto Cleanup;
    }

    // Check if the _pccsCache covers the sid of the text run
    Assert(por->_lsCharProps.fGlyphBased ||
        por->_sid==(!_pCFLi?DWORD(por->_ptp->Sid()):sidAsciiLatin));

    sidText = por->_sid;

    AssertSz(sidText!=sidMerge, "Script IDs should have been merged.");

    {
        CBaseCcs* pBaseCcs = _pccsCache->GetBaseCcs();

        if(sidText < sidLim)
        {
            if(sidText == sidDefault)
            {
                fCheckAltFont = FALSE; // Assume the author picked a font containing the glyph.  Don't fontlink.
            }
            else
            {
                fCheckAltFont = (pBaseCcs->_sids&ScriptBit(sidText)) == sidsNotSet;
            }
        }
        else
        {
            Assert(sidText==sidSurrogateA
                || sidText==sidSurrogateB
                || sidText==sidAmbiguous
                || sidText==sidEUDC
                || sidText==sidHalfWidthKana);

            if(sidText == sidAmbiguous)
            {
                fCheckAltFont = DisambiguateScriptId(hdc, pBaseCcs, por, &sidAlt, &bCharSetAlt);
            }
            else if(sidText == sidEUDC)
            {
                const UINT uiFamilyCodePage = pdi->_pDoc->GetFamilyCodePage();
                SCRIPT_ID sidForPUA;

                fCheckAltFont = ShouldSwitchFontsForPUA(hdc, uiFamilyCodePage, pBaseCcs, pCF, &sidForPUA);
                if(fCheckAltFont)
                {
                    sidText = sidAlt = sidAmbiguous; // to prevent call to UnUnifyHan
                    bCharSetAlt = DefaultCharSetFromScriptAndCodePage(sidForPUA, uiFamilyCodePage);
                }
            }
            else
            {
                fCheckAltFont = (pBaseCcs->_sids&ScriptBit(sidText)) == sidsNotSet;
            }
        }
    }

    if(!fCheckAltFont)
    {
        goto Cleanup;
    }

    // Check to see if the _pccsAltCache covers the sid of the text run

    {
        const UINT uiFamilyCodePage = pdi->_pDoc->GetFamilyCodePage();

        if(sidText == sidHan)
        {
            sidAlt = UnUnifyHan(hdc, uiFamilyCodePage, pCF->_lcid, por);
            bCharSetAlt = DefaultCharSetFromScriptAndCodePage(sidAlt, uiFamilyCodePage);
        }
        else if(sidText != sidAmbiguous)
        {
            SCRIPT_IDS sidsFace = fc().EnsureScriptIDsForFont(hdc, _pccsCache->GetBaseCcs(), FALSE);

            bCharSetAlt = DefaultCharSetFromScriptAndCodePage(sidText, uiFamilyCodePage);

            if((sidsFace&ScriptBit(sidText)) == sidsNotSet)
            {
                sidAlt = sidText;           // current face does not cover
            }
            else
            {
                sidAlt = sidAmbiguous;      // current face does cover
            }
        }

        fPickNewAltFont = !_pccsAltCache || _pCFAltCache!=pCF
            || _pccsAltCache->GetBaseCcs()->_bCharSet!=bCharSetAlt;
    }

    // Looks like we need to pick a new alternate font
    if(!fPickNewAltFont)
    {
        Assert(_pccsAltCache);
        pccs = _pccsAltCache;
    }
    else
    {
        CCharFormat cfAlt = *pCF;
        BOOL fNewlyFetchedFromRegistry;

        // sidAlt of sidAmbiguous at this point implies we have the right facename,
        // but the wrong GDI charset.  Otherwise, lookup in the registry/mlang to
        // determine an appropriate font for the given script id.
        if(sidAlt != sidAmbiguous)
        {
            fNewlyFetchedFromRegistry = SelectScriptAppropriateFont(sidAlt, bCharSetAlt, pdi->_pDoc, &cfAlt);
        }
        else
        {
            fNewlyFetchedFromRegistry = FALSE;
            cfAlt._bCharSet = bCharSetAlt;
            cfAlt._bCrcFont = cfAlt.ComputeFontCrc();
        }

        // NB (cthrash) sidAltText may differ from sidText in the case where
        // sidText is sidHan.  SelectScriptAppropriateFont makes a reasonable
        // effort to pick the optimal script amongst the unified Han scripts.

        pccs = fc().GetFontLinkCcs(hdc, pdi, pccs, &cfAlt);
        if(pccs)
        {
            // Remember the pCF from which the pccs was derived.
            _pCFAltCache = pCF;

            if(_pccsAltCache)
            {
                _pccsAltCache->Release();
            }

            _pccsAltCache = pccs;

            if(fNewlyFetchedFromRegistry)
            {
                // It's possible that the font signature for the given font would indicate
                // that it does not have the necessary glyph coverage for the sid in question.
                // Nevertheless, the user has specifically requested this face for the sid,
                // so we honor his/her choice and assume that the glyphs are covered.
                ForceScriptIdOnUserSpecifiedFont(pdi->_pDoc->_pOptionSettings, sidAlt);
            }

            pccs->GetBaseCcs()->_sids |= ScriptBit(sidAlt);
        }
    }

Cleanup:
    Assert(!pccs || pccs->GetHDC()==hdc);
    return pccs;
}

//-----------------------------------------------------------------------------
//
//  Function:   LineHasNoText
//
//  Synopsis:   Utility function which returns TRUE if there is no text (only nested
//              layout or antisynth'd goop) on the line till the specified cp.
//
//  Returns:    BOOL
//
//-----------------------------------------------------------------------------
BOOL CLineServices::LineHasNoText(LONG cp)
{
    LONG fRet = TRUE;
    for(COneRun* por=_listCurrent._pHead; por; por=por->_pNext)
    {
        if(por->Cp() >= cp)
        {
            break;
        }

        if(!por->IsNormalRun())
        {
            continue;
        }

        if(por->_lsCharProps.idObj == LSOBJID_TEXT)
        {
            fRet = FALSE;
            break;
        }
    }
    return fRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   CheckSetTables (member)
//
//  Synopsis:   Set appropriate tables for Line Services
//
//  Returns:    HRESULT
//
//-----------------------------------------------------------------------------
HRESULT CLineServices::CheckSetTables()
{
    LSERR lserr;

    lserr = CheckSetBreaking();

    if(lserr==lserrNone && _pPFFirst->_uTextJustify>styleTextJustifyInterWord)
    {
        lserr = CheckSetExpansion();

        if(lserr == lserrNone)
        {
            lserr = CheckSetCompression();
        }
    }

    RRETURN(HRFromLSERR(lserr));
}

//-----------------------------------------------------------------------------
//
//  Function:   HRFromLSERR (global)
//
//  Synopsis:   Utility function to converts a LineServices return value
//              (LSERR) into an appropriate HRESULT.
//
//  Returns:    HRESULT
//
//-----------------------------------------------------------------------------
HRESULT HRFromLSERR(LSERR lserr)
{
    HRESULT hr;

    switch(lserr)
    {
    case lserrNone:         hr = S_OK;          break;
    case lserrOutOfMemory:  hr = E_OUTOFMEMORY; break;
    default:                hr = E_FAIL;        break;
    }

    return hr;
}

//-----------------------------------------------------------------------------
//
//  Function:   LSERRFromHR (global)
//
//  Synopsis:   Utility function to converts a HRESULT into an appropriate
//              LineServices return value (LSERR).
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR LSERRFromHR(HRESULT hr)
{
    LSERR lserr;

    if(SUCCEEDED(hr))
    {
        lserr = lserrNone;
    }
    else
    {
        switch(hr)
        {
        default:
            AssertSz(FALSE, "Unmatched error code; returning lserrOutOfMemory");
        case E_OUTOFMEMORY:
            lserr = lserrOutOfMemory;
            break;
        }
    }

    return lserr;
}

// Since lnobj is worthless as far as we're concerned, we just point it back
// to the ilsobj.  lnobj's are instantiated once per object type per line.
typedef struct lnobj
{
    PILSOBJ pilsobj;
} LNOBJ;

LSERR WINAPI CLineServices::CreateILSObj(
                            PLSC plsc,          // IN
                            PCLSCBK plscbk,     // IN
                            DWORD idObj,        // IN
                            PILSOBJ* ppilsobj)  // OUT
{
    LSTRACE(CreateILSObj);

    switch(idObj)
    {
    case LSOBJID_EMBEDDED:
        *ppilsobj= (PILSOBJ)(new CEmbeddedILSObj(this, plscbk));
        break;
    case LSOBJID_NOBR:
        *ppilsobj= (PILSOBJ)(new CNobrILSObj(this, plscbk));
        break;
    default:
        AssertSz(0, "Unknown lsobj_id");
    }

    return *ppilsobj?lserrNone:lserrOutOfMemory;
}

//-----------------------------------------------------------------------------
//
//  Function:   TruncateChunk (static member, LS callback)
//
//  Synopsis:   From LS docs: Purpose    To obtain the exact position within
//              a chunk where the chunk crosses the right margin.  A chunk is
//              a group of contiguous LS objects which are of the same type.
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::TruncateChunk(/*[in]*/PCLOCCHNK plocchnk, /*[out]*/PPOSICHNK pposichnk)
{
    LSTRACE(TruncateChunk);
    LSERR lserr = lserrNone;
    long idobj;

    if(plocchnk->clschnk == 1)
    {
        idobj = 0;
    }
    else
    {
        long urColumnMax;
        long urTotal;
        OBJDIM objdim;

        urColumnMax = PDOBJ(plocchnk->plschnk[0].pdobj)->GetPLS()->_xWrappingWidth;
        urTotal = plocchnk->lsfgi.urPen;

        Assert(urTotal <= urColumnMax);

        for(idobj=0; idobj<(long)plocchnk->clschnk; idobj++)
        {
            PDOBJ pdobj = (PDOBJ) plocchnk->plschnk[idobj].pdobj;
            lserr = pdobj->QueryObjDim(&objdim);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }

            urTotal += objdim.dur;
            if(urTotal > urColumnMax)
            {
                break;
            }
        }

        Assert(urTotal > urColumnMax);
        Assert(idobj < (long)plocchnk->clschnk);
    }

    // Tell it which chunk didn't fit.
    pposichnk->ichnk = idobj;

    // Tell it which cp inside this chunk the break occurs at.
    // In our case, it's always an all-or-nothing proposition.  So include the whole dobj
    pposichnk->dcp = plocchnk->plschnk[idobj].dcp;

    Trace3("Truncate(cchnk=%d cpFirst=%d) -> ichnk=%d\n",
        plocchnk->clschnk,
        plocchnk->plschnk->cpFirst,
        idobj);

Cleanup:
    return lserr;
}

LSERR WINAPI CLineServices::ForceBreakChunk(
                               PCLOCCHNK plocchnk,     // IN
                               PCPOSICHNK pposichnk,   // IN
                               PBRKOUT pbrkout)        // OUT
{
    LSTRACE(ForceBreakChunk);

    Trace1("CLineServices::ForceBreakChunk(ichnk=%d)\n", pposichnk->ichnk);

    ZeroMemory(pbrkout, sizeof(BRKOUT));

    pbrkout->fSuccessful = fTrue;

    if(plocchnk->lsfgi.fFirstOnLine&&pposichnk->ichnk==0
        || pposichnk->ichnk==ichnkOutside)
    {
        PDOBJ pdobj  = (PDOBJ) plocchnk->plschnk[0].pdobj;

        pbrkout->posichnk.dcp = plocchnk->plschnk[0].dcp;
        pdobj->QueryObjDim(&pbrkout->objdim);
    }
    else
    {
        pbrkout->posichnk.ichnk = pposichnk->ichnk;
        pbrkout->posichnk.dcp = 0;
    }

    return lserrNone;
}

//-----------------------------------------------------------------------------
//
//  Function:   DiscardOneRuns
//
//  Synopsis:   Removes all the runs from the current list and gives them
//              back to the free list.
//
//  Returns:    Nothing
//
//-----------------------------------------------------------------------------
void CLineServices::DiscardOneRuns()
{
    COneRun* pFirst = _listCurrent._pHead;
    COneRun* pTail  = _listCurrent._pTail;

    if(pFirst)
    {
        _listCurrent.SpliceOut(pFirst, pTail);
        _listFree.SpliceIn(pFirst);
    }
}

//-----------------------------------------------------------------------------
//
//  Function:   AdvanceOneRun
//
//  Synopsis:   This is our primary function to get the next run at the
//              frontier.
//
//  Returns:    The one run.
//
//-----------------------------------------------------------------------------
COneRun* CLineServices::AdvanceOneRun(LONG lscp)
{
    COneRun* por;
    BOOL fRet = FALSE;

    // Get the memory for a one run
    por = _listFree.GetFreeOneRun(NULL);
    if(!por)
    {
        goto Cleanup;
    }

    // Setup the lscp...
    por->_lscpBase = lscp;

    if(!_treeInfo._fHasNestedElement)
    {
        // If we have run out of characters in the tree pos then we need
        // advance to the next tree pos.
        if(!_treeInfo._cchRemainingInTreePos)
        {
            if(!_treeInfo.AdvanceTreePos())
            {
                goto Cleanup;
            }
            por->_fCannotMergeRuns = TRUE;
        }

        // If we have run out of characters in the text then we need
        // to advance to the next text pos
        if(!_treeInfo._cchValid)
        {
            if(!_treeInfo.AdvanceTxtPtr())
            {
                goto Cleanup;
            }
            por->_fCannotMergeRuns = TRUE;
        }

        Assert(_treeInfo._lscpFrontier == lscp);
    }

    // If we have a nested run owner then the number of characters given to
    // the run are the number of chars in that nested run owner. Else the
    // number of chars is the minimum of the chars in the tree pos and that
    // in the text node.
    if(!_treeInfo._fHasNestedElement)
    {
        por->_lscch = min(_treeInfo._cchRemainingInTreePos, _treeInfo._cchValid);
        if(_lsMode == LSMODE_MEASURER)
        {
            por->_lscch = min(por->_lscch, MAX_CHARS_FETCHRUN_RETURNS);
        }
        else
        {
            // NB Additional 5 chars corresponds to a fudge factor.
            por->_lscch = min(por->_lscch, LONG(_pMeasurer->_li._cch+5));
        }
        AssertSz(por->_lscch>0, "Cannot measure 0 or -ve chars!");

        por->_pchBase = _treeInfo._pchFrontier;
    }
    else
    {
        // NOTE(SujalP): The number of characters _returned_ to LS will not be
        // _lscch. We will catch this case in FetchRun and feed only a single
        // char with the pch pointing to a valid location so that LS does not
        // choke on it.
        CElement* pElemNested = _treeInfo._ptpFrontier->Branch()->Element();

        por->_lscch = GetNestedElementCch(pElemNested);
        por->_fCannotMergeRuns = TRUE;
        por->_pchBase = NULL;
        por->_fCharsForNestedElement  = TRUE;
        por->_fCharsForNestedLayout   = _treeInfo._fHasNestedLayout;
        por->_fCharsForNestedRunOwner = _treeInfo._fHasNestedRunOwner;
    }

    // Update all the other information in the one run
    por->_chSynthsBefore = _treeInfo._chSynthsBefore;
    por->_lscchOriginal = por->_lscch;
    por->_pchBaseOriginal = por->_pchBase;
    por->_ptp = _treeInfo._ptpFrontier;
    por->SetSidFromTreePos(por->_ptp);

    por->_pCF = (CCharFormat*)_treeInfo._pCF;
    por->_fInnerCF = _treeInfo._fInnerCF;
    por->_pPF = _treeInfo._pPF;
    por->_fInnerPF = _treeInfo._fInnerPF;
    por->_pFF = _treeInfo._pFF;

    // At last let us go and move out frontier
    _treeInfo.AdvanceFrontier(por);

    fRet = TRUE;

Cleanup:
    if(!fRet && por)
    {
        delete por;
        por = NULL;
    }

    return por;
}

//-----------------------------------------------------------------------------
//
//  Function:   CanMergeTwoRuns
//
//  Synopsis:   Decided if the 2 runs can be merged into one
//
//  Returns:    BOOL
//
//-----------------------------------------------------------------------------
BOOL CLineServices::CanMergeTwoRuns(COneRun* por1, COneRun* por2)
{
    BOOL fRet;

    if(por1->_dwProps
        || por2->_dwProps
        || por1->_pCF!=por2->_pCF
        || por1->_bConvertMode!=por2->_bConvertMode
        || por1->_ccvBackColor.GetRawValue()!=por2->_ccvBackColor.GetRawValue()
        || por1->_pComplexRun
        || por2->_pComplexRun
        || (por1->_pchBase+por1->_lscch!=por2->_pchBase) // happens with passwords.
        )
    {
        fRet = FALSE;
    }
    else
    {
        fRet = TRUE;
    }
    return fRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   MergeIfPossibleIntoCurrentList
//
//-----------------------------------------------------------------------------
COneRun* CLineServices::MergeIfPossibleIntoCurrentList(COneRun* por)
{
    COneRun* pTail = _listCurrent._pTail;
    if(pTail!=NULL && CanMergeTwoRuns(pTail, por))
    {
        Assert(pTail->_lscpBase+pTail->_lscch == por->_lscpBase);
        Assert(pTail->_pchBase+pTail->_lscch == por->_pchBase);
        Assert(!pTail->_fNotProcessedYet); // Cannot merge into a run not yet processed

        pTail->_lscch += por->_lscch;
        pTail->_lscchOriginal += por->_lscchOriginal;

        // Since we merged our por into the previous one, let us put the
        // present one back on the free list.
        _listFree.SpliceIn(por);
        por = pTail;
    }
    else
    {
        _listCurrent.SpliceInAfterMe(pTail, por);
    }
    return por;
}

//-----------------------------------------------------------------------------
//
//  Function:   SplitRun
//
//  Synopsis:   Splits a single run into 2 runs. The original run remains
//              por and the number of chars it has is cchSplitTill, while
//              the new split off run is the one which is returned and the
//              number of characters it has is cchOriginal-cchSplit.
//
//  Returns:    The 2nd run (which we got from cutting up por)
//
//-----------------------------------------------------------------------------
COneRun* CLineServices::SplitRun(COneRun* por, LONG cchSplitTill)
{
    LONG cchDelta;

    // Create an exact copy of the run
    COneRun* porNew = _listFree.GetFreeOneRun(por);
    if(!porNew)
    {
        goto Cleanup;
    }
    por->_lscch = cchSplitTill;
    cchDelta = por->_lscchOriginal - por->_lscch;
    por->_lscchOriginal = por->_lscch;
    Assert(por->_lscch);

    // Then setup the second run so that it can be spliced in properly
    porNew->_pPrev = porNew->_pNext = NULL;
    porNew->_lscpBase += por->_lscch;
    porNew->_lscch = cchDelta;
    porNew->_lscchOriginal = porNew->_lscch;
    porNew->_pchBase = por->_pchBaseOriginal + cchSplitTill;
    porNew->_pchBaseOriginal = porNew->_pchBase;
    porNew->_fGlean = TRUE;
    porNew->_fNotProcessedYet = TRUE;
    Assert(porNew->_lscch);

Cleanup:
    return porNew;
}

//-----------------------------------------------------------------------------
//
//  Function:   AttachOneRunToCurrentList
//
//  Note: We always return the pointer to the run which is contains the
//  lscp for por. Consider the following cases:
//  1) No splitting:
//          If merged then return the ptr of the run we merged por into
//          If not merged then return por itself
//  2) Splitting:
//          Split into por and porNew
//          If por is merged then return ptr of the run we merged por into
//          If not morged then return por itself
//          Just attach/merge porNew
//
//  Returns:    The attached/merged-into run.
//
//-----------------------------------------------------------------------------
COneRun* CLineServices::AttachOneRunToCurrentList(COneRun* por)
{
    COneRun* porRet;

    Assert(por);
    Assert(por->_lscchOriginal >= por->_lscch);

    if(por->_lscchOriginal > por->_lscch)
    {
        Assert(por->IsNormalRun());
        COneRun* porNew = SplitRun(por, por->_lscch);
        if(!porNew)
        {
            porRet = NULL;
            goto Cleanup;
        }

        // Then splice in the current run and then the one we split out.
        porRet = MergeIfPossibleIntoCurrentList(por);

        // can replace this with a SpliceInAfterMe
        MergeIfPossibleIntoCurrentList(porNew);
    }
    else
    {
        porRet = MergeIfPossibleIntoCurrentList(por);
    }

Cleanup:
    return porRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   AppendSynth
//
//  Synopsis:   Appends a synthetic into the current one run store.
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLineServices::AppendSynth(COneRun* por, SYNTHTYPE synthtype, COneRun** pporOut)
{
    COneRun* pTail = _listCurrent._pTail;
    LONG     lscp  = por->_lscpBase;
    LSERR    lserr = lserrNone;
    BOOL     fAdd;
    LONG     lscpLast;

    // Atmost one node can be un-processed
    if(pTail && pTail->_fNotProcessedYet)
    {
        pTail = pTail->_pPrev;
    }

    if(pTail)
    {
        lscpLast = pTail->_lscpBase + (pTail->IsAntiSyntheticRun()?0:pTail->_lscch);
        Assert(lscp <= lscpLast);
        if(lscp == lscpLast)
        {
            fAdd = TRUE;
        }
        else
        {
            fAdd = FALSE;
            while(pTail)
            {
                Assert(pTail->_fNotProcessedYet == FALSE);
                if(pTail->_lscpBase == lscp)
                {
                    Assert(pTail->IsSyntheticRun());
                    *pporOut = pTail;
                    break;
                }
                pTail = pTail->_pNext;
            }

            AssertSz(*pporOut, "Cannot find the synthetic char which should have been there!");
        }
    }
    else
    {
        fAdd = TRUE;
    }

    if(fAdd)
    {
        COneRun* porNew;

        porNew = _listFree.GetFreeOneRun(por);
        if(!porNew)
        {
            lserr = lserrOutOfMemory;
            goto Cleanup;
        }

        // Tell our clients which run the synthetic character was added
        *pporOut = porNew;

        // Let us change our synthetic run
        porNew->MakeRunSynthetic();
        porNew->FillSynthData(synthtype);

        _listCurrent.SpliceInAfterMe(pTail, porNew);

        // Now change the original one run itself
        por->_lscpBase++; // for the synthetic character
        por->_chSynthsBefore++;

        // Update the tree info
        _treeInfo._lscpFrontier++;
        _treeInfo._chSynthsBefore++;
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   AppendAntiSynthetic
//
//  Synopsis:   Appends a anti-synthetic run
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLineServices::AppendAntiSynthetic(COneRun* por)
{
    LSERR lserr = lserrNone;
    LONG  cch   = por->_lscch;

    Assert(por->IsAntiSyntheticRun());
    Assert(por->_lscch == por->_lscchOriginal);

    // If the run is not in the list yet, please go and add it to the list
    if(por->_pNext==NULL && por->_pPrev==NULL)
    {
        _listCurrent.SpliceInAfterMe(_listCurrent._pTail, por);
    }

    // This run has now been processed
    por->_fNotProcessedYet = FALSE;

    // Update the tree info
    _treeInfo._lscpFrontier   -= cch;
    _treeInfo._chSynthsBefore -= cch;

    // Now change all the subsequent runs in the list
    por = por->_pNext;
    while(por)
    {
        por->_lscpBase       -= cch;
        por->_chSynthsBefore -= cch;
        por = por->_pNext;
    }

    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Member:     CLineServices::TerminateLine
//
//  Synopsis:   Close any open LS objects. This will add end of object
//              characters to the synthetic store for any open LS objects and
//              also optionally add a synthetic WCH_SECTIONBREAK (fAddEOS).
//              If it adds any synthetic character it will set *psynthAdded to
//              the type of the first synthetic character added. FetchRun()
//              callers should be sure to check the *psynthAdded value; if it
//              is not SYNTHTYPE_NONE then the run should be filled using
//              FillSynthRun() and returned to Line Services.
//
//-----------------------------------------------------------------------------
LSERR CLineServices::TerminateLine(COneRun* por, TL_ENDMODE tlEndMode, COneRun** pporOut)
{
    LSERR lserr = lserrNone;
    SYNTHTYPE synthtype;
    COneRun* porOut = NULL;
    COneRun* porRet;
    COneRun* pTail = _listCurrent._pTail;

    if(pTail)
    {
        int aObjRef[LSOBJID_COUNT];

        // Zero out the object refcount array.
        ZeroMemory(aObjRef, LSOBJID_COUNT*sizeof(int));

        // End any open LS objects.
        for(; pTail; pTail=pTail->_pPrev)
        {
            if(!pTail->_fIsStartOrEndOfObj)
            {
                continue;
            }

            synthtype = pTail->_synthType;
            WORD idObj = s_aSynthData[synthtype].idObj;

            // If this synthetic character starts or stops an LS object...
            if(idObj != idObjTextChp)
            {
                // Adjust the refcount up or down depending on whether the object
                // is started or ended.
                if(s_aSynthData[synthtype].fObjEnd)
                {
                    aObjRef[idObj]--;
                }
                if(s_aSynthData[synthtype].fObjStart)
                {
                    aObjRef[idObj]++;
                }

                // If the refcount is positive we have an unclosed object (we're
                // walking backward). Close it.
                if(aObjRef[idObj] > 0)
                {
                    synthtype = s_aSynthData[synthtype].typeEndObj;
                    Assert(synthtype!=SYNTHTYPE_NONE &&
                        s_aSynthData[synthtype].idObj==idObj &&
                        s_aSynthData[synthtype].fObjStart==FALSE &&
                        s_aSynthData[synthtype].fObjEnd==TRUE);

                    // If we see an open ruby object but the ruby main text
                    // has not been closed yet, then we must close it here
                    // before we can close off the ruby object by passing
                    // and ENDRUBYTEXT to LS.
                    if(idObj==LSOBJID_RUBY && _fIsRuby && !_fIsRubyText)
                    {
                        synthtype = SYNTHTYPE_ENDRUBYMAIN;
                        _fIsRubyText = TRUE;
                    }

                    // Be sure to inc lscp.
                    // BUGBUG (mikejoch) We should really be adjusting por here. As
                    // it is we aren't necessarily pointing at the correct run.
                    // This should be fixed when por positioning is fixed.
                    lserr = AppendSynth(por, synthtype, &porRet);
                    if(lserr != lserrNone)
                    {
                        // NOTE(SujalP): The linker (even in debug build) will
                        // not link in DumpList() since it is not called anywhere.
                        // This call here forces the linker to link in the DumpList
                        // function, so that we can use it during debugging.
                        goto Cleanup;
                    }

                    // Terminate line needs to return the pointer to the run
                    // belonging to the first synthetic character added.
                    if(!porOut)
                    {
                        porOut = porRet;
                    }

                    aObjRef[idObj]--;

                    Assert(aObjRef[idObj] == 0);
                }
            }
        }
    }

    if(tlEndMode != TL_ADDNONE)
    {
        // Add a synthetic section break character.  Note we add a section
        // break character as this has no width.
        synthtype = tlEndMode==TL_ADDLBREAK ? SYNTHTYPE_LINEBREAK : SYNTHTYPE_SECTIONBREAK;
        lserr = AppendSynth(por, synthtype, &porRet);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        porRet->_fNoTextMetrics = TRUE;

        if(tlEndMode == TL_ADDLBREAK)
        {
            porRet->_fMakeItASpace = TRUE;
            if(GetRenderer())
            {
                lserr = SetRenderingHighlights(porRet);
                if(lserr != lserrNone)
                {
                    goto Cleanup;
                }
            }
        }

        if(!porOut)
        {
            porOut = porRet;
        }
    }

    // Lock up the synthetic character store. We've terminated the line, so we
    // don't want anyone adding any more synthetics.
    FreezeSynth();

Cleanup:
    *pporOut = porOut;
    return lserr;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::IsSynthEOL
//
//  Synopsis:   Determines if there is some synthetic end of line mark at the
//              end of the synthetic array.
//
//  Returns:    TRUE if the synthetic array is terminated by a synthetic EOL
//              mark; otherwise FALSE.
//
//-----------------------------------------------------------------------------
BOOL CLineServices::IsSynthEOL()
{
    COneRun* pTail = _listCurrent._pTail;
    SYNTHTYPE synthEnd = SYNTHTYPE_NONE;

    if(pTail != NULL)
    {
        if(pTail->_fNotProcessedYet)
        {
            pTail = pTail->_pPrev;
        }
        while(pTail)
        {
            if(pTail->IsSyntheticRun())
            {
                synthEnd = pTail->_synthType;
                break;
            }
            pTail = pTail->_pNext;
        }
    }

    return (synthEnd==SYNTHTYPE_SECTIONBREAK
        || synthEnd==SYNTHTYPE_ENDPARA1
        || synthEnd==SYNTHTYPE_ALTENDPARA);
}

//-----------------------------------------------------------------------------
//
//  Function:   CPFromLSCPCore
//
//-----------------------------------------------------------------------------
LONG CLineServices::CPFromLSCPCore(LONG lscp, COneRun** ppor)
{
    COneRun* por = _listCurrent._pHead;
    LONG     cp  = lscp;

    while(por)
    {
        if(por->IsAntiSyntheticRun())
        {
            cp += por->_lscch;
        }
        else if(lscp>=por->_lscpBase && lscp<por->_lscpBase+por->_lscch)
        {
            break;
        }
        else if(por->IsSyntheticRun())
        {
            cp--;
        }
        por = por->_pNext;
    }

    if(ppor)
    {
        *ppor = por;
    }

    return cp;
}

//-----------------------------------------------------------------------------
//
//  Function:   LSCPFromCPCore
//
//  FUTURE(SujalP): The problem with this function is that it is computing lscp
//                  and is using it to terminate the loop. Probably the better
//                  approach would be to use Cp's to determine termination
//                  conditions. That would be a radical change and we leave
//                  that to be fixed in IE5+.
//
//                  Another change which we need to make is that we do not
//                  inc/dec the lscp (and cp in the function above) as we
//                  march along the linked list. All the loop has to do is
//                  ensure that we end up with the correct COneRun and from
//                  that it should be pretty easy for us to determine the
//                  lscp/cp to be returned.
//
//-----------------------------------------------------------------------------
LONG CLineServices::LSCPFromCPCore(LONG cp, COneRun** ppor)
{
    COneRun* por = _listCurrent._pHead;
    LONG lscp = cp;

    while(por)
    {
        if(lscp>=por->_lscpBase && lscp<por->_lscpBase+por->_lscch)
        {
            break;
        }
        if(por->IsAntiSyntheticRun())
        {
            lscp -= por->_lscch;
        }
        else if(por->IsSyntheticRun())
        {
            lscp++;
        }
        por = por->_pNext;
    }

    // If we have stopped at an anti-synthetic it means that the cp is within this
    // run. This implies that the lscp is the same for all the cp's in this run.
    if(por && por->IsAntiSyntheticRun())
    {
        Assert(por->WantsLSCPStop());
        lscp = por->_lscpBase;
    }
    else
    {
        while(por && !por->WantsLSCPStop())
        {
            por = por->_pNext;
            lscp++;
        }
    }

    Assert(!por || por->WantsLSCPStop());

    // It is possible that we can return a NULL por if there is a semi-valid
    // lscp that could be returned.
    if(ppor)
    {
        *ppor = por;
    }
    return lscp;
}

//-----------------------------------------------------------------------------
//
//  Function:   FindOneRun
//
//  Synopsis:   Given an lscp, find the one run if it exists in the current list
//
//  Returns:    The one run
//
//-----------------------------------------------------------------------------
COneRun* CLineServices::FindOneRun(LSCP lscp)
{
    COneRun* por = _listCurrent._pTail;

    if(!por)
    {
        por = NULL;
    }
    else if(lscp >= por->_lscpBase+por->_lscch)
    {
        por = NULL;
    }
    else
    {
        while(por)
        {
            if(lscp>=por->_lscpBase && lscp<por->_lscpBase+por->_lscch)
            {
                break;
            }
            por = por->_pPrev;
        }
    }
    return por;
}

//-----------------------------------------------------------------------------
//
//  Function:   FindPrevLSCP (member)
//
//  Synopsis:   Find the LSCP of the first actual character to preceed lscp.
//              Synthetic characters are ignored.
//
//  Returns:    LSCP of the character prior to lscp, ignoring synthetics. If no
//              characters preceed lscp, or lscp is beyond the end of the line,
//              then lscp itself is returned. Also returns a flag indicating if
//              any reverse objects exist between these two LSCPs.
//
//-----------------------------------------------------------------------------
LSCP CLineServices::FindPrevLSCP(LSCP lscp, BOOL* pfReverse)
{
    COneRun* por;
    LSCP lscpPrev = lscp;
    BOOL fReverse = FALSE;

    // Find the por matching this lscp.
    por = FindOneRun(lscp);

    // If lscp was outside the limits of the line just bail out.
    if(por == NULL)
    {
        goto cleanup;
    }

    Assert(lscp>=por->_lscpBase && lscp<por->_lscpBase+por->_lscch);

    if(por->_lscpBase < lscp)
    {
        // We're in the midst of a run. lscpPrev is just lscp - 1.
        lscpPrev--;
        goto cleanup;
    }

    // Loop over the pors
    while(por->_pPrev != NULL)
    {
        por = por->_pPrev;

        // If the por is a reverse object set fReverse.
        if(por->IsSyntheticRun() && s_aSynthData[por->_synthType].idObj==LSOBJID_REVERSE)
        {
            fReverse = TRUE;
        }

        // If the por is a text run then find the last lscp in it and break.
        if(por->IsNormalRun())
        {
            lscpPrev = por->_lscpBase + por->_lscch - 1;
            break;
        }
    }

cleanup:
    Assert(lscpPrev <= lscp);
    if(pfReverse != NULL)
    {
        // If we hit a reverse object but lscpPrev == lscp, then the reverse
        // object preceeds the first character in the line. In this case there
        // isn't actually a reverse object between the two LSCPs, since the
        // LSCPs are the same.
        *pfReverse = (fReverse && lscpPrev<lscp);
    }

    return lscpPrev;
}

//-----------------------------------------------------------------------------
//
//  Function:   FetchRun (member, LS callback)
//
//  Synopsis:   This is a key callback from lineservices.  LS calls this method
//              when performing LsCreateLine.  Here it is asking for a run, or
//              an embedded object -- whatever appears next in the stream.  It
//              passes us cp, and CLineServices (which we fool C++ into getting
//              to be the object of this method).  We return a bunch of stuff
//              about the next thing to put in the stream.
//
//  Returns:    lserrNone
//              lserrOutOfMemory
//
//-----------------------------------------------------------------------------
LSERR WINAPI CLineServices::FetchRun(
                        LSCP lscp,          // IN
                        LPCWSTR* ppwchRun,  // OUT
                        DWORD* pcchRun,     // OUT
                        BOOL* pfHidden,     // OUT
                        PLSCHP plsChp,      // OUT
                        PLSRUN* pplsrun )   // OUT
{
    LSERR       lserr = lserrNone;
    COneRun*    por;
    COneRun*    pTail;
    LONG        cchDelta;
    COneRun*    porOut;


    ZeroMemory(plsChp, sizeof(LSCHP)); // Otherwise, we're gonna forget and leave some bits on that we shouldn't.
    *pfHidden = FALSE;

    if(IsAdornment())
    {
        por = GetRenderer()->FetchLIRun(lscp, ppwchRun, pcchRun);
        CHPFromCF(por, por->GetCF());
        goto Cleanup;
    }

    pTail = _listCurrent._pTail;
    // If this was already cached before
    if(lscp < _treeInfo._lscpFrontier)
    {
        Assert(pTail);
        Assert(_treeInfo._lscpFrontier == pTail->_lscpBase+pTail->_lscch);
        while(pTail)
        {
            if(lscp >= pTail->_lscpBase)
            {
                // Should never get a AS run since the actual run will the run
                // following it and we should have found that one before since
                // we are looking from the end.
                Assert(!pTail->IsAntiSyntheticRun());

                // We should be in this run since 1) if check above 2) walking backwards
                AssertSz(lscp<pTail->_lscpBase+pTail->_lscch, "Inconsistent linked list");

                // Gotcha. Got a previously cached sucker
                por = pTail;
                Assert(por->_lscchOriginal == por->_lscch);
                cchDelta = lscp - por->_lscpBase;
                if(por->_fGlean)
                {
                    // We never have to reglean a synth or an antisynth run
                    Assert(por->IsNormalRun());

                    // NOTE(SujalP+MikeJoch):
                    // This can never happen because ls always fetches sequentially.
                    // If this happened it would mean that we were asked to fetch
                    // part of the run which was not gleaned. Hence the part before
                    // this one was not gleaned and hence not fetched. This violates
                    // the fact that LS will fetch all chars before the present one
                    // atleast once before it fetches the present one.
                    AssertSz(cchDelta==0, "CAN NEVER HAPPEN!!!!");

                    for(;;)
                    {
                        // We still have to be interested in gleaning
                        Assert(por->_fGlean);

                        // We will should never have a anti-synthetic run here
                        Assert(!por->IsAntiSyntheticRun());

                        // Now go and glean information into the run ...
                        lserr = GleanInfoFromTheRun(por, &porOut);
                        if(lserr != lserrNone)
                        {
                            goto Cleanup;
                        }

                        // Did the run get marked as Antisynth. If so then
                        // we need to ignore that run and go onto the next one.
                        if(por->IsAntiSyntheticRun())
                        {
                            Assert(por == porOut);

                            // The run was marked as an antisynthetic run. Be sure
                            // that no splitting was requested...
                            Assert(por->_lscch == por->_lscchOriginal);
                            Assert(por->_fNotProcessedYet);
                            AppendAntiSynthetic(por);
                            por = por->_pNext;
                        }
                        else
                        {
                            break;
                        }

                        // If we ran out of already cached runs (all the cached runs
                        // turned themselves into anti-synthetics) then we need to
                        // advance the frontier and fetch new runs from the story.
                        if(por == NULL)
                        {
                            goto NormalProcessing;
                        }
                    }

                    // The only time a different run is than the one passed in is returned is
                    // when the run has not been procesed as yet, and during processing we
                    // notice that to process it we need to append a synthetic character.
                    // The case to handle here is:
                    // <table><tr><td nowrap>abcd</td></tr></table>
                    Assert(porOut==por || por->_fNotProcessedYet);

                    if(porOut != por)
                    {
                        // If we added a synthetic, then the present run should not be split!
                        Assert(por->_lscch == por->_lscchOriginal);
                        Assert(por->_fNotProcessedYet);

                        // Remember we have to re-glean the information the next time we come around.
                        // However, we will not make the decision to append a synth that time since
                        // the synth has already been added this time and hence will fall into the
                        // else clause of this if and everything should be fine.
                        //
                        // DEPENDING ON WHETHER porOut WAS ADDED IN THE PRESENT GLEAN
                        // OR WAS ALREADY IN THE LIST, WE EITHER RETURN porOut OR por
                        por->_fGlean = TRUE;
                        por = porOut;
                    }
                    else
                    {
                        por->_fNotProcessedYet = FALSE;

                        // Did gleaning give us reason to further split the run?
                        if(por->_lscchOriginal > por->_lscch)
                        {
                            COneRun* porNew = SplitRun(por, por->_lscch);
                            if(!porNew)
                            {
                                lserr = lserrOutOfMemory;
                                goto Cleanup;
                            }
                            _listCurrent.SpliceInAfterMe(por, porNew);
                        }
                    }
                    por->_fGlean = FALSE;
                    cchDelta = 0;
                }

                // This is our quickest way outta here! We had already done all
                // the hard work before so just reuse it here
                *ppwchRun  = por->_pchBase + cchDelta;
                *pcchRun   = por->_lscch   - cchDelta;
                goto Cleanup;
            }
            pTail = pTail->_pPrev;
        } // while
        AssertSz(0, "Should never come here!");
    } // if


NormalProcessing:
    for(;;)
    {
        por = AdvanceOneRun(lscp);
        if(!por)
        {
            lserr = lserrOutOfMemory;
            goto Cleanup;
        }

        lserr = GleanInfoFromTheRun(por, &porOut);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        Assert(porOut);

        if(por->IsAntiSyntheticRun())
        {
            Assert(por == porOut);
            AppendAntiSynthetic(por);
        }
        else
        {
            break;
        }
    }

    if(por != porOut)
    {
        *ppwchRun = porOut->_pchBase;
        *pcchRun  = porOut->_lscch;
        por->_fGlean = TRUE;
        por->_fNotProcessedYet = TRUE;
        Assert(por->_lscch == por->_lscchOriginal); // be sure no splitting takes place
        Assert(porOut->_fCannotMergeRuns);
        Assert(porOut->IsSyntheticRun());

        if(por->_lscch)
        {
            COneRun* porLastSynth = porOut;
            COneRun* porTemp = porOut;

            // GleanInfoFromThrRun can add multiple synthetics to the linked
            // list, in which case we will have to jump across all of them
            // before we can add por to the list. (We need to add the por
            // because the frontier has already moved past that run).
            while(porTemp && porTemp->IsSyntheticRun())
            {
                porLastSynth = porTemp;
                porTemp = porTemp->_pNext;
            }

            // FUTURE: porLastSynth should equal pTail.  Add a check for this, and
            // remove above while loop.
            _listCurrent.SpliceInAfterMe(porLastSynth, por);
        }
        else
        {
            // Run not needed, please do not leak memory
            _listFree.SpliceIn(por);
        }

        // Finally remember that por is the run which we give to LS
        por = porOut;
    }
    else
    {
        por->_fNotProcessedYet = FALSE;
        *ppwchRun  = por->_pchBase;
        *pcchRun  = por->_lscch;
        por = AttachOneRunToCurrentList(por);
        if(!por)
        {
            lserr = lserrOutOfMemory;
            goto Cleanup;
        }
    }

Cleanup:
    if(lserr == lserrNone)
    {
        Assert(por);

        // We can never return an antisynthetic run to LS!
        Assert(!por->IsAntiSyntheticRun());

        if(por->_fCharsForNestedLayout && !IsAdornment())
        {
            // Give LS a junk character in this case. Fini will jump
            // accross the number of chars actually taken up by the
            // nested run owner.
            *ppwchRun = por->SetCharacter('A');
            *pcchRun = 1;
        }

        *pfHidden  = por->_fHidden;
        *plsChp    = por->_lsCharProps;
        *(PLSRUN*)pplsrun = por;
    }
    else
    {
        *(PLSRUN*)pplsrun = NULL;
    }
    return lserr;
}

//+----------------------------------------------------------------------------
//
//  Member:     CLineServices::AppendILSControlChar
//
//  Synopsis:   Appends an ILS object control character to the synthetic store.
//              This function allows us to begin and end line services objects
//              by inserting the control characters that begin and end them
//              into the text stream. It also keeps track of the state of the
//              object stack at the end of the synthetic store and returns the
//              synthetic type that matches the first added charcter.
//
//              A curiousity of Line Services dictates that ILS objects cannot
//              be overlapped; it is not legal to have:
//
//                  <startNOBR><startReverse><endNOBR><endReverse>
//
//              If this case were to arise the <endNOBR> would get ignored and
//              the NOBR object would continue until the line was terminated
//              by a TerminateLine() call. Furthermore, the behavior of ILS
//              objects is not always inherited; the reverse object inside of
//              a NOBR will still break.
//
//              To get around these problems it is necessary to keep the object
//              stack in good order. We define a hierarchy of objects and make
//              certain that whenever a new object is created any objects which
//              are lower in the heirarchy are closed and reopened after the
//              new object is opened. The current heirarchy is:
//
//                  Reverse objects     (highest)
//                  NOBR objects    
//                  Embedding objects   (lowest)
//
//              Additional objects (FE objects) will be inserted into this
//              heirarchy.
//
//              Embedding objects require no special handling due to their
//              special handling (recursive calls to line services). Thus the
//              only objects which currently require handling are the reverse
//              and NOBR objects.
//
//              If we apply our strategy to the overlapped case above, we will
//              end up with the following:
//
//                  <startNOBR><endNOBR><startReverse><startNOBR><endNOBR><endReverse>
//
//              As can be seen the objects are well ordered and each object's
//              start character is paired with its end character. NOBRs are
//              kept at the top of the stack, and reverse objects preceed them.
//
//              One problem which is introduced by this solution is the fact
//              that a break opprotunity is introduced between the two NOBR
//              objects. This can be fixed in the NOBR breaking code.
//
//  Returns:    An LSERR value. The function also returns synthetic character
//              at lscp in *psynthAdded. This is the first charcter added by
//              this function, NOT necessarily the character that matches idObj
//              and fOpen. For example (again using the case above) when we
//              ask to open the LSOBJID_REVERSE inside the NOBR object we will
//              return SYNTHTYPE_ENDNOBR in *psynthAdded (though we will also
//              append the SYNTHTYPE_REVERSE and START_NOBR to the store).
//
//-----------------------------------------------------------------------------
LSERR CLineServices::AppendILSControlChar(COneRun* por, SYNTHTYPE synthtype, COneRun** pporOut)
{
    Assert(synthtype != SYNTHTYPE_NONE);
    const SYNTHDATA& synthdata = s_aSynthData[synthtype];
    LSERR lserr = lserrNone;

#ifdef _DEBUG
    BOOL fOpen = synthdata.fObjStart;
#endif

    // We can only APPEND REAL OBJECTS to the store.
    Assert(synthdata.idObj != LSOBJID_TEXT);

    *pporOut = NULL;

    if(IsFrozen())
    {
        // We cannot add to the store if it is frozen.
        goto Cleanup;
    }

    // Handle the object.
    switch(synthdata.idObj)
    {
    case LSOBJID_NOBR:
        {
            // NOBR objects just need to be opened or closed. Note that we cannot
            // close an object unless it has been opened, nor can we open an object
            // if one is already open.
            Assert(!!_fNoBreakForMeasurer != !!fOpen);

#ifdef _DEBUG
            if(!fOpen)
            {
                COneRun* porTemp = _listCurrent._pTail;
                BOOL fFoundTemp = FALSE;

                while(porTemp)
                {
                    if(porTemp->IsSyntheticRun())
                    {
                        if(porTemp->_synthType != SYNTHTYPE_NOBR)
                        {
                            AssertSz(0, "Should have found an STARTNOBR before anyother synth");
                        }
                        fFoundTemp = TRUE;
                        break;
                    }
                    porTemp = porTemp->_pPrev;
                }
                AssertSz(fFoundTemp, "Did not find the STARTNOBR you are closing!");
            }
#endif
            lserr = AppendSynth(por, synthtype, pporOut);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }

            break;
        }

    case LSOBJID_REVERSE:
        {
            COneRun* porTemp = NULL;

            // If we've got an open NOBR object, we need to close it first.
            if(_fNoBreakForMeasurer)
            {
                COneRun* porNobrStart;

                lserr = AppendSynth(por, SYNTHTYPE_ENDNOBR, pporOut);
                if(lserr != lserrNone)
                {
                    goto Cleanup;
                }

                // We need to mark the starting por that it was artificially
                // terminated, so we can break appropriate in the ILS handlers.
                for(porNobrStart=(*pporOut)->_pPrev; porNobrStart; porNobrStart=porNobrStart->_pPrev)
                {
                    if(porNobrStart->IsSyntheticRun()
                        && porNobrStart->_synthType==SYNTHTYPE_NOBR)
                    {
                        break;
                    }
                }

                Assert(porNobrStart);

                // Can't break after this END-NOBR
                porNobrStart->_fIsArtificiallyTerminatedNOBR = 1;            
            }

            // Open or close the reverse object and update _nReverse.
            lserr = AppendSynth(por, synthtype, *pporOut?&porTemp:pporOut);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }

            // If there was a NOBR object before we opened or closed the
            // reverse object re-open a new NOBR object.
            if(_fNoBreakForMeasurer)
            {
                Assert(*pporOut);
                lserr = AppendSynth(por, SYNTHTYPE_NOBR, &porTemp);
                if(lserr != lserrNone)
                {
                    goto Cleanup;
                }

                // Can't break before this BEGIN-NOBR
                porTemp->_fIsArtificiallyStartedNOBR = 1;            
            }
        }
        break;

    case LSOBJID_RUBY:
        {
            // Here we close off all open LS objects, append the ruby synthetic
            // and then reopen the LS objects.
            lserr = AppendSynthClosingAndReopening(por, synthtype, pporOut);
            if(lserr != lserrNone)
            {
                goto Cleanup;
            }
        }
        break;

#ifdef _DEBUG
    default:
        // We only handle NOBR and reverse objects so far. Embedding objects
        // should not come here, and we don't support the FE objects so far.
        Assert(FALSE);
        break;
#endif
    }

Cleanup:
    // Make sure the store is still in good shape.
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   AppendSynthClosingAndReopening
//
//  Synopsis:   Appends a synthetic into the current one run store, but first
//              closes all open LS objects and then reopens them afterwards.
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLineServices::AppendSynthClosingAndReopening(COneRun* por, SYNTHTYPE synthtype, COneRun** pporOut)
{
    LSERR lserr = lserrNone;
    COneRun *porOut=NULL, *porTail;
    SYNTHTYPE curSynthtype;
    WORD idObj = s_aSynthData[synthtype].idObj;
    CStackDataAry<SYNTHTYPE, 16> arySynths;
    int i;

    int aObjRef[LSOBJID_COUNT];

    // Zero out the object refcount array.
    ZeroMemory(aObjRef, LSOBJID_COUNT*sizeof(int));

    *pporOut = NULL;

    // End any open LS objects.
    for(porTail=_listCurrent._pTail; porTail; porTail=porTail->_pPrev)
    {
        if(!porTail->_fIsStartOrEndOfObj)
        {
            continue;
        }

        curSynthtype = porTail->_synthType;
        WORD curIdObj = s_aSynthData[curSynthtype].idObj;

        // If this synthetic character starts or stops an LS object...
        if(curIdObj != idObj)
        {
            // Adjust the refcount up or down depending on whether the object
            // is started or ended.
            if(s_aSynthData[curSynthtype].fObjEnd)
            {
                aObjRef[curIdObj]--;
            }
            if(s_aSynthData[curSynthtype].fObjStart)
            {
                aObjRef[curIdObj]++;
            }

            // If the refcount is positive we have an unclosed object (we're
            // walking backward). Close it.
            if(aObjRef[curIdObj] > 0)
            {
                arySynths.AppendIndirect(&curSynthtype);
                curSynthtype = s_aSynthData[curSynthtype].typeEndObj;
                Assert(curSynthtype!=SYNTHTYPE_NONE &&
                    s_aSynthData[curSynthtype].idObj==curIdObj &&
                    s_aSynthData[curSynthtype].fObjStart==FALSE &&
                    s_aSynthData[curSynthtype].fObjEnd==TRUE);

                // NOTE (t-ramar): closing a Ruby object may require adding
                // two synths (endrubymain and endrubytext) depending on 
                // how it is currently open.  Right now, this function is being 
                // called only when Ruby synths are being appended so this situation 
                // will never occur.
                // if this function is used to append a non-Ruby synth,
                // the assert should be removed and code added that will close
                // a ruby object properly; that is, it will make sure that one or both
                // of [endrubymain] and [endrubytext] are appended appropriately.
                // Also, the code that reopens the closed objects (in the for loop below)
                // may need to be changed.
                Assert(curIdObj != LSOBJID_RUBY);

                lserr = AppendSynth(por, curSynthtype, &porOut);
                if(*pporOut == NULL) 
                {
                    *pporOut = porOut;
                }

                if(lserr != lserrNone)
                {
                    goto Cleanup;
                }

                aObjRef[curIdObj]--;

                Assert(aObjRef[curIdObj] == 0);
            }
        }
        else
        {
            break;
        }
    }

    // Append the synth that was passed in
    lserr = AppendSynth(por, synthtype, &porOut);
    if(lserr != lserrNone)
    {
        porOut = NULL;
        goto Cleanup;
    }
    if(*pporOut == NULL) 
    {
        *pporOut = porOut;
    }

    // Re-open the LS objects that we closed
    for(i=arySynths.Size(); i>0; i--)
    {
        curSynthtype = arySynths[i-1];        
        // NOTE (t-ramar): reopening a Ruby object may require adding
        // two synths (rubymain and endrubymain) depending on how it was
        // closed.  Right now, this function is being called only when
        // Ruby synths are being appended so this situation will never
        // occur.
        Assert(s_aSynthData[curSynthtype].idObj != LSOBJID_RUBY);

        Assert(curSynthtype!=SYNTHTYPE_NONE &&
            s_aSynthData[curSynthtype].fObjStart==TRUE &&
            s_aSynthData[curSynthtype].fObjEnd==FALSE);
        lserr = AppendSynth(por, curSynthtype, &porOut);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }
    }

Cleanup:
    // Make sure the store is still in good shape.
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   FigureNextPtp.
//
//  Synopsis:   Figures the ptp at the cp
//
//  Returns:    CTreePos *
//
//-----------------------------------------------------------------------------
CTreePos* CLineServices::FigureNextPtp(LONG cp)
{
    COneRun* por = _listCurrent._pTail;
    CTreePos* ptp = NULL;
    LONG cpAtEndOfOneRun;

    Assert(_treeInfo._fInited);

    if(!por)
    {
        goto Cleanup;
    }

    ptp = por->_ptp;
    if(ptp->IsBeginElementScope() && ptp->Branch()->NeedsLayout())
    {
        GetNestedElementCch(ptp->Branch()->Element(), &ptp);
    }
    cpAtEndOfOneRun = por->Cp() + (por->IsSyntheticRun()?0:por->_lscch);

    if(cpAtEndOfOneRun <= cp)
    {
        LONG cpAtEndOfPtp;

        if(!por->IsSyntheticRun() && cpAtEndOfOneRun==cp
            && por->_lscch==ptp->GetCch())
        {
            cpAtEndOfPtp = cpAtEndOfOneRun;
        }
        else
        {
            cpAtEndOfPtp = ptp->GetCp() + ptp->GetCch();
        }

        while(cpAtEndOfPtp <= cp)
        {
            Assert(ptp != _treeInfo._ptpLayoutLast);
            if(ptp->IsBeginElementScope() && ptp->Branch()->NeedsLayout())
            {
                GetNestedElementCch(ptp->Branch()->Element(), &ptp);
            }
            else
            {
                ptp = ptp->NextTreePos();
            }
            cpAtEndOfPtp = ptp->GetCp() + ptp->GetCch();
        }
    }
    else
    {
        while(por)
        {
            if(por->Cp()<=cp && por->Cp()+por->_lscch > cp)
            {
                break;
            }
            por = por->_pPrev;
        }
        Assert(por);

        // Note(SujalP): Bug63941 shows us the problem that if we ended up on an
        // antisynthetic run, then the cp could anywhere within that run and hence
        // the ptp could be anything -- not necessarily the ptp of the first cp of
        // the run. Hence, if the cp is not the first cp of the anitsynth run then
        // we will derive the ptp from basic principles.
        if(por->IsAntiSyntheticRun() && por->Cp()!=cp)
        {
            long junk;
            ptp = _treeInfo._pMarkup->TreePosAtCp(cp, &junk);
        }
        else
        {
            ptp = por->_ptp;
        }
    }

    {
        DEBUG_ONLY(long junk;)
        Assert(_treeInfo._pMarkup->TreePosAtCp(cp, &junk) == ptp);
    }
    Assert(cp>=ptp->GetCp() && cp<ptp->GetCp()+ptp->GetCch());

Cleanup:
    return ptp;
}


