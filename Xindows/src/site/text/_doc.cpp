
#include "stdafx.h"
#include "_doc.h"

/*
*  CTxtBlk::InitBlock(cb)
*
*  @mfunc
*      Initialize this text block
*
*  @rdesc
*      TRUE if success, FALSE if allocation failed
*/
BOOL CTxtBlk::InitBlock(DWORD cb) //@parm initial size of the text block
{
    _pch = NULL;
    _cch = 0;
    _ibGap = 0;

    if(cb)
    {
        _pch = (TCHAR*)MemAllocClear(cb);
    }

    MemSetName((_pch, "CTxtBlk data"));

    if(_pch)
    {
        _cbBlock = cb;
        return TRUE;
    }
    else
    {
        _cbBlock = 0;
        return FALSE;
    }
}

/*
*  CTxtBlk::FreeBlock()
*
*  @mfunc
*      Free this text block
*
*  @rdesc
*      nothing
*/
VOID CTxtBlk::FreeBlock()
{
    MemFree(_pch);
    _pch = NULL;
    _cch = 0;
    _ibGap = 0;
    _cbBlock= 0;
}

/*
*  CTxtBlk::MoveGap(ichGap)
*
*  @mfunc
*      move gap in this text block
*
*  @rdesc
*      nothing
*/
void CTxtBlk::MoveGap(DWORD ichGap) //@parm new position for the gap
{
    DWORD cbMove;
    DWORD ibGapNew = CbOfCch(ichGap);
    LPBYTE pbFrom = (LPBYTE)_pch;
    LPBYTE pbTo;

    if(ibGapNew == _ibGap)
    {
        return;
    }

    if(ibGapNew < _ibGap)
    {
        cbMove = _ibGap - ibGapNew;
        pbFrom += ibGapNew;
        pbTo = pbFrom + _cbBlock - CbOfCch(_cch);
    }
    else
    {
        cbMove = ibGapNew - _ibGap;
        pbTo = pbFrom + _ibGap;
        pbFrom = pbTo + _cbBlock - CbOfCch(_cch);
    }

    MoveMemory(pbTo, pbFrom, cbMove);
    _ibGap = ibGapNew;
}

/*
*  CTxtBlk::ResizeBlock(cbNew)
*
*  @mfunc
*      resize this text block
*
*  @rdesc
*      FALSE if block could not be resized <nl>
*      non-FALSE otherwise
*
*  @comm
*  Side Effects: <nl>
*      moves text block
*/
BOOL CTxtBlk::ResizeBlock(DWORD cbNew) //@parm the new size
{
    TCHAR* pch;
    DWORD cbMove;
    HRESULT hr;

    AssertSz(cbNew>0, "resizing block to size <= 0");

    if(cbNew < _cbBlock)
    {
        if(_ibGap != CbOfCch(_cch))
        {
            // move text after gap down so that it doesn't get dropped
            cbMove = CbOfCch(_cch) - _ibGap;
            pch = _pch + CchOfCb(_cbBlock-cbMove);
            MoveMemory(pch-CchOfCb(_cbBlock-cbNew), pch, cbMove);
        }
        _cbBlock = cbNew;
    }
    pch = _pch;
    hr = MemRealloc((void**)&pch, cbNew);
    if(hr)
    {
        return _cbBlock == cbNew; // FALSE if grow, TRUE if shrink
    }

    _pch = pch;
    if(cbNew > _cbBlock)
    {
        if(_ibGap != CbOfCch(_cch)) // Move text after gap to end so that we don't end up with two gaps
        {
            cbMove = CbOfCch(_cch) - _ibGap;
            pch += CchOfCb(_cbBlock - cbMove);
            MoveMemory(pch+CchOfCb(cbNew-_cbBlock), pch, cbMove);
        }
        _cbBlock = cbNew;
    }

    return TRUE;
}



/*
*  CTxtArray::CTxtArray
*
*  @mfunc      Text array constructor
*
*/
CTxtArray::CTxtArray() : CArray<CTxtBlk>()
{
    AssertSz(CchOfCb(cbBlockMost)-cchGapInitial>=cchBlkInitmGapI*2,
        "cchBlockMax - cchGapInitial must be at least (cchBlockInitial - cchGapInitial) * 2");

    _cchText = 0;
}

/*
*  CTxtArray::~CTxtArray
*
*  @mfunc      Text array destructor
*/
CTxtArray::~CTxtArray()
{
    DWORD itb = Count();

    while(itb--)
    {
        Assert(Elem(itb) != NULL);
        Elem(itb)->FreeBlock();
    }
}

/*
*  CTxtArray::GetCch()
*
*  @mfunc      Computes and return length of text in this text array
*
*  @rdesc      The number of character in this text array
*
*  @devnote    This call may be computationally expensive; we have to
*              sum up the character sizes of all of the text blocks in
*              the array.
*/
long CTxtArray::GetCch() const
{
    DWORD itb = Count();
    long cch = 0;

    while(itb--)
    {
        Assert(Elem(itb) != NULL);
        cch += Elem(itb)->_cch;
    }

    return cch;
}

/*
*  CTxtArray::RemoveAll
*
*  @mfunc      Removes all characters in the array
*/
void CTxtArray::RemoveAll()
{
    DWORD itb = Count();

    while(itb--)
    {
        Assert(Elem(itb) != NULL);
        Elem(itb)->FreeBlock();
    }

    Clear(AF_DELETEMEM);

    _cchText = 0;
}

/*
*  CTxtArray::AddBlock(itbNew, cb)
*
*  @mfunc      create new text block
*
*  @rdesc
*      FALSE if block could not be added
*      non-FALSE otherwise
*
*  @comm
*  Side Effects:
*      moves text block array
*/
BOOL CTxtArray::AddBlock(
     DWORD   itbNew,    //@parm index of the new block
     LONG    cb)        //@parm size of new block; if <lt>= 0, default is used
{
    CTxtBlk* ptb;

    if(cb <= 0)
    {
        cb = cbBlockInitial;
    }

    AssertSz(cb>0, "CTxtArray::AddBlock() - adding block of size zero");
    AssertSz(cb<=cbBlockMost, "CTxtArray::AddBlock() - block too big");

    ptb = Insert(itbNew, 1);

    if(!ptb || !ptb->InitBlock(cb))
    {
        Trace0("TXTARRAT::AddBlock() - unable to allocate new block\n");
        return FALSE;
    }

    return TRUE;
}

/*
*  CTxtArray::RemoveBlocks(itbFirst, ctbDel)
*
*  @mfunc      remove a range of text blocks
*
*  @rdesc
*      nothing
*
*  @comm Side Effects: <nl>
*      moves text block array
*/
VOID CTxtArray::RemoveBlocks(
     DWORD itbFirst,    //@parm index of first block to remove
     DWORD ctbDel)      //@parm number of blocks to remove
{
    DWORD itb = itbFirst;
    DWORD ctb = ctbDel;

    AssertSz(itb+ctb<=Count(), "CTxtArray::RemoveBlocks(): not enough blocks");

    while(ctb--)
    {
        Assert(Elem(itb) != NULL);
        Elem(itb++)->FreeBlock();
    }

    Remove(itbFirst, ctbDel, AF_KEEPMEM);
}

/*
*  CTxtArray::CombineBlocks(itb)
*
*  @mfunc      combine adjacent text blocks
*
*  @rdesc
*      TRUE if blocks were combined, otherwise false
*
*  @comm
*  Side Effects: <nl>
*      moves text block array
*
*  @devnote
*      scans blocks from itb - 1 through itb + 1 trying to combine
*      adjacent blocks
*/
BOOL CTxtArray::CombineBlocks(DWORD itb) //@parm index of the first block modified
{
    DWORD ctb;
    DWORD cbT;
    CTxtBlk *ptb, *ptb1;
    BOOL  fRet = FALSE;

    if(itb > 0)
    {
        itb--;
    }

    ctb = min(3, int(Count()-itb));
    if(ctb <= 1)
    {
        return FALSE;
    }

    for(; ctb > 1; ctb--)
    {
        ptb  = Elem(itb);                       // Can we combine current
        ptb1 = Elem(itb+1);                     //  and next blocks ?
        cbT = CbOfCch(ptb->_cch + ptb1->_cch + cchGapInitial);
        if(cbT <= cbBlockInitial)
        {                                           // Yes
            if(cbT != ptb->_cbBlock && !ptb->ResizeBlock(cbT))
            {
                continue;
            }
            ptb ->MoveGap(ptb->_cch);               // Move gaps at ends of
            ptb1->MoveGap(ptb1->_cch);              //  both blocks
            CopyMemory(ptb->_pch + ptb->_cch,       // Copy next block text
                ptb1->_pch, CbOfCch(ptb1->_cch));   //  into current block
            ptb->_cch += ptb1->_cch;
            ptb->_ibGap += CbOfCch(ptb1->_cch);
            RemoveBlocks(itb+1, 1);                 // Remove next block
            fRet = TRUE;
        }
        else
            itb++;
    }

    return fRet;
}

/*
*  CTxtArray::SplitBlock(itb, ichSplit, cchFirst, cchLast, fStreaming)
*
*  @mfunc      split a text block into two
*
*  @rdesc
*      FALSE if the block could not be split <nl>
*      non-FALSE otherwise
*
*  @comm
*  Side Effects: <nl>
*      moves text block array
*/
BOOL CTxtArray::SplitBlock(
     DWORD itb,          //@parm index of the block to split
     DWORD ichSplit,     //@parm character index within block at which to split
     DWORD cchFirst,     //@parm desired extra space in first block
     DWORD cchLast,      //@parm desired extra space in new block
     BOOL fStreaming)    //@parm TRUE if streaming in new text
{
    LPBYTE pbSrc;
    LPBYTE pbDst;
    CTxtBlk *ptb, *ptb1;

    AssertSz(ichSplit>0 || cchFirst>0, "CTxtArray::SplitBlock(): splitting at beginning, but not adding anything");

    AssertSz(itb>=0, "CTxtArray::SplitBlock(): negative itb");
    ptb = Elem(itb);

    // compute size for first half
    AssertSz(cchFirst+ichSplit<=CchOfCb(cbBlockMost),
        "CTxtArray::SplitBlock(): first size too large");
    cchFirst += ichSplit + cchGapInitial;
    // Does not work!, need fail to occur. jonmat cchFirst = min(cchFirst, CchOfCb(cbBlockMost));
    // because our client expects cchFirst chars.

    // BUGBUG (cthrash) I *think* this should work but I also *know* this
    // code needs to be revisited.  Basically, there are Asserts sprinkled
    // about the code requiring the _cbBlock < cbBlockMost.  We can of course
    // exceed that if we insist on tacking on cchGapInitial (See AssertSz
    // above).  My modifications make certain you don't.  It would seem fine
    // except for the comment by jonmat.

    // (cthrash) jonmat's comment makes no sense to me.  The new buffer size
    // *wants* to cchFirst + ichSplit + cchGapInitial; it seems to me that
    // it'll be ok as long as we give it a gap >= 0.
    cchFirst = min(cchFirst, (DWORD)CchOfCb(cbBlockMost));

    // compute size for second half

    AssertSz(cchLast+ptb->_cch-ichSplit<=CchOfCb(cbBlockMost),
        "CTxtArray::SplitBlock(): second size too large");
    cchLast += ptb->_cch - ichSplit + cchGapInitial;
    // Does not work!, need fail to occur. jonmat cchLast = min(cchLast, CchOfCb(cbBlockMost));
    // because our client expects cchLast chars.

    // (cthrash) see comment above with cchFirst.
    cchLast = min(cchLast, (DWORD)CchOfCb(cbBlockMost));

    // allocate second block and move text to it

    // ***** moves rgtb ***** //
    // if streaming in, allocate a block that's as big as possible so that
    // subsequent additions of text are faster
    // we always fall back to smaller allocations so this won't cause
    // unneccesary errors
    // when we're done streaming we compress blocks, so this won't leave
    // a big empty gap
    if(fStreaming)
    {
        DWORD cb = cbBlockMost;
        const DWORD cbMin = CbOfCch(cchLast);

        while(cb>=cbMin && !AddBlock(itb+1, cb))
        {
            cb -= cbBlockCombine;
        }
        if(cb >= cbMin)
        {
            goto got_block;
        }
    }
    if(!AddBlock(itb+1, CbOfCch(cchLast)))
    {
        Trace0("CTxtArray::SplitBlock(): unabled to add new block\n");
        return FALSE;
    }

got_block:
    ptb1 = Elem(itb+1); // recompute ptb after rgtb moves
    ptb = Elem(itb);    // recompute ptb after rgtb moves
    ptb1->_cch = ptb->_cch - ichSplit;
    ptb1->_ibGap = 0;
    pbDst = (LPBYTE) (ptb1->_pch - ptb1->_cch) + ptb1->_cbBlock;
    ptb->MoveGap(ptb->_cch); // make sure pch points to a continuous block of all text in ptb.
    pbSrc = (LPBYTE) (ptb->_pch + ichSplit);
    CopyMemory(pbDst, pbSrc, CbOfCch(ptb1->_cch));
    ptb->_cch = ichSplit;
    ptb->_ibGap = CbOfCch(ichSplit);

    // resize the first block
    if(CbOfCch(cchFirst) != ptb->_cbBlock)
    {
        //$ FUTURE: don't resize unless growing or shrinking considerably
        if(!ptb->ResizeBlock(CbOfCch(cchFirst)))
        {
            // Review, if this fails then we need to delete all of the Added blocks, right? jonmat
            Trace0("TXTARRA::SplitBlock(): unabled to resize block\n");
            return FALSE;
        }
    }

    return TRUE;
}

/*
*  CTxtArray::ShrinkBlocks()
*
*  @mfunc      Shrink all blocks to their minimal size
*
*  @rdesc
*      nothing
*
*/
void CTxtArray::ShrinkBlocks()
{
    DWORD itb = Count();
    CTxtBlk* ptb;

    while(itb--)
    {
        ptb = Elem(itb);
        Assert(ptb);
        ptb->ResizeBlock(CbOfCch(ptb->_cch));
    }
}

/*
*  CTxtArray::GetChunk(ppch, cch, pchChunk, cchCopy)
*
*  @mfunc
*      Get content of text chunk in this text array into a string
*
*  @rdesc
*      remaining count of characters to get
*/
LONG CTxtArray::GetChunk(
     TCHAR** ppch,          //@parm ptr to ptr to buffer to copy text chunk into
     DWORD   cch,           //@parm length of pch buffer
     TCHAR*  pchChunk,      //@parm ptr to text chunk
     DWORD   cchCopy) const //@parm count of characters in chunk
{
    if(cch>0 && cchCopy>0)
    {
        if(cch < cchCopy)
        {
            cchCopy = cch;                      // Copy less than full chunk
        }
        CopyMemory(*ppch, pchChunk, cchCopy*sizeof(TCHAR));
        *ppch   += cchCopy;                     // Adjust target buffer ptr
        cch     -= cchCopy;                     // Fewer chars to copy
    }
    return cch;                                 // Remaining count to copy
}

