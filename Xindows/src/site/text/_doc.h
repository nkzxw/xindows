
#ifndef __XINDOWS_SITE_TEXT_DOC_H__
#define __XINDOWS_SITE_TEXT_DOC_H__


#ifdef UNICODE
#define CbOfCch(_x) (DWORD)((_x)*sizeof(WCHAR))
#define CchOfCb(_x) (DWORD)((_x)/sizeof(WCHAR))
#else
#define CbOfCch(_x) (_x)
#define CchOfCb(_x) (_x)
#endif

#define cbBlockCombine      CbOfCch(3072)
#define cbBlockMost         CbOfCch(49152)
#define cbBlockInitial      CbOfCch(4096)
#define cchGapInitial       128
#define cchBlkCombmGapI     (CchOfCb(cbBlockCombine)-cchGapInitial)
#define cchBlkInitmGapI     (CchOfCb(cbBlockInitial)-cchGapInitial)

#define cchBlkInsertmGapI   (CchOfCb(cbBlockInitial)*5-cchGapInitial)

/*
*  CTxtRun
*
*  @class  Formalizes a run of text. A range of text with same attribute,
* (see CFmtDesc) or within the same line (see CLine), etc. Runs are kept
* in arrays (see CArray) and are pointed to by CRunPtr's of various kinds.
* In general the character position of a run is computed by summing the
* length of all preceding runs, altho it may be possible to start from
* some other cp, e.g., for CLines, from CDisplay::_cpFirstVisible.
*/
class CTxtRun
{
public:
    CTxtRun() { _cch = 0; }

    long _cch;
};


/*
*  CTxtBlk
*
*  @class  A text block; a chunk of UNICODE text with a buffer gap to allow
*  for easy insertions and deletions.
*
*  @base   protected | CTxtRun
*
*  @devnote    A text block may have four states: <nl>
*      NULL:   No data allocated for the block <nl>
*              <md CTxtBlk::_pch> == NULL  <nl>
*              <md CTxtRun::_cch> == 0     <nl>
*              <md CTxtBlk::_ibGap> == 0   <nl>
*              <md CTxtBlk::_cbBlock> == 0 <nl>
*
*      empty:  All of the available space is a buffer gap <nl>
*              <md CTxtBlk::_pch> != NULL  <nl>
*              <md CTxtRun::_cch> == 0     <nl>
*              <md CTxtBlk::_ibGap> == 0   <nl>
*              <md CTxtBlk::_cbBlock> <gt>= 0  <nl>
*
*      normal: There is both data and a buffer gap <nl>
*              <md CTxtBlk::_pch> != NULL  <nl>
*              <md CTxtRun::_cch> != 0     <nl>
*              <md CTxtBlk::_ibGap> != 0   <nl>
*              <md CTxtBlk::_cbBlock> <gt>= 0  <nl>
*
*      full:   The buffer gap is of zero size <nl>
*              <md CTxtBlk::_pch> != NULL  <nl>
*              <md CTxtRun::_cch> <gt>= 0  <nl>
*              <md CTxtBlk::_ibGap> <gt> 0 <nl>
*              <md CTxtBlk::_cbBlock> == _cch * sizeof(WCHAR) <nl>
*
*  The position of the buffer gap is given by _ibGap.  With _cch and _cbBlock,
*  it's possible to figure out the *size* of the gap by simply calculating:
*  <nl>
*      size = _cbBlock - (_cch * sizeof(character))
*
*/
class CTxtBlk : public CTxtRun
{
    friend class CTxtPtr;
    friend class CTxtArray;

protected:
    CTxtBlk() { InitBlock( 0 ); }
    ~CTxtBlk() { FreeBlock(); }

    //@cmember  Initializes the block to the
    //# of bytes given by <p cb>
    BOOL InitBlock(DWORD cb);
    //@cmember  Sets a block to the NULL state
    VOID FreeBlock();
    //@cmember  Moves the buffer gap in a
    //block
    VOID MoveGap(DWORD ichGap);
    //@cmember  Resizes a block to <p cbNew>
    //bytes
    BOOL ResizeBlock(DWORD cbNew);

private:
    TCHAR*  _pch;       // Pointer to the text data
    DWORD   _ibGap;     // BYTE offset of the gap
    DWORD   _cbBlock;   // Size of the block in bytes
};

/*
*  CTxtArray
*
*  @class  A dynamic array of <c CTxtBlk> classes
*
*  @base public | CArray<lt>CTxtBlk<gt>
*/
class CTxtArray : public CArray<CTxtBlk>
{
    friend class CTxtPtr;
    friend class CMarkup;

public:
    DECLARE_MEMALLOC_NEW_DELETE()

    //@cmember  Constructor
    CTxtArray();
    //@cmember  Destructor
    ~CTxtArray();
    //@cmember  Gets the total number of
    //characters in the array.
    long GetCch () const;

private:
    //@cmember Removes all data from the array
    VOID RemoveAll();

    BOOL AddBlock(DWORD itbNew, LONG cb);
    //@cmember  Removes the given number of
    //blocks
    VOID RemoveBlocks(DWORD itbFirst, DWORD ctbDel);
    //@cmember  Combines blocks adjacent to itb
    BOOL CombineBlocks(DWORD itb);
    //@cmember  Splits a block
    BOOL SplitBlock(DWORD itb, DWORD ichSplit, DWORD cchFirst, DWORD cchLast, BOOL fStreaming);
    //@cmember  Shrinks all blocks to their minimal
    //size
    VOID ShrinkBlocks();
    //@cmember  Copies a chunk of text into the
    //given location
    LONG GetChunk(TCHAR** ppch, DWORD cch, TCHAR* pchChunk, DWORD cchCopy) const;
    //@cmember  The total number of characters in the
    //this text array.
    DWORD _cchText;
};

#endif //__XINDOWS_SITE_TEXT_DOC_H__