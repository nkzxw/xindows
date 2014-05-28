
#include "stdafx.h"
#include "Array2.h"

#define celGrow     8

/*
*  CArrayBase::ArAdd
*
*  @mfunc  Adds <p celAdd> elements to the end of the array.
*
*  @rdesc  A pointer to the start of the new elements added.  If non-NULL,
*  <p pielIns> will be set to the index at which elements were added.
*/
void* CArrayBase::ArAdd(
        DWORD    celAdd,    // the number of elements to add
        DWORD*   pielIns)   // where to store the index of the first element added
{
    char* pel;
    DWORD celNew;

    if(_cel+celAdd > _celMax) // need to grow
    {
        HRESULT hr;

        // When we grow the array we grow it in units of celGrow.  However,
        // to make sure we don't grow small arrays too large, and get too much
        // unused space, we allocate only as much space as needed up to
        // celGrow.
        if(_cel+celAdd <= celGrow)
        {
            celNew = _cel + celAdd - _celMax;
        }
        else
        {
            celNew = max(DWORD(celGrow), celAdd+celGrow-celAdd%celGrow );
        }

        pel = _prgel;

        hr = MemRealloc((void**)&pel, (_celMax+celNew)*_cbElem);

        MemSetName((pel, "CArrayBase data - %d elements", celNew));

        if(hr)
        {
            return NULL;
        }

        _prgel = pel;

        pel += _cel * _cbElem;

        _celMax += celNew;
    }
    else
    {
        pel = _prgel + _cel*_cbElem;
    }

    ZeroMemory(pel, celAdd*_cbElem);

    if(pielIns)
    {
        *pielIns = _cel;
    }

    _cel += celAdd;

    return pel;
}

/*
*  CArrayBase::ArInsert
*
*  @mfunc Inserts <p celIns> new elements at index <p iel>
*
*  @rdesc A pointer to the newly inserted elements.  Will be NULL on
*  failure.
*/
void* CArrayBase::ArInsert(
       DWORD iel,       //@parm the index at which to insert
       DWORD celIns)    //@parm the number of elements to insert
{
    char*   pel;
    DWORD   celNew;
    HRESULT hr;

    AssertSz(iel<=_cel, "CArrayBase::Insert() - Insert out of range");

    if(iel >= _cel)
    {
        return ArAdd(celIns, NULL);
    }

    if(_cel+celIns > _celMax) // need to grow
    {
        AssertSz(_prgel, "CArrayBase::Insert() - Growing a non existent array !");

        celNew = max(DWORD(celGrow), celIns+celGrow-celIns%celGrow);
        pel = _prgel;
        hr = MemRealloc((void**)&pel, (_celMax+celNew)*_cbElem);
        if(hr)
        {
            AssertSz(FALSE, "CArrayBase::Insert() - Couldn't realloc line array");
            return NULL;
        }
        MemSetName((pel, "CArrayBase data - %d elements", celNew));

        _prgel = pel;
        _celMax += celNew;
    }
    pel = _prgel + iel*_cbElem;
    if(iel < _cel) // Nove Elems up to make room for new ones
    {
        memmove(pel+celIns*_cbElem, pel, (_cel-iel)*_cbElem);
        ZeroMemory(pel, celIns*_cbElem);
    }

    _cel += celIns;
    return pel;
}

/*
*  CArrayBase::Remove
*
*  @mfunc  Removes the <p celFree> elements from the array starting at index
*  <p ielFirst>.  If <p celFree> is negative, then all elements after
*  <p ielFirst> are removed.
*
*  @rdesc nothing
*/
void CArrayBase::Remove(
        DWORD       ielFirst,   //@parm the index at which elements should be removed
        LONG        celFree,    //@parm the number of elements to remove.
        ArrayFlag   flag)       //@parm what to do with the left over memory (delete or leave alone).
{
    char* pel;

    if(celFree < 0)
    {
        celFree = _cel - ielFirst;
    }

    AssertSz(ielFirst+celFree<=_cel, "CArrayBase::Free() - Freeing out of range");

    if(_cel > ielFirst+celFree)
    {
        pel = _prgel + ielFirst*_cbElem;
        memmove(pel, pel+celFree*_cbElem, (_cel-ielFirst-celFree)*_cbElem);
    }

    _cel -= celFree;

    if((flag==AF_DELETEMEM) && _cel<_celMax-celGrow)
    {
        HRESULT hr;

        // shrink array
        _celMax = _cel + celGrow - _cel%celGrow;
        pel = _prgel;
        hr = MemRealloc((void**)&pel, _celMax*_cbElem);
        // we don't care if it fails since we're shrinking
        if(!hr)
        {
            _prgel = pel;
        }
    }
}

/*
*  CArrayBase::Clear
*
*  @mfunc  Clears the entire array, potentially deleting all of the memory
*  as well.
*
*  @rdesc  nothing
*/
void CArrayBase::Clear(ArrayFlag flag)  // @parm Indicates what should be done with the memory
                                        // in the array.  One of AF_DELETEMEM or AF_KEEPMEM
{
    if(flag == AF_DELETEMEM)
    {
        MemFree(_prgel);
        _prgel = NULL;
        _celMax = 0;
    }
    _cel = 0;
}

/*
*  CArrayBase::Replace
*
*  @mfunc  Replaces the <p celRepl> elements at index <p ielRepl> with the
*  contents of the array specified by <p par>.  If <p celRepl> is negative,
*  then the entire contents of <p this> array starting at <p ielRepl> should
*  be replaced.
*
*  @rdesc  Returns TRUE on success, FALSE otherwise.
*/
BOOL CArrayBase::Replace(
         DWORD          ielRepl,//@parm the index at which replacement should occur
         LONG           celRepl,//@parm the number of elements to replace (may be negative, indicating that all).
         CArrayBase*    par)    //@parm the array to use as the replacement source
{
    DWORD celMove = 0;
    DWORD celIns = par->Count();

    if(celRepl < 0)
    {
        celRepl = _cel - ielRepl;
    }

    AssertSz(ielRepl+celRepl<=_cel, "CArrayBase::ArReplace() - Replacing out of range");

    celMove = min(celRepl, (LONG)celIns);

    if(celMove > 0)
    {
        memmove(Elem(ielRepl), par->Elem(0), celMove*_cbElem);
        celIns -= celMove;
        celRepl -= celMove;
        ielRepl += celMove;
    }

    Assert(celIns >= 0);
    Assert(celRepl >= 0);
    Assert(celIns+celMove == par->Count());

    if(celIns > 0)
    {
        Assert(celRepl == 0);
        void* pelIns = ArInsert(ielRepl, celIns);
        if(!pelIns)
        {
            return FALSE;
        }
        memmove(pelIns, par->Elem(celMove), celIns*_cbElem);
    }
    else if(celRepl > 0)
    {
        Remove(ielRepl, celRepl, AF_DELETEMEM);
    }

    return TRUE;
}