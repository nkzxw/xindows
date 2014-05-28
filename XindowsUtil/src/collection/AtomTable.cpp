
#include "stdafx.h"
#include "AtomTable.h"

BOOL _tcsequal(const TCHAR* string1, const TCHAR* string2)
{
    // This function optimizes the case where all we want to find
    // is if the strings are equal and do not care about the relative
    // ordering of the strings.
    while(*string1)
    {
        if(*string1 != *string2)
        {
            return FALSE;
        }

        string1 += 1;
        string2 += 1;

    }
    return (*string2)?(FALSE):(TRUE);
}

HRESULT CAtomTable::AddNameToAtomTable(LPCTSTR pch, long* plIndex)
{
    HRESULT     hr = S_OK;
    long        lIndex;
    CString*    pstr;

    for(lIndex=0; lIndex<Size(); lIndex++)
    {
        pstr = (CString*)Deref(sizeof(CString), lIndex);
        if(_tcsequal(pch, *pstr))
        {
            break;
        }
    }
    if(lIndex == Size())
    {
        CString cstr;

        // Not found, so add element to array.
        hr = cstr.Set(pch);
        if(hr)
        {
            goto Cleanup;
        }

        hr = AppendIndirect(&cstr);
        if(hr)
        {
            goto Cleanup;
        }

        // The array now owns the memory for the cstr, so take it away from
        // the cstr on the stack.
        cstr.TakePch();
    }

    if(plIndex)
    {
        *plIndex = lIndex;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CAtomTable::GetAtomFromName(LPCTSTR pch, long* plIndex,
        BOOL fCaseSensitive/*=TRUE*/, BOOL fStartFromGivenIndex/*=FALSE*/)
{
    long        lIndex;
    HRESULT     hr = S_OK;
    CString*    pstr;

    if(fStartFromGivenIndex)
    {
        lIndex = *plIndex;
    }
    else
    {
        lIndex = 0;
    }

    for(; lIndex<Size(); lIndex++)
    {
        pstr = (CString*)Deref(sizeof(CString), lIndex);
        if(fCaseSensitive)
        {
            if(_tcsequal(pch, *pstr))
            {
                break;
            }
        }
        else
        {
            if(_tcsicmp(pch, *pstr) == 0)
            {
                break;
            }
        }
    }

    if(lIndex == Size())
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    if(plIndex)
    {
        *plIndex = lIndex;
    }

Cleanup:    
    RRETURN(hr);
}

HRESULT CAtomTable::GetNameFromAtom(long lIndex, LPCTSTR* ppch)
{
    HRESULT     hr = S_OK;
    CString*    pcstr;

    if(Size() <= lIndex)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    pcstr = (CString*)Deref(sizeof(CString), lIndex);
    *ppch = (TCHAR*)*pcstr;

Cleanup:    
    RRETURN1(hr, DISP_E_MEMBERNOTFOUND);
}

void CAtomTable::Free()
{
    CString*    pcstr;
    long        i;

    for(i=0; i<Size(); i++)
    {
        pcstr = (CString*)Deref(sizeof(CString), i);
        pcstr->Free();
    }
    DeleteAll();
}