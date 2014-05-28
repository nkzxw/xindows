
#ifndef __XINDOWSUTIL_COLLECTION_ATOMTABLE_H__
#define __XINDOWSUTIL_COLLECTION_ATOMTABLE_H__

class XINDOWS_PUBLIC CAtomTable : public CDataAry<CString>
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    CAtomTable() : CDataAry<CString>() {}
    HRESULT AddNameToAtomTable(LPCTSTR pch, long* plIndex);
    HRESULT GetAtomFromName(LPCTSTR pch, long* plIndex,
        BOOL fCaseSensitive=TRUE, BOOL fStartFromGivenIndex=FALSE);
    HRESULT GetNameFromAtom(long lIndex, LPCTSTR* ppch);
    void Free();
};

#endif //__XINDOWSUTIL_COLLECTION_ATOMTABLE_H__