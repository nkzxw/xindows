
#ifndef __XINDOWS_SITE_UTIL_DATACACHE_H__
#define __XINDOWS_SITE_UTIL_DATACACHE_H__

// ========================  CDataCacheElem  =================================
// Each element of the data cache is described by the following:
struct CDataCacheElem
{
    // Pinter to the data, NULL if this element is free
    void* _pvData;

    // Index of the next free element for a free element
    // ref count, Crc and optional Id for a used element
    union
    {
        LONG _ielNextFree;
        struct
        {
            // BUGBUG? davidd: potential ENDIAN problems. For now ignored as
            // Free&Used elements are mutually exclusive I suppose...
            WORD _wPad;
            WORD _wCrc;
        } ielRef;
    };
    DWORD _cRef;
};

// ===========================  CDataCache  ==================================
// This array class ensures stability of the indexes. Elements are freed by
// inserting them in a free list, and the array is never shrunk.
// The first UINT of DATA is used to store the index of next element in the
// free list.
class CDataCacheBase
{
private:
    CDataCacheElem*  _pael;         // array of elements
    LONG             _cel;          // total count of elements (including free ones)
    LONG             _ielFirstFree; // first free element | FLBIT

#ifdef _DEBUG
public:
    LONG             _celsInCache;  // total count of elements (excluding free ones)
    LONG             _cMaxEls;      // maximum no. of cells in the cache                
#endif

    HRESULT CacheData(const void* pvData, LONG* piel, BOOL* pfDelete, BOOL fClone);

public:
    CDataCacheBase();
    virtual ~CDataCacheBase() {}

    //Cache this data w/o cloning.  Takes care of further memory management of data
    HRESULT CacheDataPointer(void** ppvData, LONG* piel);
    HRESULT CacheData(const void* pvData, LONG* piel) { return CacheData(pvData, piel, NULL, TRUE); }
    void    AddRefData(LONG iel);
    void    ReleaseData(LONG iel);

    long Size() const
    {
        return _cel;
    }

protected:
    CDataCacheElem* Elem(LONG iel) const
    {
        Assert(iel>=0 && iel<_cel);
        return (_pael+iel);
    }

    HRESULT Add(const void* pvData, LONG* piel, BOOL fClone);
    void    Free(LONG ielFirst);
    void    Free();
    LONG    Find(const void* pvData) const;

    virtual void    DeletePtr(void* pvData) = 0;
    virtual HRESULT InitData(CDataCacheElem* pel, const void* pvData, BOOL fClone) = 0;
    virtual void    PassivateData(CDataCacheElem* pel) = 0;
    virtual WORD    ComputeDataCrc(const void* pvData) const = 0;
    virtual BOOL    CompareData(const void* pvData1, const void* pvData2) const = 0;

#ifdef _DEBUG
    void CheckFreeChainFn() const;
#endif
};

template<class DATA>
class CDataCache : public CDataCacheBase
{
public:
    virtual ~CDataCache() { Free(); }

    const DATA& operator[](LONG iel) const
    {
        Assert(Elem(iel)->_pvData != NULL);
        return *(DATA*)(Elem(iel)->_pvData);
    }

    const DATA* ElemData(LONG iel) const
    {
        return Elem(iel)->_pvData?(DATA*)(Elem(iel)->_pvData):NULL;
    }

protected:
    virtual void DeletePtr(void* pvData)
    {
        Assert(pvData);
        delete (DATA*)pvData;
    }
    virtual HRESULT InitData(CDataCacheElem* pel, const void* pvData, BOOL fClone)
    {
        Assert(pvData);
        if(fClone)
        {
            return ((DATA*)pvData)->Clone((DATA**)&pel->_pvData);
        }
        else
        {
            pel->_pvData = (void*)pvData;
            return S_OK;
        }
    }

    virtual void PassivateData(CDataCacheElem* pel)
    {
        delete (DATA*)pel->_pvData;
    }

    virtual WORD ComputeDataCrc(const void* pvData) const
    {
        Assert(pvData);
        return ((DATA*)pvData)->ComputeCrc();
    }

    virtual BOOL CompareData(const void* pvData1, const void* pvData2) const
    {
        Assert(pvData1 && pvData2);
        return ((DATA*)pvData1)->Compare((DATA*)pvData2);
    }
};

#endif //__XINDOWS_SITE_UTIL_DATACACHE_H__