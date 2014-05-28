
#ifndef __XINDOWS_SITE_TEXT_LSCACHE_H__
#define __XINDOWS_SITE_TEXT_LSCACHE_H__

class CDocument;
class CLSCache;
class CLineServices;

// We will keep atleast these many contexts around.
#define N_CACHE_MINSIZE 1

class CLSCacheEntry
{
    friend class CLSCache;

private:
    CLineServices* _pLS;
    BOOL _fUsed;
};

class CLSCache
{
private:
    DECLARE_CDataAry (CAryLSCache, CLSCacheEntry)
    CAryLSCache _aLSEntries;
    LONG _cUsed;

public:
    CLSCache() { _cUsed = 0; }
    CLineServices* GetFreeEntry(CMarkup* pMarkup, BOOL fStartUpLSDLL);
    void ReleaseEntry(CLineServices* pLS);
    void Dispose(BOOL fDisposeAll);
};

#endif //__XINDOWS_SITE_TEXT_LSCACHE_H__