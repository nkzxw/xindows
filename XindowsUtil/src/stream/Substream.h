
#ifndef __XINDOWSUTIL_STREAM_SUBSTREAM_H__
#define __XINDOWSUTIL_STREAM_SUBSTREAM_H__

//+---------------------------------------------------------------------------
//
//  Class:      CSubstream
//
//              An IStream wrapper which restricts access to an interval
//              of bytes. Can be initialized to be read-only or writable.
//
//              In read-only mode, the class operates on a clone of the
//              source stream, so multiple read-only CSubstreams can
//              be active at arbitary positions within a single source
//              stream. In this mode, all write operations fail.
//
//              In writable mode, the class operates directly on the
//              source stream, so only one writable substream can be
//              used on single source stream. A writable substream must
//              be positioned at the end of the source stream. Operations
//              on bytes before the beginning of the substream fail, but
//              operations beyond the end of the substream are allowed;
//              they extend the length of both the substream and the
//              source stream. Writable mode is provided for symmetry
//              with read-only mode.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CSubstream : public IStream
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv);
    ULONG _ulRefs;
    STDMETHOD_(ULONG, AddRef)(void)
    {
        return ++_ulRefs;
    }
    STDMETHOD_(ULONG, Release)(void)
    {
        if(--_ulRefs == 0)
        {
            _ulRefs = ULREF_IN_DESTRUCTOR;
            delete this;
            return 0;
        }
        return _ulRefs;
    }
    ULONG GetRefs(void)
    {
        return _ulRefs;
    }

    // ctor/dtor
    CSubstream();
    ~CSubstream();

    // Initialization
    HRESULT InitRead(IStream* pStream, ULARGE_INTEGER cb);
    HRESULT InitWrite(IStream* pStream);
    HRESULT InitClone(CSubstream* pOrigstream);
    void    Detach();

    // IStream methods
    STDMETHOD(Read)(void* pv, ULONG cb, ULONG* pcbRead);
    STDMETHOD(Write)(const void* pv, ULONG cb, ULONG* pcbRead);
    STDMETHOD(Seek)(LARGE_INTEGER ibMove, DWORD dwOrigin, ULARGE_INTEGER* pibNewPosition);
    STDMETHOD(SetSize)(ULARGE_INTEGER ibNewSize);
    STDMETHOD(CopyTo)(IStream* pStream, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten);
    STDMETHOD(Commit)(DWORD dwFlags);
    STDMETHOD(Revert)();
    STDMETHOD(LockRegion)(ULARGE_INTEGER ibOffset, ULARGE_INTEGER cb, DWORD dwFlags);
    STDMETHOD(UnlockRegion)(ULARGE_INTEGER ibOffset, ULARGE_INTEGER cb, DWORD dwFlags);
    STDMETHOD(Stat)(STATSTG* pstatstg, DWORD dwFlags);
    STDMETHOD(Clone)(IStream** ppStream);

private:
    // Data members
    CSubstream*     _pOrig;             // this, or, if this is a clone, the original substream
    IStream*        _pStream;           // the underlying original stream (or a clone)

    ULARGE_INTEGER  _ibStart;           // Beginning of substream
    ULARGE_INTEGER  _ibEnd;             // End of the substream
    ULONG           _fWritable : 1;     // TRUE if writable
};

XINDOWS_PUBLIC HRESULT CreateReadOnlySubstream(CSubstream** ppStreamOut, IStream* pStreamSource, ULARGE_INTEGER cb);
XINDOWS_PUBLIC HRESULT CreateWritableSubstream(CSubstream** ppStreamOut, IStream* pStreamSource);

#endif //__XINDOWSUTIL_STREAM_SUBSTREAM_H__