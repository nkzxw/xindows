
#ifndef __XINDOWSUTIL_STREAM_DATASTREAM_H__
#define __XINDOWSUTIL_STREAM_DATASTREAM_H__

class CSubstream;

//+---------------------------------------------------------------------------
//
//  Class:      CDataStream
//
//              Useful for robust reading and writing of various data types
//              to a stream. (Unlike IStream, we don't succeed on partial
//              reads.)
//
//              Makes no attempt to be particularly efficient.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CDataStream
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    CDataStream() : _aryLocations() { memset(this, 0, sizeof(*this)); }
    CDataStream(IStream* pStream) : _aryLocations() { memset(this, 0, sizeof(*this)); Init(pStream); } // for convenience
    ~CDataStream() { ReleaseInterface(_pStream); }

    void Init(IStream* pStream)
    {
        Assert(pStream && !_pStream);
        pStream->AddRef();
        _pStream = pStream;
    }

    // Saving basic data types
    HRESULT SaveData(void* pv, ULONG cb);
    HRESULT SaveDataLater(DWORD* pdwCookie, ULONG cb);
    HRESULT SaveDataNow(DWORD dwCookie, void* pv, ULONG cb);
    HRESULT LoadData(void* pv, ULONG cb);

    HRESULT SaveDword(DWORD dw)   { RRETURN(SaveData(&dw, sizeof(DWORD))); }
    HRESULT LoadDword(DWORD* pdw) { RRETURN(LoadData(pdw, sizeof(DWORD))); }
    HRESULT SaveString(TCHAR* pch);
    HRESULT LoadString(TCHAR** ppch);
    HRESULT SaveCStr(const CString* pcstr);
    HRESULT LoadCStr(CString* pcstr);

    // Saving substreams
    HRESULT BeginSaveSubstream(IStream** ppSubstream);
    HRESULT EndSaveSubstream();
    HRESULT LoadSubstream(IStream** ppSubstream);

#ifdef _DEBUG
    // Debugging Streams
    void DumpStreamInfo();
#endif

    // Direct access to underlying stream
    IStream* _pStream;

private:
    CSubstream* _pSubstream;        // The substream being saved (if any)
    DWORD       _dwLengthCookie;    // Cookie to the position reserved for the substream length

    class CLocation
    {
    public:
        LARGEINT    _ib;
        ULONG       _cb;
        ULONG       _dwCookie;
    };

    CDataAry<CLocation> _aryLocations;
};

#endif //__XINDOWSUTIL_STREAM_DATASTREAM_H__