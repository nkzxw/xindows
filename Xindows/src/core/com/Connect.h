
#ifndef __XINDOWS_CORE_COM_CONNECT_H__
#define __XINDOWS_CORE_COM_CONNECT_H__

class CBase;
class CConnectionPt;
class CConnectionPointContainer;

//+------------------------------------------------------------------------
//
//  struct CONNECTION_POINT_INFO (cpi) contains information about
//  an object's connection points.
//
//-------------------------------------------------------------------------
struct CONNECTION_POINT_INFO
{
    const IID*  piid;   // iid of connection point
    DISPID      dispid; // dispid under which to store sinks
};

#define CPI_ENTRY(iid, dispid)  \
{                               \
    &iid,                       \
    dispid,                     \
},                              \

#define CPI_ENTRY_NULL          \
{                               \
    NULL,                       \
    DISPID_UNKNOWN,             \
},


//+---------------------------------------------------------------------------
//
//  Class:      CEnumConnections (cenumc)
//
//  Purpose:    Enumerates connections per IEnumConnections.
//
//----------------------------------------------------------------------------
class CEnumConnections : public CBaseEnum
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    //  IEnum methods
    STDMETHOD(Next)(ULONG celt, void* reelt, ULONG* pceltFetched);
    STDMETHOD(Clone)(CBaseEnum** ppenum);

    // CEnumGeneric methods
    static HRESULT Create(int ccd, CONNECTDATA* pcd, CEnumConnections** ppenum);

protected:
    CEnumConnections();
    CEnumConnections(const CEnumConnections& enumc);
    ~CEnumConnections();

    CEnumConnections& operator=(const CEnumConnections& enumc); // don't define
};

//+---------------------------------------------------------------------------
//
//  Class:      CConnectionPt (CCP)
//
//  Purpose:    Manages a connection.
//
//----------------------------------------------------------------------------
class CConnectionPt : public IConnectionPoint
{
public:
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppvObj);
    STDMETHOD_(ULONG,AddRef)();
    STDMETHOD_(ULONG,Release)();

    // IConnectionPoint methods
    STDMETHOD(GetConnectionInterface)(IID* pIID);
    STDMETHOD(GetConnectionPointContainer)(IConnectionPointContainer** ppCPC);
    STDMETHOD(Advise)(LPUNKNOWN pUnkSink, DWORD* pdwCookie);
    STDMETHOD(Unadvise)(DWORD dwCookie);
    STDMETHOD(EnumConnections)(LPENUMCONNECTIONS* ppEnum);

    inline CConnectionPointContainer* MyCPC();
    inline const CONNECTION_POINT_INFO* MyCPI();

    int _index;
};

//+---------------------------------------------------------------------------
//
//  Class:      CConnectionPointContainer (CCPC)
//
//  Purpose:    Manages classes of connections.
//
//----------------------------------------------------------------------------
#define CONNECTION_POINTS_MAX   10 // Number of entries + 1

class CConnectionPointContainer : public IConnectionPointContainer
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE();

    CConnectionPointContainer(CBase* pBase, IUnknown* pUnkOuter);
    ~CConnectionPointContainer();

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    // IConnectionPointContainer methods
    STDMETHOD(EnumConnectionPoints)(IEnumConnectionPoints** ppEnum);
    STDMETHOD(FindConnectionPoint)(REFIID riid, IConnectionPoint** ppCP);

    virtual CONNECTION_POINT_INFO* GetCPI();

    ULONG           _ulRef;     // refcount
    CBase*          _pBase;     // parent class to this class
    IUnknown*       _pUnkOuter; // delegate QI to this object
    CConnectionPt   _aCP[CONNECTION_POINTS_MAX]; // the connection points
};

inline CConnectionPointContainer* CConnectionPt::MyCPC()
{
    return CONTAINING_RECORD(this, CConnectionPointContainer, _aCP[_index]);
}

inline const CONNECTION_POINT_INFO* CConnectionPt::MyCPI()
{
    return &(MyCPC()->GetCPI()[_index]);
}



//+------------------------------------------------------------------------
//
//  Connection point utilities
//
//-------------------------------------------------------------------------
HRESULT ConnectSink(IUnknown* pUnkSource, REFIID iid, IUnknown* pUnkSink, DWORD* pdwCookie);
HRESULT ConnectSinkWithCPC(IConnectionPointContainer* pCPC, REFIID iid, IUnknown* pUnkSink, DWORD* pdwCookie);
HRESULT DisconnectSink(IUnknown* pUnkSource, REFIID iid, DWORD* pdwCookie);

#endif //__XINDOWS_CORE_COM_CONNECT_H__