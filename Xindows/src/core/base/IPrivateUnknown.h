
#ifndef __XINDOWS_CORE_BASE_IPRIVATEUNKNOWN_H__
#define __XINDOWS_CORE_BASE_IPRIVATEUNKNOWN_H__

interface IPrivateUnknown
{
public:
    STDMETHOD(PrivateQueryInterface)(REFIID riid, void** ppv) = 0;
    STDMETHOD_(ULONG, PrivateAddRef)() = 0;
    STDMETHOD_(ULONG, PrivateRelease)() = 0;
};

#endif //__XINDOWS_CORE_BASE_IPRIVATEUNKNOWN_H__