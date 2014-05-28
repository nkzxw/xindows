
#ifndef __XINDOWS_CORE_COM_EXCEPINFO_H__
#define __XINDOWS_CORE_COM_EXCEPINFO_H__

inline void InitEXCEPINFO(EXCEPINFO* pEI)
{
    if(pEI)
    {
        memset(pEI, 0, sizeof(*pEI));
    }
}

void FreeEXCEPINFO(EXCEPINFO* pEI);

class CExcepInfo : public EXCEPINFO
{
public:
    CExcepInfo() { InitEXCEPINFO(this); }
    ~CExcepInfo() { FreeEXCEPINFO(this); }
};

#endif //__XINDOWS_CORE_COM_EXCEPINFO_H__