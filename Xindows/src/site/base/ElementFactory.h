
#ifndef __XINDOWS_SITE_BASE_ELEMENTFACTORY_H__
#define __XINDOWS_SITE_BASE_ELEMENTFACTORY_H__

class CElementFactory : public CBase
{
    DECLARE_CLASS_TYPES(CElementFactory, CBase)
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CElementFactory() {}
    ~CElementFactory() {}
    CDocument* _pDoc;

    DECLARE_PRIVATE_QI_FUNCS(CBase)
    HRESULT QueryInterface(REFIID iid, void** ppv){ return PrivateQueryInterface(iid, ppv); }

    struct CLASSDESC
    {
        CBase::CLASSDESC _classdescBase;
        void* _apfnTearOff;
    };
};

#endif //__XINDOWS_SITE_BASE_ELEMENTFACTORY_H__