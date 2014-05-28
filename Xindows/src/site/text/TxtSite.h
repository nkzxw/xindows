
#ifndef __XINDOWS_SITE_TEXT_TXTSITE_H__
#define __XINDOWS_SITE_TEXT_TXTSITE_H__

class CTxtSite : public CSite
{
    DECLARE_CLASS_TYPES(CTxtSite, CSite)

public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    CTxtSite(ELEMENT_TAG etag, CDocument* pDoc);

    // Automation interfaces
#define _CTxtSite_
#include "../gen/txtedit.hdl"

    NV_DECLARE_TEAROFF_METHOD(get_onscroll, GET_onscroll, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onscroll, PUT_onscroll, (VARIANT));

    // Misc helpers

    //--------------------------------------------------------------
    // IPrivateUnknown members
    //--------------------------------------------------------------
    DECLARE_PLAIN_IUNKNOWN(CTxtSite);

    DECLARE_PRIVATE_QI_FUNCS(CBase)

    static CElement::ACCELS s_AccelsTxtSiteDesign;
    static CElement::ACCELS s_AccelsTxtSiteRun;
};

#endif //__XINDOWS_SITE_TEXT_TXTSITE_H__