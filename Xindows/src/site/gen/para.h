#ifndef __para_h__
#define __para_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

#ifndef __IHTMLParaElement_FWD_DEFINED__
#define __IHTMLParaElement_FWD_DEFINED__
typedef interface IHTMLParaElement IHTMLParaElement;
#endif     /* __IHTMLParaElement_FWD_DEFINED__ */

#ifndef __IHTMLParaElement_INTERFACE_DEFINED__

#define __IHTMLParaElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLParaElement;


MIDL_INTERFACE("3050f1f5-98b5-11cf-bb82-00aa00bdce0b")
IHTMLParaElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_align(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_align(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLParaElement_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLParaElement;



EXTERN_C const GUID DIID_DispHTMLParaElement;


#ifndef _CParaElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCParaElementalign;

#endif


#endif /*__para_h__*/

