#ifndef __e1d_h__
#define __e1d_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "txtedit.h"

/* header files for imported files */
#include "csite.h"

#ifndef __IHTMLDivPosition_FWD_DEFINED__
#define __IHTMLDivPosition_FWD_DEFINED__
typedef interface IHTMLDivPosition IHTMLDivPosition;
#endif     /* __IHTMLDivPosition_FWD_DEFINED__ */

#ifndef __IHTMLDivPosition_INTERFACE_DEFINED__

#define __IHTMLDivPosition_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLDivPosition;


MIDL_INTERFACE("3050f212-98b5-11cf-bb82-00aa00bdce0b")
IHTMLDivPosition : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_align(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_align(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLDivPosition_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLDivPosition;



EXTERN_C const GUID DIID_DispHTMLDivPosition;


#ifndef _C1DElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescC1DElementalign;

#endif


#endif /*__e1d_h__*/

