#ifndef __block_h__
#define __block_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

#ifndef __IHTMLBlockElement_FWD_DEFINED__
#define __IHTMLBlockElement_FWD_DEFINED__
typedef interface IHTMLBlockElement IHTMLBlockElement;
#endif     /* __IHTMLBlockElement_FWD_DEFINED__ */

#ifndef __IHTMLBlockElement_INTERFACE_DEFINED__

#define __IHTMLBlockElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLBlockElement;


MIDL_INTERFACE("3050f208-98b5-11cf-bb82-00aa00bdce0b")
IHTMLBlockElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_clear(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_clear(
         /* [out] */ BSTR * p) = 0;

};

#endif     /* __IHTMLBlockElement_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLBlockElement;



EXTERN_C const GUID DIID_DispHTMLBlockElement;


#ifndef _CBlockElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCBlockElementclear;

#endif


#endif /*__block_h__*/

