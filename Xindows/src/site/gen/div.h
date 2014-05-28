#ifndef __div_h__
#define __div_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "element.h"

/* header files for imported files */
#include "csite.h"

/* header files for imported files */
#include "txtedit.h"

/* header files for imported files */
#include "e1d.h"

#ifndef __IHTMLDivElement_FWD_DEFINED__
#define __IHTMLDivElement_FWD_DEFINED__
typedef interface IHTMLDivElement IHTMLDivElement;
#endif     /* __IHTMLDivElement_FWD_DEFINED__ */

#ifndef __IHTMLDivElement_INTERFACE_DEFINED__

#define __IHTMLDivElement_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLDivElement;


MIDL_INTERFACE("3050f200-98b5-11cf-bb82-00aa00bdce0b")
IHTMLDivElement : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE put_align(
         /* [in] */ BSTR v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_align(
         /* [out] */ BSTR * p) = 0;

    virtual HRESULT STDMETHODCALLTYPE put_noWrap(
         /* [in] */ VARIANT_BOOL v) = 0;

    virtual HRESULT STDMETHODCALLTYPE get_noWrap(
         /* [out] */ VARIANT_BOOL * p) = 0;

};

#endif     /* __IHTMLDivElement_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLDivElement;



EXTERN_C const GUID DIID_DispHTMLDivElement;


#ifndef _CDivElement_PROPDESCS_
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCDivElementalign;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCDivElementnoWrap;
EXTERN_C const PROPERTYDESC_NUMPROP s_propdescCDivElementnofocusrect;

#endif


#endif /*__div_h__*/

