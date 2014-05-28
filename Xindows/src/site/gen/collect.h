#ifndef __collect_h__
#define __collect_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "types.h"

/* header files for imported files */
#include "element.h"

#ifndef __IHTMLElementCollection_FWD_DEFINED__
#define __IHTMLElementCollection_FWD_DEFINED__
typedef interface IHTMLElementCollection IHTMLElementCollection;
#endif     /* __IHTMLElementCollection_FWD_DEFINED__ */

#ifndef __IHTMLElementCollection2_FWD_DEFINED__
#define __IHTMLElementCollection2_FWD_DEFINED__
typedef interface IHTMLElementCollection2 IHTMLElementCollection2;
#endif     /* __IHTMLElementCollection2_FWD_DEFINED__ */

#ifndef __IHTMLElementCollection_INTERFACE_DEFINED__

#define __IHTMLElementCollection_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLElementCollection;


MIDL_INTERFACE("3050f21f-98b5-11cf-bb82-00aa00bdce0b")
IHTMLElementCollection : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE toString(
            /* [out] */ BSTR* String) = 0;

};

#endif     /* __IHTMLElementCollection_INTERFACE_DEFINED__ */


#ifndef __IHTMLElementCollection2_INTERFACE_DEFINED__

#define __IHTMLElementCollection2_INTERFACE_DEFINED__

EXTERN_C const IID IID_IHTMLElementCollection2;


MIDL_INTERFACE("3050f5ee-98b5-11cf-bb82-00aa00bdce0b")
IHTMLElementCollection2 : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE urns(
            /* [in] */ VARIANT urn,/* [out] */ IDispatch** pdisp) = 0;

};

#endif     /* __IHTMLElementCollection2_INTERFACE_DEFINED__ */



EXTERN_C const GUID GUID_HTMLElementCollection;



EXTERN_C const GUID DIID_DispHTMLElementCollection;


#ifndef _CElementCollection_PROPDESCS_
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementCollectiontoString;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementCollectionlength;
EXTERN_C const PROPERTYDESC_BASIC_ABSTRACT s_propdescCElementCollection_newEnum;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementCollectiontags;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementCollectionitem;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCElementCollectionurns;

#endif


#endif /*__collect_h__*/

