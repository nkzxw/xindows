#ifndef __types_h__
#define __types_h__

/* Forward Declarations */

struct ENUMDESC;

/* header files for imported files */
#include "oaidl.h"

/* header files for imported files */
#include "oleidl.h"

/* header files for imported files */
#include "ocidl.h"

#ifndef __IObjectIdentity_FWD_DEFINED__
#define __IObjectIdentity_FWD_DEFINED__
typedef interface IObjectIdentity IObjectIdentity;
#endif     /* __IObjectIdentity_FWD_DEFINED__ */
typedef enum _htmlZOrder
{
    htmlZOrderFront = 0,
    htmlZOrderBack = 1,
    htmlZOrder_Max = 2147483647L
} htmlZOrder;


EXTERN_C const ENUMDESC s_enumdeschtmlZOrder;

typedef enum _htmlClear
{
    htmlClearNotSet = 0,
    htmlClearAll = 1,
    htmlClearLeft = 2,
    htmlClearRight = 3,
    htmlClearBoth = 4,
    htmlClearNone = 5,
    htmlClear_Max = 2147483647L
} htmlClear;


EXTERN_C const ENUMDESC s_enumdeschtmlClear;

typedef enum _htmlControlAlign
{
    htmlControlAlignNotSet = 0,
    htmlControlAlignLeft = 1,
    htmlControlAlignCenter = 2,
    htmlControlAlignRight = 3,
    htmlControlAlignTextTop = 4,
    htmlControlAlignAbsMiddle = 5,
    htmlControlAlignBaseline = 6,
    htmlControlAlignAbsBottom = 7,
    htmlControlAlignBottom = 8,
    htmlControlAlignMiddle = 9,
    htmlControlAlignTop = 10,
    htmlControlAlign_Max = 2147483647L
} htmlControlAlign;


EXTERN_C const ENUMDESC s_enumdeschtmlControlAlign;

typedef enum _htmlBlockAlign
{
    htmlBlockAlignNotSet = 0,
    htmlBlockAlignLeft = 1,
    htmlBlockAlignCenter = 2,
    htmlBlockAlignRight = 3,
    htmlBlockAlignJustify = 4,
    htmlBlockAlign_Max = 2147483647L
} htmlBlockAlign;


EXTERN_C const ENUMDESC s_enumdeschtmlBlockAlign;

typedef enum _htmlReadyState
{
    htmlReadyStateuninitialized = 0,
    htmlReadyStateloading = 1,
    htmlReadyStateloaded = 2,
    htmlReadyStateinteractive = 3,
    htmlReadyStatecomplete = 4,
    htmlReadyState_Max = 2147483647L
} htmlReadyState;


EXTERN_C const ENUMDESC s_enumdeschtmlReadyState;

typedef enum _htmlLoop
{
    htmlLoopLoopInfinite = -1,
    htmlLoop_Max = 2147483647L
} htmlLoop;


EXTERN_C const ENUMDESC s_enumdeschtmlLoop;


#ifndef _CBase_PROPDESCS_
EXTERN_C const PROPERTYDESC_METHOD s_methdescCBasesetAttribute;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCBasegetAttribute;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCBaseremoveAttribute;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCBaseattachEvent;
EXTERN_C const PROPERTYDESC_METHOD s_methdescCBasedetachEvent;

#endif

#ifndef _CFunctionPointer_PROPDESCS_

#endif


#endif /*__types_h__*/

