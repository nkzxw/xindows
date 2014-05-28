/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Jan 03 20:20:12 2009
 */
/* Compiler settings for IDispObserver.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __IDispObserver_h__
#define __IDispObserver_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IDispObserver_FWD_DEFINED__
#define __IDispObserver_FWD_DEFINED__
typedef interface IDispObserver IDispObserver;
#endif 	/* __IDispObserver_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IDispObserver_INTERFACE_DEFINED__
#define __IDispObserver_INTERFACE_DEFINED__

/* interface IDispObserver */
/* [local][unique][uuid][object] */ 


EXTERN_C const IID IID_IDispObserver;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3050f43d-98b5-11cf-bb82-00aa00bdce0b")
    IDispObserver : public IUnknown
    {
    public:
        virtual void STDMETHODCALLTYPE Invalidate( 
            /* [in] */ const RECT __RPC_FAR *prcInvalid,
            /* [in] */ HRGN rgnInvalid,
            /* [in] */ DWORD flags) = 0;
        
        virtual HDC STDMETHODCALLTYPE GetClientDC( 
            /* [in] */ const RECT __RPC_FAR *prc) = 0;
        
        virtual void STDMETHODCALLTYPE ReleaseClientDC( 
            /* [in] */ HDC hdc) = 0;
        
        virtual void STDMETHODCALLTYPE DrawSynchronous( 
            /* [in] */ HRGN hrgn,
            /* [in] */ HDC hdc,
            /* [in] */ const RECT __RPC_FAR *prcClip) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IDispObserverVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IDispObserver __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IDispObserver __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IDispObserver __RPC_FAR * This);
        
        void ( STDMETHODCALLTYPE __RPC_FAR *Invalidate )( 
            IDispObserver __RPC_FAR * This,
            /* [in] */ const RECT __RPC_FAR *prcInvalid,
            /* [in] */ HRGN rgnInvalid,
            /* [in] */ DWORD flags);
        
        HDC ( STDMETHODCALLTYPE __RPC_FAR *GetClientDC )( 
            IDispObserver __RPC_FAR * This,
            /* [in] */ const RECT __RPC_FAR *prc);
        
        void ( STDMETHODCALLTYPE __RPC_FAR *ReleaseClientDC )( 
            IDispObserver __RPC_FAR * This,
            /* [in] */ HDC hdc);
        
        void ( STDMETHODCALLTYPE __RPC_FAR *DrawSynchronous )( 
            IDispObserver __RPC_FAR * This,
            /* [in] */ HRGN hrgn,
            /* [in] */ HDC hdc,
            /* [in] */ const RECT __RPC_FAR *prcClip);
        
        END_INTERFACE
    } IDispObserverVtbl;

    interface IDispObserver
    {
        CONST_VTBL struct IDispObserverVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDispObserver_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDispObserver_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDispObserver_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDispObserver_Invalidate(This,prcInvalid,rgnInvalid,flags)	\
    (This)->lpVtbl -> Invalidate(This,prcInvalid,rgnInvalid,flags)

#define IDispObserver_GetClientDC(This,prc)	\
    (This)->lpVtbl -> GetClientDC(This,prc)

#define IDispObserver_ReleaseClientDC(This,hdc)	\
    (This)->lpVtbl -> ReleaseClientDC(This,hdc)

#define IDispObserver_DrawSynchronous(This,hrgn,hdc,prcClip)	\
    (This)->lpVtbl -> DrawSynchronous(This,hrgn,hdc,prcClip)

#endif /* COBJMACROS */


#endif 	/* C style interface */



void STDMETHODCALLTYPE IDispObserver_Invalidate_Proxy( 
    IDispObserver __RPC_FAR * This,
    /* [in] */ const RECT __RPC_FAR *prcInvalid,
    /* [in] */ HRGN rgnInvalid,
    /* [in] */ DWORD flags);


void __RPC_STUB IDispObserver_Invalidate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HDC STDMETHODCALLTYPE IDispObserver_GetClientDC_Proxy( 
    IDispObserver __RPC_FAR * This,
    /* [in] */ const RECT __RPC_FAR *prc);


void __RPC_STUB IDispObserver_GetClientDC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


void STDMETHODCALLTYPE IDispObserver_ReleaseClientDC_Proxy( 
    IDispObserver __RPC_FAR * This,
    /* [in] */ HDC hdc);


void __RPC_STUB IDispObserver_ReleaseClientDC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


void STDMETHODCALLTYPE IDispObserver_DrawSynchronous_Proxy( 
    IDispObserver __RPC_FAR * This,
    /* [in] */ HRGN hrgn,
    /* [in] */ HDC hdc,
    /* [in] */ const RECT __RPC_FAR *prcClip);


void __RPC_STUB IDispObserver_DrawSynchronous_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IDispObserver_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
