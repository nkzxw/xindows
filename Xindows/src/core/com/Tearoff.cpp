
#include "stdafx.h"
#include "Tearoff.h"

typedef HRESULT (STDMETHODCALLTYPE *FNQI)(void* pv, REFIID iid, void** ppv);
typedef ULONG   (STDMETHODCALLTYPE *FNAR)(void* pv);

typedef void    (STDMETHODCALLTYPE *PFNVOID)();


HRESULT STDMETHODCALLTYPE PlainQueryInterface(TEAROFF_THUNK* pthunk, REFIID iid, void** ppv)
{
    void* pv;
    const void* apfnVtbl;
    IID const* const* ppIID;

    for(ppIID=pthunk->apIID; *ppIID; ppIID++)
    {
        if(**ppIID == iid)
        {
            *ppv = pthunk;
            pthunk->ulRef += 1;
            return S_OK;
        }
    }

    if(pthunk->dwMask & 1)
    {
        pv = pthunk->pvObject2;
        apfnVtbl = pthunk->apfnVtblObject2;
    }
    else
    {
        pv = pthunk->pvObject1;
        apfnVtbl = pthunk->apfnVtblObject1;
    }

    return ((FNQI)((void**)apfnVtbl)[0])(pv, iid, ppv);
}

ULONG STDMETHODCALLTYPE PlainAddRef(TEAROFF_THUNK* pthunk)
{
    return ++pthunk->ulRef;
}

static void* s_pvCache1 = NULL;
static void* s_pvCache2 = NULL;

ULONG STDMETHODCALLTYPE PlainRelease(TEAROFF_THUNK* pthunk)
{
    DEBUG_ONLY(static long l = 0; l++;) Assert(pthunk->ulRef > 0);

    if(--pthunk->ulRef == 0)
    {
        if(pthunk->pvObject1 && (pthunk->dwMask&4)==0)
        {
            ((FNAR)((void**)pthunk->apfnVtblObject1)[2])(pthunk->pvObject1);
        }
        if(pthunk->pvObject2)
        {
            ((FNAR)((void**)pthunk->apfnVtblObject2)[2])(pthunk->pvObject2);
        }

        pthunk = (TEAROFF_THUNK*)InterlockedExchangePointer(&s_pvCache1, pthunk);

        if(pthunk)
        {
            pthunk = (TEAROFF_THUNK*)InterlockedExchangePointer(&s_pvCache2, pthunk);

            if(pthunk)
            {
                MemFree(pthunk);
            }
        }

        return 0;
    }

    return pthunk->ulRef;
}

#define THUNK_IMPLEMENT_COMPARE(n)          \
void __declspec(naked) STDMETHODCALLTYPE TearoffThunk##n() \
{										    \
    /* this = thisArg                   */  \
    __asm mov eax, [esp+4]				    \
    /* save original this               */  \
    __asm push eax						    \
    /* if this->_dwMask & 1 << n        */  \
    __asm test dword ptr [eax+28], 1<<n	    \
    /* jmp if < to Else */				    \
    __asm je $+9						    \
    /* increment offset to pvObject2    */  \
    __asm add eax, 8					    \
    /* increment offset to pvObject1    */  \
    __asm add eax, 12					    \
    /* pvObject = this->_pvObject       */  \
    __asm mov ecx, [eax]				    \
    /* thisArg = pvObject2              */  \
    __asm mov [esp+8], ecx				    \
    /* apfnObject = this->_apfnVtblObj  */  \
    __asm mov ecx, [eax+4]				    \
    /* pfn = apfnObject[n]              */  \
    __asm mov ecx, [ecx+(4*n)]			    \
    /* set eax back to the tearoff      */  \
    __asm pop eax						    \
    /* remember vtbl index of method    */  \
    __asm mov dword ptr [eax+32], n		    \
    /* jump....                         */  \
    __asm jmp ecx						    \
}

#define THUNK_IMPLEMENT_SIMPLE(n)           \
void __declspec(naked) STDMETHODCALLTYPE TearoffThunk##n() \
{										    \
    /* this = thisArg                   */  \
    __asm mov eax, [esp+4]				    \
    /* remember vtbl index of method    */  \
    __asm mov dword ptr [eax+32], n		    \
    /* pvObject = this->_pvObject       */  \
    __asm mov ecx, [eax+12]				    \
    /* thisArg = pvObject               */  \
    __asm mov [esp+4], ecx				    \
    /* apfnObject = this->_apfnVtblObj  */  \
    __asm mov ecx, [eax+16]				    \
    /* pfn = apfnObject[n]              */  \
    __asm mov ecx, [ecx+(4*n)]			    \
    /* jump....                         */  \
    __asm jmp ecx						    \
}

THUNK_ARRAY_3_TO_15(IMPLEMENT_COMPARE);
THUNK_ARRAY_16_AND_UP(IMPLEMENT_SIMPLE);

#define THUNK_EXTERN(n)     extern void STDMETHODCALLTYPE TearoffThunk##n();

THUNK_ARRAY_3_TO_15(EXTERN);
THUNK_ARRAY_16_AND_UP(EXTERN);

#define THUNK_ADDRESS(n)    &TearoffThunk##n,

static void (STDMETHODCALLTYPE *s_apfnPlainTearoffVtable[])() =
{
    PFNVOID(&PlainQueryInterface),
    PFNVOID(&PlainAddRef),
    PFNVOID(&PlainRelease),
    THUNK_ARRAY_3_TO_15(ADDRESS)
    THUNK_ARRAY_16_AND_UP(ADDRESS)
};

//+------------------------------------------------------------------------
//
//  Function:   CreateTearOffThunk
//
//  Synopsis:   Create a tearoff interface thunk. The returned object
//              must be AddRef'ed by the caller.
//
//  Arguments:  pvObject    Delegate to this object using...
//              apfnObject    ...this array of pointers to member functions.
//              pUnkOuter   Delegate IUnknown methods to this object.
//              ppvThunk    The returned thunk.
//              pvObject2   Delegate to this object instead...
//              apfnObject2   ... this array of pointers to functions...
//              dwMask        ... when the index of the method call is
//                            marked in this mask.
//
//  Notes:      The basic implementation here consists of a thunk with
//              a pointer to two different objects.  If the second object
//              is NULL, it is assumed to be the first object.  This
//              is the logic of the thunks:
//
//                  i is the index of the method that is called.
//
//                  if (i < 16)
//                  {
//                      if (dwMask & 2^i)
//                      {
//                          Delegate to pvObject2 using apfnObject2
//                      }
//                  }
//                  Delegate to pvObject1 using apfnObject1
//
//-------------------------------------------------------------------------
HRESULT CreateTearOffThunk(void* pvObject1, const void* apfn1, IUnknown* pUnkOuter,
                           void** ppvThunk, void* pvObject2, void* apfn2, DWORD dwMask,
                           const IID* const* apIID, void* appropdescsInVtblOrder)
{
    TEAROFF_THUNK* pthunk;

    Assert(ppvThunk);
    Assert(apfn1);
    Assert((!pvObject2&&!apfn2) || (pvObject2&&apfn2));
    Assert(!(dwMask&0xFFFF0000) && "Only 16 overrides allowed");
    Assert(!dwMask || (dwMask&&pvObject2));
    Assert(!pUnkOuter || (dwMask==0 && pvObject2==0));
    Assert((dwMask&(2|4))==0 || (dwMask&(2|4))==(2|4));

    if(pUnkOuter)
    {
        pvObject2 = pUnkOuter;
        apfn2 = *(void**)pUnkOuter;
        dwMask = 1;
    }

    pthunk = (TEAROFF_THUNK*)InterlockedExchangePointer(&s_pvCache1, NULL);

    if(!pthunk)
    {
        pthunk = (TEAROFF_THUNK*)InterlockedExchangePointer(&s_pvCache2, NULL);

        if(!pthunk)
        {
            pthunk = (TEAROFF_THUNK*)MemAlloc(sizeof(TEAROFF_THUNK));
            MemSetName((pthunk, "Tear-Off Thunk - owner=%08x", pvObject1));
        }
    }

    if(!pthunk)
    {
        *ppvThunk = NULL;
        RRETURN(E_OUTOFMEMORY);
    }

    pthunk->papfnVtblThis = s_apfnPlainTearoffVtable;
    pthunk->ulRef = 0;
    pthunk->pvObject1 = pvObject1;
    pthunk->apfnVtblObject1 = apfn1;
    pthunk->pvObject2 = pvObject2;
    pthunk->apfnVtblObject2 = apfn2;
    pthunk->dwMask = dwMask;
    pthunk->apIID = apIID ? apIID : (const IID* const*)&_afxGlobalData._Zero;
    pthunk->apVtblPropDesc = appropdescsInVtblOrder;

    if(pvObject1 && (dwMask&2)==0)
    {
        ((FNAR)((void**)apfn1)[1])(pvObject1);
    }
    if(pvObject2)
    {
        ((FNAR)((void**)apfn2)[1])(pvObject2);
    }
    *ppvThunk = pthunk;

    return S_OK;
}

// Short argument list version saves space in calling functions.
HRESULT CreateTearOffThunk(void* pvObject1, const void* apfn1, IUnknown* pUnkOuter,
                           void** ppvThunk, void* appropdescsInVtblOrder)
{
    return CreateTearOffThunk(pvObject1, 
        apfn1, 
        pUnkOuter, 
        ppvThunk, 
        NULL, 
        NULL, 
        0, 
        NULL,
        appropdescsInVtblOrder);
}

HRESULT InstallTearOffObject(void* pvthunk, void* pvObject, void* apfn, DWORD dwMask)
{
    TEAROFF_THUNK* pthunk = (TEAROFF_THUNK*)pvthunk;

    Assert(pthunk);
    Assert(!pthunk->pvObject2);
    Assert(!pthunk->apfnVtblObject2);
    Assert(!pthunk->dwMask);

    pthunk->pvObject2 = pvObject;
    pthunk->apfnVtblObject2 = apfn;
    pthunk->dwMask = dwMask;

    if(pvObject)
    {
        ((FNAR)((void**)apfn)[1])(pvObject);
    }

    return S_OK;
}

void DeinitTearOffCache()
{
    MemFree(s_pvCache1);
    MemFree(s_pvCache2);
}