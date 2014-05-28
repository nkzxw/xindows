
#ifndef __XINDOWS_CORE_COM_TEAROFF_H__
#define __XINDOWS_CORE_COM_TEAROFF_H__

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFNTEAROFF)(void);

typedef union _VTABLE_ENTRY
{
    PFNTEAROFF pfn;
    void* pvfn;
} VTABLE_ENTRY;

#define FIRST_VTABLE_OFFSET     0

#define VTBL_PFN(pVtbl)         (pVtbl->pvfn)
#define VTBL_THIS(pVtbl, pThis) (pThis)

#define DECLARE_CLASS_TYPES(klass, supa) \
    typedef supa super; \
    typedef klass thisclass; \
    typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFNTEAROFF)(void);

#define DECLARE_TEAROFF_METHOD(fn, FN, args) \
    STDMETHOD(fn)args

#define DECLARE_TEAROFF_METHOD_(ret, fn, FN, args) \
    STDMETHOD_(ret, fn)args

#define DECLARE_TEAROFF_TABLE(itf) \
    static HRESULT (STDMETHODCALLTYPE CVoid::*const s_apfn##itf[])(void);

#define DECLARE_TEAROFF_TABLE_PROPDESC(itf) \
    static HRESULT (STDMETHODCALLTYPE CVoid::*const s_apfnpd##itf[])(void);

#define DECLARE_TEAROFF_TABLE_NAMED(apfname) \
    static HRESULT (STDMETHODCALLTYPE CVoid::*apfname[])(void);

#define NV_DECLARE_TEAROFF_METHOD(fn, FN, args) \
    HRESULT STDMETHODCALLTYPE fn args

#define NV_DECLARE_TEAROFF_METHOD_(ret, fn, FN, args) \
    ret STDMETHODCALLTYPE fn args

#define BEGIN_TEAROFF_TABLE(klass, itf) \
    HRESULT (STDMETHODCALLTYPE  CVoid::*const klass::s_apfn##itf[])(void) = { \
    TEAROFF_METHOD(klass, &PrivateQueryInterface, privatequeryinterface, (REFIID, void**)) \
    TEAROFF_METHOD_(klass, &PrivateAddRef, privateaddref, ULONG, ()) \
    TEAROFF_METHOD_(klass, &PrivateRelease, privaterelease, ULONG, ()) \

#define BEGIN_TEAROFF_TABLE_PROPDESC(klass, itf) \
    HRESULT (STDMETHODCALLTYPE  CVoid::*const klass::s_apfnpd##itf[])(void) = { \
    TEAROFF_METHOD(klass, &PrivateQueryInterface, privatequeryinterface, (REFIID, void**)) \
    TEAROFF_METHOD_(klass, &PrivateAddRef, privateaddref, ULONG, ()) \
    TEAROFF_METHOD_(klass, &PrivateRelease, privaterelease, ULONG, ()) \

#define BEGIN_TEAROFF_TABLE_(klass, itf) \
    HRESULT (STDMETHODCALLTYPE  CVoid::*const klass::s_apfn##itf[])(void) = { \
    TEAROFF_METHOD(klass, &QueryInterface, queryinterface, (REFIID, void**)) \
    TEAROFF_METHOD_(klass, &AddRef, addref, ULONG, ()) \
    TEAROFF_METHOD_(klass, &Release, release, ULONG, ()) \

#define BEGIN_TEAROFF_TABLE_SUB_(klass, subklass, itf) \
    HRESULT (STDMETHODCALLTYPE  CVoid::*const klass::subklass::s_apfn##itf[])(void) = { \
    TEAROFF_METHOD (klass::subklass, &klass::subklass::QueryInterface, klass::subklass::queryinterface, (REFIID, void**)) \
    TEAROFF_METHOD_(klass::subklass, &klass::subklass::AddRef, klass::subklass::addref, ULONG, ()) \
    TEAROFF_METHOD_(klass::subklass, &klass::subklass::Release, klass::subklass::release, ULONG, ()) \

#define BEGIN_TEAROFF_TABLE_NAMED(klass, name) \
    HRESULT (STDMETHODCALLTYPE  CVoid::*klass::name[])(void) = { \

#define END_TEAROFF_TABLE() \
    };

#define _TEAROFF_METHOD(fn, FN, args) \
    (PFNTEAROFF)(HRESULT (STDMETHODCALLTYPE CVoid::*)args)fn,

#define _TEAROFF_METHOD_(fn, FN, ret, args) \
    (PFNTEAROFF)(ret (STDMETHODCALLTYPE CVoid::*)args)fn,

#define TEAROFF_METHOD(klass, fn, FN, args) \
    (PFNTEAROFF)(HRESULT (STDMETHODCALLTYPE CVoid::*)args)fn,

#define TEAROFF_METHOD_(klass, fn, FN, ret, args) \
    (PFNTEAROFF)(ret (STDMETHODCALLTYPE CVoid::*)args)fn,

#define TEAROFF_METHOD_SUB(klass, subklass, fn, FN, args) \
    TEAROFF_METHOD(klass::subklass, &klass::subklass::fn, &klass::subklass::FN, args)

#define TEAROFF_METHOD_NULL \
    (PFNTEAROFF)NULL,

#define cat(a,b)                    a##b

struct TEAROFF_THUNK
{
    void*			    papfnVtblThis;	    // Thunk's vtable
    ULONG			    ulRef;			    // Reference count for this thunk.
    IID const* const*   apIID;			    // Short circuit QI using these IIDs.
    void*			    pvObject1;		    // Delegate other methods to this object using...
    const void*		    apfnVtblObject1;    // ...this array of pointers to member functions.
    void*			    pvObject2;		    // Delegate methods to this object using...
    void*			    apfnVtblObject2;    // ...this array of pointers to member functions...
    DWORD			    dwMask;			    // ...the index of the method is set in the mask.
    DWORD			    n;				    // index of method into vtbl
    void*               apVtblPropDesc;     // array of propdescs in Vtbl order
};

#define GET_TEAROFF_THUNK           \
    TEAROFF_THUNK* pthunk;          \
{                                   \
    __asm mov pthunk, eax           \
}

// macro for accessing propdesc from tearoff thunks in interface methods
#define FIRST_INTERFACE_METHOD  3

#define GET_THUNK_PROPDESC          \
    GET_TEAROFF_THUNK               \
    const PROPERTYDESC* pPropDesc;  \
    Assert(pthunk);                 \
    Assert(pthunk->apVtblPropDesc); \
    pPropDesc = ((const PROPERTYDESC* const*)pthunk->apVtblPropDesc)[pthunk->n-FIRST_INTERFACE_METHOD]; \
    Assert(pPropDesc);              \

#define THUNK_ARRAY_3_TO_15(x) \
    cat(THUNK_,x)(3)   cat(THUNK_,x)(4)   cat(THUNK_,x)(5)   cat(THUNK_,x)(6)   cat(THUNK_,x)(7)   cat(THUNK_,x)(8)   cat(THUNK_,x)(9)   cat(THUNK_,x)(10)  cat(THUNK_,x)(11)  cat(THUNK_,x)(12)  cat(THUNK_,x)(13)  \
    cat(THUNK_,x)(14)  cat(THUNK_,x)(15)

#define THUNK_ARRAY_16_AND_UP(x) \
    cat(THUNK_,x)(16)  cat(THUNK_,x)(17)  cat(THUNK_,x)(18)  cat(THUNK_,x)(19)  cat(THUNK_,x)(20)  cat(THUNK_,x)(21)  cat(THUNK_,x)(22)  cat(THUNK_,x)(23)  cat(THUNK_,x)(24)  \
    cat(THUNK_,x)(25)  cat(THUNK_,x)(26)  cat(THUNK_,x)(27)  cat(THUNK_,x)(28)  cat(THUNK_,x)(29)  cat(THUNK_,x)(30)  cat(THUNK_,x)(31)  cat(THUNK_,x)(32)  cat(THUNK_,x)(33)  cat(THUNK_,x)(34)  cat(THUNK_,x)(35)  \
    cat(THUNK_,x)(36)  cat(THUNK_,x)(37)  cat(THUNK_,x)(38)  cat(THUNK_,x)(39)  cat(THUNK_,x)(40)  cat(THUNK_,x)(41)  cat(THUNK_,x)(42)  cat(THUNK_,x)(43)  cat(THUNK_,x)(44)  cat(THUNK_,x)(45)  cat(THUNK_,x)(46)  \
    cat(THUNK_,x)(47)  cat(THUNK_,x)(48)  cat(THUNK_,x)(49)  cat(THUNK_,x)(50)  cat(THUNK_,x)(51)  cat(THUNK_,x)(52)  cat(THUNK_,x)(53)  cat(THUNK_,x)(54)  cat(THUNK_,x)(55)  cat(THUNK_,x)(56)  cat(THUNK_,x)(57)  \
    cat(THUNK_,x)(58)  cat(THUNK_,x)(59)  cat(THUNK_,x)(60)  cat(THUNK_,x)(61)  cat(THUNK_,x)(62)  cat(THUNK_,x)(63)  cat(THUNK_,x)(64)  cat(THUNK_,x)(65)  cat(THUNK_,x)(66)  cat(THUNK_,x)(67)  cat(THUNK_,x)(68)  \
    cat(THUNK_,x)(69)  cat(THUNK_,x)(70)  cat(THUNK_,x)(71)  cat(THUNK_,x)(72)  cat(THUNK_,x)(73)  cat(THUNK_,x)(74)  cat(THUNK_,x)(75)  cat(THUNK_,x)(76)  cat(THUNK_,x)(77)  cat(THUNK_,x)(78)  cat(THUNK_,x)(79)  \
    cat(THUNK_,x)(80)  cat(THUNK_,x)(81)  cat(THUNK_,x)(82)  cat(THUNK_,x)(83)  cat(THUNK_,x)(84)  cat(THUNK_,x)(85)  cat(THUNK_,x)(86)  cat(THUNK_,x)(87)  cat(THUNK_,x)(88)  cat(THUNK_,x)(89)  cat(THUNK_,x)(90)  \
    cat(THUNK_,x)(91)  cat(THUNK_,x)(92)  cat(THUNK_,x)(93)  cat(THUNK_,x)(94)  cat(THUNK_,x)(95)  cat(THUNK_,x)(96)  cat(THUNK_,x)(97)  cat(THUNK_,x)(98)  cat(THUNK_,x)(99)  cat(THUNK_,x)(100) cat(THUNK_,x)(101) \
    cat(THUNK_,x)(102) cat(THUNK_,x)(103) cat(THUNK_,x)(104) cat(THUNK_,x)(105) cat(THUNK_,x)(106) cat(THUNK_,x)(107) cat(THUNK_,x)(108) cat(THUNK_,x)(109) cat(THUNK_,x)(110) cat(THUNK_,x)(111) cat(THUNK_,x)(112) \
    cat(THUNK_,x)(113) cat(THUNK_,x)(114) cat(THUNK_,x)(115) cat(THUNK_,x)(116) cat(THUNK_,x)(117) cat(THUNK_,x)(118) cat(THUNK_,x)(119) cat(THUNK_,x)(120) cat(THUNK_,x)(121) cat(THUNK_,x)(122) cat(THUNK_,x)(123) \
    cat(THUNK_,x)(124) cat(THUNK_,x)(125) cat(THUNK_,x)(126) cat(THUNK_,x)(127) cat(THUNK_,x)(128) cat(THUNK_,x)(129) cat(THUNK_,x)(130) cat(THUNK_,x)(131) cat(THUNK_,x)(132) cat(THUNK_,x)(133) cat(THUNK_,x)(134) \
    cat(THUNK_,x)(135) cat(THUNK_,x)(136) cat(THUNK_,x)(137) cat(THUNK_,x)(138) cat(THUNK_,x)(139) cat(THUNK_,x)(140) cat(THUNK_,x)(141) cat(THUNK_,x)(142) cat(THUNK_,x)(143) cat(THUNK_,x)(144) cat(THUNK_,x)(145) \
    cat(THUNK_,x)(146) cat(THUNK_,x)(147) cat(THUNK_,x)(148) cat(THUNK_,x)(149) cat(THUNK_,x)(150) cat(THUNK_,x)(151) cat(THUNK_,x)(152) cat(THUNK_,x)(153) cat(THUNK_,x)(154) cat(THUNK_,x)(155) cat(THUNK_,x)(156) \
    cat(THUNK_,x)(157) cat(THUNK_,x)(158) cat(THUNK_,x)(159) cat(THUNK_,x)(160) cat(THUNK_,x)(161) cat(THUNK_,x)(162) cat(THUNK_,x)(163) cat(THUNK_,x)(164) cat(THUNK_,x)(165) cat(THUNK_,x)(166) cat(THUNK_,x)(167) \
    cat(THUNK_,x)(168) cat(THUNK_,x)(169) cat(THUNK_,x)(170) cat(THUNK_,x)(171) cat(THUNK_,x)(172) cat(THUNK_,x)(173) cat(THUNK_,x)(174) cat(THUNK_,x)(175) cat(THUNK_,x)(176) cat(THUNK_,x)(177) cat(THUNK_,x)(178) \
    cat(THUNK_,x)(179) cat(THUNK_,x)(180) cat(THUNK_,x)(181) cat(THUNK_,x)(182) cat(THUNK_,x)(183) cat(THUNK_,x)(184) cat(THUNK_,x)(185) cat(THUNK_,x)(186) cat(THUNK_,x)(187) cat(THUNK_,x)(188) cat(THUNK_,x)(189) \
    cat(THUNK_,x)(190) cat(THUNK_,x)(191) cat(THUNK_,x)(192) cat(THUNK_,x)(193) cat(THUNK_,x)(194) cat(THUNK_,x)(195) cat(THUNK_,x)(196) cat(THUNK_,x)(197) cat(THUNK_,x)(198) cat(THUNK_,x)(199)


HRESULT CreateTearOffThunk(void* pvObject1, const void* apfn1, IUnknown* pUnkOuter, void** ppvThunk, void* appropdescsInVtblOrder=NULL);

// Indices to the IUnknown methods.
#define METHOD_QI           0
#define METHOD_ADDREF       1
#define METHOD_RELEASE      2

#define METHOD_MASK(index)  (1<<(index))

// Well known method dwMask values
#define QI_MASK             METHOD_MASK(METHOD_QI)
#define ADDREF_MASK         METHOD_MASK(METHOD_ADDREF)
#define RELEASE_MASK        METHOD_MASK(METHOD_RELEASE)

HRESULT CreateTearOffThunk(void* pvObject1, const void* apfn1, IUnknown* pUnkOuter,
                           void** ppvThunk, void* pvObject2, void* apfn2, DWORD dwMask,
                           const IID* const* apIID, void* appropdescsInVtblOrder=NULL);

// Installs pvObject2 after the tearoff is created
HRESULT InstallTearOffObject(void* pvthunk, void* pvObject, void* apfn, DWORD dwMask);

#define Data1_IAdviseSink                  0x0000010f
#define Data1_IAdviseSink2                 0x00000125
#define Data1_IAdviseSinkEx                0x3af24290
#define Data1_IBindCtx                     0x0000000e
#define Data1_IBindHost                    0xfc4801a1
#define Data1_ICDataDoc                    0xF413E4C0
#define Data1_IClassFactory                0x00000001
#define Data1_IClassFactory2               0xb196b28f
#define Data1_IConnectionPoint             0xb196b286
#define Data1_IConnectionPointContainer    0xb196b284
#define Data1_ICreateErrorInfo             0x22f03340
#define Data1_ICreateTypeInfo              0x00020405
#define Data1_ICreateTypeLib               0x00020406
#define Data1_IDataAdviseHolder            0x00000110
#define Data1_IDataObject                  0x0000010e
#define Data1_DataSource                   0x7c0ffab3
#define Data1_DataSourceListener           0x7c0ffab2
#define Data1_IDATASRCListener             0x3050f380
#define Data1_IDispatch                    0x00020400
#define Data1_IDispObserver                0x3050f442
#define Data1_IDropSource                  0x00000121
#define Data1_IDropTarget                  0x00000122
#define Data1_IEnumCallback                0x00000108
#define Data1_IEnumConnectionPoints        0xb196b285
#define Data1_IEnumConnections             0xb196b287
#define Data1_IEnumFORMATETC               0x00000103
#define Data1_IEnumGeneric                 0x00000106
#define Data1_IEnumHolder                  0x00000107
#define Data1_IEnumMoniker                 0x00000102
#define Data1_IEnumOLEVERB                 0x00000104
#define Data1_IEnumSTATDATA                0x00000105
#define Data1_IEnumSTATSTG                 0x0000000d
#define Data1_IEnumString                  0x00000101
#define Data1_IEnumOleUndoUnits            0xb3e7c340
#define Data1_IEnumUnknown                 0x00000100
#define Data1_IEnumVARIANT                 0x00020404
#define Data1_IErrorInfo                   0x1cf2b120
#define Data1_IExternalConnection          0x00000019
#define Data1_IFont                        0xbef6e002
#define Data1_IFontDisp                    0xbef6e003
#define Data1_IForm                        0x04598fc8
#define Data1_IFormExpert                  0x04598fc5
#define Data1_IGangConnectWithDefault      0x6d5140c0
#define Data1_IHlinkTarget                 0x79eac9c4
#define Data1_IInternalMoniker             0x00000011
#define Data1_ILockBytes                   0x0000000a
#define Data1_IMalloc                      0x00000002
#define Data1_IMessageFilter               0x00000016
#define Data1_IMoniker                     0x0000000f
#define Data1_IOleCommandTarget            0xb722bccb
#define Data1_IOleDocument                 0xb722bcc5
#define Data1_IOleDocumentView             0xb722bcc6
#define Data1_IOleAdviseHolder             0x00000111
#define Data1_IOleCache                    0x0000011e
#define Data1_IOleCache2                   0x00000128
#define Data1_IOleCacheControl             0x00000129
#define Data1_IOleClientSite               0x00000118
#define Data1_IOleParentUndoUnit           0xa1faf330
#define Data1_IOleContainer                0x0000011b
#define Data1_IOleControl                  0xb196b288
#define Data1_IOleControlSite              0xb196b289
#define Data1_IOleInPlaceActiveObject      0x00000117
#define Data1_IOleInPlaceFrame             0x00000116
#define Data1_IOleInPlaceObject            0x00000113
#define Data1_IOleInPlaceObjectWindowless  0x1c2056cc
#define Data1_IOleInPlaceSite              0x00000119
#define Data1_IOleInPlaceSiteEx            0x9c2cad80
#define Data1_IOleInPlaceSiteWindowless    0x922eada0
#define Data1_IOleInPlaceUIWindow          0x00000115
#define Data1_IOleItemContainer            0x0000011c
#define Data1_IOleLink                     0x0000011d
#define Data1_IOleManager                  0x0000011f
#define Data1_IOleObject                   0x00000112
#define Data1_IOlePresObj                  0x00000120
#define Data1_IOlePropertyFrame            0xb83bb801
#define Data1_IOleStandardTool             0xd97877c4
#define Data1_IOleUndoUnit                 0x894ad3b0
#define Data1_IOleUndoManager              0xd001f200
#define Data1_IOleWindow                   0x00000114
#define Data1_IPSFactory                   0x00000009
#define Data1_IPSFactoryBuffer             0xd5f569d0
#define Data1_IParseDisplayName            0x0000011a
#define Data1_IPersist                     0x0000010c
#define Data1_IPersistFile                 0x0000010b
#define Data1_IPersistHistory              0x91A565C1
#define Data1_IPersistMoniker              0x79eac9c9
#define Data1_IPersistStorage              0x0000010a
#define Data1_IPersistStream               0x00000109
#define Data1_IPersistStreamInit           0x7fd52380
#define Data1_IPersistPropertyBag          0x37D84F60
#define Data1_IPersistHTML                 0x049948d1
#define Data1_IPicture                     0x7bf80980
#define Data1_IPictureDisp                 0x7bf80981
#define Data1_IPointerInactive             0x55980ba0
#define Data1_IPropertyNotifySink          0x9bfbbc02
#define Data1_IPropertyPage                0xb196b28d
#define Data1_IPropertyPage2               0x01e44665
#define Data1_IPropertyPage3               0xb83bb803
#define Data1_IPropertyPageInPlace         0xb83bb802
#define Data1_IPropertyPageSite            0xb196b28c
#define Data1_IPropertyPageSite2           0xb83bb804
#define Data1_IProvideDynamicClassInfo     0x6d5140d1
#define Data1_IQuickActivate               0xcf51ed10
#define Data1_IRootStorage                 0x00000012
#define Data1_IRunnableObject              0x00000126
#define Data1_IRunningObjectTable          0x00000010
#define Data1_ISelectionContainer          0x6d5140c6
#define Data1_IServiceProvider             0x6d5140c1
#define Data1_ISimpleFrameSite             0x742b0e01
#define Data1_ISpecifyPropertyPages        0xb196b28b
#define Data1_IStdMarshalInfo              0x00000018
#define Data1_IStorage                     0x0000000b
#define Data1_IStream                      0x0000000c
#define Data1_ITargetContainer             0x7847EC01
#define Data1_ITargetEmbedding             0x548793c0
#define Data1_ITargetFrame                 0xd5f78c80
#define Data1_ITargetFrame2                0x3abac181
#define Data1_ITargetNotify                0x863a99a0
#define Data1_ITypeComp                    0x00020403
#define Data1_ITypeInfo                    0x00020401
#define Data1_ITypeLib                     0x00020402
#define Data1_IUnknown                     0x00000000
#define Data1_IViewObject                  0x0000010d
#define Data1_IViewObject2                 0x00000127
#define Data1_IViewObjectEx                0x3af24292
#define Data1_IWeakRef                     0x0000001a
#define Data1_ICategorizeProperties        0x4d07fc10
#define Data1_IAccessible                  0x618736e0
#define Data1_IHTMLFrameBase               0x3050f311
#define Data1_IHTMLTextSite                0x1ff6aa71
#define Data1_IHTMLImageSite               0x49A4B4C0
#define Data1_IHTMLBodyElement             0x3050f1d8
#define Data1_IHTMLFontElement             0x3050f1d9
#define Data1_IHTMLParaElement             0x3050f1f5
#define Data1_IHTMLDivElement              0x3050f200
#define Data1_IHTMLBaseFontElement         0x3050f202
#define Data1_IGetUniqueID                 0xaf8665b0
#define Data1_IBoundObjectSite             0x9BFBBC01
#define Data1_IBitmapSurface               0x3050f2ef 
#define Data1_IGdiSurface                  0x3050f2f0 
#define Data1_IDDSurface                   0x3050f2f1 
#define Data1_IBitmapSurfaceFactory        0x3050f2f2 
#define Data1_IGdiSurfaceFactory           0x3050f2f3 
#define Data1_IDDSurfaceFactory            0x3050f2f4 
#define Data1_IRGBColorTable               0x3050f2f5 
#define Data1_IHTMLDocument                0x626FC520
#define Data1_IHTMLDocument2               0x332c4425
#define Data1_IHTMLDocument3               0x3050f485
#define Data1_IMarqueeInfo                 0x0bdc6ae0
#define Data1_IObjectSafety                0xcb5bdc81
#define Data1_IShellPropSheetExt           0x000214E9
#define Data1_IHTMLScreen                  0x3050f35c
#define Data1_IHTMLViewFilter              0x3050f2f1
#define Data1_IHTMLViewFilterSite          0x3050f2f4
#define Data1_ITimerService                0x3050f35f
#define Data1_ITimer                       0x3050f360
#define Data1_ITimerSink                   0x3050f361
#define Data1_IHTMLLocation                0x163BB1E0
#define Data1_IHTMLElement                 0x3050f1ff
#define Data1_IHTMLElement2                0x3050f434
#define Data1_IInternetHostSecurityManager 0x3af280b6
#define Data1_IHTMLTxtRange                0x3050F220
#define Data1_IHTMLTextRangeMetrics        0x3050F40B
#define Data1_IHTMLTextRangeMetrics2       0x3050f4a6
#define Data1_ICustomDoc                   0x3050f3f0
#define Data1_IObjectIdentity              0xCA04B7E6
#define Data1_IBindStatusCallback          0x79eac9c1
#define Data1_IWebBridge                   0xAE24FDAD
#define Data1_IIdentityBehavior            0x3050f60c
#define Data1_IIdentityBehaviorFactory     0x3050f60d
#define Data1_IHTCAttachBehavior           0x3050f5f4
#define Data1_IProfferService              0xcb728b20
#define Data1_IPersistPropertyBag2         0x22F55881
#define Data1_IHTMLStyle                   0x3050f25e
#define Data1_IHTMLStyle2                  0x3050f4a2
#define Data1_IObjectWithSite              0xFC4801A3
#define Data1_IMarkupPointer               0x3050f49f
#define Data1_IMarkupServices              0x3050f4a0
#define Data1_IMarkupContainer             0x3050f5f9
#define Data1_IHTMLViewServices            0x3050f603
#define Data1_IScriptletHandler            0xa001a870
#define Data1_IScriptletHandlerConstructor 0xa3d52a50
#define Data1_IClassFactoryEx              0x342d1ea0
#define Data1_IHTMLObjectElement           0x3050f24f
#define Data1_IHTMLObjectElement2          0x3050f4cd
#define Data1_ISegmentList                 0x3050f605
#define Data1_ISelectionRenderingServices  0x3050f606
#define Data1_IHTMLFormElement2            0x3050f4f6
#define Data1_IHTMLCaret                   0x3050f604
#define Data1_IDispClient                  0x3050f437
#define Data1_IHTMLElementCollection2      0x3050f5ee
#define Data1_IHTMLAreasCollection2        0x3050f5ec
#define Data1_IHTMLSelectElement2          0x3050f5ed


#define QI_INHERITS(pObj, itf)              \
case Data1_##itf:                           \
    if(iid == IID_##itf)                    \
    {                                       \
        *ppv = (itf*)pObj;                  \
    }                                       \
    break;                                  \

#define QI_INHERITS2(pObj, itf, itfDerived) \
case Data1_##itf:                           \
    if(iid == IID_##itf)                    \
    {                                       \
        *ppv = (itfDerived*)pObj;           \
    }                                       \
    break;                                  \

#define QI_TEAROFF(pObj, itf, pUnkOuter)    \
case Data1_##itf:                           \
{									        \
    HRESULT hr = CreateTearOffThunk(	    \
    pObj,								    \
    (void*)pObj->s_apfn##itf,			    \
    pUnkOuter,							    \
    ppv);								    \
}                                           \
break;                                      \

#define QI_TEAROFF2(pObj, itf, itfDerived, pUnkOuter) \
case Data1_##itf:                           \
{										    \
    HRESULT hr = CreateTearOffThunk(        \
    pObj,								    \
    (void*)pObj->s_apfn##itfDerived,	    \
    pUnkOuter,							    \
    ppv);								    \
}                                           \
break;                                      \

#define QI_HTML_TEAROFF(pObj, itf, pUnkOuter) \
case Data1_##itf:                           \
if(iid == IID_##itf)                        \
{                                           \
    HRESULT hr = CreateTearOffThunk(        \
    pObj,                                   \
    (void*)pObj->s_apfnpd##itf,             \
    pUnkOuter,                              \
    ppv,                                    \
    (void*)s_ppropdescsInVtblOrder##itf);   \
    if(hr)                                  \
        RRETURN(hr);                        \
}                                           \
break;                                      \


// Usuage if IID_TEAROFF(xxxx, xxxx, xxxx)
#define IID_TEAROFF(pObj, itf, pUnkOuter)   \
    (iid == IID_##itf)                      \
    {                                       \
        hr = CreateTearOffThunk(            \
            pObj,                           \
            s_apfn##itf,                    \
            pUnkOuter,                      \
            ppv);                           \
        if(hr) RRETURN(hr);                 \
    }                                       \

#define IID_HTML_TEAROFF(pObj, itf, pUnkOuter)  \
    (iid == IID_##itf)                          \
    {                                           \
        hr = CreateTearOffThunk(                \
        pObj,                                   \
        s_apfnpd##itf,                          \
        pUnkOuter,                              \
        ppv,                                    \
        (void*)s_ppropdescsInVtblOrder##itf);   \
        if (hr)                                 \
        RRETURN(hr);                            \
    }


#define QI_CASE(itf)                                \
    case Data1_##itf:                               \



#define CONTEXTTHUNK_SETCONTEXT             \
{                                           \
    __asm  mov pUnkContext, eax             \
}

#define CONTEXTTHUNK_SETTREENODE            \
    IUnknown* pUnkContext;                  \
    CONTEXTTHUNK_SETCONTEXT                 \
    Assert(pUnkContext);                    \
    pNode = NULL;                           \
    pUnkContext->QueryInterface(CLSID_CTreeNode, (void**)&pNode);

#endif //__XINDOWS_CORE_COM_TEAROFF_H__