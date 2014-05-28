BOOL DispNonDualDIID(IID iid);
typedef HRESULT (*CustomHandler)(CBase *pBase,
                                 IServiceProvider *pSrvProvider,
                                 IDispatch *pInstance,
                                 WORD wVTblOffset,
                                 PROPERTYDESC_BASIC_ABSTRACT *pDesc,
                                 WORD wFlags,
                                 DISPPARAMS *pdispparams,
                                 VARIANT *pvarResult);


HRESULT Method_void_BSTR_VARIANT_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_BSTR_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_oDoLONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_BSTR_BSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_PropEnum (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_float (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_BSTRp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_PropEnum (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_IUnknownp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_VARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_BSTR_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_BSTRp_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_short (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_IUnknownp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_void (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_IDispatchp_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_IDispatchp_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_VARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_IDispatchp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_BSTRp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_BSTRp_BSTR_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_longp_BSTR_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_VARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_oDolong_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_longp_BSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_longp_BSTR_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_BSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_oDoVARIANTBOOL_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_longp_BSTR_BSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT G_short (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_VARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_BSTR_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_IDispatchp_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_SAFEARRAYPVARIANTP (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_oDoBSTR_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_long_long (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_oDoBSTR_oDolong (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_longp_VARIANTp_long_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_oDoBSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_oDoBSTR_oDoBSTR_oDoBSTR_oDoVARIANTBOOL (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_BSTR_o0oVARIANTp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_void_BSTR_o0oVARIANT_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTp_BSTR_oDoBSTR (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_oDoBSTR_o0oVARIANTp_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_LONG (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT GS_IDispatchp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_IDispatchpp_BSTR_o0oVARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);

HRESULT Method_VARIANTBOOLp_BSTR_VARIANTp (CBase *pBase,
            IServiceProvider *pSrvProvider,
            IDispatch *pInstance,
            WORD wVTblOffset,
            PROPERTYDESC_BASIC_ABSTRACT *pDesc,
            WORD wFlags,
            DISPPARAMS *pdispparams,
            VARIANT *pvarResult);


#define 	IDX_Method_void_BSTR_VARIANT_oDoLONG 	0
#define 	IDX_Method_VARIANTp_BSTR_oDoLONG 	1
#define 	IDX_Method_VARIANTBOOLp_BSTR_oDoLONG 	2
#define 	IDX_Method_void_BSTR_BSTR_oDoBSTR 	3
#define 	IDX_Method_VARIANTp_BSTR 	4
#define 	IDX_Method_VARIANTBOOLp_BSTR 	5
#define 	IDX_Method_VARIANTBOOLp_BSTR_IDispatchp 	6
#define 	IDX_Method_void_BSTR_IDispatchp 	7
#define 	IDX_GS_BSTR 	8
#define 	IDX_GS_PropEnum 	9
#define 	IDX_GS_VARIANT 	10
#define 	IDX_GS_VARIANTBOOL 	11
#define 	IDX_GS_long 	12
#define 	IDX_GS_float 	13
#define 	IDX_Method_BSTRp_void 	14
#define 	IDX_G_PropEnum 	15
#define 	IDX_G_VARIANT 	16
#define 	IDX_G_BSTR 	17
#define 	IDX_G_long 	18
#define 	IDX_G_IUnknownp 	19
#define 	IDX_Method_VARIANTp_VARIANTp 	20
#define 	IDX_G_IDispatchp 	21
#define 	IDX_Method_void_o0oVARIANT 	22
#define 	IDX_Method_VARIANTBOOLp_IDispatchp 	23
#define 	IDX_Method_void_BSTR_BSTR 	24
#define 	IDX_G_VARIANTBOOL 	25
#define 	IDX_Method_void_void 	26
#define 	IDX_Method_void_oDoVARIANTBOOL 	27
#define 	IDX_Method_BSTRp_long_long 	28
#define 	IDX_Method_IDispatchpp_void 	29
#define 	IDX_GS_short 	30
#define 	IDX_Method_void_IUnknownp 	31
#define 	IDX_Method_VARIANTBOOLp_void 	32
#define 	IDX_Method_IDispatchpp_IDispatchp_o0oVARIANT 	33
#define 	IDX_Method_IDispatchpp_IDispatchp 	34
#define 	IDX_Method_IDispatchpp_IDispatchp_IDispatchp 	35
#define 	IDX_Method_IDispatchpp_VARIANTBOOL 	36
#define 	IDX_Method_void_IDispatchp 	37
#define 	IDX_Method_IDispatchpp_IDispatchp_BSTR 	38
#define 	IDX_Method_IDispatchpp_oDoVARIANTBOOL 	39
#define 	IDX_Method_IDispatchpp_BSTR_IDispatchp 	40
#define 	IDX_Method_BSTRp_BSTR 	41
#define 	IDX_Method_BSTRp_BSTR_BSTR 	42
#define 	IDX_Method_longp_BSTR_o0oVARIANTp 	43
#define 	IDX_Method_VARIANTBOOLp_long 	44
#define 	IDX_Method_IDispatchpp_BSTR 	45
#define 	IDX_Method_IDispatchpp_long 	46
#define 	IDX_Method_IDispatchpp_o0oVARIANTp 	47
#define 	IDX_Method_IDispatchpp_VARIANT 	48
#define 	IDX_Method_IDispatchpp_o0oVARIANT_o0oVARIANT 	49
#define 	IDX_Method_VARIANTBOOLp_BSTR_oDolong_oDolong 	50
#define 	IDX_Method_void_long_long 	51
#define 	IDX_Method_longp_BSTR_oDolong 	52
#define 	IDX_Method_longp_BSTR_IDispatchp 	53
#define 	IDX_Method_void_BSTR 	54
#define 	IDX_Method_VARIANTBOOLp_BSTR_oDoVARIANTBOOL_o0oVARIANT 	55
#define 	IDX_Method_longp_BSTR_BSTR_oDolong 	56
#define 	IDX_Method_void_long 	57
#define 	IDX_Method_VARIANTp_long 	58
#define 	IDX_G_short 	59
#define 	IDX_Method_void_o0oVARIANTp 	60
#define 	IDX_Method_VARIANTBOOLp_BSTR_o0oVARIANT 	61
#define 	IDX_Method_void_VARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT 	62
#define 	IDX_Method_VARIANTBOOLp_BSTR_BSTR_o0oVARIANT 	63
#define 	IDX_Method_void_IDispatchp_o0oVARIANT 	64
#define 	IDX_Method_void_oDolong 	65
#define 	IDX_Method_void_SAFEARRAYPVARIANTP 	66
#define 	IDX_Method_IDispatchpp_oDoBSTR_o0oVARIANT_o0oVARIANT_o0oVARIANT 	67
#define 	IDX_Method_IDispatchpp_long_long 	68
#define 	IDX_Method_IDispatchpp_oDoBSTR_oDolong 	69
#define 	IDX_Method_longp_VARIANTp_long_o0oVARIANTp 	70
#define 	IDX_Method_void_oDoBSTR 	71
#define 	IDX_Method_VARIANTBOOLp_oDoBSTR 	72
#define 	IDX_Method_VARIANTp_oDoBSTR_oDoBSTR 	73
#define 	IDX_Method_IDispatchpp_oDoBSTR_oDoBSTR_oDoBSTR_oDoVARIANTBOOL 	74
#define 	IDX_Method_VARIANTp_BSTR_o0oVARIANTp_o0oVARIANTp 	75
#define 	IDX_Method_void_BSTR_o0oVARIANT_oDoBSTR 	76
#define 	IDX_Method_VARIANTp_BSTR_oDoBSTR 	77
#define 	IDX_Method_IDispatchpp_oDoBSTR_o0oVARIANTp_o0oVARIANTp 	78
#define 	IDX_GS_LONG 	79
#define 	IDX_Method_IDispatchpp_o0oVARIANT_o0oVARIANT_o0oVARIANT_o0oVARIANT 	80
#define 	IDX_GS_IDispatchp 	81
#define 	IDX_Method_IDispatchpp_BSTR_o0oVARIANTp 	82
#define 	IDX_Method_VARIANTBOOLp_BSTR_VARIANTp 	83
