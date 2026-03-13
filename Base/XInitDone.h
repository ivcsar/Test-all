#pragma once

const GUID IID_IInitDone = 
    {0xab634001,0xf13d,0x11d0,{0xa4,0x59,0x00,0x40,0x95,0xe1,0xda,0xea}};
const GUID IID_IPropertyProfile =
    {0xab634002,0xf13d,0x11d0,{0xa4,0x59,0x00,0x40,0x95,0xe1,0xda,0xea}};
const GUID IID_IAsyncEvent =
    {0xab634004,0xf13d,0x11d0,{0xa4,0x59,0x00,0x40,0x95,0xe1,0xda,0xea}};
const GUID IID_ILanguageExtender =
    {0xab634003,0xf13d,0x11d0,{0xa4,0x59,0x00,0x40,0x95,0xe1,0xda,0xea}};
const GUID IID_IStatusLine =
    {0xab634005,0xf13d,0x11d0,{0xa4,0x59,0x00,0x40,0x95,0xe1,0xda,0xea}};
const GUID IID_IExtWndsSupport =
    {0xefe19ea0,0x09e4,0x11d2,{0xa6,0x01,0x00,0x80,0x48,0xda,0x00,0xde}};
const GUID IID_IPropertyLink =
    {0x52512a61,0x2a9d,0x11d1,{0xa4,0xd6,0x00,0x40,0x95,0xe1,0xda,0xea}};

class IInitDone : public IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE Init(IDispatch * pDispatch) = 0;
	virtual HRESULT STDMETHODCALLTYPE Done(void) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetInfo(SAFEARRAY **pInfo) = 0;
};

struct __declspec(uuid("efe19ea0-09e4-11d2-a601-008048da00de"))
IExtWndsSupport : IUnknown
{
	STDMETHOD(GetAppMainFrame)(HWND *hwnd) = 0;
	STDMETHOD(GetAppMDIFrame)(HWND *hwnd) = 0;
	STDMETHOD(CreateAddInWindow)(
		BSTR bstrProgID,
		BSTR bstrWindowName,
		long dwStyles,
		long dwExStyles,
		struct tagRECT *rctl,
		long Flags,
		wireHWND *pHwnd,
		IDispatch **pDisp) = 0;
};

class ILanguageExtender : public IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE RegisterExtensionAs(BSTR * pExtensionName) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetNProps(long *plProps) = 0;
	virtual HRESULT STDMETHODCALLTYPE FindProp(BSTR bstrPropName, long *plPropNum) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetPropName(long lPropNum, long lPropAlias, BSTR *pbstrPropName) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetPropVal(long lPropNum, VARIANT *pvarPropVal) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetPropVal(long lPropNum, VARIANT *varPropVal) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsPropReadable(long lPropNum, BOOL *pboolPropRead) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsPropWritable(long lPropNum, BOOL *pboolPropWrite) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetNMethods(long *plMethods) = 0;
	virtual HRESULT STDMETHODCALLTYPE FindMethod(BSTR bstrMethodName, long *plMethodNum) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetMethodName(long lMethodNum, long lMethodAlias, BSTR *pbstrMethodName) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetNParams(long lMethodNum, long *plParams) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetParamDefValue(long lMethodNum, long lParamNum, VARIANT *pvarParamDefValue) = 0;
	virtual HRESULT STDMETHODCALLTYPE HasRetVal(long lMethodNum, BOOL *pboolRetValue) = 0;
	virtual HRESULT STDMETHODCALLTYPE CallAsProc(long lMethodNum, SAFEARRAY **paParams) = 0;
	virtual HRESULT STDMETHODCALLTYPE CallAsFunc(long lMethodNum, VARIANT *pvarRetValue, SAFEARRAY **paParams) = 0;
};

////////////// CXInitDone
class CXInitDone : public IInitDone, ILanguageExtender
{
	ULONG m_nRef;
public:
	ULONG STDMETHODCALLTYPE AddRef(void)
	{
		m_nRef++;
		return m_nRef;
	}
	ULONG STDMETHODCALLTYPE Release(void)
	{
		m_nRef--;
		if(m_nRef == 0) 
			delete this;
		return m_nRef;
	}
	HRESULT STDMETHODCALLTYPE QueryInterface( 
                /* [in] */ REFIID riid,
                /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		if(riid==IID_IInitDone)
			*ppvObject = static_cast<IInitDone*>(this);
		else if(riid==IID_IUnknown)
			*ppvObject = static_cast<IInitDone*>(this);
		else if(riid==IID_ILanguageExtender)
			*ppvObject = static_cast<ILanguageExtender*>(this);
		else 
			return E_NOINTERFACE;
		AddRef();
		return S_OK;
	}
	HWND m_hwnd1C;
	LPDISPATCH m_lpDispatch1C;
	IErrorLog * m_iErrorLog;
public:
	CXInitDone();
	~CXInitDone();
public:
	STDMETHOD(Init)(IDispatch* pDispatch);
	STDMETHOD(Done)(void);
	STDMETHOD(GetInfo)(SAFEARRAY **pInfo);
	STDMETHOD(RegisterExtensionAs)(BSTR* pExtensionName);
	
	STDMETHOD(GetNProps)(long *plProps);
	STDMETHOD(FindProp)(BSTR bstrPropName, long *plPropNum);
	STDMETHOD(GetPropName)(long lPropNum, long lPropAlias, BSTR *pbstrPropName);
	STDMETHOD(GetPropVal)(long lPropNum, VARIANT *pvarPropVal);
	STDMETHOD(SetPropVal)(long lPropNum, VARIANT *varPropVal);
	STDMETHOD(IsPropReadable)(long lPropNum, BOOL *pboolPropRead);
	STDMETHOD(IsPropWritable)(long lPropNum, BOOL *pboolPropWrite);

	STDMETHOD(GetNMethods)(long *plMethods);
	STDMETHOD(FindMethod)(BSTR bstrMethodName, long *plMethodNum);
	STDMETHOD(GetMethodName)(long lMethodNum, long lMethodAlias, BSTR *pbstrMethodName);
	STDMETHOD(GetNParams)(long lMethodNum, long *plParams);
	STDMETHOD(GetParamDefValue)(long lMethodNum, long lParamNum, VARIANT *pvarParamDefValue);
	STDMETHOD(HasRetVal)(long lMethodNum, BOOL *pboolRetValue);
	STDMETHOD(CallAsProc)(long lMethodNum, SAFEARRAY **paParams);
	STDMETHOD(CallAsFunc)(long lMethodNum, VARIANT *pvarRetValue, SAFEARRAY **paParams);

public:
	HRESULT LoadXCom(BSTR, BSTR, void **pObject);
	BOOL AttachAddin(BSTR dllname, BSTR bstrCLSID); 
	VOID ExcepInfoEvent(DWORD hResult);
	VOID ExcepInfoEvent(HRESULT hResult, BSTR errorDescription);
	BOOL InitWnd();
protected:
	LONG LoadParameters(long nMethod, SAFEARRAY **parameters, VARIANT **varPars);
	void FreeParameters(VARIANT *varPars, LONG cArgs);
	void Unload(BSTR dllname);
	CMethod *GetMethodInfo(long lMethodNum);
	wchar_t *m_Version;
};

