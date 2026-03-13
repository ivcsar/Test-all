#pragma once
#include "..\stdafx.h"
#include "..\XCom.h"
#include "OCIdl.h"
#include "DispEx.h"
//#include "OCIdl.idl"

const GUID IID_ICtx { 0xE7210190, 0x61F4, 0x11D4, { 0x94, 0x1D, 0x00, 0x80, 0x48, 0xDA, 0x11, 0xF9 } };
MIDL_INTERFACE("E7210190-61F4-11D4-941D-008048DA11F9") ICtx2Disp
{
	IUnknown * getContext();
};
struct __declspec(uuid("FD7B6CC3-DC8E-11D2-B8D0-008048DA0335")) IPreConnect
{

};
const GUID IID_FD7B6CC3{ 0 };

class __declspec(uuid("FD7B6CC3-DC8E-11D2-B8D0-008048DA0335")) CEvents : public IRunnableObject, ICtx2Disp, IConnectionPoint, IConnectionPointContainer, IDispatchEx
{
	ULONG nRef;
public:
	CEvents()
	{
		nRef = 0;
	}
	~CEvents()
	{

	}
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv);
	ULONG __stdcall AddRef(void)
	{
		nRef++;
		return nRef;
	}
	ULONG __stdcall Release(void)
	{
		if (nRef > 0)
			nRef--;
		if (nRef == 0)
		{
			delete this;
			return 0;
		}
		return nRef;
	}
	HRESULT __stdcall GetTypeInfoCount(UINT *pInfoCount)
	{
		//if(pInfoCount)
		//	*pInfoCount = 0;
		return E_NOTIMPL;
	}
	HRESULT __stdcall GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo)
	{
		//ITypeInfo *pI;
		//pI = new ITypeInfo();
		//pI->
		return E_NOTIMPL;
	}
	HRESULT __stdcall GetIDsOfNames(const IID &, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid);
	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
	HRESULT __stdcall GetRunningClass(/* out */ LPCLSID lpClsid);
	HRESULT STDMETHODCALLTYPE Run(LPBINDCTX pbc);
	BOOL STDMETHODCALLTYPE IsRunning(void);
	HRESULT STDMETHODCALLTYPE LockRunning(BOOL fLock, BOOL fLastUnlockCloses);
	HRESULT STDMETHODCALLTYPE SetContainedObject(BOOL fContained);
	HRESULT __stdcall GetConnectionInterface(IID * pIID);
	HRESULT __stdcall GetConnectionPointContainer(IConnectionPointContainer ** ppCPC);
	HRESULT __stdcall Advise(IUnknown * pUnkSink, DWORD * pdwCookie);
	HRESULT __stdcall Unadvise(DWORD dwCookie);
	HRESULT __stdcall EnumConnections(IEnumConnections ** ppEnum);

	HRESULT STDMETHODCALLTYPE GetDispID(BSTR bstrName, DWORD, DISPID *pid);
	HRESULT STDMETHODCALLTYPE InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	HRESULT STDMETHODCALLTYPE DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	HRESULT STDMETHODCALLTYPE DeleteMemberByDispID(DISPID id);

	HRESULT STDMETHODCALLTYPE GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	HRESULT STDMETHODCALLTYPE GetMemberName(DISPID id, BSTR *pbstrName);
	HRESULT STDMETHODCALLTYPE GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	HRESULT STDMETHODCALLTYPE GetNameSpaceParent(IUnknown **ppunk);

	// IConnectionPointContainer
	HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints **ppEnum);
	HRESULT STDMETHODCALLTYPE FindConnectionPoint(__RPC__in REFIID riid, __RPC__deref_out_opt IConnectionPoint **ppCP);
};

class CEventsEnumConnectionPoint : public IEnumConnectionPoints
{
	ULONG m_nRef;
public:
	CEventsEnumConnectionPoint()
	{
		m_nRef = 0;
	}
public:
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv);

public:
	HRESULT STDMETHODCALLTYPE Next(
		/* [in] */ ULONG cConnections,
		/* [length_is][size_is][out] */ LPCONNECTIONPOINT *ppCP,
		/* [out] */ ULONG *pcFetched);

	HRESULT STDMETHODCALLTYPE Skip(
		/* [in] */ ULONG cConnections);

	HRESULT STDMETHODCALLTYPE Reset(void);

	HRESULT STDMETHODCALLTYPE Clone(
		/* [out] */ __RPC__deref_out_opt IEnumConnectionPoints **ppEnum);
};