#pragma once
#include "..\stdafx.h"
#include "..\XCom.h"

class CXFunctions : public IDispatch
{
	ULONG nRef;
	struct METHOD_DESCRIPTION
	{
		wchar_t * methodName;
		int nMethod;
	};
	static METHOD_DESCRIPTION MethodsList[];
public:
	CXFunctions()
	{
		nRef = 0;
		IncrementRef();
	}
	~CXFunctions()
	{
		DecrementRef();
	}
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv)
	{
		if(riid == IID_IUnknown)
			*ppv = dynamic_cast<IUnknown*>(this);
		else if(riid == IID_IDispatch)
			*ppv = dynamic_cast<IDispatch*>(this);
		else
		{
			*ppv = NULL;
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}
	ULONG __stdcall AddRef(void)
	{
		nRef++;
		return nRef;
	}
	ULONG __stdcall Release(void)
	{
		if(nRef > 0) 
			nRef--;
		if(nRef == 0)
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
public:
	///// methods
	HRESULT XGetMessage(DISPPARAMS * parameters, VARIANT * pVarResult);
	ULONG MemSize();
	HRESULT ExecInNewThread(DISPPARAMS * parameters, VARIANT * pVarResult);
	HRESULT SetMemory(DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO*);
	HRESULT Do(DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO*);
	HRESULT Unload(BSTR dllname);
};