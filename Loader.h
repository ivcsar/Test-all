#pragma once
#include "stdafx.h"

class CLoader :
	public IDispatch
{
	ULONG m_nRef;
public:
	CLoader(void);
	~CLoader(void);
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
		m_nRef++;
		return m_nRef;
	}
	ULONG __stdcall Release(void)
	{
		if(m_nRef > 0) 
			m_nRef--;
		if(m_nRef == 0)
		{
			delete this;
			return 0;
		}
		return m_nRef;
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
protected:
	HRESULT ExcepInfo(DWORD rc, EXCEPINFO*);
	HRESULT CreateObject(HMODULE hModule, BSTR classID, CLSID &riid, void ** ppv);
	HRESULT TypeInfo(BSTR dllname, CLSID & clsid);
	HRESULT LoadXCom(BSTR dllname, BSTR classID, void **pObject);
};

