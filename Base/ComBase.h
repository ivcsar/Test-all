#pragma once
#include "stdafx.h"

#ifndef __COMBASE__H__
#define __COMBASE__H__

#include "..\XNative\NativeBase.h"

class ComBase : public IDispatch
{
	LONG m_nRef;
public:
	CMethod *m_pComMethods;
	UINT m_nComMethodCount;
public:
	ComBase();
	virtual ~ComBase();
public:
	virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppv);
	virtual ULONG __stdcall AddRef(void);
	virtual ULONG __stdcall Release(void);
	virtual HRESULT __stdcall GetTypeInfoCount(UINT *pInfoCount);
	virtual HRESULT __stdcall GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo);
public:
	virtual HRESULT __stdcall GetIDsOfNames(const IID &, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid);
	virtual HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *) = 0;
};
#endif