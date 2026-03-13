#pragma once
#include "stdafx.h"
#include "Thunk\Replacement.h"

class CLogHook : public IDispatch
{
	LONG m_nRef;
public:
	CLogHook();
	~CLogHook();
protected:
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv);
	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);
	HRESULT __stdcall GetTypeInfoCount(UINT *pInfoCount);
	HRESULT __stdcall GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo);
public:
	HRESULT __stdcall GetIDsOfNames(const IID &, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid);
	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);

private:
	TList <Replacement::REPLACEMENT> *m_replacement;
	static int ReplaceWrite(LPCSTR c_dllname, LPCSTR win_dllname, void*);
	void SetHook();
};