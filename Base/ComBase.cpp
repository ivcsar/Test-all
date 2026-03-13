#include "ComBase.h"
#include "XCom.h"

ComBase::ComBase()
{
	m_nRef = 0;
	m_nComMethodCount = 0;
	m_pComMethods = NULL;
}
ComBase::~ComBase()
{

}

HRESULT __stdcall ComBase::QueryInterface(const IID &riid, void **ppv)
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
ULONG __stdcall ComBase::AddRef(void)
{
	IncrementRef();
	m_nRef++;
	return m_nRef;
}
ULONG __stdcall ComBase::Release(void)
{
	DecrementRef();
	if(m_nRef > 0) 
		m_nRef--;
	if(m_nRef == 0)
	{
		delete this;
		return 0;
	}
	return m_nRef;
}
HRESULT __stdcall ComBase::GetTypeInfoCount(UINT *pInfoCount)
{
	return E_NOTIMPL;
}
HRESULT __stdcall ComBase::GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo)
{
	return E_NOTIMPL;
}
HRESULT ComBase::GetIDsOfNames(const IID &riid, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
{
	HRESULT rc = DISP_E_MEMBERNOTFOUND;
	UINT i, j;

	for(i = 0; i < this->m_nComMethodCount; i++)
	{
		for(j=0; j < nMethodName; j++)
		{
			if(_wcsicmp(methodName[j], m_pComMethods[i].method) == 0)
			{
				dispid[j] = m_pComMethods[i].id;
				rc = S_OK;
			}
		}
	}
	return rc;
}
#pragma endregion
