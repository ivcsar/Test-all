#include "..\stdafx.h"
#include "Events.h"

STDAPI GetEvents(LPVOID *ppv)
{
	CEvents *events = new CEvents();
	events->AddRef();
	*ppv = (void*)events;
	return S_OK;
}
HRESULT CEvents::GetConnectionInterface(IID * pIID)
{
	return S_OK;
}
HRESULT CEvents::GetConnectionPointContainer(IConnectionPointContainer ** ppCPC)
{
	return S_OK;
}
HRESULT CEvents::Advise(IUnknown * pUnkSink, DWORD * pdwCookie)
{
	return S_OK;
}
HRESULT CEvents::Unadvise(DWORD dwCookie)
{
	return S_OK;
}
HRESULT CEvents::EnumConnections(IEnumConnections ** ppEnum)
{
	return S_OK;
}

HRESULT CEvents::QueryInterface(const IID &riid, void **ppv)
{
	if (riid == IID_IDispatch)
		*ppv = this;
	else if (riid == IID_IUnknown) 
		*ppv = this;
	else if (riid == IID_IRunnableObject)
		*ppv = dynamic_cast<IRunnableObject*>(this);
	else if (riid == IID_IConnectionPointContainer)
		*ppv = dynamic_cast<IConnectionPointContainer*>(this);
	else if (riid == IID_IDispatchEx)
		*ppv = dynamic_cast<IDispatchEx*>(this);
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

HRESULT CEvents::GetIDsOfNames(const IID &iid, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
{
	int i;

	//METHOD_DESCRIPTION * pMethodsList = MethodsList;

	//for (i = 0; i < sizeof(MethodsList) / sizeof(MethodsList[0]); i++)
	//{
	//	if (_wcsicmp(*methodName, pMethodsList->methodName) == 0)
	//	{
	//		*dispid = pMethodsList->nMethod;
	//		return S_OK;
	//	}
	//	pMethodsList++;
	//}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CEvents::Invoke(DISPID nf,	const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult,	EXCEPINFO *pExcepInfo,	UINT *nArgErr)
{
	HRESULT rc = S_OK;

	return rc;
}

HRESULT CEvents::GetRunningClass(LPCLSID lpClsid)
{
	*lpClsid = __uuidof(CEvents);
	return S_OK;
}
HRESULT CEvents::Run(LPBINDCTX pbc)
{
	return S_OK;
}
BOOL CEvents::IsRunning(void)
{
	return TRUE;
}
HRESULT CEvents::LockRunning(BOOL fLock, BOOL fLastUnlockCloses)
{
	return S_OK;
}
HRESULT CEvents::SetContainedObject(BOOL fContained)
{
	return S_OK;
}

HRESULT CEvents::GetDispID(BSTR bstrName, DWORD, DISPID *pid) 
{ 
	return S_OK; 
}
HRESULT CEvents::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller) 
{ 
	return S_OK; 
}
HRESULT CEvents::DeleteMemberByName(BSTR bstrName, DWORD grfdex) 
{ 
	return S_OK; 
}
HRESULT CEvents::DeleteMemberByDispID(DISPID id) 
{ 
	return S_OK; 
}
HRESULT CEvents::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex) 
{ 
	return S_OK; 
}
HRESULT CEvents::GetMemberName(DISPID id, BSTR *pbstrName) 
{ 
	return S_OK; 
}
HRESULT CEvents::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid) 
{ 
	return S_OK; 
}
HRESULT CEvents::GetNameSpaceParent(IUnknown **ppunk) 
{ 
	return S_OK; 
}

HRESULT CEvents::EnumConnectionPoints(__RPC__deref_out_opt IEnumConnectionPoints **ppEnum)
{
	CEventsEnumConnectionPoint *pConnectionPoint = new CEventsEnumConnectionPoint();
	pConnectionPoint->AddRef();
	*ppEnum = pConnectionPoint;
	return S_OK;
}
HRESULT CEvents::FindConnectionPoint(__RPC__in REFIID riid, __RPC__deref_out_opt IConnectionPoint **ppCP)
{
	return S_OK;
}

ULONG STDMETHODCALLTYPE CEventsEnumConnectionPoint::AddRef()
{
	m_nRef++;
	return m_nRef;
}
ULONG STDMETHODCALLTYPE CEventsEnumConnectionPoint::Release()
{
	m_nRef--;
	if (m_nRef == 0)
	{
		delete this;
		return 0;
	}
	else
		return m_nRef;
}

HRESULT CEventsEnumConnectionPoint::QueryInterface(const IID &riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = this;
	else if (riid == IID_IEnumConnectionPoints)
		*ppv = this;
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CEventsEnumConnectionPoint::Next(
	/* [in] */ ULONG cConnections,
	/* [length_is][size_is][out] */ LPCONNECTIONPOINT *ppCP,
	/* [out] */ ULONG *pcFetched)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CEventsEnumConnectionPoint::Skip(
	/* [in] */ ULONG cConnections)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CEventsEnumConnectionPoint::Reset(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CEventsEnumConnectionPoint::Clone(
	/* [out] */ __RPC__deref_out_opt IEnumConnectionPoints **ppEnum)
{
	return S_OK;
}
