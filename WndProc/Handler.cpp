#include "..\stdafx.h"
#include "Handler.h"

#pragma region CHandler
CHandler::CHandler(LPDISPATCH pObject, BSTR method)
{
	LONG nref = pObject->AddRef();
	m_pDispatch = pObject;
	m_rc = m_pDispatch->GetIDsOfNames(IID_NULL, &method, 1, LOCALE_SYSTEM_DEFAULT, &m_Dispid);
	//if(varMsg)
	//{
	//	BSTR msgName = get_BSTR(*varMsg);
	//	if(msgName==L"")
	//	{
	//		m_Msg = V_UI4(varMsg);
	//	}
	//	else if(_wcsicmp(msgName, L"KEYUP") == 0)
	//		m_Msg = WM_KEYUP;
	//	else if(_wcsicmp(msgName, L"KEYDOWN") == 0)
	//		m_Msg = WM_KEYDOWN;
	//	else if(_wcsicmp(msgName, L"SETTEXT") == 0)
	//		m_Msg = WM_SETTEXT;
	//	else
	//		m_Msg = 0;
	//}
	//else
	//	m_Msg = 0;
}
CHandler::~CHandler()
{
	LONG nref;
	nref = m_pDispatch->Release();
	if(nref > 0)
	{
		nref = nref + 1;
	}
}
void CHandler::CallHandler(int code, WPARAM wParam, LPARAM lParam)
{
	DISPPARAMS pars = {0};
	VARIANT varResult = {0};
	EXCEPINFO excepInfo = {0};
	UINT nArgErr;
	HRESULT rc;
	
	if(!SUCCEEDED(m_rc))
		return;
	pars.cArgs = 3;
	pars.rgvarg = new VARIANT[pars.cArgs];
	UINT np = pars.cArgs - 1;
	V_VT(pars.rgvarg + np) = VT_I4;
	V_I4(pars.rgvarg + np) = code;
	np--;
	V_VT(pars.rgvarg + np) = VT_I4;
	V_I4(pars.rgvarg + np) = wParam;
	np--;
	V_VT(pars.rgvarg + np) = VT_I4;
	V_I4(pars.rgvarg + np) = lParam;

	LONG nRef = m_pDispatch->AddRef();
	rc = m_pDispatch->Invoke(m_Dispid, GUID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &pars, &varResult, &excepInfo, &nArgErr);
	nRef = m_pDispatch->Release();

	for(int i=0; i<pars.cArgs; i++)
	{
		::VariantClear(pars.rgvarg + i);
	}
	delete pars.rgvarg;
}
#pragma endregion

#pragma region CHandlerParent
CHandlerParent::CHandlerParent(void *parent)
{
	m_parent = parent;
	m_handlers = new TList<CHandler>();
}
CHandlerParent::~CHandlerParent()
{
	delete m_handlers;
}
	
int CHandlerParent::Compare(void * ptr)
{
	if(m_parent == ptr)
		return 0;
	else
		return 1;
}
void CHandlerParent::CallHandlers(int code, WPARAM wParam, LPARAM lParam)
{
	CHandler * pHandler = NULL;
	m_handlers->MoveFirst();
	while((pHandler=m_handlers->Next()))
	{
		pHandler->CallHandler(code, wParam, lParam);
	}
}
void CHandlerParent::Add(LPDISPATCH pObject, BSTR method)
{
	CHandler *handler = new CHandler(pObject, method);
	m_handlers->Add(handler);
}
BOOL CHandlerParent::IsEmpty()
{
	if(!m_handlers || m_handlers->Count() == 0)
		return TRUE;
	else
		return FALSE;
}
#pragma endregion