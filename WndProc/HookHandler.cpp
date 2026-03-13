#include "..\stdafx.h"
#include "HookHandler.h"
#include "Handler.h"

CHookHandler::CHookHandler(HHOOK hook, HOOKPROC proc)
{
	m_hook = (HHOOK)hook;
	m_HookProc = proc;
	m_parents = new TListKey<CHandlerParent, CXWndProc*>();
}
CHookHandler::~CHookHandler()
{
	if(m_hook)
		::UnhookWindowsHookEx(m_hook);
	delete m_parents;
}

void CHookHandler::AddHandler(void *parent, IDispatch *pObject, BSTR method)
{
	CHandlerParent *pHandlerParent = m_parents->FindByKey((CXWndProc*)parent);
	if(!pHandlerParent)
	{
		pHandlerParent = new CHandlerParent(parent);
		m_parents->Add(pHandlerParent);
	}
	pHandlerParent->Add(pObject, method);
}

void CHookHandler::RemoveHandlerParent(void *parent)
{
	m_parents->RemoveByKey((CXWndProc*)parent);
}

BOOL CHookHandler::IsEmpty()
{
	if(!m_parents || m_parents->Count() == 0)
		return TRUE;
	else
		return FALSE;
}

void CHookHandler::CallWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	//DISPPARAMS pars = {0};
	//VARIANT varResult = {0};
	//EXCEPINFO excepInfo = {0};
	//UINT nArgErr;
	//HRESULT rc;
	//
	//if(!SUCCEEDED(m_rc))
	//	return;
	//if(!(m_Msg == uMsg || m_Msg==0)) 
	//	return; 

	//DWORD id = ::GetCurrentThreadId();

	//LONG nRef;
	//IStream *pStream;
	//rc = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, m_pDispatch, &pStream);
	//if(rc == S_OK)
	//{
	//	IDispatch * pObject;
	//	rc = CoGetInterfaceAndReleaseStream(pStream, IID_IDispatch, (void**)&pObject);
	//	if(rc == S_OK)
	//	{
	//		pars.cArgs = 4;
	//		pars.rgvarg = new VARIANT[pars.cArgs];
	//		UINT np = pars.cArgs - 1;
	//		V_VT(pars.rgvarg + np) = VT_UI4;
	//		V_UI4(pars.rgvarg + np) = 0;
	//		np--;
	//		V_VT(pars.rgvarg + np) = VT_UI4;
	//		V_UI4(pars.rgvarg + np) = (UINT)uMsg;
	//		np--;
	//		V_VT(pars.rgvarg + np) = VT_UI4;
	//		V_UI4(pars.rgvarg + np) = (UINT)wParam;
	//		np--;
	//		if(uMsg == WM_SETTEXT)
	//		{
	//			wchar_t * wCaption = new wchar_t[2048];
	//			::MultiByteToWideChar(0, 0, (char*)lParam, -1, wCaption, 2048);
	//			V_VT(pars.rgvarg + np) = VT_BSTR;
	//			V_BSTR(pars.rgvarg + np) = ::SysAllocString(wCaption);
	//			delete wCaption;
	//		}
	//		else
	//		{
	//			V_VT(pars.rgvarg + np) = VT_UI4;
	//			V_UI4(pars.rgvarg + np) = (UINT)lParam;
	//		}
	//		rc = pObject->Invoke(m_nMethod, GUID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &pars, &varResult, &excepInfo, &nArgErr);
	//		for(int i=0; i<pars.cArgs; i++)
	//		{
	//			::VariantClear(pars.rgvarg + i);
	//		}
	//		delete pars.rgvarg;
	//		nRef = pObject->Release();
	//	}
	//}
}

void CHookHandler::CallHandlers(int code, WPARAM wParam, LPARAM lParam)
{
	CHandlerParent *pHandlerParent = NULL;
	m_parents->MoveFirst();
	while((pHandlerParent=m_parents->Next()))
	{
		pHandlerParent->CallHandlers(code, wParam, lParam);
	}
}

