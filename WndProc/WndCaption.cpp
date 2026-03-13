#include "WndCaption.h"
#include <wchar.h>

CXWndCaption* CXWndCaption::m_Ptr = NULL;

HWND GetAppWnd();

//void SetCaptionHook(wchar_t *caption)
//{
//	CXWndCaption::setCaptionHook(caption);
//}

CXWndCaption::CXWndCaption(BSTR wCaption, VARIANT *wincap)
{
	//m_hwnd = CXInitDone::getptr()->GetHWND();
	//HINSTANCE xcomH = ::GetModuleHandleA("XComLoader.dll");
	wchar_t* wcap;
	
	m_Ptr = this;
	m_wCaption = NULL;
	m_WndProc = NULL;

	if (wincap)
	{
		V_VT(wincap) = VT_EMPTY;
		//V_BOOL(wincap) = FALSE;
	}

	m_hwnd = GetAppWnd();
	if (!m_hwnd)
		return;

	if (wincap)
	{
		wcap = new wchar_t[1024];
		int lwcap = ::GetWindowTextW(m_hwnd, wcap, 1024);
		wcap[lwcap] = ' ';
		::GetClassNameW(m_hwnd, wcap + lwcap + 1, 1024 - lwcap);
		V_VT(wincap) = VT_BSTR;
		V_BSTR(wincap) = ::SysAllocString(wcap);
		delete wcap;
	}

	if(wCaption && *wCaption != L'\0')
	{
		m_wCaption = ::SysAllocString(wCaption);
		::SetWindowTextW(m_hwnd, m_wCaption);
		m_WndProc = (WNDPROC)GetWindowLongPtrW(m_hwnd, GWLP_WNDPROC);
		SetWindowLongPtrW(m_hwnd, GWLP_WNDPROC, (LONG_PTR)WndProc);			// ставим новую функцию
	}
	else
	{
		wcap = new wchar_t[1024];
		int lw = ::GetWindowTextW(m_hwnd, wcap, 1024);
		m_WndProc = (WNDPROC)GetWindowLongPtrW(m_hwnd, GWLP_WNDPROC);
		SetWindowLongPtrW(m_hwnd, GWLP_WNDPROC, (LONG_PTR)WndProc);			// ставим новую функцию
		::SetWindowTextW(m_hwnd, wcap);
		delete wcap;
	}
	//IncrementRef();
}
CXWndCaption::~CXWndCaption()
{
	if(m_WndProc)
		::SetWindowLongPtrW(m_hwnd, GWLP_WNDPROC, (LONG_PTR)m_WndProc);
	if(m_wCaption)
		::SysFreeString(m_wCaption);
	m_hwnd = NULL;
	m_Ptr = NULL;
	//DecrementRef();
}
void CXWndCaption::Release()
{
	if(m_Ptr)
		delete m_Ptr;
}
LRESULT CXWndCaption::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 0; 
	
	if(uMsg==WM_SETTEXT && hwnd==m_Ptr->m_hwnd)
	{
		if(m_Ptr->m_wCaption) 
			lParam = (LONG_PTR)m_Ptr->m_wCaption;
		else
		{
			if(wmemcmp(		L"1С:Предприятие", (wchar_t*)lParam, 14)==0)
				lParam = (LONG_PTR)(((wchar_t*)lParam) + 17);
			else if(wmemcmp(L"1C:Enterprise", (wchar_t*)lParam, 13)==0)
				lParam = (LONG_PTR)(((wchar_t*)lParam) + 16);
		}
	}
	lResult = CallWindowProcW(m_Ptr->m_WndProc, hwnd, uMsg, wParam, lParam);
	return lResult;
}
BOOL CXWndCaption::setCaptionHook(BSTR caption, VARIANT *result)
{
	if(caption)
	{
		Release();
		new CXWndCaption(caption, result);
		if (m_Ptr->m_hwnd)
			return TRUE;
	}
	else
	{
		if (!m_Ptr)
		{
			new CXWndCaption(NULL, result);
			if (m_Ptr->m_hwnd)
				return TRUE;
		}
	}
	return FALSE;
}
// by Hook - не получилось
LRESULT CXWndCaption::CallGetMsgProc(int code, WPARAM wParam, LPARAM lParam)
{
		
	if(code < 0)
		return ::CallNextHookEx(NULL, code, wParam, lParam);
	if(code == HC_ACTION)
	{
		MSG * pMsg = (MSG*)lParam;
		if(pMsg->message == WM_SETTEXT)
		{
			wchar_t *wcap = (wchar_t*)pMsg->lParam;
			pMsg->lParam = (LONG_PTR)L"Мой заголовок";
		}

	}
	return ::CallNextHookEx(NULL, code, wParam, lParam);
}
void CXWndCaption::SetWndCaptionByHook()
{
	static HHOOK hook = NULL;
	if(!hook)
		hook = ::SetWindowsHookW(WH_GETMESSAGE, CallGetMsgProc);
}