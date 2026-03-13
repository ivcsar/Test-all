#include "..\stdafx.h"
#include <Psapi.h>
#include "..\XCom.h"
#include "XWndProc.h"
#include "HookHandler.h"
#include "WndCaption.h"
#include "..\resource.h"

#pragma comment (lib, "Psapi.lib")

TListKey <CHookHandler, HOOKPROC> gl_HookList;
extern HINSTANCE gl_hInstance;

enum XWNDPROC_METHODS
{
    emSetWndProcedure = 1,
	emSetKeyboardProc,
	emSetCaptionHook,
};
struct XWNDPROC_VTABLE
{
	wchar_t * method;
	int nMethod;
} MethodsList [] = 
{
	L"SetWndProcedure", emSetWndProcedure,
	L"SetKeyboardProc",	emSetKeyboardProc,
	L"SetCaptionHook",	emSetCaptionHook,
};

// функция создания объекта
STDAPI GetWndProc(LPVOID *ppv)
{
	CXWndProc *pWndProc = new CXWndProc();
	*ppv = (void*)pWndProc;
	return S_OK;
}

CXWndProc::CXWndProc()
{
	m_nRef = 0;
	AddRef();
}
CXWndProc::~CXWndProc()
{
	CHookHandler * pHook = NULL;
	gl_HookList.MoveFirst();
	while((pHook = gl_HookList.Next()))
	{
		pHook->RemoveHandlerParent(this);
		if(pHook->IsEmpty())
		{
			gl_HookList.Remove(pHook);
		}
	}
}

#pragma region CXWndProc IDispatch
HRESULT __stdcall CXWndProc::QueryInterface(const IID &riid, void **ppv)
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
ULONG __stdcall CXWndProc::AddRef(void)
{
	IncrementRef();
	m_nRef++;
	return m_nRef;
}
ULONG __stdcall CXWndProc::Release(void)
{
	if(m_nRef > 0) 
		m_nRef--;
	if(m_nRef == 0)
	{
		delete this;
		return 0;
	}
	DecrementRef();
	return m_nRef;
}
HRESULT __stdcall CXWndProc::GetTypeInfoCount(UINT *pInfoCount)
{
	//if (pInfoCount)
	//{
	//	*pInfoCount = 1;
	//	return NOERROR;
	//}
	//else
	//	return E_INVALIDARG;
	return E_NOTIMPL;
}
HRESULT __stdcall CXWndProc::GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo)
{
	return E_NOTIMPL;
}
HRESULT CXWndProc::GetIDsOfNames(const IID &riid, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
{
	HRESULT rc = DISP_E_MEMBERNOTFOUND;
	UINT i, j;

	for(i = 0; i < sizeof(MethodsList) / sizeof(MethodsList[0]); i++)
	{
		for(j=0; j < nMethodName; j++)
		{
			if(_wcsicmp(methodName[j], MethodsList[i].method) == 0)
			{
				dispid[j] = MethodsList[i].nMethod;
				rc = S_OK;
			}
		}
	}
	return rc;
}

HRESULT CXWndProc::Invoke(DISPID nf, const IID &, LCID, WORD propertyType, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT rc = S_FALSE;
	VARIANT var = {0};

	if(nf == emSetWndProcedure)
	{
		rc = SetWndProcedure(parameters, &var);
		if (IsVarVal(pVarResult))
		{
			*pVarResult = var;
		}
	}
	else if(nf == emSetKeyboardProc)
	{
		rc = SetKeyboardProc(parameters);
	}
	else if(nf == emSetCaptionHook)
	{
		//if(parameters->cArgs > 0)
		//	SetCaption(get_BSTR(parameters->rgvarg[parameters->cArgs - 1]));
		//else
		//	SetCaption(L"");
		int rt = CXWndCaption::setCaptionHook(L"", pVarResult);
		rc = S_OK;
	}
	return rc;
}
#pragma endregion

HRESULT CXWndProc::SetWndProcedure(DISPPARAMS * parameters, VARIANT * pVarResult)
{
	HRESULT rc;

	DWORD id = ::GetCurrentThreadId();

	if(parameters->cArgs > 1)
	{
		//HWND hwnd = CXInitDone::getptr()->GetHWND();
		//char * caption = new char[1024];
		//::GetWindowText(hwnd, caption, 1024);
		//::MessageBox(hwnd, caption, NULL, 0);
		//delete caption;
		//CWndProc * pWndProc = gl_WndProcList.Find(hwnd);
		//if(!pWndProc)
		//{
		//	pWndProc = new CWndProc(hwnd, CXWndProc::XWndProc);
		//	gl_WndProcList.Add(pWndProc);
		//}
		//LPDISPATCH pObject = parameters->rgvarg[parameters->cArgs - 1].pvarVal->pdispVal;
		//BSTR method = get_BSTR(parameters->rgvarg[parameters->cArgs - 2]);
		//if(parameters->cArgs > 2)
		//	pWndProc->AddHandler(this, pObject, method, &parameters->rgvarg[parameters->cArgs - 3]);
		//else
		//	pWndProc->AddHandler(this, pObject, method, NULL);
		rc = S_OK;
	}
	else
	{
		rc = DISP_E_BADPARAMCOUNT;
	}
	return rc;
}

LRESULT __stdcall XKeyboardHook(int code, WPARAM wParam, LPARAM lParam)
{
	if(code < 0)
	{
		return ::CallNextHookEx(NULL, code, wParam, lParam);
	}

	if(code == HC_ACTION)
	{
		KEYBOARD_LPARAM *key = (KEYBOARD_LPARAM*)&lParam;
	//	//unsigned int transitionState = 0;
	//	//unsigned int preKeyState = 0;
	//	//DWORD hw = HIWORD(lParam);
	//	//if((KF_UP & hw) == KF_UP) 
	//	//	transitionState = 1;
	//	//if((0x4000 & hw) == 0x4000) 
	//	//	preKeyState = 1;
	//	//DWORD lw = LOWORD(lParam);
	//	//unsigned int scanCode = hw & 0x00FF;
	//	//if(scanCode != key->code || transitionState != key->transState || preKeyState != key->preKeyState)
	//	//	::MessageBoxA(::GetActiveWindow(), "Отличаются!", 0, 0);
		CHookHandler * pHook = NULL;
		pHook = gl_HookList.FindByKey(XKeyboardHook);
		if(pHook)
		{
			pHook->CallHandlers(wParam, key->preKeyState, key->code);
		}
	}
	return ::CallNextHookEx(NULL, code, wParam, lParam); 
}

HRESULT CXWndProc::SetKeyboardProc(DISPPARAMS * parameters)
{
	HRESULT hr;

	if(parameters->cArgs > 1)
	{
		LPDISPATCH pObject = parameters->rgvarg[parameters->cArgs - 1].pvarVal->pdispVal;
		BSTR method = get_BSTR(parameters->rgvarg[parameters->cArgs - 2]);
		//LRESULT (__stdcall CHook::*pKeyProc)(int, WPARAM, LPARAM) = &CHook::KeyboardHook;
		//void *procStub = (void*)&CHook::KeyboardHook;

		CHookHandler * pHookHandler = gl_HookList.FindByKey(XKeyboardHook);
		if(!pHookHandler)
		{
			HHOOK hook = ::SetWindowsHookEx(WH_KEYBOARD, XKeyboardHook, NULL, ::GetCurrentThreadId());
			pHookHandler = new CHookHandler(hook, XKeyboardHook);
			gl_HookList.Add(pHookHandler);
		}
		pHookHandler->AddHandler(this, pObject, method);
		hr = S_OK;
	}
	else
		hr = DISP_E_BADPARAMCOUNT;
	return hr;
}
void CXWndProc::SetCaption(BSTR caption)
{
	char *wndCaptionDllname = "XWndCaption.dll";
	//HANDLE hExe = GetModuleHandle(NULL);
	BYTE *pimage = reinterpret_cast<BYTE*>(gl_hInstance);
	IMAGE_OPTIONAL_HEADER *pIOH = 
		reinterpret_cast<IMAGE_OPTIONAL_HEADER*>(pimage + 
		reinterpret_cast<PIMAGE_DOS_HEADER>(gl_hInstance)->e_lfanew + 4 + 
		sizeof(IMAGE_FILE_HEADER));

	if (pIOH->Magic == 0x020B)	// 64
		wndCaptionDllname = "XWndCaption64.dll";
	else
		wndCaptionDllname = "XWndCaption.dll";

	char *modulename = new char[1024];
	DWORD lModuleName = ::GetModuleFileNameA(gl_hInstance, modulename, 1024);
	LONG i;
	i = lModuleName - 1;
	while(i > 0)
	{
		if(modulename[i] == '\\')
		{
			i++;
			break;
		}
		i--;
	}
	memmove(modulename + i, wndCaptionDllname, strlen(wndCaptionDllname) + 1);

	//HRSRC hRes = FindResource(gl_hInstance, MAKEINTRESOURCE(IDR_XWND_CAPTION_DLL), "XWND_CAPTION_DLL"); 
	//// загружаем ресурс
	//HGLOBAL hLoaded = LoadResource(gl_hInstance, hRes);
	//// получаем адрес загруженного ресурса
	//LPVOID lpDllResource = LockResource(hLoaded);
	//// получаем размер загруженного ресурса
	//DWORD dwSize = SizeofResource(gl_hInstance, hRes);
	//// теперь сохраним загруженный ресурс в виде файла
	//HANDLE hFile = CreateFile(modulename, GENERIC_WRITE, FILE_SHARE_WRITE, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
	//if(hFile == INVALID_HANDLE_VALUE)
	//{

	//}
	//else
	//{
	//	WriteFile(hFile, lpDllResource, dwSize, &dwSize, 0);
	//	CloseHandle(hFile);
	//}
	//// освобождаем загруженный ресурс
	//UnlockResource(hLoaded);
	
	HMODULE hDll = ::GetModuleHandleA(modulename);
	if(!hDll)
		hDll = ::LoadLibraryA(modulename);
	if(hDll)
	{
		void (_cdecl * pSetCaptionHook)(wchar_t*);
		pSetCaptionHook = (void (_cdecl*)(wchar_t*))GetProcAddress(hDll, "SetCaptionHook");
		if(pSetCaptionHook)
		{
			pSetCaptionHook(caption);
		}
	}
	delete [] modulename;
}