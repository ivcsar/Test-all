// XFunctions.cpp : Implementation of CXFunctions
//#define _WIN32_DCOM 1
//#define _CRT_SECURE_NO_DEPRECATE 1
//#define OEMRESOURCE 1

#include "..\stdafx.h"
#include "process.h"
#include "stdio.h"
#include <iostream>
#include "XFunctions.h"
//#include "..\XLib\XLib.h"
HRESULT XMLRecordCountLong(BSTR xmlFilename, ULONG parlevel, ULONG *nRecCount);
HRESULT get_Property(DISPPARAMS *pars, VARIANT *retValue);
HRESULT Execute(DISPPARAMS *pars, VARIANT *retValue);
HRESULT get_RefCount(DISPPARAMS *pars, VARIANT *retValue);
HRESULT IsValueInList(DISPPARAMS *pars, VARIANT *retValue);

#include <Psapi.h>

#pragma comment (lib, "psapi.lib")

#define BUFSIZE 4096

enum ENUM_METHOD_NUMBERS
{
    emGetMessage = 1,
	emSleep,
    emXmlRecordCount,
	emGetMilliseconds,
	emMemSize,
	emExecInNewThread,
	emSetGlobal,
	emCurrentUser,
	emProcessID,
	emVersion,
	emDo,
	emUnload,
	emUpdateWindow,
	emGetProperty,
	emGetRefCount,
	emExecute,
	emCompStrings,
	emIsValueInList
};
CXFunctions::METHOD_DESCRIPTION CXFunctions::MethodsList [] =
{
	L"XGetMessage",			emGetMessage,
	L"GetWinMessage",		emGetMessage,
	L"Execute",				emExecute,
	L"XSleep",				emSleep,
	L"Sleep",				emSleep,
	L"XmlRecordCount",		emXmlRecordCount,
	L"XmlRecordsCount",		emXmlRecordCount,
	L"XmlRecordsCounter",	emXmlRecordCount,
	L"XGetMilliseconds",	emGetMilliseconds,
	L"MemSize",				emMemSize,
	L"ExecInNewThread",		emExecInNewThread,
	L"SetGlobal",			emSetGlobal,
	L"SetMemory",			emSetGlobal,
	L"CurrentUser",			emCurrentUser,
	L"ProcessID",			emProcessID,
	L"Version",				emVersion,
	L"Do",					emDo,
	L"ВыполнитьВЦикле",		emDo,
	L"Unload",				emUnload, 
	L"UpdateWindow",		emUpdateWindow,
	L"GetProperty",			emGetProperty,
	L"GetRefCount",			emGetRefCount,
	L"CompStrings",			emCompStrings,
	L"IsValueInList",		emIsValueInList
};

STDAPI functions(LPVOID *ppv)
{
	CXFunctions *cf = new CXFunctions();
	cf->AddRef();
	*ppv = (void*)cf;
	return S_OK;
}

//////////////// CXFunctions
HRESULT CXFunctions::GetIDsOfNames(const IID &iid, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
{	
	int i;

	METHOD_DESCRIPTION * pMethodsList = MethodsList;

	for(i = 0; i < sizeof(MethodsList) / sizeof(MethodsList[0]); i++)
	{
		if(_wcsicmp(*methodName, pMethodsList->methodName) == 0)
		{
			*dispid = pMethodsList->nMethod;
			return S_OK;
		}
		pMethodsList++;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CXFunctions::Invoke(DISPID nf, 
							const IID &, 
							LCID, 
							WORD, 
							DISPPARAMS * parameters, 
							VARIANT * pVarResult, 
							EXCEPINFO *pExcepInfo, 
							UINT *nArgErr)
{
	HRESULT result = S_OK;

	switch(nf)
	{
		case emGetMessage:		// XGetMessage
			result = XGetMessage(parameters, pVarResult);
			break;
		case emExecute:
			result = Execute(parameters, pVarResult);
			break;
		case emGetMilliseconds:		// XGetMilliseconds
		{
			ULONG ms; 
			ms = ::GetTickCount();
			if(IsVarVal(pVarResult))	// для размещения результата передан не NULL
			{
				V_VT(pVarResult) = VT_UI4;
				V_UI4(pVarResult) = ms;
			}
			break;
		}
		case emSleep:			// XSleep
			::Sleep(parameters->rgvarg[0].pvarVal->intVal);
			break;
		case emXmlRecordCount:
		{
			ULONG recCount = 0;
			if (parameters->cArgs == 1)			// по умолчанию - 2-ой уровень
			{
				result = XMLRecordCountLong(parameters->rgvarg[0].pvarVal->bstrVal, 2, &recCount);
				if(result == S_OK && IsVarVal(pVarResult))	// для размещения результата передан не NULL
				{
					V_VT(pVarResult) = VT_UI4;
					V_UI4(pVarResult) = recCount;
				}
			}
			else if (parameters->cArgs == 2)
			{
				result = XMLRecordCountLong(parameters->rgvarg[1].pvarVal->bstrVal, 
											parameters->rgvarg[0].pvarVal->lVal, 
											&recCount); 
				if(result == S_OK && IsVarVal(pVarResult))	// для размещения результата передан не NULL
				{
					V_VT(pVarResult) = VT_UI4;
					V_UI4(pVarResult) = recCount;
				}
			}
			else
				result = DISP_E_BADPARAMCOUNT;
			break;
		}
		case emMemSize:
			if(IsVarVal(pVarResult))
			{
				V_VT(pVarResult)  = VT_UI4;
				V_UI4(pVarResult) = MemSize();
			}
			break;
		case emExecInNewThread:
			ExecInNewThread(parameters, pVarResult);
			break;
		case emSetGlobal:
			result = SetMemory(parameters, pVarResult, pExcepInfo);
			break;
		case emCurrentUser:
		{
			wchar_t * pwUserName = new wchar_t[1024];
			DWORD nb = 1024;
			if(::GetUserNameW(pwUserName, &nb))
			{
				if(IsVarVal(pVarResult))
				{
					V_VT(pVarResult) = VT_BSTR;
					V_BSTR(pVarResult) = ::SysAllocString(pwUserName);
				}
			}
			delete pwUserName;
			break;
		}
		case emProcessID:
			if(IsVarVal(pVarResult))
			{
				wchar_t wbuf[64] = {0}; 
				wsprintfW(wbuf, L"%d : %d", ::GetCurrentProcessId(), ::GetCurrentThreadId());
				V_VT(pVarResult) = VT_BSTR;
				V_BSTR(pVarResult) = ::SysAllocString(wbuf);
			}
			break;
		case emVersion:
			if(IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_BSTR;
				V_BSTR(pVarResult) = ::SysAllocString(L"123.89");
			}
			break;
		case emDo:
			result = Do(parameters, pVarResult, pExcepInfo);
			break;
		//case emSetWndProc:
		//	WndProcNamespace::CSetWndProc *pcset = new WndProcNamespace::CSetWndProc();
		//	bool ret = pcset->Init(pars->rgvarg[1].pvarVal->pdispVal, get_BSTR(pars->rgvarg[0]));
		//	if(IsVarVal(pvarRetValue))
		//	{
		//		V_VT(pvarRetValue) = VT_BOOL;
		//		V_BOOL(pvarRetValue) = ret;
		//	}
		//	if(!ret)
		//		delete pcset;
		//	hr = S_OK;
		case emUnload:
			result = Unload(get_BSTR(parameters->rgvarg[0]));
			break;
		case emUpdateWindow:
			UpdateWindow(::GetActiveWindow());
			break;
		case emGetProperty:
			result = get_Property(parameters, pVarResult);
			break;
		case emGetRefCount:
			result = get_RefCount(parameters, pVarResult);
			break;
		case emCompStrings:
		{
			int cmpres = wcscmp(get_BSTR(parameters->rgvarg[1]), get_BSTR(parameters->rgvarg[0]));
			if (IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_INT;
				V_INT(pVarResult) = cmpres;
			}

			result = S_OK;
			break;
		}
		case emIsValueInList:
		{
			result = IsValueInList(parameters, pVarResult);
			break;
		}
		default:
			result = DISP_E_MEMBERNOTFOUND;
			break;
	}
	return result;
}

HRESULT CXFunctions::XGetMessage(DISPPARAMS *parameters, VARIANT *pVarResult)
{
	HWND hwnd = NULL;
	MSG msg;
	bool cbKey = false;

	while(::PeekMessage(&msg, hwnd, 0, 0, PM_REMOVE))
	{
		::TranslateMessage(&msg);
		::DispatchMessage(&msg);
		if(msg.message == WM_KEYDOWN && msg.wParam == 3)	//	Ctrl + Break
		{
			cbKey = true;
		}
	}
	if(IsVarVal(pVarResult))
	{
		V_VT(pVarResult) = VT_BOOL;
		V_BOOL(pVarResult) = cbKey;
	}

//	if(parameters->cArgs > 0)
//		hwnd = reinterpret_cast<HWND>(parameters->rgvarg[0].pvarVal->lVal);
//	while(::PeekMessage(&msg, hwnd, 0, 0, PM_REMOVE))
//	{
//		if(msg.message == WM_NCMOUSEMOVE)
//		{
//			int x = LOWORD(msg.lParam);
//			int y = HIWORD(msg.lParam);
//			
//		}
//		if(hwnd == msg.hwnd)
//		{
//			::TranslateMessage(&msg);
//			::DispatchMessage(&msg);
//			if(msg.message == WM_KEYDOWN && msg.wParam == 3)	//	Ctrl + Break
//			{
//				cbKey = true;
//			}
//		}
//		else
//		{
//			::TranslateMessage(&msg);
//			::DispatchMessage(&msg);
//			if(msg.message == WM_KEYDOWN && msg.wParam == 3)	//	Ctrl + Break
//			{
//				cbKey = true;
//			}
//			if(msg.message == WM_ENABLE)
//			{
//				long wParam;
//				long lParam;
//
//				wParam = msg.wParam;
//				lParam = msg.lParam;
//			}
//		}
//	}
//	if(hwnd)
//	{
//		POINT cp;
//		::GetCursorPos(&cp);
//		RECT rect;
//		::GetWindowRect(hwnd, &rect);
//		HCURSOR hcurs;
//		hcurs = ::GetCursor();
//		if(cp.x >= rect.left && cp.x <= rect.right && cp.y >= rect.top && cp.y <= rect.bottom)
//		{
////			::SetSystemCursor(hcurs, OCR_NORMAL);
//		}
//		else
//		{
////			::SetSystemCursor(hcurs, OCR_WAIT);
//		}
//		::UpdateWindow(hwnd);
//	}
//	if(IsVarVal(pVarResult))
//	{
//		V_VT(pVarResult) = VT_BOOL;
//		V_BOOL(pVarResult) = cbKey;
//	}
	return S_OK;
}

ULONG CXFunctions::MemSize()
{
//typedef struct _MEMORYSTATUS {
//DWORD dwLength;
//DWORD dwMemoryLoad;
//SIZE_T dwTotalPhys;
//SIZE_T dwAvailPhys;
//SIZE_T dwTotalPageFile;
//SIZE_T dwAvailPageFile;
//SIZE_T dwTotalVirtual;
//SIZE_T dwAvailVirtual;
//} MEMORYSTATUS, *LPMEMORYSTATUS;

	//SIZE_T msize = 0;
	//MEMORYSTATUS mst = {0};
	//::GlobalMemoryStatus(&mst);
	//HANDLE hd = ::GetCurrentProcess();
	DWORD curId = ::GetCurrentProcessId();
	PROCESS_MEMORY_COUNTERS pmc = {0};
	HANDLE hProcess = ::OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, curId);
	if(hProcess)
	{
		::GetProcessMemoryInfo(hProcess, &pmc, sizeof(PROCESS_MEMORY_COUNTERS));
		::CloseHandle(hProcess);
	}
	return (ULONG)(pmc.WorkingSetSize / 1024);
}
struct THREADPARAMETERS
{
	BSTR methodName;
	DISPPARAMS params;
	IStream *pStream;
};

void __cdecl StartThread(void *parameters)
{
	int i;
	HRESULT rc;
	DISPID dispId;
	VARIANT varResult = {0};
	EXCEPINFO exInfo = {0};
	UINT nArgErr = 0;
	THREADPARAMETERS *pPars = (THREADPARAMETERS*)parameters;
	IDispatch *pObject;
	static HANDLE mutex = NULL;

	DWORD processID = ::GetCurrentProcessId();
	DWORD threadID = ::GetCurrentThreadId();

	rc = CoGetInterfaceAndReleaseStream(pPars->pStream, IID_IDispatch, (void**)&pObject);
	if(rc != S_OK)
		return;
	rc = pObject->GetIDsOfNames(GUID_NULL, &pPars->methodName, 1, 0, &dispId);
	if(rc == S_OK)
	{
		VARIANT * vcounter = pPars->params.rgvarg + pPars->params.cArgs - 1;
		VariantInit(vcounter);
		vcounter->vt = VT_UI4;
		for(unsigned long j=0; ; j++)
		{
			vcounter->ulVal = j;
			rc = pObject->Invoke(dispId, GUID_NULL, 0, DISPATCH_METHOD, &pPars->params, &varResult, &exInfo, &nArgErr);
			if(rc!=S_OK)
				break;
			if(varResult.vt==VT_BOOL && varResult.boolVal==false)
				break;
		}
	}
	
	i = pObject->Release();
	i = pPars->pStream->Release();

	for(i=0; i < pPars->params.cArgs; i++)
	{
		VariantClear(pPars->params.rgvarg + i);
	}
	
	::SysFreeString(pPars->methodName);
	
	delete pPars;
	
	//_endthread();
}

HRESULT CXFunctions::ExecInNewThread(DISPPARAMS *pars, VARIANT *retValue)
{
	HRESULT rc;
	THREADPARAMETERS *tpars = new THREADPARAMETERS;
	ULONG nRef;
	DWORD i;

	int narg = pars->cArgs - 1;

	IDispatch *pObject = pars->rgvarg[narg].pvarVal->pdispVal;
	nRef = pObject->AddRef();
	IStream *pStream;
	rc = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pObject, &pStream);
	nRef = pObject->Release();
	if(rc != S_OK)
	{
	}
	pStream->AddRef();
	tpars->pStream = pStream;

	//DISPID dispid;
	//LPOLESTR createMethod = L"Create";
	//VARIANT varResult = {0};
	//EXCEPINFO exInfo = {0};
	//UINT nArgErr = 0;
	//rc = pObject->GetIDsOfNames(GUID_NULL, &createMethod, 1, 0, &dispid);
	//if(rc == S_OK)
	//{
	//	DISPPARAMS dpars = {0};
	//	rc = pObject->Invoke(dispid, GUID_NULL, 0, DISPATCH_METHOD, &dpars, &varResult, &exInfo, &nArgErr);
	//	if(rc != S_OK)
	//	{
	//		delete tpars;
	//		return rc;
	//	}
	//}

	narg--;
	tpars->methodName = ::SysAllocString(pars->rgvarg[narg].pvarVal->bstrVal);

	tpars->params.cNamedArgs = 0;
	tpars->params.rgdispidNamedArgs = NULL;
	tpars->params.cArgs = narg + 1;
	tpars->params.rgvarg = new VARIANTARG[tpars->params.cArgs + 1];
	i = 0;
	narg--;
	while(narg >= 0)
	{
		VariantInit(tpars->params.rgvarg + i);
		rc = VariantCopyInd(tpars->params.rgvarg + i, pars->rgvarg + i);
		narg--;
		i++;
	}
	
	uintptr_t id = _beginthread(StartThread, 0, tpars);
	//hThread = ::CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)StartThread, tpars, 0, &threadID);

	return S_OK;
}

HRESULT CXFunctions::SetMemory(DISPPARAMS *pars, VARIANT * varResult, EXCEPINFO * pExcepInfo)
{
	LPWSTR filename = (pars->rgvarg[pars->cArgs-1]).pvarVal->bstrVal;
	HANDLE hFile = ::CreateFileMappingW(INVALID_HANDLE_VALUE, NULL, PAGE_EXECUTE_READWRITE, 0, 1, filename);
	if(hFile)
	{
		LPVOID map = MapViewOfFile(hFile, FILE_MAP_WRITE, 0, 0, 0); 

		::CloseHandle(hFile);
		return S_OK;
	}
	else
	{
		DWORD dwError = ::GetLastError();
		set_EXCEPINFO(pExcepInfo, L"SetMemory", dwError);
		return DISP_E_EXCEPTION;
	}
}
HRESULT CXFunctions::Do(DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo)
{
	if(parameters->cArgs < 2)
		return DISP_E_BADPARAMCOUNT;
	
	LPOLESTR members;
	HRESULT rc;
	UINT nArgErr = 0;

	IDispatch * processor; 
	DISPID  dispidRead = 0;
	VARIANT varIsBreak = {0};
	DISPPARAMS readPars = {0};

	processor = parameters->rgvarg[1].pvarVal->pdispVal;
	processor->AddRef();
	members = get_BSTR(parameters->rgvarg[0]);

	rc = processor->GetIDsOfNames(IID_NULL, &members, 1, LOCALE_SYSTEM_DEFAULT, &dispidRead); 
	if(rc==S_OK)
	{
		readPars.cArgs = 1;
		readPars.rgvarg = new VARIANT[1];
		V_VT(readPars.rgvarg) = VT_UI4;
		for (UINT i=0; ; i++)
		{
			V_UI4(readPars.rgvarg) = i;
			rc = processor->Invoke(dispidRead, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &readPars, &varIsBreak, pExcepInfo, &nArgErr);
			if(rc != S_OK)
				break;
			if(varIsBreak.vt==VT_I4)
			{
				if(varIsBreak.intVal < 0)
					break;
				if(varIsBreak.intVal > 0)
				{
					::Sleep(varIsBreak.intVal);
				}
			}
		}
		VariantClear(readPars.rgvarg);
		delete readPars.rgvarg;
	}
	processor->Release();
	return rc;
}
HRESULT CXFunctions::Unload(BSTR dllname)
{
	HRESULT (__stdcall* pDllCanUnloadNow)();
	HRESULT hr;

	HMODULE hModule = ::GetModuleHandleW(dllname);
	if(hModule)
	{
		pDllCanUnloadNow = (HRESULT (__stdcall*)())GetProcAddress(hModule, "DllCanUnloadNow");
		if(pDllCanUnloadNow)
		{
			hr = pDllCanUnloadNow();
			if(hr == S_OK)
				::FreeLibrary(hModule);
		}
	}
	return S_OK;
}

