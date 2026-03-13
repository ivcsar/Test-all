#include "stdafx.h"
#include "..\XNative\NativeBase.h"
#include "XInitDone.h"
#include <Psapi.h>
#include "XCom.h"
#include "Thunk\XInit.h"
#include "Thunk\XCLR.h"

using namespace std;

CXInitDone * m_pInit = NULL;

HRESULT __stdcall GetInitDone(void ** ppv)
{
	CXInitDone *pInit = new CXInitDone();
	pInit->AddRef();
	*ppv = (void*)pInit;
	return S_OK;
}
HWND GetAppWnd()
{
	if (m_pInit)
	{
		m_pInit->InitWnd();
		return m_pInit->m_hwnd1C;
	}
	else
	{
		return NULL;
	}
}
BOOL __stdcall EnumWndFunction(HWND hwnd, LPARAM lParam)
{
	TCHAR wbuf[64] = { 0 };

	::GetClassName(hwnd, wbuf, sizeof(wbuf)-1);
	if (_stricmp(wbuf, "V8MDIClient") == 0)
	{
		TCHAR cbuf[1024];
		::GetWindowText(hwnd, cbuf, sizeof(cbuf) - 1);
		return FALSE;
	}
	else
		return TRUE;
}
BOOL CXInitDone::InitWnd()
{
	BOOL rc = FALSE;
	DWORD nRef;
	HWND hwnd;
	DWORD processID, currentPID;
	char *wbuf;

	//IExtWndsSupport *pExtWndsSupport;

	if (m_hwnd1C)
		return TRUE;

	//hwnd = GetTopWindow(NULL);
	hwnd = GetActiveWindow();
	//TCHAR wbuf[256];
	//::GetWindowText(hwnd, wbuf, sizeof(wbuf) - 1);
	//LPOLESTR methodname = L"WriteLogEvent";
	//DISPID dispId=0;
	//rc = m_lpDispatch1C->GetIDsOfNames(GUID_NULL, &methodname, 1, 0, &dispId);

	currentPID =::GetCurrentProcessId();
	::GetWindowThreadProcessId(hwnd, &processID);
	if (currentPID == processID)
	{
		wbuf = new char[256];
		::GetClassName(hwnd, wbuf, 256);
		if (_stricmp("V8TopLevelFrame", wbuf) == 0)
		{
			m_hwnd1C = hwnd;
			rc = TRUE;
		}
		delete wbuf;
	}

	//if (m_lpDispatch1C)
	//{
	//	rc = m_lpDispatch1C->QueryInterface(IID_IExtWndsSupport, (void **)&pExtWndsSupport);
	//	if (rc == S_OK)
	//	{
	//		m_hwnd1C = NULL;
	//		rc = pExtWndsSupport->GetAppMainFrame(&m_hwnd1C);
	//		if (rc!=S_OK)
	//		{
	//			rc = pExtWndsSupport->GetAppMDIFrame(&m_hwnd1C);
	//	//		if (!SUCCEEDED(rc) && m_hwnd1C==NULL)
	//	//			m_hwnd1C = ::GetActiveWindow();
	//		}
	//		nRef = pExtWndsSupport->Release();
	//	}
	//	
	//	return TRUE;
	//}
	return rc;
}
void WriteLog(char * textEvent, BOOL isModuleName)
{
	//char logname[1024] = {0};
	//char moduleName[256] = {0};

	//DWORD lModule = GetModuleFileName(get_HINSTANCE(), logname, 1024);
	//if(lModule == 0)
	//	return;
	//DWORD lPath = lModule - 1;
	//while(logname[lPath] != '\\')
	//{
	//	lPath--;
	//}
	//if(isModuleName)
	//	memmove(moduleName, logname + lPath + 1, lModule-lPath);
	//memmove(logname + lPath + 1, "xlog.txt", 9);
	//ofstream *of = new ofstream(logname, ios::app);
	//if(isModuleName)
	//	*of << ::GetCurrentProcessId() << " : " << ::GetCurrentThreadId() << " : " << textEvent << " : " << moduleName;
	//else
	//	*of << ::GetCurrentProcessId() << " : " << ::GetCurrentThreadId() << " : " << textEvent;
	//*of << std::endl;
	//of->close();
}

enum ENUM_METHOD_NUMBERS
{
    emLoadXCom = 0,
	emAttachAddin,
    emLoadAssembly,
	emUnloadAssemblyObject,
	emExecuteAssemblyCode,
	emLoadAssemblyObject,
	emVersion,
	emUnload,
	emTestError,
	emRegister,
	emGetCommandLine
};
struct CMethod MethodList [] =
{
	CMethod(L"loadxcom", emLoadXCom, 2, true), 
	CMethod(L"AttachAddin", emAttachAddin, 2, true),
	CMethod(L"loadassembly", emLoadAssembly, 1, true),
	CMethod(L"unloadassemblyobject", emUnloadAssemblyObject, 1, true),
	CMethod(L"executeassemblycode", emExecuteAssemblyCode, 3, true),
	CMethod(L"LoadAssemblyObject", emLoadAssemblyObject, 2, true),
	CMethod(L"Version", emVersion, 0, true),
	CMethod(L"Unload",  emUnload, 1, false),
	CMethod(L"TestError",  emTestError, 0, false),
	CMethod(L"Register", emRegister, 2, true),
	CMethod(L"GetCommandLine", emGetCommandLine, 0, true)
}; 

CXInitDone::CXInitDone()
{
	m_nRef = 0;
	m_pInit = this;
	m_hwnd1C = NULL;
	m_lpDispatch1C = NULL;
	m_Version = NULL;
	WriteLog("Create", TRUE);
}

CXInitDone::~CXInitDone()
{
	if(m_Version)
		::SysFreeString(m_Version);
	WriteLog("Delete", TRUE);
	m_pInit = 0;
}

#pragma region IInitDone
STDMETHODIMP CXInitDone::Init(IDispatch* pDispatch)
{
	IExtWndsSupport *pExtWndsSupport;
	HRESULT rc;
	DWORD nRef;
	
	m_lpDispatch1C = pDispatch;
	m_lpDispatch1C->AddRef();
	m_lpDispatch1C->QueryInterface(IID_IErrorLog, (void**)&m_iErrorLog);
	//InitWnd();
	
	return S_OK;
}

STDMETHODIMP CXInitDone::Done(void)
{	
	m_iErrorLog->Release();
	m_lpDispatch1C->Release();
	m_lpDispatch1C = NULL;
	m_hwnd1C = NULL;
	return S_OK;
}

STDMETHODIMP CXInitDone::GetInfo(SAFEARRAY **pInfo)
{
	VARIANT varInfo = {0};
	SAFEARRAYBOUND rgsabound[1];

	rgsabound[0].lLbound = 0;
	rgsabound[0].cElements = 1;

	long lIndex = 0;

	V_VT(&varInfo) = VT_I4;
	V_I4(&varInfo) = 1000;		// будет создаваться только один объект компоненты
	
	SAFEARRAY *pArray = reinterpret_cast<SAFEARRAY*>(*pInfo);				// ::SafeArrayCreate(VT_VARIANT, 1, rgsabound);

	::SafeArrayPutElement(pArray, &lIndex, &varInfo);
	VARIANT var;
	SafeArrayGetElement(pArray, &lIndex, &var);

	return S_OK;
}
#pragma endregion

#pragma region  ILanguageExtender
STDMETHODIMP CXInitDone::RegisterExtensionAs(BSTR *pExtensionName)
{
	*pExtensionName = SysAllocString(L"XInitDone");
	return S_OK;
}

STDMETHODIMP CXInitDone::GetNMethods(long *plMethods)
{
	*plMethods = sizeof(MethodList) / sizeof(CMethod);
	return S_OK;
}

STDMETHODIMP CXInitDone::FindMethod(BSTR bstrMethodName, long *plMethodNum)
{
	HRESULT rc = S_FALSE;
	
	size_t lwcz = wcslen(bstrMethodName); 
	CMethod *pm = MethodList;
	for(int i=0; i < sizeof(MethodList) / sizeof(CMethod); i++)
	{
		if (lwcz==pm->lw && _wcsicmp(bstrMethodName, pm->method)==0)
		{
			*plMethodNum = pm->id;
			rc = S_OK;
			break;
		}
		pm++;
	}
	return rc;
}
// заплатка
CMethod *GetMethodNum(long lMethodNum)
{
	CMethod * pm = MethodList;
	for(int i=0; i < sizeof(MethodList) / sizeof(CMethod); i++)
	{
		if (pm->id == lMethodNum) 
		{
			return pm;
		}
		pm++;
	}
	return NULL;
}
STDMETHODIMP CXInitDone::GetMethodName(long lMethodNum, long lMethodAlias, BSTR *pbstrMethodName)
{
	CMethod *pm = GetMethodNum(lMethodNum);
	*pbstrMethodName = ::SysAllocString(pm->method);
	return S_OK;
}

STDMETHODIMP CXInitDone::GetNParams(long lMethodNum, long *plParams)
{
	CMethod * pm = GetMethodNum(lMethodNum);
	if(pm)
	{
		*plParams = pm->nParams;
		return S_OK;
	}
	else
		return S_FALSE;
}
STDMETHODIMP CXInitDone::GetParamDefValue(long lMethodNum, long lParamNum, VARIANT *pvarParamDefValue)
{
	if(lMethodNum == emLoadXCom)
	{
		if(lParamNum == 0 || lParamNum == 1) 
		{
			V_VT(pvarParamDefValue) = VT_NULL;
			V_BSTR(pvarParamDefValue) = L"";
		}
	}
	else if(lMethodNum == emRegister)
	{
		if(lParamNum == 0 || lParamNum == 1) 
		{
			V_VT(pvarParamDefValue) = VT_NULL;
			//V_BSTR(pvarParamDefValue) = L"";
		}
	}
	return S_OK;
}
STDMETHODIMP CXInitDone::HasRetVal(long lMethodNum, BOOL *pboolRetValue)
{
	CMethod * pm = GetMethodNum(lMethodNum);
	if(pm)
	{
		*pboolRetValue = (BOOL)(pm->hr);
		return S_OK;
	}
	else
		return S_FALSE;
}
STDMETHODIMP CXInitDone::CallAsProc(long lMethodNum, SAFEARRAY **paParams)
{
	LONG cArgs = 0;
	VARIANT *varArgs;
	HRESULT hr = S_FALSE;
	
	cArgs = LoadParameters(lMethodNum, paParams, &varArgs);
	if(lMethodNum == emUnload)
	{
		Unload(get_BSTR(varArgs[0]));
		hr = S_OK;
	}
	else if(lMethodNum == emTestError)
	{
		ExcepInfoEvent(2, L"Тестирование интерфейса IErrorLog");
		hr = S_FALSE;
	}
	else if(lMethodNum == emRegister)
	{
		hr = XInit::Register(get_BSTR(varArgs[0]), get_BSTR(varArgs[1])); 
	}
	FreeParameters(varArgs, cArgs);
	return hr;
}

STDMETHODIMP CXInitDone::CallAsFunc(long lMethodNum, VARIANT *pvarRetValue, SAFEARRAY **paParams)
{
	VARIANT *varArgs;
	LONG cArgs = 0;
	HRESULT hr;
	
	cArgs = LoadParameters(lMethodNum, paParams, &varArgs);
	if(lMethodNum == emLoadXCom)	// LoadXCom
	{
		//DISPID dispid = 0;
		//OLECHAR *methodname = L"AttachAddin";
		//hr = this->m_lpDispatch1C->GetIDsOfNames(GUID_NULL, &methodname, 1, 0, &dispid);

		IUnknown *pObject;
		if (cArgs == 2)
		{
			hr = LoadXCom(get_BSTR(varArgs[0]), get_BSTR(varArgs[1]), (void**)&pObject);
		}
		else if (cArgs == 1)
		{
			hr = LoadXCom(L"", get_BSTR(varArgs[0]), (void**)&pObject);
		}

		if(hr == S_OK)
		{
			if(IsVarVal(pvarRetValue))
			{
				V_VT(pvarRetValue) = VT_UNKNOWN;
				V_UNKNOWN(pvarRetValue) = pObject;
			}
		}
		else
		{
			if(IsVarVal(pvarRetValue))
				VariantClear(pvarRetValue);
			hr = S_OK;
		}
	}
	else if(lMethodNum == emAttachAddin)
	{

		BOOL bRet;
		bRet = AttachAddin(get_BSTR(varArgs[0]), get_BSTR(varArgs[1]));
		if(IsVarVal(pvarRetValue))
		{
			V_VT(pvarRetValue) = VT_BOOL;
			V_BOOL(pvarRetValue) = bRet;
		}
		hr = S_OK;
	}
	else if(lMethodNum == emLoadAssembly)	// 
	{
		//hr = LoadAssembly(get_BSTR(varArgs[0]), pvarRetValue);
		//if(!SUCCEEDED(hr))
		//{
		//	ExcepInfoEvent(hr, ErrorDescription(hr));
		//	hr = S_OK;
		//}
	}
	else if(lMethodNum == emUnloadAssemblyObject)
	{
		hr = UnloadAssemblyObject(get_BSTR(varArgs[0]));
		if(hr == S_OK)
		{
			V_VT(pvarRetValue) = VT_BOOL;
			V_BOOL(pvarRetValue) = true;
			return S_OK;
		}
		else
		{
			V_VT(pvarRetValue) = VT_BOOL;
			V_BOOL(pvarRetValue) = false;
			return S_FALSE;
		}
		hr = S_OK;
	}
	else if(lMethodNum == emExecuteAssemblyCode)		// executeassemblyobject
	{
		//try
		//{
		//	result = ExecuteAssemblyCode(get_BSTR(pars->rgvarg[2]), get_BSTR(pars->rgvarg[1]), &pars->rgvarg[0], pExcepInfo, pvarRetValue);
		//	if(FAILED(result))
		//	{
		//		result = DISP_E_EXCEPTION;
		//	}
		//}
		//catch(int err)
		//{
		//	::MessageBox(NULL, "errr", NULL, 0);
		//}
		//if(FAILED(result))
		//{
		//	set_EXCEPINFO(pExcepInfo, NULL, result);
		//	return DISP_E_EXCEPTION;
		//}
		hr = S_OK;
	}
	else if(lMethodNum == emLoadAssemblyObject)
	{
		
		hr = LoadAssemblyObject(get_BSTR(varArgs[0]), get_BSTR(varArgs[1]), pvarRetValue);
		if(!SUCCEEDED(hr))
		{
			ExcepInfoEvent(hr, ErrorDescription(hr));
			hr = S_OK;
		}
	}
	else if (lMethodNum == emRegister)
	{
		hr = XInit::Register(get_BSTR(varArgs[0]), get_BSTR(varArgs[1]));
		if (hr == S_OK)
		{
			V_VT(pvarRetValue) = VT_BOOL;
			V_BOOL(pvarRetValue) = true;
			return S_OK;
		}
		else
		{
			V_VT(pvarRetValue) = VT_BOOL;
			V_BOOL(pvarRetValue) = false;
			return S_FALSE;
		}
	}
	else if(lMethodNum == emVersion)
	{
		if(IsVarVal(pvarRetValue))
		{
			V_VT(pvarRetValue) = VT_BSTR;
			wchar_t * moduleName = new wchar_t[1024];
			::GetModuleFileNameW(get_HINSTANCE(), moduleName, 1024);
			V_BSTR(pvarRetValue) = ::SysAllocString(moduleName);
			delete moduleName;
		}
		hr = S_OK;
	}
	else if (lMethodNum == emGetCommandLine)
	{
		wchar_t * wsCmdline = ::GetCommandLineW();
		if (IsVarVal(pvarRetValue))
		{
			V_VT(pvarRetValue) = VT_BSTR;
			V_BSTR(pvarRetValue) = ::SysAllocString(wsCmdline);
		}
	}

	FreeParameters(varArgs, cArgs);
	return S_OK;
}
#pragma endregion

LONG CXInitDone::LoadParameters(long nMethod, SAFEARRAY **parameters, VARIANT **varPars)
{
	LONG i, cArgs=0;
	SAFEARRAY *parArray = NULL;

	parArray = (SAFEARRAY*)(*parameters);
	if(parArray)
	{
		cArgs = parArray->rgsabound[0].cElements;
		*varPars = new VARIANT [cArgs];
		for(i = 0; i < cArgs; i++)
		{
			::SafeArrayGetElement(parArray, &i, (*varPars) + i);
		}
	}
	return cArgs;
}

void CXInitDone::FreeParameters(VARIANT *varPars, LONG cArgs)
{
	LONG i;

	if(cArgs)
	{
		for(i = 0; i < cArgs; i++)
		{
			VariantClear(varPars + i);
		}
		delete varPars;
	}
}

HRESULT CXInitDone::LoadXCom(BSTR bstrDllname, BSTR bstrUUID, void **pObject)
{
	HRESULT (__stdcall *get_Object)(void**);
	HRESULT ret = S_FALSE;
	HMODULE hLib;
	BOOL isLoaded = FALSE;
	IID riid = {0};
	
	if (wcslen((wchar_t*)bstrDllname) == 0)
	{
		//hLib = LoadLibraryEx(dllname, NULL, LOAD_IGNORE_CODE_AUTHZ_LEVEL);
		hLib = get_HINSTANCE();
	}
	else
	{
		hLib = ::GetModuleHandleW(bstrDllname);
		if(!hLib)
		{
			hLib = ::LoadLibraryW(bstrDllname);
			if(!hLib)
			{
				ret = GetLastError();
				ExcepInfoEvent(ret);
				return ret;
			}
			else
				isLoaded = TRUE;
		}
	}

	char *classID = NULL;
	int lc = ::WideCharToMultiByte(CP_ACP, 0, bstrUUID, -1, NULL, 0, 0, NULL);
	if(lc > 0)
	{
		classID = new char[lc + 1];
		::WideCharToMultiByte(CP_ACP, 0, bstrUUID, -1, classID, lc + 1, 0, NULL);
	}

	if(classID)
	{
		if(classID[0] == '{');
		else
		{
			get_Object = (HRESULT (__stdcall*)(void**)) GetProcAddress(hLib, classID);
			if(get_Object)
			{
				ret = get_Object(pObject);
			}
		}
		delete classID;
	}
	if(ret != S_OK)
	{
		if(isLoaded)
			::FreeLibrary(hLib);
		ExcepInfoEvent(ret);
	}
	return ret;
}

VOID CXInitDone::ExcepInfoEvent(DWORD hResult)
{
	HRESULT ret;
	EXCEPINFO excepInfo = {0};

	wchar_t * errbuf = new wchar_t [1024];
	DWORD ln = FormatMessageW(	FORMAT_MESSAGE_FROM_SYSTEM, 
								NULL, 
								hResult, 
								MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
								errbuf, 
								1024, 
								NULL);
	if(ln > 0)
		excepInfo.bstrDescription = ::SysAllocString(errbuf);
	else
		excepInfo.bstrDescription = ::SysAllocString(L"Unknown error");
	delete errbuf;

	excepInfo.wCode = hResult;
	excepInfo.scode = hResult;
	excepInfo.bstrSource = ::SysAllocString(L"XCom");

	//ULONG nRef = m_lpDispatch1C->AddRef();
	//IErrorLog * iErr = NULL;
	//HRESULT ret = m_lpDispatch1C->QueryInterface(IID_IErrorLog, (void**)&iErr);
	excepInfo.scode = excepInfo.wCode;	// выдается предупреждение Ошибка при выполнении метода компоненты
	//pExcepInfo.scode = 0;		// при нуле не выдается предупреждение
	//#define ADDIN_E_NONE	    1000
	//#define ADDIN_E_ORDINARY  1001
	//#define ADDIN_E_ATTENTION 1002
	//#define ADDIN_E_IMPORTANT 1003
	//#define ADDIN_E_VERY_IMPORTANT 1004
	//#define ADDIN_E_INFO 1005
	//#define ADDIN_E_FAIL 1006
	//#define ADDIN_E_MSGBOX_ATTENTION 1007
	//#define ADDIN_E_MSGBOX_INFO 1008
	//#define ADDIN_E_MSGBOX_FAIL 1009
	excepInfo.wCode = 1006;
	ret = m_iErrorLog->AddError(L"Error", &excepInfo);
	//nRef = m_lpDispatch1C->Release();
}

VOID CXInitDone::ExcepInfoEvent(HRESULT hResult, BSTR errorDescription)
{
	HRESULT ret;
	EXCEPINFO excepInfo = {0};
	excepInfo.bstrDescription = errorDescription;
	excepInfo.wCode = hResult;
	excepInfo.scode = hResult;
	excepInfo.bstrSource = ::SysAllocString(L"XCom");
	excepInfo.scode = excepInfo.wCode;	// выдается предупреждение Ошибка при выполнении метода компоненты
	excepInfo.scode = 0;				// при нуле не выдается предупреждение
	excepInfo.wCode = 1006;
	ret = m_iErrorLog->AddError(L"Error", &excepInfo);
}

BOOL CXInitDone::AttachAddin(BSTR dllname, BSTR bstrCLSID)
{
	//HRESULT hr;
	//IID riid = {0};

	//hr = TypeInfo(dllname, riid);
	//if(hr==S_OK)
	//{
	//	new XInit::CXInit(riid, dllname);
	//	return TRUE;
	//}
	//else
	//{
	//	ExcepInfoEvent(hr);
	//	return FALSE;
	//}

	return FALSE;
}

void CXInitDone::Unload(BSTR dllname)
{
	HMODULE hModule;
	HRESULT hr;
	HRESULT (__stdcall* pDllCanUnloadNow)();
	hModule = ::GetModuleHandleW(dllname);
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
	
	DWORD i;
	DWORD lBaseModule, lModule, lPath;
	char *baseModule = new char[1024];
	HMODULE *hModules = new HMODULE[500];
	char *moduleName = new char[1024];
	
	HANDLE hProcess = GetCurrentProcess();

	lBaseModule = ::GetModuleFileNameA(NULL, baseModule, 1024);
	WriteLog(baseModule, FALSE);
	lBaseModule--;
	while(baseModule[lBaseModule] != '\\')
	{
		lBaseModule--;
	}

	DWORD cbNeeded = 0;

	::EnumProcessModules(hProcess, hModules, sizeof(HMODULE[500]), &cbNeeded);
			
	for(i=0; i < cbNeeded / sizeof(HMODULE); i++)
	{
		if(hModules[i] == 0) 
			break;
		lModule = GetModuleFileName(hModules[i], moduleName, 1024);
		if(lModule == 0)
			break;
		lPath = lModule - 1;
		while(moduleName[lPath] != '\\')
		{
			lPath--;
		}
		WriteLog(moduleName + lPath + 1, FALSE);
	}
	
	delete moduleName;
	delete hModules;
	delete baseModule;
}

STDMETHODIMP CXInitDone::GetNProps(long *plProps)
{
	return 1;
}
STDMETHODIMP CXInitDone::FindProp(BSTR bstrPropName, long *plPropNum)
{
	*plPropNum = 1;
	return S_OK;
}
STDMETHODIMP CXInitDone::GetPropName(long lPropNum, long lPropAlias, BSTR *pbstrPropName)
{
	return S_FALSE;
}
STDMETHODIMP CXInitDone::GetPropVal(long lPropNum, VARIANT *pvarPropVal)
{
	if(lPropNum == 1)
	{
		V_VT(pvarPropVal) = VT_BSTR;
		V_BSTR(pvarPropVal) = ::SysAllocString(m_Version);
		return S_OK;
	}
	return S_FALSE;
}
STDMETHODIMP CXInitDone::SetPropVal(long lPropNum, VARIANT *varPropVal)
{
	if(lPropNum == 1)
	{
		if(m_Version)
			::SysFreeString(m_Version);
		m_Version = ::SysAllocString(get_BSTR(*varPropVal));
		return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CXInitDone::IsPropReadable(long lPropNum, BOOL *pboolPropRead)
{
	if(lPropNum==1)
	{
		*pboolPropRead=(BOOL)1;
		return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CXInitDone::IsPropWritable(long lPropNum, BOOL *pboolPropWrite)
{
	if(lPropNum == 1)
	{
		*pboolPropWrite = (BOOL)TRUE;
		return S_OK;
	}
	return S_FALSE;
}
/*
class XFactory : public IClassFactory
{
	LONG m_nRef;
public:
	XFactory()
	{
		m_nRef = 0;
	}
	ULONG __stdcall AddRef()
	{
		m_nRef++;
		return m_nRef;
	}
	ULONG __stdcall Release()
	{
		m_nRef--;
		if(m_nRef <= 0)
			delete this;
		return m_nRef;
	}
	HRESULT __stdcall LockServer(BOOL flock)
	{
		return S_OK;
	}
	HRESULT __stdcall CreateInstance(IUnknown *pUnk, const CLSID & riid, void ** ppv)
	{
		CXInitDone *pInitDone = new CXInitDone();
		pInitDone->AddRef();
		*ppv = (void*)pInitDone;
		return S_OK;
	}
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv)
	{
		return S_OK;
	}
};

HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID riid, LPVOID *ppv)
{
	XFactory * pFactory = new XFactory;
	pFactory->AddRef();
	*ppv = (void*)pFactory;
	return S_OK;
}*/

