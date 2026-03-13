#include "stdafx.h"
#include "Loader.h"
#include "XCom.h"

enum ENUM_METHOD_NUMBERS
{
    emLoadXCom = 1,
	emAttachAddin,
};
struct METHOD_DESCRIPTION
{
	wchar_t * methodName;
	int nMethod;
} MethodsList [] = 
{
	L"LoadXCom",	emLoadXCom,
	L"AttachAddin", emAttachAddin,
};

EXTERN_C HRESULT __stdcall GetLoader(void**ppv)
{
	CLoader * pLoader = new CLoader();
	pLoader->AddRef();
	*ppv = (void*)pLoader;
	return S_OK;
}
CLoader::CLoader(void)
{
	//Increment();
	m_nRef = 0;
}

CLoader::~CLoader(void)
{
	//Decrement();
}

HRESULT CLoader::ExcepInfo(DWORD rc, EXCEPINFO * pInfo)
{
	wchar_t * errbuf = new wchar_t [1024];
	DWORD ln = FormatMessageW(	FORMAT_MESSAGE_FROM_SYSTEM, 
								NULL, 
								rc, 
								MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
								errbuf, 
								1024, 
								NULL);
	if(ln > 0)
		pInfo->bstrDescription = ::SysAllocString(errbuf);
	else
		pInfo->bstrDescription = ::SysAllocString(L"Unknown error");
	delete errbuf;

	pInfo->wCode = rc;
	pInfo->scode = rc;
	pInfo->bstrSource = ::SysAllocString(L"Loader");

	return (HRESULT)rc;

}

HRESULT CLoader::LoadXCom(BSTR bstrDllname, BSTR bstrUUID, void **pObject)
{
	HRESULT ret;
	HMODULE hDll;
	BOOL isLoaded = FALSE;
	IID riid = {0};
	
	hDll = ::GetModuleHandleW(bstrDllname);
	if(!hDll)
	{
		LONG i = wcslen(bstrDllname) - 1;
		while(i>=0 && bstrDllname[i]!=L'\\') 
			i--;
		if(i < 0)
		{
			wchar_t *moduleName = new wchar_t[1024];
			i = ::GetModuleFileNameW(get_HINSTANCE(), moduleName, 1024);
			while(i>=0 && moduleName[i]!=L'\\') 
				i--;
			wmemmove(moduleName + i + 1, bstrDllname, wcslen(bstrDllname) + 1);
			hDll = ::LoadLibraryW(moduleName);
			
			delete moduleName;
		}
		else
			hDll = ::LoadLibraryW(bstrDllname);
		if(!hDll)
		{
			ret = GetLastError();
			return ret;
		}
		else
			isLoaded = TRUE;
	}
	ret = CreateObject(hDll, bstrUUID, riid, pObject);
	if(ret != S_OK)
	{
		if(isLoaded)
			::FreeLibrary(hDll);
		return ret;
	}
	return ret;
}

HRESULT CLoader::TypeInfo(BSTR dllname, CLSID & clsid)
{
	ITypeLib *ptlib;
	HRESULT hr = LoadTypeLib(dllname, &ptlib);
	if(SUCCEEDED(hr))
	{
		UINT imax = ptlib->GetTypeInfoCount();
		for(UINT i=0; i <= imax; i++ )
		{
			TYPEKIND tk;
			hr = ptlib->GetTypeInfoType(i,&tk);
			if(SUCCEEDED(hr) && tk==TKIND_COCLASS)
			{
				ITypeInfo* pti;
				hr = ptlib->GetTypeInfo(i,&pti);
				if( SUCCEEDED(hr) )
				{
					TYPEATTR* pta;
					hr = pti->GetTypeAttr(&pta);
					if(SUCCEEDED(hr))
					{
						clsid = pta->guid;
						i = imax + 1;
			// pta->guid - Это и есть искомый CLSID для REFCLSID

			//            hr = CoCreateInstance(pta->guid,NULL,CLSCTX_INPROC_SERVER, IID_IUnknown, (void**) &pObj);

			// или        hr = pti->CreateInstance(NULL, IID_IUnknown, (void**) &pObj);

			// а дальше начинается самое-самое. Объект-то ты в любом случае создашь.

			// НО: интерфейсы использовать, особенно, если ты их не знаешь, это очень и очень сложно.

			// Хотя через ITypeInfo все возможно.
						pti->ReleaseTypeAttr(pta);
					}
					pti->Release();
				}
			}
		}
		ptlib->Release();
	}
	return hr;
}

HRESULT CLoader::CreateObject(HMODULE hModule, BSTR classID, CLSID &riid, void ** ppv)
{
	HRESULT hResult = S_OK;
	IClassFactory * pFactory = NULL;
	HRESULT (__stdcall *pCreateObject)(LPVOID *ppv);

	//if(classID == NULL || classID[0]==L'\0')
	//{
	//	wchar_t * wDllname = new wchar_t[2048];
	//	::GetModuleFileNameW(hModule, wDllname, 2048);
	//	hResult = TypeInfo(wDllname, riid);
	//	delete wDllname;
	//	if(hResult==S_OK)
	//		hResult = XInit::CreateInstance(hModule, riid, IID_IUnknown, ppv);
	//}
	//else if(classID[0] == L'{') 
	//{
	//	hResult = ::CLSIDFromString(classID, &riid);
	//	if(hResult==S_OK)
	//		hResult = XInit::CreateInstance(hModule, riid, IID_IUnknown, ppv);
	//}
	//else
	//{
	//	char *prcName = new char[2048];
	//	int lPrcName = ::WideCharToMultiByte(CP_ACP, 0, classID, -1, prcName, 2048, NULL, NULL);
	//	if(lPrcName > 0)
	//	{
	//		pCreateObject = (HRESULT (__stdcall*)(LPVOID*))::GetProcAddress(hModule, prcName);
	//		if(!pCreateObject)
	//			hResult = ::GetLastError();
	//		else
	//			hResult = pCreateObject(ppv);
	//	}
	//	else 
	//		hResult = GetLastError();
	//	delete prcName;
	//}
	return hResult;
}

HRESULT CLoader::GetIDsOfNames(const IID &, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
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
HRESULT CLoader::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT hResult = S_FALSE;

	if(nf == emLoadXCom)
	{
		BSTR dllname = NULL;
		BSTR classID = NULL;
		void *pObject;

		if(parameters->cArgs > 0)
			dllname = get_BSTR(parameters->rgvarg[parameters->cArgs-1]);
		if(parameters->cArgs > 1)
			classID = get_BSTR(parameters->rgvarg[parameters->cArgs-2]);
		hResult = LoadXCom(dllname, classID, &pObject);
		if(hResult == S_OK)
		{
			if(IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_UNKNOWN;
				V_UNKNOWN(pVarResult) = (IUnknown*)pObject;
			}
		}
		else
		{
			ExcepInfo(hResult, pExcepInfo);
			hResult = DISP_E_EXCEPTION;
		}
	}
	else if(nf == emAttachAddin)
	{
		hResult = S_OK;
	}
	return hResult;
}

