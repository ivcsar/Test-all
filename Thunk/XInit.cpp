#include "..\stdafx.h"
#include "..\XCom.h"
#include "..\TList.h"
#include "XInit.h"

using namespace XInit;
using namespace Replacement;
TList <CXInit> gl_InitList;

#pragma region CXInit

void CXInit::Init()
{
	GetCXInit() = this;
	m_wModulename = NULL;
	m_classID = NULL;
	m_wClassID = NULL;
	m_registration = vrg_XCom;
	m_pCoClassList = NULL;

	m_pFuncReplace = new CFuncReplace();
	m_pFuncReplace->Add(L"CLSIDFromProgID",  &CLSIDFromProgID_Wrap);
	m_pFuncReplace->Add(L"CoCreateInstance", &CreateInstance_Wrap);
	m_pFuncReplace->Add(L"GetProcAddress",	 &GetProcAddress_Wrap);
	m_pFuncReplace->Add(L"CoFreeLibrary",	 &CoFreeLibrary_Wrap);		// чтобы ПодключитьВнешнююКомпоненту не FreeLibrary эту DLL
}
CXInit::~CXInit()
{
	if(m_wModulename)
		::SysFreeString(m_wModulename);
	if(m_classID)
		delete m_classID;
	if(m_wClassID)
		::SysFreeString(m_wClassID);
	if(m_pFuncReplace)
		delete m_pFuncReplace;
	if(m_pCoClassList)
		delete m_pCoClassList;
}

CXInit::CXInit(wchar_t *funcname)
{
	Init();
	m_registration = vrg_XCom;
	ClassID(funcname);
	m_pFuncReplace->Replace();
}

CXInit::CXInit(wchar_t *wszDllname, wchar_t *wszClassID)
{
	Init();
	m_registration = vrg_AttachAddin;
	m_wModulename = ::SysAllocString(wszDllname);

	if(wszClassID && wcslen(wszClassID) > 0)
	{
		m_pCoClassList = new CoClassList();
		CoClassInfo *pClassInfo = NULL;

		wchar_t *wszToken = wszClassID;
		wchar_t *wbuf = new wchar_t[2048];
		wmemset(wbuf, L'\0', 2048);
		UINT i = 0;
		while(*wszToken)
		{
			if(*wszToken == L'|')
			{
				pClassInfo = new CoClassInfo;
				pClassInfo->m_bstrClassName = ::SysAllocString(wbuf);
				m_pCoClassList->Add(pClassInfo);
				i = 0;
				wmemset(wbuf, L'\0', 2048);
			}
			else
			{
				wbuf[i] = *wszToken;
				i++;
			}
			wszToken++;
		}
		if(i > 0)
		{
			pClassInfo = new CoClassInfo;
			pClassInfo->m_bstrClassName = ::SysAllocString(wbuf);
			m_pCoClassList->Add(pClassInfo);
		}
		delete wbuf;
	}
	else		// нет списка классов
	{
		HRESULT hr;
		m_pCoClassList = new CoClassList();
		hr = XInit::TypeInfo(wszDllname, m_pCoClassList);
		if(m_pCoClassList->Count() > 0)	// вариант ПодключитьВнешнююКомпоненту()
		{
			m_pFuncReplace->Add(L"LoadStringW", &LoadStringW_Wrap);
		}
		else			// вариант new COMObject()
		{
			m_registration = vrg_NET;		// будем считать NET сборкой, имя класса будет передаваться в new COMObject
		}
	}
	m_pFuncReplace->Replace();
}
#pragma endregion

#pragma region static functions
// static
HRESULT XInit::TypeInfo(wchar_t *dllname, CoClassList *classList)
{

	ITypeLib *ptlib;
	HRESULT hr = LoadTypeLib(dllname, &ptlib);
	if(SUCCEEDED(hr))
	{
		UINT uiCount = ptlib->GetTypeInfoCount();
		for(UINT i=0; i <= uiCount; i++)
		{
			TYPEKIND tk;
			hr = ptlib->GetTypeInfoType(i,&tk);
			if(SUCCEEDED(hr) && tk==TKIND_COCLASS)
			{
				ITypeInfo* pti;
				hr = ptlib->GetTypeInfo(i, &pti);
				if(SUCCEEDED(hr))
				{
					TYPEATTR* pta;
					hr = pti->GetTypeAttr(&pta);
					if(SUCCEEDED(hr))
					{
						BSTR bstrName = NULL;
						BSTR bstrDoc = NULL;
						DWORD dwHelpContext = 0;
						BSTR bstrHelpFile = NULL;
						hr = ptlib->GetDocumentation(i, &bstrName, &bstrDoc, &dwHelpContext, &bstrHelpFile);  
						if(SUCCEEDED(hr))
						{
							CoClassInfo *pClassInfo = new CoClassInfo;
							pClassInfo->m_clsid = pta->guid;
							pClassInfo->m_bstrClassName = ::SysAllocString(bstrName);
							classList->Add(pClassInfo);
						}
						if(bstrName)
							::SysFreeString(bstrName);
						if(bstrDoc)
							::SysFreeString(bstrDoc);
						if(bstrHelpFile)
							::SysFreeString(bstrHelpFile);
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

HANDLE gl_hMutex;

BOOL XInit::LockInit()
{
	HANDLE hMutex = ::CreateMutexA(NULL, TRUE, "Global\\XInit_Replacement");
	DWORD rc = ::GetLastError();
	if(rc == ERROR_ALREADY_EXISTS)
	{
		rc = ::WaitForSingleObject(hMutex, 20000);
		if(rc == WAIT_TIMEOUT)
			return FALSE;
	}
	gl_hMutex = hMutex;
	return TRUE;
}

void XInit::UnlockInit()
{
	CXInit *pInit = CXInit::GetCXInit();
	delete pInit;
	CXInit::GetCXInit() = 0;
	::ReleaseMutex(gl_hMutex);
	::CloseHandle(gl_hMutex);
}

HRESULT XInit::Register(WCHAR *dllname, WCHAR *classID)
{
	HRESULT rc = S_FALSE;

	if(dllname && wcslen(dllname) > 0)
	{
		if(LockInit())
		{
			new CXInit(dllname, classID);
			rc = S_OK;
		}
	}
	else	// регистрируется эта библиотека
	{
		if(LockInit())
		{
			new XInit::CXInit(classID);
			rc = S_OK;
		}
	}

	return rc;
}

HRESULT XInit::CreateInstance(HMODULE hModule, const CLSID & clsid, REFIID riid, void ** ppv)
{
	HRESULT (__stdcall *pDllGetClassObject)(REFCLSID rclsid, REFIID riid, LPVOID *ppv);
	IClassFactory * pFactory = NULL;
	HRESULT hResult;

	pDllGetClassObject = (HRESULT (__stdcall*)(REFCLSID, REFIID, LPVOID*))GetProcAddress(hModule, "DllGetClassObject");
	if(!pDllGetClassObject)
	{
		hResult = GetLastError();
		return hResult;
	}
	hResult = pDllGetClassObject(clsid, IID_IClassFactory, (void**)&pFactory);
	if(hResult == S_OK)
	{
		pFactory->LockServer(true);
		hResult = pFactory->CreateInstance(NULL, riid, ppv);
		pFactory->LockServer(false);
		pFactory->Release();
	}
	return hResult;
}
#pragma endregion

void CXInit::ClassID(OLECHAR *wszClassID)
{
	if(m_classID && *m_classID != 0)
	{
	}
	else
	{
		size_t i = wcslen(wszClassID);
		while(i > 0 && wszClassID[i-1] != L'.')
			i--;

		int lc = ::WideCharToMultiByte(CP_ACP, 0, wszClassID + i, -1, NULL, 0, 0, NULL);
		if(lc > 0)
		{
			m_classID = new char[lc];
			::WideCharToMultiByte(CP_ACP, 0, wszClassID + i, -1, m_classID, lc, 0, NULL);
		}
	}
}

void CXInit::ClassCorrection(LPWSTR lpwBuffer)
{
	CoClassInfo *pClassInfo;

	bool isReg = FALSE;
	wchar_t *wszToken = lpwBuffer;
	wchar_t *wbuf = new wchar_t[2048];
	wmemset(wbuf, L'\0', 2048);
	UINT i = 0;
	while(*wszToken)
	{
		if(*wszToken == L'|')
		{
			pClassInfo = m_pCoClassList->Find(wbuf);
			if(pClassInfo)
			{
				pClassInfo->m_isReg = true;
				isReg = true;
			}
			i = 0;
			wmemset(wbuf, L'\0', 2048);
		}
		else
		{
			wbuf[i] = *wszToken;
			i++;
		}
		wszToken++;
	}
	if(i > 0)
	{
		pClassInfo = m_pCoClassList->Find(wbuf);
		if(pClassInfo)
		{
			pClassInfo->m_isReg = true;
			isReg = true;
		}
	}
	if(!isReg && m_pCoClassList->Count() == 1)
	{
		m_pCoClassList->MoveFirst();
		pClassInfo = m_pCoClassList->Next();
		if(pClassInfo)
		{
			::SysFreeString(pClassInfo->m_bstrClassName);
			pClassInfo->m_bstrClassName = ::SysAllocString(wbuf);
			pClassInfo->m_isReg = true;
		}
	}
	delete wbuf;

	m_pCoClassList->MoveFirst();
	while((pClassInfo=m_pCoClassList->Next()))
	{
		if(!pClassInfo->m_isReg)
			m_pCoClassList->Remove(pClassInfo);
	}
}

HRESULT CXInit::ClsidFromProgId(LPCOLESTR lpszProgID, LPCLSID pclsid)
{

	if(m_registration == XInit::vrg_XCom)		// 
	{
		ClassID((OLECHAR*)lpszProgID);
		*pclsid = IID_XCom;
	}
	else if(m_registration == XInit::vrg_AttachAddin)
	{
		CoClassInfo *pClassInfo = NULL;
		pClassInfo = m_pCoClassList->Find((void*)(lpszProgID + 6));
		if(pClassInfo)
		{
			*pclsid = pClassInfo->m_clsid;
			m_pCoClassList->Remove(pClassInfo);
		}
	}
	else if(m_registration == XInit::vrg_NET)
	{
		m_wClassID = ::SysAllocString(lpszProgID);			
	}
	else
		return CO_E_CLASSSTRING;

	return S_OK;
}