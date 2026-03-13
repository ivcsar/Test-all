#include "XCom.h"
#include "LogHook.h"
#include "Thunk\Replacement.h"
#include "winternl.h"

using namespace Replacement;

STDAPI GetLogHook(LPVOID *ppv)
{
	CLogHook *pLogHook = new CLogHook();
	*ppv = (void*)pLogHook;
	return S_OK;
}

#pragma region CLogHook IDispatch
HRESULT __stdcall CLogHook::QueryInterface(const IID &riid, void **ppv)
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
ULONG __stdcall CLogHook::AddRef(void)
{
	IncrementRef();
	m_nRef++;
	return m_nRef;
}
ULONG __stdcall CLogHook::Release(void)
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
HRESULT __stdcall CLogHook::GetTypeInfoCount(UINT *pInfoCount)
{
	return E_NOTIMPL;
}
HRESULT __stdcall CLogHook::GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo)
{
	return E_NOTIMPL;
}
HRESULT CLogHook::GetIDsOfNames(const IID &riid, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid)
{
	HRESULT rc = DISP_E_MEMBERNOTFOUND;

	if(_wcsicmp(methodName[0], L"SetHook") == 0)
	{
		*dispid = 1;
		rc = S_OK;
	}
	return rc;
}

HRESULT CLogHook::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT rc = S_FALSE;

	if(nf == 1)
	{
		SetHook();
		rc = S_OK;
	}
	return rc;
}
#pragma endregion

CLogHook::CLogHook()
{
	m_replacement = NULL;
	m_nRef = 0;
	AddRef();
}

CLogHook::~CLogHook()
{
	REPLACEMENT * repl;
	DWORD protectAttr = 0;
	
	if(m_replacement)
	{
		m_replacement->MoveFirst();
		while((repl = m_replacement->Next()))
		{
			VirtualProtect(repl->m_funplace, sizeof(void*), PAGE_EXECUTE_READWRITE, &protectAttr);
			memmove(repl->m_funplace, &repl->m_oldptr, sizeof(void*));
			VirtualProtect(repl->m_funplace, sizeof(void*), protectAttr, &protectAttr);
		}
		delete m_replacement;
		m_replacement = NULL;
	}
}
BOOL WINAPI WriteFile_Stub(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten, LPOVERLAPPED lpOverlapped)
{
	return WriteFile(hFile, lpBuffer, nNumberOfBytesToWrite, lpNumberOfBytesWritten, lpOverlapped);
}
int fprintf_Stub(FILE *fp, const char *format, va_list alist)
{
	return fprintf(fp, format, alist);
}
HFILE OpenFile_Stub(LPCSTR lpFileName, LPOFSTRUCT lpOfstruct, UINT uiStyle)
{
	return OpenFile(lpFileName, lpOfstruct, uiStyle);
}
int CLogHook::ReplaceWrite(LPCSTR c_dllname, LPCSTR win_dllname, void *pLogHook)
{
	void *faddr = ReplaceFunctionAddr(c_dllname, win_dllname, "WriteFile",  WriteFile_Stub);
	if(faddr)
		((CLogHook*)pLogHook)->m_replacement->Add((REPLACEMENT*)faddr);
	
	faddr = ReplaceFunctionAddr(c_dllname, win_dllname, "OpenFile", OpenFile_Stub);
	if(faddr)
		((CLogHook*)pLogHook)->m_replacement->Add((REPLACEMENT*)faddr);

	return 0;
}
typedef struct HandleInfo{
    ULONG Pid;
    USHORT  ObjectType;
    USHORT  HandleValue;
    PVOID ObjectPointer;
    ULONG AccessMask;
} HANDLEINFO, *PHANDLEINFO;

typedef struct SystemHandleInfo {
    ULONG nHandleEntries;
    HANDLEINFO HandleInfo[1];
} SYSTEMHANDLEINFO, *PSYSTEMHANDLEINFO;

void CLogHook::SetHook()
{
	NTSTATUS (WINAPI * pNtQuerySystemInformation)(
		SYSTEM_INFORMATION_CLASS SystemInformationClass,
		PVOID SystemInformation,
		ULONG SystemInformationLength,
		PULONG ReturnLength);
	PSYSTEMHANDLEINFO pSystemHandleInfo;
	ULONG ulSystemInfoLength = 0;

	HINSTANCE hinstLib = LoadLibrary("ntdll");
    if(hinstLib) 
    { 
        pNtQuerySystemInformation = (NTSTATUS (WINAPI*) (SYSTEM_INFORMATION_CLASS, PVOID, ULONG, PULONG))GetProcAddress(hinstLib, "NtQuerySystemInformation"); 
		if(pNtQuerySystemInformation)
		{
			pSystemHandleInfo = new SYSTEMHANDLEINFO[1000];
			pNtQuerySystemInformation((SYSTEM_INFORMATION_CLASS)16,
										pSystemHandleInfo,
										1000 * sizeof(SYSTEMHANDLEINFO),
										&ulSystemInfoLength);
			SYSTEMHANDLEINFO *pInfo = pSystemHandleInfo;
			for(unsigned long i=0; i < pSystemHandleInfo->nHandleEntries; i++)
			{
				USHORT handle = pInfo->HandleInfo[i].HandleValue;
			}
		}
		::FreeLibrary(hinstLib);
	}
	m_replacement = new TList<REPLACEMENT>();
	EnumDlls(CLogHook::ReplaceWrite, (void*)this, NULL);
}