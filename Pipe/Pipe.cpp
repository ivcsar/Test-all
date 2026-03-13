#include "Pipe.h"
#include "process.h"
#include "..\XCom.h"

enum ENUM_METHOD_NUMBERS
{
	emWaitMessage=1,
	emSendMessage,
	emGetMessage, 
	emSleep,
};
CMethod PipeMethods [] =
{
	CMethod(L"WaitMessage", emWaitMessage, 1, false), 
	CMethod(L"SendMessage", emSendMessage, 3, false),
	CMethod(L"GetMessage",  emGetMessage,  0, true),
	CMethod(L"Sleep",		emSleep,	1, false),
}; 

// функция создания объекта
STDAPI GetPipe(LPVOID *ppv)
{
	CPipe *pipe = new CPipe();
	*ppv = (void*)pipe;
	return S_OK;
}

CPipe::CPipe(void)
{
	AddRef();
	m_pComMethods = PipeMethods;
	m_nComMethodCount = sizeof(PipeMethods) / sizeof(CMethod);

	//m_nMethods = m_nComMethodCount;
	//m_cMethods = m_pComMethods;

	m_wTestVersion = NULL;
	m_status = 1;
	m_hEvent = NULL;
	m_hPipeMutex = NULL;
	m_pipename = NULL;

	m_pStream = NULL;
	m_dispid = 0;

	m_message = NULL;

	InitializeCriticalSection(&m_CriticalSection);
}

CPipe::~CPipe(void)
{
	if(m_wTestVersion)
		::SysFreeString(m_wTestVersion);

	if(m_pipename)
		::SysFreeString(m_pipename);
	SetMessage(NULL, 0);
	SetStatus(0);
	if(m_hEvent)
	{
		::SetEvent(m_hEvent);
		::CloseHandle(m_hEvent);
	}
	if(m_hPipeMutex)
	{
		::WaitForSingleObject(m_hPipeMutex, INFINITE);
		::ReleaseMutex(m_hPipeMutex);
		::CloseHandle(m_hPipeMutex);
	}
	::DeleteCriticalSection(&m_CriticalSection);
}

/* ILanguageExtenderBase
//---------------------------------------------------------------------------//
bool CPipe::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
    return AllocWSTR(wsExtensionName, L"Pipe");
}
//---------------------------------------------------------------------------//
bool CPipe::GetParamDefValue(const long lMethodNum, 
							const long lParamNum,
							tVariant *pvarParamDefValue)
{ 
	bool result = false;
	if(lMethodNum == emSendMessage)
	{
		if(lParamNum == 0)
		{
			TV_VT(pvarParamDefValue)= VTYPE_PWSTR;
			TV_WSTR(pvarParamDefValue) = L"";
			result = true;
		}
	}

    return result;
} 
//---------------------------------------------------------------------------//
bool CPipe::CallAsProc(	const long lMethodNum,
						tVariant *paParams, 
						const long lSizeArray)
{ 
	if(lMethodNum == emWaitMessage)
	{
		WaitMessage(paParams[0].pwstrVal);
		return true;
	}
	else if(lMethodNum == emSendMessage)
	{
		SendPipeMessage(paParams[0].pwstrVal, paParams[1].pwstrVal, paParams[2].pwstrVal);
		return true;		
	}
	else if(lMethodNum == emSleep)
	{
		::Sleep(paParams[0].intVal);
		return true;
	}
    return false;
}

//---------------------------------------------------------------------------
bool CPipe::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    bool ret = false;

    if(lMethodNum == emGetMessage)
    {
		wchar_t *wmessage = GetPipeMessage();
		if(wmessage)
		{
			WCHAR_T *wstmp; 
			AllocWSTR(&wstmp, wmessage);
			SetWSTR(pvarRetValue, wstmp);
			::SysFreeString(wmessage);
		}
		else
		{
			TV_VT(pvarRetValue) = VTYPE_EMPTY;
		}
		ret = true;
    }
    return ret; 
}
*/

HRESULT CPipe::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT rc = S_FALSE;
	VARIANT var = {0};

	if(nf == emWaitMessage)
	{
		if(parameters->cArgs == 1)
		{
			WaitMessage(get_BSTR(parameters->rgvarg[0]));
			rc = S_OK;
		}
		else
			rc = DISP_E_BADPARAMCOUNT;
	}
	else if(nf== emSendMessage)
	{
		if(parameters->cArgs == 3)
		{
			bool rcb = SendPipeMessage(get_BSTR(parameters->rgvarg[2]), get_BSTR(parameters->rgvarg[1]), get_BSTR(parameters->rgvarg[0]), pExcepInfo);
			if(rcb)
				rc = S_OK;
			else
				rc = DISP_E_EXCEPTION;
		}
		else
			rc = DISP_E_BADPARAMCOUNT;
	}
	else if(nf == emSleep)
	{
		::Sleep(parameters->rgvarg[0].pvarVal->lVal);
		rc = S_OK;
	}
    else if(nf == emGetMessage)
    {
		wchar_t *wmessage = GetPipeMessage();
		if(wmessage)
		{
			if(IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_BSTR;
				V_BSTR(pVarResult) = wmessage;
			}
		}
		rc = S_OK;
    }

	return rc;
}
#pragma endregion

bool CPipe::SendPipeMessage(wchar_t *server, wchar_t *pipename, wchar_t *message, EXCEPINFO *pExcepInfo)
{
	HANDLE hPipe;
	wchar_t *filename;
	wchar_t *serv;

	if(server==NULL || *server == L'\0')
		serv = L".";
	else
		serv = server;
	filename = new wchar_t[1024];
	wsprintfW(filename, L"\\\\%s\\pipe\\%s", serv, pipename);

	hPipe = CreateFileW(filename, 				// pipe name 
			GENERIC_READ | GENERIC_WRITE,	// read and write access 
			0,								// no sharing 
			NULL,							// default security attributes
			OPEN_EXISTING,					// opens existing pipe 
			0,								// default attributes 
			NULL);							// no template file 
	// Break if the pipe handle is valid. 
	if (hPipe == INVALID_HANDLE_VALUE) 
	{
		DWORD erc = ::GetLastError();
		set_EXCEPINFO(pExcepInfo, NULL, erc);
		//this->CallException(erc, filename);
		delete [] filename;
		return true;
	}
	delete [] filename;
	
	// Send a message to the pipe server. 
	DWORD cbWritten = 0;

	BOOL rc = WriteFile(hPipe, message, wcslen(message) * sizeof(wchar_t), &cbWritten, NULL);
 
	CloseHandle(hPipe); 

	return true;
}

void WaitClientMessage(void *parameters)
{
	HANDLE hPipe;
	BOOL rc;
	DWORD cbRead;
	DWORD dwErr;

	BYTE *pBuf;
	BYTE *buffer;
	DWORD lBuf;
	DWORD ncount = 0;
	DWORD i;
	DWORD buffer_size = 4096;
	OVERLAPPED ovrs = {0};

	CPipe * serv = (CPipe*)parameters;
	::WaitForSingleObject(serv->m_hPipeMutex, INFINITE);		// захват

	ovrs.hEvent = serv->m_hEvent;

	hPipe = ::CreateNamedPipeW(serv->m_pipename, 
						PIPE_ACCESS_DUPLEX			// read/write access 
						| FILE_FLAG_OVERLAPPED,
						PIPE_TYPE_BYTE |			// message-type pipe 
						//PIPE_READMODE_BYTE |		// message-read mode 
						PIPE_WAIT,					// pipe mode
						1,							// PIPE_UNLIMITED_INSTANCES,
						buffer_size,				// out buffer size
						buffer_size,				// in buffer size
						0,							// default time out
						NULL);						// security attributes
	
	if(hPipe == INVALID_HANDLE_VALUE)
		return;  
	
	while(1)
	{
		rc = ConnectNamedPipe(hPipe, &ovrs);
		if(rc)
		{
		}
		else
		{
			dwErr = GetLastError();
			DWORD dwWait = 0;
			if(dwErr == ERROR_IO_PENDING)
			{
				dwWait = ::WaitForSingleObject(ovrs.hEvent, INFINITE);
				if(dwWait == WAIT_OBJECT_0)		// уничтожается класс Pipe
				{
					if(serv->GetStatus() == 0)
					{
						::DisconnectNamedPipe(hPipe);
						break;
					}
				}
			}
			else if(dwErr == ERROR_PIPE_CONNECTED)
			{
				::SetEvent(ovrs.hEvent);
			}
			else
			{
				continue;
			}
			ncount = 0;
			pBuf = new BYTE[buffer_size];
			lBuf = buffer_size;
			buffer = new BYTE[buffer_size];
			while (true) 
			{ 
				rc = ::ReadFile(hPipe, buffer, buffer_size, &cbRead, NULL);
				if(!rc || cbRead == 0)
				{
					break;
				}
				if(ncount + cbRead > lBuf)
				{
					BYTE * mem = new BYTE [ncount];
					memmove(mem, pBuf, ncount);
					delete pBuf;
					lBuf += buffer_size;  
					pBuf  = new BYTE [lBuf];
					memset(pBuf, 0, lBuf);
					memmove(pBuf, mem, ncount);
					delete mem;
				}
				for(i = 0; i < cbRead; i++)
				{
					pBuf[ncount] = buffer[i];
					ncount++;
				}
			}
			delete buffer;
			serv->SetMessage(pBuf, ncount);
			delete pBuf;
			::DisconnectNamedPipe(hPipe);
			::SetEvent(ovrs.hEvent);
		}		
	}
	CloseHandle(hPipe);

	ReleaseMutex(serv->m_hPipeMutex);
}
//void CPipe::CreateStreamObject(DISPPARAMS *pars)
//{
//	LONG nRef;
//	HRESULT rc;
//
//	if(pars->cArgs < 3)
//		return;
//
//	BSTR method = get_BSTR(pars->rgvarg[pars->cArgs - 3]);
//
//	IDispatch *pObject = pars->rgvarg[pars->cArgs - 2].pvarVal->pdispVal;
//	nRef = pObject->AddRef();
//	rc = pObject->GetIDsOfNames(IID_NULL, &method, 1, LOCALE_SYSTEM_DEFAULT, &m_dispid);
//	if(rc == S_OK)
//	{
//		rc = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pObject, &m_pStream);
//		nRef = pObject->Release();
//		m_pStream->AddRef();
//	}
//	pObject->Release();
//}
//
//BOOL CPipe::ExecuteStreamObject()
//{
//	HRESULT rc;
//
//	if(!m_pStream)
//		return FALSE;
//
//	IDispatch *pObject;
//	rc = CoGetInterfaceAndReleaseStream(m_pStream, IID_IDispatch, (void**)&pObject);
//	if(rc == S_OK)
//	{
//		DISPPARAMS pars = {0};
//		EXCEPINFO excepInfo = {0};
//		VARIANT varResult = {0};
//		UINT nArgErr = 0;
//		rc = pObject->Invoke(m_dispid, GUID_NULL, 0, DISPATCH_METHOD, &pars, &varResult, &excepInfo, &nArgErr);
//		pObject->Release();
//	}
//	return TRUE;
//}
HRESULT CPipe::WaitMessage(wchar_t *pipename)
{
	wchar_t * wstr;

	wstr = new wchar_t[1024];
	wsprintfW(wstr, L"\\\\.\\pipe\\%s", pipename);
	m_pipename = ::SysAllocString(wstr);
	delete wstr;
	//CreateStreamObject(parameters);
	m_hPipeMutex = CreateMutex(NULL, FALSE, NULL);
	m_hEvent = ::CreateEventW(NULL, TRUE, FALSE, NULL);

	_beginthread(WaitClientMessage, 0, this);

	return S_OK;
}

void CPipe::SetMessage(BYTE *buffer, DWORD ln)
{
	EnterCriticalSection(&m_CriticalSection);
	if(m_message)
	{
		::SysFreeString(m_message);
		m_message = 0;
	}
	if(ln > 0)
	{
		m_message = ::SysAllocStringByteLen((char*)buffer, ln);
	}
	::LeaveCriticalSection(&m_CriticalSection);
}

wchar_t * CPipe::GetPipeMessage()
{
	wchar_t *message = NULL;
	EnterCriticalSection(&m_CriticalSection);
	if(m_message)
	{
		//AllocWSTR(&message, m_message);
		message = ::SysAllocString(m_message);
	}
	::LeaveCriticalSection(&m_CriticalSection);
	return message;
}
DWORD CPipe::GetStatus()
{
	DWORD status;
	EnterCriticalSection(&m_CriticalSection);
	status = m_status;
	::LeaveCriticalSection(&m_CriticalSection);
	return status;
}
void CPipe::SetStatus(DWORD status)
{
	EnterCriticalSection(&m_CriticalSection);
	m_status = status;
	::LeaveCriticalSection(&m_CriticalSection);
}

void WaitEvent(void * parameters)
{
	//DWORD rc;
	//THREADPARAMETERS * tpars = (THREADPARAMETERS*)parameters;

	//rc = WaitForSingleObject(tpars->hEvent, INFINITE);
	//::CloseHandle(tpars->hEvent);

	//IDispatch *pObject;
	//rc = CoGetInterfaceAndReleaseStream(tpars->pStream, IID_IDispatch, (void**)&pObject);
	//if(rc == S_OK)
	//{
	//	DISPID dispid;
	//	rc = pObject->GetIDsOfNames(GUID_NULL, &tpars->method, 1, 0, &dispid);
	//	if(rc == S_OK)
	//	{
	//		DISPPARAMS pars = {0};
	//		EXCEPINFO excepInfo = {0};
	//		VARIANT varResult = {0};
	//		UINT nArgErr = 0;
	//		rc = pObject->Invoke(dispid, GUID_NULL, 0, DISPATCH_METHOD, &pars, &varResult, &excepInfo, &nArgErr);
	//	}
	//	pObject->Release();
	//}

	//delete tpars;
}

//HRESULT CPipe::AddEvent(DISPPARAMS *pars, VARIANT *var)
//{
//	//THREADPARAMETERS * tpars = new THREADPARAMETERS;
//
//	//tpars->pipename = ::SysAllocString(_bstr_t("Global\\") + get_BSTR(pars->rgvarg[pars->cArgs - 1]));
//	//IDispatch *pObject = pars->rgvarg[pars->cArgs - 2].pvarVal->pdispVal;
//	//IStream *pStream;
//	//DWORD rc = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, pObject, &pStream);
//	//tpars->pStream = pStream;
//	//tpars->method = ::SysAllocString(get_BSTR(pars->rgvarg[pars->cArgs - 3]));
//	//
//	//m_hEvent = ::CreateEventW(NULL, TRUE, FALSE, tpars->pipename);
//	//tpars->hEvent = m_hEvent;
//	//_beginthread(WaitEvent, 0, tpars);
//	
//	return S_OK;
//}

//HRESULT CPipe::CallEvent(DISPPARAMS *pars, VARIANT *var)
//{
//	//BOOL rc;
//	//if(pars->cArgs > 0)
//	//{
//	//	wchar_t * pipename = _bstr_t("Global\\") + get_BSTR(pars->rgvarg[pars->cArgs - 1]);
//	//	HANDLE hEvent = OpenEventW(EVENT_MODIFY_STATE | SYNCHRONIZE, FALSE, pipename);
//	//	rc = ::SetEvent(hEvent);		// to signaled
//	//	rc = ::CloseHandle(hEvent);
//	//}
//	//else
//	//{
//	//}
//
//	return S_OK;
//}