#include "..\XCom.h"
#include "Lock.h"
#include "winternl.h"

CMethod MutexMethods[] = 
{ 
	CMethod(L"Lock", 1, 1, true), 
	CMethod(L"Unlock", 2, 0, false),
	CMethod(L"CloseMutex", 3, 1, true),
	CMethod(L"FileInformation", 4, 1, true)
};

// функция создания объекта
STDAPI Lock(LPVOID *ppv)
{
	CLock *lock = new CLock();
	lock->AddRef();
	*ppv = (void*)lock;
	return S_OK;
}

CLock::CLock(void)
{
	m_pComMethods = MutexMethods;
	m_nComMethodCount = 4;

	m_hMutex = NULL;
}

CLock::~CLock(void)
{
	if(m_hMutex)
	{
		::ReleaseMutex(m_hMutex);
		::CloseHandle(m_hMutex);
		m_hMutex = NULL;
	}
}

HRESULT CLock::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT rc = S_FALSE;

	if(nf==1)
	{
		if(parameters->cArgs < 1)
			rc = DISP_E_BADPARAMCOUNT;
		else
		{
			m_hMutex = ::CreateMutexW(NULL, TRUE, get_BSTR(parameters->rgvarg[parameters->cArgs - 1]));
			DWORD dwErr = ::GetLastError();
			if(dwErr == ERROR_ALREADY_EXISTS)
			{
				if(parameters->cArgs < 2)
					dwErr = ::WaitForSingleObject(m_hMutex, INFINITE);
				else
					dwErr = ::WaitForSingleObject(m_hMutex, parameters->rgvarg[parameters->cArgs - 2].pvarVal->lVal);
				if(IsVarVal(pVarResult))
				{
					V_VT(pVarResult) = VT_BOOL;
					if (dwErr == WAIT_OBJECT_0)
						V_BOOL(pVarResult) = TRUE;
					else 
						V_BOOL(pVarResult) = FALSE;
				}
			}
			else
			{
				if(IsVarVal(pVarResult))
				{
					V_VT(pVarResult) = VT_BOOL;
					V_BOOL(pVarResult) = TRUE;
				}
			}
			rc = S_OK;
		}
	}
	else if(nf == 2)
	{
		::ReleaseMutex(m_hMutex);
		::CloseHandle(m_hMutex);
		m_hMutex = NULL;
		rc = S_OK;
	}
	else if (nf == 3)		// closeMutex
	{
		HANDLE hMutex;

		hMutex = ::OpenMutexW(MUTEX_ALL_ACCESS, TRUE, get_BSTR(parameters->rgvarg[parameters->cArgs - 1]));
		DWORD dwErr = ::GetLastError();
		if (hMutex)
		{
			::ReleaseMutex(hMutex);
			::CloseHandle(hMutex);
			if (IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_BOOL;
				V_BOOL(pVarResult) = TRUE;
			}
		}
		else
		{
			if (IsVarVal(pVarResult))
			{
				V_VT(pVarResult) = VT_BOOL;
				V_BOOL(pVarResult) = FALSE;
			}
		}
		rc = S_OK;
	}
	else if (nf == 4)
	{
		FileInformation(get_BSTR(parameters->rgvarg[parameters->cArgs - 1]));
		rc = S_OK;
	}
	return rc;
}



int CLock::FileInformation(BSTR filename)
{
	
	return 0;
}
