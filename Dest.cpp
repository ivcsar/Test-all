#include "Dest.h"

STDAPI GetDest(LPVOID *ppv)
{
	CDest *pDest = new CDest;
	*ppv = (void*)pDest;
	return S_OK;
}

CDest::CDest(void)
{
	AddRef();
	m_str = NULL;
}

CDest::~CDest(void)
{
	if(m_str)
		::SysFreeString(m_str);
}

HRESULT CDest::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	return S_OK;
}

//---------------------------------------------------------------------------//
bool CDest::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
    return AllocWSTR(wsExtensionName, L"Dest");
}
//---------------------------------------------------------------------------//
bool CDest::CallAsProc(	const long lMethodNum,
						tVariant *paParams, 
						const long lSizeArray)
{ 
    return true;
}

//---------------------------------------------------------------------------
bool CDest::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    return true; 
}
