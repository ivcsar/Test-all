#include "..\stdafx.h"
#include "..\XCom.h"
#include "stdio.h"

HRESULT get_Property(DISPPARAMS *pars, VARIANT *retValue)
{
	HRESULT rc;
	LPOLESTR method;
	DISPID dispid=0;
	VARIANT varResult = {0};
	EXCEPINFO exInfo = {0};
	UINT nArgErr = 0;
	DISPPARAMS dispars = { 0 };

	int narg = pars->cArgs - 1;

	IDispatch *pObject = pars->rgvarg[narg].pvarVal->pdispVal;
	method = get_BSTR(pars->rgvarg[narg - 1]);
	rc = pObject->GetIDsOfNames(GUID_NULL, &method, 1, 0, &dispid);
	if(rc == S_OK)
	{
		rc = pObject->Invoke(dispid, GUID_NULL, 0, DISPATCH_PROPERTYGET, &dispars, retValue, &exInfo, &nArgErr);
		if(rc != S_OK)
		{
			;
		}
	}
	return S_OK;
}
HRESULT get_RefCount(DISPPARAMS *pars, VARIANT *retValue)
{

	IDispatch *pObject = pars->rgvarg[pars->cArgs - 1].pvarVal->pdispVal;
	long nRefCount = pObject->AddRef();
	nRefCount = pObject->Release();
	if (IsVarVal(retValue))
	{
		V_VT(retValue) = VT_INT;
		V_INT(retValue) = nRefCount;
	}
	return S_OK;
}

HRESULT Execute(DISPPARAMS *pars, VARIANT *retValue)
{
	//HRESULT rc;
	//DWORD i;
	//LPOLESTR method;
	//DISPID dispid;
	//EXCEPINFO exInfo = {0};
	//UINT nArgErr = 0;
	//DISPPARAMS dispparams = { 0 };


	//int narg = pars->cArgs - 1;

	//IDispatch *pObject = pars->rgvarg[narg].pvarVal->pdispVal;
	//method = get_BSTR(pars->rgvarg[narg - 1]);
	//if (wcslen(method) == 0)
	//	return 0;

	////pObject->AddRef();

	//rc = pObject->GetIDsOfNames(GUID_NULL, &method, 1, 0, &dispid);
	//if(rc == S_OK)
	//{
	//	dispparams.cArgs = pars->cArgs - 2;
	//	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	//	for (i = 0; i < dispparams.cArgs; i++)
	//	{
	//		*(dispparams.rgvarg + i) = *(pars->rgvarg + i);
	//	}
	//	rc = pObject->Invoke(dispid, GUID_NULL, 0, DISPATCH_METHOD, &dispparams, retValue, &exInfo, &nArgErr);
	//	
	//	//pars->cArgs = pars->cArgs + 2;

	//	delete dispparams.rgvarg;

	//}

	////pObject->Release();

	HRESULT rc;
	LPOLESTR method;
	DISPID dispid;
	EXCEPINFO exInfo = { 0 };
	UINT nArgErr = 0;

	int narg = pars->cArgs - 1;

	IDispatch *pObject = pars->rgvarg[narg].pvarVal->pdispVal;
	method = get_BSTR(pars->rgvarg[narg - 1]);
	if (wcslen(method) == 0)
		return 0;

	rc = pObject->GetIDsOfNames(GUID_NULL, &method, 1, 0, &dispid);
	if (rc == S_OK)
	{
		pars->cArgs = pars->cArgs - 2;
		rc = pObject->Invoke(dispid, GUID_NULL, 0, DISPATCH_METHOD, pars, retValue, &exInfo, &nArgErr);
		//for (int i = 0; i < pars->cArgs; i++)
		//{
		//	if (pars->rgvarg[narg].pvarVal->vt == VT_DISPATCH)
		//	{
		//		ULONG nRef = pars->rgvarg[narg].pvarVal->pdispVal->Release();
		//		nRef;
		//	}
		//}
		pars->cArgs = pars->cArgs + 2;

	}
	
	return S_OK;
}

HRESULT IsValueInList(DISPPARAMS *pars, VARIANT *retValue)
{
	int i;
	int vt = (pars->rgvarg[pars->cArgs - 1].vt & 0xFF);
	int isRef = (pars->rgvarg[pars->cArgs - 1].vt & 0xFF00);

	if (IsVarVal(retValue))
	{
		V_VT(retValue) = VT_BOOL;
		V_BOOL(retValue) = FALSE;
	}

	if (vt == VT_BSTR)
	{

	}
	else
	{
		if (isRef)
		{
			LONG lVal = pars->rgvarg[pars->cArgs - 1].pvarVal->lVal;

			for (i = pars->cArgs - 2; i >= 0; i--)
			{
				if (pars->rgvarg[i].pvarVal->lVal == lVal)
				{
					if (IsVarVal(retValue))
					{
						V_VT(retValue) = VT_BOOL;
						V_BOOL(retValue) = TRUE;
					}
				}
			}
		}
		else
		{
			for (i = pars->cArgs - 2; i >= 0; i--)
			{
				if (pars->rgvarg[i].pintVal == pars->rgvarg[pars->cArgs - 1].pintVal)
				{

				}
			}
		}
	}

	return S_OK;
}