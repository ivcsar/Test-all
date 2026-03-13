#include "..\stdafx.h"
#include "XCLR.h"

#include "Parser.h"
wchar_t *Domain_Name = L"SVU_Domain_name"; 

#define MSCOREELIB 
#ifdef MSCOREELIB
#include <metahost.h>
#include <mscoree.h>
#pragma comment(lib, "mscoree.lib")
#import "mscorlib.tlb" raw_interfaces_only				\
    high_property_prefixes("_get","_put","_putref")		\
    rename("ReportEvent", "InteropServices_ReportEvent")

using namespace mscorlib; 

void AssemblyList(_AppDomain * domain, LPWSTR snum);
HRESULT GetAssembly(_AppDomain * domain, _Assembly ** pAssembly);

HRESULT GetAppDomain(BSTR filename, _AppDomain ** pAppDomain)
{
	ICLRMetaHost *pMetaHost = NULL;
	HRESULT hr;
	WCHAR *pwVersion;
	DWORD dwLength = 0;
	DWORD nRef;

	//ICLRMetaHostPolicy *pMetaHostPolicy = NULL;
	//ICLRDebugging      *pCLRDebugging   = NULL;
	//hr = CLRCreateInstance (CLSID_CLRMetaHostPolicy, IID_ICLRMetaHostPolicy,(LPVOID*)&pMetaHostPolicy);
	//hr = CLRCreateInstance (CLSID_CLRDebugging, IID_ICLRDebugging,(LPVOID*)&pCLRDebugging);
	pwVersion = new WCHAR[256];
	hr = ::GetFileVersion(filename, pwVersion, 256, &dwLength);

	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (LPVOID*)&pMetaHost);
	if(SUCCEEDED(hr))
	{
		ICLRRuntimeInfo *pRuntimeInfo = NULL;
		hr = pMetaHost->GetRuntime(pwVersion, IID_PPV_ARGS(&pRuntimeInfo)); 
		if(SUCCEEDED(hr)) 
		{
			// Check if the specified runtime can be loaded into the process. This  
			// method will take into account other runtimes that may already be  
			// loaded into the process and set pbLoadable to TRUE if this runtime can  
			// be loaded in an in-process side-by-side fashion.  
			BOOL isLoadable; 
			hr = pRuntimeInfo->IsLoadable(&isLoadable); 
			if(SUCCEEDED(hr) && isLoadable)
			{
				ICorRuntimeHost * pCorRuntimeHost = NULL;
				hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pCorRuntimeHost)); 
				if(SUCCEEDED(hr))
				{
					pCorRuntimeHost->Start();
					IUnknown * pUnkDomain = NULL;
					hr = pCorRuntimeHost->GetDefaultDomain(&pUnkDomain); 
					if(SUCCEEDED(hr))
					{
						hr = pUnkDomain->QueryInterface(__uuidof(_AppDomain), (void**)pAppDomain); 
						nRef = pUnkDomain->Release();
					}
					pCorRuntimeHost->Stop();
					pCorRuntimeHost->Release();
				}
			}
			pRuntimeInfo->Release();
		}
		pMetaHost->Release();
	}
	delete pwVersion;

	return hr;
}

HRESULT LoadAssemblyObject(BSTR dllname, BSTR typeName, VARIANT * pVarObject)
{
	_AppDomain * pAppDomain = NULL;
	HRESULT hr;

	hr = GetAppDomain(dllname, &pAppDomain);
	if(SUCCEEDED(hr))
	{
		_ObjectHandle *pObjectHandle = NULL;
		hr = pAppDomain->CreateInstanceFrom(dllname, typeName, &pObjectHandle);
		if(hr == S_OK)
		{
			hr = pObjectHandle->Unwrap(pVarObject);
			pObjectHandle->Release();
		}
		pAppDomain->Release();
	}
	return hr;
}

HRESULT LoadAssembly(BSTR dllname, VARIANT * pAsm)
{
	_AppDomain * pAppDomain;
	HRESULT hr;
	
	hr = GetAppDomain(dllname, &pAppDomain);
	if(SUCCEEDED(hr))
	{
		_Assembly * pAsm = NULL;
		hr = pAppDomain->Load_2(dllname, &pAsm);
		if(SUCCEEDED(hr))
		{
			void * pObject = NULL;
			hr = pAsm->QueryInterface(IID_IDispatch, &pObject);
			if(hr == S_OK);
		}
		pAppDomain->Release();
	}
	return hr;
}

BSTR ErrorDescription(HRESULT hResult)
{
	BSTR errorString = NULL;
	wchar_t * errbuf = new wchar_t[1024];
	if(SUCCEEDED(::LoadStringRC(hResult, errbuf, 1024, 0)))
	{
		errorString = ::SysAllocString(errbuf);
	}
	else
	{
		errorString = ::SysAllocString(L"Unknown error");
	}
	delete errbuf;
	return errorString;
}

// не работает
HRESULT UnloadAssemblyObject(BSTR dllname)
{
	HRESULT result = S_OK;
	ULONG nRef;
	_AppDomain * pAppDomain = NULL;
	SAFEARRAY *pAssemblies;

	result = GetAppDomain(dllname, &pAppDomain);
	if (result == S_OK)
	{
		result = pAppDomain->GetAssemblies(&pAssemblies);
		if (result == S_OK)
		{ 
			for (LONG j = 0; j < pAssemblies->rgsabound[0].cElements; j++)
			{
				IUnknown * punk;
				::SafeArrayGetElement(pAssemblies, &j, &punk);

				_AssemblyPtr pAssembly;
				punk->QueryInterface(__uuidof(_Assembly), (void**)&pAssembly);
				nRef = punk->Release();

				BSTR location;
				pAssembly->get_Location(&location);
				//::MessageBoxW(NULL, location, dllname, 0);
				if (_wcsicmp(location, dllname) == 0)
				{
					while (pAssembly->Release());
					//HMODULE hModule = ::GetModuleHandleW(location);
					//::FreeLibrary(hModule);
				}
				else
					nRef = pAssembly->Release();
				::SysFreeString(location);
			}
			SafeArrayDestroyData(pAssemblies);
			SafeArrayDestroy(pAssemblies);
		}
		while(pAppDomain->Release());
	}

	//ICorRuntimeHost *pCorRuntimeHost = NULL;
	//IUnknown *pUnkDomain = NULL;
	//HDOMAINENUM hdc;
	//result = CorBindToRuntimeEx(NULL, L"wks", 
	//			STARTUP_LOADER_OPTIMIZATION_SINGLE_DOMAIN | STARTUP_CONCURRENT_GC, 
	//			CLSID_CorRuntimeHost, IID_ICorRuntimeHost, 
	//			(void **)&pCorRuntimeHost); 
	//
	//pCorRuntimeHost->Start();
	//result = pCorRuntimeHost->EnumDomains(&hdc);
	//if(result == S_OK)
	//{
	//	BSTR friendly_name = NULL;
	//	while(pCorRuntimeHost->NextDomain(hdc, &pUnkDomain) == S_OK)
	//	{
	//		result = pUnkDomain->QueryInterface(__uuidof(_AppDomain), (void**)&pAppDomain);
	//		if(result == S_OK)
	//		{
	//			result = pAppDomain->get_FriendlyName(&friendly_name);
	//			nRef = pAppDomain->Release();
	//			if(result == S_OK)
	//			{
	//				if(wcscmp(friendly_name, Domain_Name) == 0)
	//				{
	//					::SysFreeString(friendly_name);
	//					while(pUnkDomain->Release() > 0);
	//					result = pCorRuntimeHost->UnloadDomain(pUnkDomain);
	//					break;
	//				}
	//				::SysFreeString(friendly_name);
	//				friendly_name = NULL;
	//			}
	//		}
	//		nRef = pUnkDomain->Release();
	//	}
	//	pCorRuntimeHost->CloseEnum(hdc);
	//}
	//pCorRuntimeHost->Stop();
	//nRef = pCorRuntimeHost->Release();

	return result;
}

_TypePtr GetVariantType(VARIANT target)
{
	HRESULT rc;
	ULONG nRef;
	_TypePtr rtptr = NULL;

	_ObjectPtr pObject;
	rc = target.punkVal->QueryInterface(__uuidof(_Object), (void**)&pObject); // QI вариант дл€ iface объектов
	if(!FAILED(rc))
	{
		rc = pObject->GetType(&rtptr);
		if(pObject != target.punkVal)
			nRef = pObject->Release();
	}
	return rtptr;
}

HRESULT VariantInvoke(VARIANT target, LPWSTR methodname, mscorlib::BindingFlags invokeAttr, VARIANT * pRetVariant, int nArgs, ...)
{
	HRESULT rc;
	ULONG nRef;
	LONG i;
	
	_ObjectPtr pObject;
	rc = target.punkVal->QueryInterface(__uuidof(_Object), (void**)&pObject); // QI вариант дл€ iface объектов
	if(FAILED(rc))
	{
		return (rc);
	}
	else
	{
		_TypePtr pType;
		rc = pObject->GetType(&pType);
		if(FAILED(rc))
		{
			return rc;
		}
		else
		{
			va_list argList;
			va_start(argList, nArgs);
			
			VARIANT * vargs = new VARIANT[nArgs];
			VARIANT * cvargs = vargs;
			SAFEARRAY * args = ::SafeArrayCreateVector(VT_VARIANT, 0, nArgs);
			for(i=0, cvargs = vargs; i < nArgs; i++, cvargs++)
			{
				V_VT(cvargs) = VT_BSTR;
				V_BSTR(cvargs) = va_arg(argList, BSTR);
				::SafeArrayPutElement(args, &i, cvargs);
			}

			BSTR mname = ::SysAllocString(methodname);
			rc = pType->InvokeMember_3(mname, invokeAttr, NULL, target, args, pRetVariant);
			::SysFreeString(mname);
			delete vargs;
			::SafeArrayDestroy(args);
			nRef = pType->Release();
		}
		nRef = pObject->Release();
	}
	return rc;
}
LONG GetNRef(_TypePtr type)
{
	LONG nRef = type->AddRef();
	nRef = type->Release();
	return nRef;
}
LONG GetNRef(VARIANT vtype)
{
	LONG nRef = vtype.punkVal->AddRef();
	nRef = vtype.punkVal->Release();
	return nRef;
}

HRESULT ExecuteCode(_AssemblyPtr pSystem, VARIANT parameters, VARIANT provider, BSTR dynClassname, BSTR dynMethodname, BSTR dynCode, VARIANTARG * assemblies, VARIANT * varResult)
{
	HRESULT rc;
	ULONG nRef;
	LONG i;
	VARIANT refAssemblies = {0};
	VARIANT var = {0};

	rc = VariantInvoke(parameters, L"ReferencedAssemblies", BindingFlags_GetProperty, &refAssemblies, 0, NULL);
	if(!FAILED(rc))
	{
		if(assemblies != NULL)
		{
			if((assemblies->vt & VT_ARRAY)==VT_ARRAY && (assemblies->vt & VT_BSTR)==VT_BSTR)
			{
				for(i = 0; i < assemblies->parray->rgsabound[0].cElements; i++)
				{
					BSTR wbStr;
					if(SafeArrayGetElement(assemblies->parray, &i, &wbStr) == S_OK)
					{	
						rc = VariantInvoke(refAssemblies, L"Add", BindingFlags_InvokeMethod, &var, 1, wbStr);
					}
				}
			}
			else if(assemblies->vt == VT_BSTR)
			{
				rc = VariantInvoke(refAssemblies, L"Add", BindingFlags_InvokeMethod, &var, 1, assemblies->bstrVal);
			}
		}
		BSTR fullName;
		pSystem->get_FullName(&fullName);
		rc = VariantInvoke(refAssemblies, L"Add", BindingFlags_InvokeMethod, &var, 1, fullName);
		::SysFreeString(fullName);
	}
	return rc;
}

void AssemblyList(_AppDomain * domain, LPWSTR snum)
{
	HRESULT rc;
	SAFEARRAY * assemblies; 
	ULONG nRef;

	rc = domain->GetAssemblies(&assemblies);
	if(rc == S_OK)
	{
		for(LONG j=0; j < assemblies->rgsabound[0].cElements; j++)
		{
			IUnknown * punk;
			::SafeArrayGetElement(assemblies, &j, &punk);
			_AssemblyPtr pAssembly; 
			punk->QueryInterface(__uuidof(_Assembly), (void**)&pAssembly);
			nRef = punk->Release();
			BSTR codebase;
			pAssembly->get_CodeBase(&codebase);
			::MessageBoxW(NULL, codebase, snum, 0);
			::SysFreeString(codebase);
			nRef = pAssembly->Release();
		}
	}
}

HRESULT GetAssembly(_AppDomain * domain, _Assembly ** pAssembly)
{
	HRESULT rc;
	SAFEARRAY * assemblies; 
	ULONG nRef;
	_Assembly * pAsm;

	rc = domain->GetAssemblies(&assemblies);
	if(rc == S_OK)
	{
		for(LONG j=0; j < assemblies->rgsabound[0].cElements; j++)
		{
			IUnknown * punk;
			::SafeArrayGetElement(assemblies, &j, &punk);
			rc = punk->QueryInterface(__uuidof(_Assembly), (void**)&pAsm);
			if(rc == S_OK)
			{
				*pAssembly = pAsm;
				//BSTR fullname;
				////rc = pAsm->get_FullName(&fullname);
				//rc = pAsm->get_CodeBase(&fullname);
				//if(rc==S_OK)
				//	::SysFreeString(fullname);
			}
			nRef = punk->Release();
			break;
		}
	}
	::SafeArrayDestroy(assemblies);
	return rc;
}

HRESULT ExecuteAssemblyCode(BSTR classname, BSTR code, VARIANTARG *assemblies, EXCEPINFO * pExcepInfo, VARIANT * vResult)
{
	HRESULT result;
	LONG i, j;
    ICorRuntimeHost *pCorRuntimeHost = NULL; 
	ULONG nRef;
	//XCom.dll,RunAssemblyCode /RunCode=D:\VS\CPP\XCom\RunCode\bin\Debug\RunCode.dll  /Type = classname  /Code=1
    result = CorBindToRuntimeEx(NULL, L"wks", 
		STARTUP_LOADER_OPTIMIZATION_SINGLE_DOMAIN | STARTUP_CONCURRENT_GC, 
		CLSID_CorRuntimeHost, IID_ICorRuntimeHost, 
        (void **)&pCorRuntimeHost); 
	
	result = pCorRuntimeHost->Start(); 
	if(result != S_OK)
	{
		return result;
	}
    IUnknown *pUnkDomain = NULL; 
	result = pCorRuntimeHost->CreateDomain(Domain_Name, NULL, &pUnkDomain); 
	//result = pCorRuntimeHost->GetDefaultDomain(&pUnkDomain); 
	//result = pCorRuntimeHost->CurrentDomain(&pUnkDomain); 
	if(!FAILED(result))
	{
		_AppDomain *pAppDomain = NULL; 

		result = pUnkDomain->QueryInterface(__uuidof(_AppDomain), (void**)&pAppDomain); 
		if(!FAILED(result))
		{
			//AssemblyList(pAppDomain, L"1");
			LPWSTR sysDir = new WCHAR[1204];
			DWORD ld = 0;
			result = GetCORSystemDirectory(sysDir, 1024, &ld);
			result = GetCORVersion(sysDir, 1024, &ld);
			delete sysDir;

			_Assembly * system;
			_ObjectHandlePtr pCompileParameters; 
			BSTR sysDllname = _bstr_t("C:\\WINDOWS\\Microsoft.NET\\Framework\\v2.0.50727\\System.dll"); 
			result = pAppDomain->CreateInstanceFrom(sysDllname, _bstr_t("System.CodeDom.Compiler.CompilerParameters"), &pCompileParameters);
			if(!FAILED(result))
			{
				VARIANT parameters;
				result = pCompileParameters->Unwrap(&parameters);
				nRef = GetNRef(parameters);
					
				BSTR member; 
				SAFEARRAY *args;
				VARIANT cvar = {0};
				VARIANT rvar = {0};
				args = ::SafeArrayCreateVector(VT_VARIANT, 0, 0);
				_TypePtr tPtr = GetVariantType(parameters);
				result = tPtr->InvokeMember_3(_bstr_t("ReferencedAssemblies"), BindingFlags_GetProperty, NULL, parameters, args, &cvar);
				nRef = tPtr->Release();
				::SafeArrayDestroy(args);

				if(!FAILED(result))		// работаем с parameters
				{
					// добавление сборок в ReferencedAssemblies 
					if(assemblies != NULL)
					{
						member = _bstr_t("Add");
						tPtr = GetVariantType(cvar);
						args = SafeArrayCreateVector(VT_VARIANT, 0, 1);
						j = 0;		// индекс
						BSTR wbStr;
						if((assemblies->vt & VT_ARRAY)==VT_ARRAY && (assemblies->vt & VT_BSTR)==VT_BSTR)
						{
							for(i = 0; i < assemblies->parray->rgsabound[0].cElements; i++)
							{
								if(SafeArrayGetElement(assemblies->parray, &i, &wbStr) == S_OK)
								{	
									::SafeArrayPutElement(args, &j, wbStr);
									tPtr->InvokeMember_3(member, BindingFlags_InvokeMethod, NULL, cvar, args, &rvar);
								}
							}
						}
						else if(assemblies->vt == VT_BSTR)
						{
							::SafeArrayPutElement(args, &j, assemblies->bstrVal);
							tPtr->InvokeMember_3(member, BindingFlags_InvokeMethod, NULL, cvar, args, &rvar);
						}
						::SafeArrayDestroy(args);
						::SysFreeString(member);
						nRef = tPtr->Release();
					}
					nRef = cvar.punkVal->Release();
					VariantClear(&cvar);
					VariantClear(&rvar);
					// добавление сборок 

					// получим сборку System.dll
					pAppDomain->GetAssemblies(&args);
					IUnknown *punk;
					j = 1;
					::SafeArrayGetElement(args, &j, (void**)&punk);
					result = punk->QueryInterface(__uuidof(_Assembly), (void**)&system);
					nRef = punk->Release();
					::SafeArrayDestroy(args);
					// получили сборку System.dll

					if(!FAILED(result))
					{				
						// компилируем
						VARIANT provider = {0};
						result = system->CreateInstance(_bstr_t("Microsoft.CSharp.CSharpCodeProvider"), &provider);
						if(!FAILED(result))
						{
							args = SafeArrayCreateVector(VT_VARIANT, 0, 2);
							j = 0;
							::SafeArrayPutElement(args, &j, &parameters); 
							V_VT(&cvar) = VT_BSTR;
							V_BSTR(&cvar) = code;
							j++;
							::SafeArrayPutElement(args, &j, &cvar); 
							tPtr = GetVariantType(provider);
							VARIANT compileResults = {0};
							result = tPtr->InvokeMember_3(_bstr_t("CompileAssemblyFromSource"), BindingFlags_InvokeMethod, NULL, provider, args, &compileResults);
							nRef = tPtr->Release();
							::SafeArrayDestroy(args);
							
							if(!FAILED(result))
							{
								_TypePtr typeResults = GetVariantType(compileResults);
								VARIANT errors;
								result = typeResults->InvokeMember_3(_bstr_t("Errors"), BindingFlags_GetProperty, NULL, compileResults, NULL, &errors);
								if(!FAILED(result))
								{
									_TypePtr typeErrors = GetVariantType(errors);
									result = typeErrors->InvokeMember_3(_bstr_t("HasErrors"), BindingFlags_GetProperty, NULL, errors, NULL, &rvar);
									if(!FAILED(result))
									{
										if(rvar.vt == VT_BOOL && rvar.boolVal)
										{
											VARIANT vEnumerator;
											result = typeErrors->InvokeMember_3(_bstr_t("GetEnumerator"), BindingFlags_InvokeMethod, NULL, errors, NULL, &vEnumerator);
											if(!FAILED(result))
											{
												nRef = GetNRef(vEnumerator);
												mscorlib::IEnumeratorPtr enumerator;
												result = vEnumerator.punkVal->QueryInterface(__uuidof(mscorlib::IEnumerator), (void**)&enumerator);
												if(!FAILED(result))
												{
													VARIANT er = {0};
													VARIANT_BOOL vrb = 1;
													//_bstr_t ertext("");
													nRef = tPtr->AddRef();
													BSTR wbstr;
													tPtr->get_FullName(&wbstr);
													::SysFreeString(wbstr);
													nRef = tPtr->Release();
													tPtr = NULL;
													while(true)
													{
														if(FAILED(enumerator->MoveNext(&vrb)))
															break;
														if(vrb)
														{
															if(FAILED(enumerator->get_Current(&er)))
																break;
															if(!tPtr)
																tPtr = GetVariantType(er);
															//BSTR fullname;
															//tPtr->get_FullName(&fullname);

															nRef = tPtr->AddRef();
															nRef = tPtr->Release();

															tPtr->InvokeMember_3(_bstr_t("Line"), BindingFlags_GetProperty, NULL, er, NULL, &rvar);
															//ertext += "Line: ";
 															tPtr->InvokeMember_3(_bstr_t("ErrorText"), BindingFlags_GetProperty, NULL, er, NULL, &cvar);
															//ertext += "Error: ";
															//ertext += cvar.bstrVal;
															//ertext += "\n";
															nRef = er.punkVal->Release();
															//nRef = tPtr->Release();
														}
														else
															break;
													}
													nRef = tPtr->Release();
													//pExcepInfo->bstrDescription = ertext;
												}
												nRef = enumerator->Release();
											}
											nRef = vEnumerator.punkVal->Release();
										}
										else	// нет ошибок
										{
											result = typeResults->InvokeMember_3(_bstr_t("CompiledAssembly"), BindingFlags_GetProperty, NULL, compileResults, NULL, &rvar);
											if(!FAILED(result))
											{
												tPtr = GetVariantType(rvar);
												args = SafeArrayCreateVector(VT_VARIANT, 0, 1);
												j = 0;
												SafeArrayPutElement(args, &j, classname);
												result = tPtr->InvokeMember_3(_bstr_t("CreateInstance"), BindingFlags_InvokeMethod, NULL, rvar, args, &cvar);
												SafeArrayDestroy(args);
												if(!FAILED(result))
												{
													nRef = cvar.punkVal->Release();
												}
												nRef = tPtr->Release();
												nRef = rvar.punkVal->Release();
											}
										}
									}
									nRef = typeErrors->Release();
									nRef = errors.punkVal->Release();
								}
								nRef = typeResults->Release();
								nRef = compileResults.punkVal->Release();
							}
							nRef = provider.punkVal->Release();
						}
						nRef = system->Release();
					}
				}
				nRef = parameters.punkVal->Release();
				nRef = pCompileParameters->Release();
			}
			nRef = pAppDomain->Release();
		}
		nRef = pUnkDomain->Release();
	}
	pCorRuntimeHost->Stop();
	nRef = pCorRuntimeHost->Release();
	return result;
}
	// RunCode
			//wchar_t *cdir = new wchar_t[1024];
			//DWORD lw = ::GetModuleFileNameW(get_HINSTANCE(), cdir, 1024);
			//while(lw > 0 && cdir[lw]!='\\' && cdir[lw]!='/') 
			//	lw--;
			//wmemcpy(cdir + lw + ((lw==0) ? 0 : 1), L"RunCode.dll", 12);
			//BSTR runCodeDll = ::SysAllocString(cdir);
			//delete cdir;

			//mscorlib::_ObjectHandle *pRunCode = NULL;
			//BSTR type = ::SysAllocString(L"RunCode.RunCode"); 
			//result = pAppDomain->CreateInstanceFrom(runCodeDll, type, &pRunCode);
			//::SysFreeString(type);
			//::SysFreeString(runCodeDll);
			//if(!FAILED(result))
			//{
			//	result = RunCode(pRunCode, classname, L"Run", code, assemblies, vResult);
			//	nRef = pRunCode->Release();
			//}

#endif

void CALLBACK RunAssemblyCodeW(HWND hwnd, HINSTANCE hinst, LPWSTR cmdLine, int nCmdShow)
{
	HRESULT result;
	VARIANT vResult = {0};

	//XCom.dll,RunAssemblyCode /RunCode=D:\VS\CPP\XCom\RunCode\bin\Debug\RunCode.dll  /Type = classname  /Code=1
	//Parser *parser = new Parser(cmdLine);
	//BSTR classname = parser->GetValue(L"Type");
	//BSTR code = parser->GetValue(L"Code");
	//EXCEPINFO excepInfo = {0};
	//result = ExecuteAssemblyCode(classname, code, NULL, &excepInfo, &vResult);

	//delete parser;
}

#ifndef MSCOREELIB
#pragma managed
#using <System.dll>
using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Runtime::Remoting;
using namespace System::Reflection;
using namespace System::CodeDom::Compiler;
using namespace Microsoft::CSharp; 
//using namespace System::Windows::Forms;

HRESULT ExecuteAssemblyCode(BSTR classname, BSTR code, VARIANTARG *assemblies, tagEXCEPINFO * pExcepInfo, VARIANT * vResult)
{
	LONG i;
	HRESULT rc = S_OK;

	CSharpCodeProvider ^ codeCompiler = gcnew CSharpCodeProvider();
	// добавл€ем ссылки на сборки 
	CompilerParameters ^ parameters = gcnew CompilerParameters();
	// добавл€ем ссылку на нашу сборку SimpleMacro 
	if(assemblies != NULL)
	{
		if((assemblies->vt & VT_ARRAY)==VT_ARRAY && (assemblies->vt & VT_BSTR)==VT_BSTR)
		{
			for(i=0; i < assemblies->parray->rgsabound[0].cElements; i++)
			{
				BSTR wbStr;
				if(SafeArrayGetElement(assemblies->parray, &i, &wbStr) == S_OK)
				{
					String ^ assemblyName = Marshal::PtrToStringBSTR(IntPtr(wbStr));
					parameters->ReferencedAssemblies->Add(assemblyName);
				}
			}
		}
		else if(assemblies->vt == VT_BSTR)
		{
			String ^ assemblyName = Marshal::PtrToStringBSTR(IntPtr(assemblies->bstrVal));
			parameters->ReferencedAssemblies->Add(assemblyName);
		}
	}

	String ^ executingAssembly = Assembly::GetExecutingAssembly()->Location;
	parameters->ReferencedAssemblies->Add(executingAssembly);
	// компилируем 
	CompilerResults ^ results = codeCompiler->CompileAssemblyFromSource(parameters, Marshal::PtrToStringBSTR(IntPtr(code)));
	// есть ли ошибки? 
	if (results->Errors->HasErrors)
	{
		String ^ errs = "";
		for(i=0; i < results->Errors->Count; i++)
		{
			errs += String::Format("Line:{0:d},Error:{1}\n", results->Errors[0]->Line, results->Errors[0]->ErrorText);
		}
		IntPtr errPtr = Marshal::StringToBSTR(errs);
		pExcepInfo->bstrDescription = (BSTR)((void*)errPtr);
		pExcepInfo->bstrSource = ::SysAllocString(L"Execute assembly code");
		//throw gcnew ArgumentException("ќшибки при компил€ции.");
		rc = DISP_E_EXCEPTION;
	}
	else
	{
		//создаем класс 
		Object ^ cobj = results->CompiledAssembly->CreateInstance(Marshal::PtrToStringBSTR(IntPtr(classname)));
		if(cobj == nullptr)
		{
			pExcepInfo->bstrSource = ::SysAllocString(L"Execute assembly code");
			pExcepInfo->bstrDescription = ::SysAllocString(L"ќшибка при создании класса");
			rc = DISP_E_EXCEPTION;
		}
		else
		{
			Object ^ resObj = cobj->GetType()->InvokeMember("Run", BindingFlags::InvokeMethod, nullptr, cobj, nullptr);
			if(IsVarVal(vResult))
			{
				if(resObj == nullptr)
				{
				}
				else
				{
					Marshal::GetNativeVariantForObject(resObj, IntPtr(vResult));
				}
			}
			rc = S_OK;
		}
		String ^ ssert = results->CompiledAssembly->Location;
		BOOL rd = ::DeleteFileW((BSTR)(void*)Marshal::StringToBSTR(ssert));
	}
	return rc;
}
HRESULT LoadAssemblyObject(BSTR dllname, BSTR type, VARIANT * pObject)
{
	return S_OK;
}

HRESULT UnloadAssemblyObject(BSTR dllname)
{
	return S_OK;
}
#pragma unmanaged
#endif
