#include "stdafx.h"
#include "RabbitMQWrapper.h"
void Register(const wchar_t *dllname, const wchar_t *typeobject)
{
	//CRabbitMQWrapper ^ rabbit = gcnew CRabbitMQWrapper();
	//System::String ^ strDllname = gcnew System::String(dllname);
	//rabbit->SendMessage(strDllname);
	//delete strDllname;
	//delete rabbit;

	RabbitMQ::Client::ConnectionFactory^ connection = gcnew ConnectionFactory();
	connection->CreateConnection();
	delete connection;
}
#pragma managed(push, off)

#include "NativeAddIn.h"

extern LONG gl_nRef;

enum ENUM_METHOD_NUMBERS
{
    emRegister = 0,
	emUnload,
	emUnloadAssembly,
	emTestError,
	emVersion
};
CMethod gl_Methods [] =
{
	CMethod(L"Register", emRegister, 2, true), 
}; 

// CAddInNative
CAddInNative::CAddInNative()
{
	this->m_cMethods = gl_Methods;
	this->m_nMethods = sizeof(gl_Methods) / sizeof(CMethod);

	//FILE *fp = fopen("\\\\svu\\File_DBS\\ExtProc\\XCom\\XCom.txt", "a+");
	//fprintf(fp, "Register %u %u %u\n", ::GetCurrentProcessId(), GetCurrentThreadId(), gl_nRef);
	//fclose(fp);

}
//---------------------------------------------------------------------------//
CAddInNative::~CAddInNative()
{
}
// ILanguageExtenderBase
bool CAddInNative::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
	return AllocWSTR(wsExtensionName, L"NativeAddIn");
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetParamDefValue(const long lMethodNum, 
									const long lParamNum,
									tVariant *pvarParamDefValue)
{ 
	bool result = false;
	if(lMethodNum == emRegister)
	{
		if(lParamNum==1 || lParamNum==0)
		{
			TV_VT(pvarParamDefValue)= VTYPE_PWSTR;
			TV_WSTR(pvarParamDefValue) = L"";
		}
	}

    return result;
} 
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsProc(	const long lMethodNum,
								tVariant* paParams, 
								const long lSizeArray)
{ 
    return false;
}
//---------------------------------------------------------------------------
bool CAddInNative::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    bool result = true;

	if (lMethodNum == emRegister)
	{
		const wchar_t *argdllname = 0;
		const wchar_t *argtypeobject = 0;

		if (lSizeArray > 0)
			argdllname = paParams[0].pwstrVal;
		if (lSizeArray > 1)
			argtypeobject = paParams[1].pwstrVal;

		Register(argdllname, argtypeobject);
	}

	return result;
}
