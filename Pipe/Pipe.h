#pragma once

//#include "..\NativeBase.h"
#include "..\ComBase.h"

class CPipe : 
	public ComBase
{
public:
	CPipe(void);
	~CPipe(void);
    // ILanguageExtenderBase
    //bool ADDIN_API RegisterExtensionAs(WCHAR_T**);
    //bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
    //                        tVariant *pvarParamDefValue);   
    //virtual bool ADDIN_API CallAsProc(const long lMethodNum,
    //                tVariant* paParams, const long lSizeArray);
    //virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
    //            tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
	/*
	long ADDIN_API GetNProps();
	long ADDIN_API FindProp(const WCHAR_T* wsPropName);
	const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
	bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
	bool ADDIN_API SetPropVal(const long lPropNum, tVariant *varPropVal);
	bool ADDIN_API IsPropReadable(const long lPropNum);
	bool ADDIN_API IsPropWritable(const long lPropNum);
	*/
	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
private:
	CRITICAL_SECTION m_CriticalSection;
public:
	ULONG m_status;
	HANDLE m_hEvent;
	HANDLE m_hPipeMutex;
	wchar_t * m_pipename;
	IStream *m_pStream;
	DISPID m_dispid;
	wchar_t *m_message;
	bool SendPipeMessage(wchar_t *server, wchar_t *pipename, wchar_t *message, EXCEPINFO *);
	HRESULT WaitMessage(wchar_t *pipename);
	//HRESULT AddEvent(DISPPARAMS * parameters, VARIANT *pVarResult);
	//HRESULT CallEvent(DISPPARAMS * parameters, VARIANT *pVarResult);
	//HRESULT GetStatus(VARIANT * pVarResult);
	void SetStatus(DWORD status);
	DWORD GetStatus();
	//void CreateStreamObject(DISPPARAMS * parameters);
	//BOOL ExecuteStreamObject();
	wchar_t *GetPipeMessage();
	void SetMessage(BYTE * buffer, DWORD ln);
	
	wchar_t *m_wTestVersion;
};
