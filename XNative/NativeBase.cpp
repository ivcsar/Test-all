#include "NativeBase.h"

CNativeBase::CNativeBase(void)
{
    m_iMemory = 0;
    m_iConnect = 0;
	m_nMethods = 0;
	m_cMethods = NULL;
}

CNativeBase::~CNativeBase(void)
{

}

//---------------------------------------------------------------------------//
bool CNativeBase::Init(void* pConnection)
{ 
    m_iConnect = (IAddInDefBase*)pConnection;
    return (m_iConnect != NULL);
}
//---------------------------------------------------------------------------//
long CNativeBase::GetInfo()
{ 
    // Component should put supported component technology version 
    // This component supports 2.0 version
    return 2000; 
}
//---------------------------------------------------------------------------//
void CNativeBase::Done()
{
}
//---------------------------------------------------------------------------//
long CNativeBase::GetNProps()
{ 
    // You may delete next lines and add your own implementation code here
    return 0;
}
//---------------------------------------------------------------------------//
long CNativeBase::FindProp(const WCHAR_T* wsPropName)
{ 
    return -1;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CNativeBase::GetPropName(long lPropNum, long lPropAlias)
{ 
    return NULL;
}
//---------------------------------------------------------------------------//
bool CNativeBase::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CNativeBase::SetPropVal(const long lPropNum, tVariant *varPropVal)
{ 
	return false;
}
//---------------------------------------------------------------------------//
bool CNativeBase::IsPropReadable(const long lPropNum)
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CNativeBase::IsPropWritable(const long lPropNum)
{
    return false;
}
//---------------------------------------------------------------------------//
long CNativeBase::GetNMethods()
{ 
    return (m_nMethods);
}
//---------------------------------------------------------------------------//
long CNativeBase::FindMethod(const WCHAR_T *wsMethodName)
{ 
	long i;
	
	CMethod *pMethod = m_cMethods;
	for(i=0; i < m_nMethods; i++)
	{
		if(_wcsicmp(pMethod->method, wsMethodName) == 0)
		{
			return pMethod->id;
		}
		pMethod++;
	}
	return -1;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CNativeBase::GetMethodName(const long lMethodNum, const long lMethodAlias)
{ 
	long i;
	wchar_t *wszMethodName = NULL;
	
	CMethod *pMethod = m_cMethods;
	for(i=0; i < m_nMethods; i++)
	{
		if(pMethod->id == lMethodNum)
		{
			AllocWSTR(&wszMethodName, pMethod->method);
			break;
		}
		pMethod++;
	}
	return wszMethodName;
}
//---------------------------------------------------------------------------//
long CNativeBase::GetNParams(const long lMethodNum)
{ 
	CMethod *pMethod = m_cMethods;
	for(long i=0; i < (m_nMethods); i++)
	{
		if(pMethod->id == lMethodNum)
		{
			return pMethod->nParams;
		}
		pMethod++;
	}
    return 0;
}
//---------------------------------------------------------------------------//
bool CNativeBase::GetParamDefValue(	const long lMethodNum, 
								const long lParamNum,
								tVariant *pvarParamDefValue)
{ 
	bool result = false;
	TV_VT(pvarParamDefValue)= VTYPE_EMPTY;
    return result;
} 
//---------------------------------------------------------------------------//
bool CNativeBase::HasRetVal(const long lMethodNum)
{ 
	CMethod *pMethod = m_cMethods;
	for(long i=0; i < (m_nMethods); i++)
	{
		if(pMethod->id == lMethodNum)
		{
			return pMethod->hr;
		}
		pMethod++;
	}
    return false;
}
//---------------------------------------------------------------------------//
bool CNativeBase::CallAsProc(const long lMethodNum,
						tVariant *paParams, 
						const long lSizeArray)
{ 
    return false;
}

//---------------------------------------------------------------------------
bool CNativeBase::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    return false;
}
//---------------------------------------------------------------------------//
void CNativeBase::SetLocale(const WCHAR_T* loc)
{
	_wsetlocale(LC_ALL, loc);
}
/////////////////////////////////////////////////////////////////////////////
// LocaleBase
//---------------------------------------------------------------------------//
bool CNativeBase::setMemManager(void* mem)
{
    m_iMemory = (IMemoryManager*)mem;
    return m_iMemory != 0;
}
//---------------------------------------------------------------------------//
void CNativeBase::addError(uint32_t wcode, 
					 const wchar_t *source, 
					 const wchar_t *descriptor, 
					 long code)
{
    if (m_iConnect)
    {
        //WCHAR_T *err = 0;
        //WCHAR_T *descr = 0;
        //
        //::convToShortWchar(&err, source);
        //::convToShortWchar(&descr, descriptor);

		//m_iConnect->AddError(wcode, err, descr, code);
        //delete[] err;
        //delete[] descr;
    }
}

void CNativeBase::CallException(DWORD errc, wchar_t *source)
{
	wchar_t *errbuf = new wchar_t [2048];
	DWORD ln = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 
							NULL, 
							errc, 
							MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
							errbuf, 
							2048, 
							NULL);
	BSTR werrstr = ::SysAllocString(errbuf);
	delete [] errbuf;
	bool rc;

	if(m_iConnect)
	{
		long errNo = RESULT_FROM_ERRNO(errc);
		rc = m_iConnect->AddError(1004, (wchar_t*)werrstr, source, 0);	// если вместо 0 поставить errNo - все падает
		//m_iConnect->SetStatusLine(werrstr);
	}

	::SysFreeString(werrstr);
}

void CNativeBase::SetWSTR(tVariant *tvar, wchar_t *pwstr)
{
	TV_VT(tvar) = VTYPE_PWSTR;
	TV_WSTR(tvar) = pwstr;
	tvar->wstrLen = wcslen(pwstr);
}

bool CNativeBase::AllocWSTR(WCHAR_T **pwstr, WCHAR_T *wstr)
{
	size_t lwcz = wcslen(wstr);

    if (m_iMemory)
    {
        if(m_iMemory->AllocMemory((void**)pwstr, (lwcz+1) * sizeof(WCHAR_T)))
		{
			wmemmove(*pwstr, wstr, lwcz);
			(*pwstr)[lwcz] = L'\0';
			return true;
		}
    }
	return false;
}