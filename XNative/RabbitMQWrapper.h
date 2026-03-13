#pragma once
#using <System.dll>
#using "..\..\RabbitMQ\bin\RabbitMQ.Client.dll" 

using namespace System::Windows::Forms;
using namespace System::Runtime::InteropServices;
using namespace RabbitMQ::Client;

//public ref class Tessellator : System::IDisposable
//{
//internal: // эти поля не попадут в метаданные
//
//	Unmanaged::Tessellator * tess;
//
//public:
//
//	Tessellator(int numSteps)
//	{
//		tess = new Unmanaged::Tessellator(numSteps);
//	}
//
//	~Tessellator() // IDisposable::Dispose ()
//	{
//		delete tess;
//	}
//};

public ref class CRabbitMQWrapper
{
public:
	CRabbitMQWrapper();
	virtual ~CRabbitMQWrapper();
	void SendMessage(System::String ^ dllname);
};