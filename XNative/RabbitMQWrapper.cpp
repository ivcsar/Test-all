using namespace std;

#include "RabbitMQWrapper.h"

CRabbitMQWrapper::CRabbitMQWrapper()
{
	System::Windows::Forms::MessageBox::Show("Constructor");
}

CRabbitMQWrapper::~CRabbitMQWrapper()
{
	System::Windows::Forms::MessageBox::Show("Destructor");
}

void CRabbitMQWrapper::SendMessage(System::String ^ dllname)
{
	System::Windows::Forms::MessageBox::Show(dllname);
};

