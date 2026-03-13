#include "..\stdafx.h"

//#define MSCORCLR
#ifdef MSCORCLR
#pragma managed
#using <System.dll>
using namespace System;
using namespace System::Runtime::InteropServices;
//using namespace System::Runtime::Remoting;
using namespace System::Reflection;

public delegate void ObjectProcessingDelegate(Object ^ parent, Object ^ data, Object ^ _object);

//[Guid("C8959CD0-03A6-4542-BBA6-6508283FE651")]
//[InterfaceType(ComInterfaceType(IDispatchImplType]))
//public interface class IProcHandler
//{
//	event ObjectProcessingDelegate ^ proc;
//};
[ComVisibleAttribute(true)]
public ref class ProcHandler
{
public:
	virtual event ObjectProcessingDelegate ^ proc;
	void RunEvent(Object ^ parent, Object ^ data, Object ^ _object)
	{
		{
			proc(parent, data, _object);
		}
	}
};

STDAPI GetProcHandler(LPVOID *ppv)
{
	ProcHandler ^ proc = gcnew ProcHandler();

	*ppv = Marshal::GetIDispatchForObject(proc).ToPointer();

	return S_OK;
}


public delegate void ClickEventHandler(int, double);
// define interface with nested interface
public interface class Interface_A {
	void Function_1();

	interface class Interface_Nested_A 
	{
		void Function_2();
	};
};

// interface with a base interface
public interface class Interface_B : Interface_A {
	property int Property_Block;
	event ClickEventHandler^ OnClick;
	static void Function_3() { Console::WriteLine("in Function_3"); }
};

// implement nested interface
public ref class MyClass : public Interface_A::Interface_Nested_A {
public:
	virtual void Function_2() { Console::WriteLine("in Function_2"); }
};

// implement interface and base interface
public ref class MyClass2 : public Interface_B {
private:
	int MyInt;

public:
	// implement non-static function
	virtual void Function_1() { Console::WriteLine("in Function_1"); }

	// implement property
	property int Property_Block {
		virtual int get() { return MyInt; }
		virtual void set(int value) { MyInt = value; }
	}
	// implement event
	virtual event ClickEventHandler^ OnClick;

	void FireEvents() {
		OnClick(7, 3.14159);
	}
};

// class that defines method called when event occurs
ref class EventReceiver {
public:
	void OnMyClick(int i, double d) {
		Console::WriteLine("OnClick: {0}, {1}", i, d);
	}
};
#pragma unmanaged
#endif
