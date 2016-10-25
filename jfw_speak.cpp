#include <windows.h>
#include <ole2.h>
#include <oaidl.h>
#include <objbase.h>
#include <AtlBase.h>
#include <AtlConv.h>
#include <iostream>
#include <string>
#include <vector>
#include <deque>
#include <mutex>
#include <thread>
#include <chrono>
#include "dictationbridge-core/master/master.h"

#pragma comment(lib, "ole32.lib")

#define ERR(x, msg) do { \
if(x != S_OK) {\
printf(msg "\n");\
printf("%x\n", res);\
return;\
}\
} while(0)

std::mutex queue_m;
std::deque<BSTR> queue;


void WINAPI textCallback(HWND hwnd, DWORD startPosition, LPCWSTR text) {
	printf("Got text\n");
	auto s = SysAllocString(text);
	std::lock_guard<std::mutex> g(queue_m);
	queue.push_back(s);
}

void main() {
	HRESULT res;
	res = OleInitialize(NULL);
	ERR(res, "Couldn't initialize OLE");
	CLSID JFWClass;
	IDispatch *jfw = nullptr;
	res = CLSIDFromProgID(L"FreedomSci.JawsApi", &JFWClass);
	ERR(res, "Couldn't get Jaws interface ID");
	res = CoCreateInstance(JFWClass, NULL, CLSCTX_ALL, __uuidof(IDispatch), (void**)&jfw);
	ERR(res, "Couldn't create Jaws interface");
	auto started = DBMaster_Start();
	if(!started) {
		printf("Couldn't start DictationBridge-core\n");
		return;
	}
	DBMaster_SetTextInsertedCallback(textCallback);
	/* Call jaws. IDL is:
	HRESULT SayString(
	[in] BSTR StringToSpeak, 
	[in, optional, defaultvalue(-1)] VARIANT_BOOL bFlush, 
	[out, retval] VARIANT_BOOL* vbSuccess);
	*/
	LPOLESTR name = L"SayString";
	DISPID id;
	DISPPARAMS params;
	VARIANTARG args[3];
	args[1].vt = VT_BSTR;
	args[0].vt = VT_BOOL;
	args[0].boolVal = 0;
	args[2].vt = VT_BOOL|VT_BYREF;
	VARIANT_BOOL ignored;
	args[2].pboolVal = &ignored;
	params.rgvarg = args;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;
	params.cNamedArgs = 0;
	res = jfw->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &id);
	ERR(res, "Couldn't get SayString");
	while(1) {
		BSTR s;
		{
			std::lock_guard<std::mutex> g(queue_m);
			if(queue.empty()) continue;
			s = queue.front();
			queue.pop_front();
		}
		printf("Saying text.\n");
		args[1].bstrVal = s;
		res = jfw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
		ERR(res, "Couldn't invoke");
		SysFreeString(s);
		std::this_thread::sleep_for(std::chrono::milliseconds(25));
	}
	DBMaster_Stop();
	OleUninitialize();
}
