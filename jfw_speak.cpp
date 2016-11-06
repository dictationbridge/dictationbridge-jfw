#include <windows.h>
#include <ole2.h>
#include <oaidl.h>
#include <objbase.h>
#include <AtlBase.h>
#include <AtlConv.h>
#include <iostream>
#include <string>
#include <vector>
#include <stdlib.h>
#include "dictationbridge-core/master/master.h"

#pragma comment(lib, "ole32.lib")

#define ERR(x, msg) do { \
if(x != S_OK) {\
printf(msg "\n");\
printf("%x\n", res);\
exit(1);\
}\
} while(0)

DISPID id;
DISPPARAMS params;
VARIANTARG args[2];
IDispatch *jfw = nullptr;

void initSpeak() {
	CLSID JFWClass;
	auto res = CLSIDFromProgID(L"FreedomSci.JawsApi", &JFWClass);
	ERR(res, "Couldn't get Jaws interface ID");
	res = CoCreateInstance(JFWClass, NULL, CLSCTX_ALL, __uuidof(IDispatch), (void**)&jfw);
	ERR(res, "Couldn't create Jaws interface");
	/* Setup to call jaws. IDL is:
	HRESULT SayString(
	[in] BSTR StringToSpeak, 
	[in, optional, defaultvalue(-1)] VARIANT_BOOL bFlush, 
	[out, retval] VARIANT_BOOL* vbSuccess);
	*/
	LPOLESTR name = L"SayString";
	args[1].vt = VT_BSTR;
	args[0].vt = VT_BOOL;
	args[0].boolVal = 0;
	params.rgvarg = args;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;
	params.cNamedArgs = 0;
	res = jfw->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &id);
	ERR(res, "Couldn't get SayString");
}

void speak(wchar_t const * text) {
	auto s = SysAllocString(text);
	args[1].bstrVal = s;
	jfw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
	SysFreeString(s);
}

void WINAPI textCallback(HWND hwnd, DWORD startPosition, LPCWSTR text) {
	speak(text);
}

void main() {
	HRESULT res;
	res = OleInitialize(NULL);
	ERR(res, "Couldn't initialize OLE");
	initSpeak();
	auto started = DBMaster_Start();
	if(!started) {
		printf("Couldn't start DictationBridge-core\n");
		return;
	}
	DBMaster_SetTextInsertedCallback(textCallback);
	MSG msg;
	while(GetMessage(&msg, NULL, NULL, NULL) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	DBMaster_Stop();
	OleUninitialize();
}
