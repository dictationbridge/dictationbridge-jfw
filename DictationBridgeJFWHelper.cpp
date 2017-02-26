#define UNICODE
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
#pragma comment(lib, "oleacc.lib")


#define ERR(x, msg) do { \
if(x != S_OK) {\
MessageBox(NULL, msg L"\n", NULL, NULL);\
exit(1);\
}\
} while(0)

std::string wideToString(wchar_t const * text, unsigned int length) {
	auto tmp = new char[length*2+1];
	auto resultingLen = WideCharToMultiByte(CP_UTF8, NULL, text,
		length, tmp, length*2,
		NULL, NULL);
	tmp[resultingLen] = '\0';
	std::string ret(tmp);
	delete[] tmp;
	return ret;
}

std::string BSTRToString(BSTR text) {
	unsigned int len = SysStringLen(text);
	return wideToString(text, len);
}

DISPID id;
DISPPARAMS params;
VARIANTARG args[2];
IDispatch *jfw = nullptr;

void initSpeak() {
	CLSID JFWClass;
	auto res = CLSIDFromProgID(L"FreedomSci.JawsApi", &JFWClass);
	ERR(res, L"Couldn't get Jaws interface ID");
	res = CoCreateInstance(JFWClass, NULL, CLSCTX_ALL, __uuidof(IDispatch), (void**)&jfw);
	ERR(res, L"Couldn't create Jaws interface");
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
	ERR(res, L"Couldn't get SayString");
}

void speak(wchar_t const * text) {
	auto s = SysAllocString(text);
	args[1].bstrVal = s;
	jfw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
	SysFreeString(s);
}

void WINAPI textCallback(HWND hwnd, DWORD startPosition, LPCWSTR textUnprocessed) {
	//We need to replace \r with nothing.
	int len = wcslen(textUnprocessed);
	wchar_t* text = new wchar_t[len+1];
	int copied = 0;
	for(int i = 0; i < len; i++) {
		if(textUnprocessed[i] != L'\r') {
			text[copied] = textUnprocessed[i];
			copied += 1;
		}
	}
	text[copied] = 0;
	if(wcscmp(text, L"\n\n") == 0
		|| wcscmp(text, L"") == 0 //new paragraph in word.
	) {
		speak(L"New paragraph.");
	}
	else if(wcscmp(text, L"\n")  == 0) {
		speak(L"New line.");
	}
	else {
		speak(text);
	}
}

void WINAPI textDeletedCallback(HWND hwnd, DWORD startPosition, LPCWSTR text) {
	int len = wcslen(text);
	int neededLen = len + strlen("Deleted ");
	wchar_t* tmp = new wchar_t[neededLen+1];
	wcscpy(tmp, L"Deleted ");
	int offset = strlen("Deleted ");
	for(int i = 0; i < len; i++) tmp[i+offset] = text[i];
	tmp[neededLen] = 0;
	speak(tmp);
	delete[] tmp;
}

//These are string constants for the microphone status, as well as the status itself:
//The pointer below is set to the last one we saw.
const char* MICROPHONE_OFF = "Dragon's microphone is off;";
const char* MICROPHONE_ON = "Normal mode: You can dictate and use voice";
const char* MICROPHONE_SLEEPING = "The microphone is asleep;";

const char* microphoneState = nullptr;

void announceMicrophoneState(const char* state) {
	if(state == MICROPHONE_ON) speak(L"Microphone on.");
	else if(state == MICROPHONE_OFF) speak(L"Microphone off.");
	else if(state == MICROPHONE_SLEEPING) speak(L"Microphone sleeping.");
	else speak(L"Microphone in unknown state.");
}

wchar_t processNameBuffer[1024] = {0};

void CALLBACK nameChanged(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime) {
	//First, is it coming from natspeak.exe?
	DWORD procId;
	GetWindowThreadProcessId(hwnd, &procId);
	auto procHandle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, procId);
	//We can't recover from this failing, so abort.
	if(procHandle == NULL) return;
	DWORD len = 1024;
	auto res = QueryFullProcessImageName(procHandle, 0, processNameBuffer, &len);
	CloseHandle(procHandle);
	if(res == 0) return;
	auto processName = wideToString(processNameBuffer, (unsigned int) len);
	if(processName.find("dragonbar.exe") == std::string::npos
		&& processName.find("natspeak.exe") == std::string::npos) return;
	//Attempt to get the new text.
	IAccessible* acc = nullptr;
	VARIANT child;
	HRESULT hres = AccessibleObjectFromEvent(hwnd, idObject, idChild, &acc, &child);
	if(hres != S_OK || acc == nullptr) return;
	BSTR nameBSTR;
	hres = acc->get_accName(child, &nameBSTR);
	acc->Release();
	if(hres != S_OK) return;
	auto name = BSTRToString(nameBSTR);
	SysFreeString(nameBSTR);
	const char* possibles[] = {MICROPHONE_ON, MICROPHONE_OFF, MICROPHONE_SLEEPING};
	const char* newState = microphoneState;
	for(int i = 0; i < 3; i++) {
		if(name.find(possibles[i]) != std::string::npos) {
			newState = possibles[i];
			break;
		}
	}
	if(newState != microphoneState) {
		announceMicrophoneState(newState);
		microphoneState = newState;
	}
}

int keepRunning = 1; // Goes to 0 on WM_CLOSE.
LPCTSTR msgWindowClassName = L"DictationBridgeJFWHelper";

LRESULT CALLBACK exitProc(_In_ HWND hwnd, _In_ UINT msg, _In_ WPARAM wparam, _In_ LPARAM lparam) {
	if(msg == WM_CLOSE) keepRunning = 0;
	return DefWindowProc(hwnd, msg, wparam, lparam);
}

int CALLBACK WinMain(_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPSTR lpCmdLine,
	_In_ int nCmdShow) {
	// First, is a core running?
	if(FindWindow(msgWindowClassName, NULL)) {
		MessageBox(NULL, L"Core already running.", NULL, NULL);
		return 0;
	}
	WNDCLASS windowClass = {0};
	windowClass.lpfnWndProc = exitProc;
	windowClass.hInstance = hInstance;
	windowClass.lpszClassName = msgWindowClassName;
	auto msgWindowClass = RegisterClass(&windowClass);
	if(msgWindowClass == 0) {
		MessageBox(NULL, L"Failed to register window class.", NULL, NULL);
		return 0;
	}
	auto msgWindowHandle = CreateWindow(msgWindowClassName, NULL, NULL, NULL, NULL, NULL, NULL, HWND_MESSAGE, NULL, GetModuleHandle(NULL), NULL);
	if(msgWindowHandle == 0) {
		MessageBox(NULL, L"Failed to create message-only window.", NULL, NULL);
		return 0;
	}
	HRESULT res;
	res = OleInitialize(NULL);
	ERR(res, L"Couldn't initialize OLE");
	initSpeak();
	auto started = DBMaster_Start();
	if(!started) {
		printf("Couldn't start DictationBridge-core\n");
		return 1;
	}
	DBMaster_SetTextInsertedCallback(textCallback);
	DBMaster_SetTextDeletedCallback(textDeletedCallback);
	if(SetWinEventHook(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, NULL, nameChanged, 0, 0, WINEVENT_OUTOFCONTEXT) == 0) {
		printf("Couldn't register to receive events\n");
		return 1;
	}
	MSG msg;
	while(GetMessage(&msg, NULL, NULL, NULL) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
		if(keepRunning == 0) break;
	}
	DBMaster_Stop();
	OleUninitialize();
	DestroyWindow(msgWindowHandle);
	return 0;
}
