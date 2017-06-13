#define UNICODE
#include <windows.h>
#include <TlHelp32.h>
#include <ole2.h>
#include <AtlBase.h>
#include <algorithm>
#include <sstream>
#include <string>
using namespace std;
#include "FSAPI.h"
#include "dictationbridge-core/master/master.h"
#include "combool.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleacc.lib")


#define ERR(x, msg) do { \
if(x != S_OK) {\
MessageBox(NULL, msg L"\n", NULL, NULL);\
exit(1);\
}\
} while(0)

CComPtr<IJawsApi> pJfw =nullptr;
HWINEVENTHOOK hNatspeakNameChangedHook;
HWINEVENTHOOK hDragonBarNameChangedHook;

HRESULT FindIdsIfProcessIsRunning(__in_z LPCWSTR wzExeName, __out DWORD** ppdwProcessIds, __out DWORD* pcProcessIds)
{
	HRESULT hr = S_OK;
	DWORD er = ERROR_SUCCESS;
	HANDLE hSnapshot = INVALID_HANDLE_VALUE;
	bool fContinue = false;
	PROCESSENTRY32 peData = { sizeof(peData) };
	hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	if (hSnapshot == INVALID_HANDLE_VALUE)
	{
		er = GetLastError();
		hr = HRESULT_FROM_WIN32(er);
			goto LExit;
	}
	
	fContinue = Process32First(hSnapshot, &peData);
	while (fContinue)
	{
		if (wcsicmp(wzExeName, peData.szExeFile) == 0)
		{
			if (!*ppdwProcessIds)
			{
				hr = S_OK;
				*ppdwProcessIds = static_cast<DWORD*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(DWORD)));
				
				if (*ppdwProcessIds == nullptr)
				{
					hr = E_OUTOFMEMORY;
					goto LExit;
			}
			}
			else
			{
				DWORD* pdwReAllocReturnedPids = static_cast<DWORD*>(HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, *ppdwProcessIds, sizeof(DWORD) * ((*pcProcessIds) + 1)));
				
				if (pdwReAllocReturnedPids == nullptr)
				{
					hr = E_OUTOFMEMORY;
					goto LExit;
			}
				
				*ppdwProcessIds = pdwReAllocReturnedPids;
			}
			
			(*ppdwProcessIds)[*pcProcessIds] = peData.th32ProcessID;
			++(*pcProcessIds);
		}
		fContinue = Process32NextW(hSnapshot, &peData);
	}
	
	er = ::GetLastError();
	if (er == ERROR_NO_MORE_FILES && *pcProcessIds >0)
	{
		hr = S_OK;
	}
	else
	{
		hr = HRESULT_FROM_WIN32(er);
	}

LExit:
	if (hSnapshot != nullptr)
	{
		CloseHandle(hSnapshot);
	}
	
	return hr;
}

void initJAWSIfRunning() 
{
	DWORD *prgProcessIds = nullptr;
	DWORD cProcessIds = 0, iProcessId;
	auto res = FindIdsIfProcessIsRunning(L"jfw.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(res) && cProcessIds >= 1)
	{
		HeapFree(GetProcessHeap(), 0, prgProcessIds);
		CLSID JFWClass;
		res = S_FALSE;
		res = CLSIDFromProgID(L"FreedomSci.JawsApi", &JFWClass);
		ERR(res, L"Couldn't get Jaws interface ID");
		res = pJfw.CoCreateInstance(JFWClass);
		ERR(res, L"Couldn't create Jaws interface");
	}
	}

void speak(std::wstring text) {
	CComBSTR bS = CComBSTR(text.size(), text.data());
	CComBool silence = false;
	CComBool bResult;
	pJfw->SayString(bS, silence, &bResult);
}

//These are string constants for the microphone status, as well as the status itself:
//The pointer below is set to the last one we saw.
std::wstring MICROPHONE_OFF = L"Dragon's microphone is off;";
std::wstring MICROPHONE_ON = L"Normal mode: You can dictate and use voice";
std::wstring MICROPHONE_SLEEPING = L"The microphone is asleep;";

std::wstring microphoneState;
//This is a constant for the text indicating dragon hasn't understood what a user has dictated.
const std::wstring DictationWasNotUnderstood = L"<???>";

void announceMicrophoneState(const std::wstring state) {
	if (state == MICROPHONE_ON) speak(L"Microphone on.");
	else if (state == MICROPHONE_OFF) speak(L"Microphone off.");
	else if (state == MICROPHONE_SLEEPING) speak(L"Microphone sleeping.");
	else speak(L"Microphone in unknown state.");
}

wchar_t processNameBuffer[1024] = { 0 };

void CALLBACK nameChanged(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime) {
	//First, is it coming from natspeak.exe?
	DWORD procId;
	GetWindowThreadProcessId(hwnd, &procId);
	auto procHandle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, procId);
	//We can't recover from this failing, so abort.
	if (procHandle == NULL) return;
	DWORD len = 1024;
	auto res = QueryFullProcessImageName(procHandle, 0, processNameBuffer, &len);
	CloseHandle(procHandle);
	if (res == 0) return;
	std::wstring processName = processNameBuffer;
	if (processName.find(L"dragonbar.exe") == std::string::npos
		&& processName.find(L"natspeak.exe") == std::string::npos) return;
	//Attempt to get the new text.
	CComPtr<IAccessible> pAcc;
	CComVariant vChild;
	HRESULT hres = AccessibleObjectFromEvent(hwnd, idObject, idChild, &pAcc, &vChild);
	if (hres != S_OK) return;
	CComBSTR bName;
	hres = pAcc->get_accName(vChild, &bName);
	if (hres != S_OK) return;
	std::wstring name = bName;
	//check to see whether Dragon understood the user.
	if (name.compare(DictationWasNotUnderstood) == 0)
	{
		speak(L"I do not understand.");
		return;
	}
	const std::wstring possibles[] = { MICROPHONE_ON, MICROPHONE_OFF, MICROPHONE_SLEEPING };
	std::wstring newState = microphoneState;
	for (int i = 0; i < 3; i++) {
		if (name.find(possibles[i]) != std::string::npos) {
			newState = possibles[i];
			break;
		}
	}
	if (newState != microphoneState) {
		announceMicrophoneState(newState);
		microphoneState = newState;
	}
}

void InitializeWindowsCallbackForDragonProcesses()
{
	DWORD *prgProcessIds = nullptr;
	DWORD cProcessIds = 0, iProcessId;
	auto res = FindIdsIfProcessIsRunning(L"natspeak.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(res) && cProcessIds >0)
	{
		for (iProcessId = 0; iProcessId < cProcessIds; ++iProcessId)
		{
			hNatspeakNameChangedHook = SetWinEventHook(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, NULL, nameChanged, prgProcessIds[iProcessId], 0, WINEVENT_OUTOFCONTEXT);
			}
		HeapFree(GetProcessHeap(), 0, static_cast<LPVOID>(prgProcessIds));
		prgProcessIds = NULL;
	}

	//Find an initialize the DragonBar hook.
	cProcessIds = 0, iProcessId = 0;
	res = FindIdsIfProcessIsRunning(L"dragonbar.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(res) && cProcessIds >0)
	{
		for (iProcessId = 0; iProcessId < cProcessIds; ++iProcessId)
		{
			hDragonBarNameChangedHook = SetWinEventHook(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, NULL, nameChanged, prgProcessIds[iProcessId], 0, WINEVENT_OUTOFCONTEXT);
		}
		HeapFree(GetProcessHeap(), 0, static_cast<LPVOID>(prgProcessIds));
		prgProcessIds = NULL;
	}
	}

void WINAPI textCallback(HWND hwnd, DWORD startPosition, LPCWSTR textUnprocessed) {
	//We need to replace \r with nothing.
std::wstring text =textUnprocessed;
text.erase(std::remove_if(begin(text), end(text), [] (wchar_t checkingCharacter) {
		return checkingCharacter == '\r';
	}), end(text));

	if(text.compare(L"\n\n") ==0 
|| text.compare(L"") ==0 //new paragraph in word.
	) {
		speak(L"New paragraph.");
	}
	else if(text.compare(L"\n") ==0) {
		speak(L"New line.");
	}
	else {
speak(text.c_str());
}
}

void WINAPI textDeletedCallback(HWND hwnd, DWORD startPosition, LPCWSTR text) {
std::wstringstream deletedText;
deletedText << "Deleted ";
deletedText << text;
	speak(deletedText.str().c_str());
}

int keepRunning = 1; // Goes to 0 on WM_CLOSE.
LPCTSTR msgWindowClassName = L"DictationBridgeJFWHelper";

LRESULT CALLBACK exitProc(_In_ HWND hwnd, _In_ UINT msg, _In_ WPARAM wparam, _In_ LPARAM lparam) {
	if(msg == WM_CLOSE) keepRunning = 0;
	return DefWindowProc(hwnd, msg, wparam, lparam);
}
void StartProcessTracking()
{
	
}
void TerminateProcessTracking()
{
	
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
	HRESULT res = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	ERR(res, L"Couldn't initialize COM.");
	
	//Initialize JAWS if it is running.
	initJAWSIfRunning();

	auto started = DBMaster_Start();
	if(!started) {
		printf("Couldn't start DictationBridge-core\n");
		return 1;
	}
	DBMaster_SetTextInsertedCallback(textCallback);
	DBMaster_SetTextDeletedCallback(textDeletedCallback);
	
	//register to receive events from both the natspeak and DragonBar processes.
	InitializeWindowsCallbackForDragonProcesses();
	StartProcessTracking();
	
	MSG msg;
	while(GetMessage(&msg, NULL, NULL, NULL) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
		if(keepRunning == 0) break;
	}
	
	//Shutdown all subsystems.
	DBMaster_Stop();
	if (hNatspeakNameChangedHook != 0)
	{
		UnhookWinEvent(hNatspeakNameChangedHook);
		hNatspeakNameChangedHook = nullptr;
	}
	
	if (hDragonBarNameChangedHook != 0)
	{
		UnhookWinEvent(hDragonBarNameChangedHook);
		hDragonBarNameChangedHook = nullptr;
	}
	TerminateProcessTracking();
CoUninitialize();
	DestroyWindow(msgWindowHandle);
	return 0;
}