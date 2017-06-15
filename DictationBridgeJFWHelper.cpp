#define UNICODE
#include <windows.h>
#include <TlHelp32.h>
#include <ole2.h>
#include <AtlBase.h>
#include <algorithm>
#include <comdef.h>
#include <sstream>
#include <string>
#include <map>
#include <utility>
using namespace std;
#include "FSAPI.h"
#include "dictationbridge-core/master/master.h"
#include "combool.h"
#include "ProcessMonitor.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleacc.lib")
#pragma comment(lib, "comsuppw.lib")

#define ERR(x, msg) do { \
if(x != S_OK) {\
MessageBox(NULL, msg L"\n", NULL, NULL);\
exit(1);\
}\
} while(0)

CComPtr<IJawsApi> pJfw =nullptr;
ProcessMonitor *pProcessMonitor;
map<wstring, HWINEVENTHOOK> ProcessWinEventHooks; //Hold the WinEvent hooks for each process.
//variables for WMI.
CComPtr<					  IWbemLocator> pLoc = nullptr;
CComPtr<	IWbemServices> pSvc = nullptr;
CComPtr<IUnsecuredApartment> pUnsecApp = nullptr;
CComPtr<IUnknown> pStubUnk = nullptr;
CComPtr<IWbemObjectSink> pStubSink = nullptr;

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

void initJAWS() 
{
		CLSID JFWClass;
HRESULT hr = S_FALSE;
		hr = CLSIDFromProgID(L"FreedomSci.JawsApi", &JFWClass);
		ERR(hr, L"Couldn't get Jaws interface ID");
		hr = pJfw.CoCreateInstance(JFWClass);
		ERR(hr, L"Couldn't create Jaws interface");
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

HRESULT SetWinEventHookForProcess(_In_ DWORD eventMin, _In_ DWORD eventMax, _In_ WINEVENTPROC pfnWinEventProc, _In_ DWORD idProcess)
{
	HRESULT hr = S_FALSE;
		HWINEVENTHOOK hook = SetWinEventHook(eventMin, eventMax, NULL, pfnWinEventProc, idProcess, 0, WINEVENT_OUTOFCONTEXT);
		if (hook != 0)
		{
			//Add to the map so that we can unhook later.
			ProcessWinEventHooks.insert(make_pair(L"natspeak.exe", hook));
		hr =S_OK;
		}
		
		return hr;
}

void InitializeWindowsHooksForDragonProcesses()
{
	DWORD *prgProcessIds = nullptr;
	DWORD cProcessIds = 0, iProcessId;
	auto res = FindIdsIfProcessIsRunning(L"natspeak.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(res) && cProcessIds >0)
	{
		for (iProcessId = 0; iProcessId < cProcessIds; ++iProcessId)
		{
			res = SetWinEventHookForProcess(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, nameChanged, prgProcessIds[iProcessId]);
			}
		HeapFree(GetProcessHeap(), 0, static_cast<LPVOID>(prgProcessIds));
		prgProcessIds = NULL;
		cProcessIds = 0;
	}

	//Find an initialize the DragonBar hook.
	res = S_FALSE;
	res = FindIdsIfProcessIsRunning(L"dragonbar.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(res) && cProcessIds >0)
	{
		for (iProcessId = 0; iProcessId < cProcessIds; ++iProcessId)
		{
			res = SetWinEventHookForProcess(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, nameChanged, prgProcessIds[iProcessId]);
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

//constants for process we need to keep track of.
const LPCWSTR jawsProcessName = L"jfw.exe";
const LPCWSTR natspeakProcessName = L"natspeak.exe";
const LPCWSTR dragonbarProcessName = L"natspeak.exe";

void WINAPI processCreatedCallback(DWORD processID, LPCWSTR processName)
{
	MessageBox(NULL, L"Callback called.", L"Process creation", MB_OK | MB_ICONERROR);
	return;
}

void WINAPI processDeletedCallback(LPCWSTR processName)
{
	if (wcsicmp(processName, jawsProcessName) == 0)
	{
		//JAWS has exited, so release the API pointer.
		pJfw.Release();
		pJfw = nullptr;
	}
	else if (wcsicmp(processName, natspeakProcessName) == 0)
	{
		//the natspeak process has terminated, so unhook the winevent for that process.
		auto processHook = ProcessWinEventHooks.find(processName);
		if (processHook != end(ProcessWinEventHooks))
		{
			UnhookWinEvent(processHook->second);
			ProcessWinEventHooks.erase(processHook);
		}
	}
	else if (wcsicmp(processName, dragonbarProcessName) == 0)
	{
		//the dragonbar process has terminated, so unhook the winevent for that process.
		auto processHook = ProcessWinEventHooks.find(processName);
		if (processHook != end(ProcessWinEventHooks))
		{
			UnhookWinEvent(processHook->second);
			ProcessWinEventHooks.erase(processHook);
		}
	}
	return;
}

void StartProcessTracking()
{
	CComBSTR bRootNamespace = L"ROOT\\CIMV2";
	CComBSTR bWQL = L"WQL";
	CComBSTR bWQLQuery = "SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'";
//Com and security are initialized in WinMain.
	HRESULT hr = S_OK;
	// Obtain the initial locator to WMI
	hr = pLoc.CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER);
	ERR(hr, L"Failed to create IWbemLocator object.");

	// Connect to WMI through the IWbemLocator::ConnectServer method
	// Connect to the local root\cimv2 namespace
	// and obtain pointer pSvc to make IWbemServices calls.
	hr = pLoc->ConnectServer(bRootNamespace, NULL, NULL, 0, NULL, 0, 0, &pSvc);
	ERR(hr, L"Could not connect to the WMI root\\\\cimv2 namespace.");

	// Set security levels on the proxy 
	hr = CoSetProxyBlanket(pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE);
	ERR(hr, L"Could not set proxy blanket.");

	// Receive event notifications
	// Use an unsecured apartment for security
	hr = pUnsecApp.CoCreateInstance(CLSID_UnsecuredApartment, NULL, CLSCTX_LOCAL_SERVER);
	ERR(hr, L"Unable to create the unsecured apartment.");

	pProcessMonitor = new ProcessMonitor;
	pProcessMonitor->SetProcessCreatedCallback(&processCreatedCallback);
	pProcessMonitor->SetProcessDeletedCallback(&processDeletedCallback);
	pProcessMonitor->AddRef();

	hr = pUnsecApp->CreateObjectStub(pProcessMonitor, &pStubUnk);
	ERR(hr, L"Could not create the object forwarder sink for use by WMI.");


	hr = pStubUnk->QueryInterface(IID_IWbemObjectSink, (void **)&pStubSink);
	ERR(hr, L"could not obtain the IWbemObjectSink interface.");

	// The ExecNotificationQueryAsync method will call
	// The EventQuery::Indicate method when an event occurs
	hr = pSvc->ExecNotificationQueryAsync(bWQL, bWQLQuery, WBEM_FLAG_SEND_STATUS, NULL, pStubSink);
	ERR(hr, L"ExecNotificationQueryAsync failed.");
}

void TerminateProcessTracking()
{
	HRESULT res = S_OK;
	res = pSvc->CancelAsyncCall(pStubSink);
	ERR(res, L"Unable to cancel the process tracking.");
}

HRESULT InitializeCom()
{
	HRESULT hr = S_OK;
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (SUCCEEDED(hr))
	{
		// Set general COM security levels
		hr = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	}
		return hr;
}

void UnhookAllWinEventProcessSpecificHooks()
{
	for (auto & specificProcessHook : ProcessWinEventHooks)
	{
		UnhookWinEvent(specificProcessHook.second);
}
	ProcessWinEventHooks.clear();
}

int CALLBACK WinMain(_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPSTR lpCmdLine,
	_In_ int nCmdShow) 
{
	// First, is a core running?
	if(FindWindow(msgWindowClassName, NULL)) {
		MessageBox(NULL, L"Core already running.", NULL, NULL);
			return 0;
	}
	//Next, initialize COM.
	HRESULT hr = InitializeCom();

	WNDCLASS windowClass = {0};
	windowClass.lpfnWndProc = exitProc;
	windowClass.hInstance = hInstance;
	windowClass.lpszClassName = msgWindowClassName;
	auto msgWindowClass = RegisterClass(&windowClass);
	if(msgWindowClass == 0) {
		MessageBox(NULL, L"Failed to register window class.", NULL, NULL);
		CoUninitialize();
		return 0;
	}
	auto msgWindowHandle = CreateWindow(msgWindowClassName, NULL, NULL, NULL, NULL, NULL, NULL, HWND_MESSAGE, NULL, GetModuleHandle(NULL), NULL);
	if(msgWindowHandle == 0) {
		MessageBox(NULL, L"Failed to create message-only window.", NULL, NULL);
		CoUninitialize();
		return 0;
	}
	
	//Initialize JAWS if it is running.
	DWORD *prgProcessIds = nullptr;
	DWORD cProcessIds = 0, iProcessId;
	hr = FindIdsIfProcessIsRunning(L"jfw.exe", &prgProcessIds, &cProcessIds);
	if (SUCCEEDED(hr) && cProcessIds >1)
	{
		HeapFree(GetProcessHeap(), 0, prgProcessIds);
		prgProcessIds = NULL;
		cProcessIds = 0;
		initJAWS();
	}

	auto started = DBMaster_Start();
	if(!started) {
		printf("Couldn't start DictationBridge-core\n");
		CoUninitialize();
		return 1;
	}
	DBMaster_SetTextInsertedCallback(textCallback);
	DBMaster_SetTextDeletedCallback(textDeletedCallback);
	
	//register to receive events from both the natspeak and DragonBar processes.
	InitializeWindowsHooksForDragonProcesses();
	StartProcessTracking();
	
	MSG msg;
	while(GetMessage(&msg, NULL, NULL, NULL) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
		if(keepRunning == 0) break;
	}
	
	//Shutdown all subsystems.
	DBMaster_Stop();
	UnhookAllWinEventProcessSpecificHooks();
	TerminateProcessTracking();
CoUninitialize();
	DestroyWindow(msgWindowHandle);
	return 0;
}