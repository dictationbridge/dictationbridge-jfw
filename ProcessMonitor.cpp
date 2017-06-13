// ProcessMonitor.cpp

#include "ProcessMonitor.h"

ULONG ProcessMonitor::AddRef()
{
    return InterlockedIncrement(&m_lRef);
}

ULONG ProcessMonitor::Release()
{
    LONG lRef = InterlockedDecrement(&m_lRef);
    if(lRef == 0)
        delete this;
    return lRef;
}

HRESULT ProcessMonitor::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IWbemObjectSink)
    {
        *ppv = (IWbemObjectSink *) this;
        AddRef();
        return WBEM_S_NO_ERROR;
    }
    else return E_NOINTERFACE;
}


HRESULT ProcessMonitor::Indicate(long lObjectCount,
    IWbemClassObject **apObjArray)
{
 HRESULT hres = S_OK;
 CComVariant vData = NULL;
 CComPtr<IWbemClassObject> pTargetInstance; 
 CComBSTR bClassPropertyName = L"__CLASS";
 CComBSTR bTargetInstancePropertyName = L"TargetInstance";
 CComBSTR bProcessName = L"Name";
 CComBSTR bProcessId = L"ProcessId";
 wstring sClass;
 wstring sProcessName;
 DWORD dProcessId;

 for (int i = 0; i < lObjectCount; i++)
{
	hres = apObjArray[i]->Get(bClassPropertyName, 0, &vData, 0, 0);
	if (SUCCEEDED(hres))
	{
		sClass = vData.bstrVal;
		if (sClass.compare(L"__InstanceCreationEvent") ==0 || sClass.compare(L"__InstanceDeletionEvent") ==0)
		{
			//get the target instance property.
			hres = apObjArray[i]->Get(bTargetInstancePropertyName, 0, &vData, 0, 0);
			if (SUCCEEDED(hres))
			{
				//Obtained the TargetInstance property, now query fo the IWBEMCLASSOBJECT interface.				CComPtr<IWbemClassObject> pTargetInstance;
				IUnknown* str = vData.punkVal;
				hres = str->QueryInterface(IID_IWbemClassObject, reinterpret_cast< void** >(&pTargetInstance));
				if (SUCCEEDED(hres))
				{
					//Obtain the process name.
					hres =pTargetInstance->Get(bProcessName, 0, &vData, 0, 0);
					if (SUCCEEDED(hres))
					{
						sProcessName = vData.bstrVal;
						if (sClass.compare(L"__InstanceCreationEvent") == 0)
						{
							//We are creating a process therefore we need the process Id.
							hres = pTargetInstance->Get(bProcessId, 0, &vData, 0, 0);
							if (SUCCEEDED(hres))
							{
								dProcessId = vData.lVal;
								if (processCreatedCallback != nullptr)
								{
									processCreatedCallback(dProcessId, sProcessName.c_str());
}
								else 
								{ 
									MessageBox(NULL, L"The process created callback is null.", L"Process creation", MB_OK | MB_ICONERROR);
								}
							}
							else
							{
								MessageBox(NULL, L"Unable to obtain the process Id.", L"Process creation", MB_OK | MB_ICONERROR);
							}
						}
						else if (sClass.compare(L"__InstanceDeletionEvent") == 0)
						{
							//We are deleting a process.
							if (processDeletedCallback != nullptr)
							{
								processDeletedCallback(sProcessName.c_str());
							}
							else
							{
								MessageBox(NULL, L"The process deleted callback is null.", L"Process creation", MB_OK | MB_ICONERROR);
							}
						}
					}
					else
					{
						MessageBox(NULL, L"Unable to obtain the process name.", L"Process creation", MB_OK | MB_ICONERROR);
					}
				}
				else
				{
					MessageBox(NULL, L"Unable to query for the IWBEMClassObject interface.", L"Process creation", MB_OK | MB_ICONERROR);
				}
			}
			else
			{
				MessageBox(NULL, L"Unable to obtain the target instance property.", L"Process creation", MB_OK | MB_ICONERROR);
			}
		}
	}
	else
	{
		MessageBox(NULL, L"Unable to obtain the class of the event.", L"Process creation", MB_OK | MB_ICONERROR);
	}
 }    

return WBEM_S_NO_ERROR;
}

HRESULT ProcessMonitor::SetStatus(
            /* [in] */ LONG lFlags,
            /* [in] */ HRESULT hResult,
            /* [in] */ BSTR strParam,
            /* [in] */ IWbemClassObject __RPC_FAR *pObjParam
        )
{
    return WBEM_S_NO_ERROR;
}    

void ProcessMonitor::SetProcessCreatedCallback(TProcessCreatedCallback callback)
{
processCreatedCallback =callback;
}

void ProcessMonitor::SetProcessDeletedCallback(TProcessDeletedCallback callback)
{
	processDeletedCallback = callback;
}
// end of ProcessMonitor.cpp