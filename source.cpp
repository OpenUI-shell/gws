#include <iostream>
#include <fstream>
#include <windows.h>
#include <wbemidl.h>
#include <comdef.h>
#include <psapi.h>
#include <string>

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "psapi.lib")

// Global flag for program termination on console close
volatile bool g_bQuit = false;

// Console control handler
BOOL WINAPI ConsoleHandler(DWORD dwCtrlType) {
    if (dwCtrlType == CTRL_CLOSE_EVENT ||
        dwCtrlType == CTRL_C_EVENT ||
        dwCtrlType == CTRL_BREAK_EVENT ||
        dwCtrlType == CTRL_LOGOFF_EVENT ||
        dwCtrlType == CTRL_SHUTDOWN_EVENT)
    {
        g_bQuit = true;
        return TRUE;
    }
    return FALSE;
}

// Converts a TCHAR string to a std::wstring.
// If _UNICODE is defined, TCHAR is wchar_t and we construct directly;
// otherwise, we convert using MultiByteToWideChar.
std::wstring convertToWString(const TCHAR* str) {
#ifdef _UNICODE
    return std::wstring(str);
#else
    if (str == nullptr)
        return L"";
    int length = lstrlen(str);
    int requiredSize = MultiByteToWideChar(CP_ACP, 0, str, length, NULL, 0);
    std::wstring wideStr(requiredSize, L'\0');
    MultiByteToWideChar(CP_ACP, 0, str, length, &wideStr[0], requiredSize);
    return wideStr;
#endif
}

class ProcessEventSink : public IWbemObjectSink {
    LONG refCount;
    std::string appsJson; // Accumulates JSON content

public:
    ProcessEventSink() : refCount(1), appsJson("") {}

    ULONG STDMETHODCALLTYPE AddRef() {
        return InterlockedIncrement(&refCount);
    }
    ULONG STDMETHODCALLTYPE Release() {
        LONG count = InterlockedDecrement(&refCount);
        if (count == 0)
            delete this;
        return count;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == IID_IWbemObjectSink) {
            *ppv = static_cast<IWbemObjectSink*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    HRESULT STDMETHODCALLTYPE Indicate(LONG lObjectCount, IWbemClassObject** ppObjArray) {
        std::wcout << L"Event received!" << std::endl;

        for (LONG i = 0; i < lObjectCount; i++) {
            VARIANT vtInstance;
            HRESULT hr = ppObjArray[i]->Get(L"TargetInstance", 0, &vtInstance, nullptr, nullptr);
            if (FAILED(hr)) {
                std::wcout << L"Failed to retrieve TargetInstance." << std::endl;
                continue;
            }
            if ((vtInstance.vt == VT_UNKNOWN || vtInstance.vt == VT_DISPATCH) && vtInstance.punkVal != nullptr) {
                IWbemClassObject* pProcess = nullptr;
                hr = vtInstance.punkVal->QueryInterface(IID_IWbemClassObject, (void**)&pProcess);
                if (SUCCEEDED(hr) && pProcess) {
                    VARIANT vtPID;
                    hr = pProcess->Get(L"ProcessId", 0, &vtPID, nullptr, nullptr);
                    if (SUCCEEDED(hr)) {
                        DWORD pid = vtPID.uintVal;
                        VariantClear(&vtPID);

                        HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
                        if (hProcess) {
                            DWORD exitCode;
                            if (GetExitCodeProcess(hProcess, &exitCode) && exitCode == STILL_ACTIVE) {
                                TCHAR processName[MAX_PATH] = TEXT("<unknown>");
                                DWORD dwSize = MAX_PATH;
                                if (QueryFullProcessImageName(hProcess, 0, processName, &dwSize)) {
                                    std::wcout << L"New process detected:" << std::endl;
                                    std::wcout << L"  PID          : " << pid << std::endl;
                                    std::wcout << L"  Process Name : " << processName << std::endl;

                                    // Convert processName (TCHAR array) to std::wstring then to std::string
                                    std::wstring processNameW = convertToWString(processName);
                                    std::string processNameStr(processNameW.begin(), processNameW.end());

                                    // Append this process info as a JSON object using the PID as key.
                                    if (!appsJson.empty()) {
                                        appsJson += ",\n";
                                    }
                                    appsJson += "    \"" + std::to_string(pid) + "\": {\n";
                                    appsJson += "      \"pid\": " + std::to_string(pid) + ",\n";
                                    appsJson += "      \"name\": \"" + processNameStr + "\",\n";
                                    appsJson += "      \"GUI\": " + std::string(IsGuiProcess(hProcess) ? "true" : "false") + ",\n";
                                    appsJson += "      \"service\": " + std::string(IsService(hProcess) ? "true" : "false") + "\n";
                                    appsJson += "    }";
                                }
                            }
                            CloseHandle(hProcess);
                        }
                    } else {
                        std::wcout << L"Failed to retrieve ProcessId." << std::endl;
                    }
                    pProcess->Release();
                } else {
                    std::wcout << L"Failed to QI for IWbemClassObject from TargetInstance." << std::endl;
                }
            } else {
                std::wcout << L"TargetInstance is not an embedded object." << std::endl;
            }
            VariantClear(&vtInstance);
        }

        // Wrap and write JSON
        std::string fullJson = "{\n  \"apps\": {\n" + appsJson + "\n  }\n}\n";
        std::ofstream outFile("apps.json");
        if (outFile) {
            outFile << fullJson;
        } else {
            std::wcerr << L"Failed to open apps.json for writing." << std::endl;
        }
        return WBEM_S_NO_ERROR;
    }
    HRESULT STDMETHODCALLTYPE SetStatus(LONG, HRESULT, BSTR, IWbemClassObject*) {
        return WBEM_S_NO_ERROR;
    }
private:
    // Placeholder for GUI check
    bool IsGuiProcess(HANDLE hProcess) {
        // For this example, we simply check if there's a console window.
        HWND hwnd = GetConsoleWindow();
        return (hwnd != NULL);
    }
    // Placeholder for service check
    bool IsService(HANDLE hProcess) {
        return false;
    }
};

int main() {
    SetConsoleCtrlHandler(ConsoleHandler, TRUE);

    HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitializeEx failed." << std::endl;
        return 1;
    }
    IWbemLocator* locator = nullptr;
    IWbemServices* services = nullptr;
    hr = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&locator);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create WbemLocator instance." << std::endl;
        return 1;
    }
    hr = locator->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &services);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to connect to WMI server." << std::endl;
        locator->Release();
        return 1;
    }
    hr = CoSetProxyBlanket(services, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr,
                           RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to set proxy blanket." << std::endl;
        services->Release();
        locator->Release();
        return 1;
    }
    ProcessEventSink* eventSink = new ProcessEventSink();
    IUnsecuredApartment* apartment = nullptr;
    IWbemObjectSink* stubSink = nullptr;
    hr = CoCreateInstance(CLSID_UnsecuredApartment, nullptr, CLSCTX_LOCAL_SERVER,
                          IID_IUnsecuredApartment, (LPVOID*)&apartment);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create UnsecuredApartment instance." << std::endl;
        services->Release();
        locator->Release();
        return 1;
    }
    hr = apartment->CreateObjectStub(eventSink, (IUnknown**)&stubSink);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create object stub." << std::endl;
        apartment->Release();
        services->Release();
        locator->Release();
        return 1;
    }
    hr = services->ExecNotificationQueryAsync(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"),
        WBEM_FLAG_SEND_STATUS,
        nullptr,
        stubSink
    );
    if (FAILED(hr)) {
        std::wcerr << L"Failed to subscribe to WMI events." << std::endl;
        stubSink->Release();
        apartment->Release();
        services->Release();
        locator->Release();
        return 1;
    }
    std::wcout << L"Listening for new GUI processes..." << std::endl;
    while (!g_bQuit) {
        Sleep(100);
    }
    services->CancelAsyncCall(stubSink);
    stubSink->Release();
    apartment->Release();
    services->Release();
    locator->Release();
    eventSink->Release();
    CoUninitialize();
    std::wcout << L"Program terminated." << std::endl;
    return 0;
}
