#include <windows.h>
#include <iostream>
#include <strsafe.h>
#include <tchar.h>
#include "ClassFactory.h"


const CLSID CLSID_TextReceiver = static_cast<CLSID>(GUID_TextReceiver);

void GetModulePath(TCHAR* path, DWORD size) {
    GetModuleFileName(NULL, path, size);
}

HRESULT RegisterCOMServer() {
    TCHAR modulePath[MAX_PATH];
    GetModulePath(modulePath, MAX_PATH);

    HKEY hKey;
    LONG lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT, _T("CLSID\\{E2124D11-569D-4C83-8C6D-61434C0003B1}"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(lResult);
    }

    lResult = RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("TextReceiver COM Object"), (DWORD)(_tcslen(_T("TextReceiver COM Object")) + 1) * sizeof(TCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return HRESULT_FROM_WIN32(lResult);
    }

    HKEY hSubKey;
    lResult = RegCreateKeyEx(hKey, _T("LocalServer32"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return HRESULT_FROM_WIN32(lResult);
    }

    lResult = RegSetValueEx(hSubKey, NULL, 0, REG_SZ, (BYTE*)modulePath, (DWORD)(_tcslen(modulePath) + 1) * sizeof(TCHAR));
    RegCloseKey(hSubKey);
    RegCloseKey(hKey);

    return HRESULT_FROM_WIN32(lResult);
}

HRESULT UnregisterCOMServer() {
    LONG lResult = RegDeleteTree(HKEY_CLASSES_ROOT, _T("CLSID\\{E2124D11-569D-4C83-8C6D-61434C0003B1}"));
    return HRESULT_FROM_WIN32(lResult);
}

HRESULT RegisterInterface() {
    HKEY hKey;
    LONG lResult;

    lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT, _T("Interface\\{32bb8320-b41b-11cf-a6bb-0080c7b2d682}"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(lResult);
    }

    lResult = RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("ITextReceiver"), (DWORD)(_tcslen(_T("ITextReceiver")) + 1) * sizeof(TCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return HRESULT_FROM_WIN32(lResult);
    }

    HKEY hProxyStubClsid32Key;
    lResult = RegCreateKeyEx(hKey, _T("ProxyStubClsid32"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hProxyStubClsid32Key, NULL);
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return HRESULT_FROM_WIN32(lResult);
    }

    lResult = RegSetValueEx(hProxyStubClsid32Key, NULL, 0, REG_SZ, (BYTE*)_T("{00020424-0000-0000-C000-000000000046}"), (DWORD)(_tcslen(_T("{00020424-0000-0000-C000-000000000046}")) + 1) * sizeof(TCHAR));

    RegCloseKey(hProxyStubClsid32Key);
    RegCloseKey(hKey);

    return HRESULT_FROM_WIN32(lResult);
}

int main(int argc, char* argv[]) {
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::wcout << L"Failed to initialize COM. HRESULT: " << hr << std::endl;
        return -1;
    }
    if (argc > 1) {
        HRESULT hr;
        if (strcmp(argv[1], "/RegServer") == 0) {
            hr = RegisterCOMServer();
            if (SUCCEEDED(hr)) {
                std::wcout << L"Server registered successfully" << std::endl;
            }
            else {
                std::wcout << L"Failed to register server. HRESULT: " << hr << std::endl;
            }
            hr = RegisterInterface();
            if (SUCCEEDED(hr)) {
                std::wcout << L"Interface registered successfully" << std::endl;
            }
            else {
                std::wcout << L"Failed to register interface. HRESULT: " << hr << std::endl;
            }
        }

        if (strcmp(argv[1], "/UnregServer") == 0) {
            hr = UnregisterCOMServer();
            if (SUCCEEDED(hr)) {
                std::wcout << L"Server unregistered successfully" << std::endl;
            }
            else {
                std::wcout << L"Failed to unregister server. HRESULT: " << hr << std::endl;
            }
            
        }
    }
    
    CClassFactory* pFactory = new CClassFactory();
    DWORD regID = 0;
    hr = CoRegisterClassObject(CLSID_TextReceiver, static_cast<IClassFactory*>(pFactory),
        CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &regID);

    if (FAILED(hr)) {
        std::wcout << L"Failed to register class object. HRESULT: " << hr << std::endl;
        pFactory->Release();
        CoUninitialize();
        return -1;
    }

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    CoRevokeClassObject(regID);
    pFactory->Release();

    CoUninitialize();
    return 0;
}
