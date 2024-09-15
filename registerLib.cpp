#include <windows.h>
#include <iostream>
#include <OleAuto.h>
#include <filesystem>

HRESULT RegisterTypeLibrary(LPCOLESTR typelibPath) {
    ITypeLib* pTypeLib = nullptr;

    HRESULT hr = LoadTypeLib(typelibPath, &pTypeLib);
    if (FAILED(hr)) {
        std::wcout << L"Failed to load type library. HRESULT: " << hr << std::endl;
        return hr;
    }

    hr = RegisterTypeLib(pTypeLib, typelibPath, NULL);
    if (SUCCEEDED(hr)) {
        std::wcout << L"Type library registered successfully." << std::endl;
    }
    else {
        std::wcout << L"Failed to register type library. HRESULT: " << hr << std::endl;
    }

    if (pTypeLib) {
        pTypeLib->Release();
    }

    return hr;
}

int wmain() {
    try {
        wchar_t exePath[MAX_PATH];
        GetModuleFileName(NULL, exePath, MAX_PATH);
        std::filesystem::path currentPath = exePath;
        currentPath = currentPath.parent_path();
        std::filesystem::path typelibPath = currentPath / "idl files" / "TextReceiver.tlb";
        std::wcout << L"Type library path: " << typelibPath.c_str() << std::endl;

        HRESULT hr = CoInitialize(NULL);
        if (SUCCEEDED(hr)) {
            RegisterTypeLibrary(typelibPath.c_str());
            CoUninitialize();
        }
        else {
            std::wcout << L"Failed to initialize COM. HRESULT: " << hr << std::endl;
        }
    }
    catch (const std::filesystem::filesystem_error& ex) {
        std::wcerr << L"Filesystem error: " << ex.what() << std::endl;
    }
    catch (const std::exception& ex) {
        std::wcerr << L"General error: " << ex.what() << std::endl;
    }

    return 0;
}
