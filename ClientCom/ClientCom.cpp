
#include <windows.h>
#include <tchar.h>
#include <iostream>
#include <strsafe.h>
#include <comdef.h> 
#include "ITextReceiver.h"

int main() {
    std::cout << "Initializing COM" << std::endl;

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cerr << "Unable to initialize COM. HR = " << std::hex << hr << std::endl;
        return -1;
    }

    CLSID clsid = static_cast<CLSID>(GUID_TextReceiver);

    IClassFactory* pCF = nullptr;
    hr = CoGetClassObject(clsid, CLSCTX_LOCAL_SERVER, NULL, IID_IClassFactory, (void**)&pCF);
    if (FAILED(hr)) {
        std::cerr << "Failed to GetClassObject server instance. HR = " << std::hex << hr << std::endl;
        CoUninitialize();
        return -1;
    }

    IUnknown* pUnk = nullptr;
    hr = pCF->CreateInstance(NULL, IID_IUnknown, (void**)&pUnk);
    pCF->Release();

    if (FAILED(hr)) {
        std::cerr << "Failed to create server instance. HR = " << std::hex << hr << std::endl;
        CoUninitialize();
        return -1;
    }

    std::cout << "Instance created" << std::endl;

    ITextReceiver* pTextReceiver = nullptr;
    hr = pUnk->QueryInterface(IID_ITextReceiver, (LPVOID*)&pTextReceiver);
    pUnk->Release();

    if (FAILED(hr)) {
        std::cerr << "QueryInterface() for ITextReceiver failed. HR = " << std::hex << hr << std::endl;
        CoUninitialize();
        return -1;
    }

    BSTR message = SysAllocString(L"Hello, Andry!");
    hr = pTextReceiver->ReceiveText(message);
    SysFreeString(message);

    if (SUCCEEDED(hr)) {
        std::cout << "Message sent successfully" << std::endl;
    }
    else {
        std::cerr << "Failed to send message. HR = " << std::hex << hr << std::endl;
    }
    pTextReceiver->Release();

    std::cout << "Releasing instance" << std::endl;

    CoUninitialize();
    std::cout << "Shutting down COM" << std::endl;

    return 0;
}
