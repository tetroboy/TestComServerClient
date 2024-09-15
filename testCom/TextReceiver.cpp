#include "TextReceiver.h"
#include <iostream>

IID iid_ITextReceiver = static_cast<IID>(IID_ITextReceiver);

CTextReceiver::CTextReceiver() : m_refCount(1) {}

HRESULT __stdcall CTextReceiver::QueryInterface(REFIID riid, void** ppvObject) {
    if (riid == IID_IUnknown ) {
        
        *ppvObject = static_cast<ITextReceiver*>(this);
        AddRef();
        
        return S_OK;
    }
    else if (riid == IID_ITextReceiver) {
        *ppvObject = static_cast<ITextReceiver*>(this);
        AddRef();
        std::cout << "Text receiver \n";
        return S_OK;
    }
    LPOLESTR wszIID = nullptr;
    StringFromIID(riid, &wszIID);
    std::wcout << L"Requested IID: " << wszIID << std::endl;
    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall CTextReceiver::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

ULONG __stdcall CTextReceiver::Release() {
    ULONG ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) {
        delete this;
    }
    return ref;
}

HRESULT __stdcall CTextReceiver::ReceiveText(BSTR text) {
    MessageBoxW(NULL, text, L"Received Text", MB_OK);
    return S_OK;
}