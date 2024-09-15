#include "ClassFactory.h"

CClassFactory::CClassFactory() : m_refCount(1) {}

HRESULT __stdcall CClassFactory::QueryInterface(REFIID riid, void** ppvObject) {
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppvObject = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall CClassFactory::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

ULONG __stdcall CClassFactory::Release() {
    ULONG ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) {
        delete this;
    }
    return ref;
}

HRESULT __stdcall CClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
    if (pUnkOuter != nullptr) {
        return CLASS_E_NOAGGREGATION;
    }

    CTextReceiver* pReceiver = new CTextReceiver();
    if (!pReceiver) {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = pReceiver->QueryInterface(riid, ppvObject);
    pReceiver->Release();

    return hr;
}

HRESULT __stdcall CClassFactory::LockServer(BOOL Lock) {
    if (Lock) {
        InterlockedIncrement(&m_refCount);
    }
    else {
        InterlockedDecrement(&m_refCount);
    }
    return S_OK;
}