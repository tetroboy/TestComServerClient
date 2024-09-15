#pragma once

#include <Unknwn.h>
#include "TextReceiver.h"

class CClassFactory : public IClassFactory {
private:
    long m_refCount;

public:
    CClassFactory();

    
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG __stdcall AddRef() override;
    ULONG __stdcall Release() override;

    
    HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override;
    HRESULT __stdcall LockServer(BOOL fLock) override;
};
