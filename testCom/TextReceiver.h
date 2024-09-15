#pragma once
#include "ITextReceiver.h"
#include <windows.h>

class CTextReceiver : public ITextReceiver {
private:
    long m_refCount;

public:
    CTextReceiver();

   
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG __stdcall AddRef() override;
    ULONG __stdcall Release() override;

    HRESULT __stdcall ReceiveText(BSTR text) override;
};
