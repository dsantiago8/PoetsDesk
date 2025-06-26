#include <windows.h>
#include <sapi.h>

void StartSpeechRecognition(HWND hEdit) {
    if (FAILED(::CoInitialize(NULL))) {
        MessageBox(NULL, L"Failed to initialize COM", L"Error", MB_OK);
        return;
    }

    ISpRecognizer* pRecognizer = nullptr;
    ISpRecoContext* pRecoContext = nullptr;
    ISpRecoGrammar* pGrammar = nullptr;

    HRESULT hr = CoCreateInstance(CLSID_SpSharedRecognizer, NULL, CLSCTX_ALL, IID_ISpRecognizer, (void**)&pRecognizer);
    if (SUCCEEDED(hr)) hr = pRecognizer->CreateRecoContext(&pRecoContext);

    if (SUCCEEDED(hr)) {
        pRecoContext->SetNotifyWin32Event();
        HANDLE hEvent = pRecoContext->GetNotifyEventHandle();

        ULONGLONG interest = (1ULL << SPEI_RECOGNITION) | ((1ULL << SPEI_RESERVED1) | (1ULL << SPEI_RESERVED2));
        pRecoContext->SetInterest(interest, interest);

        hr = pRecoContext->CreateGrammar(1, &pGrammar);
    }

    if (SUCCEEDED(hr)) {
        hr = pGrammar->LoadDictation(NULL, SPLO_STATIC);
        if (SUCCEEDED(hr)) hr = pGrammar->SetDictationState(SPRS_ACTIVE);
    }

    if (SUCCEEDED(hr)) {
        MessageBox(NULL, L"Speak now...", L"Speech", MB_OK);

        while (true) {
            WaitForSingleObject(pRecoContext->GetNotifyEventHandle(), INFINITE);
            SPEVENT evt;
            ULONG fetched = 0;

            while (SUCCEEDED(pRecoContext->GetEvents(1, &evt, &fetched)) && fetched > 0) {
                if (evt.eEventId == SPEI_RECOGNITION && evt.lParam) {
                    ISpRecoResult* pResult = reinterpret_cast<ISpRecoResult*>(evt.lParam);
                    pResult->AddRef();

                    WCHAR* pText = nullptr;
                    if (SUCCEEDED(pResult->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &pText, NULL))) {
                        SendMessageW(hEdit, EM_REPLACESEL, TRUE, (LPARAM)pText);
                        CoTaskMemFree(pText);
                    }

                    pResult->Release();
                }
            }
        }
    }

    // Cleanup (this will not be reached unless the loop exits, which it won't normally)
    if (pGrammar) pGrammar->Release();
    if (pRecoContext) pRecoContext->Release();
    if (pRecognizer) pRecognizer->Release();
    ::CoUninitialize();
}
