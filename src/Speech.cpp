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

    if (SUCCEEDED(CoCreateInstance(CLSID_SpSharedRecognizer, NULL, CLSCTX_ALL, IID_ISpRecognizer, (void**)&pRecognizer)) &&
        SUCCEEDED(pRecognizer->CreateRecoContext(&pRecoContext))) {

        pRecoContext->SetNotifyWin32Event();
        HANDLE hEvent = pRecoContext->GetNotifyEventHandle();
        ULONGLONG interest = SPFEI(SPEI_RECOGNITION);
        pRecoContext->SetInterest(interest, interest);
        pRecoContext->CreateGrammar(1, &pGrammar);

        if (SUCCEEDED(pGrammar->LoadDictation(NULL, SPLO_STATIC)) &&
            SUCCEEDED(pGrammar->SetDictationState(SPRS_ACTIVE))) {

            MessageBox(NULL, L"Speak now...", L"Speech", MB_OK);

            while (true) {
                WaitForSingleObject(hEvent, INFINITE);
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
                    // Manually clear event struct
                    ::CoTaskMemFree(evt.pwszSourceText);
                }
            }
        }
    }

    if (pGrammar) pGrammar->Release();
    if (pRecoContext) pRecoContext->Release();
    if (pRecognizer) pRecognizer->Release();
    ::CoUninitialize();
}
