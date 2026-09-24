#include "word_export.h"

#include "util.h"

#include <comdef.h>
#include <vector>
#include <objbase.h>
#include <oleauto.h>

namespace {

// Word ExportAsFixedFormat：wdExportFormatPDF = 17
constexpr long kWdExportFormatPdf = 17;

VARIANT MakeOptional() {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_ERROR;
    v.scode = DISP_E_PARAMNOTFOUND;
    return v;
}

VARIANT MakeBoolVariant(bool value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BOOL;
    v.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return v;
}

VARIANT MakeI4(long value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = value;
    return v;
}

VARIANT MakeBstr(const std::wstring& s) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocStringLen(s.c_str(), static_cast<UINT>(s.size()));
    return v;
}

void ClearVariant(VARIANT& v) {
    VariantClear(&v);
}

HRESULT GetDispId(IDispatch* disp, LPCWSTR name, DISPID* id) {
    LPOLESTR names[1] = {const_cast<LPOLESTR>(name)};
    return disp->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, id);
}

HRESULT SetPropertyBool(IDispatch* disp, LPCWSTR name, bool value) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    DISPID putId = DISPID_PROPERTYPUT;
    VARIANT arg = MakeBoolVariant(value);
    DISPPARAMS params{};
    params.rgvarg = &arg;
    params.cArgs = 1;
    params.rgdispidNamedArgs = &putId;
    params.cNamedArgs = 1;
    hr = disp->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, nullptr,
                      nullptr, nullptr);
    ClearVariant(arg);
    return hr;
}

HRESULT InvokeMethod(IDispatch* disp, LPCWSTR name, VARIANT* result, std::vector<VARIANT>& args) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    // IDispatch 参数为逆序
    DISPPARAMS params{};
    if (!args.empty()) {
        params.rgvarg = args.data();
        params.cArgs = static_cast<UINT>(args.size());
    }
    return disp->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, result,
                        nullptr, nullptr);
}

HRESULT GetDispatchProperty(IDispatch* disp, LPCWSTR name, IDispatch** out) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    VARIANT result;
    VariantInit(&result);
    DISPPARAMS params{};
    hr = disp->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &result,
                      nullptr, nullptr);
    if (FAILED(hr)) {
        return hr;
    }
    if (result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return E_FAIL;
    }
    *out = result.pdispVal;
    result.pdispVal->AddRef();
    VariantClear(&result);
    return S_OK;
}

}  // namespace

bool ExportWordToPdf(const std::wstring& wordPath, const std::wstring& pdfPath,
                     std::wstring& errorMsg) {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool comInited = SUCCEEDED(hr);
    if (hr == RPC_E_CHANGED_MODE) {
        errorMsg = L"COM 初始化失败：线程模型冲突。";
        return false;
    }
    if (FAILED(hr) && hr != S_FALSE) {
        errorMsg = L"COM 初始化失败：" + util::FormatHresult(hr);
        return false;
    }

    bool ok = false;
    IDispatch* app = nullptr;
    IDispatch* documents = nullptr;
    IDispatch* document = nullptr;

    CLSID clsid{};
    hr = CLSIDFromProgID(L"Word.Application", &clsid);
    if (FAILED(hr)) {
        errorMsg = L"未找到 Microsoft Word。请安装 Word 后重试。";
        goto cleanup;
    }

    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          reinterpret_cast<void**>(&app));
    if (FAILED(hr) || !app) {
        errorMsg = L"无法启动 Word：" + util::FormatHresult(hr);
        goto cleanup;
    }

    if (FAILED(SetPropertyBool(app, L"Visible", false))) {
        // 非致命
    }
    if (FAILED(SetPropertyBool(app, L"DisplayAlerts", false))) {
        // 非致命
    }

    hr = GetDispatchProperty(app, L"Documents", &documents);
    if (FAILED(hr) || !documents) {
        errorMsg = L"无法访问 Word 文档集合：" + util::FormatHresult(hr);
        goto cleanup;
    }

    {
        std::vector<VARIANT> openArgs;
        openArgs.push_back(MakeBstr(wordPath));  // FileName

        VARIANT openResult;
        VariantInit(&openResult);
        hr = InvokeMethod(documents, L"Open", &openResult, openArgs);
        for (auto& v : openArgs) {
            ClearVariant(v);
        }
        if (FAILED(hr)) {
            errorMsg = L"无法打开 Word 文档：" + util::FormatHresult(hr);
            goto cleanup;
        }
        if (openResult.vt != VT_DISPATCH || !openResult.pdispVal) {
            VariantClear(&openResult);
            errorMsg = L"打开 Word 文档失败。";
            goto cleanup;
        }
        document = openResult.pdispVal;
    }

    {
        std::vector<VARIANT> exportArgs;
        exportArgs.push_back(MakeI4(kWdExportFormatPdf));  // ExportFormat
        exportArgs.push_back(MakeBstr(pdfPath));           // OutputFileName

        VARIANT exportResult;
        VariantInit(&exportResult);
        hr = InvokeMethod(document, L"ExportAsFixedFormat", &exportResult, exportArgs);
        for (auto& v : exportArgs) {
            ClearVariant(v);
        }
        VariantClear(&exportResult);
        if (FAILED(hr)) {
            errorMsg = L"导出 PDF 失败：" + util::FormatHresult(hr);
            goto cleanup;
        }
    }

    ok = true;

cleanup:
    if (document) {
        std::vector<VARIANT> closeArgs;
        closeArgs.push_back(MakeI4(0));  // SaveChanges：wdDoNotSaveChanges
        VARIANT closeResult;
        VariantInit(&closeResult);
        InvokeMethod(document, L"Close", &closeResult, closeArgs);
        for (auto& v : closeArgs) {
            ClearVariant(v);
        }
        VariantClear(&closeResult);
        document->Release();
    }
    if (documents) {
        documents->Release();
    }
    if (app) {
        std::vector<VARIANT> quitArgs;
        quitArgs.push_back(MakeI4(0));  // SaveChanges：wdDoNotSaveChanges
        VARIANT quitResult;
        VariantInit(&quitResult);
        InvokeMethod(app, L"Quit", &quitResult, quitArgs);
        for (auto& v : quitArgs) {
            ClearVariant(v);
        }
        VariantClear(&quitResult);
        app->Release();
    }
    if (comInited) {
        CoUninitialize();
    }

    return ok;
}
