#include "excel_export.h"

#include "util.h"

#include <windows.h>

#include <comdef.h>
#include <objbase.h>
#include <oleauto.h>

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

namespace {

constexpr long kXlTypePdf = 0;
constexpr long kXlQualityStandard = 0;
constexpr long kXlSheetVisible = -1;

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

void ClearVariant(VARIANT& v) { VariantClear(&v); }

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

HRESULT SetPropertyI4(IDispatch* disp, LPCWSTR name, long value) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    DISPID putId = DISPID_PROPERTYPUT;
    VARIANT arg = MakeI4(value);
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

HRESULT GetLongProperty(IDispatch* disp, LPCWSTR name, long* out) {
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
    if (result.vt == VT_I4) {
        *out = result.lVal;
        VariantClear(&result);
        return S_OK;
    }
    if (result.vt == VT_I2) {
        *out = result.iVal;
        VariantClear(&result);
        return S_OK;
    }
    VARIANT coerced;
    VariantInit(&coerced);
    hr = VariantChangeType(&coerced, &result, 0, VT_I4);
    VariantClear(&result);
    if (FAILED(hr)) {
        return hr;
    }
    *out = coerced.lVal;
    VariantClear(&coerced);
    return S_OK;
}

HRESULT GetItemDispatch(IDispatch* collection, long index1Based, IDispatch** out) {
    *out = nullptr;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = index1Based;

    DISPPARAMS params{};
    params.rgvarg = &arg;
    params.cArgs = 1;

    VARIANT result;
    VariantInit(&result);

    // Excel 集合的 Item 是默认属性，用 DISPID_VALUE + PROPERTYGET（命名 "Item" 常返回 0x80020003）
    HRESULT hr = collection->Invoke(DISPID_VALUE, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        VariantInit(&result);
        hr = collection->Invoke(DISPID_VALUE, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET | DISPATCH_METHOD, &params, &result, nullptr,
                                nullptr);
    }
    if (FAILED(hr)) {
        // 少数环境仍暴露名为 Item 的方法
        VariantClear(&result);
        VariantInit(&result);
        DISPID id = 0;
        LPOLESTR names[1] = {const_cast<LPOLESTR>(L"Item")};
        hr = collection->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
        if (SUCCEEDED(hr)) {
            hr = collection->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params,
                                    &result, nullptr, nullptr);
            if (FAILED(hr)) {
                VariantClear(&result);
                VariantInit(&result);
                hr = collection->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                        &params, &result, nullptr, nullptr);
            }
        }
    }

    VariantClear(&arg);

    if (FAILED(hr)) {
        return hr;
    }
    if (result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return E_FAIL;
    }
    *out = result.pdispVal;  // 所有权转出，不再 Clear result 的 pdispVal
    return S_OK;
}

HRESULT SetPropertyDouble(IDispatch* disp, LPCWSTR name, double value) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    DISPID putId = DISPID_PROPERTYPUT;
    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_R8;
    arg.dblVal = value;
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

HRESULT GetBstrProperty(IDispatch* disp, LPCWSTR name, std::wstring& out) {
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
    if (result.vt == VT_BSTR && result.bstrVal) {
        out.assign(result.bstrVal, SysStringLen(result.bstrVal));
    } else {
        hr = E_FAIL;
    }
    VariantClear(&result);
    return hr;
}

HRESULT SetPropertyBstr(IDispatch* disp, LPCWSTR name, const std::wstring& value) {
    DISPID id = 0;
    HRESULT hr = GetDispId(disp, name, &id);
    if (FAILED(hr)) {
        return hr;
    }
    DISPID putId = DISPID_PROPERTYPUT;
    VARIANT arg = MakeBstr(value);
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

double InchesToPoints(IDispatch* app, double inches) {
    std::vector<VARIANT> args;
    VARIANT inchArg;
    VariantInit(&inchArg);
    inchArg.vt = VT_R8;
    inchArg.dblVal = inches;
    args.push_back(inchArg);
    VARIANT result;
    VariantInit(&result);
    const HRESULT hr = InvokeMethod(app, L"InchesToPoints", &result, args);
    for (auto& v : args) {
        ClearVariant(v);
    }
    if (FAILED(hr)) {
        VariantClear(&result);
        return inches * 72.0;
    }
    double pts = inches * 72.0;
    if (result.vt == VT_R8) {
        pts = result.dblVal;
    } else if (result.vt == VT_R4) {
        pts = result.fltVal;
    } else {
        VARIANT coerced;
        VariantInit(&coerced);
        if (SUCCEEDED(VariantChangeType(&coerced, &result, 0, VT_R8))) {
            pts = coerced.dblVal;
        }
        VariantClear(&coerced);
    }
    VariantClear(&result);
    return pts;
}

// 读取 Range 的 Row / Column（1-based）
bool GetRangeRowCol(IDispatch* range, long* row, long* col) {
    if (!range) {
        return false;
    }
    long r = 0;
    long c = 0;
    if (FAILED(GetLongProperty(range, L"Row", &r)) ||
        FAILED(GetLongProperty(range, L"Column", &c))) {
        return false;
    }
    if (row) {
        *row = r;
    }
    if (col) {
        *col = c;
    }
    return true;
}

bool GetRangeRowColCount(IDispatch* range, long* rows, long* cols) {
    if (!range) {
        return false;
    }
    long r = 0;
    long c = 0;
    IDispatch* rowsDisp = nullptr;
    IDispatch* colsDisp = nullptr;
    if (FAILED(GetDispatchProperty(range, L"Rows", &rowsDisp)) || !rowsDisp ||
        FAILED(GetLongProperty(rowsDisp, L"Count", &r))) {
        if (rowsDisp) {
            rowsDisp->Release();
        }
        return false;
    }
    rowsDisp->Release();
    if (FAILED(GetDispatchProperty(range, L"Columns", &colsDisp)) || !colsDisp ||
        FAILED(GetLongProperty(colsDisp, L"Count", &c))) {
        if (colsDisp) {
            colsDisp->Release();
        }
        return false;
    }
    colsDisp->Release();
    if (r < 1 || c < 1) {
        return false;
    }
    if (rows) {
        *rows = r;
    }
    if (cols) {
        *cols = c;
    }
    return true;
}

// 确保工作表内所有图形/图片可打印（浮动图 + Pictures 集合）
void EnsureShapesPrintable(IDispatch* sheet) {
    auto enableCollection = [](IDispatch* sheetObj, LPCWSTR propName) {
        IDispatch* coll = nullptr;
        if (FAILED(GetDispatchProperty(sheetObj, propName, &coll)) || !coll) {
            return;
        }
        long count = 0;
        if (FAILED(GetLongProperty(coll, L"Count", &count)) || count < 1) {
            coll->Release();
            return;
        }
        for (long i = 1; i <= count; ++i) {
            IDispatch* item = nullptr;
            if (FAILED(GetItemDispatch(coll, i, &item)) || !item) {
                continue;
            }
            SetPropertyI4(item, L"Visible", -1);  // msoTrue
            SetPropertyBool(item, L"PrintObject", true);
            // xlMoveAndSize = 1，随单元格打印
            SetPropertyI4(item, L"Placement", 1);
            item->Release();
        }
        coll->Release();
    };

    enableCollection(sheet, L"Shapes");
    enableCollection(sheet, L"Pictures");
    enableCollection(sheet, L"ChartObjects");
}

// 计算内容末行末列（UsedRange + 图形占位），避免 UsedRange 虚高导致页数暴增
bool ComputeContentExtent(IDispatch* sheet, long* outLastRow, long* outLastCol) {
    long lastRow = 1;
    long lastCol = 1;

    IDispatch* usedRange = nullptr;
    if (SUCCEEDED(GetDispatchProperty(sheet, L"UsedRange", &usedRange)) && usedRange) {
        long startRow = 1;
        long startCol = 1;
        long rowCount = 1;
        long colCount = 1;
        GetRangeRowCol(usedRange, &startRow, &startCol);
        GetRangeRowColCount(usedRange, &rowCount, &colCount);
        lastRow = startRow + rowCount - 1;
        lastCol = startCol + colCount - 1;
        // 限制异常虚高的 UsedRange（常见于曾经编辑到很远的单元格）
        if (lastRow > 5000) {
            lastRow = 5000;
        }
        if (lastCol > 100) {
            lastCol = 100;
        }
        usedRange->Release();
    }

    IDispatch* shapes = nullptr;
    if (SUCCEEDED(GetDispatchProperty(sheet, L"Shapes", &shapes)) && shapes) {
        long count = 0;
        if (SUCCEEDED(GetLongProperty(shapes, L"Count", &count))) {
            for (long i = 1; i <= count; ++i) {
                IDispatch* shape = nullptr;
                if (FAILED(GetItemDispatch(shapes, i, &shape)) || !shape) {
                    continue;
                }
                IDispatch* br = nullptr;
                if (SUCCEEDED(GetDispatchProperty(shape, L"BottomRightCell", &br)) && br) {
                    long r = 0;
                    long c = 0;
                    if (GetRangeRowCol(br, &r, &c)) {
                        if (r > lastRow) {
                            lastRow = r;
                        }
                        if (c > lastCol) {
                            lastCol = c;
                        }
                    }
                    br->Release();
                }
                shape->Release();
            }
        }
        shapes->Release();
    }

    if (outLastRow) {
        *outLastRow = lastRow < 1 ? 1 : lastRow;
    }
    if (outLastCol) {
        *outLastCol = lastCol < 1 ? 1 : lastCol;
    }
    return true;
}

HRESULT GetSheetRange(IDispatch* sheet, long row1, long col1, long row2, long col2,
                      IDispatch** outRange) {
    *outRange = nullptr;
    // Range(Cells(r1,c1), Cells(r2,c2))
    IDispatch* c1 = nullptr;
    IDispatch* c2 = nullptr;

    auto getCells = [&](long r, long c, IDispatch** cellOut) -> HRESULT {
        IDispatch* cells = nullptr;
        HRESULT hr = GetDispatchProperty(sheet, L"Cells", &cells);
        if (FAILED(hr) || !cells) {
            return FAILED(hr) ? hr : E_FAIL;
        }
        // Cells(row, col) —— IDispatch 逆序 [col, row]
        VARIANT vCol = MakeI4(c);
        VARIANT vRow = MakeI4(r);
        VARIANT args[2] = {vCol, vRow};
        DISPPARAMS params{};
        params.rgvarg = args;
        params.cArgs = 2;
        VARIANT result;
        VariantInit(&result);
        hr = cells->Invoke(DISPID_VALUE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, nullptr, nullptr);
        ClearVariant(vCol);
        ClearVariant(vRow);
        cells->Release();
        if (FAILED(hr) || result.vt != VT_DISPATCH || !result.pdispVal) {
            VariantClear(&result);
            return FAILED(hr) ? hr : E_FAIL;
        }
        *cellOut = result.pdispVal;
        return S_OK;
    };

    HRESULT hr = getCells(row1, col1, &c1);
    if (FAILED(hr)) {
        return hr;
    }
    hr = getCells(row2, col2, &c2);
    if (FAILED(hr)) {
        c1->Release();
        return hr;
    }

    DISPID rangeId = 0;
    hr = GetDispId(sheet, L"Range", &rangeId);
    if (FAILED(hr)) {
        c1->Release();
        c2->Release();
        return hr;
    }
    VARIANT a2;
    VariantInit(&a2);
    a2.vt = VT_DISPATCH;
    a2.pdispVal = c2;
    c2->AddRef();
    VARIANT a1;
    VariantInit(&a1);
    a1.vt = VT_DISPATCH;
    a1.pdispVal = c1;
    c1->AddRef();
    VARIANT args[2] = {a2, a1};
    DISPPARAMS params{};
    params.rgvarg = args;
    params.cArgs = 2;
    VARIANT result;
    VariantInit(&result);
    hr = sheet->Invoke(rangeId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params,
                       &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        VariantInit(&result);
        hr = sheet->Invoke(rangeId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params,
                           &result, nullptr, nullptr);
    }
    ClearVariant(a1);
    ClearVariant(a2);
    c1->Release();
    c2->Release();
    if (FAILED(hr) || result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return FAILED(hr) ? hr : E_FAIL;
    }
    *outRange = result.pdispVal;
    return S_OK;
}

void MarkRowsFromRange(IDispatch* range, std::vector<uint8_t>& used, long maxRow) {
    if (!range || maxRow < 1) {
        return;
    }

    auto markArea = [&](IDispatch* area) {
        if (!area) {
            return;
        }
        long startRow = 1;
        long startCol = 1;
        long rowCount = 1;
        long colCount = 1;
        if (!GetRangeRowCol(area, &startRow, &startCol) ||
            !GetRangeRowColCount(area, &rowCount, &colCount)) {
            return;
        }
        const long endRow = startRow + rowCount - 1;
        for (long r = startRow; r <= endRow && r <= maxRow; ++r) {
            if (r >= 1 && r < static_cast<long>(used.size())) {
                used[static_cast<size_t>(r)] = 1;
            }
        }
    };

    IDispatch* areas = nullptr;
    if (SUCCEEDED(GetDispatchProperty(range, L"Areas", &areas)) && areas) {
        long areaCount = 0;
        if (SUCCEEDED(GetLongProperty(areas, L"Count", &areaCount))) {
            for (long i = 1; i <= areaCount; ++i) {
                IDispatch* area = nullptr;
                if (SUCCEEDED(GetItemDispatch(areas, i, &area)) && area) {
                    markArea(area);
                    area->Release();
                }
            }
        }
        areas->Release();
    } else {
        markArea(range);
    }
}

HRESULT GetSpecialCells(IDispatch* range, long cellType, IDispatch** out) {
    *out = nullptr;
    std::vector<VARIANT> args;
    args.push_back(MakeI4(cellType));
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = InvokeMethod(range, L"SpecialCells", &result, args);
    for (auto& v : args) {
        ClearVariant(v);
    }
    if (FAILED(hr) || result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return FAILED(hr) ? hr : E_FAIL;
    }
    *out = result.pdispVal;
    return S_OK;
}

// 收集有内容的行（常量/公式/图形）
void CollectUsedRows(IDispatch* sheet, long lastRow, std::vector<uint8_t>& used) {
    if (lastRow < 1) {
        return;
    }
    used.assign(static_cast<size_t>(lastRow) + 1, 0);

    IDispatch* usedRange = nullptr;
    if (SUCCEEDED(GetDispatchProperty(sheet, L"UsedRange", &usedRange)) && usedRange) {
        IDispatch* special = nullptr;
        if (SUCCEEDED(GetSpecialCells(usedRange, 2, &special)) && special) {
            MarkRowsFromRange(special, used, lastRow);
            special->Release();
        }
        special = nullptr;
        if (SUCCEEDED(GetSpecialCells(usedRange, -4123, &special)) && special) {
            MarkRowsFromRange(special, used, lastRow);
            special->Release();
        }
        usedRange->Release();
    }

    IDispatch* shapes = nullptr;
    if (SUCCEEDED(GetDispatchProperty(sheet, L"Shapes", &shapes)) && shapes) {
        long scount = 0;
        if (SUCCEEDED(GetLongProperty(shapes, L"Count", &scount))) {
            for (long i = 1; i <= scount; ++i) {
                IDispatch* shape = nullptr;
                if (FAILED(GetItemDispatch(shapes, i, &shape)) || !shape) {
                    continue;
                }
                IDispatch* tl = nullptr;
                IDispatch* br = nullptr;
                GetDispatchProperty(shape, L"TopLeftCell", &tl);
                GetDispatchProperty(shape, L"BottomRightCell", &br);
                long r1 = 0;
                long r2 = 0;
                long c = 0;
                if (tl) {
                    GetRangeRowCol(tl, &r1, &c);
                    tl->Release();
                }
                if (br) {
                    GetRangeRowCol(br, &r2, &c);
                    br->Release();
                }
                if (r1 < 1) {
                    r1 = r2;
                }
                if (r2 < 1) {
                    r2 = r1;
                }
                if (r1 > r2) {
                    std::swap(r1, r2);
                }
                for (long r = r1; r <= r2 && r <= lastRow; ++r) {
                    if (r >= 1) {
                        used[static_cast<size_t>(r)] = 1;
                    }
                }
                shape->Release();
            }
        }
        shapes->Release();
    }
}

struct ContentBlock {
    long row1 = 1;
    long row2 = 1;
};

// 按连续内容行识别表格块；间隔超过 2 个空行视为新表
std::vector<ContentBlock> FindContentBlocks(IDispatch* sheet, long lastRow) {
    std::vector<ContentBlock> blocks;
    std::vector<uint8_t> used;
    CollectUsedRows(sheet, lastRow, used);
    if (lastRow < 1 || used.size() < 2) {
        return blocks;
    }

    long start = -1;
    long emptyRun = 0;
    for (long r = 1; r <= lastRow; ++r) {
        if (used[static_cast<size_t>(r)]) {
            if (start < 0) {
                start = r;
            }
            emptyRun = 0;
        } else if (start > 0) {
            ++emptyRun;
            // 超过 2 个空行：结束当前表格
            if (emptyRun > 2) {
                blocks.push_back({start, r - emptyRun});
                start = -1;
                emptyRun = 0;
            }
        }
    }
    if (start > 0) {
        long end = lastRow;
        while (end >= start && !used[static_cast<size_t>(end)]) {
            --end;
        }
        if (end >= start) {
            blocks.push_back({start, end});
        }
    }
    return blocks;
}

// 隐藏无含内容的空行，避免导出空白页（关闭工作簿时不保存，不改原文件）
void HideEmptyRows(IDispatch* sheet, long lastRow) {
    if (lastRow < 1) {
        return;
    }
    std::vector<uint8_t> used;
    CollectUsedRows(sheet, lastRow, used);

    bool any = false;
    for (long r = 1; r <= lastRow; ++r) {
        if (used[static_cast<size_t>(r)]) {
            any = true;
            break;
        }
    }
    if (!any) {
        return;
    }

    IDispatch* rows = nullptr;
    if (FAILED(GetDispatchProperty(sheet, L"Rows", &rows)) || !rows) {
        return;
    }
    for (long r = 1; r <= lastRow; ++r) {
        if (used[static_cast<size_t>(r)]) {
            continue;
        }
        IDispatch* row = nullptr;
        if (SUCCEEDED(GetItemDispatch(rows, r, &row)) && row) {
            SetPropertyBool(row, L"Hidden", true);
            row->Release();
        }
    }
    rows->Release();
}

// 列号转 Excel 列字母（1 -> A）
std::wstring ColumnIndexToLetters(long col) {
    std::wstring s;
    while (col > 0) {
        const long rem = (col - 1) % 26;
        s.insert(s.begin(), static_cast<wchar_t>(L'A' + rem));
        col = (col - 1) / 26;
    }
    return s;
}

// fitOnePageTall=true：整块表格缩放到一页高，避免表内分页
void ConfigureBlockPageSetup(IDispatch* sheet, IDispatch* app, long row1, long row2, long lastCol,
                             bool fitOnePageTall) {
    EnsureShapesPrintable(sheet);

    if (row1 < 1) {
        row1 = 1;
    }
    if (row2 < row1) {
        row2 = row1;
    }
    if (lastCol < 1) {
        lastCol = 1;
    }

    SetPropertyBool(app, L"PrintCommunication", false);

    IDispatch* pageSetup = nullptr;
    if (FAILED(GetDispatchProperty(sheet, L"PageSetup", &pageSetup)) || !pageSetup) {
        SetPropertyBool(app, L"PrintCommunication", true);
        return;
    }

    const std::wstring printArea = L"$A$" + std::to_wstring(row1) + L":$" +
                                   ColumnIndexToLetters(lastCol) + L"$" + std::to_wstring(row2);
    SetPropertyBstr(pageSetup, L"PrintArea", printArea);
    SetPropertyBool(pageSetup, L"Draft", false);

    SetPropertyBool(pageSetup, L"Zoom", false);
    SetPropertyI4(pageSetup, L"FitToPagesWide", 1);
    if (fitOnePageTall) {
        SetPropertyI4(pageSetup, L"FitToPagesTall", 1);
    } else {
        SetPropertyBool(pageSetup, L"FitToPagesTall", false);
    }

    const long rowSpan = row2 - row1 + 1;
    SetPropertyI4(pageSetup, L"Orientation", (lastCol > 8 || rowSpan > 60) ? 2 : 1);

    const double margin = InchesToPoints(app, 0.3);
    SetPropertyDouble(pageSetup, L"LeftMargin", margin);
    SetPropertyDouble(pageSetup, L"RightMargin", margin);
    SetPropertyDouble(pageSetup, L"TopMargin", margin);
    SetPropertyDouble(pageSetup, L"BottomMargin", margin);

    SetPropertyBool(pageSetup, L"CenterHorizontally", true);
    SetPropertyBool(pageSetup, L"CenterVertically", true);

    // 清除自动分页，避免表内被拆开
    std::vector<VARIANT> empty;
    VARIANT resetResult;
    VariantInit(&resetResult);
    InvokeMethod(sheet, L"ResetAllPageBreaks", &resetResult, empty);
    VariantClear(&resetResult);

    pageSetup->Release();
    SetPropertyBool(app, L"PrintCommunication", true);
}

std::wstring MakeTempPdfPath() {
    wchar_t tempDir[MAX_PATH]{};
    GetTempPathW(MAX_PATH, tempDir);
    wchar_t tempFile[MAX_PATH]{};
    if (GetTempFileNameW(tempDir, L"x2p", 0, tempFile) == 0) {
        return {};
    }
    DeleteFileW(tempFile);
    std::wstring path = tempFile;
    const size_t dot = path.find_last_of(L'.');
    if (dot != std::wstring::npos) {
        path = path.substr(0, dot) + L".pdf";
    } else {
        path += L".pdf";
    }
    return path;
}

HRESULT ExportSheetOrWorkbookToPdf(IDispatch* target, const std::wstring& pdfPath) {
    // ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas)
    // IgnorePrintAreas=False：使用我们设置的 PrintArea（完整内容且不虚高）
    std::vector<VARIANT> exportArgs;
    exportArgs.push_back(MakeBoolVariant(false));      // IgnorePrintAreas
    exportArgs.push_back(MakeBoolVariant(true));       // IncludeDocProperties
    exportArgs.push_back(MakeI4(kXlQualityStandard));  // Quality
    exportArgs.push_back(MakeBstr(pdfPath));           // Filename
    exportArgs.push_back(MakeI4(kXlTypePdf));          // Type

    VARIANT exportResult;
    VariantInit(&exportResult);
    HRESULT hr = InvokeMethod(target, L"ExportAsFixedFormat", &exportResult, exportArgs);
    for (auto& v : exportArgs) {
        ClearVariant(v);
    }
    VariantClear(&exportResult);
    return hr;
}

bool SelectAllWorksheets(IDispatch* worksheets, long count, std::wstring& errorMsg) {
    for (long i = 1; i <= count; ++i) {
        IDispatch* sheet = nullptr;
        HRESULT hr = GetItemDispatch(worksheets, i, &sheet);
        if (FAILED(hr) || !sheet) {
            errorMsg = L"无法获取工作表：" + util::FormatHresult(hr);
            return false;
        }
        SetPropertyI4(sheet, L"Visible", kXlSheetVisible);

        std::vector<VARIANT> selArgs;
        selArgs.push_back(MakeBoolVariant(i == 1));  // 首表替换，其后追加
        VARIANT selResult;
        VariantInit(&selResult);
        hr = InvokeMethod(sheet, L"Select", &selResult, selArgs);
        for (auto& v : selArgs) {
            ClearVariant(v);
        }
        VariantClear(&selResult);
        sheet->Release();
        if (FAILED(hr)) {
            std::vector<VARIANT> empty;
            VARIANT allResult;
            VariantInit(&allResult);
            hr = InvokeMethod(worksheets, L"Select", &allResult, empty);
            VariantClear(&allResult);
            if (FAILED(hr)) {
                errorMsg = L"无法选中全部工作表：" + util::FormatHresult(hr);
                return false;
            }
            return true;
        }
    }
    return true;
}

}  // namespace

bool ExportExcelToPdf(const std::wstring& excelPath, const std::wstring& pdfPath,
                      std::vector<std::wstring>& sheetPdfPaths, std::wstring& errorMsg) {
    sheetPdfPaths.clear();

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
    IDispatch* workbooks = nullptr;
    IDispatch* workbook = nullptr;
    IDispatch* worksheets = nullptr;
    long count = 0;

    CLSID clsid{};
    hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        errorMsg = L"未找到 Microsoft Excel。请安装 Excel 后重试。";
        goto cleanup;
    }

    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          reinterpret_cast<void**>(&app));
    if (FAILED(hr) || !app) {
        errorMsg = L"无法启动 Excel：" + util::FormatHresult(hr);
        goto cleanup;
    }

    SetPropertyBool(app, L"Visible", false);
    SetPropertyBool(app, L"DisplayAlerts", false);
    SetPropertyBool(app, L"ScreenUpdating", false);

    hr = GetDispatchProperty(app, L"Workbooks", &workbooks);
    if (FAILED(hr) || !workbooks) {
        errorMsg = L"无法访问 Excel 工作簿集合：" + util::FormatHresult(hr);
        goto cleanup;
    }

    {
        std::vector<VARIANT> openArgs;
        openArgs.push_back(MakeBstr(excelPath));
        VARIANT openResult;
        VariantInit(&openResult);
        hr = InvokeMethod(workbooks, L"Open", &openResult, openArgs);
        for (auto& v : openArgs) {
            ClearVariant(v);
        }
        if (FAILED(hr)) {
            errorMsg = L"无法打开 Excel 文件：" + util::FormatHresult(hr);
            goto cleanup;
        }
        if (openResult.vt != VT_DISPATCH || !openResult.pdispVal) {
            VariantClear(&openResult);
            errorMsg = L"打开 Excel 文件失败。";
            goto cleanup;
        }
        workbook = openResult.pdispVal;
    }

    hr = GetDispatchProperty(workbook, L"Worksheets", &worksheets);
    if (FAILED(hr) || !worksheets) {
        errorMsg = L"无法访问工作表集合：" + util::FormatHresult(hr);
        goto cleanup;
    }

    hr = GetLongProperty(worksheets, L"Count", &count);
    if (FAILED(hr) || count < 1) {
        errorMsg = L"工作簿中没有可导出的工作表。";
        goto cleanup;
    }

    // 按「表格块」导出：每块尽量 Fit 到一页，避免表内分页；多块分别导出再拼 PNG
    for (long i = 1; i <= count; ++i) {
        IDispatch* sheet = nullptr;
        hr = GetItemDispatch(worksheets, i, &sheet);
        if (FAILED(hr) || !sheet) {
            errorMsg = L"无法获取工作表：" + util::FormatHresult(hr);
            goto cleanup;
        }

        SetPropertyI4(sheet, L"Visible", kXlSheetVisible);

        long lastRow = 1;
        long lastCol = 1;
        ComputeContentExtent(sheet, &lastRow, &lastCol);

        std::vector<ContentBlock> blocks = FindContentBlocks(sheet, lastRow);
        if (blocks.empty()) {
            blocks.push_back({1, lastRow});
        }

        for (size_t bi = 0; bi < blocks.size(); ++bi) {
            const ContentBlock& block = blocks[bi];
            // 每块表格尽量缩放到一页（宽+高），减少表内分页
            ConfigureBlockPageSetup(sheet, app, block.row1, block.row2, lastCol, true);

            {
                std::vector<VARIANT> selArgs;
                selArgs.push_back(MakeBoolVariant(true));
                VARIANT selResult;
                VariantInit(&selResult);
                InvokeMethod(sheet, L"Select", &selResult, selArgs);
                for (auto& v : selArgs) {
                    ClearVariant(v);
                }
                VariantClear(&selResult);
            }

            const bool singleOutput = (count == 1 && blocks.size() == 1);
            const std::wstring partPdf = singleOutput ? pdfPath : MakeTempPdfPath();
            if (partPdf.empty()) {
                sheet->Release();
                errorMsg = L"无法创建临时 PDF 路径。";
                goto cleanup;
            }

            IDispatch* contentRange = nullptr;
            GetSheetRange(sheet, block.row1, 1, block.row2, lastCol, &contentRange);
            IDispatch* exportTarget = contentRange ? contentRange : sheet;
            hr = ExportSheetOrWorkbookToPdf(exportTarget, partPdf);
            if (contentRange) {
                contentRange->Release();
            }
            if (FAILED(hr)) {
                sheet->Release();
                errorMsg = L"导出工作表失败：" + util::FormatHresult(hr);
                goto cleanup;
            }
            sheetPdfPaths.push_back(partPdf);
        }

        // 多表格时再导出一份「合并 PDF」：不连续打印区域，每块各起一页
        if (count == 1 && blocks.size() > 1) {
            std::wstring unionArea;
            for (size_t bi = 0; bi < blocks.size(); ++bi) {
                if (bi > 0) {
                    unionArea += L",";
                }
                unionArea += L"$A$" + std::to_wstring(blocks[bi].row1) + L":$" +
                             ColumnIndexToLetters(lastCol) + L"$" +
                             std::to_wstring(blocks[bi].row2);
            }
            EnsureShapesPrintable(sheet);
            SetPropertyBool(app, L"PrintCommunication", false);
            IDispatch* pageSetup = nullptr;
            if (SUCCEEDED(GetDispatchProperty(sheet, L"PageSetup", &pageSetup)) && pageSetup) {
                SetPropertyBstr(pageSetup, L"PrintArea", unionArea);
                SetPropertyBool(pageSetup, L"Zoom", false);
                SetPropertyI4(pageSetup, L"FitToPagesWide", 1);
                // 多块时按页高自适应；每块已在临时 PDF 中尽量一页
                SetPropertyBool(pageSetup, L"FitToPagesTall", false);
                SetPropertyBool(pageSetup, L"CenterHorizontally", true);
                SetPropertyBool(pageSetup, L"CenterVertically", true);
                pageSetup->Release();
            }
            SetPropertyBool(app, L"PrintCommunication", true);
            std::vector<VARIANT> empty;
            VARIANT resetResult;
            VariantInit(&resetResult);
            InvokeMethod(sheet, L"ResetAllPageBreaks", &resetResult, empty);
            VariantClear(&resetResult);
            hr = ExportSheetOrWorkbookToPdf(sheet, pdfPath);
            if (FAILED(hr)) {
                CopyFileW(sheetPdfPaths.front().c_str(), pdfPath.c_str(), FALSE);
            }
        }

        sheet->Release();
    }

    // 多页签：再导出一份完整工作簿 PDF
    if (count > 1) {
        std::wstring selErr;
        if (SelectAllWorksheets(worksheets, count, selErr)) {
            hr = ExportSheetOrWorkbookToPdf(workbook, pdfPath);
            if (FAILED(hr)) {
                CopyFileW(sheetPdfPaths.front().c_str(), pdfPath.c_str(), FALSE);
            }
        } else {
            CopyFileW(sheetPdfPaths.front().c_str(), pdfPath.c_str(), FALSE);
        }
    }

    ok = true;

cleanup:
    if (!ok) {
        for (const auto& p : sheetPdfPaths) {
            if (p != pdfPath) {
                DeleteFileW(p.c_str());
            }
        }
        sheetPdfPaths.clear();
    }

    if (worksheets) {
        worksheets->Release();
    }
    if (workbook) {
        std::vector<VARIANT> closeArgs;
        closeArgs.push_back(MakeBoolVariant(false));
        VARIANT closeResult;
        VariantInit(&closeResult);
        InvokeMethod(workbook, L"Close", &closeResult, closeArgs);
        for (auto& v : closeArgs) {
            ClearVariant(v);
        }
        VariantClear(&closeResult);
        workbook->Release();
    }
    if (workbooks) {
        workbooks->Release();
    }
    if (app) {
        std::vector<VARIANT> quitArgs;
        VARIANT quitResult;
        VariantInit(&quitResult);
        InvokeMethod(app, L"Quit", &quitResult, quitArgs);
        VariantClear(&quitResult);
        app->Release();
    }
    if (comInited) {
        CoUninitialize();
    }

    return ok;
}
