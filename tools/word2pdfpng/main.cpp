// word2pdfpng — 选择 Word/Excel 文档，在同目录生成 PDF 与 PNG
// 作者：dianyin
// 用法：双击选择文件；或 word2pdfpng.exe "C:\path\file.xlsx"

#include "excel_export.h"
#include "pdf_to_png.h"
#include "util.h"
#include "word_export.h"

#include <windows.h>
#include <commdlg.h>
#include <shellapi.h>

#include <iterator>
#include <string>
#include <vector>

#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shell32.lib")

namespace {

const wchar_t kAppTitle[] = L"Word/Excel 转 PDF/PNG";

enum class OfficeKind { Word, Excel, Unknown };

OfficeKind DetectOfficeKind(const std::wstring& path) {
    const std::wstring ext = util::GetExtensionLower(path);
    if (ext == L".doc" || ext == L".docx" || ext == L".docm") {
        return OfficeKind::Word;
    }
    if (ext == L".xls" || ext == L".xlsx" || ext == L".xlsm") {
        return OfficeKind::Excel;
    }
    return OfficeKind::Unknown;
}

bool PickOfficeFile(HWND owner, std::wstring& path) {
    wchar_t fileBuf[MAX_PATH * 4]{};

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter =
        L"Office 文档 (*.doc;*.docx;*.xls;*.xlsx)\0*.doc;*.docx;*.xls;*.xlsx\0"
        L"Word 文档 (*.doc;*.docx)\0*.doc;*.docx\0"
        L"Excel 工作簿 (*.xls;*.xlsx;*.xlsm)\0*.xls;*.xlsx;*.xlsm\0"
        L"所有文件 (*.*)\0*.*\0\0";
    ofn.lpstrFile = fileBuf;
    ofn.nMaxFile = static_cast<DWORD>(std::size(fileBuf));
    ofn.lpstrTitle = L"请选择要转换的 Word 或 Excel 文件";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;

    if (!GetOpenFileNameW(&ofn)) {
        const DWORD err = CommDlgExtendedError();
        if (err != 0) {
            std::wstring msg = L"无法打开文件选择对话框：";
            msg += util::FormatWin32(err);
            MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONERROR);
        }
        return false;
    }
    path = util::NormalizeAbsolutePath(fileBuf);
    return true;
}

void ShowSuccess(HWND owner, const std::wstring& pdfPath, const std::vector<std::wstring>& pngPaths,
                 bool silent) {
    if (silent) {
        return;
    }
    std::wstring msg = L"转换完成。\n\nPDF：\n";
    msg += pdfPath;
    msg += L"\n\nPNG：\n";
    if (!pngPaths.empty()) {
        msg += pngPaths.front();
    }
    MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONINFORMATION);
}

void ShowError(HWND owner, const std::wstring& msg, bool silent) {
    if (silent) {
        return;
    }
    MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONERROR);
}

void CleanupTempPdfs(const std::vector<std::wstring>& sheetPdfs, const std::wstring& keepPdf) {
    for (const auto& p : sheetPdfs) {
        if (p != keepPdf) {
            DeleteFileW(p.c_str());
        }
    }
}

int RunConversion(HWND owner, const std::wstring* cmdPath) {
    const bool silent = cmdPath != nullptr && !cmdPath->empty();
    std::wstring officePath;
    if (silent) {
        officePath = util::NormalizeAbsolutePath(*cmdPath);
        if (GetFileAttributesW(officePath.c_str()) == INVALID_FILE_ATTRIBUTES) {
            ShowError(owner, L"文件不存在：" + officePath, false);
            return 1;
        }
    } else if (!PickOfficeFile(owner, officePath)) {
        return 0;
    }

    const OfficeKind kind = DetectOfficeKind(officePath);
    if (kind == OfficeKind::Unknown) {
        ShowError(owner, L"不支持的文件类型。请选择 .doc / .docx / .xls / .xlsx / .xlsm。", silent);
        return 1;
    }

    const std::wstring pdfPath = util::ReplaceExtension(officePath, L".pdf");

    HCURSOR waitCursor = LoadCursorW(nullptr, IDC_WAIT);
    HCURSOR oldCursor = SetCursor(waitCursor);

    std::wstring err;
    std::vector<std::wstring> pdfsForPng;

    if (kind == OfficeKind::Word) {
        if (!ExportWordToPdf(officePath, pdfPath, err)) {
            SetCursor(oldCursor);
            ShowError(owner, err, silent);
            return 1;
        }
        pdfsForPng.push_back(pdfPath);
    } else {
        std::vector<std::wstring> sheetPdfs;
        if (!ExportExcelToPdf(officePath, pdfPath, sheetPdfs, err)) {
            SetCursor(oldCursor);
            ShowError(owner, err, silent);
            return 1;
        }
        pdfsForPng = sheetPdfs.empty() ? std::vector<std::wstring>{pdfPath} : sheetPdfs;
    }

    std::vector<std::wstring> pngPaths;
    if (!ConvertPdfsToPngFiles(pdfsForPng, officePath, pngPaths, err)) {
        if (kind == OfficeKind::Excel) {
            CleanupTempPdfs(pdfsForPng, pdfPath);
        }
        SetCursor(oldCursor);
        if (!silent) {
            std::wstring msg = L"PDF 已生成，但 PNG 失败：\n";
            msg += err;
            msg += L"\n\nPDF 路径：\n";
            msg += pdfPath;
            MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONWARNING);
        }
        return 1;
    }

    if (kind == OfficeKind::Excel) {
        CleanupTempPdfs(pdfsForPng, pdfPath);
    }

    SetCursor(oldCursor);
    ShowSuccess(owner, pdfPath, pngPaths, silent);
    return 0;
}

}  // namespace

int WINAPI wWinMain(_In_ HINSTANCE, _In_opt_ HINSTANCE, _In_ PWSTR, _In_ int) {
    std::wstring cmdPath;
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv) {
        if (argc >= 2 && argv[1] && argv[1][0] != L'\0') {
            cmdPath = argv[1];
        }
        LocalFree(argv);
    }

    if (!cmdPath.empty()) {
        return RunConversion(nullptr, &cmdPath);
    }
    return RunConversion(nullptr, nullptr);
}
