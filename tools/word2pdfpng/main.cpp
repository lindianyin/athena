// word2pdfpng — 选择 Word 文档，在同目录生成 PDF 与 PNG
// 作者：dianyin

#include "pdf_to_png.h"
#include "util.h"
#include "word_export.h"

#include <windows.h>
#include <commdlg.h>

#include <iterator>
#include <string>
#include <vector>

#pragma comment(lib, "comdlg32.lib")

namespace {

const wchar_t kAppTitle[] = L"Word 转 PDF/PNG";

// 弹出打开文件对话框，返回是否选择了文件
bool PickWordFile(HWND owner, std::wstring& path) {
    wchar_t fileBuf[MAX_PATH * 4]{};  // 预留较长路径

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"Word 文档 (*.doc;*.docx)\0*.doc;*.docx\0所有文件 (*.*)\0*.*\0\0";
    ofn.lpstrFile = fileBuf;
    ofn.nMaxFile = static_cast<DWORD>(std::size(fileBuf));
    ofn.lpstrTitle = L"请选择要转换的 Word 文件";
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

void ShowSuccess(HWND owner, const std::wstring& pdfPath, const std::vector<std::wstring>& pngPaths) {
    std::wstring msg = L"转换完成。\n\nPDF：\n";
    msg += pdfPath;
    msg += L"\n\nPNG：\n";
    if (!pngPaths.empty()) {
        msg += pngPaths.front();
    }
    MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONINFORMATION);
}

void ShowError(HWND owner, const std::wstring& msg) {
    MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONERROR);
}

int RunConversion(HWND owner) {
    std::wstring wordPath;
    if (!PickWordFile(owner, wordPath)) {
        return 0;  // 用户取消，不提示
    }

    const std::wstring pdfPath = util::ReplaceExtension(wordPath, L".pdf");

    HCURSOR waitCursor = LoadCursorW(nullptr, IDC_WAIT);
    HCURSOR oldCursor = SetCursor(waitCursor);

    std::wstring err;
    if (!ExportWordToPdf(wordPath, pdfPath, err)) {
        SetCursor(oldCursor);
        ShowError(owner, err);
        return 1;
    }

    std::vector<std::wstring> pngPaths;
    if (!ConvertPdfToPngFiles(pdfPath, wordPath, pngPaths, err)) {
        SetCursor(oldCursor);
        std::wstring msg = L"PDF 已生成，但 PNG 失败：\n";
        msg += err;
        msg += L"\n\nPDF 路径：\n";
        msg += pdfPath;
        MessageBoxW(owner, msg.c_str(), kAppTitle, MB_OK | MB_ICONWARNING);
        return 1;
    }

    SetCursor(oldCursor);
    ShowSuccess(owner, pdfPath, pngPaths);
    return 0;
}

}  // namespace

int WINAPI wWinMain(_In_ HINSTANCE, _In_opt_ HINSTANCE, _In_ PWSTR, _In_ int) {
    return RunConversion(nullptr);
}
