#pragma once

#include <string>

// 通过 Word COM 将 .doc/.docx 导出为 PDF。本机需安装 Microsoft Word。
// pdfPath：输出 PDF 完整路径（宽字符，支持中文路径）
bool ExportWordToPdf(const std::wstring& wordPath, const std::wstring& pdfPath, std::wstring& errorMsg);
