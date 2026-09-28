#pragma once

#include <string>
#include <vector>

// 通过 Excel COM 将全部工作表页签导出为 PDF。
// pdfPath：合并后的主 PDF（尽量包含全部页签）。
// sheetPdfPaths：每个页签各自的临时 PDF（用于 PNG 逐表拼接）；调用方用完后应删除其中 != pdfPath 的文件。
// 本机需安装 Microsoft Excel。
bool ExportExcelToPdf(const std::wstring& excelPath, const std::wstring& pdfPath,
                      std::vector<std::wstring>& sheetPdfPaths, std::wstring& errorMsg);
