#pragma once

#include <string>
#include <vector>

// 使用 Windows.Data.Pdf + WIC 将 PDF 渲染为 PNG（Win10+）；多页按顺序竖向拼接为一张长图
bool ConvertPdfToPngFiles(const std::wstring& pdfPath, const std::wstring& wordPathForNaming,
                          std::vector<std::wstring>& pngPaths, std::wstring& errorMsg);
