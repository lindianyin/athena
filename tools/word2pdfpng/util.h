#pragma once

#include <windows.h>

#include <cstdint>
#include <string>
#include <vector>

namespace util {

// 宽字符串：取不含扩展名的文件名（不含路径）
std::wstring GetFileBaseName(const std::wstring& fullPath);

// 宽字符串：取所在目录（含盘符/UNC 前缀，含末尾反斜杠，除非是根路径）
std::wstring GetDirectory(const std::wstring& fullPath);

// 宽字符串：取文件名（含扩展名）
std::wstring GetFileName(const std::wstring& fullPath);

// 转为带盘符的绝对路径（失败则返回原字符串）
std::wstring NormalizeAbsolutePath(const std::wstring& path);

// 将扩展名替换为 newExt（应包含点，如 L".pdf"）
std::wstring ReplaceExtension(const std::wstring& fullPath, const std::wstring& newExt);

// 生成 PNG 路径：与 Word 同目录、同主文件名，扩展名为 .png（单页或多页长图均为一个文件）
std::wstring MakePngPath(const std::wstring& wordPath);

// HRESULT / Win32 错误转用户可读宽字符串（中文前缀）
std::wstring FormatHresult(HRESULT hr);
std::wstring FormatWin32(DWORD err);

}  // namespace util
