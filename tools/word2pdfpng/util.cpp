#include "util.h"

#include <shlwapi.h>
#include <windows.h>

#include <cstdio>
#include <cwctype>
#include <iterator>

#pragma comment(lib, "shlwapi.lib")

namespace util {

std::wstring GetFileBaseName(const std::wstring& fullPath) {
    wchar_t name[MAX_PATH]{};
    wchar_t ext[MAX_PATH]{};
    if (_wsplitpath_s(fullPath.c_str(), nullptr, 0, nullptr, 0, name, MAX_PATH, ext, MAX_PATH) != 0) {
        return L"output";
    }
    return std::wstring(name);
}

std::wstring GetDirectory(const std::wstring& fullPath) {
    wchar_t drive[_MAX_DRIVE]{};
    wchar_t dir[MAX_PATH]{};
    if (_wsplitpath_s(fullPath.c_str(), drive, _MAX_DRIVE, dir, MAX_PATH, nullptr, 0, nullptr, 0) !=
        0) {
        return L".\\";
    }
    std::wstring result = drive;
    result += dir;
    if (!result.empty() && result.back() != L'\\' && result.back() != L'/') {
        result += L'\\';
    }
    return result;
}

std::wstring GetFileName(const std::wstring& fullPath) {
    wchar_t fname[MAX_PATH]{};
    wchar_t ext[MAX_PATH]{};
    if (_wsplitpath_s(fullPath.c_str(), nullptr, 0, nullptr, 0, fname, MAX_PATH, ext, MAX_PATH) !=
        0) {
        return L"";
    }
    return std::wstring(fname) + ext;
}

std::wstring NormalizeAbsolutePath(const std::wstring& path) {
    if (path.empty()) {
        return path;
    }
    wchar_t buf[MAX_PATH * 4]{};
    const DWORD len = GetFullPathNameW(path.c_str(), static_cast<DWORD>(std::size(buf)), buf,
                                       nullptr);
    if (len == 0 || len >= std::size(buf)) {
        return path;
    }
    return buf;
}

std::wstring GetExtensionLower(const std::wstring& fullPath) {
    wchar_t ext[MAX_PATH]{};
    if (_wsplitpath_s(fullPath.c_str(), nullptr, 0, nullptr, 0, nullptr, 0, ext, MAX_PATH) != 0) {
        return L"";
    }
    for (wchar_t* p = ext; *p; ++p) {
        *p = static_cast<wchar_t>(towlower(*p));
    }
    return ext;
}

std::wstring ReplaceExtension(const std::wstring& fullPath, const std::wstring& newExt) {
    return GetDirectory(fullPath) + GetFileBaseName(fullPath) + newExt;
}

std::wstring MakePngPath(const std::wstring& wordPath) {
    return GetDirectory(wordPath) + GetFileBaseName(wordPath) + L".png";
}

std::wstring FormatHresult(HRESULT hr) {
    wchar_t buf[64]{};
    swprintf_s(buf, L"错误代码 0x%08X", static_cast<unsigned>(hr));
    return buf;
}

std::wstring FormatWin32(DWORD err) {
    wchar_t* msg = nullptr;
    const DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                        FORMAT_MESSAGE_IGNORE_INSERTS;
    if (FormatMessageW(flags, nullptr, err, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                       reinterpret_cast<LPWSTR>(&msg), 0, nullptr) &&
        msg) {
        std::wstring s(msg);
        LocalFree(msg);
        while (!s.empty() && (s.back() == L'\r' || s.back() == L'\n')) {
            s.pop_back();
        }
        return s;
    }
    return FormatHresult(HRESULT_FROM_WIN32(err));
}

}  // namespace util
