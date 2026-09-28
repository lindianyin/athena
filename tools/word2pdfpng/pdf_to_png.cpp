#include "pdf_to_png.h"

#include "util.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>

#include <winrt/base.h>
#include <winrt/Windows.Data.Pdf.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Windows.Security.Cryptography.h>
#include <winrt/Windows.Storage.h>
#include <winrt/Windows.Storage.Streams.h>

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <vector>

using namespace winrt;
using namespace winrt::Windows::Data::Pdf;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Graphics::Imaging;
using namespace winrt::Windows::Security::Cryptography;
using namespace winrt::Windows::Storage;
using namespace winrt::Windows::Storage::Streams;

namespace {

constexpr int32_t kMaxStitchedHeight = 65500;
// PDF 页渲染分辨率（DIP 基准为 96）；提高到 220 DPI 以改善 PNG 清晰度
constexpr float kRenderDpi = 220.f;
constexpr float kDipDpi = 96.f;

struct BitmapView {
    const uint8_t* data = nullptr;
    uint32_t width = 0;
    uint32_t height = 0;
    uint32_t stride = 0;
};

SoftwareBitmap ToBgra8Premultiplied(const SoftwareBitmap& bitmap) {
    if (bitmap.BitmapPixelFormat() == BitmapPixelFormat::Bgra8 &&
        bitmap.BitmapAlphaMode() == BitmapAlphaMode::Premultiplied) {
        return bitmap;
    }
    return SoftwareBitmap::Convert(bitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
}

// 通过 CopyToBuffer + DataReader 取像素，避免对 BitmapBuffer 引用 QI IBufferByteAccess（E_NOINTERFACE）。
bool CopySoftwareBitmapPixels(const SoftwareBitmap& bitmap, std::vector<uint8_t>& pixels,
                                BitmapView& view) {
    const SoftwareBitmap bgra = ToBgra8Premultiplied(bitmap);
    const uint32_t width = static_cast<uint32_t>(bgra.PixelWidth());
    const uint32_t height = static_cast<uint32_t>(bgra.PixelHeight());
    if (width == 0 || height == 0) {
        return false;
    }

    const uint32_t byteCount = width * height * 4;
    Buffer buffer(byteCount);
    bgra.CopyToBuffer(buffer);

    pixels.resize(byteCount);
    DataReader reader = DataReader::FromBuffer(buffer);
    reader.ReadBytes(pixels);

    view.data = pixels.data();
    view.width = width;
    view.height = height;
    view.stride = width * 4;
    return true;
}

void BlitCopyCentered(const BitmapView& src, uint8_t* dst, uint32_t dstStride, int32_t canvasWidth,
                      int32_t dstY) {
    const int32_t offsetX = (canvasWidth - static_cast<int32_t>(src.width)) / 2;
    for (uint32_t y = 0; y < src.height; ++y) {
        const uint8_t* srcRow = src.data + static_cast<size_t>(y) * src.stride;
        uint8_t* dstRow =
            dst + static_cast<size_t>(dstY + y) * dstStride + static_cast<size_t>(offsetX) * 4;
        std::memcpy(dstRow, srcRow, static_cast<size_t>(src.width) * 4);
    }
}

void BlitScaleCentered(const BitmapView& src, uint8_t* dst, uint32_t dstStride, int32_t canvasWidth,
                       int32_t dstY, int32_t dstW, int32_t dstH) {
    if (dstW <= 0 || dstH <= 0) {
        return;
    }
    const int32_t offsetX = (canvasWidth - dstW) / 2;
    for (int32_t y = 0; y < dstH; ++y) {
        const int32_t sy = static_cast<int32_t>(static_cast<int64_t>(y) * src.height / dstH);
        const uint8_t* srcRow = src.data + static_cast<size_t>(sy) * src.stride;
        uint8_t* dstRow =
            dst + static_cast<size_t>(dstY + y) * dstStride + static_cast<size_t>(offsetX) * 4;
        for (int32_t x = 0; x < dstW; ++x) {
            const int32_t sx = static_cast<int32_t>(static_cast<int64_t>(x) * src.width / dstW);
            std::memcpy(dstRow + static_cast<size_t>(x) * 4, srcRow + static_cast<size_t>(sx) * 4, 4);
        }
    }
}

bool IsMostlyBlankPage(const SoftwareBitmap& bitmap) {
    std::vector<uint8_t> pixels;
    BitmapView view{};
    if (!CopySoftwareBitmapPixels(bitmap, pixels, view) || view.width == 0 || view.height == 0 ||
        !view.data) {
        return true;
    }

    uint64_t total = 0;
    uint64_t white = 0;
    // 抽样检测：步长加大，空白页几乎全白
    constexpr uint32_t kStepX = 16;
    constexpr uint32_t kStepY = 16;
    for (uint32_t y = 0; y < view.height; y += kStepY) {
        const uint8_t* row = view.data + static_cast<size_t>(y) * view.stride;
        for (uint32_t x = 0; x < view.width; x += kStepX) {
            const uint8_t* p = row + static_cast<size_t>(x) * 4;
            ++total;
            if (p[0] >= 250 && p[1] >= 250 && p[2] >= 250) {
                ++white;
            }
        }
    }
    if (total == 0) {
        return true;
    }
    // 超过 99.2% 近似白色视为空白页
    return (white * 1000) >= (total * 992);
}

bool StitchPagesVertically(const std::vector<SoftwareBitmap>& pages, SoftwareBitmap& stitched,
                           std::wstring& errorMsg) {
    if (pages.empty()) {
        errorMsg = L"\u6ca1\u6709\u53ef\u62fc\u63a5\u7684\u9875\u9762\u3002";
        return false;
    }

    int32_t maxWidth = 0;
    int64_t totalHeight = 0;
    for (const auto& page : pages) {
        maxWidth = (std::max)(maxWidth, page.PixelWidth());
        totalHeight += page.PixelHeight();
    }
    if (maxWidth <= 0 || totalHeight <= 0) {
        errorMsg = L"\u9875\u9762\u5c3a\u5bf8\u65e0\u6548\u3002";
        return false;
    }

    double stitchScale = 1.0;
    if (totalHeight > kMaxStitchedHeight) {
        stitchScale = static_cast<double>(kMaxStitchedHeight) / static_cast<double>(totalHeight);
    }

    const int32_t canvasWidth = (std::max)(
        1, static_cast<int32_t>(std::ceil(static_cast<double>(maxWidth) * stitchScale)));

    int32_t canvasHeight = 0;
    std::vector<int32_t> scaledHeights;
    scaledHeights.reserve(pages.size());
    for (const auto& page : pages) {
        const int32_t h = (std::max)(
            1, static_cast<int32_t>(std::round(static_cast<double>(page.PixelHeight()) * stitchScale)));
        scaledHeights.push_back(h);
        canvasHeight += h;
    }
    if (canvasHeight > kMaxStitchedHeight) {
        canvasHeight = kMaxStitchedHeight;
    }

    const uint32_t dstStride = static_cast<uint32_t>(canvasWidth) * 4;
    std::vector<uint8_t> canvasPixels(static_cast<size_t>(dstStride) * static_cast<size_t>(canvasHeight),
                                        0xFF);

    const bool useScale = stitchScale < 0.999999;
    int32_t y = 0;
    std::vector<uint8_t> pagePixels;
    pagePixels.reserve(static_cast<size_t>(maxWidth) * 4);

    for (size_t i = 0; i < pages.size(); ++i) {
        BitmapView src{};
        if (!CopySoftwareBitmapPixels(pages[i], pagePixels, src)) {
            errorMsg = L"\u65e0\u6cd5\u8bfb\u53d6\u9875\u9762\u50cf\u7d20\u3002";
            return false;
        }

        const int32_t dstH = scaledHeights[i];
        if (y + dstH > canvasHeight) {
            break;
        }
        if (useScale) {
            const int32_t dstW = (std::max)(
                1, static_cast<int32_t>(std::round(static_cast<double>(src.width) * stitchScale)));
            BlitScaleCentered(src, canvasPixels.data(), dstStride, canvasWidth, y, dstW, dstH);
        } else {
            BlitCopyCentered(src, canvasPixels.data(), dstStride, canvasWidth, y);
        }
        y += dstH;
    }

    const auto outBuffer = CryptographicBuffer::CreateFromByteArray(canvasPixels);
    stitched = SoftwareBitmap::CreateCopyFromBuffer(outBuffer, BitmapPixelFormat::Bgra8, canvasWidth,
                                                    canvasHeight, BitmapAlphaMode::Premultiplied);
    return true;
}

// 将 SoftwareBitmap 编码为 PNG 字节流（内存），避免直接写文件时 Flush 未落盘导致 0 字节文件。
bool EncodeSoftwareBitmapToPngStream(const SoftwareBitmap& bitmap, InMemoryRandomAccessStream& pngStream,
                                     std::wstring& errorMsg) {
    SoftwareBitmap encodeBitmap = bitmap;
    if (bitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
        bitmap.BitmapAlphaMode() != BitmapAlphaMode::Premultiplied) {
        encodeBitmap = SoftwareBitmap::Convert(bitmap, BitmapPixelFormat::Bgra8,
                                               BitmapAlphaMode::Premultiplied);
    }

    pngStream = InMemoryRandomAccessStream{};
    BitmapEncoder encoder =
        BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), pngStream).get();
    encoder.SetSoftwareBitmap(encodeBitmap);
    encoder.FlushAsync().get();

    pngStream.Seek(0);
    if (pngStream.Size() == 0) {
        errorMsg = L"\u0050\u004e\u0047 \u7f16\u7801\u7ed3\u679c\u4e3a\u7a7a\u3002";
        return false;
    }
    return true;
}

bool WriteStreamToStorageFile(IRandomAccessStream const& dataStream, const StorageFile& outFile,
                              std::wstring& errorMsg) {
    dataStream.Seek(0);
    const uint64_t byteCount = dataStream.Size();
    if (byteCount == 0) {
        errorMsg = L"\u65e0\u6cd5\u5199\u5165\u7a7a\u7684 PNG \u6570\u636e\u3002";
        return false;
    }

    auto outStream = outFile.OpenAsync(FileAccessMode::ReadWrite).get();
    outStream.Size(0);
    outStream.Seek(0);
    RandomAccessStream::CopyAndCloseAsync(dataStream, outStream).get();
    return true;
}

}  // namespace

bool ConvertPdfToPngFiles(const std::wstring& pdfPath, const std::wstring& wordPathForNaming,
                          std::vector<std::wstring>& pngPaths, std::wstring& errorMsg) {
    return ConvertPdfsToPngFiles({pdfPath}, wordPathForNaming, pngPaths, errorMsg);
}

bool ConvertPdfsToPngFiles(const std::vector<std::wstring>& pdfPaths,
                           const std::wstring& pathForNaming, std::vector<std::wstring>& pngPaths,
                           std::wstring& errorMsg) {
    pngPaths.clear();
    if (pdfPaths.empty()) {
        errorMsg = L"没有可转换的 PDF。";
        return false;
    }

    try {
        init_apartment(apartment_type::single_threaded);
    } catch (...) {
    }

    try {
        std::vector<SoftwareBitmap> pageBitmaps;

        for (const auto& pdfPath : pdfPaths) {
            StorageFile pdfFile = StorageFile::GetFileFromPathAsync(pdfPath).get();
            PdfDocument pdfDoc = PdfDocument::LoadFromFileAsync(pdfFile).get();
            const uint32_t pageCount = pdfDoc.PageCount();
            if (pageCount == 0) {
                continue;
            }

            for (uint32_t i = 0; i < pageCount; ++i) {
                PdfPage page = pdfDoc.GetPage(i);
                InMemoryRandomAccessStream bmpStream;
                PdfPageRenderOptions options;
                const auto pageSize = page.Size();
                const uint32_t destW = (std::max)(
                    1u, static_cast<uint32_t>(
                            std::lround(pageSize.Width * kRenderDpi / kDipDpi)));
                const uint32_t destH = (std::max)(
                    1u, static_cast<uint32_t>(
                            std::lround(pageSize.Height * kRenderDpi / kDipDpi)));
                options.DestinationWidth(destW);
                options.DestinationHeight(destH);
                page.RenderToStreamAsync(bmpStream, options).get();
                bmpStream.Seek(0);

                BitmapDecoder decoder = BitmapDecoder::CreateAsync(bmpStream).get();
                SoftwareBitmap bitmap =
                    ToBgra8Premultiplied(decoder.GetSoftwareBitmapAsync().get());
                // 跳过几乎全白的空白页，避免长图中间大片留白
                if (!IsMostlyBlankPage(bitmap)) {
                    pageBitmaps.push_back(bitmap);
                }
                page.Close();
            }
        }

        if (pageBitmaps.empty()) {
            errorMsg = L"PDF 没有可渲染的页面。";
            return false;
        }

        SoftwareBitmap stitched{nullptr};
        if (!StitchPagesVertically(pageBitmaps, stitched, errorMsg)) {
            return false;
        }

        InMemoryRandomAccessStream pngStream;
        if (!EncodeSoftwareBitmapToPngStream(stitched, pngStream, errorMsg)) {
            return false;
        }

        const std::wstring pngPath = util::MakePngPath(pathForNaming);
        const std::wstring dir = util::GetDirectory(pngPath);
        const std::wstring fileName = util::GetFileName(pngPath);
        StorageFolder folder = StorageFolder::GetFolderFromPathAsync(dir).get();
        StorageFile outFile =
            folder.CreateFileAsync(fileName, CreationCollisionOption::ReplaceExisting).get();

        if (!WriteStreamToStorageFile(pngStream, outFile, errorMsg)) {
            return false;
        }

        pngPaths.push_back(pngPath);
        return true;
    } catch (const winrt::hresult_error& e) {
        errorMsg = L"PDF 转 PNG 失败：" + std::wstring(e.message().c_str());
        return false;
    } catch (...) {
        errorMsg = L"PDF 转 PNG 时发生未知错误。";
        return false;
    }
}
