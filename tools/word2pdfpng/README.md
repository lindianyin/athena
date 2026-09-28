# word2pdfpng

Windows 桌面小工具：选择 `.doc` / `.docx` / `.xls` / `.xlsx`（及 `.xlsm`），在与源文件**相同目录**下生成同名 PDF 和 PNG。

## 依赖

- **Microsoft Word**（转换 Word 时需要，COM 导出 PDF）
- **Microsoft Excel**（转换 Excel 时需要，COM 导出 PDF）
- **Windows 10 或更高**（PNG 由系统 `Windows.Data.Pdf` + WIC 从 PDF 渲染）
- **MSVC**（Visual Studio 2019/2022，含「使用 C++ 的桌面开发」和 Windows 10 SDK）

## 编译

在 `tools/word2pdfpng` 目录下双击或命令行运行：

```bat
build.bat
```

成功后得到单个可执行文件 `word2pdfpng.exe`（静态链接 CRT `/MT`，无需额外 DLL）。

## 使用

1. 双击 `word2pdfpng.exe`
2. 在对话框中选择 Word 或 Excel 文件
3. 等待转换（鼠标为等待状态）
4. 完成后弹出消息框，显示 PDF 与 PNG 的完整路径

## 输出命名

| 类型 | 规则 |
|------|------|
| PDF | 与源文件同主文件名，扩展名为 `.pdf` |
| PNG | 与源文件同主文件名，扩展名为 `.png`（**始终一个文件**） |

多页/多工作表打印页时，各页按页码顺序**自上而下拼接**为一张竖向长图写入该 PNG；不会生成 `文件名_1.png`、`文件名_2.png` 等多文件。

- 画布宽度取各页渲染结果的最大宽度；较窄的页在白色背景上**水平居中**放置。
- 若拼接后的总高度超过约 **65500 像素**（极长文档），会按比例整体缩小后再编码，以避免内存与 PNG 编码器限制。

## 限制

- Excel 会导出**全部工作表页签**（隐藏表会先显示再导出）；PDF 尽量合并为一份，PNG 按页签顺序竖向拼接
- Excel 导出范围受页面设置影响；已设置 `IgnorePrintAreas` 尽量导出整表
- 转换 Word 需安装 Word；转换 Excel 需安装 Excel
- 首次调用 Office 可能较慢
- 路径支持中文；请确保对目标目录有写权限
- PNG 渲染分辨率默认约 **220 DPI**（可在 `pdf_to_png.cpp` 中修改 `kRenderDpi`）
- 页数很多或单页很高时，长图体积与内存占用会明显增大；触发高度上限时会自动缩小整图

## 目录结构

```
tools/word2pdfpng/
  main.cpp           入口、文件对话框、流程
  word_export.cpp    Word COM → PDF
  excel_export.cpp   Excel COM → PDF
  pdf_to_png.cpp     PDF → PNG（WinRT，多页竖向拼接）
  util.cpp           路径与错误信息
  build.bat          MSVC 编译脚本
  README.md
```
