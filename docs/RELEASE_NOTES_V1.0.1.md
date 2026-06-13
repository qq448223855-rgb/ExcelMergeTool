# Excel 合并工具 V1.0.1

## 中文

本次版本重点优化了大文件处理速度、内存占用和 macOS 应用体积。

- 普通工作簿改用低内存流式合并。
- 包含合并单元格的工作簿自动使用兼容模式。
- 文件行数改用轻量 XLSX 元数据扫描。
- 重复单元格样式使用缓存，减少重复对象创建。
- 保持公式、公式文本、单元格格式、列宽和合并单元格。
- 精简未使用的 Qt 网络、TLS、SVG 和图片插件。
- 继续支持全部 Qt 基础系统语言。
- 当前版本支持 Apple 芯片 macOS。

性能测试使用 4 个工作簿，共 70,453 行、37 列。V1.0.0 基线约为
32.6 秒，V1.0.1 实测约 26.1 秒。文件信息读取由每个约 5 秒降至
约 0.2–0.3 秒。

## English

This release focuses on merge performance, memory usage, and macOS package
size.

- Uses low-memory streaming for ordinary workbooks.
- Automatically falls back to compatibility mode for merged cells.
- Reads row counts through lightweight XLSX metadata scanning.
- Caches repeated cell styles to reduce object creation.
- Preserves formulas, formula-like text, formatting, column widths, and merged
  cells.
- Removes unused Qt networking, TLS, SVG, and image plugins.
- Keeps all Qt base system-language translations.
- Supports Apple Silicon macOS.

The benchmark uses four workbooks with 70,453 rows and 37 columns. The V1.0.0
baseline is approximately 32.6 seconds, while V1.0.1 completes in about
26.1 seconds. File metadata loading drops from roughly five seconds per file
to approximately 0.2-0.3 seconds.
