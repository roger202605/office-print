# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-04-23

### Fixed
- Excel 页码输入从错误的"1-2"单框格式改为正确的双输入框（Tab5起始页 + Tab结束页）
- PDF 打印机不再尝试操作单双面选项（PDF打印对话框无此控件）
- Excel 不再误操作 Tab4（打印范围下拉框），保持默认"活动工作表"

### Added
- `is_pdf_printer()` 打印机类型判断
- 自动判断打印机类型决定是否操作单双面

## [1.0.0] - 2026-04-23

### Added
- PPT 指定页打印（Tab4选自定义范围 + Tab1输入页码）
- Excel 指定页打印（双输入框格式）
- Excel COM 打印机设置（枚举端口名）
- PDF 虚拟打印机另存为对话框处理（Alt+N + 剪贴板粘贴路径）
- 旧格式自动转换（.doc→.docx, .ppt→.pptx, .xls→.xlsx）
- 中文路径兼容（复制到 TEMP 英文路径）
- 单面/双面切换支持
- BlockInput 锁定用户输入防止干扰
- anti_sleep 阻止系统休眠

## [T0] - 2026-04-22

### Added
- Word 指定页打印基准版本（纯 Tab 键导航方案）
- Word 单面/双面切换
- 此版本永久冻结
