---
name: office-print
version: 1.1.0
description: |
  Word/PPT/Excel 指定页码范围打印，支持单双面切换、PDF虚拟打印机输出。
  触发词：打印Word指定页、打印PPT某几页、打印Excel第几页、打印页码范围、print word pages、
  Word打印、PPT打印、Excel打印、office打印、指定页打印、打印第几页、打印范围、
  单面打印、双面打印、打印到PDF、Word print、Excel print、PowerPoint print
author: Roger
tags: [打印, Word, Excel, PowerPoint, Office, PDF, 办公自动化]
license: MIT
minVersion: 2026.1.0
---

# Office Print - 指定页码范围打印

一句话：告诉 AI 打印哪个文件的哪几页，自动完成，不用手动操作打印对话框。

## 功能特点

- **Word 打印指定页**：输入"2-3"格式页码，Tab5定位页码输入框
- **PPT 打印指定页**：选"自定义范围"→输入页码，Tab4+Tab1定位
- **Excel 打印指定页**：两个独立输入框（起始页+结束页），Tab5→Tab定位
- **单面/双面切换**：自动判断打印机类型，PDF打印机跳过单双面选项
- **PDF虚拟打印机**：自动处理另存为对话框，输出到桌面
- **旧格式支持**：.doc/.ppt/.xls 自动转 .docx/.pptx/.xlsx
- **中文路径兼容**：自动复制到英文临时路径操作

## 使用方法

### 基础用法

```
用户：帮我打印 report.docx 的第2到3页
用户：打印 data.xlsx 第1-2页到 PDF 打印机
用户：把 slides.pptx 第3-5页单面打印出来
```

### 命令行调用

```bash
python print_pages.py <文件路径> <起始页> <结束页> [打印机名] [single/double]
```

### 示例

```bash
# Word 打印第2-3页（默认打印机）
python print_pages.py doc.docx 2 3

# Excel 打印第1-2页到 PDF 打印机
python print_pages.py sheet.xlsx 1 2 "Microsoft Print to PDF"

# PPT 打印第3-5页，单面
python print_pages.py slides.pptx 3 5 "Ricoh SP 330" single
```

## 核心原理

**为什么用 Tab 导航而不是 COM 参数？**

Office COM 的 PrintOut 方法，其 Range/From/To/Pages 参数在 Python pywin32 和 VBS 调用中均无法正确传递——无论传什么值，实际只打印第1页。这是 Office COM 的已知限制，不是代码 bug。

**解决方案：COM+Tab 混合策略**

| 职责 | 方法 | 说明 |
|------|------|------|
| 设置打印机 | COM | Word直接设，Excel需端口名 |
| 打开文档 | COM | 包括旧格式转换 |
| 输入页码 | Tab导航 | 因为COM参数失效 |
| 选择单双面 | Tab导航 | 仅物理打印机 |
| 触发打印 | Tab导航 | Shift+Tab回打印按钮 |

### 各软件 Tab 顺序

**Word (Ctrl+P 后):**

| Tab# | 控件 | 说明 |
|------|------|------|
| 0 | 打印按钮 | Enter直接打印 |
| 1 | 份数 | |
| 2 | 打印机 | |
| 3 | 打印属性 | |
| 4 | 打印范围 | |
| 5 | 页数输入框 | 输入"2-3"格式 |
| 6 | 单双面 | down→up→enter切单面 |

**PPT (Ctrl+P 后):**

| Tab# | 控件 | 说明 |
|------|------|------|
| 0 | 打印按钮 | |
| 1 | 份数 | |
| 2 | 打印机 | COM只读，需Tab导航 |
| 3 | 打印属性 | |
| 4 | 打印范围下拉 | Down×2+Enter选"自定义" |
| 5 | 页码输入框 | 输入"2-3"格式 |

**Excel (Ctrl+P 后):**

| Tab# | 控件 | 说明 |
|------|------|------|
| 0 | 打印按钮 | |
| 1 | 份数 | |
| 2 | 打印机 | |
| 3 | 打印属性 | |
| 4 | 打印范围 | 不动！默认"活动工作表" |
| 5 | 起始页输入框 | "从第X页" |
| 6 | 结束页输入框 | Tab一下，"至第Y页" |
| 7 | 单双面 | 仅物理打印机 |

## 配置要求

| 依赖 | 版本 | 说明 |
|------|------|------|
| Python | 3.8+ | 运行脚本 |
| pywin32 | 最新 | COM自动化 |
| pyautogui | 最新 | 键盘导航 |
| Microsoft Office | 2016+ | Word/Excel/PPT |

## 注意事项

- **操作期间勿动鼠标键盘**：脚本使用 BlockInput 锁定输入，防止干扰
- **PDF打印机没有单双面选项**：自动跳过，不会误操作
- **Excel不要动Tab4**：默认"打印活动工作表"即可，选自定义范围反而出错
- **Excel页码是两个独立框**：起始页和结束页分开输入，不是"1-2"格式
- **中文路径自动处理**：文件会复制到 TEMP 目录操作
- **旧格式自动转换**：.doc/.ppt/.xls 先转 .docx/.pptx/.xlsx

## 版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| T0 | 2026-04-22 | 基准版本，仅 Word 打印（冻结） |
| 1.0.0 | 2026-04-23 | 新增 PPT/Excel 支持，PDF打印机兼容 |
| 1.1.0 | 2026-04-23 | 修复 Excel 双输入框逻辑，PDF打印机跳过单双面 |
