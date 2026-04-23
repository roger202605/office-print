# Office Print - 指定页码范围打印

[![ClawHub](https://img.shields.io/badge/ClawHub-office--print-blue)](https://clawhub.ai/)
[![Version](https://img.shields.io/badge/version-1.1.0-green)]()
[![License](https://img.shields.io/badge/license-MIT-orange)]()

> 告诉 AI 打印哪个文件的哪几页，自动完成，不用手动操作打印对话框。

## 为什么需要这个 Skill？

日常办公最常见的操作之一就是"打印某几页"。但无论是 Word、Excel 还是 PPT，打印指定页码都需要：打开文件 → Ctrl+P → 找到页码输入框 → 输入 → 选单双面 → 点打印。听起来简单，但如果你用的是 AI 智能体（WorkBuddy/OpenClaw），它面对的是 Office COM API 的一个巨坑：

**`PrintOut` 的 `From/To/Pages` 参数全部失效——无论 Python 还是 VBS，传什么值都只打第1页。**

这个 Skill 用 COM+Tab 混合策略绕过了这个限制，让 AI 真正能帮你打印指定页。

## 支持的文件格式

| 格式 | 扩展名 | 状态 |
|------|--------|------|
| Word | .doc, .docx | ✅ 已验证 |
| PowerPoint | .ppt, .pptx | ✅ 已验证 |
| Excel | .xls, .xlsx | ✅ 已验证 |
| PDF | .pdf | 🔜 规划中 |

## 快速开始

### 安装

将 `office-print` 文件夹复制到 `~/.workbuddy/skills/` 目录即可。

### 使用

对 WorkBuddy 说：

```
帮我打印 report.docx 的第2到3页
打印 data.xlsx 第1-2页
把 slides.pptx 第3-5页单面打印出来
打印到 PDF 打印机
```

### 命令行

```bash
python print_pages.py <文件路径> <起始页> <结束页> [打印机名] [single/double]

# 示例
python print_pages.py doc.docx 2 3
python print_pages.py sheet.xlsx 1 2 "Microsoft Print to PDF"
python print_pages.py slides.pptx 3 5 "Ricoh SP 330" single
```

## 核心原理

### COM PrintOut 为什么不行？

经过 exhaustive 测试，以下所有方案均无法正确传递页码参数：

| 尝试方案 | 语言 | 结果 |
|----------|------|------|
| `PrintOut(Range=2, From="1", To="3")` | Python | 只打第1页 |
| `doc.PrintOut False, False, 2, "", "2", "3"` | VBS | 只打第1页 |
| `doc.PrintOut False, False, 3, ..., "2-3"` | VBS | 报"打印范围无效" |
| `wdPrintCurrentPage` + `Selection.GoTo` | VBS | 报"打印范围无效" |
| `ExportAsFixedFormat(Range=1, From=1, To=2)` | Python | 只导出1页 |

**结论：Word COM 的 PrintOut Range/From/To/Pages 参数在实际调用中均无法正确传递，无论 Python 还是 VBS。**

### COM+Tab 混合方案

| 职责 | 方法 | 原因 |
|------|------|------|
| 设置打印机 | COM | 可靠，Excel需端口名 |
| 打开文档 | COM | 含旧格式转换 |
| 输入页码 | Tab导航 | COM参数失效 |
| 选择单双面 | Tab导航 | COM无法控制 |
| 触发打印 | Tab导航 | 确保焦点正确 |

## 关键踩坑记录

### 1. Excel 的页码是两个独立输入框

Excel 的打印范围不是像 Word 那样在一个框里输入"1-2"，而是**两个独立的输入框**——"从第X页"和"至第Y页"，中间有个"至"字分隔。

正确操作：Tab5 → 输入起始页 → Tab一下 → 输入结束页

### 2. PDF 打印机没有单双面选项

PDF 虚拟打印机（如 Microsoft Print to PDF）的打印对话框中**不存在单双面选项**。如果在 PDF 打印机上尝试操作单双面，Tab 焦点会错位，导致后续操作全部失败。

解决方案：先判断打印机类型，PDF 打印机跳过单双面操作。

### 3. Excel 不要动 Tab4

Tab4 是打印范围下拉框（"打印活动工作表"/"打印整个工作簿"），**默认"活动工作表"就是对的**。选"自定义范围"反而会出问题。页码输入框在 Tab5 和 Tab6。

### 4. Excel 设打印机需要端口名

Excel COM 设置 ActivePrinter 时，必须带端口名：`"Ricoh SP 330 在 Ne00:"`，注意是中文"在"不是英文"on"。脚本会自动枚举 Ne00-Ne19 寻找正确端口。

### 5. Shift+Tab 回退次数因操作而异

- Word 不切单双面：Shift+Tab×5
- Word 切单双面：Shift+Tab×6
- Excel 不切单双面：Shift+Tab×6
- Excel 切单双面：Shift+Tab×7

## 依赖

- Python 3.8+
- pywin32（COM 自动化）
- pyautogui（键盘导航）
- Microsoft Office 2016+

## 许可证

MIT License
