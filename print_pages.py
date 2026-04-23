"""
Office Print T1 v2 - COM+Tab 混合方案
支持: Word(.doc/.docx), PPT(.ppt/.pptx), Excel(.xls/.xlsx)
打印机: 任意，默认 Ricoh SP 330
单双面: double(默认)/single (仅物理打印机有效，PDF打印机无此选项)

版本历史:
  T1 v1: Word+PPT初步版本，Excel未验证
  T1 v2: Excel验证完成，修复页码输入(两个独立框)，PDF打印机不操作单双面

核心思路:
  COM 负责能做的: 设置打印机、打开文档、关闭文档、格式转换
  Tab 负责必须做的: 输入页码、选择单双面（仅物理打印机）
  
  关键: Excel的页码是两个独立输入框(从+至)，不是Word/PPT的"1-2"格式

用法: python print_pages.py <文件路径> <起始页> <结束页> [打印机名] [single/double]
"""
import sys
import os
import time
import shutil
import subprocess
import ctypes
import pyautogui
from datetime import datetime

pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = False

PRINTER_DEFAULT = "Ricoh SP 330"
TEMP_BASE = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'office_print')


# ══════════════════════════════════════════════════════════════
# 基础工具
# ══════════════════════════════════════════════════════════════

def kill_office():
    """清理残留的 Office 进程"""
    for proc in ['WINWORD.EXE', 'POWERPNT.EXE', 'EXCEL.EXE', 'Acrobat.exe']:
        subprocess.run(['taskkill', '/F', '/IM', proc],
                       capture_output=True, timeout=10)
    time.sleep(0.5)


def block_input():
    try:
        ctypes.windll.user32.BlockInput(True)
    except:
        pass


def unblock_input():
    try:
        ctypes.windll.user32.BlockInput(False)
    except:
        pass


def anti_sleep():
    """阻止系统休眠"""
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)


def force_foreground(hwnd):
    """强制窗口到前台"""
    import win32gui, win32con
    fg = win32gui.GetForegroundWindow()
    fg_tid = ctypes.windll.user32.GetWindowThreadProcessId(fg, None)
    tg_tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
    if fg_tid != tg_tid:
        ctypes.windll.user32.AttachThreadInput(fg_tid, tg_tid, True)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.2)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(0.3)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except:
        win32gui.BringWindowToTop(hwnd)
    if fg_tid != tg_tid:
        ctypes.windll.user32.AttachThreadInput(fg_tid, tg_tid, False)
    time.sleep(0.5)


def find_window_by_class(cls_name):
    """按窗口类名查找窗口句柄"""
    import win32gui
    result = [None]
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == cls_name:
            result[0] = hwnd
    win32gui.EnumWindows(cb, None)
    return result[0]


def is_pdf_printer(printer_name):
    """判断是否是PDF虚拟打印机（PDF打印机没有单双面选项）"""
    if not printer_name:
        return False
    pdf_keywords = ['PDF', 'XPS', 'OneNote', 'Fax']
    return any(kw in printer_name.upper() for kw in pdf_keywords)


def set_excel_printer(excel, printer_name):
    """Excel设置打印机需要带端口名: '打印机名 在 NeXX:'"""
    if not printer_name:
        return False
    for i in range(20):
        port_name = f"Ne{i:02d}"
        test_name = f"{printer_name} \u5728 {port_name}:"  # "在"不是"on"
        try:
            excel.ActivePrinter = test_name
            print(f"[Excel] printer set: {port_name}")
            return True
        except:
            continue
    print(f"[Excel] WARN: cannot set printer {printer_name}")
    return False


def prepare_file(file_path):
    """复制文件到英文临时路径，.doc/.ppt/.xls 自动转新格式"""
    os.makedirs(TEMP_BASE, exist_ok=True)
    ext = os.path.splitext(file_path)[1].lower()
    temp_file = os.path.join(TEMP_BASE, f'source{ext}')
    shutil.copy2(file_path, temp_file)

    # 旧格式转换
    format_map = {
        '.doc': ('Word.Application', 16),   # wdFormatXMLDocument
        '.ppt': ('PowerPoint.Application', 24),  # ppSaveAsOpenXMLPresentation
        '.xls': ('Excel.Application', 51),  # xlOpenXMLWorkbook=51
    }

    if ext in format_map:
        app_name, fmt = format_map[ext]
        new_ext = 'x' + ext
        temp_new = os.path.join(TEMP_BASE, f'source{new_ext}')
        import win32com.client
        app = win32com.client.Dispatch(app_name)
        if app_name == 'Word.Application':
            app.Visible = False
            d = app.Documents.Open(temp_file)
            d.SaveAs(temp_new, FileFormat=fmt)
            d.Close(SaveChanges=False)
        elif app_name == 'PowerPoint.Application':
            pres = app.Presentations.Open(temp_file)
            pres.SaveAs(temp_new, FileFormat=fmt)
            pres.Close()
        elif app_name == 'Excel.Application':
            app.Visible = False
            wb = app.Workbooks.Open(temp_file)
            wb.SaveAs(temp_new, FileFormat=fmt)
            wb.Close(SaveChanges=False)
        app.Quit()
        temp_file = temp_new

    return temp_file


def handle_pdf_save(from_page, to_page, printer_name):
    """处理 PDF 打印机的另存为对话框"""
    if not is_pdf_printer(printer_name):
        return

    time.sleep(2)
    print(f"[PDF Save] saving...")

    ts = datetime.now().strftime('%H%M%S')
    pdf_name = f"print_p{from_page}-{to_page}_{ts}.pdf"
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    pdf_path = os.path.join(desktop, pdf_name)

    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except:
            pass

    time.sleep(0.5)
    pyautogui.hotkey('alt', 'n')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)

    ps_cmd = f"Set-Clipboard -Value '{pdf_path}'"
    subprocess.run(['powershell', '-Command', ps_cmd],
                   capture_output=True, timeout=10)
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(4)

    if os.path.exists(pdf_path):
        size = os.path.getsize(pdf_path)
        print(f"[PDF Save] OK: {pdf_name} ({size} bytes)")
    else:
        print(f"[PDF Save] failed")


# ══════════════════════════════════════════════════════════════
# Word 打印
# ══════════════════════════════════════════════════════════════
#
# Word Tab顺序 (Ctrl+P后):
#   Tab0=打印按钮, Tab1=份数, Tab2=打印机, Tab3=打印机属性
#   Tab4=打印范围, Tab5=页数输入框(输入"2-3"格式)
#   Tab6=单双面, Tab7=对照, Tab8=横向纵向
#   Tab9=纸张, Tab10=边距, Tab11=缩放, Tab12=页面设置
#

def print_word(file_path, from_page, to_page, printer_name=None, duplex='double'):
    """Word 指定页打印"""
    import win32com.client

    word = None
    doc = None
    try:
        block_input()

        print(f"[Word] preparing file...")
        temp_file = prepare_file(file_path)

        print(f"[Word] starting...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        doc = word.Documents.Open(temp_file)
        time.sleep(2)

        # 设置打印机 (COM直接设置，Word不需要端口名)
        if printer_name:
            word.ActivePrinter = printer_name
            print(f"[Word] printer: {printer_name}")

        # 激活窗口
        word.Activate()
        time.sleep(0.5)
        hwnd = find_window_by_class('OpusApp')
        if hwnd:
            force_foreground(hwnd)
            time.sleep(0.3)

        # Ctrl+P
        print(f"[Word] Ctrl+P...")
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(3)

        # Tab 5 -> 页数输入框
        for _ in range(5):
            pyautogui.press('tab')
            time.sleep(0.08)
        time.sleep(0.2)

        # 输入页码 (Word用"2-3"格式)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.05)
        pyautogui.write(f"{from_page}-{to_page}", interval=0.05)
        time.sleep(0.2)
        print(f"[Word] pages: {from_page}-{to_page}")

        # 单面/双面 (仅物理打印机)
        if duplex == 'single' and not is_pdf_printer(printer_name):
            pyautogui.press('tab')   # -> Tab6 单双面
            time.sleep(0.2)
            pyautogui.press('down')
            time.sleep(0.2)
            pyautogui.press('up')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(0.2)
            print(f"[Word] single-sided")
            shift_count = 6
        else:
            print(f"[Word] duplex (default)" if not is_pdf_printer(printer_name) else "[Word] PDF printer, skip duplex")
            shift_count = 5

        # Shift+Tab 回到打印按钮
        for _ in range(shift_count):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(0.08)
        time.sleep(0.3)

        # Enter 打印
        print(f"[Word] printing!")
        pyautogui.press('enter')

        # PDF打印机另存为
        handle_pdf_save(from_page, to_page, printer_name)

        time.sleep(3)

        try:
            doc.Close(SaveChanges=False)
        except:
            pass
        try:
            word.Quit()
        except:
            pass
        try:
            os.remove(temp_file)
        except:
            pass

        print(f"[Word] done!")
        return True

    except Exception as e:
        print(f"[Word] error: {e}")
        try:
            if doc: doc.Close(SaveChanges=False)
        except:
            pass
        try:
            if word: word.Quit()
        except:
            pass
        return False
    finally:
        unblock_input()


# ══════════════════════════════════════════════════════════════
# PowerPoint 打印
# ══════════════════════════════════════════════════════════════
#
# PPT Tab顺序 (Ctrl+P后):
#   Tab0=打印按钮, Tab1=份数, Tab2=打印机下拉, Tab3=打印机属性
#   Tab4=打印范围下拉(全部/所选/自定义)
#   [选自定义后] Tab5=页码输入框(输入"2-3"格式)
#   Tab6=单双面, Tab7=打印顺序, Tab8=颜色/灰度
#
# 注意: PPT COM不支持ActivePrinter设置(只读)
#

def print_ppt(file_path, from_page, to_page, printer_name=None, duplex='double'):
    """PPT 指定页打印"""
    import win32com.client

    ppt = None
    pres = None
    try:
        block_input()

        print(f"[PPT] preparing file...")
        temp_file = prepare_file(file_path)

        print(f"[PPT] starting...")
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(temp_file)
        time.sleep(2)

        # 激活窗口
        hwnd = find_window_by_class('PPTFrameClass')
        if hwnd:
            force_foreground(hwnd)
            time.sleep(0.5)

        # Ctrl+P
        print(f"[PPT] Ctrl+P...")
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(3)

        # Tab 4 -> 打印范围下拉框
        for _ in range(4):
            pyautogui.press('tab')
            time.sleep(0.1)

        # 选"自定义范围": Down*2 + Enter
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Tab 1 -> 页码输入框
        pyautogui.press('tab')
        time.sleep(0.2)

        # 输入页码 (PPT用"2-3"格式)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.05)
        pyautogui.write(f"{from_page}-{to_page}", interval=0.05)
        time.sleep(0.2)
        print(f"[PPT] pages: {from_page}-{to_page}")

        # 单双面 (仅物理打印机)
        if duplex == 'single' and not is_pdf_printer(printer_name):
            pyautogui.press('tab')   # -> Tab6 单双面
            time.sleep(0.2)
            pyautogui.press('down')
            time.sleep(0.2)
            pyautogui.press('up')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(0.2)
            print(f"[PPT] single-sided")
            shift_count = 6
        else:
            print(f"[PPT] duplex (default)" if not is_pdf_printer(printer_name) else "[PPT] PDF printer, skip duplex")
            shift_count = 5

        # Shift+Tab 回到打印按钮
        for _ in range(shift_count):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(0.08)
        time.sleep(0.3)

        # Enter 打印
        print(f"[PPT] printing!")
        pyautogui.press('enter')

        # PDF打印机另存为
        handle_pdf_save(from_page, to_page, printer_name)

        time.sleep(3)

        try:
            pres.Close()
        except:
            pass
        try:
            ppt.Quit()
        except:
            pass
        try:
            os.remove(temp_file)
        except:
            pass

        print(f"[PPT] done!")
        return True

    except Exception as e:
        print(f"[PPT] error: {e}")
        try:
            if pres: pres.Close()
        except:
            pass
        try:
            if ppt: ppt.Quit()
        except:
            pass
        return False
    finally:
        unblock_input()


# ══════════════════════════════════════════════════════════════
# Excel 打印
# ══════════════════════════════════════════════════════════════
#
# Excel Tab顺序 (Ctrl+P后):
#   Tab0=打印按钮, Tab1=份数, Tab2=打印机, Tab3=打印机属性
#   Tab4=打印范围(不动!默认"打印活动工作表"即可)
#   Tab5=起始页输入框("从第X页")
#   Tab6=结束页输入框("至第Y页")
#   Tab7=单双面
#
# 关键: Excel的页码是两个独立输入框，不是"1-2"格式！
#   Tab5输入起始页 -> Tab一下 -> Tab6输入结束页
#
# Excel设打印机需要端口名: "打印机名 在 NeXX:"
#

def print_excel(file_path, from_page, to_page, printer_name=None, duplex='double'):
    """Excel 指定页打印"""
    import win32com.client

    excel = None
    wb = None
    try:
        block_input()

        print(f"[Excel] preparing file...")
        temp_file = prepare_file(file_path)

        print(f"[Excel] starting...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(temp_file)
        time.sleep(2)

        # 设置打印机 (需要带端口名)
        if printer_name:
            set_excel_printer(excel, printer_name)

        # 激活窗口
        hwnd = find_window_by_class('XLMAIN')
        if hwnd:
            force_foreground(hwnd)
            time.sleep(0.5)

        # Ctrl+P
        print(f"[Excel] Ctrl+P...")
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(3)

        # Tab 5 -> 起始页输入框("从")
        print(f"[Excel] Tab 5 -> FROM page...")
        for _ in range(5):
            pyautogui.press('tab')
            time.sleep(0.12)
        time.sleep(0.3)

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.05)
        pyautogui.write(str(from_page), interval=0.08)
        time.sleep(0.3)

        # Tab -> 结束页输入框("至")
        print(f"[Excel] Tab -> TO page...")
        pyautogui.press('tab')
        time.sleep(0.3)

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.05)
        pyautogui.write(str(to_page), interval=0.08)
        time.sleep(0.3)

        print(f"[Excel] pages: {from_page} to {to_page}")

        # 单双面 (仅物理打印机)
        if duplex == 'single' and not is_pdf_printer(printer_name):
            pyautogui.press('tab')   # -> Tab7 单双面
            time.sleep(0.2)
            pyautogui.press('down')
            time.sleep(0.2)
            pyautogui.press('up')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(0.2)
            print(f"[Excel] single-sided")
            shift_count = 7
        else:
            print(f"[Excel] duplex (default)" if not is_pdf_printer(printer_name) else "[Excel] PDF printer, skip duplex")
            shift_count = 6

        # Shift+Tab 回到打印按钮
        for _ in range(shift_count):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(0.1)
        time.sleep(0.3)

        # Enter 打印
        print(f"[Excel] printing!")
        pyautogui.press('enter')

        # PDF打印机另存为
        handle_pdf_save(from_page, to_page, printer_name)

        time.sleep(3)

        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        try:
            os.remove(temp_file)
        except:
            pass

        print(f"[Excel] done!")
        return True

    except Exception as e:
        print(f"[Excel] error: {e}")
        try:
            if wb: wb.Close(SaveChanges=False)
        except:
            pass
        try:
            if excel: excel.Quit()
        except:
            pass
        return False
    finally:
        unblock_input()


# ══════════════════════════════════════════════════════════════
# PDF 打印 (Acrobat) - 待实现
# ══════════════════════════════════════════════════════════════

def print_pdf(file_path, from_page, to_page, printer_name=None, duplex='double'):
    """PDF 指定页打印 - 暂未实现"""
    print(f"[PDF] not implemented yet")
    return False


# ══════════════════════════════════════════════════════════════
# 统一入口
# ══════════════════════════════════════════════════════════════

def get_file_type(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.doc', '.docx'):
        return 'word'
    elif ext in ('.ppt', '.pptx'):
        return 'ppt'
    elif ext in ('.xls', '.xlsx'):
        return 'excel'
    elif ext == '.pdf':
        return 'pdf'
    else:
        return None


def print_pages(file_path, from_page, to_page, printer_name=None, duplex='double'):
    """统一入口"""
    anti_sleep()
    kill_office()

    file_type = get_file_type(file_path)
    if not file_type:
        print(f"Error: unsupported format {os.path.splitext(file_path)[1]}")
        print(f"Supported: .doc, .docx, .ppt, .pptx, .xls, .xlsx, .pdf")
        return False

    if not os.path.exists(file_path):
        print(f"Error: file not found: {file_path}")
        return False

    if not printer_name:
        printer_name = PRINTER_DEFAULT

    pdf_flag = "(PDF)" if is_pdf_printer(printer_name) else ""
    print(f"{'='*50}")
    print(f"File: {os.path.basename(file_path)}")
    print(f"Type: {file_type.upper()}")
    print(f"Pages: {from_page}-{to_page}")
    print(f"Printer: {printer_name} {pdf_flag}")
    print(f"Duplex: {'single' if duplex == 'single' and not is_pdf_printer(printer_name) else 'double' if not is_pdf_printer(printer_name) else 'N/A(PDF)'}")
    print(f"{'='*50}")

    dispatch = {
        'word': print_word,
        'ppt': print_ppt,
        'excel': print_excel,
        'pdf': print_pdf,
    }

    return dispatch[file_type](file_path, from_page, to_page, printer_name, duplex)


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Office Print T1 v2 - COM+Tab hybrid")
        print()
        print("Usage: python print_pages.py <file> <from> <to> [printer] [single/double]")
        print()
        print("Supported: .doc, .docx, .ppt, .pptx, .xls, .xlsx, .pdf")
        print()
        print("Examples:")
        print('  python print_pages.py doc.docx 2 3')
        print('  python print_pages.py doc.docx 2 3 "Ricoh SP 330"')
        print('  python print_pages.py doc.docx 2 3 "Ricoh SP 330" single')
        print('  python print_pages.py sheet.xlsx 1 2 "Microsoft Print to PDF"')
        sys.exit(1)

    file_path = sys.argv[1]
    from_page = int(sys.argv[2])
    to_page = int(sys.argv[3])
    printer = sys.argv[4] if len(sys.argv) > 4 else None
    duplex = sys.argv[5] if len(sys.argv) > 5 else 'double'

    ok = print_pages(file_path, from_page, to_page, printer, duplex)
    sys.exit(0 if ok else 1)
