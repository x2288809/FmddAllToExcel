import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES
from datetime import datetime
from openpyxl import Workbook
from docx import Document

# ===================== 核心解析函数 =====================
def parse_content(text_content):
    lines = [line.strip() for line in text_content.splitlines() if line.strip() != '']
    items = []
    index = 0
    total_lines = len(lines)
    pattern = re.compile(r'^\d+\.$')

    while index < total_lines:
        current_line = lines[index]
        if pattern.match(current_line):
            seq_num = current_line.replace('.', '')
            if index + 1 < total_lines:
                title = lines[index + 1]
                index += 2
            else:
                title = ''
                index += 1
                continue

            content = []
            while index < total_lines and not pattern.match(lines[index]):
                content.append(lines[index])
                index += 1

            copy_text = '\n'.join(content)
            items.append({
                "序号": seq_num,
                "标题": title,
                "文案": copy_text
            })
        else:
            index += 1
    return items

def read_txt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        with open(file_path, 'r', encoding='gbk') as f:
            return f.read()

def read_docx(file_path):
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs if para.text.strip()])

def create_excel(parsed_data, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "内容表格"

    headers = [
        "序号", "标题", "文案", "话题", "大字报",
        "图片1", "图片2", "图片3", "图片4", "图片5",
        "视频", "使用说明", "指定用户"
    ]
    ws.append(headers)

    for item in parsed_data:
        row = [
            item["序号"], item["标题"], item["文案"],
            "", "", "", "", "", "", "", "", "", ""
        ]
        ws.append(row)

    wb.save(output_path)

# ===================== 标签页类（和你现有格式完全一致）=====================
class ContentExcelTab:
    def __init__(self, notebook):
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text="文案转Excel")

        self.file_path = tk.StringVar()

        # 标题
        ttk.Label(self.tab, text="📝 TXT / Word 转 Excel 工具",
                 font=("微软雅黑", 14, "bold")).pack(pady=15)

        # 文件路径行
        frame_path = ttk.Frame(self.tab)
        frame_path.pack(fill="x", padx=30, pady=5)
        ttk.Label(frame_path, text="文件路径：").pack(side="left", padx=5)
        tk.Entry(frame_path, textvariable=self.file_path, font=("微软雅黑", 10))\
            .pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(frame_path, text="选择文件", command=self.select_file).pack(side="left", padx=5)

        # 拖拽区域
        drop_frame = ttk.LabelFrame(self.tab, text="👇 拖拽文件到这里", padding=10)
        drop_frame.pack(fill="both", expand=True, padx=30, pady=8)
        self.drop_label = ttk.Label(drop_frame, text="拖入 .txt / .docx 文件",
                                   font=("微软雅黑", 12), foreground="#409eff")
        self.drop_label.pack(expand=True, fill="both")

        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind("<<Drop>>", self.on_drop)

        # 开始按钮
        self.run_btn = ttk.Button(self.tab, text="✅ 开始转换",
                                  command=self.start_convert, width=20)
        self.run_btn.pack(pady=10, ipadx=10, ipady=5)

        # 日志
        log_frame = ttk.LabelFrame(self.tab, text="处理日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=30, pady=5)
        self.log_text = tk.Text(log_frame, height=4, font=("微软雅黑", 9))
        self.log_text.pack(fill="both", expand=True)

        self.log("工具已启动，支持 TXT / DOCX 格式解析")
        self.log("要求文本格式为序号开头，分割文案，如下：\n 1. 标题\n文案内容（可多行）\n\n2. 标题\n文案内容...\n...\n")

    def log(self, msg):
        time_str = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{time_str}] {msg}\n")
        self.log_text.see("end")

    def select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("支持的文件", "*.txt;*.docx"), ("TXT 文件", "*.txt"), ("Word 文件", "*.docx")]
        )
        if path:
            self.file_path.set(path)
            self.log(f"已选择：{os.path.basename(path)}")

    def on_drop(self, e):
        path = e.data.strip("{}").replace("\\", "/")
        if path.lower().endswith((".txt", ".docx")):
            self.file_path.set(path)
            self.log(f"已拖拽：{os.path.basename(path)}")

    def start_convert(self):
        path = self.file_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("错误", "请选择有效文件")
            return

        self.run_btn.config(state=tk.DISABLED)
        self.log("开始处理文件...")

        try:
            if path.endswith(".txt"):
                content = read_txt(path)
            elif path.endswith(".docx"):
                content = read_docx(path)
            else:
                messagebox.showerror("错误", "不支持此文件格式")
                return

            data = parse_content(content)
            if not data:
                self.log("未解析到有效内容，请检查文件格式")
                messagebox.showwarning("提示", "未解析到内容")
                return

            output_dir = os.path.dirname(path)
            output_name = f"文案表格_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            output_path = os.path.join(output_dir, output_name)

            create_excel(data, output_path)
            self.log(f"✅ 转换完成：{output_name}")
            messagebox.showinfo("完成", f"Excel 已生成！\n{output_path}")

        except Exception as e:
            self.log(f"❌ 出错：{str(e)}")
            messagebox.showerror("转换失败", str(e))

        finally:
            self.run_btn.config(state=tk.NORMAL)