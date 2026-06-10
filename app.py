from tkinterdnd2 import TkinterDnD
from tkinter import ttk, messagebox

# 导入所有功能模块（每个模块自带UI）
import video_excel
import image_excel
import zip_excel
import mlumourl_excel
import content_excel
import extend_excel

class MainApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("全能Excel转换工具 v3.1_2026-06-10")
        self.geometry("790x860")
        self.resizable(False, False) # (宽是否可拖动缩放, 高是否可拖动缩放)

        # 全局样式
        style = ttk.Style()
        style.configure('.', font=('微软雅黑', 10))
        style.configure('TNotebook.Tab', padding=[12, 6], font=('微软雅黑', 10, 'bold'))

        # 标签容器
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # 按顺序添加标签：视频 → 图片 → ZIP → URL → 文案转Excel → 拓展文件转Excel
        video_excel.VideoExcelTab(notebook)
        image_excel.ImageExcelTab(notebook)
        zip_excel.ZipExcelTab(notebook)
        mlumourl_excel.UrlExcelTab(notebook)
        content_excel.ContentExcelTab(notebook)
        extend_excel.ExtendExcelTab(notebook)

if __name__ == "__main__":
    try:
        app = MainApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("启动失败", f"程序出错：{str(e)}")