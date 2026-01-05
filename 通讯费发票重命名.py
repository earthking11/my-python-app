import pdfplumber
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import traceback  # 新增：用于捕获详细报错内容

class PDFRenamerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF发票重命名工具")
        self.root.geometry("800x650")
        
        # 存储选中的文件列表
        self.selected_files = []
        
        # 默认姓名
        self.name_var = tk.StringVar(value="张三")
        
        # 创建界面组件
        self.create_widgets()
        
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="PDF发票重命名工具", font=("Arial", 16))
        title_label.pack(pady=10)
        
        # 姓名输入框架
        name_frame = tk.Frame(self.root)
        name_frame.pack(pady=5, fill=tk.X, padx=20)
        
        tk.Label(name_frame, text="自定义姓名:", font=("Arial", 12)).pack(anchor=tk.W)
        
        # 姓名输入框
        name_entry = tk.Entry(name_frame, textvariable=self.name_var, width=30, font=("Arial", 12))
        name_entry.pack(pady=5)
        
        # 文件选择框架
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, fill=tk.X, padx=20)
        
        # 选择文件按钮
        self.select_btn = tk.Button(file_frame, text="选择PDF文件", command=self.select_files, 
                                   bg="lightblue", font=("Arial", 12))
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # 选择文件夹按钮
        self.select_folder_btn = tk.Button(file_frame, text="选择文件夹(批量处理)", 
                                          command=self.select_folder, bg="lightgreen", font=("Arial", 12))
        self.select_folder_btn.pack(side=tk.LEFT, padx=5)
        
        # 预览按钮
        self.preview_btn = tk.Button(file_frame, text="预览PDF内容", command=self.preview_pdf,
                                    bg="lightyellow", font=("Arial", 12), state=tk.DISABLED)
        self.preview_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表框
        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        # 文件列表标题
        tk.Label(list_frame, text="待处理文件列表:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        
        # 创建带滚动条的列表框
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 按钮框架
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        # 开始重命名按钮
        self.rename_btn = tk.Button(button_frame, text="开始重命名", command=self.start_rename,
                                   bg="lightcoral", font=("Arial", 12), state=tk.DISABLED)
        self.rename_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空列表按钮
        self.clear_btn = tk.Button(button_frame, text="清空列表", command=self.clear_list,
                                  bg="lightyellow", font=("Arial", 12))
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=10)
        
        # 结果显示区域
        result_frame = tk.Frame(self.root)
        result_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        tk.Label(result_frame, text="处理结果:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, height=12)
        result_scrollbar = tk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择PDF发票文件",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if files:
            self.selected_files.extend(files)
            self.update_file_list()
            self.rename_btn.config(state=tk.NORMAL)
            self.preview_btn.config(state=tk.NORMAL if len(self.selected_files) >= 1 else tk.DISABLED)
    
    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder_path:
            pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
                        if f.lower().endswith('.pdf')]
            self.selected_files.extend(pdf_files)
            self.update_file_list()
            if pdf_files:
                self.rename_btn.config(state=tk.NORMAL)
    
    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for file_path in self.selected_files:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))
    
    def clear_list(self):
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.rename_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.result_text.delete(1.0, tk.END)
        self.progress['value'] = 0
    
    def preview_pdf(self):
        if not self.selected_files:
            return
        file_path = self.selected_files[0]
        try:
            with pdfplumber.open(file_path) as pdf:
                text = pdf.pages[0].extract_text()
                self.show_preview_window(text)
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")
    
    def show_preview_window(self, text):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("PDF内容预览")
        text_widget = tk.Text(preview_window, wrap=tk.WORD)
        text_widget.insert(tk.END, text if text else "未能识别到文字")
        text_widget.pack(fill=tk.BOTH, expand=True)

    def start_rename(self):
        if not self.name_var.get().strip():
            messagebox.showwarning("警告", "请输入姓名")
            return
        
        self.rename_btn.config(state=tk.DISABLED)
        self.result_text.delete(1.0, tk.END)
        
        rename_thread = threading.Thread(target=self.rename_files_loop)
        rename_thread.daemon = True
        rename_thread.start()
    
    def rename_files_loop(self):
        total_files = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            try:
                # 核心改动：调用修复后的重命名逻辑
                result = self.execute_rename_logic(file_path)
                if result:
                    self.result_text.insert(tk.END, f"✓ 成功: {os.path.basename(file_path)} -> {result}\n")
                else:
                    self.result_text.insert(tk.END, f"⚠ 跳过: {os.path.basename(file_path)} (可能已存在或格式不符)\n")
            except Exception as e:
                # 改进：捕获详细报错信息
                err_msg = str(e)
                full_trace = traceback.format_exc()
                self.result_text.insert(tk.END, f"✗ 失败 {os.path.basename(file_path)}:\n原因: {err_msg}\n")
                print(full_trace) # 控制台记录详细堆栈
            
            self.progress['value'] = ((i + 1) / total_files) * 100
            self.result_text.see(tk.END)
            self.root.update_idletasks()
        
        self.rename_btn.config(state=tk.NORMAL)
        messagebox.showinfo("完成", "处理结束！")

    def execute_rename_logic(self, file_path):
        """修复了文件占用逻辑的核心函数"""
        directory = os.path.dirname(file_path)
        new_name = None
        
        # 1. 先读取并关闭文件
        with pdfplumber.open(file_path) as pdf:
            text = pdf.pages[0].extract_text()
            if not text:
                return None

            # 提取账期
            month_match = re.search(r"账期[:：]\d{4}(\d{2})", text)
            month = str(int(month_match.group(1))) if month_match else "未知"

            # 提取金额
            amount = "未知"
            # 尝试多种金额匹配模式
            patterns = [
                r"\(小写\)[¥]?\s*(\d+\.?\d*)",
                r"价税合计[（(]小写[）)]\s*[¥]?\s*(\d+\.?\d*)",
                r"[¥]\s*(\d+\.?\d*)",
                r"(\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+(?:\.\d{2}))"
            ]
            for p in patterns:
                m = re.search(p, text)
                if m:
                    amount = m.group(1).replace(',', '')
                    break

            custom_name = self.name_var.get().strip() or "未知"
            new_name = f"{custom_name}{month}月通讯费发票 {amount}元.pdf"

        # 2. 退出 with 后，文件句柄已释放，