import pdfplumber
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import sys

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
        """选择PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF发票文件",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if files:
            self.selected_files.extend(files)
            self.update_file_list()
            self.rename_btn.config(state=tk.NORMAL)
            if len(files) == 1:  # 只有一个文件时启用预览按钮
                self.preview_btn.config(state=tk.NORMAL)
            else:
                self.preview_btn.config(state=tk.DISABLED)
    
    def select_folder(self):
        """选择文件夹进行批量处理"""
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        
        if folder_path:
            # 获取文件夹中的所有PDF文件
            pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
                        if f.lower().endswith('.pdf')]
            self.selected_files.extend(pdf_files)
            self.update_file_list()
            if pdf_files:
                self.rename_btn.config(state=tk.NORMAL)
                self.preview_btn.config(state=tk.DISABLED)  # 文件夹模式禁用预览
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.selected_files:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))
    
    def clear_list(self):
        """清空文件列表"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.rename_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.result_text.delete(1.0, tk.END)
        self.progress['value'] = 0
    
    def preview_pdf(self):
        """预览选中PDF的内容"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择一个PDF文件")
            return
        
        # 只处理第一个选中的文件
        file_path = self.selected_files[0]
        try:
            with pdfplumber.open(file_path) as pdf:
                # 读取第一页内容
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                
                # 弹出预览窗口
                self.show_preview_window(text)
                
        except Exception as e:
            messagebox.showerror("错误", f"预览文件时出错: {str(e)}")
    
    def show_preview_window(self, text):
        """显示PDF内容预览窗口"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("PDF内容预览")
        preview_window.geometry("600x400")
        
        # 创建文本框显示内容
        text_widget = tk.Text(preview_window, wrap=tk.WORD)
        scrollbar = tk.Scrollbar(preview_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 插入文本内容
        text_widget.insert(tk.END, text)
        text_widget.config(state=tk.DISABLED)  # 设置为只读
    
    def start_rename(self):
        """开始重命名，使用线程避免界面冻结"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要处理的PDF文件")
            return
        
        # 检查姓名是否为空
        if not self.name_var.get().strip():
            messagebox.showwarning("警告", "请输入姓名")
            return
        
        # 禁用按钮防止重复点击
        self.rename_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.select_folder_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        
        # 在新线程中执行重命名
        rename_thread = threading.Thread(target=self.rename_files)
        rename_thread.daemon = True
        rename_thread.start()
    
    def rename_files(self):
        """重命名文件的主逻辑"""
        total_files = len(self.selected_files)
        processed = 0
        
        for i, file_path in enumerate(self.selected_files):
            try:
                # 调用原有的重命名函数
                result = self.rename_invoice_pdf(file_path)
                if result:
                    self.result_text.insert(tk.END, f"✓ 重命名成功: {os.path.basename(file_path)} -> {result}\n")
                else:
                    self.result_text.insert(tk.END, f"✗ 重命名失败: {os.path.basename(file_path)}\n")
            except Exception as e:
                self.result_text.insert(tk.END, f"✗ 处理失败 {os.path.basename(file_path)}: {str(e)}\n")
            
            # 更新进度
            processed += 1
            progress_percent = (processed / total_files) * 100
            self.progress['value'] = progress_percent
            
            # 滚动到底部显示最新结果
            self.result_text.see(tk.END)
            
            # 更新GUI
            self.root.update_idletasks()
        
        # 完成后启用按钮
        self.rename_btn.config(state=tk.NORMAL)
        self.select_btn.config(state=tk.NORMAL)
        self.select_folder_btn.config(state=tk.NORMAL)
        if len(self.selected_files) == 1:
            self.preview_btn.config(state=tk.NORMAL)
        
        messagebox.showinfo("完成", "所有文件处理完成！")
    
    def rename_invoice_pdf(self, file_path):
        """PDF重命名的核心函数，改进金额提取逻辑，增加Windows兼容性"""
        directory = os.path.dirname(file_path)
        
        try:
            # 先读取PDF内容
            extracted_data = self.extract_pdf_data(file_path)
            if not extracted_data:
                print(f"无法从文件 {file_path} 中提取文字。")
                return None

            month, amount = extracted_data

            # 使用自定义姓名构造新文件名
            custom_name = self.name_var.get().strip()
            if not custom_name:
                custom_name = "未知"
            
            new_name = f"{custom_name}{month}月通讯费发票 {amount}元.pdf"
            new_path = os.path.join(directory, new_name)

            # 检查目标文件是否已存在
            if os.path.exists(new_path):
                print(f"文件已存在，跳过: {new_name}")
                return None

            # 重命名文件 - 在Windows上使用重试机制
            if self.safe_rename_file(file_path, new_path):
                return new_name
            else:
                return None

        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {e}")
            return None

    def extract_pdf_data(self, file_path):
        """从PDF中提取月份和金额信息"""
        try:
            with pdfplumber.open(file_path) as pdf:
                # 仅读取第一页内容（通常关键信息都在第一页）
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                
                if not text:
                    return None

                # 1. 提取账期月份 (匹配 账期:202506 或 账期:202501)
                month_match = re.search(r"账期[:：]\d{4}(\d{2})", text)
                if month_match:
                    # 将 06 转换为 6，01 转换为 1
                    month = str(int(month_match.group(1)))
                else:
                    month = "未知"

                # 2. 提取金额 - 使用多种可能的模式
                amount = "未知"
                
                # 模式1: (小写)¥184.00
                amount_match = re.search(r"\(小写\)[¥]?\s*(\d+\.?\d*)", text)
                if amount_match:
                    amount = amount_match.group(1)
                else:
                    # 模式2: 价税合计(小写)¥184.00
                    amount_match = re.search(r"价税合计[（(]小写[）)]\s*[¥]?\s*(\d+\.?\d*)", text)
                    if amount_match:
                        amount = amount_match.group(1)
                    else:
                        # 模式3: ¥184.00
                        amount_match = re.search(r"[¥]\s*(\d+\.?\d*)", text)
                        if amount_match:
                            amount = amount_match.group(1)
                        else:
                            # 模式4: 简单的数字金额匹配（带小数点）
                            amount_match = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+(?:\.\d{2}))", text)
                            if amount_match:
                                amount = amount_match.group(1).replace(',', '')  # 去除千分位逗号

                return month, amount
        except Exception as e:
            print(f"提取PDF数据时出错: {e}")
            return None

    def safe_rename_file(self, old_path, new_path, max_retries=5, delay=0.5):
        """
        安全重命名文件，对于Windows平台增加重试机制
        """
        for attempt in range(max_retries):
            try:
                # 确保文件未被占用
                if sys.platform.startswith('win'):  # Windows平台
                    # 短暂延迟，确保文件句柄被释放
                    time.sleep(delay)
                
                # 执行重命名
                os.rename(old_path, new_path)
                return True  # 成功重命名
                
            except PermissionError as e:
                print(f"权限错误 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(delay * (attempt + 1))  # 递增延迟
                else:
                    print(f"重命名失败，已达到最大重试次数: {old_path}")
                    return False
                    
            except FileNotFoundError as e:
                print(f"文件未找到: {e}")
                return False
                
            except OSError as e:
                print(f"操作系统错误: {e}")
                if attempt < max_retries - 1:
                    time.sleep(delay * (attempt + 1))
                else:
                    print(f"重命名失败，已达到最大重试次数: {old_path}")
                    return False
                    
            except Exception as e:
                print(f"重命名时发生未知错误: {e}")
                return False
        
        return False

def main():
    root = tk.Tk()
    app = PDFRenamerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()