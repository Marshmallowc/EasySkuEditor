import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
from pathlib import Path
import os
import re
import shutil

class SKUManager(ctk.CTk):
    
    def __init__(self):
        super().__init__()  

        # 设置主题和外观
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # 添加主题切换功能
        self.theme_button = ctk.CTkButton(
            self,
            text="切换主题",
            command=self.toggle_theme,
            width=100
        )
        self.theme_button.pack(anchor="ne", padx=10, pady=5)

        # 配置主窗口
        self.title("SKU 文件管理器")
        self.geometry("1200x700")  # 增加窗口大小
        
        # 设置最小窗口大小
        self.minsize(800, 600)
        
        # 创建主框架
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 初始化数据
        self.current_file = None
        self.df = None
        
        # 定义支持的分隔符
        self.separator_pattern = r'[,;，；|\t/／、]+'
        
        # 初始化界面
        self.create_welcome_screen()

        # 添加快捷键支持
        self.bind("<Control-s>", lambda e: self.save_file())
        self.bind("<Control-o>", lambda e: self.show_open_file_screen())
        self.bind("<Escape>", lambda e: self.clear_frame() or self.create_welcome_screen())
        
        # 添加状态栏
        self.status_bar = ctk.CTkLabel(
            self,
            text="就绪",
            anchor="w"
        )
        self.status_bar.pack(side="bottom", fill="x", padx=5, pady=2)

    def clear_frame(self):
        """清空主框架"""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def validate_input(self, input_text, input_type="text"):
        """增强的输入数据验证"""
        if not input_text.strip():
            messagebox.showerror("错误", "输入不能为空")
            return False
        
        # 检查特殊字符
        if re.search(r'[<>{}\\]', input_text):
            messagebox.showerror("错误", "输入包含非法字符")
            return False
        
        if input_type == "number":
            values = re.split(self.separator_pattern, input_text)
            for value in values:
                value = value.strip()
                if not value:
                    continue
                # 增强数字验证
                if not re.match(r'^-?\d*\.?\d+$', value):
                    messagebox.showerror("错误", f"'{value}' 不是有效的数字")
                    return False
                # 检查数值范围
                try:
                    num = float(value)
                    if abs(num) > 1e10:
                        messagebox.showerror("错误", f"数值 {value} 超出合理范围")
                        return False
                except ValueError:
                    messagebox.showerror("错误", f"无法转换数值: {value}")
                    return False
        return True

    def convert_data(self, raw_data, data_type, selected_attr):
        """统一的数据转换处理"""
        try:
            if data_type == "number":
                # 添加数值范围检查
                data = []
                for x in raw_data:
                    try:
                        num = float(x) if '.' in x else int(x)
                        if abs(num) > 1e10:  # 设置合理的数值范围
                            raise ValueError(f"数值 {num} 超出合理范围")
                        data.append(num)
                    except ValueError as e:
                        messagebox.showerror("错误", f"数值转换错误: {str(e)}")
                        return None
                
                if selected_attr not in self.df.columns or self.df[selected_attr].empty:
                    self.df[selected_attr] = pd.Series(dtype='float64')
            else:
                # 添加文本长度检查
                data = []
                for text in raw_data:
                    if len(text) > 1000:  # 设置最大文本长度
                        messagebox.showerror("错误", f"文本长度超过限制: {text[:50]}...")
                        return None
                    data.append(text)
                
                if selected_attr not in self.df.columns:
                    self.df[selected_attr] = pd.Series(dtype='object')
            return data
        except Exception as e:
            messagebox.showerror("错误", f"数据转换错误: {str(e)}")
            return None

    def save_file(self):
        """统一的文件保存处理"""
        try:
            # 添加文件备份功能
            backup_path = f"{self.current_file}.bak"
            if os.path.exists(self.current_file):
                shutil.copy2(self.current_file, backup_path)
            
            if self.current_file.endswith('.xlsx'):
                with pd.ExcelWriter(self.current_file, mode='w', engine='openpyxl') as writer:
                    self.df.to_excel(writer, index=False)
            else:
                self.df.to_csv(self.current_file, index=False, encoding='utf-8-sig')  # 添加 BOM 以支持中文
            
            # 保存成功后删除备份
            if os.path.exists(backup_path):
                os.remove(backup_path)
            return True
        except Exception as e:
            # 发生错误时恢复备份
            if os.path.exists(backup_path):
                shutil.copy2(backup_path, self.current_file)
                os.remove(backup_path)
            messagebox.showerror("错误", f"保存文件时发生错误：{str(e)}")
            return False

    def update_preview(self, preview_text):
        """更新预览区域"""
        preview_text.delete("1.0", tk.END)
        if self.df is not None and not self.df.empty:
            # 使用 to_string 方法来获取更好的格式化输出
            preview_text.insert("1.0", self.df.to_string())
        else:
            preview_text.insert("1.0", "暂无数据")

    def create_welcome_screen(self):
        """创建更现代的欢迎界面"""
        # 欢迎标题
        title = ctk.CTkLabel(
            self.main_frame,
            text="SKU 文件管理器",
            font=("Helvetica", 28, "bold")
        )
        title.pack(pady=40)

        # 功能描述
        desc = ctk.CTkLabel(
            self.main_frame,
            text="简单高效的SKU数据管理工具",
            font=("Helvetica", 14),
            text_color="gray"
        )
        desc.pack(pady=10)

        # 按钮容器
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.pack(pady=30)

        # 创建文件按钮
        create_btn = ctk.CTkButton(
            button_frame,
            text="创建新文件",
            command=self.show_create_file_screen,
            width=250,
            height=50,
            font=("Helvetica", 14)
        )
        create_btn.pack(pady=10)

        # 打开文件按钮
        open_btn = ctk.CTkButton(
            button_frame,
            text="打开现有文件",
            command=self.show_open_file_screen,
            width=250,
            height=50,
            font=("Helvetica", 14)
        )
        open_btn.pack(pady=10)

    def create_edit_interface(self):
        """创建编辑界面的通用元素"""
        # 返回按钮
        back_btn = ctk.CTkButton(
            self.main_frame,
            text="返回",
            command=lambda: [self.clear_frame(), self.create_welcome_screen()],
            width=100
        )
        back_btn.pack(anchor="nw", padx=10, pady=10)

        # 文件信息
        file_info = ctk.CTkLabel(
            self.main_frame,
            text=f"当前文件: {os.path.basename(self.current_file)}"
        )
        file_info.pack(pady=10)

        # 预览区域
        preview_frame = ctk.CTkFrame(self.main_frame)
        preview_frame.pack(fill="x", padx=20, pady=10)
        
        # 添加表格标题
        title_label = ctk.CTkLabel(
            preview_frame,
            text="数据预览",
            font=("Helvetica", 14, "bold")
        )
        title_label.pack(pady=5)
        
        # 优化预览框显示
        preview_text = ctk.CTkTextbox(
            preview_frame,
            height=250,  # 增加高度以显示更多内容
            wrap="none",
            font=("Courier", 12)  # 使用等宽字体使表格对齐
        )
        preview_text.pack(fill="both", expand=True, padx=5, pady=5)  # 允许水平和垂直扩展

        # 修改预览内容显示方式
        def update_preview():
            preview_text.delete("1.0", tk.END)
            if self.df is not None and not self.df.empty:
                # 使用 to_string 方法来获取更好的格式化输出
                preview_text.insert("1.0", self.df.to_string())
            else:
                preview_text.insert("1.0", "暂无数据")

        update_preview()
        return preview_text

    def show_create_file_screen(self):
        """显示创建文件界面"""
        self.clear_frame()
        
        # 返回按钮
        back_btn = ctk.CTkButton(
            self.main_frame,
            text="返回",
            command=lambda: [self.clear_frame(), self.create_welcome_screen()],
            width=100
        )
        back_btn.pack(anchor="nw", padx=10, pady=10)

        # 文件配置框架
        config_frame = ctk.CTkFrame(self.main_frame)
        config_frame.pack(fill="x", padx=20, pady=10)

        # 文件名输入
        ctk.CTkLabel(config_frame, text="文件名:").pack(side="left", padx=5)
        filename_entry = ctk.CTkEntry(config_frame, width=200)
        filename_entry.pack(side="left", padx=5)

        # 文件类型选择
        file_type = tk.StringVar(value=".xlsx")
        xlsx_radio = ctk.CTkRadioButton(
            config_frame,
            text="Excel (.xlsx)",
            variable=file_type,
            value=".xlsx"
        )
        csv_radio = ctk.CTkRadioButton(
            config_frame,
            text="CSV (.csv)",
            variable=file_type,
            value=".csv"
        )
        xlsx_radio.pack(side="left", padx=20)
        csv_radio.pack(side="left", padx=20)

        # 属性配置
        attr_frame = ctk.CTkFrame(self.main_frame)
        attr_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(attr_frame, text="属性 (用逗号、分号等分隔):").pack(side="left", padx=5)
        attr_entry = ctk.CTkEntry(attr_frame, width=400)
        attr_entry.pack(side="left", padx=5)

        def create_file():
            """创建文件处理函数"""
            filename = filename_entry.get()
            if not filename:
                messagebox.showerror("错误", "请输入文件名")
                return

            attributes = [
                attr.strip()
                for attr in re.split(self.separator_pattern, attr_entry.get())
                if attr.strip()
            ]

            if not attributes:
                messagebox.showerror("错误", "请输入至少一个属性")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=file_type.get(),
                initialfile=filename,
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("CSV files", "*.csv")
                ]
            )

            if file_path:
                self.df = pd.DataFrame(columns=attributes)
                self.current_file = file_path
                if not self.save_file():
                    return
                self.show_edit_screen()

        # 创建按钮
        create_btn = ctk.CTkButton(
            self.main_frame,
            text="创建文件",
            command=create_file,
            width=200
        )
        create_btn.pack(pady=20)

    def show_open_file_screen(self):
        """显示打开文件界面"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv")
            ]
        )

        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    self.df = pd.read_excel(file_path)
                else:
                    self.df = pd.read_csv(file_path)
                
                self.current_file = file_path
                self.show_edit_screen()
            except Exception as e:
                messagebox.showerror("错误", f"打开文件时发生错误：{str(e)}")

    def show_edit_screen(self):
        """显示编辑界面"""
        self.clear_frame()
        
        # 创建左右分栏布局
        content_frame = ctk.CTkFrame(self.main_frame)
        content_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 左侧：预览和搜索区域
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 文件信息和预览
        file_info = ctk.CTkLabel(
            left_frame,
            text=f"当前文件: {os.path.basename(self.current_file)}",
            font=("Helvetica", 12)
        )
        file_info.pack(pady=5)
        
        preview_text = ctk.CTkTextbox(
            left_frame,
            height=400,
            wrap="none",
            font=("Courier", 12)
        )
        preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 添加这行来立即显示数据
        self.update_preview(preview_text)  # 在创建预览框后立即更新内容
        
        # 右侧：操作区域
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side="right", fill="y", padx=5, pady=5)
        
        # 返回按钮
        back_btn = ctk.CTkButton(
            right_frame,
            text="返回主菜单",
            command=lambda: [self.clear_frame(), self.create_welcome_screen()],
            width=120
        )
        back_btn.pack(pady=5)
        
        # 属性选择和数据输入区域
        input_section = ctk.CTkFrame(right_frame)
        input_section.pack(fill="x", pady=10)
        
        ctk.CTkLabel(input_section, text="选择属性").pack(pady=2)
        attr_dropdown = ctk.CTkOptionMenu(
            input_section,
            values=list(self.df.columns),
            width=200
        )
        attr_dropdown.pack(pady=5)
        
        # 数据类型选择
        type_var = tk.StringVar(value="text")
        type_frame = ctk.CTkFrame(input_section)
        type_frame.pack(pady=5)
        ctk.CTkRadioButton(
            type_frame,
            text="文本",
            variable=type_var,
            value="text"
        ).pack(side="left", padx=10)
        ctk.CTkRadioButton(
            type_frame,
            text="数字",
            variable=type_var,
            value="number"
        ).pack(side="left", padx=10)
        
        # 数据输入框
        ctk.CTkLabel(input_section, text="输入数据").pack(pady=2)
        data_entry = ctk.CTkEntry(
            input_section,
            placeholder_text="用逗号、分号等分隔多个数据",
            width=200,
            height=30
        )
        data_entry.pack(pady=5)
        
        # 在按钮区域之前，定义 update_data 函数
        def update_data():
            """更新数据处理函数"""
            try:
                selected_attr = attr_dropdown.get()
                if not selected_attr:
                    messagebox.showerror("错误", "请选择一个属性")
                    return

                input_text = data_entry.get()
                if not self.validate_input(input_text, type_var.get()):
                    return

                raw_data = [
                    value.strip()
                    for value in re.split(self.separator_pattern, input_text)
                    if value.strip()
                ]

                data = self.convert_data(raw_data, type_var.get(), selected_attr)
                if data is None:
                    return

                # 找到选中列的第一个空值位置
                if selected_attr not in self.df.columns:
                    start_idx = 0
                else:
                    # 使用 pd.isna 来检测空值（包括 None, NaN 等）
                    mask = pd.isna(self.df[selected_attr])
                    start_idx = 0 if mask.all() else mask.idxmax() if mask.any() else len(self.df)

                # 确保 DataFrame 有足够的行
                needed_length = max(len(self.df), start_idx + len(data))
                if needed_length > len(self.df):
                    new_rows = pd.DataFrame(index=range(len(self.df), needed_length))
                    self.df = pd.concat([self.df, new_rows])

                # 从找到的位置开始插入数据
                for i, value in enumerate(data):
                    self.df.at[start_idx + i, selected_attr] = value

                # 保存并更新界面
                if self.save_file():
                    self.update_preview(preview_text)
                    data_entry.delete(0, tk.END)
                    messagebox.showinfo("成功", "数据已更新")

            except Exception as e:
                messagebox.showerror("错误", f"更新数据时发生错误：{str(e)}")

        # 功能按钮区
        button_frame = ctk.CTkFrame(right_frame)
        button_frame.pack(fill="x", pady=10)
        
        ctk.CTkButton(
            button_frame,
            text="更新数据",
            command=update_data,  # 现在 update_data 已定义
            width=120
        ).pack(pady=5)
        
        ctk.CTkButton(
            button_frame,
            text="清空输入",
            command=lambda: data_entry.delete(0, tk.END),
            width=120
        ).pack(pady=5)
        
        ctk.CTkButton(
            button_frame,
            text="统计信息",
            command=self.show_statistics,
            width=120
        ).pack(pady=5)

    def show_statistics(self):
        """显示数据统计信息"""
        try:
            stats_window = ctk.CTkToplevel(self)
            stats_window.title("数据统计")
            stats_window.geometry("600x400")
            
            # 创建统计信息文本框
            stats_text = ctk.CTkTextbox(
                stats_window,
                wrap="none",
                font=("Courier", 12)
            )
            stats_text.pack(fill="both", expand=True, padx=10, pady=10)
            
            # 计算统计信息
            stats_info = []
            stats_info.append(f"总行数: {len(self.df)}")
            stats_info.append(f"总列数: {len(self.df.columns)}")
            
            for column in self.df.columns:
                stats_info.append(f"\n{column} 统计信息:")
                if pd.api.types.is_numeric_dtype(self.df[column]):
                    stats = self.df[column].describe()
                    for stat_name, value in stats.items():
                        stats_info.append(f"{stat_name}: {value:.2f}")
                else:
                    unique_count = self.df[column].nunique()
                    stats_info.append(f"唯一值数量: {unique_count}")
            
            stats_text.insert("1.0", "\n".join(stats_info))
            
        except Exception as e:
            messagebox.showerror("错误", f"生成统计信息时发生错误：{str(e)}")

    def toggle_theme(self):
        """切换深色/浅色主题"""
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")

if __name__ == "__main__":
    app = SKUManager()
    app.mainloop()