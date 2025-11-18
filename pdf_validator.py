# -*- coding: utf-8 -*-
"""
PDF Payslip Validator - GUI版本
验证拆分后的PDF工资单文件名与内容是否匹配

功能：
1. 导入包含拆分PDF的文件夹
2. 验证文件名与PDF内容的匹配度
3. 生成带颜色标记的Excel报告
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

# 导入核心验证模块
from validator_core import validate_folder, generate_excel_report


# ===== GUI应用 =====
class PDFValidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF工资单验证工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        self.folder_path = None
        self.results = None

        self.setup_ui()

    def setup_ui(self):
        # 标题
        title_frame = tk.Frame(self.root, bg="#4472C4", height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="PDF工资单验证工具",
            font=("Arial", 18, "bold"),
            bg="#4472C4",
            fg="white"
        )
        title_label.pack(expand=True)

        # 主框架
        main_frame = tk.Frame(self.root, padx=30, pady=30)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 说明文字
        info_text = (
            "此工具用于验证拆分后的PDF工资单文件名与内容是否匹配\n\n"
            "验证内容：\n"
            "• 文件名中的编号与PDF中的编号是否一致\n"
            "• 文件名中的姓名与PDF中的姓名是否一致\n\n"
            "将生成带颜色标记的Excel报告：\n"
            "• 绿色 = 匹配 ✓\n"
            "• 红色 = 不匹配 ✗"
        )

        info_label = tk.Label(
            main_frame,
            text=info_text,
            justify=tk.LEFT,
            font=("Arial", 10),
            bg="white"
        )
        info_label.pack(pady=(0, 20))

        # 选择文件夹按钮
        select_btn = tk.Button(
            main_frame,
            text="选择PDF文件夹",
            command=self.select_folder,
            font=("Arial", 12, "bold"),
            bg="#4472C4",
            fg="white",
            padx=20,
            pady=10,
            cursor="hand2"
        )
        select_btn.pack(pady=10)

        # 进度显示
        self.progress_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10),
            fg="#666"
        )
        self.progress_label.pack(pady=5)

        # 进度条
        self.progress_bar = ttk.Progressbar(
            main_frame,
            mode='determinate',
            length=400
        )
        self.progress_bar.pack(pady=10)

        # 状态标签
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10, "bold"),
            fg="#007ACC"
        )
        self.status_label.pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder:
            self.folder_path = folder
            self.validate_pdfs()

    def update_progress(self, current, total, filename):
        """更新进度显示"""
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = current
        self.progress_label.config(text=f"正在验证: {current}/{total} - {filename}")
        self.root.update_idletasks()

    def validate_pdfs(self):
        if not self.folder_path:
            return

        # 重置显示
        self.progress_bar['value'] = 0
        self.status_label.config(text="正在验证中...", fg="#007ACC")
        self.root.update_idletasks()

        try:
            # 执行验证
            self.results = validate_folder(self.folder_path, self.update_progress)

            if not self.results:
                messagebox.showwarning("警告", "未找到PDF文件！")
                self.status_label.config(text="未找到PDF文件", fg="#FF0000")
                return

            # 生成Excel报告
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(
                self.folder_path,
                f"验证报告_{timestamp}.xlsx"
            )

            generate_excel_report(self.results, report_path)

            # 显示结果
            total = len(self.results)
            matched = sum(1 for r in self.results if r['overall_match'])
            unmatched = total - matched

            result_msg = (
                f"验证完成！\n\n"
                f"总文件数: {total}\n"
                f"匹配: {matched}\n"
                f"不匹配: {unmatched}\n"
                f"匹配率: {matched/total*100:.1f}%\n\n"
                f"报告已保存至:\n{report_path}"
            )

            messagebox.showinfo("验证完成", result_msg)
            self.status_label.config(
                text=f"验证完成 - {matched}/{total} 匹配",
                fg="#00AA00" if matched == total else "#FF6600"
            )

            # 询问是否打开报告
            if messagebox.askyesno("打开报告", "是否打开Excel报告？"):
                os.system(f'xdg-open "{report_path}"' if os.name != 'nt' else f'start excel "{report_path}"')

        except Exception as e:
            messagebox.showerror("错误", f"验证过程出错:\n{str(e)}")
            self.status_label.config(text="验证失败", fg="#FF0000")


def main():
    """主函数"""
    root = tk.Tk()
    app = PDFValidatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
