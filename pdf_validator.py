# -*- coding: utf-8 -*-
"""
Validador de Nóminas PDF - Versión GUI
Valida que los nombres de archivo de las nóminas PDF divididas coincidan con su contenido

Funcionalidades:
1. Importar carpeta que contiene PDFs divididos
2. Validar la coincidencia entre nombre de archivo y contenido del PDF
3. Generar reporte Excel con marcas de color
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
        self.root.title("Herramienta de Validación de Nóminas PDF")
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
            text="Herramienta de Validación de Nóminas PDF",
            font=("Arial", 18, "bold"),
            bg="#4472C4",
            fg="white"
        )
        title_label.pack(expand=True)

        # 主框架
        main_frame = tk.Frame(self.root, padx=30, pady=30)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Texto de descripción
        info_text = (
            "Esta herramienta valida que los nombres de archivo de las nóminas PDF\n"
            "divididas coincidan con su contenido\n\n"
            "Contenido de validación：\n"
            "• El código del archivo coincide con el código del PDF\n"
            "• El nombre del archivo coincide con el nombre del PDF\n\n"
            "Se generará un reporte Excel con marcas de color：\n"
            "• Verde = Coincide ✓\n"
            "• Rojo = No coincide ✗"
        )

        info_label = tk.Label(
            main_frame,
            text=info_text,
            justify=tk.LEFT,
            font=("Arial", 10),
            bg="white"
        )
        info_label.pack(pady=(0, 20))

        # Botón de selección de carpeta
        select_btn = tk.Button(
            main_frame,
            text="Seleccionar carpeta de PDFs",
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
        folder = filedialog.askdirectory(title="Seleccionar carpeta que contiene archivos PDF")
        if folder:
            self.folder_path = folder
            self.validate_pdfs()

    def update_progress(self, current, total, filename):
        """Actualizar visualización de progreso"""
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = current
        self.progress_label.config(text=f"Validando: {current}/{total} - {filename}")
        self.root.update_idletasks()

    def validate_pdfs(self):
        if not self.folder_path:
            return

        # Reiniciar visualización
        self.progress_bar['value'] = 0
        self.status_label.config(text="Validando...", fg="#007ACC")
        self.root.update_idletasks()

        try:
            # Ejecutar validación
            self.results = validate_folder(self.folder_path, self.update_progress)

            if not self.results:
                messagebox.showwarning("Advertencia", "¡No se encontraron archivos PDF!")
                self.status_label.config(text="No se encontraron archivos PDF", fg="#FF0000")
                return

            # Generar reporte Excel
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(
                self.folder_path,
                f"Reporte_Validacion_{timestamp}.xlsx"
            )

            generate_excel_report(self.results, report_path)

            # Mostrar resultados
            total = len(self.results)
            matched = sum(1 for r in self.results if r['overall_match'])
            unmatched = total - matched

            result_msg = (
                f"¡Validación completada!\n\n"
                f"Total de archivos: {total}\n"
                f"Coinciden: {matched}\n"
                f"No coinciden: {unmatched}\n"
                f"Tasa de coincidencia: {matched/total*100:.1f}%\n\n"
                f"Reporte guardado en:\n{report_path}"
            )

            messagebox.showinfo("Validación completada", result_msg)
            self.status_label.config(
                text=f"Validación completada - {matched}/{total} coinciden",
                fg="#00AA00" if matched == total else "#FF6600"
            )

            # Preguntar si abrir el reporte
            if messagebox.askyesno("Abrir reporte", "¿Desea abrir el reporte Excel?"):
                os.system(f'xdg-open "{report_path}"' if os.name != 'nt' else f'start excel "{report_path}"')

        except Exception as e:
            messagebox.showerror("Error", f"Error durante la validación:\n{str(e)}")
            self.status_label.config(text="Validación fallida", fg="#FF0000")


def main():
    """Función principal"""
    root = tk.Tk()
    app = PDFValidatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
