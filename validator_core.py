# -*- coding: utf-8 -*-
"""
Validador de Nóminas PDF - Módulo de validación central
No depende de GUI, procesamiento de lógica pura
"""

import os
import re
import unicodedata
from datetime import datetime
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ===== Funciones de utilidad =====
def strip_accents(s: str) -> str:
    """Eliminar acentos"""
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def normalize_name(name: str) -> str:
    """Normalizar nombre: quitar acentos, convertir a mayúsculas, eliminar espacios extra"""
    name = strip_accents(name)
    name = name.upper()
    name = re.sub(r'\s+', ' ', name)
    name = name.strip()
    # Eliminar comas y guiones bajos
    name = name.replace(',', ' ').replace('_', ' ')
    name = re.sub(r'\s+', ' ', name)
    return name.strip()


def parse_filename(filename: str) -> dict:
    """
    Analizar nombre de archivo
    Formato: 809_MARTINEZ MONTERO_ LAURA MARIA_Payslip_01092024.pdf
    Retorna: {'codigo': '809', 'nombre': 'MARTINEZ MONTERO LAURA MARIA', 'fecha': '01092024'}
    """
    # Eliminar extensión .pdf
    name = filename.replace('.pdf', '').replace('.PDF', '')

    # Intentar coincidir formato: código_nombre_Payslip_fecha
    # Código puede ser 2-6 dígitos
    pattern = r'^(\d{2,6})_(.*?)_Payslip_(.+)$'
    match = re.match(pattern, name, re.IGNORECASE)

    if match:
        codigo = match.group(1)
        nombre = normalize_name(match.group(2))
        fecha = match.group(3).replace('_', '')  # Eliminar posibles guiones bajos

        return {
            'codigo': codigo,
            'nombre': nombre,
            'fecha': fecha,
            'valid': True
        }
    else:
        return {
            'codigo': '',
            'nombre': '',
            'fecha': '',
            'valid': False,
            'error': 'Formato de nombre de archivo no cumple con la norma'
        }


def extract_pdf_info(pdf_path: str) -> dict:
    """
    Extraer información clave del PDF
    Retorna: {'codigo': '809', 'nombre': 'MARTINEZ MONTERO, LAURA MARIA', 'nif': '...', 'periodo': '...'}
    """
    try:
        with fitz.open(pdf_path) as doc:
            # Normalmente la información está en la primera página
            if doc.page_count > 0:
                page = doc[0]
                text = page.get_text('text')

                # Extraer código (formato 809/1)
                codigo = None
                codigo_pattern = r'\n\s*(\d{2,6})\/\d\s*\n'
                match = re.search(codigo_pattern, text)
                if match:
                    codigo = match.group(1)

                # Si no se encuentra, intentar otros patrones
                if not codigo:
                    # Intentar encontrar línea de código independiente
                    for line in text.split('\n')[:30]:
                        line = line.strip()
                        if re.match(r'^\d{2,6}$', line):
                            codigo = line
                            break

                # Extraer nombre (usualmente nombre completo en mayúsculas)
                nombre = None
                nombre_pattern = r'\n\s*\d{2,6}\/\d\s*\n\s*([A-ZÁÉÍÓÚÜÑ ,.\'\\-]+?)\s*\n'
                match = re.search(nombre_pattern, text)
                if match:
                    nombre = normalize_name(match.group(1))

                # Si no se encuentra, buscar líneas en mayúsculas que contengan coma (los nombres suelen tener coma)
                if not nombre:
                    lines = text.split('\n')[:60]
                    for line in lines:
                        line = line.strip()
                        if ',' in line and line.isupper() and len(line.split()) >= 2:
                            # Asegurar que no sea dirección u otra información
                            if 'CL ' not in line and 'PZ ' not in line and 'BARCELONA' not in line:
                                nombre = normalize_name(line)
                                break

                # Extraer NIF
                nif = None
                nif_pattern = r'N\.I\.F\.\s*([A-Z0-9]{8,})'
                match = re.search(nif_pattern, text)
                if match:
                    nif = match.group(1)

                # Extraer período
                periodo = None
                periodo_pattern = r'PERÍODO\s+(\d{1,2}\s+\w+\s+\d{1,2}\s+\w+\s+\d{4})'
                match = re.search(periodo_pattern, text, re.IGNORECASE)
                if match:
                    periodo = match.group(1)

                # Extraer Nº. Afiliación S.S.
                afiliacion = None
                afiliacion_pattern = r'Afiliación\s+S\.S\.\s+(\d+)'
                match = re.search(afiliacion_pattern, text, re.IGNORECASE)
                if match:
                    afiliacion = match.group(1)

                return {
                    'codigo': codigo or 'NO ENCONTRADO',
                    'nombre': nombre or 'NO ENCONTRADO',
                    'nif': nif or 'NO ENCONTRADO',
                    'periodo': periodo or 'NO ENCONTRADO',
                    'afiliacion': afiliacion or 'NO ENCONTRADO',
                    'valid': True
                }
            else:
                return {
                    'valid': False,
                    'error': 'PDF vacío'
                }

    except Exception as e:
        return {
            'valid': False,
            'error': f'Error al leer PDF: {str(e)}'
        }


def compare_info(filename_info: dict, pdf_info: dict) -> dict:
    """
    Comparar nombre de archivo y contenido del PDF
    Retorna resultado de validación
    """
    result = {
        'codigo_match': False,
        'nombre_match': False,
        'overall_match': False,
        'errors': []
    }

    # Comparar código
    if filename_info.get('codigo') and pdf_info.get('codigo'):
        if filename_info['codigo'] == pdf_info['codigo']:
            result['codigo_match'] = True
        else:
            result['errors'].append(f"Código no coincide: Archivo={filename_info['codigo']}, PDF={pdf_info['codigo']}")
    else:
        result['errors'].append("Información de código faltante")

    # Comparar nombre (después de normalizar)
    if filename_info.get('nombre') and pdf_info.get('nombre'):
        fn_nombre = normalize_name(filename_info['nombre'])
        pdf_nombre = normalize_name(pdf_info['nombre'])

        # Coincidencia exacta o muy similar
        if fn_nombre == pdf_nombre:
            result['nombre_match'] = True
        else:
            # Verificar si existe relación de contención (porque el formato puede variar ligeramente)
            fn_parts = set(fn_nombre.split())
            pdf_parts = set(pdf_nombre.split())

            # Si al menos el 80% de las palabras coinciden, considerar similar
            if fn_parts and pdf_parts:
                intersection = fn_parts & pdf_parts
                union = fn_parts | pdf_parts
                similarity = len(intersection) / len(union)

                if similarity >= 0.8:
                    result['nombre_match'] = True
                    result['errors'].append(f"Nombre coincide parcialmente (similitud {similarity:.0%}): Archivo={fn_nombre}, PDF={pdf_nombre}")
                else:
                    result['errors'].append(f"Nombre no coincide: Archivo={fn_nombre}, PDF={pdf_nombre}")
            else:
                result['errors'].append(f"Nombre no coincide: Archivo={fn_nombre}, PDF={pdf_nombre}")
    else:
        result['errors'].append("Información de nombre faltante")

    # Coincidencia general: código y nombre coinciden
    result['overall_match'] = result['codigo_match'] and result['nombre_match']

    return result


def validate_folder(folder_path: str, progress_callback=None) -> list:
    """
    Validar todos los archivos PDF en la carpeta
    Retorna lista de resultados de validación
    """
    results = []

    # Obtener todos los archivos PDF
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)

    if total_files == 0:
        return results

    for idx, filename in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx + 1, total_files, filename)

        pdf_path = os.path.join(folder_path, filename)

        # Analizar nombre de archivo
        filename_info = parse_filename(filename)

        # Extraer información del PDF
        pdf_info = extract_pdf_info(pdf_path)

        # Comparar información
        if filename_info.get('valid') and pdf_info.get('valid'):
            comparison = compare_info(filename_info, pdf_info)
        else:
            comparison = {
                'codigo_match': False,
                'nombre_match': False,
                'overall_match': False,
                'errors': []
            }
            if not filename_info.get('valid'):
                comparison['errors'].append(filename_info.get('error', 'Fallo al analizar nombre de archivo'))
            if not pdf_info.get('valid'):
                comparison['errors'].append(pdf_info.get('error', 'Fallo al analizar PDF'))

        # Consolidar resultados
        result = {
            'filename': filename,
            'fn_codigo': filename_info.get('codigo', ''),
            'fn_nombre': filename_info.get('nombre', ''),
            'fn_fecha': filename_info.get('fecha', ''),
            'pdf_codigo': pdf_info.get('codigo', ''),
            'pdf_nombre': pdf_info.get('nombre', ''),
            'pdf_nif': pdf_info.get('nif', ''),
            'pdf_periodo': pdf_info.get('periodo', ''),
            'pdf_afiliacion': pdf_info.get('afiliacion', ''),
            'codigo_match': comparison['codigo_match'],
            'nombre_match': comparison['nombre_match'],
            'overall_match': comparison['overall_match'],
            'errors': '; '.join(comparison['errors']) if comparison['errors'] else ''
        }

        results.append(result)

    return results


def generate_excel_report(results: list, output_path: str):
    """
    Generar reporte Excel, marcar resultados de validación con colores
    Verde: Coincidencia completa
    Rojo: No coincide
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de Validación"

    # Definir estilos
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Encabezados
    headers = [
        'Nombre de archivo',
        'Archivo-Código',
        'Archivo-Nombre',
        'Archivo-Fecha',
        'PDF-Código',
        'PDF-Nombre',
        'PDF-NIF',
        'PDF-Período',
        'PDF-Nº Seg. Social',
        'Código coincide',
        'Nombre coincide',
        'Resultado validación',
        'Descripción de error'
    ]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    # 数据行
    for row_num, result in enumerate(results, 2):
        # 基本信息
        ws.cell(row=row_num, column=1, value=result['filename']).border = border
        ws.cell(row=row_num, column=2, value=result['fn_codigo']).border = border
        ws.cell(row=row_num, column=3, value=result['fn_nombre']).border = border
        ws.cell(row=row_num, column=4, value=result['fn_fecha']).border = border
        ws.cell(row=row_num, column=5, value=result['pdf_codigo']).border = border
        ws.cell(row=row_num, column=6, value=result['pdf_nombre']).border = border
        ws.cell(row=row_num, column=7, value=result['pdf_nif']).border = border
        ws.cell(row=row_num, column=8, value=result['pdf_periodo']).border = border
        ws.cell(row=row_num, column=9, value=result['pdf_afiliacion']).border = border

        # Coincidencia de código
        codigo_cell = ws.cell(row=row_num, column=10, value='✓' if result['codigo_match'] else '✗')
        codigo_cell.border = border
        codigo_cell.alignment = Alignment(horizontal='center')
        codigo_cell.fill = green_fill if result['codigo_match'] else red_fill

        # Coincidencia de nombre
        nombre_cell = ws.cell(row=row_num, column=11, value='✓' if result['nombre_match'] else '✗')
        nombre_cell.border = border
        nombre_cell.alignment = Alignment(horizontal='center')
        nombre_cell.fill = green_fill if result['nombre_match'] else red_fill

        # Resultado general
        overall_cell = ws.cell(row=row_num, column=12, value='Coincide' if result['overall_match'] else 'No coincide')
        overall_cell.border = border
        overall_cell.alignment = Alignment(horizontal='center')
        overall_cell.fill = green_fill if result['overall_match'] else red_fill
        overall_cell.font = Font(bold=True)

        # Descripción de error
        error_cell = ws.cell(row=row_num, column=13, value=result['errors'])
        error_cell.border = border
        error_cell.alignment = Alignment(wrap_text=True)
        if result['errors']:
            error_cell.fill = yellow_fill

    # 调整列宽
    column_widths = [35, 12, 30, 15, 12, 30, 15, 25, 18, 10, 10, 12, 50]
    for col_num, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(col_num)].width = width

    # Congelar primera fila
    ws.freeze_panes = 'A2'

    # Añadir información estadística
    total = len(results)
    matched = sum(1 for r in results if r['overall_match'])
    unmatched = total - matched

    stats_row = len(results) + 3
    ws.cell(row=stats_row, column=1, value="Información estadística").font = Font(bold=True, size=12)
    ws.cell(row=stats_row + 1, column=1, value=f"Total de archivos: {total}")
    ws.cell(row=stats_row + 2, column=1, value=f"Coinciden: {matched}").fill = green_fill
    ws.cell(row=stats_row + 3, column=1, value=f"No coinciden: {unmatched}").fill = red_fill
    ws.cell(row=stats_row + 4, column=1, value=f"Tasa de coincidencia: {matched/total*100:.1f}%" if total > 0 else "N/A")

    # Guardar
    wb.save(output_path)
