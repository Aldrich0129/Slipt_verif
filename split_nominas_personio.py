# -*- coding: utf-8 -*-
"""
Script: split_nominas_personio_gui_v2_3.py

Cambios:
- Ventanas para seleccionar PDF/carpeta salida/(opcional) ejemplos.
- Modo por defecto: Heurístico (STRICT_MODE=False) + Fusionar cabeceras iguales (MERGE_DUP_HEADERS=True).
- Nombre de salida SOLO con fecha de primer día del mes:
  * Mes completo o solo mes+año -> DDMMYYYY (sin separadores): 01 09 2025 -> '01092025'
  * Periodo parcial dentro del mismo mes (p.ej., 1-22 feb 2025) -> DD_MM_YYYY: '01_02_2025'
- CSV resumen y log detallado. (Opc) comparación con PDFs de ejemplo.

Requisitos:
  pip install pymupdf
"""

import os
import re
import csv
import difflib
import unicodedata
import calendar
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF

# ===== Config (no se preguntan estos modos) =====
STRICT_MODE = False          # Heurístico
MERGE_DUP_HEADERS = True     # Fusionar cabeceras iguales consecutivas
DEBUG_MODE = False           # Cambiar a True para volcados por página (debug_text/)

# ===== Patrones y utilidades =====
MONTHS_ES = {
    'enero': '01','febrero': '02','marzo': '03','abril': '04','mayo': '05','junio': '06',
    'julio': '07','agosto': '08','septiembre': '09','setiembre': '09','octubre': '10',
    'noviembre': '11','diciembre': '12'
}
MONTH_PATTERN = re.compile(r"(?i)\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b")
YEAR_PATTERN  = re.compile(r"\b(20\d{2})\b")

RE_CODE_NAME_BLOCK = re.compile(r"\n\s*(\d{2,6})\/(\d)\s*\n\s*([A-ZÁÉÍÓÚÜÑ ,.'\-]+?)\s*\n")
RE_PERIODO = re.compile(r"(?i)per[ií]odo")
RE_DAY_MONTH = re.compile(r"\b(\d{1,2})\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b", re.I)
RE_NIF = re.compile(r"(?i)\b(NIF|DNI|N\.I\.F\.)\b\s*[:\-]?\s*([A-Z0-9]{5,})")
RE_AFILIACION = re.compile(r"(?i)afiliaci[oó]n\s*s\.s\.\s*\n\s*([0-9]{3,12})")
RE_GENERIC_CODE = re.compile(r"\b\d{3,6}\b")

def strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def sane(s: str) -> str:
    s = strip_accents(s)
    s = re.sub(r"[^A-Za-z0-9_\-]", "_", s)
    return re.sub(r"_+", "_", s).strip("._")[:200]

def get_lines(text: str):
    return [l.strip() for l in text.splitlines()]

def get_top_lines(text: str, n=25):
    return [l for l in get_lines(text) if l.strip()][:n]

def find_month_year_in_window(lines, start_idx=None, lookahead=12):
    """Busca mes y año en un rango; tolera líneas separadas."""
    L = len(lines)
    if start_idx is None:
        start_idx = 0
    end_idx = min(L, start_idx + lookahead)
    months = [(i, MONTH_PATTERN.search(lines[i]).group(1))
              for i in range(start_idx, end_idx) if MONTH_PATTERN.search(lines[i])]
    years  = [(i, YEAR_PATTERN.search(lines[i]).group(1))
              for i in range(start_idx, end_idx) if YEAR_PATTERN.search(lines[i])]
    best_pair = (None, None); best_distance = 10**9
    for mi, m in months:
        forward  = [(yi, y) for yi, y in years if yi >= mi and (yi - mi) <= 6]
        backward = [(yi, y) for yi, y in years if yi < mi and (mi - yi) <= 3]
        candidates = forward or backward
        for yi, y in candidates:
            d = abs(yi - mi)
            if d < best_distance:
                best_distance = d
                best_pair = (m, y)
    return best_pair

def find_month_year_anywhere(lines):
    months = [(i, MONTH_PATTERN.search(lines[i]).group(1)) for i in range(len(lines)) if MONTH_PATTERN.search(lines[i])]
    years  = [(i, YEAR_PATTERN.search(lines[i]).group(1)) for i in range(len(lines)) if YEAR_PATTERN.search(lines[i])]
    if not months or not years:
        return (None, None)
    best_pair = (None, None); best_distance = 10**9
    for mi, m in months:
        forward = [(yi, y) for yi, y in years if yi >= mi and (yi - mi) <= 8]
        pool = forward or years
        for yi, y in pool:
            d = abs(yi - mi)
            if d < best_distance:
                best_distance = d
                best_pair = (m, y)
    return best_pair

def extract_periodo_mes_anio(lines):
    """Devuelve ('mes', 'anio', idxPeriodo) si se deduce; si no, (None, None, None)."""
    idx = None
    for i, l in enumerate(lines):
        if RE_PERIODO.search(l):
            idx = i; break
    mes = anio = None
    if idx is not None:
        mes, anio = find_month_year_in_window(lines, idx, lookahead=12)
    if not (mes and anio):
        m2, y2 = find_month_year_anywhere(lines)
        mes = mes or m2; anio = anio or y2
    return (mes, anio, idx)

def extract_days_near_period(lines, idx_period, target_month):
    """Busca días (números) asociados al mes en la ventana a partir de PERÍODO."""
    days = []
    if idx_period is None:
        window = lines[:40]
    else:
        window = lines[idx_period: idx_period+12]
    txt = "\n".join(window)
    for m in RE_DAY_MONTH.finditer(txt):
        d = m.group(1)
        mon = m.group(2).lower()
        # Normaliza 'setiembre' a 'septiembre'
        if mon == 'setiembre': mon = 'septiembre'
        if mon == target_month:
            try:
                days.append(int(d))
            except:
                pass
    return sorted(set(days))

def build_suffix(period_lines):
    """
    Devuelve SIEMPRE 'DDMMYYYY' con día 01, sin guiones bajos,
    tanto para mes completo como para periodos parciales.
    """
    mes, anio, idxp = extract_periodo_mes_anio(period_lines)
    if not (mes and anio):
        return 'SIN_FECHA'

    mes_norm = mes.lower()
    if mes_norm == 'setiembre':
        mes_norm = 'septiembre'

    mm = MONTHS_ES.get(mes_norm)
    if not mm:
        return 'SIN_FECHA'

    # Formato compacto, sin guiones bajos:
    return f"01{mm}{anio}"


def detect_header(text: str):
    """Extrae código, nombre, PERÍODO y NIF. None si no parece cabecera."""
    U = text.upper()
    header_tokens = 0
    if "RECIBO DE NÓMINA" in U or "RECIBO DE NOMINA" in U:
        header_tokens += 1
    if RE_PERIODO.search(U):
        header_tokens += 1

    codigo = nombre = periodo_str = nif = None

    m = RE_CODE_NAME_BLOCK.search(text)
    if m:
        codigo = m.group(1)
        nombre = m.group(3).strip()

    for l in get_top_lines(text, 50):
        mn = RE_NIF.search(l)
        if mn and not nif:
            nif = mn.group(2).strip()

    lines = get_lines(text)
    mes, anio, _ = extract_periodo_mes_anio(lines)
    if mes and anio:
        # formato visible 'mes año' para el CSV
        periodo_str = f"{mes} {anio}"

    # Fallbacks
    if not codigo:
        for l in get_top_lines(text, 50):
            ma = RE_AFILIACION.search(l)
            if ma:
                codigo = ma.group(1)[:6]; break
        if not codigo:
            for l in get_top_lines(text, 50):
                mg = RE_GENERIC_CODE.search(l)
                if mg:
                    codigo = mg.group(); break

    if not nombre:
        cands = [l for l in get_top_lines(text, 60) if l.isupper() and "," in l and len(l.split())>=2]
        if cands:
            nombre = cands[0]
        else:
            ups = [l for l in get_top_lines(text, 60) if l.isupper() and len(l.split())>=2]
            if ups:
                nombre = sorted(ups, key=len, reverse=True)[0]

    score = (1 if header_tokens>=1 else 0) + (1 if nombre else 0) + (1 if periodo_str else 0)
    if score >= 2:
        return {
            'codigo':  codigo or 'SIN_CODIGO',
            'nombre':  nombre or 'SIN_NOMBRE',
            'periodo': periodo_str or 'SIN_PERIODO',
            'nif':     nif or 'SIN_NIF',
            'lines':   lines
        }
    return None

def split_name(nombre_completo: str):
    if not nombre_completo or nombre_completo=='SIN_NOMBRE':
        return 'DESCONOCIDO','DESCONOCIDO'
    partes = [p for p in nombre_completo.replace(',', ' ').split() if p]
    if len(partes)>=2:
        nombre = partes[-1]; apellidos = ' '.join(partes[:-1])
        return apellidos, nombre
    return nombre_completo, ''

def normalize_text_for_diff(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()

def compare_pdfs(p1: str, p2: str):
    def extract(path):
        out = []
        with fitz.open(path) as d:
            for pg in d:
                out.append(normalize_text_for_diff(pg.get_text('text')))
        return "\n".join(out)
    a = extract(p1); b = extract(p2)
    if a == b:
        return True, ''
    diff = difflib.unified_diff(b.splitlines(), a.splitlines(), lineterm='')
    return False, "\n".join(list(diff)[:60])

def LOG(msg, log_list):
    print(msg)
    log_list.append(msg)

# ----- Guardado de bloque -----
def save_block(doc, pages_idx, cabecera, out_dir, ex_dir, rows, log):
    ap, no = split_name(cabecera['nombre'])
    # Sufijo de fecha SOLO (según reglas)
    suffix = build_suffix(cabecera['lines'])  # '01092025' o '01_02_2025' o 'SIN_FECHA'
    # Identificador base: código o NIF si no hay código
    base_id = cabecera['codigo'] if cabecera['codigo'] != 'SIN_CODIGO' else (cabecera['nif'] if cabecera['nif'] != 'SIN_NIF' else 'SINID')
    fname = f"{sane(base_id)}_{sane(ap)}_{sane(no)}_Payslip_{sane(suffix)}.pdf"
    out_pdf = os.path.join(out_dir, fname)

    with fitz.open() as newdoc:
        for p in pages_idx:
            newdoc.insert_pdf(doc, from_page=p, to_page=p)
        newdoc.save(out_pdf)

    # Comparación con ejemplo
    comp, diff = 'SIN EJEMPLO',''
    if ex_dir:
        exact = os.path.join(ex_dir, os.path.basename(out_pdf))
        alt = None
        if not os.path.exists(exact):
            pref = f"{sane(base_id)}_"
            for candf in os.listdir(ex_dir):
                if candf.startswith(pref) and candf.lower().endswith('.pdf'):
                    alt = os.path.join(ex_dir, candf); break
        example = exact if os.path.exists(exact) else alt
        if example and os.path.exists(example):
            eq, d = compare_pdfs(out_pdf, example)
            comp = 'OK' if eq else 'DIFERENTE'
            diff = d

    rows.append([os.path.basename(out_pdf), base_id, cabecera['nombre'], cabecera['periodo'], len(pages_idx), comp, diff])
    LOG(f"Guardado: {os.path.basename(out_pdf)} (páginas: {len(pages_idx)})", log)
    return out_pdf

# ===== UI: ventanas =====
root = tk.Tk(); root.withdraw()
messagebox.showinfo(
    "Instrucciones",
    "1) Selecciona el PDF consolidado de nóminas.\n"
    "2) Selecciona la carpeta de salida para los PDFs individuales.\n"
    "3) (Opcional) Carpeta con nóminas de ejemplo para comparar."
)

pdf_path = filedialog.askopenfilename(title="Selecciona el PDF consolidado", filetypes=[("PDF", "*.pdf")])
if not pdf_path:
    messagebox.showerror("Error", "No se seleccionó PDF."); raise SystemExit(1)

out_dir = filedialog.askdirectory(title="Selecciona la carpeta de salida")
if not out_dir:
    messagebox.showerror("Error", "No se seleccionó carpeta de salida."); raise SystemExit(1)

ex_dir = filedialog.askdirectory(title="(Opcional) Carpeta con nóminas de ejemplo (Cancelar si no hay)")
if not ex_dir:
    ex_dir = None

csv_path = os.path.join(out_dir, 'resumen_nominas.csv')
log_path = os.path.join(out_dir, 'log_nominas.txt')
if DEBUG_MODE:
    dbg_dir = os.path.join(out_dir, 'debug_text')
    os.makedirs(dbg_dir, exist_ok=True)

# ===== Proceso =====
if not os.path.exists(pdf_path):
    messagebox.showerror('Error', f'No existe: {pdf_path}'); raise SystemExit(1)

log = []
rows = []

try:
    with fitz.open(pdf_path) as doc:
        current = None
        pages = []

        for idx in range(doc.page_count):
            page = doc[idx]
            text = page.get_text('text') or ''
            if DEBUG_MODE:
                dbg_dir = os.path.join(out_dir, 'debug_text'); os.makedirs(dbg_dir, exist_ok=True)
                with open(os.path.join(dbg_dir, f'page_{idx+1:04d}.txt'), 'w', encoding='utf-8') as f:
                    f.write(text)

            cand = detect_header(text)
            has_header = cand is not None

            # Heurístico + Fusión
            if has_header:
                if current is not None and pages:
                    same = MERGE_DUP_HEADERS and cand and current and \
                           (cand['codigo'] == current['codigo']) and \
                           ( (cand['periodo'] or '') == (current['periodo'] or '') )
                    if same:
                        pages.append(idx)
                        LOG(f"{idx+1:04d}: cabecera repetida fusionada con {current['codigo']}", log)
                        continue
                    # Guardar bloque anterior y abrir nuevo
                    save_block(doc, pages, current, out_dir, ex_dir, rows, log)
                    current = None; pages = []
                current = cand; pages = [idx]
                LOG(f"{idx+1:04d}: inicio -> {current['codigo']} | {current['nombre']} | {current['periodo']}", log)
            else:
                if current is not None:
                    pages.append(idx)
                    LOG(f"{idx+1:04d}: continuación de {current['codigo']}", log)
                else:
                    LOG(f"{idx+1:04d}: huérfana (sin cabecera y sin bloque activo)", log)

        # Guardar último bloque
        if current is not None and pages:
            save_block(doc, pages, current, out_dir, ex_dir, rows, log)

    # Escribir CSV y log
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(["archivo","codigo","nombre","periodo","paginas","comparacion","diferencias"])
        w.writerows(rows)

    with open(log_path, 'w', encoding='utf-8') as f:
        for l in log:
            f.write(l + "\n")
        f.write(f"\nTotal bloques: {len(rows)}\n")

    messagebox.showinfo('Completado', f"Total nóminas: {len(rows)}\nResumen: {csv_path}\nLog: {log_path}")

except Exception as e:
    messagebox.showerror('Error', f'Se produjo un error: {e}')
    raise