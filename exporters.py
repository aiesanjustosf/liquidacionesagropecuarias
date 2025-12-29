# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from parser import Liquidacion


VENTAS_COLUMNS = [
    "Fecha dd/mm/aaaa","Cpbte","Tipo","Suc.","Número",
    "Razón Social o Denominación Cliente ",
    "Tipo Doc.","CUIT","Domicilio","C.P.","Pcia","Cond Fisc",
    "Cód. Neto","Neto Gravado","Alíc.","IVA Liquidado","IVA Débito",
    "Cód. NG/EX","Conceptos NG/EX","Cód. P/R","Perc./Ret.","Pcia P/R","Total"
]

COMPRAS_COLUMNS = [
    "Fecha Emisión ","Fecha Recepción","Cpbte","Tipo","Suc.","Número",
    "Razón Social/Denominación Proveedor",
    "Tipo Doc.","CUIT","Domicilio","C.P.","Pcia","Cond Fisc",
    "Cód. Neto","Neto Gravado","Alíc.","IVA Liquidado","IVA Crédito",
    "Cód. NG/EX","Conceptos NG/EX","Cód. P/R","Perc./Ret.","Pcia P/R","Total"
]


def _digits_to_int_or_none(s: str):
    s = (s or "").strip()
    if not s:
        return None
    if s.isdigit():
        try:
            return int(s)
        except Exception:
            return s
    return s


def build_ventas_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        rows.append({
            "Fecha dd/mm/aaaa": l.fecha,
            "Cpbte": l.tipo_cbte,          # F1/F2
            "Tipo": l.letra,              # A
            "Suc.": l.pv,
            "Número": l.numero,
            "Razón Social o Denominación Cliente ": (l.comprador.razon_social or "").strip(),
            "Tipo Doc.": 80,
            "CUIT": _digits_to_int_or_none(l.comprador.cuit),
            "Domicilio": (l.comprador.domicilio or "").strip(),
            "C.P.": "",
            "Pcia": "",
            "Cond Fisc": l.comprador.cond_fisc,
            "Cód. Neto": l.cod_neto_venta,
            "Neto Gravado": float(l.neto or 0.0),
            "Alíc.": float(l.alic_iva or 0.0),          # VALOR 10.5 (no 0.105)
            "IVA Liquidado": float(l.iva or 0.0),
            "IVA Débito": float(l.iva or 0.0),
            "Cód. NG/EX": "",
            "Conceptos NG/EX": None,
            "Cód. P/R": "",
            "Perc./Ret.": None,
            "Pcia P/R": "",
            "Total": float(l.total or 0.0),
        })

        # Retención IVA: SOLO RA07 (Ganancias ignoradas)
        if float(l.ret_iva or 0.0) != 0.0:
            amt = float(l.ret_iva or 0.0)
            rows.append({
                "Fecha dd/mm/aaaa": l.fecha,
                "Cpbte": "RV",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social o Denominación Cliente ": (l.comprador.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": _digits_to_int_or_none(l.comprador.cuit),
                "Domicilio": (l.comprador.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.comprador.cond_fisc,
                "Cód. Neto": "",
                "Neto Gravado": None,
                "Alíc.": None,
                "IVA Liquidado": None,
                "IVA Débito": None,
                "Cód. NG/EX": "",
                "Conceptos NG/EX": None,
                "Cód. P/R": "RA07",
                "Perc./Ret.": amt,
                "Pcia P/R": "",
                "Total": amt,     # siempre con total
            })

    return pd.DataFrame(rows, columns=VENTAS_COLUMNS)


def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows = []
    for l in liqs:
        pv = (l.pv or "").zfill(4) if (l.pv or "").isdigit() else (l.pv or "")
        num = (l.numero or "").zfill(8) if (l.numero or "").isdigit() else (l.numero or "")
        comprobante = f"{pv}-{num}" if pv and num else ""

        rows.append({
            "FECHA": l.fecha,
            "COMPROBANTE": comprobante,  # SOLO 3302-29912534
            "ACOPIO": (l.acopio.razon_social or "").strip(),
            "TIPO DE GRANO": l.grano,
            "CAMPAÑA": l.campaña or "",
            "CANTIDAD DE KILOS": float(l.kilos or 0.0),
            "PRECIO": float(l.precio or 0.0),
            "LOCALIDAD": l.localidad,
            "ME - Nro comprobante": l.me_nro_comprobante,
            "ME - Grado": l.me_grado,
            "ME - Factor": l.me_factor if l.me_factor is not None else None,
            "ME - Contenido proteico": l.me_contenido_proteico if l.me_contenido_proteico is not None else None,
            "ME - Procedencia": l.me_procedencia,
            "ME - Peso (kg)": l.me_peso_kg if l.me_peso_kg is not None else None,
        })
    return pd.DataFrame(rows)


def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    """
    Modelo compras (HWCpra1):
    - Proveedor = acopio (encabezado)
    - Cpbte = ND
    - Tipo = A
    - Tipo de movimiento: 203 por defecto; si IVA 21% => 202
    - Exento (alíc 0%) puede ir en la misma línea (en NG/EX)
    - Si hay dos alícuotas (10.5 y 21) => líneas separadas.
    """
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        exento_total = 0.0
        by_alic = {}  # alic -> (neto, iva)

        for d in (l.deducciones or []):
            alic = float(d.alic or 0.0)
            if abs(alic) < 0.000001:
                exento_total += float(d.total if d.total else d.neto)
            else:
                by_alic.setdefault(alic, [0.0, 0.0])
                by_alic[alic][0] += float(d.neto or 0.0)
                by_alic[alic][1] += float(d.iva or 0.0)

        alics_sorted = sorted(by_alic.keys())
        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0
                mov = 202 if abs(alic - 21.0) < 0.001 else 203
                total = (neto or 0.0) + (iva or 0.0) + (exento_here or 0.0)

                rows.append({
                    "Fecha Emisión ": l.fecha,
                    "Fecha Recepción": l.fecha,
                    "Cpbte": "ND",
                    "Tipo": l.letra,  # A
                    "Suc.": l.pv,
                    "Número": l.numero,
                    "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                    "Tipo Doc.": 80,
                    "CUIT": _digits_to_int_or_none(l.acopio.cuit),
                    "Domicilio": (l.acopio.domicilio or "").strip(),
                    "C.P.": "",
                    "Pcia": "",
                    "Cond Fisc": l.acopio.cond_fisc,
                    "Cód. Neto": mov,
                    "Neto Gravado": float(neto or 0.0),
                    "Alíc.": float(alic or 0.0),
                    "IVA Liquidado": float(iva or 0.0),
                    "IVA Crédito": float(iva or 0.0),
                    "Cód. NG/EX": 203 if exento_here else "",
                    "Conceptos NG/EX": float(exento_here) if exento_here else None,
                    "Cód. P/R": "",
                    "Perc./Ret.": None,
                    "Pcia P/R": "",
                    "Total": float(total or 0.0),
                })
        else:
            # Only exento
            mov = 203
            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": "ND",
                "Tipo": l.letra,  # A
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": _digits_to_int_or_none(l.acopio.cuit),
                "Domicilio": (l.acopio.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.acopio.cond_fisc,
                "Cód. Neto": mov,
                "Neto Gravado": 0.0,
                "Alíc.": None,
                "IVA Liquidado": 0.0,
                "IVA Crédito": 0.0,
                "Cód. NG/EX": 203,
                "Conceptos NG/EX": float(exento_total) if exento_total else None,
                "Cód. P/R": "",
                "Perc./Ret.": None,
                "Pcia P/R": "",
                "Total": float(exento_total or 0.0),
            })

    return pd.DataFrame(rows, columns=COMPRAS_COLUMNS)


# ------------------------- XLSX with correct formats -------------------------

def _set_col_widths(ws, widths: List[float]):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_font = Font(bold=True)
    center = Alignment(vertical="center")

    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = header_font
        c.alignment = center

    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    col_idx = {name: i + 1 for i, name in enumerate(df.columns)}

    # Detect sheet type by columns
    cols_set = set(df.columns)

    # Formats
    fmt_amount = '#,##0.00'
    fmt_alic = '0.000'
    fmt_cuit = '0'

    def apply_format(col_name: str, fmt: str):
        if col_name not in col_idx:
            return
        j = col_idx[col_name]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=j)
            if cell.value is None or cell.value == "":
                continue
            cell.number_format = fmt

    # Ventas
    if set(VENTAS_COLUMNS).issubset(cols_set):
        for nm in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
            apply_format(nm, fmt_amount)
        apply_format("Alíc.", fmt_alic)
        apply_format("CUIT", fmt_cuit)

        _set_col_widths(ws, [
            14, 8, 6, 6, 12, 42, 9, 14, 30, 8, 8, 10,
            10, 14, 10, 14, 14, 10, 14, 10, 14, 10, 14
        ])

    # Compras
    elif set(COMPRAS_COLUMNS).issubset(cols_set):
        for nm in ["Neto Gravado", "IVA Liquidado", "IVA Crédito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
            apply_format(nm, fmt_amount)
        apply_format("Alíc.", fmt_alic)
        apply_format("CUIT", fmt_cuit)

        _set_col_widths(ws, [
            14, 14, 8, 6, 6, 12, 42, 9, 14, 30, 8, 8, 10,
            10, 14, 10, 14, 14, 10, 14, 10, 14, 10, 14
        ])

    # CPNs (una sola hoja)
    else:
        # aplicar formatos por nombre si existen
        if "CANTIDAD DE KILOS" in col_idx:
            apply_format("CANTIDAD DE KILOS", fmt_amount)
        if "PRECIO" in col_idx:
            apply_format("PRECIO", '"$"#,##0.00')
        for nm in ["ME - Factor", "ME - Contenido proteico", "ME - Peso (kg)"]:
            if nm in col_idx:
                apply_format(nm, fmt_amount)

        # widths razonables
        widths = []
        for i, name in enumerate(df.columns, start=1):
            widths.append(min(max(len(str(name)) + 2, 12), 45))
        _set_col_widths(ws, widths)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
