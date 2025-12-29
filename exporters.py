# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any, Optional

import pandas as pd
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


# Excel formats (ARG: 1.000,00 / alícuota 10,500)
FMT_AMOUNT = "#.##0,00"
FMT_ALIQ = "0,000"
FMT_MONEY = '"$"#.##0,00'


def _set_col_widths(ws, df: pd.DataFrame):
    # ancho simple: basado en header
    for i, col in enumerate(df.columns, start=1):
        w = max(10, min(60, len(str(col)) + 2))
        ws.column_dimensions[get_column_letter(i)].width = w


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
            "CUIT": (l.comprador.cuit or "").strip(),
            "Domicilio": (l.comprador.domicilio or "").strip(),
            "C.P.": "",
            "Pcia": "",
            "Cond Fisc": l.comprador.cond_fisc,
            "Cód. Neto": l.cod_neto_venta,
            "Neto Gravado": float(l.neto or 0),
            "Alíc.": float(l.alic_iva or 0),
            "IVA Liquidado": float(l.iva or 0),
            "IVA Débito": float(l.iva or 0),
            "Cód. NG/EX": "",
            "Conceptos NG/EX": "",
            "Cód. P/R": "",
            "Perc./Ret.": "",
            "Pcia P/R": "",
            "Total": float(l.total or 0),
        })

        # RET IVA: SOLO RA07 (Ganancias ignoradas)
        if float(l.ret_iva or 0) != 0:
            rows.append({
                "Fecha dd/mm/aaaa": l.fecha,
                "Cpbte": "RV",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social o Denominación Cliente ": (l.comprador.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": (l.comprador.cuit or "").strip(),
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
                "Conceptos NG/EX": "",
                "Cód. P/R": "RA07",
                "Perc./Ret.": float(l.ret_iva or 0),
                "Pcia P/R": "",
                "Total": float(l.ret_iva or 0),
            })

    return pd.DataFrame(rows, columns=VENTAS_COLUMNS)


def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for l in liqs:
        comprobante = f"{l.pv}-{l.numero}"  # SOLO 3302-29912534
        rows.append({
            "FECHA": l.fecha,
            "COMPROBANTE": comprobante,
            "ACOPIO": (l.acopio.razon_social or "").strip(),
            "CUIT ACOPIO": (l.acopio.cuit or "").strip(),
            "COMPRADOR": (l.comprador.razon_social or "").strip(),
            "CUIT COMPRADOR": (l.comprador.cuit or "").strip(),
            "TIPO DE GRANO": l.grano,
            "CAMPAÑA": l.campaña or "",
            "CANTIDAD DE KILOS": float(l.kilos or 0),
            "PRECIO": float(l.precio or 0),
            "NETO": float(l.neto or 0),
            "IVA": float(l.iva or 0),
            "TOTAL": float(l.total or 0),
            "RET IVA": float(l.ret_iva or 0),
            # sección ME (simple)
            "ME - Nro comprobante": l.me_nro_comprobante,
            "ME - Grado": l.me_grado,
            "ME - Factor": l.me_factor if l.me_factor is not None else "",
            "ME - Contenido proteico": l.me_contenido_proteico if l.me_contenido_proteico is not None else "",
            "ME - Peso (kg)": l.me_peso_kg if l.me_peso_kg is not None else "",
            "ME - Procedencia": l.me_procedencia,
        })
    return pd.DataFrame(rows)


def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    """
    Modelo compras:
    - Proveedor = acopio
    - Cpbte = ND
    - Tipo = letra (A)
    - Mov (202/203) va en Cód. Neto
    - Exento va en NG/EX (Cód. NG/EX = 203)
    - Total SIEMPRE informado
    - Percepciones (si existieran) irían en Cód. P/R + Perc./Ret. (P007), pero acá no las forzamos
    """
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        exento_total = 0.0
        by_alic: Dict[float, List[float]] = {}  # alic -> [neto, iva]

        for d in l.deducciones:
            if float(d.alic or 0) == 0:
                exento_total += float(d.total if d.total else d.neto)
            else:
                by_alic.setdefault(float(d.alic), [0.0, 0.0])
                by_alic[float(d.alic)][0] += float(d.neto or 0)
                by_alic[float(d.alic)][1] += float(d.iva or 0)

        alics_sorted = sorted(by_alic.keys())

        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0
                mov = 202 if abs(alic - 21.0) < 0.001 else 203
                total = float(neto or 0) + float(iva or 0) + float(exento_here or 0)

                rows.append({
                    "Fecha Emisión ": l.fecha,
                    "Fecha Recepción": l.fecha,
                    "Cpbte": "ND",
                    "Tipo": l.letra,          # A
                    "Suc.": l.pv,
                    "Número": l.numero,
                    "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                    "Tipo Doc.": 80,
                    "CUIT": (l.acopio.cuit or "").strip(),
                    "Domicilio": (l.acopio.domicilio or "").strip(),
                    "C.P.": "",
                    "Pcia": "",
                    "Cond Fisc": l.acopio.cond_fisc,
                    "Cód. Neto": mov,
                    "Neto Gravado": float(neto or 0),
                    "Alíc.": float(alic or 0),
                    "IVA Liquidado": float(iva or 0),
                    "IVA Crédito": float(iva or 0),
                    "Cód. NG/EX": 203 if exento_here else "",
                    "Conceptos NG/EX": float(exento_here) if exento_here else "",
                    "Cód. P/R": "",
                    "Perc./Ret.": "",
                    "Pcia P/R": "",
                    "Total": float(total or 0),
                })
        else:
            mov = 203
            total = float(exento_total or 0)
            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": "ND",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": (l.acopio.cuit or "").strip(),
                "Domicilio": (l.acopio.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.acopio.cond_fisc,
                "Cód. Neto": mov,
                "Neto Gravado": 0.0,
                "Alíc.": "",
                "IVA Liquidado": 0.0,
                "IVA Crédito": 0.0,
                "Cód. NG/EX": 203,
                "Conceptos NG/EX": float(exento_total or 0),
                "Cód. P/R": "",
                "Perc./Ret.": "",
                "Pcia P/R": "",
                "Total": total,
            })

    return pd.DataFrame(rows, columns=COMPRAS_COLUMNS)


def df_to_xlsx_bytes(
    df: pd.DataFrame,
    sheet_name: str,
    col_formats: Optional[Dict[str, str]] = None,
) -> bytes:
    """
    Escribe XLSX y aplica formatos EXCEL (no strings):
    - Montos: 1.000,00 => #.##0,00
    - Alícuotas: 10,500 => 0,000
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        bold = Font(bold=True)
        for cell in ws[1]:
            cell.font = bold
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"

        _set_col_widths(ws, df)

        if col_formats:
            # map header -> excel number_format
            header_index = {str(c.value): c.column for c in ws[1]}
            for col_name, fmt in col_formats.items():
                if col_name not in header_index:
                    continue
                col_idx = header_index[col_name]
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=col_idx).number_format = fmt

    return output.getvalue()


def ventas_xlsx_bytes(liqs: List[Liquidacion]) -> bytes:
    df = build_ventas_rows(liqs)
    formats = {
        "Neto Gravado": FMT_AMOUNT,
        "IVA Liquidado": FMT_AMOUNT,
        "IVA Débito": FMT_AMOUNT,
        "Perc./Ret.": FMT_AMOUNT,
        "Total": FMT_AMOUNT,
        "Alíc.": FMT_ALIQ,
    }
    return df_to_xlsx_bytes(df, "Ventas", formats)


def cpns_xlsx_bytes(liqs: List[Liquidacion]) -> bytes:
    df = build_cpns_rows(liqs)
    formats = {
        "CANTIDAD DE KILOS": FMT_AMOUNT,
        "PRECIO": FMT_MONEY,
        "NETO": FMT_AMOUNT,
        "IVA": FMT_AMOUNT,
        "TOTAL": FMT_AMOUNT,
        "RET IVA": FMT_AMOUNT,
    }
    return df_to_xlsx_bytes(df, "CPNs", formats)


def gastos_xlsx_bytes(liqs: List[Liquidacion]) -> bytes:
    df = build_gastos_rows(liqs)
    formats = {
        "Neto Gravado": FMT_AMOUNT,
        "IVA Liquidado": FMT_AMOUNT,
        "IVA Crédito": FMT_AMOUNT,
        "Conceptos NG/EX": FMT_AMOUNT,
        "Perc./Ret.": FMT_AMOUNT,
        "Total": FMT_AMOUNT,
        "Alíc.": FMT_ALIQ,
    }
    return df_to_xlsx_bytes(df, "Gastos", formats)
