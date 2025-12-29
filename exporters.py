# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from parser import Liquidacion

# Formatos (Excel AR)
FMT_NUM2 = "#.##0,00"      # 87.300.000,00
FMT_NUM3 = "0,000"         # 10,000
FMT_MONEY = '"$"#.##0,00'  # $1.000,00
FMT_CUIT = "0"


VENTAS_COLUMNS = [
    "Fecha dd/mm/aaaa","Cpbte","Tipo","Suc.","Número",
    "Razón Social o Denominación Cliente ",
    "Tipo Doc.","CUIT","Domicilio","C.P.","Pcia","Cond Fisc",
    "Cód. Neto","Neto Gravado","Alíc.","IVA Liquidado","IVA Débito",
    "Cód. NG/EX","Conceptos NG/EX","Cód. P/R","Perc./Ret.","Pcia P/R","Total"
]

# OJO: se QUITA la columna "Tipo" (la que te salía con 203)
COMPRAS_COLUMNS = [
    "Fecha Emisión ","Fecha Recepción","Cpbte","Suc.","Número",
    "Razón Social/Denominación Proveedor",
    "Tipo Doc.","CUIT","Domicilio","C.P.","Pcia","Cond Fisc",
    "Cód. Neto","Neto Gravado","Alíc.","IVA Liquidado","IVA Crédito",
    "Cód. NG/EX","Conceptos NG/EX","Cód. P/R","Perc./Ret.","Pcia P/R","Total"
]


def _set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _apply_formats_by_header(ws, header_row=1):
    headers = [c.value for c in ws[header_row]]
    idx = {h: i + 1 for i, h in enumerate(headers)}

    money_headers = {
        "Neto Gravado", "IVA Liquidado", "IVA Débito", "IVA Crédito",
        "Conceptos NG/EX", "Perc./Ret.", "Total"
    }
    aliq_headers = {"Alíc."}
    cuit_headers = {"CUIT"}

    for r in range(header_row + 1, ws.max_row + 1):
        for h in money_headers:
            if h in idx:
                ws.cell(row=r, column=idx[h]).number_format = FMT_NUM2
        for h in aliq_headers:
            if h in idx:
                ws.cell(row=r, column=idx[h]).number_format = FMT_NUM3
        for h in cuit_headers:
            if h in idx:
                ws.cell(row=r, column=idx[h]).number_format = FMT_CUIT


def build_ventas_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        # Línea principal (venta)
        rows.append({
            "Fecha dd/mm/aaaa": l.fecha,
            "Cpbte": l.tipo_cbte,          # F1/F2
            "Tipo": l.letra,              # A
            "Suc.": l.pv,
            "Número": l.numero,
            "Razón Social o Denominación Cliente ": (l.comprador.razon_social or "").strip(),
            "Tipo Doc.": 80,
            "CUIT": l.comprador.cuit,
            "Domicilio": (l.comprador.domicilio or "").strip(),
            "C.P.": "",
            "Pcia": "",
            "Cond Fisc": l.comprador.cond_fisc,
            "Cód. Neto": l.cod_neto_venta,
            "Neto Gravado": l.neto,
            "Alíc.": l.alic_iva,
            "IVA Liquidado": l.iva,
            "IVA Débito": l.iva,
            "Cód. NG/EX": "",
            "Conceptos NG/EX": "",
            "Cód. P/R": "",
            "Perc./Ret.": "",
            "Pcia P/R": "",
            "Total": l.total,
        })

        # Retenciones: van en Perc./Ret. (Cód. P/R), NO en NG/EX
        def add_ret(code: str, amount: float):
            rows.append({
                "Fecha dd/mm/aaaa": l.fecha,
                "Cpbte": "RV",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social o Denominación Cliente ": (l.comprador.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": l.comprador.cuit,
                "Domicilio": (l.comprador.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.comprador.cond_fisc,
                "Cód. Neto": "",
                "Neto Gravado": "",
                "Alíc.": "",
                "IVA Liquidado": "",
                "IVA Débito": "",
                "Cód. NG/EX": "",
                "Conceptos NG/EX": "",
                "Cód. P/R": code,
                "Perc./Ret.": amount,
                "Pcia P/R": "",
                "Total": amount,
            })

        if (l.ret_iva or 0) != 0:
            add_ret("RA07", l.ret_iva)
        if (l.ret_gan or 0) != 0:
            add_ret("RA05", l.ret_gan)

    return pd.DataFrame(rows, columns=VENTAS_COLUMNS)


def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    """
    Modelo compras:
    - Cpbte = ND (fijo)
    - Se quita columna "Tipo"
    - Proveedor = comprador/acopio (NO vendedor)
    - Percepción IVA (si existe): Cód. P/R = P007 y monto en Perc./Ret.
    """
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        # Deducciones agrupadas por alícuota
        exento_total = 0.0
        by_alic = {}  # alic -> (neto, iva)

        for d in l.deducciones:
            if (d.alic or 0) == 0:
                exento_total += (d.total if d.total else d.neto)
            else:
                by_alic.setdefault(d.alic, [0.0, 0.0])
                by_alic[d.alic][0] += d.neto
                by_alic[d.alic][1] += d.iva

        alics_sorted = sorted(by_alic.keys())

        def add_line(neto, alic, iva, exento_here):
            mov = 202 if abs((alic or 0) - 21.0) < 0.001 else 203
            total = (neto or 0) + (iva or 0) + (exento_here or 0)

            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": "ND",          # FIX
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (l.comprador.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": l.comprador.cuit,
                "Domicilio": (l.comprador.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.comprador.cond_fisc,
                "Cód. Neto": mov,
                "Neto Gravado": neto,
                "Alíc.": alic,
                "IVA Liquidado": iva,
                "IVA Crédito": iva,
                "Cód. NG/EX": 203 if exento_here else "",
                "Conceptos NG/EX": exento_here if exento_here else "",
                "Cód. P/R": "",
                "Perc./Ret.": "",
                "Pcia P/R": "",
                "Total": total,
            })

        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0
                add_line(neto, alic, iva, exento_here)
        else:
            # Solo exento
            add_line(0.0, "", 0.0, exento_total if exento_total else 0.0)

        # Percepción IVA -> P007 en código y monto en Perc./Ret.
        perc_iva = float(getattr(l, "perc_iva", 0.0) or 0.0)
        if perc_iva != 0.0:
            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": "ND",
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (l.comprador.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": l.comprador.cuit,
                "Domicilio": (l.comprador.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.comprador.cond_fisc,
                "Cód. Neto": "",
                "Neto Gravado": "",
                "Alíc.": "",
                "IVA Liquidado": "",
                "IVA Crédito": "",
                "Cód. NG/EX": "",
                "Conceptos NG/EX": "",
                "Cód. P/R": "P007",
                "Perc./Ret.": perc_iva,
                "Pcia P/R": "",
                "Total": perc_iva,
            })

    return pd.DataFrame(rows, columns=COMPRAS_COLUMNS)


def build_excel_ventas(liqs: List[Liquidacion]) -> BytesIO:
    df = build_ventas_rows(liqs)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ventas")
        ws = writer.book["Ventas"]
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        _apply_formats_by_header(ws, header_row=1)
        _set_col_widths(ws, [14,8,6,7,12,40,10,14,22,8,8,10,10,14,8,14,14,10,14,10,14,10,14])
    out.seek(0)
    return out


def build_excel_gastos(liqs: List[Liquidacion]) -> BytesIO:
    df = build_gastos_rows(liqs)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Gastos")
        ws = writer.book["Gastos"]
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        _apply_formats_by_header(ws, header_row=1)
        _set_col_widths(ws, [14,14,8,7,12,40,10,14,22,8,8,10,10,14,8,14,14,10,14,10,14,10,14])
    out.seek(0)
    return out


def build_excel_cpns(liqs: List[Liquidacion]) -> BytesIO:
    """
    CPNs en UNA sola hoja.
    - COMPROBANTE: 3302-29912534 (pv-numero)
    - Si hay varios ítems de Mercadería Entregada -> varias filas (repitiendo datos base)
    - Retenciones: importes (no %), vienen del parser corregido
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "CPNs"

    headers = [
        "FECHA",
        "COE",
        "COMPROBANTE",
        "ACOPIO/COMPRADOR",
        "CUIT COMPRADOR",
        "TIPO DE GRANO",
        "CAMPAÑA",
        "KILOS",
        "PRECIO",
        "NETO",
        "IVA",
        "TOTAL",
        "RET IVA",
        "RET GAN",
        "ME - Nro comprobante",
        "ME - Grado",
        "ME - Factor",
        "ME - Contenido proteico",
        "ME - Procedencia",
        "ME - Peso (kg)",
    ]
    ws.append(headers)

    bold = Font(bold=True)
    for c in ws[1]:
        c.font = bold
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"

    for l in liqs:
        comp = f"{l.pv}-{l.numero}" if (l.pv and l.numero) else ""

        # lista de mercadería entregada (si no existe, caer al legacy)
        me_items = getattr(l, "me_items", None) or []
        if not me_items and (l.me_nro_comprobante or ""):
            me_items = [{
                "nro": l.me_nro_comprobante,
                "grado": l.me_grado,
                "factor": l.me_factor,
                "prot": l.me_contenido_proteico,
                "peso": l.me_peso_kg,
                "proced": l.me_procedencia,
            }]

        # si no hay nada, igual una fila base
        if not me_items:
            me_items = [{}]

        for it in me_items:
            ws.append([
                l.fecha or "",
                l.coe or "",
                comp,
                (l.comprador.razon_social or "").strip(),
                (l.comprador.cuit or "").replace("-", ""),
                l.grano or "",
                l.campaña or "",
                float(l.kilos or 0),
                float(l.precio or 0),
                float(l.neto or 0),
                float(l.iva or 0),
                float(l.total or 0),
                float(l.ret_iva or 0),
                float(l.ret_gan or 0),
                it.get("nro", "") or "",
                it.get("grado", "") or "",
                it.get("factor", "") if it.get("factor", None) is not None else "",
                it.get("prot", "") if it.get("prot", None) is not None else "",
                it.get("proced", "") or "",
                it.get("peso", "") if it.get("peso", None) is not None else "",
            ])

    # formatos
    idx = {h: i + 1 for i, h in enumerate(headers)}
    for r in range(2, ws.max_row + 1):
        ws.cell(r, idx["KILOS"]).number_format = FMT_NUM2
        ws.cell(r, idx["PRECIO"]).number_format = FMT_MONEY
        ws.cell(r, idx["NETO"]).number_format = FMT_NUM2
        ws.cell(r, idx["IVA"]).number_format = FMT_NUM2
        ws.cell(r, idx["TOTAL"]).number_format = FMT_NUM2
        ws.cell(r, idx["RET IVA"]).number_format = FMT_NUM2
        ws.cell(r, idx["RET GAN"]).number_format = FMT_NUM2
        ws.cell(r, idx["ME - Factor"]).number_format = FMT_NUM2
        ws.cell(r, idx["ME - Contenido proteico"]).number_format = FMT_NUM2
        ws.cell(r, idx["ME - Peso (kg)"]).number_format = FMT_NUM2

    _set_col_widths(ws, [
        12, 14, 16, 40, 14, 18, 12, 14, 14, 14, 14, 14, 12, 12,
        18, 12, 12, 20, 22, 14
    ])

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

