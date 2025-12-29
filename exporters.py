# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any

import pandas as pd
from openpyxl import load_workbook

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


def build_ventas_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        rows.append({
            "Fecha dd/mm/aaaa": l.fecha,
            "Cpbte": l.tipo_cbte,
            "Tipo": l.letra,
            "Suc.": l.pv,
            "Número": l.numero,
            "Razón Social o Denominación Cliente ": (l.acopio.razon_social or "").strip(),
            "Tipo Doc.": 80,
            "CUIT": l.acopio.cuit,
            "Domicilio": (l.acopio.domicilio or "").strip(),
            "C.P.": "",
            "Pcia": "",
            "Cond Fisc": l.acopio.cond_fisc,
            "Cód. Neto": l.cod_neto_venta,
            "Neto Gravado": float(l.neto or 0.0),
            "Alíc.": float(l.alic_iva or 0.0),
            "IVA Liquidado": float(l.iva or 0.0),
            "IVA Débito": float(l.iva or 0.0),
            "Cód. NG/EX": None,
            "Conceptos NG/EX": None,
            "Cód. P/R": None,
            "Perc./Ret.": None,
            "Pcia P/R": None,
            "Total": float(l.total or 0.0),
        })

        # SOLO RA07 (IVA). Una sola línea por comprobante.
        amt = float(l.ret_iva or 0.0)
        if abs(amt) > 1e-9:
            rows.append({
                "Fecha dd/mm/aaaa": l.fecha,
                "Cpbte": "RV",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social o Denominación Cliente ": (l.acopio.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": l.acopio.cuit,
                "Domicilio": (l.acopio.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": l.acopio.cond_fisc,
                "Cód. Neto": None,
                "Neto Gravado": None,
                "Alíc.": None,
                "IVA Liquidado": None,
                "IVA Débito": None,
                "Cód. NG/EX": None,
                "Conceptos NG/EX": None,
                "Cód. P/R": "RA07",
                "Perc./Ret.": amt,
                "Pcia P/R": None,
                "Total": amt,
            })

    df = pd.DataFrame(rows, columns=VENTAS_COLUMNS)

    # Forzar columnas numéricas a numérico (evita que Excel las trate como texto)
    num_cols = ["Neto Gravado", "Alíc.", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows = []
    for l in liqs:
        # Si querés específicamente "3302-29912534": pv-numero
        comprobante = f"{l.pv}-{l.numero}"
        rows.append({
            "Fecha": l.fecha,
            "COE": l.coe,
            "Comprobante": comprobante,
            "Acopio/Comprador": (l.acopio.razon_social or "").strip(),
            "CUIT Comprador": l.acopio.cuit,
            "Tipo de grano": l.grano,
            "Campaña": l.campaña or "",
            "Kilos": float(l.kilos or 0.0),
            "Precio": float(l.precio or 0.0),
            "Subtotal": float(l.neto or 0.0),
            "Alic IVA": float(l.alic_iva or 0.0),
            "IVA": float(l.iva or 0.0),
            "Total": float(l.total or 0.0),
        })
    df = pd.DataFrame(rows)
    for c in ["Kilos", "Precio", "Subtotal", "Alic IVA", "IVA", "Total"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        exento_total = 0.0
        by_alic = {}
        for d in l.deducciones:
            if (d.alic or 0) == 0:
                exento_total += (d.total if d.total else d.neto)
            else:
                by_alic.setdefault(d.alic, [0.0, 0.0])
                by_alic[d.alic][0] += d.neto
                by_alic[d.alic][1] += d.iva

        alics_sorted = sorted(by_alic.keys())
        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0
                mov = 202 if abs(alic - 21.0) < 0.001 else 203
                total = (neto or 0) + (iva or 0) + (exento_here or 0)

                rows.append({
                    "Fecha Emisión ": l.fecha,
                    "Fecha Recepción": l.fecha,
                    "Cpbte": mov,
                    "Tipo": "",
                    "Suc.": l.pv,
                    "Número": l.numero,
                    "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                    "Tipo Doc.": 80,
                    "CUIT": l.acopio.cuit,
                    "Domicilio": (l.acopio.domicilio or "").strip(),
                    "C.P.": "",
                    "Pcia": "",
                    "Cond Fisc": l.acopio.cond_fisc,
                    "Cód. Neto": mov,
                    "Neto Gravado": float(neto or 0.0),
                    "Alíc.": float(alic or 0.0),
                    "IVA Liquidado": float(iva or 0.0),
                    "IVA Crédito": float(iva or 0.0),
                    "Cód. NG/EX": 203 if exento_here else None,
                    "Conceptos NG/EX": float(exento_here) if exento_here else None,
                    "Cód. P/R": None,
                    "Perc./Ret.": None,
                    "Pcia P/R": None,
                    "Total": float(total or 0.0),
                })
        else:
            mov = 203
            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": mov,
                "Tipo": "",
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (l.acopio.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": l.acopio.cuit,
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
                "Cód. P/R": None,
                "Perc./Ret.": None,
                "Pcia P/R": None,
                "Total": float(exento_total or 0.0),
            })

    df = pd.DataFrame(rows, columns=COMPRAS_COLUMNS)
    num_cols = ["Neto Gravado", "Alíc.", "IVA Liquidado", "IVA Crédito", "Conceptos NG/EX", "Perc./Ret.", "Total"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def _apply_formats(xlsx_bytes: bytes, sheet_name: str) -> bytes:
    wb = load_workbook(BytesIO(xlsx_bytes))
    ws = wb[sheet_name]

    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    money_fmt = '#.##0,00'
    aliq_fmt = '0,000'
    price_fmt = '"$"#.##0,00'

    def set_fmt(colname: str, fmt: str):
        if colname not in col:
            return
        j = col[colname]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=j)
            if isinstance(cell.value, (int, float)) and cell.value is not None:
                cell.number_format = fmt

    if sheet_name == "Ventas":
        set_fmt("Neto Gravado", money_fmt)
        set_fmt("Alíc.", aliq_fmt)
        set_fmt("IVA Liquidado", money_fmt)
        set_fmt("IVA Débito", money_fmt)
        set_fmt("Perc./Ret.", money_fmt)
        set_fmt("Total", money_fmt)

    if sheet_name == "CPNs":
        set_fmt("Kilos", money_fmt)
        set_fmt("Precio", price_fmt)
        set_fmt("Subtotal", money_fmt)
        set_fmt("Alic IVA", aliq_fmt)
        set_fmt("IVA", money_fmt)
        set_fmt("Total", money_fmt)

    if sheet_name == "Gastos":
        set_fmt("Neto Gravado", money_fmt)
        set_fmt("Alíc.", aliq_fmt)
        set_fmt("IVA Liquidado", money_fmt)
        set_fmt("IVA Crédito", money_fmt)
        set_fmt("Conceptos NG/EX", money_fmt)
        set_fmt("Perc./Ret.", money_fmt)
        set_fmt("Total", money_fmt)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    return _apply_formats(output.getvalue(), sheet_name)
