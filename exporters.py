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

        # SOLO RA07 (IVA). Ganancias NO se exportan.
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
                "Cód. Neto": "",
                "Neto Gravado": "",
                "Alíc.": "",
                "IVA Liquidado": "",
                "IVA Débito": "",
                "Cód. NG/EX": "",
                "Conceptos NG/EX": "",
                "Cód. P/R": "RA07",
                "Perc./Ret.": amt,
                "Pcia P/R": "",
                "Total": amt,
            })

    return pd.DataFrame(rows, columns=VENTAS_COLUMNS)


def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows = []
    for l in liqs:
        comprobante = f"{l.pv}-{l.numero}"  # si querés solo 3302-29912534
        rows.append({
            "FECHA": l.fecha,
            "COE": l.coe,
            "COMPROBANTE": comprobante,
            "ACOPIO": (l.acopio.razon_social or "").strip(),
            "CUIT": l.acopio.cuit,
            "TIPO DE GRANO": l.grano,
            "CAMPAÑA": l.campaña or "",
            "CANTIDAD DE KILOS": l.kilos,
            "PRECIO": l.precio,
            "NETO": l.neto,
            "ALIC IVA": l.alic_iva,
            "IVA": l.iva,
            "TOTAL": l.total,
        })
    return pd.DataFrame(rows)


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
                "Alíc.": "",
                "IVA Liquidado": 0.0,
                "IVA Crédito": 0.0,
                "Cód. NG/EX": 203,
                "Conceptos NG/EX": exento_total if exento_total else "",
                "Cód. P/R": "",
                "Perc./Ret.": "",
                "Pcia P/R": "",
                "Total": exento_total if exento_total else 0.0,
            })

    return pd.DataFrame(rows, columns=COMPRAS_COLUMNS)


def _apply_number_formats(xlsx_bytes: bytes, sheet_name: str) -> bytes:
    """
    Fuerza formato visual:
      - Montos: 1.000,00  -> '#.##0,00'
      - Alícuotas: 10,000 -> '0,000'
      - Precio con $: '"$"#.##0,00'
    Mantiene los valores como numéricos.
    """
    bio = BytesIO(xlsx_bytes)
    wb = load_workbook(bio)
    ws = wb[sheet_name]

    # header -> col idx
    header = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(header)}

    money_fmt = '#.##0,00'
    aliq_fmt = '0,000'
    price_fmt = '"$"#.##0,00'

    def set_fmt(colname: str, fmt: str):
        if colname not in col:
            return
        j = col[colname]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=j)
            if isinstance(cell.value, (int, float)):
                cell.number_format = fmt

    # Ventas
    if sheet_name == "Ventas":
        set_fmt("Neto Gravado", money_fmt)
        set_fmt("IVA Liquidado", money_fmt)
        set_fmt("IVA Débito", money_fmt)
        set_fmt("Perc./Ret.", money_fmt)
        set_fmt("Total", money_fmt)
        set_fmt("Alíc.", aliq_fmt)

    # CPNs (según el build_cpns_rows que te dejé arriba)
    if sheet_name == "CPNs":
        set_fmt("CANTIDAD DE KILOS", money_fmt)
        set_fmt("PRECIO", price_fmt)
        set_fmt("NETO", money_fmt)
        set_fmt("IVA", money_fmt)
        set_fmt("TOTAL", money_fmt)
        set_fmt("ALIC IVA", aliq_fmt)

    # Gastos
    if sheet_name == "Gastos":
        set_fmt("Neto Gravado", money_fmt)
        set_fmt("IVA Liquidado", money_fmt)
        set_fmt("IVA Crédito", money_fmt)
        set_fmt("Conceptos NG/EX", money_fmt)
        set_fmt("Perc./Ret.", money_fmt)
        set_fmt("Total", money_fmt)
        set_fmt("Alíc.", aliq_fmt)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    xlsx = output.getvalue()
    return _apply_number_formats(xlsx, sheet_name=sheet_name)
