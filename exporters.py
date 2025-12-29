# exporters.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any, Optional
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from parser import Liquidacion


# Formatos Excel (AR): miles con punto, decimales con coma
FMT_AMOUNT = "#.##0,00"
FMT_ALIQ = "0,000"
FMT_CURR = '"$"#.##0,00'
FMT_CUIT = "0"


VENTAS_COLUMNS = [
    "Fecha dd/mm/aaaa", "Cpbte", "Tipo", "Suc.", "Número",
    "Razón Social o Denominación Cliente ",
    "Tipo Doc.", "CUIT", "Domicilio", "C.P.", "Pcia", "Cond Fisc",
    "Cód. Neto", "Neto Gravado", "Alíc.", "IVA Liquidado", "IVA Débito",
    "Cód. NG/EX", "Conceptos NG/EX",
    "Cód. P/R", "Perc./Ret.", "Pcia P/R",
    "Total",
]

# IMPORTANTE: se elimina la columna "Tipo" (la que te quedaba como 3ra/4ta en Gastos)
COMPRAS_COLUMNS = [
    "Fecha Emisión ", "Fecha Recepción", "Cpbte", "Suc.", "Número",
    "Razón Social/Denominación Proveedor",
    "Tipo Doc.", "CUIT", "Domicilio", "C.P.", "Pcia", "Cond Fisc",
    "Cód. Neto", "Neto Gravado", "Alíc.", "IVA Liquidado", "IVA Crédito",
    "Cód. NG/EX", "Conceptos NG/EX",
    "Cód. P/R", "Perc./Ret.", "Pcia P/R",
    "Total",
]


def _to_int_cuit(cuit: str) -> Any:
    d = "".join(ch for ch in (cuit or "") if ch.isdigit())
    if not d:
        return ""
    try:
        return int(d)
    except Exception:
        return d


def build_ventas_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        cliente = l.acopio

        rows.append({
            "Fecha dd/mm/aaaa": l.fecha,
            "Cpbte": l.tipo_cbte,          # F1/F2
            "Tipo": l.letra,              # A
            "Suc.": l.pv,
            "Número": l.numero,
            "Razón Social o Denominación Cliente ": (cliente.razon_social or "").strip(),
            "Tipo Doc.": 80,
            "CUIT": _to_int_cuit(cliente.cuit),
            "Domicilio": (cliente.domicilio or "").strip(),
            "C.P.": "",
            "Pcia": "",
            "Cond Fisc": cliente.cond_fisc,
            "Cód. Neto": l.cod_neto_venta,
            "Neto Gravado": float(l.neto or 0.0),
            "Alíc.": float(l.alic_iva or 0.0),
            "IVA Liquidado": float(l.iva or 0.0),
            "IVA Débito": float(l.iva or 0.0),
            "Cód. NG/EX": "",
            "Conceptos NG/EX": "",
            "Cód. P/R": "",
            "Perc./Ret.": "",
            "Pcia P/R": "",
            "Total": float(l.total or 0.0),
        })

        def add_pr(code: str, amount: float):
            if amount is None:
                return
            amt = float(amount or 0.0)
            if abs(amt) < 1e-9:
                return
            rows.append({
                "Fecha dd/mm/aaaa": l.fecha,
                "Cpbte": "RV",
                "Tipo": l.letra,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social o Denominación Cliente ": (cliente.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": _to_int_cuit(cliente.cuit),
                "Domicilio": (cliente.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": cliente.cond_fisc,
                "Cód. Neto": "",
                "Neto Gravado": "",
                "Alíc.": "",
                "IVA Liquidado": "",
                "IVA Débito": "",
                "Cód. NG/EX": "",
                "Conceptos NG/EX": "",
                "Cód. P/R": code,
                "Perc./Ret.": amt,
                "Pcia P/R": "",
                "Total": amt,
            })

        # Retenciones (montos)
        add_pr("RA07", l.ret_iva)
        add_pr("RA05", l.ret_gan)

        # Percepción IVA (si existiera) -> P007
        add_pr("P007", l.percep_iva)

    return pd.DataFrame(rows, columns=VENTAS_COLUMNS)


def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        # Comprobante SOLO pv-numero (ej: 3302-29912534)
        comprobante = f"{l.pv}-{l.numero}"

        # Si hay varios registros de Mercadería Entregada: varias líneas
        me_items = l.mercaderia_entregada or []
        if not me_items:
            me_items = [None]

        for it in me_items:
            rows.append({
                "FECHA": l.fecha,
                "COE": l.coe,
                "COMPROBANTE": comprobante,
                "ACOPIO/COMPRADOR": (l.acopio.razon_social or "").strip(),
                "CUIT COMPRADOR": _to_int_cuit(l.acopio.cuit),
                "TIPO DE GRANO": l.grano,
                "CAMPAÑA": l.campaña or "",
                "CANTIDAD DE KILOS": float(l.kilos or 0.0),
                "PRECIO": float_confirm(l.precio),
                "NETO": float_confirm(l.neto),
                "IVA": float_confirm(l.iva),
                "TOTAL": float_confirm(l.total),
                "ME - Nro comprobante": (it.nro_comprobante if it else ""),
                "ME - Grado": (it.grado if it else ""),
                "ME - Factor": (it.factor if it else ""),
                "ME - Contenido proteico": (it.contenido_proteico if it else ""),
                "ME - Peso (kg)": (it.peso_kg if it else ""),
                "ME - Procedencia": (it.procedencia if it else ""),
                "LOCALIDAD": l.localidad,
            })

    cols = [
        "FECHA", "COE", "COMPROBANTE",
        "ACOPIO/COMPRADOR", "CUIT COMPRADOR",
        "TIPO DE GRANO", "CAMPAÑA",
        "CANTIDAD DE KILOS", "PRECIO",
        "NETO", "IVA", "TOTAL",
        "ME - Nro comprobante", "ME - Grado", "ME - Factor",
        "ME - Contenido proteico", "ME - Peso (kg)", "ME - Procedencia",
        "LOCALIDAD",
    ]
    return pd.DataFrame(rows, columns=cols)


def float_confirm(v: Optional[float]) -> float:
    try:
        return float(v or 0.0)
    except Exception:
        return 0.0


def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    """
    Modelo compras (HWCpra1):
    - Proveedor = acopio (COMPRADOR)
    - Movimiento: 203 por defecto; si alícuota 21% => 202
    - Exento (alíc 0%) se carga en NG/EX con código 203
    - Percepción IVA (IVA RG 4310/2018) -> Cód. P/R = P007 y monto en Perc./Ret.
    """
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        prov = l.acopio

        exento_total = 0.0
        by_alic: Dict[float, List[float]] = {}  # alic -> [neto, iva]

        for d in (l.deducciones or []):
            alic = float(d.alic or 0.0)
            if abs(alic) < 1e-9:
                exento_total += float(d.total if d.total else d.neto or 0.0)
            else:
                by_alic.setdefault(alic, [0.0, 0.0])
                by_alic[alic][0] += float(d.neto or 0.0)
                by_alic[alic][1] += float(d.iva or 0.0)

        alics_sorted = sorted(by_alic.keys())

        # Percepción IVA (si existe)
        percep_iva = float(l.percep_iva or 0.0)
        pr_code = "P007" if percep_iva > 0 else ""
        pr_amt = percep_iva if percep_iva > 0 else ""

        def emit_line(mov: int, alic: Any, neto: Any, iva: Any, exento_here: Any, put_pr: bool):
            total = float_confirm(neto) + float_confirm(iva) + float_confirm(exento_here)
            if put_pr and percep_iva > 0:
                total += percep_iva

            rows.append({
                "Fecha Emisión ": l.fecha,
                "Fecha Recepción": l.fecha,
                "Cpbte": mov,
                "Suc.": l.pv,
                "Número": l.numero,
                "Razón Social/Denominación Proveedor": (prov.razon_social or "").strip(),
                "Tipo Doc.": 80,
                "CUIT": _to_int_cuit(prov.cuit),
                "Domicilio": (prov.domicilio or "").strip(),
                "C.P.": "",
                "Pcia": "",
                "Cond Fisc": prov.cond_fisc,

                "Cód. Neto": mov,
                "Neto Gravado": neto,
                "Alíc.": alic,
                "IVA Liquidado": iva,
                "IVA Crédito": iva,

                "Cód. NG/EX": 203 if float_confirm(exento_here) > 0 else "",
                "Conceptos NG/EX": exento_here if float_confirm(exento_here) > 0 else "",

                "Cód. P/R": pr_code if put_pr else "",
                "Perc./Ret.": pr_amt if put_pr else "",
                "Pcia P/R": "",

                "Total": total,
            })

        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0
                mov = 202 if abs(alic - 21.0) < 0.001 else 203
                emit_line(
                    mov=mov,
                    alic=alic,
                    neto=neto,
                    iva=iva,
                    exento_here=exento_here,
                    put_pr=(idx == 0),  # percepción solo en primera línea
                )
        else:
            mov = 203
            emit_line(
                mov=mov,
                alic="",
                neto=0.0,
                iva=0.0,
                exento_here=exento_total if exento_total else 0.0,
                put_pr=True,  # si solo hay exento y percepción, igual va
            )

    return pd.DataFrame(rows, columns=COMPRAS_COLUMNS)


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    """
    Exporta a XLSX aplicando formatos:
    - Montos: #.##0,00
    - Alícuotas: 0,000
    - Moneda: "$"#.##0,00 (según columna)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    bold = Font(bold=True)
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = bold
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"

    # data
    for row in df.itertuples(index=False):
        ws.append(list(row))

    # Map formatos por nombre de columna
    formats: Dict[str, str] = {}

    # Generales
    for name in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "IVA Crédito", "Conceptos NG/EX", "Perc./Ret.", "Total",
                 "CANTIDAD DE KILOS", "NETO", "IVA", "TOTAL"]:
        if name in df.columns:
            formats[name] = FMT_AMOUNT

    if "Alíc." in df.columns:
        formats["Alíc."] = FMT_ALIQ

    if "PRECIO" in df.columns:
        formats["PRECIO"] = FMT_CURR

    # en preview usan Precio/Kg; en CPNs export usamos PRECIO
    if "CUIT" in df.columns:
        formats["CUIT"] = FMT_CUIT
    if "CUIT COMPRADOR" in df.columns:
        formats["CUIT COMPRADOR"] = FMT_CUIT

    # Aplicar formatos
    col_index = {name: i + 1 for i, name in enumerate(df.columns)}
    for col_name, fmt in formats.items():
        j = col_index[col_name]
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=j).number_format = fmt

    # Auto ancho simple
    for j, name in enumerate(df.columns, start=1):
        max_len = len(str(name))
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=j).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(j)].width = min(max_len + 2, 55)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
