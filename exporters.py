# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import List, Dict, Any, Tuple
import pandas as pd

from parser import Liquidacion, parse_number


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
        # Main sale row
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

        # Retentions: IVA and Ganancias only, each in its own line
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

    df = pd.DataFrame(rows, columns=VENTAS_COLUMNS)
    return df

def build_cpns_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    rows = []
    for l in liqs:
        comprobante = f"{l.tipo_cbte}-{l.letra}-{l.pv}-{l.numero}"
        rows.append({
            "FECHA": l.fecha,
            "COMPROBANTE": comprobante,
            "ACOPIO": (l.acopio.razon_social or "").strip(),
            "TIPO DE GRANO": l.grano,
            "CAMPAÑA": l.campaña or "",
            "CANTIDAD DE KILOS": l.kilos,
            "PRECIO": l.precio,
            "LOCALIDAD": l.localidad,
            "ME - Nro comprobante": l.me_nro_comprobante,
            "ME - Grado": l.me_grado,
            "ME - Factor": l.me_factor if l.me_factor is not None else "",
            "ME - Contenido proteico": l.me_contenido_proteico if l.me_contenido_proteico is not None else "",
            "ME - Procedencia": l.me_procedencia,
            "ME - Peso (kg)": l.me_peso_kg if l.me_peso_kg is not None else "",
        })
    return pd.DataFrame(rows)

def build_gastos_rows(liqs: List[Liquidacion]) -> pd.DataFrame:
    """
    Modelo compras (HWCpra1):
    - Proveedor = acopio (encabezado)
    - Tipo de movimiento: 203 por defecto; si IVA 21% => 202
    - Exento (alíc 0%) puede ir en la misma línea.
    - Si hay dos alícuotas (10.5 y 21) => líneas separadas.
    """
    rows: List[Dict[str, Any]] = []

    for l in liqs:
        # Aggregate deductions by aliquot
        exento_total = 0.0
        by_alic = {}  # alic -> (neto, iva)
        for d in l.deducciones:
            if (d.alic or 0) == 0:
                # treat as exento amount: total (or neto)
                exento_total += (d.total if d.total else d.neto)
            else:
                neto = d.neto
                iva = d.iva
                by_alic.setdefault(d.alic, [0.0, 0.0])
                by_alic[d.alic][0] += neto
                by_alic[d.alic][1] += iva

        # Decide lines
        alics_sorted = sorted(by_alic.keys())
        if alics_sorted:
            for idx, alic in enumerate(alics_sorted):
                neto, iva = by_alic[alic]
                exento_here = exento_total if idx == 0 else 0.0  # attach exento to first line
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
            # Only exento
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

    df = pd.DataFrame(rows, columns=COMPRAS_COLUMNS)
    return df

def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()
