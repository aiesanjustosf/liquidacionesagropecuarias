# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
import streamlit as st
from PIL import Image

from parser import parse_liquidacion_pdf, Liquidacion
from exporters import build_ventas_rows, build_cpns_rows, build_gastos_rows, df_to_xlsx_bytes

APP_TITLE = "IA liquidaciones agropecuarias"
FOOTER = "Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete"
PRIVACY_NOTE = "La app no almacena datos, toda la información está protegida."

LOGO_PATH = "assets/logo_aie.png"

# ---------------- Page config ----------------
logo_img = Image.open(LOGO_PATH)
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=logo_img,
    layout="wide",
)

# ---------------- Header ----------------
col1, col2 = st.columns([1, 7])
with col1:
    st.image(logo_img, use_container_width=True)
with col2:
    st.title(APP_TITLE)

st.caption(PRIVACY_NOTE)

st.divider()

# ---------------- Uploader ----------------
uploaded_files = st.file_uploader(
    "Subí una o varias liquidaciones (PDF)",
    type=["pdf"],
    accept_multiple_files=True,
)

liqs: list[Liquidacion] = []
errors: list[str] = []

if uploaded_files:
    for uf in uploaded_files:
        try:
            liq = parse_liquidacion_pdf(uf.read(), uf.name)
            liqs.append(liq)
        except Exception as e:
            errors.append(f"{uf.name}: {e}")

# ---------------- Preview grid ----------------
st.subheader("Vista previa (datos detectados)")

if errors:
    st.error("Algunos archivos no pudieron procesarse:")
    for e in errors:
        st.write(f"- {e}")

if not liqs:
    st.info("Subí PDFs para ver la vista previa y habilitar las descargas.")
else:
    preview_rows = []
    for l in liqs:
        preview_rows.append({
            "Archivo": l.filename,
            "Fecha": l.fecha,
            "Localidad": l.localidad,
            "Tipo": l.tipo_cbte,
            "C.O.E.": l.coe,
            "PV": l.pv,
            "Número": l.numero,
            "Comprador (acopio)": l.comprador.razon_social,
            "CUIT comprador": l.comprador.cuit,
            "Grano": l.grano,
            "Kilos": l.kilos,
            "Precio": l.precio,
            "Neto": l.neto,
            "IVA": l.iva,
            "Total": l.total,
            "Ret IVA": l.ret_iva,
            "Ret Gan": l.ret_gan,
            "Campaña": l.campaña,
            "ME Nro": l.me_nro_comprobante,
            "ME Procedencia": l.me_procedencia,
            "ME Peso (kg)": l.me_peso_kg,
        })
    preview_df = pd.DataFrame(preview_rows)
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Descargas")

    ventas_df = build_ventas_rows(liqs)
    cpns_df = build_cpns_rows(liqs)
    gastos_df = build_gastos_rows(liqs)

    colA, colB, colC = st.columns(3)

    with colA:
        st.write("**Excel Ventas** (modelo HWVta1modelo)")
        st.download_button(
            label="Descargar Ventas.xlsx",
            data=df_to_xlsx_bytes(ventas_df, "Ventas"),
            file_name="VENTAS_liquidaciones.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(f"Reglas: F1/F2 según tipo de liquidación; retenciones IVA/Ganancias como RV (RA07/RA05).")

    with colB:
        st.write("**Excel CPNs**")
        st.download_button(
            label="Descargar CPNs.xlsx",
            data=df_to_xlsx_bytes(cpns_df, "CPNs"),
            file_name="CPNs_liquidaciones.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Incluye Mercadería Entregada; Campaña es opcional (solo si el PDF la trae).")

    with colC:
        st.write("**Excel Gastos** (modelo compras HWCpra1)")
        st.download_button(
            label="Descargar Gastos.xlsx",
            data=df_to_xlsx_bytes(gastos_df, "Gastos"),
            file_name="GASTOS_liquidaciones.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Proveedor = acopio; movimiento 203 (o 202 si IVA 21%). Exento puede ir en la misma línea.")

# ---------------- Footer ----------------
st.divider()
st.caption(FOOTER)
