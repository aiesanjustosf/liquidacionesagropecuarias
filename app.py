# ia_liquidaciones_granos.py
# Conversor / parser de Liquidaciones de Granos (PDF) -> Excel (Holistor / template propio)
# Incluye detección de NC (AJUSTE CRÉDITO + AJUSTE UNIFICADO) => TipoComprobante F2 + importes en negativo
#
# Reqs:
#   pip install streamlit pandas pdfplumber pillow openpyxl
# Opcional (fallback OCR si el PDF viene como imagen):
#   pip install pytesseract
#   y tener Tesseract instalado en el sistema.

import re
from io import BytesIO
from datetime import datetime
from typing import Optional, Dict, Any, List

import pandas as pd
import streamlit as st

try:
    import pdfplumber
except Exception as e:
    raise RuntimeError("Falta pdfplumber. Instalá: pip install pdfplumber") from e


# -----------------------------
# Helpers de texto / números
# -----------------------------
def normalize_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", (s or "").replace("\u00a0", " ")).strip()


def parse_num_ar(value: str) -> Optional[float]:
    """
    Convierte números tipo AR: 1.234.567,89  o  1234,56  o  1,234,567.89 (raro)
    a float. Devuelve None si no puede.
    """
    if value is None:
        return None
    s = normalize_spaces(value)
    s = s.replace("$", "").replace("ARS", "").replace("USD", "").strip()
    # quitar cualquier cosa que no sea dígito, separadores o signo
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if not s:
        return None

    # Heurística:
    # - Si hay coma y punto, asumimos miles con punto y decimales con coma (formato AR clásico)
    # - Si solo hay coma, asumimos coma decimal
    # - Si solo hay punto, asumimos punto decimal
    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s and "." not in s:
            s = s.replace(",", ".")
        # else: queda con punto decimal o entero
        return float(s)
    except Exception:
        return None


def find_first(pattern: str, text: str, flags=re.IGNORECASE | re.MULTILINE) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


# -----------------------------
# Extracción de texto del PDF
# -----------------------------
def extract_pdf_text(file_bytes: bytes) -> str:
    """
    Extrae texto con pdfplumber. Si el PDF es "imagen" y no trae texto,
    intenta OCR por página (opcional, si está pytesseract).
    """
    text_parts: List[str] = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            t = t.strip()
            if t:
                text_parts.append(t)
            else:
                # Fallback OCR (opcional)
                try:
                    import pytesseract  # noqa
                    from PIL import Image  # noqa

                    img = page.to_image(resolution=200).original
                    ocr = pytesseract.image_to_string(img, lang="eng")  # OCR lib soporta eng en este entorno
                    if ocr and ocr.strip():
                        text_parts.append(ocr.strip())
                except Exception:
                    # Si no hay pytesseract o falla OCR, seguimos
                    pass

    return "\n".join(text_parts)


# -----------------------------
# Detección NC y negativización
# -----------------------------
def detectar_nc_liquidacion(texto: str) -> bool:
    t = (texto or "").upper()
    return ("CONDICIONES DE LA OPERACIÓN - AJUSTE CRÉDITO" in t) and ("AJUSTE UNIFICADO" in t)


def negativizar_importes(registro: Dict[str, Any], campos_monetarios: List[str]) -> None:
    """
    Pone en negativo SOLO importes monetarios. Usa -abs() para evitar doble negativo.
    Modifica registro in-place.
    """
    for k in campos_monetarios:
        if k in registro and registro[k] is not None:
            try:
                registro[k] = -abs(float(registro[k]))
            except Exception:
                pass


# -----------------------------
# Parser (adaptable)
# -----------------------------
def parse_liquidacion(texto: str) -> Dict[str, Any]:
    """
    Parser generalista por regex. Ajustá patrones según el layout real del PDF.
    Te dejo un set base + placeholders.
    """
    t = normalize_spaces(texto)

    # Señales de NC
    es_nc = detectar_nc_liquidacion(texto)

    # Campos comunes (ejemplos). Ajustá a tu formato real:
    # Número de comprobante / liquidación:
    nro = (
        find_first(r"(?:NRO\.?|N°|NUMERO)\s*(?:DE\s*)?(?:LIQUIDACI[ÓO]N|COMPROBANTE)\s*[:\-]?\s*([A-Z0-9\-\/]+)", t)
        or find_first(r"(?:LIQUIDACI[ÓO]N)\s*[:\-]?\s*([A-Z0-9\-\/]+)", t)
    )

    # Fecha (varios formatos típicos dd/mm/aaaa)
    fecha_str = (
        find_first(r"(?:FECHA)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", t)
        or find_first(r"(\d{2}/\d{2}/\d{4})", t)
    )
    fecha = None
    if fecha_str:
        try:
            fecha = datetime.strptime(fecha_str, "%d/%m/%Y").date()
        except Exception:
            fecha = None

    # CUIT (si aparece)
    cuit = find_first(r"(?:CUIT)\s*[:\-]?\s*([0-9\-]{11,13})", t)

    # Razón social (placeholders: depende del PDF)
    productor = find_first(r"(?:PRODUCTOR|VENDEDOR)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ0-9 \.\-\/]+)", t)
    comprador = find_first(r"(?:COMPRADOR|DESTINATARIO)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ0-9 \.\-\/]+)", t)

    # Importes (adaptá estas etiquetas a las tuyas)
    # Ejemplo: "TOTAL A PAGAR", "TOTAL", "NETO", "IVA", "RETENCIONES", etc.
    total = parse_num_ar(find_first(r"(?:TOTAL\s*A\s*PAGAR|TOTAL)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")
    neto = parse_num_ar(find_first(r"(?:NETO\s*(?:GRAVADO)?)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")
    iva = parse_num_ar(find_first(r"(?:IVA)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")
    retenciones = parse_num_ar(find_first(r"(?:RETENCIONES?)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")
    deducciones = parse_num_ar(find_first(r"(?:DEDUCCIONES?)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")
    otros = parse_num_ar(find_first(r"(?:OTROS\s*(?:CONCEPTOS?)?)\s*[:\-]?\s*\$?\s*([0-9\.\,]+)", t) or "")

    # Armado del registro base
    registro: Dict[str, Any] = {
        "TipoComprobante": "F1",  # default (AJUSTÁ a tu lógica normal)
        "EsNC": es_nc,
        "Numero": nro,
        "Fecha": fecha_str,
        "CUIT": cuit,
        "Productor": productor,
        "Comprador": comprador,
        # Importes
        "Neto": neto,
        "IVA": iva,
        "Retenciones": retenciones,
        "Deducciones": deducciones,
        "Otros": otros,
        "Total": total,
    }

    if es_nc:
        registro["TipoComprobante"] = "F2"

        campos_monetarios = [
            "Neto", "IVA", "Retenciones", "Deducciones", "Otros", "Total",
            # Si tenés más columnas monetarias (Comisión, Flete, Gastos, Sellos, etc.) agregalas acá
        ]
        negativizar_importes(registro, campos_monetarios)

    return registro


# -----------------------------
# Export Excel
# -----------------------------
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Liquidaciones") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        ws.freeze_panes = "A2"
    return out.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Liquidaciones de Granos → Excel", layout="wide")

st.title("Liquidaciones de Granos (PDF) → Excel")
st.caption("Detección automática de NC: AJUSTE UNIFICADO + CONDICIONES DE LA OPERACIÓN - AJUSTE CRÉDITO ⇒ F2 e importes en negativo.")

files = st.file_uploader("Subí una o varias liquidaciones (PDF)", type=["pdf"], accept_multiple_files=True)

if files:
    registros = []
    errores = []

    for f in files:
        try:
            b = f.read()
            texto = extract_pdf_text(b)
            if not texto.strip():
                raise ValueError("No se pudo extraer texto del PDF (posible PDF escaneado sin OCR).")

            reg = parse_liquidacion(texto)
            reg["Archivo"] = f.name
            registros.append(reg)
        except Exception as e:
            errores.append({"Archivo": f.name, "Error": str(e)})

    if registros:
        df = pd.DataFrame(registros)

        st.subheader("Resultado")
        st.dataframe(df, use_container_width=True)

        excel_bytes = to_excel_bytes(df)
        st.download_button(
            "Descargar Excel",
            data=excel_bytes,
            file_name="liquidaciones_granos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Resumen rápido
        st.markdown(
            f"- Total procesadas: **{len(registros)}**  \n"
            f"- NC detectadas (F2): **{int(df['EsNC'].sum()) if 'EsNC' in df.columns else 0}**"
        )

    if errores:
        st.subheader("Errores")
        st.dataframe(pd.DataFrame(errores), use_container_width=True)

else:
    st.info("Subí PDFs para procesarlos. Si vienen escaneados, instalá OCR (pytesseract + tesseract).")
