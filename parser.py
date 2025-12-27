# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
import unicodedata
from typing import Dict, List, Optional, Tuple

import pdfplumber


# ------------------------- Helpers -------------------------

def _norm(s: str) -> str:
    """Uppercase, remove accents, normalize spaces."""
    if s is None:
        return ""
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = re.sub(r"\s+", " ", s)
    return s

def parse_number(raw: str) -> Optional[float]:
    """
    Parse numbers that may be:
    - 2,585.00 (US)
    - 27,14 (EU)
    - 2585000.00
    """
    if raw is None:
        return None
    s = str(raw)
    s = s.strip()
    if not s:
        return None
    # Keep minus, digits, separators
    s = re.sub(r"[^0-9\-,.]", "", s)
    if not s or s in {".", ",", "-", "-.", "-,"}:
        return None

    # If both separators exist, decide by last separator
    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_dot > last_comma:
            # dot is decimal, commas thousands
            s = s.replace(",", "")
        else:
            # comma is decimal, dots thousands
            s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        # comma decimal if ends with ,dd
        if re.search(r",\d{1,3}$", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        # dot only or none
        s = s.replace(",", "")

    try:
        return float(s)
    except Exception:
        return None

def parse_cuit_digits(raw: str) -> str:
    if raw is None:
        return ""
    d = re.sub(r"\D", "", str(raw))
    return d

# ------------------------- Domain -------------------------

@dataclass
class Party:
    razon_social: str = ""
    domicilio: str = ""
    localidad: str = ""
    cuit: str = ""
    iva: str = ""

    @property
    def cond_fisc(self) -> str:
        v = _norm(self.iva)
        if "RI" in v or "RESP" in v:
            return "RI"
        if "EX" in v:
            return "EX"
        if "CF" in v or "CONSUMIDOR" in v:
            return "CF"
        return self.iva.strip() if self.iva else ""

@dataclass
class DeductionLine:
    concepto: str
    neto: float
    alic: float
    iva: float
    total: float

@dataclass
class Liquidacion:
    filename: str
    fecha: str
    localidad: str
    tipo_cbte: str   # F1/F2
    letra: str       # A
    coe: str
    pv: str
    numero: str
    comprador: Party   # acopio
    vendedor: Party
    grano: str
    cod_neto_venta: str
    kilos: float
    precio: float
    neto: float
    alic_iva: float
    iva: float
    total: float
    campaña: str
    me_nro_comprobante: str
    me_procedencia: str
    me_peso_kg: Optional[float]
    ret_iva: float
    ret_gan: float
    deducciones: List[DeductionLine]

# ------------------------- Parsing -------------------------

GRAIN_CODES = {
    "SOJA": "123",
    "MAIZ": "124",
    "TRIGO": "161",
    "GIRASOL": "157",
    "ARVEJA": "120",
    "SORGO": "151",
    "CAMELINA SATIVA": "162",
}

def _detect_tipo_cbte(full_text_norm: str) -> str:
    if "LIQUIDACION SECUNDARIA" in full_text_norm:
        return "F2"
    return "F1"

def _extract_header_date_loc(page_text: str) -> Tuple[str, str]:
    # Usually: "20/11/2025, VIDELA" or "20/11/2025 - VIDELA"
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s*[,–\-]\s*([^\n]+)", page_text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # fallback: first date anywhere
    m2 = re.search(r"(\d{2}/\d{2}/\d{4})", page_text)
    return (m2.group(1) if m2 else ""), ""

def _slice_parties_block(page_text: str) -> str:
    # Between "COMPRADOR" and "CONDICIONES" or "Actuó Corredor"
    start = page_text.find("COMPRADOR")
    if start == -1:
        start = 0
    end_candidates = [page_text.find("CONDICIONES"), page_text.find("Actuó Corredor"), page_text.find("ACTUÓ CORREDOR")]
    end_candidates = [e for e in end_candidates if e != -1]
    end = min(end_candidates) if end_candidates else len(page_text)
    return page_text[start:end]

def _extract_two_values(lines: List[str], label: str) -> List[str]:
    out: List[str] = []
    for ln in lines:
        if ln.strip().startswith(label):
            out.append(ln.split(label, 1)[1].strip())
    return out

def _extract_parties(page_text: str, header_localidad: str) -> Tuple[Party, Party]:
    """
    pdfplumber's extract_text() may interleave columns inconsistently.
    Empirically, in many LP/LS PDFs:
      - "Razón Social:" values tend to appear in COMPRADOR then VENDEDOR order.
      - "Domicilio/Localidad/CUIT/IVA" lines may appear in the opposite order.
    Strategy:
      1) Extract both razones (as roles: comprador=v1, vendedor=v2).
      2) Extract other fields as two values (as indexes 0/1).
      3) Choose which index belongs to the comprador by matching the header locality (Fecha y Localidad).
    """
    block = _slice_parties_block(page_text)
    lines = [l.rstrip() for l in block.splitlines() if l.strip()]

    # --- Razones (role order) ---
    razones: List[str] = []
    for ln in lines:
        if ln.strip().startswith("Razón Social:") or ln.strip().startswith("Razon Social:"):
            val = ln.split(":", 1)[1].strip()
            razones.append(val)

    # Try to attach wrapped continuation: if we have exactly 2 razones, and there's a standalone line
    # that starts with "DE " or "DEL " right after the reasons block, attach to the first (common in cooperativas).
    razones_clean: List[str] = []
    for r in razones[:2]:
        razones_clean.append(r)

    for ln in lines:
        if ":" not in ln and ln.strip() and (ln.strip().upper().startswith("DE ") or ln.strip().upper().startswith("DEL ")):
            if len(razones_clean) >= 1 and " DE " not in _norm(razones_clean[0]):
                razones_clean[0] = (razones_clean[0] + " " + ln.strip()).strip()
            break

    def pad2(arr: List[str]) -> List[str]:
        arr = arr[:2]
        while len(arr) < 2:
            arr.append("")
        return arr

    razones_clean = pad2(razones_clean)

    domicilios = pad2(_extract_two_values(lines, "Domicilio:"))
    localidades = pad2(_extract_two_values(lines, "Localidad:"))
    cuits = pad2(_extract_two_values(lines, "C.U.I.T.:") or _extract_two_values(lines, "C.U.I.T:"))
    ivas = pad2(_extract_two_values(lines, "I.V.A.:") or _extract_two_values(lines, "I.V.A:"))

    # Determine which index (0/1) is the comprador by matching header locality
    hloc = _norm(header_localidad)
    idx_comp = 0
    if hloc:
        if _norm(localidades[0]) == hloc:
            idx_comp = 0
        elif _norm(localidades[1]) == hloc:
            idx_comp = 1

    idx_vend = 1 - idx_comp

    comprador = Party(
        razon_social=razones_clean[0],
        domicilio=domicilios[idx_comp],
        localidad=localidades[idx_comp],
        cuit=parse_cuit_digits(cuits[idx_comp]),
        iva=ivas[idx_comp],
    )
    vendedor = Party(
        razon_social=razones_clean[1],
        domicilio=domicilios[idx_vend],
        localidad=localidades[idx_vend],
        cuit=parse_cuit_digits(cuits[idx_vend]),
        iva=ivas[idx_vend],
    )

    return comprador, vendedor

def _extract_grain(page_text: str) -> Tuple[str, str]:
    # Look near conditions line or any "- MAIZ" occurrences.
    m = re.search(r"\b(Soja|Ma[ií]z|Trigo|Girasol|Arveja|Sorgo|Camelina\s*Sativa)\b", page_text, flags=re.IGNORECASE)
    if not m:
        return "", ""
    grain_raw = m.group(1)
    gnorm = _norm(grain_raw).replace("Í", "I")
    # normalize variants
    if "MAIZ" in gnorm:
        gname = "MAIZ"
    elif "SOJA" in gnorm:
        gname = "SOJA"
    elif "TRIGO" in gnorm:
        gname = "TRIGO"
    elif "GIRASOL" in gnorm:
        gname = "GIRASOL"
    elif "ARVEJA" in gnorm:
        gname = "ARVEJA"
    elif "SORGO" in gnorm:
        gname = "SORGO"
    elif "CAMELINA" in gnorm:
        gname = "CAMELINA SATIVA"
    else:
        gname = gnorm
    return gname.title() if gname != "MAIZ" else "Maíz", GRAIN_CODES.get(gname, "")

def _extract_operation_numbers(page_text: str) -> Tuple[float, float, float, float, float]:
    """
    Returns: kilos, precio, neto, alic_iva, iva, total
    """
    # Operation line: "10000 Kg $258.50 $2585000.00 10.5 $271425.00 $2856425.00"
    m = re.search(
        r"\n\s*([0-9][0-9.,]*)\s*Kg\s*\$?\s*([0-9][0-9.,]*)\s*\$?\s*([0-9][0-9.,]*)\s*([0-9][0-9.,]*)\s*\$?\s*([0-9][0-9.,]*)\s*\$?\s*([0-9][0-9.,]*)",
        page_text,
        flags=re.IGNORECASE,
    )
    if not m:
        return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    kilos = parse_number(m.group(1)) or 0.0
    precio = parse_number(m.group(2)) or 0.0
    neto = parse_number(m.group(3)) or 0.0
    alic = parse_number(m.group(4)) or 0.0
    iva = parse_number(m.group(5)) or 0.0
    total = parse_number(m.group(6)) or 0.0
    return kilos, precio, neto, alic, iva, total

def _extract_campaign(page_text: str) -> str:
    m = re.search(r"Campaña\s*[:\-]\s*([^\n]+)", page_text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""

def _extract_me(page_text: str) -> Tuple[str, str, Optional[float]]:
    # Isolate between MERCADERIA ENTREGADA and OPERACIÓN
    up = page_text.upper()
    start = up.find("MERCADERIA ENTREGADA")
    if start == -1:
        start = up.find("MERCADERÍA ENTREGADA")
    end = up.find("OPERACIÓN", start)
    if end == -1:
        end = up.find("OPERACION", start)
    if start == -1 or end == -1 or end <= start:
        return "", "", None
    sec = page_text[start:end]

    # Nro comprobante (first long integer)
    nro = ""
    m = re.search(r"\b(\d{10,14})\b", sec)
    if m:
        nro = m.group(1).strip()

    # Procedencia: capture between "Localidad:" and the comprobante number line
    proced = ""
    mproc = re.search(r"Localidad:\s*(.*?)\n\s*(\d{10,14})\b", sec, flags=re.IGNORECASE | re.DOTALL)
    if mproc:
        proced_raw = mproc.group(1).replace("\n", " ").strip()
        proced = re.sub(r"\s+", " ", proced_raw)

        # Sometimes province abbreviation sits on a standalone short line immediately after the comprobante row
        sec_lines = [l.strip() for l in sec.splitlines() if l.strip()]
        for i, l in enumerate(sec_lines):
            if nro and nro in l:
                if i + 1 < len(sec_lines):
                    nxt = sec_lines[i + 1].strip()
                    if re.fullmatch(r"[A-Za-z]{2,3}", nxt):
                        proced = (proced + " " + nxt).strip()
                break

    # Peso (kg): take the last numeric token of the comprobante row
    peso = None
    if nro:
        for line in sec.splitlines():
            if nro in line:
                toks = re.findall(r"[-]?\d[\d.,]*", line)
                if toks:
                    peso = parse_number(toks[-1])
                break

    return nro, proced, peso

def _extract_retenciones(page_text: str) -> Tuple[float, float]:
    ret_iva = 0.0
    ret_gan = 0.0
    # IVA
    for pat in [r"RET\s*IVA[^\n]*\$\s*([0-9][0-9.,]*)", r"I\.V\.A\.[^\n]*RET\s*IVA[^\n]*\$\s*([0-9][0-9.,]*)"]:
        m = re.search(pat, page_text, flags=re.IGNORECASE)
        if m:
            ret_iva = parse_number(m.group(1)) or 0.0
            break
    # GAN
    m = re.search(r"RET\s*GAN[^\n]*\$\s*([0-9][0-9.,]*)", page_text, flags=re.IGNORECASE)
    if m:
        ret_gan = parse_number(m.group(1)) or 0.0
    return ret_iva, ret_gan

def _extract_deducciones(page_text: str) -> List[DeductionLine]:
    # Between DEDUCCIONES and RETENCIONES
    up = page_text.upper()
    s = up.find("DEDUCCIONES")
    e = up.find("RETENCIONES")
    if s == -1 or e == -1 or e <= s:
        return []
    sec = page_text[s:e]
    lines = [l.strip() for l in sec.splitlines() if l.strip()]
    out: List[DeductionLine] = []

    # Heuristic parsing:
    # - Lines with percentages have: <CONCEPTO> ... <neto> <alic%> <iva> <total>
    # - 0% lines: <CONCEPTO> ... <total> 0% <iva> <total>
    for ln in lines:
        # skip headers
        if ln.lower().startswith("concepto") or "base cálculo" in ln.lower():
            continue

        # normalized spaces
        ln2 = re.sub(r"\s+", " ", ln)

        # pattern with alic% and iva and total at end
        m = re.search(r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+([0-9][0-9.,]*)%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$", ln2)
        if m:
            concepto = m.group(1).strip()
            neto = parse_number(m.group(2)) or 0.0
            alic = parse_number(m.group(3)) or 0.0
            iva = parse_number(m.group(4)) or 0.0
            total = parse_number(m.group(5)) or 0.0
            out.append(DeductionLine(concepto=concepto, neto=neto, alic=alic, iva=iva, total=total))
            continue

        # pattern 0% where neto not explicit
        m0 = re.search(r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+0%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$", ln2)
        if m0:
            concepto = m0.group(1).strip()
            total = parse_number(m0.group(2)) or 0.0
            iva = parse_number(m0.group(3)) or 0.0
            total2 = parse_number(m0.group(4)) or total
            out.append(DeductionLine(concepto=concepto, neto=total, alic=0.0, iva=iva, total=total2))
            continue

    # de-duplicate obvious header artifacts
    cleaned: List[DeductionLine] = []
    for d in out:
        if _norm(d.concepto) in {"COMISION O GASTOS", "ADMINISTRATIVOS", "OTRAS DEDUCCIONES"}:
            continue
        cleaned.append(d)
    return cleaned

def parse_liquidacion_pdf(pdf_bytes: bytes, filename: str) -> Liquidacion:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        page_text = page.extract_text() or ""

    full_norm = _norm(page_text)
    fecha, localidad = _extract_header_date_loc(page_text)
    tipo_cbte = _detect_tipo_cbte(full_norm)

    mcoe = re.search(r"C\.O\.E\.\s*:\s*([0-9]{8,})", page_text, flags=re.IGNORECASE)
    coe = mcoe.group(1).strip() if mcoe else ""
    pv = coe[:4] if len(coe) >= 4 else ""
    numero = coe[4:12] if len(coe) >= 12 else (coe[4:] if len(coe) > 4 else "")

    comprador, vendedor = _extract_parties(page_text, localidad)

    grano, cod_neto_venta = _extract_grain(page_text)
    kilos, precio, neto, alic_iva, iva, total = _extract_operation_numbers(page_text)

    campaña = _extract_campaign(page_text)
    me_nro, me_proced, me_peso = _extract_me(page_text)

    ret_iva, ret_gan = _extract_retenciones(page_text)
    deducciones = _extract_deducciones(page_text)

    return Liquidacion(
        filename=filename,
        fecha=fecha,
        localidad=localidad,
        tipo_cbte=tipo_cbte,
        letra="A",
        coe=coe,
        pv=pv,
        numero=numero,
        comprador=comprador,
        vendedor=vendedor,
        grano=grano,
        cod_neto_venta=cod_neto_venta,
        kilos=kilos,
        precio=precio,
        neto=neto,
        alic_iva=alic_iva,
        iva=iva,
        total=total,
        campaña=campaña,
        me_nro_comprobante=me_nro,
        me_procedencia=me_proced,
        me_peso_kg=me_peso,
        ret_iva=ret_iva,
        ret_gan=ret_gan,
        deducciones=deducciones,
    )
