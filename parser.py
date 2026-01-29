# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
import unicodedata
from typing import List, Optional, Tuple

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
    Supports negative.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None

    # Keep minus, digits, separators
    s = re.sub(r"[^0-9\-,.]", "", s)
    if not s or s in {".", ",", "-", "-.", "-,"}:
        return None

    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_dot > last_comma:
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        if re.search(r",\d{1,3}$", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", "")

    try:
        return float(s)
    except Exception:
        return None


def parse_cuit_digits(raw: str) -> str:
    if raw is None:
        return ""
    s = str(raw)

    m = re.search(r"\b(\d{2})\D?(\d{8})\D?(\d)\b", s)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"

    m2 = re.search(r"\b(\d{11})\b", s)
    if m2:
        return m2.group(1)

    d = re.sub(r"\D", "", s)
    return d[:11] if len(d) >= 11 else d


def _neg_abs(v: Optional[float]) -> float:
    """Forzar negativo en importes monetarios: -abs(x)."""
    try:
        return -abs(float(v or 0.0))
    except Exception:
        return 0.0


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
    acopio: Party
    comprador: Party
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
    me_grado: str
    me_factor: Optional[float]
    me_contenido_proteico: Optional[float]
    me_peso_kg: Optional[float]
    me_procedencia: str
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


def _is_nc(full_text_norm: str) -> bool:
    """
    NC cuando aparecen AMBAS:
      - CONDICIONES DE LA OPERACION - AJUSTE CREDITO (guion flexible)
      - AJUSTE UNIFICADO
    """
    pat = r"CONDICIONES DE LA OPERACION\s*[-–—]\s*AJUSTE CREDITO"
    return bool(re.search(pat, full_text_norm)) and ("AJUSTE UNIFICADO" in full_text_norm)


def _detect_tipo_cbte(full_text_norm: str) -> str:
    if "LIQUIDACION SECUNDARIA" in full_text_norm:
        return "F2"
    return "F1"


def _extract_header_date_loc(page_text: str) -> Tuple[str, str]:
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s*[,–\-]\s*([^\n]+)", page_text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m2 = re.search(r"(\d{2}/\d{2}/\d{4})", page_text)
    return (m2.group(1) if m2 else ""), ""


def _party_from_text(side_text: str) -> Party:
    if not side_text:
        return Party()
    txt = side_text.replace("\r", "\n")
    lines = [l.strip() for l in txt.split("\n") if l.strip()]

    def _take_multiline_value(label_regex: str, stop_regex: str) -> str:
        for i, ln in enumerate(lines):
            m = re.search(label_regex, ln, flags=re.IGNORECASE)
            if not m:
                continue
            val = (m.group(1) or "").strip()
            j = i + 1
            while j < len(lines):
                ln2 = lines[j]
                if re.search(stop_regex, ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bRaz[oó]n\s+Social\b\s*:", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bDomicilio\b\s*:", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bC\.U\.I\.T\b", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bI\.V\.A\b", ln2, flags=re.IGNORECASE):
                    break
                if ":" not in ln2:
                    val = (val + " " + ln2).strip() if val else ln2
                    j += 1
                    continue
                break
            return re.sub(r"\s+", " ", val).strip()
        return ""

    razon = _take_multiline_value(
        r"Raz[oó]n\s+Social\s*:\s*(.*)$",
        r"^(Domicilio|Localidad|C\.U\.I\.T|I\.V\.A|IIBB|Ingresos\s+Brutos)\b",
    )
    domicilio = _take_multiline_value(
        r"Domicilio\s*:\s*(.*)$",
        r"^(Localidad|C\.U\.I\.T|I\.V\.A|IIBB|Ingresos\s+Brutos)\b",
    )

    def _rgx_first(pat: str) -> str:
        m = re.search(pat, txt, flags=re.IGNORECASE)
        return m.group(1).strip() if m else ""

    localidad = _rgx_first(r"Localidad\s*:\s*([^\n]+)")
    cuit = parse_cuit_digits(_rgx_first(r"C\.U\.I\.T\.?\s*:?\s*([^\n]+)"))
    iva = _rgx_first(r"I\.V\.A\.?\s*:?\s*([^\n]+)")

    def _cut_at_labels(v: str) -> str:
        v2 = (v or "").strip()
        if not v2:
            return ""
        for lab in ["RAZON SOCIAL", "DOMICILIO", "C.U.I.T", "I.V.A", "LOCALIDAD"]:
            idx = _norm(v2).find(lab)
            if idx > 0:
                v2 = v2[:idx].strip()
        return re.sub(r"\s+", " ", v2).strip()

    razon = _cut_at_labels(razon)
    domicilio = _cut_at_labels(domicilio)

    return Party(razon_social=razon, domicilio=domicilio, localidad=localidad, cuit=cuit, iva=iva)


def _group_words_to_lines(words: List[dict]) -> List[str]:
    if not words:
        return []
    ws = sorted(words, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
    lines: List[List[str]] = []
    current: List[str] = []
    current_top = ws[0].get("top", 0)

    def flush():
        nonlocal current
        if current:
            lines.append(current)
            current = []

    for w in ws:
        top = w.get("top", 0)
        if abs(top - current_top) > 3:
            flush()
            current_top = top
        current.append(w.get("text", ""))
    flush()
    return [" ".join(l).strip() for l in lines if " ".join(l).strip()]


def _extract_parties_from_layout(page: pdfplumber.page.Page) -> Tuple[Party, Party, Party]:
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
    width = float(page.width)

    comprador_x0 = None
    vendedor_x0 = None
    for w in words:
        t = _norm(w.get("text", ""))
        if t == "COMPRADOR" and comprador_x0 is None:
            comprador_x0 = float(w.get("x0", 0))
        if t == "VENDEDOR" and vendedor_x0 is None:
            vendedor_x0 = float(w.get("x0", 0))

    if comprador_x0 is not None and vendedor_x0 is not None and vendedor_x0 > comprador_x0:
        x_split = (comprador_x0 + vendedor_x0) / 2.0
    else:
        x_split = width / 2.0

    def find_top(token: str) -> Optional[float]:
        tnorm = _norm(token)
        tops = [w["top"] for w in words if _norm(w.get("text", "")) == tnorm]
        return min(tops) if tops else None

    y_compr = find_top("COMPRADOR")
    y_actuo = find_top("ACTUÓ") or find_top("ACTUO")
    y_start = y_compr if y_compr is not None else 0.0
    y_end = y_actuo if y_actuo is not None else (y_start + 180.0)

    block_words = [w for w in words if (w.get("top", 0) >= y_start and w.get("top", 0) <= y_end)]
    left_words = [w for w in block_words if w.get("x0", 0) < x_split]
    right_words = [w for w in block_words if w.get("x0", 0) >= x_split]

    left_text = "\n".join(_group_words_to_lines(left_words))
    right_text = "\n".join(_group_words_to_lines(right_words))

    comprador = _party_from_text(left_text)
    vendedor = _party_from_text(right_text)

    acopio = Party(
        razon_social=comprador.razon_social,
        domicilio=comprador.domicilio,
        localidad=comprador.localidad,
        cuit=comprador.cuit,
        iva=comprador.iva,
    )

    return acopio, comprador, vendedor


def _extract_grain(page_text: str) -> Tuple[str, str]:
    m = re.search(r"\b(Soja|Ma[ií]z|Trigo|Girasol|Arveja|Sorgo|Camelina\s*Sativa)\b", page_text, flags=re.IGNORECASE)
    if not m:
        return "", ""
    grain_raw = m.group(1)
    gnorm = _norm(grain_raw).replace("Í", "I")

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


def _extract_operation_numbers_standard(page_text: str) -> Tuple[float, float, float, float, float, float]:
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


def _extract_operation_from_ajuste_credito(full_text: str) -> Optional[Tuple[float, float, float, float, float, float]]:
    """
    En Ajuste Unificado: buscar "CONDICIONES DE LA OPERACIÓN - AJUSTE CRÉDITO" y tomar
    de OPERACIÓN: Subtotal, Importe IVA, Operación c/IVA.
    No se calculan montos. Si falta alícuota en la fila, se fija en 10.5.
    """
    pat_start = r"CONDICIONES DE LA OPERACI[ÓO]N\s*[-–—]\s*AJUSTE CR[ÉE]DITO"
    m = re.search(pat_start, full_text, flags=re.IGNORECASE)
    if not m:
        return None
    start = m.start()

    m_end = re.search(r"\n\s*CONDICIONES DE LA OPERACI[ÓO]N\b", full_text[m.end():], flags=re.IGNORECASE)
    end = m.end() + m_end.start() if m_end else len(full_text)
    sec = full_text[start:end]

    lines = [re.sub(r"\s+", " ", l.strip()) for l in sec.splitlines() if l.strip()]
    op_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"\bOPERACI[ÓO]N\b", ln, flags=re.IGNORECASE):
            op_idx = i
            break
    if op_idx is None:
        return None

    data_line = None
    for j in range(op_idx + 1, min(len(lines), op_idx + 12)):
        if re.search(r"\bKg\b", lines[j], flags=re.IGNORECASE) and re.search(r"\d", lines[j]):
            data_line = lines[j]
            break
    if not data_line:
        return None

    toks = re.findall(r"[-]?\d[\d.,]*", data_line)
    vals = [parse_number(t) for t in toks]
    vals = [v for v in vals if v is not None]
    if len(vals) < 5:
        return None

    kilos = float(vals[0] or 0.0)
    precio = float(vals[1] or 0.0)
    subtotal = float(vals[2] or 0.0)

    if len(vals) >= 6:
        alic = float(vals[3] or 0.0)
        importe_iva = float(vals[4] or 0.0)
        total = float(vals[5] or 0.0)
    else:
        # Si no viene la alícuota en el PDF, se fija en 10,500 (sin calcular montos)
        alic = 10.5
        importe_iva = float(vals[3] or 0.0)
        total = float(vals[4] or 0.0)

    return kilos, precio, subtotal, alic, importe_iva, total


def _extract_campaign(page_text: str) -> str:
    m = re.search(r"Campaña\s*[:\-]\s*([^\n]+)", page_text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _extract_me(page_text: str) -> Tuple[str, str, Optional[float], Optional[float], Optional[float], str]:
    up = page_text.upper()
    start = up.find("MERCADERIA ENTREGADA")
    if start == -1:
        start = up.find("MERCADERÍA ENTREGADA")
    end = up.find("OPERACIÓN", start)
    if end == -1:
        end = up.find("OPERACION", start)
    if start == -1 or end == -1 or end <= start:
        return "", "", None, None, None, ""

    sec = page_text[start:end]

    nro = ""
    grado = ""
    factor = None
    prot = None
    peso = None

    mrow = re.search(
        r"\b(\d{10,14})\b\s+([A-Z0-9]{1,4})\s+([0-9][0-9.,]*)\s+([0-9][0-9.,]*)\s+([0-9][0-9.,]*)",
        sec,
        flags=re.IGNORECASE,
    )
    if mrow:
        nro = mrow.group(1).strip()
        grado = mrow.group(2).strip()
        factor = parse_number(mrow.group(3))
        prot = parse_number(mrow.group(4))
        peso = parse_number(mrow.group(5))
    else:
        m2 = re.search(r"\b(\d{10,14})\b", sec)
        if m2:
            nro = m2.group(1).strip()
            for line in sec.splitlines():
                if nro in line:
                    toks2 = re.findall(r"[-]?\d[\d.,]*", line)
                    if toks2:
                        peso = parse_number(toks2[-1])
                    break

    proced = ""
    mloc = re.search(r"Localidad\s*:\s*([^\n]+)", sec, flags=re.IGNORECASE)
    if mloc:
        proced = re.sub(r"\s+", " ", mloc.group(1).strip())

    return nro, grado, factor, prot, peso, proced


# ------------------------- RETENCIONES -------------------------

_STOP_KWS = {"IMPORTES", "TOTAL", "FIRMA", "OTROS", "GRAV", "LIQUID", "DEDUCC", "IMPUESTO"}

def _looks_like_continuation(line: str) -> bool:
    u = (line or "").upper().strip()
    if not u:
        return False
    if any(k in u for k in _STOP_KWS):
        return False
    if "REGIMEN" in u:
        return True
    if re.search(r"[A-ZÁÉÍÓÚÑ]", u):
        return False
    return bool(re.search(r"\d", u))


def _extract_retencion_iva_from_retenciones_table(full_text: str) -> float:
    up = full_text.upper()
    s = up.find("RETENCIONES")
    if s == -1:
        return 0.0

    sec = full_text[s:s + 1200]
    lines = sec.splitlines()

    m_total = re.search(r"Total\s+Retenciones\s+Afip\s*:\s*\$?\s*([-]?\d[\d.,]*)", sec, flags=re.IGNORECASE)
    total_afip = float(parse_number(m_total.group(1)) or 0.0) if m_total else None

    iva_idx = None
    for i, ln in enumerate(lines):
        u = ln.upper()
        if "%" in u and re.search(r"\bI\.?\s*V\.?\s*A\.?\b", ln, flags=re.IGNORECASE):
            if "IVA RG" in u:
                continue
            iva_idx = i
            break

    if iva_idx is None:
        return total_afip or 0.0

    block_lines = [lines[iva_idx].strip()]
    for j in range(iva_idx + 1, min(len(lines), iva_idx + 3)):
        ln = lines[j].strip()
        if _looks_like_continuation(ln):
            block_lines.append(ln)
        else:
            break
    block = " ".join(block_lines)

    dols = [parse_number(m.group(1)) for m in re.finditer(r"\$\s*([-]?\d[\d.,]*)", block)]
    dols = [v for v in dols if v is not None and abs(v) > 0.0001]
    if dols:
        return float(max(dols, key=lambda x: abs(x)))

    nums = [parse_number(n) for n in re.findall(r"[-]?\d[\d.,]*", block)]
    nums = [v for v in nums if v is not None and abs(v) > 0.0001]
    if nums:
        return float(min(nums, key=lambda x: abs(x)))

    return total_afip or 0.0


def _extract_retenciones(full_text: str) -> Tuple[float, float]:
    ret_iva = _extract_retencion_iva_from_retenciones_table(full_text)
    ret_gan = 0.0
    return ret_iva, ret_gan


# ------------------------- DEDUCCIONES -------------------------

def _extract_deducciones(page_text: str) -> List[DeductionLine]:
    up = page_text.upper()
    s = up.find("DEDUCCIONES")
    e = up.find("RETENCIONES")
    if s == -1 or e == -1 or e <= s:
        return []

    sec = page_text[s:e]
    raw_lines = [l.strip() for l in sec.splitlines() if l.strip()]

    out: List[DeductionLine] = []
    pending_detail = ""
    buffer = ""

    pat_full = re.compile(
        r"^(.*?)\s+\$?\s*([-]?\d[\d.,]*)\s+([-]?\d[\d.,]*)%?\s+\$?\s*([-]?\d[\d.,]*)\s+\$?\s*([-]?\d[\d.,]*)\s*$"
    )
    pat_zero = re.compile(
        r"^(.*?)\s+\$?\s*([-]?\d[\d.,]*)\s+0%?\s+\$?\s*([-]?\d[\d.,]*)\s+\$?\s*([-]?\d[\d.,]*)\s*$"
    )

    def looks_like_header(ln: str) -> bool:
        n = _norm(ln)
        if n.startswith("DEDUCCIONES"):
            return True
        if "CONCEPTO" in n and ("ALICUOTA" in n or "ALÍCUOTA" in n):
            return True
        if "BASE" in n and ("CALC" in n or "CÁLC" in n):
            return True
        if "NETO" in n and "IVA" in n and "TOTAL" in n:
            return True
        return False

    for ln in raw_lines:
        if looks_like_header(ln):
            continue

        ln2 = re.sub(r"\s+", " ", ln)

        if re.match(r"^\d{3,7}\|", ln2) and ("$" not in ln2) and (not pat_full.search(ln2)) and (not pat_zero.search(ln2)):
            pending_detail = ln2
            continue

        candidate = (buffer + " " + ln2).strip() if buffer else ln2

        m = pat_full.search(candidate)
        m0 = None if m else pat_zero.search(candidate)

        if not m and not m0:
            if len(candidate) < 220:
                buffer = candidate
                continue
            buffer = ""
            continue

        buffer = ""

        if m:
            concepto = m.group(1).strip()
            neto = parse_number(m.group(2)) or 0.0
            alic = parse_number(m.group(3)) or 0.0
            iva = parse_number(m.group(4)) or 0.0
            total = parse_number(m.group(5)) or 0.0
        else:
            concepto = m0.group(1).strip()
            total_base = parse_number(m0.group(2)) or 0.0
            iva = parse_number(m0.group(3)) or 0.0
            total = parse_number(m0.group(4)) or total_base
            neto = total_base
            alic = 0.0

        if pending_detail:
            concepto = f"{pending_detail} {concepto}".strip()
            pending_detail = ""

        out.append(DeductionLine(
            concepto=concepto,
            neto=float(neto),
            alic=float(alic),
            iva=float(iva),
            total=float(total),
        ))

    return out


def _extract_arancel_4905(page_text: str) -> float:
    up = page_text.upper()
    s = up.find("RETENCIONES")
    if s == -1:
        return 0.0

    e_candidates = [
        up.find("IMPORTES TOTALES", s),
        up.find("IMPORTE NETO A PAGAR", s),
        up.find("CONDICIONES DE LA OPERACION", s),
        up.find("CONDICIONES DE LA OPERACIÓN", s),
    ]
    e_candidates = [e for e in e_candidates if e != -1]
    e = min(e_candidates) if e_candidates else len(page_text)

    sec = page_text[s:e]
    lines = [re.sub(r"\s+", " ", l.strip()) for l in sec.splitlines() if l.strip()]

    def amounts_after_dollar(line: str) -> List[float]:
        vals: List[float] = []
        for tok in re.findall(r"\$\s*([-]?\d[\d.,]*)", line):
            v = parse_number(tok)
            if v is None:
                continue
            vals.append(float(v))
        return vals

    pending = False
    for ln in lines:
        n = _norm(ln)

        if ("58487" in ln) or ("ARANCEL" in n and "RESOLUCION" in n):
            pending = True
            vals = amounts_after_dollar(ln)
            for v in reversed(vals):
                if abs(v) > 0:
                    return v
            continue

        if pending:
            vals = amounts_after_dollar(ln)
            for v in reversed(vals):
                if abs(v) > 0:
                    return v

    return 0.0


def parse_liquidacion_pdf(pdf_bytes: bytes, filename: str) -> Liquidacion:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        page0 = pdf.pages[0]
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        page0_text = page0.extract_text() or ""

    full_norm = _norm(full_text)

    fecha, localidad = _extract_header_date_loc(page0_text)

    es_nc = _is_nc(full_norm)
    tipo_cbte = "F2" if es_nc else _detect_tipo_cbte(full_norm)

    mcoe = re.search(r"C\.O\.E\.\s*:\s*([0-9]{8,})", full_text, flags=re.IGNORECASE)
    coe = mcoe.group(1).strip() if mcoe else ""

    pv = coe[:4] if len(coe) >= 4 else ""
    numero = coe[4:12] if len(coe) >= 12 else (coe[4:] if len(coe) > 4 else "")

    try:
        acopio, comprador, vendedor = _extract_parties_from_layout(page0)
    except Exception:
        acopio = Party()
        comprador = Party()
        vendedor = Party()

    grano, cod_neto_venta = _extract_grain(full_text)

    op_adj = None
    if ("AJUSTE UNIFICADO" in full_norm) and re.search(r"AJUSTE\s+CREDITO", full_norm):
        op_adj = _extract_operation_from_ajuste_credito(full_text)

    if op_adj:
        kilos, precio, neto, alic_iva, iva, total = op_adj
    else:
        kilos, precio, neto, alic_iva, iva, total = _extract_operation_numbers_standard(full_text)

    campaña = _extract_campaign(full_text)
    me_nro, me_grado, me_factor, me_prot, me_peso, me_proced = _extract_me(full_text)

    ret_iva, ret_gan = _extract_retenciones(full_text)
    deducciones = _extract_deducciones(full_text)

    ar_4905 = _extract_arancel_4905(full_text)
    if ar_4905 and float(ar_4905) != 0.0:
        deducciones.append(DeductionLine(
            concepto="58487|Arancel Resolución 49/05",
            neto=float(ar_4905),
            alic=0.0,
            iva=0.0,
            total=float(ar_4905),
        ))

    if es_nc:
        neto = _neg_abs(neto)
        iva = _neg_abs(iva)
        total = _neg_abs(total)
        ret_iva = _neg_abs(ret_iva)
        ret_gan = _neg_abs(ret_gan)

        fixed: List[DeductionLine] = []
        for d in deducciones:
            fixed.append(DeductionLine(
                concepto=d.concepto,
                neto=_neg_abs(d.neto),
                alic=float(d.alic),
                iva=_neg_abs(d.iva),
                total=_neg_abs(d.total),
            ))
        deducciones = fixed

    return Liquidacion(
        filename=filename,
        fecha=fecha,
        localidad=localidad,
        tipo_cbte=tipo_cbte,
        letra="A",
        coe=coe,
        pv=pv,
        numero=numero,
        acopio=acopio,
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
        me_grado=me_grado,
        me_factor=me_factor,
        me_contenido_proteico=me_prot,
        me_peso_kg=me_peso,
        me_procedencia=me_proced,
        ret_iva=ret_iva,
        ret_gan=ret_gan,
        deducciones=deducciones,
    )
