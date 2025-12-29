# parser.py
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
    s = str(raw)

    m = re.search(r"\b(\d{2})\D?(\d{8})\D?(\d)\b", s)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"

    m2 = re.search(r"\b(\d{11})\b", s)
    if m2:
        return m2.group(1)

    d = re.sub(r"\D", "", s)
    return d[:11] if len(d) >= 11 else d


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
class MercaderiaEntregadaItem:
    nro_comprobante: str = ""
    grado: str = ""
    factor: Optional[float] = None
    contenido_proteico: Optional[float] = None
    peso_kg: Optional[float] = None
    procedencia: str = ""


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

    # Datos de encabezado
    acopio: Party       # COMPRADOR (izquierda)
    vendedor: Party     # VENDEDOR (derecha)

    grano: str
    cod_neto_venta: str
    kilos: float
    precio: float
    neto: float
    alic_iva: float
    iva: float
    total: float
    campaña: str

    mercaderia_entregada: List[MercaderiaEntregadaItem]

    percep_iva: float
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

    return Party(
        razon_social=_cut_at_labels(razon),
        domicilio=_cut_at_labels(domicilio),
        localidad=localidad,
        cuit=cuit,
        iva=iva,
    )


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


def _extract_parties_from_layout(page: pdfplumber.page.Page) -> Tuple[Party, Party]:
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
    return comprador, vendedor


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


def _extract_operation_numbers(page_text: str) -> Tuple[float, float, float, float, float, float]:
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


def _extract_me_items(full_text: str) -> List[MercaderiaEntregadaItem]:
    up = full_text.upper()
    start = up.find("MERCADERIA ENTREGADA")
    if start == -1:
        start = up.find("MERCADERÍA ENTREGADA")
    if start == -1:
        return []

    end = up.find("OPERACIÓN", start)
    if end == -1:
        end = up.find("OPERACION", start)
    if end == -1 or end <= start:
        return []

    sec = full_text[start:end]
    items: List[MercaderiaEntregadaItem] = []

    for raw_ln in sec.splitlines():
        ln = re.sub(r"\s+", " ", (raw_ln or "").strip())
        if not ln:
            continue
        if "LOCALIDAD" not in ln.upper():
            continue

        # Procedencia
        proc = ""
        mloc = re.search(r"Localidad\s*:\s*(.*)$", ln, flags=re.IGNORECASE)
        if mloc:
            proc = re.sub(r"\s+", " ", mloc.group(1).strip())
            ln = ln[:mloc.start()].strip()

        parts = ln.split()
        if not parts:
            continue

        # nro comprobante (puede ser "Sin informar")
        idx = 0
        if len(parts) >= 2 and parts[0].lower() == "sin" and parts[1].lower() == "informar":
            nro = "Sin informar"
            idx = 2
        else:
            nro = parts[0]
            idx = 1

        grado = parts[idx] if len(parts) > idx else ""
        num_tokens = parts[idx + 1:] if len(parts) > (idx + 1) else []
        nums = []
        for t in num_tokens:
            v = parse_number(t)
            if v is not None:
                nums.append(v)

        factor = nums[0] if len(nums) >= 1 else None
        prot = nums[1] if len(nums) >= 2 else None
        peso = nums[2] if len(nums) >= 3 else None

        items.append(MercaderiaEntregadaItem(
            nro_comprobante=nro,
            grado=grado,
            factor=factor,
            contenido_proteico=prot,
            peso_kg=peso,
            procedencia=proc,
        ))

    return items


def _extract_retenciones(full_text: str) -> Tuple[float, float]:
    """
    Lee el MONTO de la columna 'Retenciones' del cuadro RETENCIONES.
    Soporta 2 variantes reales del PDF:
      A) ... 5% $ 479,167.92 (Regimen...) 9,583,358.40   -> toma 479,167.92
      B) ... 5%  <base> <ret>                              -> toma el que cumpla ret≈base*5%
    Ganancias sólo se devuelve si el monto > 0.
    """

    up = (full_text or "").upper()
    s = up.find("RETENCIONES")
    if s == -1:
        return 0.0, 0.0

    # Cortar la sección RETENCIONES
    end_candidates = []
    for token in ["GRADO", "CONDICIONES", "OTROS", "MERCADERIA ENTREGADA", "MERCADERÍA ENTREGADA",
                  "IMPORTES TOTALES", "OPERACIÓN", "OPERACION", "FIRMA"]:
        idx = up.find(token, s + 1)
        if idx != -1:
            end_candidates.append(idx)
    e = min(end_candidates) if end_candidates else len(full_text)

    sec = full_text[s:e]
    sec_compact = re.sub(r"\s+", " ", sec)

    def _parse_after_percent(snippet: str) -> float:
        """
        Dado un snippet que contiene '... <p>% ...', extrae el monto de retención.
        Prioridad:
          1) primer número precedido por '$' después del '%'
          2) si no hay '$': tomar 2 números y elegir el que cumpla ret≈base*(p/100)
        """
        mperc = re.search(r"(\d[\d.,]*)\s*%", snippet)
        if not mperc:
            return 0.0
        p = parse_number(mperc.group(1)) or 0.0
        if abs(p) < 1e-12:
            return 0.0

        tail = snippet[mperc.end():]  # texto luego del %
        tail = tail.strip()

        # 1) caso con $ (retención explícita)
        m_money = re.search(r"\$\s*([0-9][0-9.,]*)", tail)
        if m_money:
            return float(parse_number(m_money.group(1)) or 0.0)

        # 2) caso sin $ (suele venir: base ret)
        nums = [parse_number(x) for x in re.findall(r"[-]?\d[\d.,]*", tail)]
        nums = [float(x) for x in nums if x is not None]

        # usar los 2 primeros "grandes" (base/ret)
        big = [x for x in nums if abs(x) >= 1]
        if len(big) >= 2:
            a, b = big[0], big[1]
            # decidir cuál es base y cuál ret por relación matemática
            # probar (a base, b ret) y (b base, a ret)
            exp1 = a * (p / 100.0)
            err1 = abs(b - exp1) / max(1.0, abs(exp1))

            exp2 = b * (p / 100.0)
            err2 = abs(a - exp2) / max(1.0, abs(exp2))

            if err1 <= err2 and err1 <= 0.02:
                return float(b)
            if err2 < err1 and err2 <= 0.02:
                return float(a)

        return 0.0

    def _sum_rows(row_regex: str) -> float:
        total = 0.0
        for m in re.finditer(row_regex, sec_compact, flags=re.IGNORECASE):
            # tomar una ventana razonable desde el match para capturar % y montos
            snippet = sec_compact[m.start(): m.start() + 220]
            amt = _parse_after_percent(snippet)
            if amt and abs(amt) > 1e-9:
                total += amt
        return float(total)

    # IVA: sumar todas las filas de IVA del cuadro
    ret_iva = _sum_rows(r"(?:I\.?V\.?A\.?|IVA).*?(?:RET|RETENCION)")

    # GAN: sumar, pero en tus PDFs suele ser 0; si es 0 no devuelve nada
    ret_gan = _sum_rows(r"(?:GANANCIAS|RET\.?\s*GAN|RET\s*GAN).*?(?:RET|RETENCION)")

    # si es ~0, dejar en 0 para que no se exporte
    if abs(ret_iva) < 1e-6:
        ret_iva = 0.0
    if abs(ret_gan) < 1e-6:
        ret_gan = 0.0

    return ret_iva, ret_gan


def _extract_percep_iva(full_text: str) -> float:
    # IVA RG 4310/2018: $ 123,45
    m = re.search(r"IVA\s*RG\s*4310/2018\s*:\s*\$?\s*([0-9][0-9.,]*)", full_text, flags=re.IGNORECASE)
    if m:
        return float(parse_number(m.group(1)) or 0.0)
    return 0.0


def _extract_deducciones(page_text: str) -> List[DeductionLine]:
    up = page_text.upper()
    s = up.find("DEDUCCIONES")
    e = up.find("RETENCIONES")
    if s == -1 or e == -1 or e <= s:
        return []
    sec = page_text[s:e]
    lines = [l.strip() for l in sec.splitlines() if l.strip()]
    out: List[DeductionLine] = []

    for ln in lines:
        if ln.lower().startswith("concepto") or "base cálculo" in ln.lower():
            continue
        ln2 = re.sub(r"\s+", " ", ln)

        m = re.search(
            r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+([0-9][0-9.,]*)%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$",
            ln2
        )
        if m:
            out.append(DeductionLine(
                concepto=m.group(1).strip(),
                neto=parse_number(m.group(2)) or 0.0,
                alic=parse_number(m.group(3)) or 0.0,
                iva=parse_number(m.group(4)) or 0.0,
                total=parse_number(m.group(5)) or 0.0,
            ))
            continue

        m0 = re.search(
            r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+0%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$",
            ln2
        )
        if m0:
            total = parse_number(m0.group(2)) or 0.0
            iva = parse_number(m0.group(3)) or 0.0
            total2 = parse_number(m0.group(4)) or total
            out.append(DeductionLine(
                concepto=m0.group(1).strip(),
                neto=total,
                alic=0.0,
                iva=iva,
                total=total2,
            ))
            continue

    cleaned: List[DeductionLine] = []
    for d in out:
        if _norm(d.concepto) in {"COMISION O GASTOS ADMINISTRATIVOS", "OTRAS DEDUCCIONES"}:
            continue
        cleaned.append(d)
    return cleaned


def parse_liquidacion_pdf(pdf_bytes: bytes, filename: str) -> Liquidacion:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        page0 = pdf.pages[0]
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        page0_text = page0.extract_text() or ""

    full_norm = _norm(full_text)

    fecha, localidad = _extract_header_date_loc(page0_text)
    tipo_cbte = _detect_tipo_cbte(full_norm)

    mcoe = re.search(r"C\.O\.E\.\s*:\s*([0-9]{8,})", full_text, flags=re.IGNORECASE)
    coe = mcoe.group(1).strip() if mcoe else ""
    pv = coe[:4] if len(coe) >= 4 else ""
    numero = coe[4:12] if len(coe) >= 12 else (coe[4:] if len(coe) > 4 else "")

    try:
        comprador, vendedor = _extract_parties_from_layout(page0)
    except Exception:
        comprador = Party()
        vendedor = Party()

    grano, cod_neto_venta = _extract_grain(full_text)
    kilos, precio, neto, alic_iva, iva, total = _extract_operation_numbers(full_text)

    campaña = _extract_campaign(full_text)
    me_items = _extract_me_items(full_text)

    ret_iva, ret_gan = _extract_retenciones(full_text)
    percep_iva = _extract_percep_iva(full_text)
    deducciones = _extract_deducciones(full_text)

    return Liquidacion(
        filename=filename,
        fecha=fecha,
        localidad=localidad,
        tipo_cbte=tipo_cbte,
        letra="A",
        coe=coe,
        pv=pv,
        numero=numero,
        acopio=comprador,   # COMPRADOR
        vendedor=vendedor,  # VENDEDOR
        grano=grano,
        cod_neto_venta=cod_neto_venta,
        kilos=kilos,
        precio=precio,
        neto=neto,
        alic_iva=alic_iva,
        iva=iva,
        total=total,
        campaña=campaña,
        mercaderia_entregada=me_items,
        percep_iva=percep_iva,
        ret_iva=ret_iva,
        ret_gan=ret_gan,
        deducciones=deducciones,
    )
