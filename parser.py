# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

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

    # Prefer explicit CUIT patterns if present (with or without hyphens)
    m = re.search(r"\b(\d{2})\D?(\d{8})\D?(\d)\b", s)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"

    # Otherwise take the first 11-digit token, if any
    m2 = re.search(r"\b(\d{11})\b", s)
    if m2:
        return m2.group(1)

    # Fallback: strip non-digits and keep only the first 11
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
class Liquidacion:
    filename: str
    fecha: str
    localidad: str
    tipo_cbte: str   # F1/F2
    letra: str       # A
    coe: str
    pv: str
    numero: str

    # Datos de encabezado (Acopiador/Consignatario)
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

    # Percepción IVA (si existe)
    perc_iva: float

    # MERCADERIA ENTREGADA (legacy: primer item si existe)
    me_nro_comprobante: str
    me_grado: str
    me_factor: Optional[float]
    me_contenido_proteico: Optional[float]
    me_peso_kg: Optional[float]
    me_procedencia: str

    # MERCADERIA ENTREGADA (lista - múltiples filas)
    me_items: List[Dict[str, Any]]

    # Retenciones
    ret_iva: float
    ret_gan: float

    # Deducciones
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


def _party_from_text(side_text: str) -> Party:
    """Parsea los campos habituales dentro del recuadro de COMPRADOR / VENDEDOR."""
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

            # Continuaciones: líneas siguientes que no sean otro campo
            j = i + 1
            while j < len(lines):
                ln2 = lines[j]
                if re.search(stop_regex, ln2, flags=re.IGNORECASE):
                    break

                # Evitar concatenar nuevos rótulos cuando se mezclan columnas
                if re.search(r"\bRaz[oó]n\s+Social\b\s*:", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bDomicilio\b\s*:", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bC\.U\.I\.T\b", ln2, flags=re.IGNORECASE):
                    break
                if re.search(r"\bI\.V\.A\b", ln2, flags=re.IGNORECASE):
                    break

                # Si la línea es continuación (no tiene ':'), anexar
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

    # Sanitización defensiva: cortar si aparece otro rótulo
    def _cut_at_labels(v: str) -> str:
        v2 = (v or "").strip()
        if not v2:
            return ""
        for lab in ["RAZON SOCIAL", "DOMICILIO", "C.U.I.T", "I.V.A", "LOCALIDAD", "COMPRADOR", "VENDEDOR"]:
            idx = _norm(v2).find(lab)
            if idx > 0:
                v2 = v2[:idx].strip()
        return re.sub(r"\s+", " ", v2).strip()

    razon = _cut_at_labels(razon)
    domicilio = _cut_at_labels(domicilio)

    return Party(razon_social=razon, domicilio=domicilio, localidad=localidad, cuit=cuit, iva=iva)


def _group_words_to_lines(words: List[dict]) -> List[str]:
    """Agrupa words de pdfplumber por línea usando coordenada 'top'."""
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
    """
    Extrae (por layout, evitando mezcla de columnas):
      - acopio (asumido = comprador)
      - comprador (recuadro izquierdo)
      - vendedor (recuadro derecho)

    Incluye SWAP automático si quedan invertidos.
    """
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

    if comprador_x0 is not None and vendedor_x0 is not None and vendedor_x0 != comprador_x0:
        x_split = (comprador_x0 + vendedor_x0) / 2.0
    else:
        x_split = width / 2.0

    def find_top(token: str) -> Optional[float]:
        tnorm = _norm(token)
        tops = [w["top"] for w in words if _norm(w.get("text", "")) == tnorm]
        return min(tops) if tops else None

    y_compr = find_top("COMPRADOR")
    y_vend = find_top("VENDEDOR")
    y_actuo = find_top("ACTUÓ") or find_top("ACTUO")

    y_start_candidates = [y for y in [y_compr, y_vend] if y is not None]
    y_start = min(y_start_candidates) if y_start_candidates else 0.0
    y_end = y_actuo if y_actuo is not None else (y_start + 200.0)

    block_words = [w for w in words if (w.get("top", 0) >= y_start and w.get("top", 0) <= y_end)]
    left_words = [w for w in block_words if w.get("x0", 0) < x_split]
    right_words = [w for w in block_words if w.get("x0", 0) >= x_split]

    left_text = "\n".join(_group_words_to_lines(left_words))
    right_text = "\n".join(_group_words_to_lines(right_words))

    # SWAP automático si se detectan rótulos cruzados
    L = _norm(left_text)
    R = _norm(right_text)
    if ("VENDEDOR" in L and "COMPRADOR" in R) and not ("COMPRADOR" in L and "VENDEDOR" in R):
        left_text, right_text = right_text, left_text

    comprador = _party_from_text(left_text)
    vendedor = _party_from_text(right_text)

    # Acopio = comprador
    acopio = Party(
        razon_social=comprador.razon_social,
        domicilio=comprador.domicilio,
        localidad=comprador.localidad,
        cuit=comprador.cuit,
        iva=comprador.iva,
    )
    return acopio, comprador, vendedor


def _extract_grain(page_text: str) -> Tuple[str, str]:
    m = re.search(
        r"\b(Soja|Ma[ií]z|Trigo|Girasol|Arveja|Sorgo|Camelina\s*Sativa)\b",
        page_text,
        flags=re.IGNORECASE,
    )
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

    return (gname.title() if gname != "MAIZ" else "Maíz"), GRAIN_CODES.get(gname, "")


def _extract_operation_numbers(page_text: str) -> Tuple[float, float, float, float, float, float]:
    """
    Returns: kilos, precio, neto, alic_iva, iva, total

    Línea típica:
      "<kilos> Kg $<precio> $<neto> <alic> $<iva> $<total>"
    """
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


def _extract_me_items(full_text: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Devuelve:
      - items: lista de filas MERCADERIA ENTREGADA (nro, grado, factor, prot, peso, proced)
      - procedencia (si aparece como Localidad:)
    """
    up = (full_text or "").upper()
    start = up.find("MERCADERIA ENTREGADA")
    if start == -1:
        start = up.find("MERCADERÍA ENTREGADA")
    end = up.find("OPERACIÓN", start)
    if end == -1:
        end = up.find("OPERACION", start)

    if start == -1 or end == -1 or end <= start:
        return [], ""

    sec = full_text[start:end]

    proced = ""
    mloc = re.search(r"Localidad\s*:\s*([^\n]+)", sec, flags=re.IGNORECASE)
    if mloc:
        proced = re.sub(r"\s+", " ", mloc.group(1).strip())

    items: List[Dict[str, Any]] = []
    for m in re.finditer(
        r"\b(\d{10,14})\b\s+([A-Z0-9]{1,4})\s+([0-9][0-9.,]*)\s+([0-9][0-9.,]*)\s+([0-9][0-9.,]*)",
        sec,
        flags=re.IGNORECASE,
    ):
        items.append({
            "nro": m.group(1).strip(),
            "grado": m.group(2).strip(),
            "factor": parse_number(m.group(3)),
            "prot": parse_number(m.group(4)),
            "peso": parse_number(m.group(5)),
            "proced": proced,
        })

    return items, proced


def _extract_percepcion_iva(full_text: str) -> float:
    """
    Busca percepción IVA en el texto.
    Tolera rótulos: "Percepción IVA", "Percep. IVA", etc.
    """
    if not full_text:
        return 0.0

    best = 0.0
    for ln in full_text.splitlines():
        ln_norm = _norm(ln)
        if ("PERCEP" in ln_norm or "PERCEPC" in ln_norm) and "IVA" in ln_norm:
            nums = re.findall(r"[-]?\d[\d.,]*", ln)
            if nums:
                v = parse_number(nums[-1]) or 0.0
                if abs(v) > abs(best):
                    best = v

    # Fallback: búsqueda global si no se detectó por línea
    if best == 0.0:
        m = re.search(r"Percep\w*\s*IVA.*?\$?\s*([0-9][0-9.,]*)", full_text, flags=re.IGNORECASE | re.DOTALL)
        if m:
            best = parse_number(m.group(1)) or 0.0

    return float(best or 0.0)


def _extract_retenciones(page_text: str) -> Tuple[float, float]:
    """
    Extrae retenciones IVA y Ganancias del bloque RETENCIONES.
    IMPORTANTE: toma importes ($ xxxx) y NO porcentajes (ej. 5%).
    """
    up = (page_text or "").upper()
    s = up.find("RETENCIONES")
    if s == -1:
        return 0.0, 0.0

    # cortar antes de otros bloques
    e_candidates = [
        up.find("GRADO", s),
        up.find("CONDICIONES", s),
        up.find("OTROS", s),
    ]
    e_candidates = [e for e in e_candidates if e != -1]
    e = min(e_candidates) if e_candidates else len(page_text)

    sec = page_text[s:e]

    def last_currency_amount(line: str) -> Optional[float]:
        # toma el último monto con $ (evita que agarre 5% u otros números sueltos)
        matches = re.findall(r"\$\s*([0-9][0-9.,]*)", line)
        if not matches:
            return None
        return parse_number(matches[-1])

    ret_iva = 0.0
    ret_gan = 0.0

    for ln in sec.splitlines():
        ln_norm = _norm(ln)

        # IVA: líneas que contengan IVA y algún monto con $
        if ("IVA" in ln_norm) and ("$" in ln):
            v = last_currency_amount(ln)
            if v is not None:
                ret_iva = max(ret_iva, float(v))

        # Ganancias
        if ("GANANCIAS" in ln_norm) and ("$" in ln):
            v = last_currency_amount(ln)
            if v is not None:
                ret_gan = max(ret_gan, float(v))

    # fallback por si el pdf partió el $ en otra línea
    if ret_iva == 0.0:
        m = re.search(r"(RETENCION.*IVA|I\.V\.A\.).*?\$\s*([0-9][0-9.,]*)", sec, flags=re.IGNORECASE)
        if m:
            ret_iva = float(parse_number(m.group(2)) or 0.0)

    if ret_gan == 0.0:
        m = re.search(r"GANANCIAS.*?\$\s*([0-9][0-9.,]*)", sec, flags=re.IGNORECASE)
        if m:
            ret_gan = float(parse_number(m.group(1)) or 0.0)

    return float(ret_iva or 0.0), float(ret_gan or 0.0)


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

    for ln in lines:
        # skip headers
        if ln.lower().startswith("concepto") or "base cálculo" in ln.lower() or "base calculo" in ln.lower():
            continue

        ln2 = re.sub(r"\s+", " ", ln)

        # pattern with alic% and iva and total at end
        m = re.search(
            r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+([0-9][0-9.,]*)%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$",
            ln2
        )
        if m:
            concepto = m.group(1).strip()
            neto = parse_number(m.group(2)) or 0.0
            alic = parse_number(m.group(3)) or 0.0
            iva = parse_number(m.group(4)) or 0.0
            total = parse_number(m.group(5)) or 0.0
            out.append(DeductionLine(concepto=concepto, neto=neto, alic=alic, iva=iva, total=total))
            continue

        # pattern 0% where neto not explicit
        m0 = re.search(
            r"^(.*?)\s+\$?\s*([0-9][0-9.,]*)\s+0%?\s+\$?\s*([0-9][0-9.,]*)\s+\$?\s*([0-9][0-9.,]*)\s*$",
            ln2
        )
        if m0:
            concepto = m0.group(1).strip()
            total = parse_number(m0.group(2)) or 0.0
            iva = parse_number(m0.group(3)) or 0.0
            total2 = parse_number(m0.group(4)) or total
            out.append(DeductionLine(concepto=concepto, neto=total, alic=0.0, iva=iva, total=total2))
            continue

    cleaned: List[DeductionLine] = []
    for d in out:
        if _norm(d.concepto) in {"COMISION O GASTOS", "ADMINISTRATIVOS", "OTRAS DEDUCCIONES"}:
            continue
        cleaned.append(d)
    return cleaned


def parse_liquidacion_pdf(pdf_bytes: bytes, filename: str) -> Liquidacion:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        page0 = pdf.pages[0]
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        page0_text = page0.extract_text() or ""

    full_norm = _norm(full_text)

    # Encabezado fecha/localidad
    fecha, localidad = _extract_header_date_loc(page0_text)
    tipo_cbte = _detect_tipo_cbte(full_norm)

    # COE / punto de venta / número
    mcoe = re.search(r"C\.O\.E\.\s*:\s*([0-9]{8,})", full_text, flags=re.IGNORECASE)
    coe = mcoe.group(1).strip() if mcoe else ""
    pv = coe[:4] if len(coe) >= 4 else ""
    numero = coe[4:12] if len(coe) >= 12 else (coe[4:] if len(coe) > 4 else "")

    # Partes (comprador/vendedor) por layout
    try:
        acopio, comprador, vendedor = _extract_parties_from_layout(page0)
    except Exception:
        acopio = Party()
        comprador = Party()
        vendedor = Party()

    # Grano / operación
    grano, cod_neto_venta = _extract_grain(full_text)
    kilos, precio, neto, alic_iva, iva, total = _extract_operation_numbers(full_text)

    campaña = _extract_campaign(full_text)

    # Mercadería entregada (lista + legacy)
    me_items, me_proced = _extract_me_items(full_text)
    if me_items:
        me0 = me_items[0]
        me_nro = me0.get("nro", "") or ""
        me_grado = me0.get("grado", "") or ""
        me_factor = me0.get("factor", None)
        me_prot = me0.get("prot", None)
        me_peso = me0.get("peso", None)
    else:
        me_nro, me_grado, me_factor, me_prot, me_peso = "", "", None, None, None

    # Retenciones / deducciones
    ret_iva, ret_gan = _extract_retenciones(full_text)
    deducciones = _extract_deducciones(full_text)

    # Percepción IVA
    perc_iva = _extract_percepcion_iva(full_text)

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
        kilos=float(kilos or 0),
        precio=float(precio or 0),
        neto=float(neto or 0),
        alic_iva=float(alic_iva or 0),
        iva=float(iva or 0),
        total=float(total or 0),
        campaña=campaña,
        perc_iva=float(perc_iva or 0),
        me_nro_comprobante=me_nro,
        me_grado=me_grado,
        me_factor=me_factor,
        me_contenido_proteico=me_prot,
        me_peso_kg=me_peso,
        me_procedencia=me_proced,
        me_items=me_items,
        ret_iva=float(ret_iva or 0),
        ret_gan=float(ret_gan or 0),
        deducciones=deducciones,
    )
