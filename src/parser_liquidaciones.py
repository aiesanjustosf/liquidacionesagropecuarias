from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
import re
from typing import Optional, List, Dict, Tuple

import pdfplumber

from .utils_numbers import parse_ar_number


@dataclass
class MercaderiaEntregada:
    nro_comprobante: Optional[str] = None
    procedencia: Optional[str] = None
    peso_kg: Optional[float] = None
    puerto: Optional[str] = None
    grado: Optional[str] = None
    factor: Optional[str] = None
    contenido_proteico: Optional[str] = None


@dataclass
class Deduccion:
    detalle: str
    base: Optional[float] = None
    alicuota: Optional[float] = None
    iva: Optional[float] = None
    total: Optional[float] = None
    exento: Optional[float] = None


@dataclass
class LiquidacionDoc:
    filename: str = ""
    tipo_comprobante: str = "F1"  # F1 / F2
    coe: Optional[str] = None

    fecha: Optional[str] = None
    localidad: Optional[str] = None

    comprador_rs: Optional[str] = None
    comprador_cuit: Optional[str] = None
    comprador_cf: Optional[str] = None
    comprador_dom: Optional[str] = None

    vendedor_rs: Optional[str] = None
    vendedor_cuit: Optional[str] = None

    grano: Optional[str] = None
    campania: Optional[str] = None
    kilos: Optional[float] = None
    precio_kg: Optional[float] = None

    subtotal: Optional[float] = None
    alicuota_iva: Optional[float] = None
    iva: Optional[float] = None
    total: Optional[float] = None

    ret_iva: Optional[float] = None
    ret_gan: Optional[float] = None

    mercaderia_entregada: MercaderiaEntregada = field(default_factory=MercaderiaEntregada)
    deducciones: List[Deduccion] = field(default_factory=list)


_GRANO_MAP = {
    "SOJA": "Soja",
    "MAIZ": "Maíz",
    "MAÍZ": "Maíz",
    "TRIGO": "Trigo",
    "GIRASOL": "Girasol",
    "ARVEJA": "Arveja",
    "SORGO": "Sorgo",
    "CAMELINA SATIVA": "Camelina Sativa",
    "CAMELINA": "Camelina Sativa",
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _upper_ascii(s: str) -> str:
    s = (s or "").upper()
    s = s.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U").replace("Ü", "U").replace("Ñ", "N")
    return s


def _get_full_text(pdf_bytes: bytes) -> str:
    parts: List[str] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            parts.append(p.extract_text() or "")
    return "\n".join(parts)


def _detect_tipo_comprobante(full_text: str) -> str:
    t = _upper_ascii(full_text)
    if "LIQUIDACION SECUNDARIA DE GRANOS" in t:
        return "F2"
    return "F1"


def _extract_coe(full_text: str) -> Optional[str]:
    m = re.search(r"C\.O\.E\.?\s*:\s*(\d{12,})", full_text)
    if not m:
        m = re.search(r"\bCOE\b\s*:\s*(\d{12,})", full_text, flags=re.I)
    return m.group(1) if m else None


def _extract_fecha_localidad(full_text: str) -> Tuple[Optional[str], Optional[str]]:
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s*[,–\-]\s*([A-ZÁÉÍÓÚÜÑ\.\-\s]+)", full_text)
    if m:
        return m.group(1), _norm(m.group(2)).title()
    m = re.search(r"(\d{2}/\d{2}/\d{4})", full_text)
    return (m.group(1) if m else None, None)


def _extract_campania(full_text: str) -> Optional[str]:
    m = re.search(r"Campaña\s*:\s*([^\n\r]+)", full_text, flags=re.I)
    if not m:
        return None
    val = _norm(m.group(1))
    val = re.split(r"\b(Procedencia|Peso|Grado|Puerto|Flete)\b", val, flags=re.I)[0].strip()
    return val or None


def _parse_party_block(block: str) -> Dict[str, Optional[str]]:
    def _clean_val(val: Optional[str]) -> Optional[str]:
        if not val:
            return val
        v = _norm(val)
        # Si se pegó con otra columna, cortar por etiquetas recurrentes
        for stop in ["Razón Social:", "Razon Social:", "Domicilio:", "C.U.I.T.", "I.V.A."]:
            if stop in v:
                v = v.split(stop)[0].strip()
        return v or None

    rs = None
    cuit = None
    dom = None
    iva = None

    m = re.search(r"Razón Social\s*:\s*([^\n\r]+)", block, flags=re.I)
    if m:
        rs = _clean_val(m.group(1))

    m = re.search(r"C\.U\.I\.T\.?\s*:\s*(\d{11})", block, flags=re.I)
    if m:
        cuit = m.group(1)

    m = re.search(r"Domicilio\s*:\s*([^\n\r]+)", block, flags=re.I)
    if m:
        dom = _clean_val(m.group(1))

    m = re.search(r"I\.V\.A\.?\s*:\s*([A-Z]+)", block, flags=re.I)
    if m:
        iva = _norm(m.group(1)).upper()

    cf = "RI" if iva == "RI" else (iva or None)
    return {"rs": rs, "cuit": cuit, "dom": dom, "cf": cf}


def _extract_parties(full_text: str) -> Dict[str, Optional[str]]:
    t = full_text
    t_up = _upper_ascii(t)

    buyer = {"rs": None, "cuit": None, "dom": None, "cf": None}
    seller = {"rs": None, "cuit": None, "dom": None, "cf": None}

    if "COMPRADOR" in t_up and "VENDEDOR" in t_up:
        idx_c = t_up.find("COMPRADOR")
        idx_v = t_up.find("VENDEDOR", idx_c + 1)
        if idx_c != -1 and idx_v != -1:
            buyer = _parse_party_block(t[idx_c:idx_v])
            seller = _parse_party_block(t[idx_v:])

    if not buyer["rs"] or not buyer["cuit"]:
        rs_all = re.findall(r"Razón Social\s*:\s*([^\n\r]+)", t, flags=re.I)
        cuits = re.findall(r"C\.U\.I\.T\.?\s*:\s*(\d{11})", t, flags=re.I)
        doms = re.findall(r"Domicilio\s*:\s*([^\n\r]+)", t, flags=re.I)
        ivas = re.findall(r"I\.V\.A\.?\s*:\s*([A-Z]+)", t, flags=re.I)

        if len(rs_all) >= 1: buyer["rs"] = _norm(rs_all[0]).split("Razón Social:")[0].split("Razon Social:")[0].strip() or None
        if len(rs_all) >= 2: seller["rs"] = _norm(rs_all[1]).split("Razón Social:")[0].split("Razon Social:")[0].strip() or None

        if len(cuits) >= 1: buyer["cuit"] = cuits[0]
        if len(cuits) >= 2: seller["cuit"] = cuits[1]

        if len(doms) >= 1: buyer["dom"] = _norm(doms[0])
        if len(doms) >= 2: seller["dom"] = _norm(doms[1])

        if len(ivas) >= 1:
            iva = _norm(ivas[0]).upper()
            buyer["cf"] = "RI" if iva == "RI" else iva

    return {
        "buyer_rs": buyer["rs"],
        "buyer_cuit": buyer["cuit"],
        "buyer_dom": buyer["dom"],
        "buyer_cf": buyer["cf"],
        "seller_rs": seller["rs"],
        "seller_cuit": seller["cuit"],
    }


def _extract_grano(full_text: str) -> Optional[str]:
    lines = (full_text or "").splitlines()
    for i, line in enumerate(lines):
        if re.search(r"\bGrano\b", line, flags=re.I):
            # buscar en esta y próximas 2 líneas (extract_text suele partir en renglones)
            for j in range(i, min(i + 3, len(lines))):
                m = re.search(r"\d+\s*-\s*([A-ZÁÉÍÓÚÜÑ\s]+)", lines[j])
                if m:
                    key = _upper_ascii(_norm(m.group(1)))
                    for k, v in _GRANO_MAP.items():
                        if k in key:
                            return v

    # fallback global
    m = re.search(r"\d+\s*-\s*([A-ZÁÉÍÓÚÜÑ\s]+)", full_text)
    if m:
        key = _upper_ascii(_norm(m.group(1)))
        for k, v in _GRANO_MAP.items():
            if k in key:
                return v
    return None


def _extract_kilos(full_text: str) -> Optional[float]:
    m = re.search(r"Peso\s*:?\s*([\d\.,]+)\s*kg", full_text, flags=re.I)
    if not m:
        m = re.search(r"Cantidad\s*:?\s*([\d\.,]+)\s*kg", full_text, flags=re.I)
    return parse_ar_number(m.group(1)) if m else None


def _extract_precio_kg(full_text: str) -> Optional[float]:
    m = re.search(r"Precio\s*/\s*Kg\s*:?\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
    if not m:
        m = re.search(r"Precio\s*/\s*kg\s*:?\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
    return parse_ar_number(m.group(1)) if m else None


def _extract_totales(full_text: str) -> Dict[str, Optional[float]]:
    subtotal: Optional[float] = None
    iva: Optional[float] = None
    total: Optional[float] = None
    alic: Optional[float] = None

    # Caso típico: línea OPERACIÓN con 3 importes + alícuota
    m = re.search(
        r"OPERACI[ÓO]N.*?\$\s*([\d\.,]+)\s+([\d\.,]+)\s+\$\s*([\d\.,]+)\s+\$\s*([\d\.,]+)",
        full_text,
        flags=re.I | re.S
    )
    if m:
        subtotal = parse_ar_number(m.group(1))
        alic = parse_ar_number(m.group(2))
        iva = parse_ar_number(m.group(3))
        total = parse_ar_number(m.group(4))

    # Alternativa con etiquetas
    if subtotal is None:
        m = re.search(r"Subtotal\s*:\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
        if m:
            subtotal = parse_ar_number(m.group(1))

    if iva is None or alic is None:
        m = re.search(r"IVA\s+([\d\.,]+)\s*%\s*:\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
        if m:
            alic = parse_ar_number(m.group(1))
            iva = parse_ar_number(m.group(2))

    if total is None:
        m = re.search(r"Total\s*Operación\s*:\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
        if m:
            total = parse_ar_number(m.group(1))
        else:
            m = re.search(r"Operación\s*c/IVA\s*\$?\s*([\d\.,]+)", full_text, flags=re.I)
            if m:
                total = parse_ar_number(m.group(1))

    return {"subtotal": subtotal, "iva": iva, "total": total, "alicuota": alic}


def _extract_mercaderia_entregada(full_text: str) -> MercaderiaEntregada:
    me = MercaderiaEntregada()
    m = re.search(r"N[º°]\s*de\s*Comprobante\s*(?:Grado|:)\s*(\d+)", full_text, flags=re.I)
    if not m:
        m = re.search(r"N[º°]\s*de\s*Comprobante\s*:?\s*(\d+)", full_text, flags=re.I)
    if m:
        me.nro_comprobante = m.group(1)

    m = re.search(r"Procedencia\s*de\s*la\s*Mercaderia\s*\n?.*?\b([A-ZÁÉÍÓÚÜÑ\-\s]+)\b", full_text, flags=re.I)
    # Si no funciona, usar el "Procedencia:" clásico
    m2 = re.search(r"Procedencia\s*:?\s*([^\n\r]+)", full_text, flags=re.I)
    if m2:
        me.procedencia = _norm(m2.group(1))

    m = re.search(r"Puerto\s*:?\s*([^\n\r]+)", full_text, flags=re.I)
    if m:
        me.puerto = _norm(m.group(1))

    m = re.search(r"Grado\s*:?\s*([A-Z0-9]+)", full_text, flags=re.I)
    if m:
        me.grado = _norm(m.group(1))

    m = re.search(r"Factor\s*:?\s*([A-Z0-9\.,]+)", full_text, flags=re.I)
    if m:
        me.factor = _norm(m.group(1))

    m = re.search(r"Contenido\s*Proteico\s*:?\s*([A-Z0-9\.,]+)", full_text, flags=re.I)
    if m:
        me.contenido_proteico = _norm(m.group(1))

    m = re.search(r"MERCADERIA\s+ENTREGADA.*?\b([\d\.,]+)\b\s*\n?FE", _upper_ascii(full_text), flags=re.S)
    # En varios PDFs el peso en mercadería entregada aparece como ".... 10000 FE"
    if m:
        me.peso_kg = parse_ar_number(m.group(1))
    else:
        m = re.search(r"Mercadería\s+Entregada.*?Peso\s*:?\s*([\d\.,]+)\s*kg", full_text, flags=re.I | re.S)
        if m:
            me.peso_kg = parse_ar_number(m.group(1))

    return me


def _extract_deducciones(full_text: str) -> List[Deduccion]:
    """
    Extrae deducciones desde el bloque DEDUCCIONES hasta RETENCIONES/IMPORTES.
    Formato observado en extract_text:
      <detalle...> $ <base> <alicuota>% $ <iva> $ <total>
    """
    up = _upper_ascii(full_text)
    start = up.find("DEDUCCIONES")
    if start == -1:
        return []
    end = up.find("RETENCIONES", start)
    if end == -1:
        end = up.find("IMPORTES", start)
    block = full_text[start:end] if end != -1 else full_text[start:]

    deducs: List[Deduccion] = []
    for line in block.splitlines():
        ln = _norm(line)
        if not ln or ln.upper().startswith("DEDUCCIONES") or "BASE CÁLCULO" in ln.upper() or "BASE CALCULO" in ln.upper():
            continue
        # Buscar patrón principal: ... $ base 10.5% $ iva $ total
        m = re.search(r"^(?P<detalle>.+?)\s+\$\s*(?P<base>[\d\.,]+)\s+(?P<alic>[\d\.,]+)\s*%\s+\$\s*(?P<iva>[\d\.,]+)\s+\$\s*(?P<tot>[\d\.,]+)$", ln)
        if not m:
            continue

        detalle = _norm(m.group("detalle"))
        base = parse_ar_number(m.group("base"))
        alic = parse_ar_number(m.group("alic"))
        iva = parse_ar_number(m.group("iva"))
        tot = parse_ar_number(m.group("tot"))

        exento = None
        if alic is not None and abs(alic) < 0.0001:
            # Exento: guardar importe en exento
            exento = tot if tot is not None else base
            base = None
            iva = None
            alic = 0.0

        deducs.append(Deduccion(detalle=detalle, base=base, alicuota=alic, iva=iva, total=tot, exento=exento))

    return deducs


def _extract_retenciones_from_tables(pdf_bytes: bytes) -> Tuple[Optional[float], Optional[float]]:
    ret_iva = 0.0
    ret_gan = 0.0
    found = False

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for tbl in (page.extract_tables() or []):
                flat = " ".join([_norm(str(c)) for row in tbl for c in row if c is not None])
                flat_up = _upper_ascii(flat)
                if "RETENCIONES" not in flat_up and "RETENCION" not in flat_up:
                    continue

                for row in tbl:
                    row_txt = " ".join([_norm(str(c)) for c in row if c is not None])
                    row_up = _upper_ascii(row_txt)
                    nums = re.findall(r"[\$]?\s*[\d\.,]+", row_txt)
                    amt = parse_ar_number(nums[-1]) if nums else None
                    if amt is None:
                        continue

                    if "GAN" in row_up:
                        ret_gan += amt
                        found = True

                    if "IVA" in row_up and "4310" not in row_up and "RG" not in row_up:
                        ret_iva += amt
                        found = True

    if not found:
        return None, None
    return (ret_iva if ret_iva != 0 else None, ret_gan if ret_gan != 0 else None)


def _extract_retenciones_from_block(full_text: str) -> Tuple[Optional[float], Optional[float]]:
    """
    Extrae retenciones del bloque RETENCIONES.
    - Soporta casos donde el % queda al final de una línea y el importe real aparece en la(s) línea(s) siguiente(s).
    - Evita capturar porcentajes y evita asociar el importe de IVA a GAN o viceversa.
    """
    up = _upper_ascii(full_text)
    start = up.find("RETENCIONES")
    if start == -1:
        return None, None
    end = up.find("IMPORTES", start)
    if end == -1:
        end = up.find("IMPORTE", start)
    block = full_text[start:end] if end != -1 else full_text[start:]

    lines = [_norm(l) for l in block.splitlines() if _norm(l)]
    ret_iva: Optional[float] = None
    ret_gan: Optional[float] = None

    def money_candidates(s: str):
        cands = re.findall(r"\d[\d\.,]*\d", s)
        filtered = []
        for c in cands:
            if re.search(rf"{re.escape(c)}\s*%", s):
                continue
            filtered.append(c)
        return filtered

    last_used_idx = -1

    def take_next_line_amount(base_from: int, stop_words: List[str]) -> Tuple[Optional[float], int]:
        for j in range(base_from, min(base_from + 6, len(lines))):
            lup = _upper_ascii(lines[j])
            if any(sw in lup for sw in stop_words):
                break
            cands = money_candidates(lines[j])
            if len(cands) >= 2:
                return parse_ar_number(cands[-1]), j
            if len(cands) == 1:
                v = parse_ar_number(cands[0])
                if v not in (None, 0.0):
                    return v, j
        return None, -1

    for i, ln in enumerate(lines):
        lup = _upper_ascii(ln)

        if ("RET IVA" in lup) or ("IVA" in lup and "RET" in lup and "4310" not in lup):
            cands = money_candidates(ln)
            amt = None
            used_idx = -1
            if cands:
                amt = parse_ar_number(cands[-1])
                if amt in (None, 0.0):
                    amt, used_idx = take_next_line_amount(max(i + 1, last_used_idx + 1), stop_words=["GAN"])
            else:
                amt, used_idx = take_next_line_amount(max(i + 1, last_used_idx + 1), stop_words=["GAN"])

            if amt not in (None, 0.0):
                ret_iva = amt
                if used_idx != -1:
                    last_used_idx = used_idx

        if ("RET GAN" in lup) or ("GAN" in lup and "RET" in lup):
            cands = money_candidates(ln)
            amt = None
            used_idx = -1
            if cands:
                amt = parse_ar_number(cands[-1])
                if amt in (None, 0.0):
                    amt, used_idx = take_next_line_amount(max(i + 1, last_used_idx + 1), stop_words=["IVA"])
            else:
                amt, used_idx = take_next_line_amount(max(i + 1, last_used_idx + 1), stop_words=["IVA"])

            if amt not in (None, 0.0):
                ret_gan = amt
                if used_idx != -1:
                    last_used_idx = used_idx

    return ret_iva, ret_gan


def _extract_retenciones_regex(full_text: str) -> Tuple[Optional[float], Optional[float]]:
    """
    Fallback conservador: busca líneas con 'RET IVA' / 'RET GAN' y toma el ÚLTIMO importe "grande".
    Evita capturar números de certificado (1, 0, etc.).
    """
    ret_iva = None
    ret_gan = None

    def last_big_amount(line: str) -> Optional[float]:
        # priorizar candidatos con separador o longitud > 2 (evita "10", "1")
        cands = re.findall(r"\d[\d\.,]*\d", line)
        good = []
        for c in cands:
            if re.search(rf"{re.escape(c)}\s*%", line):
                continue
            if ("," in c or "." in c) or len(re.sub(r"\D", "", c)) > 2:
                val = parse_ar_number(c)
                if val is not None:
                    good.append(val)
        if not good:
            return None
        # elegir el último no-cero
        for v in reversed(good):
            if v != 0:
                return v
        return None

    for line in full_text.splitlines():
        ln = _norm(line)
        lup = _upper_ascii(ln)
        if ("RET IVA" in lup) and ("4310" not in lup):
            v = last_big_amount(ln)
            if v is not None:
                ret_iva = v
        if "RET GAN" in lup:
            v = last_big_amount(ln)
            if v is not None:
                ret_gan = v

    return ret_iva, ret_gan


def parse_liquidacion_pdf(pdf_bytes: bytes, filename: str = "") -> LiquidacionDoc:
    text = _get_full_text(pdf_bytes)
    doc = LiquidacionDoc(filename=filename)

    doc.tipo_comprobante = _detect_tipo_comprobante(text)
    doc.coe = _extract_coe(text)
    doc.fecha, doc.localidad = _extract_fecha_localidad(text)

    parties = _extract_parties(text)
    doc.comprador_rs = parties["buyer_rs"]
    doc.comprador_cuit = parties["buyer_cuit"]
    doc.comprador_dom = parties["buyer_dom"]
    doc.comprador_cf = parties["buyer_cf"]

    doc.vendedor_rs = parties["seller_rs"]
    doc.vendedor_cuit = parties["seller_cuit"]

    doc.grano = _extract_grano(text)
    doc.campania = _extract_campania(text)

    doc.kilos = _extract_kilos(text)
    doc.precio_kg = _extract_precio_kg(text)

    tots = _extract_totales(text)
    doc.subtotal = tots["subtotal"]
    doc.alicuota_iva = tots["alicuota"] if tots["alicuota"] is not None else 10.5
    doc.iva = tots["iva"]
    doc.total = tots["total"]

    doc.mercaderia_entregada = _extract_mercaderia_entregada(text)
    doc.deducciones = _extract_deducciones(text)

    # Retenciones: 1) bloque RETENCIONES (línea), 2) tablas, 3) regex
    b_iva, b_gan = _extract_retenciones_from_block(text)
    t_iva, t_gan = _extract_retenciones_from_tables(pdf_bytes)
    r_iva, r_gan = _extract_retenciones_regex(text)

    doc.ret_iva = b_iva if b_iva is not None else (t_iva if t_iva is not None else r_iva)
    doc.ret_gan = b_gan if b_gan is not None else (t_gan if t_gan is not None else r_gan)

    return doc
