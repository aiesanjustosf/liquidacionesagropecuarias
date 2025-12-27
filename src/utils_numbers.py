from __future__ import annotations
import re
from typing import Optional

def parse_ar_number(s: str) -> Optional[float]:
    """
    Convierte números con formato AR/ES o mixto a float.
    Soporta:
      - 1.234.567,89 (AR)
      - 129,250.00 (US)
      - 27,14
      - 0.00
    """
    if s is None:
        return None

    txt = str(s).strip()
    if txt == "":
        return None

    txt = txt.replace("$", "").replace(" ", "")
    # Mantener solo dígitos, separadores y signo
    txt = re.sub(r"[^0-9\-\.,]", "", txt)
    if txt in {"", "-", ",", "."}:
        return None

    # Caso mixto: tiene coma y punto
    if "," in txt and "." in txt:
        last_comma = txt.rfind(",")
        last_dot = txt.rfind(".")
        if last_dot > last_comma:
            # decimal punto (US): quitar comas de miles
            txt = txt.replace(",", "")
        else:
            # decimal coma (AR): quitar puntos de miles y convertir coma a punto
            txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        # decimal coma (AR)
        txt = txt.replace(".", "").replace(",", ".")
    else:
        # solo punto o sin separadores: float lo interpreta
        pass

    try:
        return float(txt)
    except Exception:
        return None
