"""
Parser del mapa horario por Sala.
Respeta celdas combinadas del Excel para mostrar bloques iguales a la hoja.
"""
from openpyxl import load_workbook
from typing import List, Dict, Any, Optional, Tuple
import re

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
RE_CODIGO = re.compile(r"\[([A-Za-z0-9][A-Za-z0-9\-_/]*)\]")


def _limpiar_texto_clase(texto: str) -> str:
    if not texto:
        return ""
    t = str(texto).replace("\r", "").strip()
    t = "\n".join(line.strip() for line in t.split("\n") if line.strip())
    return t


def _es_fila_sala(texto: str) -> bool:
    if not texto or not isinstance(texto, str):
        return False
    t = texto.strip()
    if not t or t.lower().startswith("mapa horario"):
        return False
    if t.startswith("Sala "):
        return True
    if RE_CODIGO.search(t) and re.search(r"(?i)\bsala\b", t):
        return True
    return False


def _extraer_codigo_edificio(texto_sala: str) -> Tuple[Optional[str], Optional[str]]:
    codigo = None
    matches = RE_CODIGO.findall(texto_sala or "")
    for m in matches:
        if re.match(r"(?i)^B-", m) or re.match(r"(?i)^[A-Z]+-\d+", m):
            codigo = m.strip()
            break
    if not codigo and matches:
        codigo = matches[0].strip()
    if not codigo:
        return None, None
    edificio = None
    m2 = re.match(r"^(B-\d+)", codigo, re.IGNORECASE)
    if m2:
        edificio = m2.group(1).upper()
    return codigo, edificio


def _build_merge_map(ws) -> Dict[Tuple[int, int], Tuple[int, int, int, int]]:
    """
    Mapa (row, col) -> (min_row, min_col, max_row, max_col) del rango combinado.
    """
    m = {}
    for rng in ws.merged_cells.ranges:
        min_r, min_c, max_r, max_c = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                m[(r, c)] = (min_r, min_c, max_r, max_c)
    return m


def _cell_value(ws, row: int, col: int, merge_map: dict):
    """Valor visible de la celda (si está combinada, el de la esquina superior izquierda)."""
    key = (row, col)
    if key in merge_map:
        mr, mc, _, _ = merge_map[key]
        return ws.cell(mr, mc).value
    return ws.cell(row, col).value


def _merge_span_rows(row: int, col: int, merge_map: dict) -> int:
    """Cuántas filas abarca la combinación vertical desde esta fila (0 si no inicia aquí)."""
    key = (row, col)
    if key not in merge_map:
        return 1
    mr, mc, xr, xc = merge_map[key]
    if row == mr and col == mc:
        return xr - mr + 1
    # celda cubierta por merge que no es el origen
    return 0


def parsear_aulas_desde_excel(ruta_archivo: str) -> List[Dict[str, Any]]:
    """
    Extrae aulas con horario fila a fila (todas las franjas del Excel)
    y metadatos de rowspan por celda para render igual al Excel.
    horario[dia] = [{"hora", "clase", "rowspan"}, ...]
    """
    # data_only + sin read_only para poder leer merged_cells
    wb = load_workbook(ruta_archivo, data_only=True)
    ws = wb[wb.sheetnames[0]]
    merge_map = _build_merge_map(ws)
    max_row = ws.max_row or 1
    max_col = min(ws.max_column or 8, 12)

    aulas: List[Dict] = []
    codigos_vistos = set()

    sala_actual = None
    leyendo_horas = False
    horario_actual: Dict[str, list] = {d: [] for d in DIAS}
    # Para no repetir el texto en filas cubiertas por rowspan (solo en origen)
    skip_until: Dict[int, int] = {}  # col_index -> last_row inclusive covered

    def guardar_sala():
        nonlocal sala_actual, horario_actual, leyendo_horas, skip_until
        if sala_actual and sala_actual["codigo"] not in codigos_vistos:
            dias_hab = [d for d in DIAS if horario_actual[d]]
            sala_actual["horario"] = {d: list(horario_actual[d]) for d in DIAS}
            sala_actual["dias_habilitados"] = dias_hab
            aulas.append(sala_actual)
            codigos_vistos.add(sala_actual["codigo"])
        sala_actual = None
        horario_actual = {d: [] for d in DIAS}
        leyendo_horas = False
        skip_until = {}

    for r in range(1, max_row + 1):
        celda_a = ws.cell(r, 1).value
        texto_a = str(celda_a).strip() if celda_a is not None else ""

        if _es_fila_sala(texto_a):
            if sala_actual:
                guardar_sala()
            codigo, edificio = _extraer_codigo_edificio(texto_a)
            if not codigo:
                sala_actual = None
                leyendo_horas = False
                continue
            sala_actual = {
                "codigo": codigo,
                "nombre": texto_a,
                "edificio": edificio,
                "capacidad": None,
            }
            horario_actual = {d: [] for d in DIAS}
            leyendo_horas = False
            skip_until = {}
            continue

        if texto_a.lower() == "horas":
            if sala_actual:
                leyendo_horas = True
            continue

        if not (sala_actual and leyendo_horas and texto_a and re.match(r"^\d{1,2}:\d{2}", texto_a)):
            continue

        hora = texto_a
        for i, dia in enumerate(DIAS):
            col = i + 2  # B=2 ... H=8
            if col > max_col:
                break

            # Si esta columna está cubierta por un merge iniciado arriba, marcar vacío con rowspan 0
            if col in skip_until and r <= skip_until[col]:
                horario_actual[dia].append({
                    "hora": hora,
                    "clase": "",
                    "rowspan": 0,  # cubierta por celda superior
                })
                continue

            span = _merge_span_rows(r, col, merge_map)
            raw = _cell_value(ws, r, col, merge_map)
            clase = _limpiar_texto_clase(raw) if raw else ""

            if span > 1 and clase:
                skip_until[col] = r + span - 1
                horario_actual[dia].append({
                    "hora": hora,
                    "clase": clase,
                    "rowspan": span,
                })
            elif span == 0:
                # no debería llegar aquí por skip_until
                horario_actual[dia].append({"hora": hora, "clase": "", "rowspan": 0})
            else:
                # celda normal 1 fila
                horario_actual[dia].append({
                    "hora": hora,
                    "clase": clase,
                    "rowspan": 1 if clase else 1,
                })

    if sala_actual:
        guardar_sala()

    wb.close()
    return aulas
