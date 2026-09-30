"""
Parser del archivo de mapa horario por Sala.
Compatible con Excel de edificios 101, 106, etc.
Formato: "Sala SALÓN DE CLASES - [B-106A-101]" + filas de horas por día.
"""
from openpyxl import load_workbook
from typing import List, Dict, Any, Optional
import re


DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


def _limpiar_texto_clase(texto: str) -> str:
    """Resume el texto de la celda de clase para mostrar."""
    if not texto:
        return ""
    t = str(texto).replace("\n", " ").strip()
    if " | " in t:
        partes = t.split(" | ")
        codigo = partes[0].strip()
        resto = partes[1].strip() if len(partes) > 1 else ""
        if len(resto) > 40:
            resto = resto[:37] + "..."
        return f"{codigo} | {resto}" if resto else codigo
    if len(t) > 50:
        t = t[:47] + "..."
    return t


def _extraer_codigo_y_edificio(texto_sala: str, row=None) -> tuple:
    """
    Extrae código de aula y edificio desde el texto de la fila Sala.
    Ej: "Sala SALÓN DE CLASES - [B-106A-101]" -> ("B-106A-101", "B-106")
    También intenta columnas extra si existen (formatos antiguos).
    """
    codigo = None
    edificio = None
    capacidad = None

    # Formato preferido: código entre corchetes
    m = re.search(r"\[([^\]]+)\]", texto_sala or "")
    if m:
        codigo = m.group(1).strip()

    # Columnas extra (algunos Excel antiguos traen código en col K, edificio en J)
    if row is not None and len(row) > 10:
        if row[10] and not codigo:
            codigo = str(row[10]).strip()
        if row[9]:
            edificio = str(row[9]).strip()
        if len(row) > 11 and row[11] is not None:
            try:
                capacidad = int(row[11])
            except (ValueError, TypeError):
                capacidad = None

    if not codigo:
        return None, None, None

    # Edificio a partir del código: B-106A-101 -> B-106 ; B-106T-302 -> B-106
    if not edificio:
        m2 = re.match(r"^(B-\d+)", codigo, re.IGNORECASE)
        if m2:
            edificio = m2.group(1).upper()
        else:
            # fallback: parte antes del último guion con letras
            partes = codigo.split("-")
            if len(partes) >= 2:
                edificio = f"{partes[0]}-{partes[1][:3]}" if partes[1] else partes[0]

    return codigo, edificio, capacidad


def parsear_aulas_desde_excel(ruta_archivo: str) -> List[Dict[str, Any]]:
    """
    Lee el Excel y devuelve lista de aulas con:
    codigo, nombre, edificio, capacidad, horario, dias_habilitados
    """
    wb = load_workbook(ruta_archivo, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    aulas: List[Dict] = []
    codigos_vistos = set()

    sala_actual = None
    leyendo_horas = False
    horario_actual: Dict[str, list] = {d: [] for d in DIAS}

    def guardar_sala():
        nonlocal sala_actual, horario_actual, leyendo_horas
        if sala_actual and sala_actual["codigo"] not in codigos_vistos:
            dias_habilitados = [d for d in DIAS if horario_actual[d]]
            sala_actual["horario"] = {d: list(horario_actual[d]) for d in DIAS}
            sala_actual["dias_habilitados"] = dias_habilitados
            aulas.append(sala_actual)
            codigos_vistos.add(sala_actual["codigo"])
        sala_actual = None
        horario_actual = {d: [] for d in DIAS}
        leyendo_horas = False

    for row in ws.iter_rows(min_row=1, max_col=12, values_only=True):
        celda_a = row[0] if row else None

        # Convertir a texto si es necesario
        texto_a = str(celda_a).strip() if celda_a is not None else ""

        # Nueva sala
        if texto_a.startswith("Sala "):
            if sala_actual:
                guardar_sala()

            codigo, edificio, capacidad = _extraer_codigo_y_edificio(texto_a, row)
            if not codigo:
                sala_actual = None
                leyendo_horas = False
                continue

            # Nombre legible: quitar prefijo Sala y corchetes si se desea
            nombre = texto_a
            sala_actual = {
                "codigo": codigo,
                "nombre": nombre,
                "edificio": edificio,
                "capacidad": capacidad,
            }
            horario_actual = {d: [] for d in DIAS}
            leyendo_horas = False
            continue

        # Encabezado de días
        if texto_a.lower() == "horas":
            if sala_actual:
                leyendo_horas = True
            continue

        # Franjas horarias
        if sala_actual and leyendo_horas and texto_a and re.match(r"^\d{1,2}:\d{2}", texto_a):
            hora = texto_a
            for i, dia in enumerate(DIAS):
                if i + 1 >= len(row):
                    break
                valor = row[i + 1]
                if valor and str(valor).strip():
                    clase = _limpiar_texto_clase(valor)
                    if clase:
                        horario_actual[dia].append({
                            "hora": hora,
                            "clase": clase,
                        })

    if sala_actual:
        guardar_sala()

    wb.close()
    return aulas
