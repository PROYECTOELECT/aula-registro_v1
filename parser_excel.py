"""
Parser del archivo de mapa horario por Sala.
Extrae las aulas y su horario semanal (días y franjas con clases).
"""
from openpyxl import load_workbook
from typing import List, Dict, Any
import json
import re


DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


def _limpiar_texto_clase(texto: str) -> str:
    """Resume el texto de la celda de clase para mostrar."""
    if not texto:
        return ""
    t = str(texto).replace("\n", " ").strip()
    # Tomar primera línea / código + nombre corto
    if " | " in t:
        partes = t.split(" | ")
        codigo = partes[0].strip()
        resto = partes[1].strip() if len(partes) > 1 else ""
        # Acortar
        if len(resto) > 40:
            resto = resto[:37] + "..."
        return f"{codigo} | {resto}" if resto else codigo
    if len(t) > 50:
        t = t[:47] + "..."
    return t


def parsear_aulas_desde_excel(ruta_archivo: str) -> List[Dict[str, Any]]:
    """
    Lee el Excel y devuelve lista de aulas con:
    codigo, nombre, edificio, capacidad, horario
    horario = {
      "Lunes": [{"hora": "7:00:00 - 7:15:00", "clase": "..."}, ...],
      ...
    }
    """
    wb = load_workbook(ruta_archivo, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    aulas: List[Dict] = []
    codigos_vistos = set()

    # Estado actual del bloque de sala
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
        celda_a = row[0]

        if not celda_a or not isinstance(celda_a, str):
            # Si estamos en un bloque y la fila tiene datos de días, igual procesar
            if sala_actual and leyendo_horas:
                pass
            else:
                continue

        celda_a = str(celda_a).strip()

        # Nueva sala
        if celda_a.startswith("Sala "):
            # Guardar la anterior
            if sala_actual:
                guardar_sala()

            codigo = row[10]
            edificio = row[9]
            capacidad = row[11]

            if not codigo:
                sala_actual = None
                continue

            codigo = str(codigo).strip()
            try:
                capacidad = int(capacidad) if capacidad is not None else None
            except (ValueError, TypeError):
                capacidad = None

            sala_actual = {
                "codigo": codigo,
                "nombre": celda_a,
                "edificio": str(edificio).strip() if edificio else None,
                "capacidad": capacidad,
            }
            horario_actual = {d: [] for d in DIAS}
            leyendo_horas = False
            continue

        # Fila de encabezado de días
        if celda_a == "Horas" or celda_a.lower() == "horas":
            leyendo_horas = True
            continue

        # Filas de franjas horarias (empiezan con hora)
        if sala_actual and leyendo_horas and re.match(r"^\d{1,2}:\d{2}", celda_a):
            hora = celda_a
            for i, dia in enumerate(DIAS):
                valor = row[i + 1]  # columnas B-H = índices 1-7
                if valor and str(valor).strip():
                    clase = _limpiar_texto_clase(valor)
                    if clase:
                        horario_actual[dia].append({
                            "hora": hora,
                            "clase": clase,
                        })

    # Última sala
    if sala_actual:
        guardar_sala()

    wb.close()
    return aulas
