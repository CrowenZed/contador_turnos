#!/usr/bin/env python3
"""Prepara VACACIONES.xlsx para cálculos automáticos de horas en Excel.

Uso:
    python preparar_vacaciones.py

Archivos esperados en el directorio actual:
    - VACACIONES.xlsx
    - Turnos.xlsx

Salida:
    - VACACIONES_preparado.xlsx
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
import re
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


# Zona editable: festivos de ejemplo que se rellenan automáticamente en la hoja FESTIVOS
# (la forma principal de mantenimiento es editar FESTIVOS dentro del propio Excel).
FESTIVOS_INICIALES = [
    # "2026-06-24",  # Sant Joan (ejemplo)
]

VACACIONES_FILE = Path("VACACIONES.xlsx")
TURNOS_FILE = Path("Turnos.xlsx")
OUTPUT_FILE = Path("VACACIONES_preparado.xlsx")


@dataclass
class TurnosMapping:
    turno: int
    manana_laborable: int
    tarde_laborable: int
    manana_sabado: int
    tarde_sabado: int
    manana_festivo: int
    tarde_festivo: int


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("ñ", "n")
    text = re.sub(r"\s+", " ", text)
    return text


def is_date_like(value: object) -> bool:
    if isinstance(value, (datetime, date)):
        return True
    if isinstance(value, (int, float)):
        return False
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return False
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
            try:
                datetime.strptime(value, fmt)
                return True
            except ValueError:
                continue
    return False


def detect_main_sheet(wb) -> Worksheet:
    for ws in wb.worksheets:
        if ws.title.upper() not in {"HORAS_TURNOS", "HORAS_TURNOS_RAW", "FESTIVOS"}:
            return ws
    return wb.worksheets[0]


def detect_date_header_row(ws: Worksheet, scan_rows: int = 20) -> int:
    best_row = 1
    best_count = 0
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        count = 0
        for c in range(1, ws.max_column + 1):
            if is_date_like(ws.cell(row=r, column=c).value):
                count += 1
        if count > best_count:
            best_count = count
            best_row = r
    if best_count == 0:
        raise ValueError("No se ha podido detectar una fila de cabecera con fechas en VACACIONES.xlsx")
    return best_row


def detect_date_columns(ws: Worksheet, header_row: int) -> List[int]:
    cols = []
    for c in range(1, ws.max_column + 1):
        if is_date_like(ws.cell(row=header_row, column=c).value):
            cols.append(c)
    if not cols:
        raise ValueError("No se han detectado columnas de fecha en la hoja principal")
    return cols


def clear_or_create_sheet(wb, name: str) -> Worksheet:
    if name in wb.sheetnames:
        ws = wb[name]
        wb.remove(ws)
    return wb.create_sheet(name)


def copy_turnos_sheet_to_aux(wb_vac, wb_turnos) -> Worksheet:
    src = wb_turnos[wb_turnos.sheetnames[0]]
    aux = clear_or_create_sheet(wb_vac, "HORAS_TURNOS_RAW")

    for r in range(1, src.max_row + 1):
        for c in range(1, src.max_column + 1):
            s = src.cell(row=r, column=c)
            d = aux.cell(row=r, column=c, value=s.value)
            if s.has_style:
                d.font = copy(s.font)
                d.fill = copy(s.fill)
                d.border = copy(s.border)
                d.alignment = copy(s.alignment)
                d.number_format = s.number_format
                d.protection = copy(s.protection)
        aux.row_dimensions[r].height = src.row_dimensions[r].height

    for key, dim in src.column_dimensions.items():
        aux.column_dimensions[key].width = dim.width

    return aux


def detect_turnos_mapping(aux_ws: Worksheet) -> TurnosMapping:
    header_row = 1
    best = (0, 1)
    for r in range(1, min(15, aux_ws.max_row) + 1):
        non_empty = sum(1 for c in range(1, aux_ws.max_column + 1) if aux_ws.cell(r, c).value not in (None, ""))
        if non_empty > best[0]:
            best = (non_empty, r)
    header_row = best[1]

    headers: Dict[str, int] = {}
    for c in range(1, aux_ws.max_column + 1):
        headers[normalize_header(aux_ws.cell(header_row, c).value)] = c

    def find_col(*keywords: str) -> Optional[int]:
        for h, idx in headers.items():
            if all(k in h for k in keywords):
                return idx
        return None

    turno = find_col("turno") or find_col("codigo") or 1

    m_lab = find_col("manana", "labor") or find_col("mañana", "labor")
    t_lab = find_col("tarde", "labor")
    m_sab = find_col("manana", "sab") or find_col("mañana", "sab")
    t_sab = find_col("tarde", "sab")
    m_fes = find_col("manana", "fest") or find_col("mañana", "fest")
    t_fes = find_col("tarde", "fest")

    if not all([m_lab, t_lab, m_sab, t_sab, m_fes, t_fes]):
        raise ValueError(
            "No se han encontrado columnas necesarias en Turnos.xlsx. "
            "Se esperaban cabeceras que incluyan: turno, mañana laborable, tarde laborable, "
            "mañana sábado, tarde sábado, mañana festivo, tarde festivo."
        )

    std = clear_or_create_sheet(aux_ws.parent, "HORAS_TURNOS")
    std.append([
        "Turno",
        "Manana_Laborable",
        "Tarde_Laborable",
        "Manana_Sabado",
        "Tarde_Sabado",
        "Manana_Festivo",
        "Tarde_Festivo",
    ])

    for r in range(header_row + 1, aux_ws.max_row + 1):
        turno_value = aux_ws.cell(r, turno).value
        if turno_value in (None, ""):
            continue
        std.append([
            turno_value,
            aux_ws.cell(r, m_lab).value,
            aux_ws.cell(r, t_lab).value,
            aux_ws.cell(r, m_sab).value,
            aux_ws.cell(r, t_sab).value,
            aux_ws.cell(r, m_fes).value,
            aux_ws.cell(r, t_fes).value,
        ])

    for c in range(1, 8):
        std.column_dimensions[get_column_letter(c)].width = 20

    return TurnosMapping(1, 2, 3, 4, 5, 6, 7)


def ensure_festivos_sheet(wb) -> Worksheet:
    ws = clear_or_create_sheet(wb, "FESTIVOS")
    ws["A1"] = "Fecha"
    ws.column_dimensions["A"].width = 16

    for i, raw in enumerate(FESTIVOS_INICIALES, start=2):
        ws.cell(i, 1, datetime.strptime(raw, "%Y-%m-%d"))
        ws.cell(i, 1).number_format = "dd/mm/yyyy"

    ws.freeze_panes = "A2"
    return ws


def col_ref(c: int) -> str:
    return get_column_letter(c)


def build_day_formula(row: int, day_col: int, date_header_row: int, hours_col_index: int) -> str:
    day_cell = f"{col_ref(day_col)}{row}"
    date_cell = f"{col_ref(day_col)}${date_header_row}"
    return (
        f"IFERROR(INDEX(HORAS_TURNOS!$B:$G,"
        f"MATCH({day_cell},HORAS_TURNOS!$A:$A,0),"
        f"IF(OR(WEEKDAY({date_cell},2)=7,COUNTIF(FESTIVOS!$A:$A,{date_cell})>0),{hours_col_index},"
        f"IF(WEEKDAY({date_cell},2)=6,{hours_col_index-2},{hours_col_index-4}))),0)"
    )


def apply_weekly_columns(ws: Worksheet, date_header_row: int, date_cols: Sequence[int]) -> Tuple[List[int], List[int]]:
    week_starts = list(range(0, len(date_cols), 7))
    inserts: List[Tuple[int, int]] = []
    for i in week_starts:
        chunk = date_cols[i : i + 7]
        if not chunk:
            continue
        inserts.append((chunk[-1] + 1, i // 7 + 1))

    morning_cols = []
    afternoon_cols = []

    for insert_at, week_number in reversed(inserts):
        ws.insert_cols(insert_at, amount=2)
        ws.cell(date_header_row, insert_at, f"Horas mañana S{week_number}")
        ws.cell(date_header_row, insert_at + 1, f"Horas tarde S{week_number}")

        if insert_at > 1:
            for rr in range(1, ws.max_row + 1):
                origin = ws.cell(rr, insert_at - 1)
                target = ws.cell(rr, insert_at)
                target2 = ws.cell(rr, insert_at + 1)
                if origin.has_style:
                    target._style = copy(origin._style)
                    target2._style = copy(origin._style)

    updated_date_cols = detect_date_columns(ws, date_header_row)
    for i in range(0, len(updated_date_cols), 7):
        chunk = updated_date_cols[i : i + 7]
        if not chunk:
            continue
        morning_cols.append(chunk[-1] + 1)
        afternoon_cols.append(chunk[-1] + 2)

    hour_format = "[h]:mm"
    for row in range(date_header_row + 1, ws.max_row + 1):
        row_has_data = any(ws.cell(row, c).value not in (None, "") for c in updated_date_cols)
        if not row_has_data:
            continue

        for week_idx, start_idx in enumerate(range(0, len(updated_date_cols), 7)):
            chunk = updated_date_cols[start_idx : start_idx + 7]
            if not chunk:
                continue

            morning_col = morning_cols[week_idx]
            afternoon_col = afternoon_cols[week_idx]

            morning_terms = [build_day_formula(row, day_col, date_header_row, 5) for day_col in chunk]
            afternoon_terms = [build_day_formula(row, day_col, date_header_row, 6) for day_col in chunk]

            ws.cell(row, morning_col, f"={' + '.join(morning_terms)}")
            ws.cell(row, afternoon_col, f"={' + '.join(afternoon_terms)}")
            ws.cell(row, morning_col).number_format = hour_format
            ws.cell(row, afternoon_col).number_format = hour_format

    return morning_cols, afternoon_cols


def append_final_totals(ws: Worksheet, date_header_row: int, morning_cols: Sequence[int], afternoon_cols: Sequence[int]) -> None:
    start = ws.max_column + 1
    ws.cell(date_header_row, start, "Total mañana")
    ws.cell(date_header_row, start + 1, "Total tarde")
    ws.cell(date_header_row, start + 2, "Total global")

    for r in range(date_header_row + 1, ws.max_row + 1):
        if not any(ws.cell(r, c).value not in (None, "") for c in list(morning_cols) + list(afternoon_cols)):
            continue

        man_refs = ",".join(f"{col_ref(c)}{r}" for c in morning_cols) or "0"
        tar_refs = ",".join(f"{col_ref(c)}{r}" for c in afternoon_cols) or "0"

        ws.cell(r, start, f"=SUM({man_refs})")
        ws.cell(r, start + 1, f"=SUM({tar_refs})")
        ws.cell(r, start + 2, f"={col_ref(start)}{r}+{col_ref(start+1)}{r}")

        ws.cell(r, start).number_format = "[h]:mm"
        ws.cell(r, start + 1).number_format = "[h]:mm"
        ws.cell(r, start + 2).number_format = "[h]:mm"


def main() -> None:
    if not VACACIONES_FILE.exists():
        raise FileNotFoundError(f"No se encontró {VACACIONES_FILE}")
    if not TURNOS_FILE.exists():
        raise FileNotFoundError(f"No se encontró {TURNOS_FILE}")

    wb_vac = load_workbook(VACACIONES_FILE)
    wb_turnos = load_workbook(TURNOS_FILE, data_only=True)

    main_ws = detect_main_sheet(wb_vac)
    date_header_row = detect_date_header_row(main_ws)
    date_cols = detect_date_columns(main_ws, date_header_row)

    raw_aux = copy_turnos_sheet_to_aux(wb_vac, wb_turnos)
    detect_turnos_mapping(raw_aux)

    wb_vac["HORAS_TURNOS"].sheet_state = "hidden"
    wb_vac["HORAS_TURNOS_RAW"].sheet_state = "hidden"
    ensure_festivos_sheet(wb_vac)

    morning_cols, afternoon_cols = apply_weekly_columns(main_ws, date_header_row, date_cols)
    append_final_totals(main_ws, date_header_row, morning_cols, afternoon_cols)

    wb_vac.save(OUTPUT_FILE)
    print(f"Archivo generado: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
