import os
import datetime
import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx")

# ==============================================================================
# 1. CONFIGURACIÓN MAESTRA DEL SISTEMA
# ==============================================================================

UNITS = [
    ("Proteína Gym", "Cronograma General", "1E3A8A", "DBEAFE"),
    ("Proteína Box", "Horarios Box", "065F46", "D1FAE5"),
    ("Boxeo", "Escuela y Entrenamiento de Boxeo", "7C2D12", "FFEDD5"),
]

# Lista única de profesores (Sección 5 + datos reales de la planilla adjunta)
PROFESORES = [
    # (Nombre, Estado, Color Fondo Hex, Color Texto Hex, Unidades Habituales, Rol / Referencia)
    ("Agustín", "Activo", "D0E2FF", "002D9C", "Proteína Gym", "Profesor Sala"),
    ("Robert", "Activo", "DEFBE6", "0E6027", "Proteína Gym, Proteína Box, Boxeo", "Profesor Sala / Ventas"),
    ("Silvina", "Activo", "D9FBFB", "005D5D", "Proteína Gym", "Ventas / Atención (Rob/Sil)"),
    ("Valentina", "Activo", "FFD6E8", "9F1853", "Proteína Gym, Boxeo", "Ventas / Boxeo"),
    ("Jessi", "Activo", "FFF2D6", "8E5A00", "Proteína Gym", "Ventas / Recepción"),
    ("Luna", "Activo", "E8DAFF", "491D8B", "Proteína Gym", "Profesora Sala / Fin de Semana"),
    ("Irina", "Activo", "FFD7D9", "A2191F", "Proteína Gym", "Profesora Sala / Ventas"),
    ("Cecilia", "Activo", "E0E7FF", "3730A3", "Proteína Gym, Proteína Box", "Coordinadora / Profesora"),
    ("Oscar", "Activo", "ECFCCB", "3F6212", "Proteína Gym", "Profesor Turno Noche"),
    ("Yami", "Activo", "CFFAFE", "155E75", "Proteína Box", "Profesora Proteína Box"),
    ("Lucas", "Activo", "F3E8FF", "6B21A8", "Proteína Box", "Profesor Proteína Box"),
    ("Fede", "Activo", "FFEDD5", "9A3412", "Proteína Box, Boxeo", "Profesor Proteína Box / Boxeo"),
    ("Martín", "Activo", "CCFBF1", "115E59", "Boxeo", "Profesor Titular Boxeo"),
]

MAX_PROF_ROWS = 35  # Filas reservadas en PROFES (filas 5 a 39) para altas automáticas

# Período de trabajo: Noviembre 2026 a Diciembre 2027 (14 meses - Sección 6.1)
MONTHS = [
    ("NOVIEMBRE 2026", 2026, 11),
    ("DICIEMBRE 2026", 2026, 12),
    ("ENERO 2027", 2027, 1),
    ("FEBRERO 2027", 2027, 2),
    ("MARZO 2027", 2027, 3),
    ("ABRIL 2027", 2027, 4),
    ("MAYO 2027", 2027, 5),
    ("JUNIO 2027", 2027, 6),
    ("JULIO 2027", 2027, 7),
    ("AGOSTO 2027", 2027, 8),
    ("SEPTIEMBRE 2027", 2027, 9),
    ("OCTUBRE 2027", 2027, 10),
    ("NOVIEMBRE 2027", 2027, 11),
    ("DICIEMBRE 2027", 2027, 12),
]

# Franjas horarias de 30 minutos (07:00 a 23:00 -> 32 franjas - Sección 6.3)
TIME_SLOTS = []
for h in range(7, 23):
    TIME_SLOTS.append(f"{h:02d}:00 a {h:02d}:30")
    TIME_SLOTS.append(f"{h:02d}:30 a {(h+1):02d}:00")

DAYS_ES = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]

# Calendario oficial de Feriados Nacionales (Nov 2026 - Dic 2027 - Sección 9.3)
FERIADOS_LIST = [
    (datetime.date(2026, 11, 23), "Día de la Soberanía Nacional (Trasladado)", "Feriado Nacional"),
    (datetime.date(2026, 12, 8), "Inmaculada Concepción de María", "Feriado Inamovible"),
    (datetime.date(2026, 12, 25), "Navidad", "Feriado Inamovible"),
    (datetime.date(2027, 1, 1), "Año Nuevo", "Feriado Inamovible"),
    (datetime.date(2027, 2, 8), "Carnaval (Lunes)", "Feriado Inamovible"),
    (datetime.date(2027, 2, 9), "Carnaval (Martes)", "Feriado Inamovible"),
    (datetime.date(2027, 3, 24), "Día Nacional de la Memoria por la Verdad y la Justicia", "Feriado Inamovible"),
    (datetime.date(2027, 3, 26), "Viernes Santo", "Feriado Inamovible"),
    (datetime.date(2027, 4, 2), "Día del Veterano y de los Caídos en la Guerra de Malvinas", "Feriado Inamovible"),
    (datetime.date(2027, 5, 1), "Día del Trabajador", "Feriado Inamovible"),
    (datetime.date(2027, 5, 25), "Día de la Revolución de Mayo", "Feriado Inamovible"),
    (datetime.date(2027, 6, 17), "Paso a la Inmortalidad del Gral. Martín Miguel de Güemes", "Feriado Trasladable"),
    (datetime.date(2027, 6, 20), "Paso a la Inmortalidad del Gral. Manuel Belgrano (Dom+Feriado)", "Feriado Inamovible"),
    (datetime.date(2027, 7, 9), "Día de la Independencia", "Feriado Inamovible"),
    (datetime.date(2027, 8, 16), "Paso a la Inmortalidad del Gral. José de San Martín", "Feriado Trasladable"),
    (datetime.date(2027, 10, 11), "Día del Respeto a la Diversidad Cultural", "Feriado Trasladable"),
    (datetime.date(2027, 11, 22), "Día de la Soberanía Nacional", "Feriado Trasladable"),
    (datetime.date(2027, 12, 8), "Inmaculada Concepción de María", "Feriado Inamovible"),
    (datetime.date(2027, 12, 25), "Navidad", "Feriado Inamovible"),
]

# ==============================================================================
# 2. PLANTILLA SEMANAL REAL (EXTRAÍDA DE LA PLANILLA ADJUNTA + BOXEO)
# ==============================================================================

def get_weekly_assignment(unit_name, weekday_idx, slot_idx):
    """
    Devuelve la asignación por defecto (string con nombres separados por coma)
    según la unidad, el día de la semana (0=Lunes .. 6=Domingo) y el índice de franja (0..31).
    slot_idx 0 = 07:00-07:30, 1 = 07:30-08:00, ..., 31 = 22:30-23:00.
    """
    if unit_name == "Proteína Gym":
        if weekday_idx == 0:  # LUNES
            if 0 <= slot_idx <= 9:    return "Robert, Silvina"
            if 10 <= slot_idx <= 19:  return "Agustín, Jessi"
            if 20 <= slot_idx <= 27:  return "Cecilia, Luna, Valentina"
            if 28 <= slot_idx <= 29:  return "Cecilia, Oscar, Valentina"
            if 30 <= slot_idx <= 31:  return "Oscar"
        elif weekday_idx == 1:  # MARTES
            if 0 <= slot_idx <= 9:    return "Agustín, Valentina"
            if 10 <= slot_idx <= 15:  return "Agustín, Jessi"
            if 16 <= slot_idx <= 19:  return "Cecilia, Irina, Jessi"
            if 20 <= slot_idx <= 27:  return "Cecilia, Irina, Robert, Silvina"
            if 28 <= slot_idx <= 29:  return "Cecilia, Oscar, Robert, Silvina"
            if 30 <= slot_idx <= 31:  return "Oscar"
        elif weekday_idx == 2:  # MIÉRCOLES
            if 0 <= slot_idx <= 9:    return "Robert, Jessi"
            if 10 <= slot_idx <= 11:  return "Robert, Silvina"
            if 12 <= slot_idx <= 19:  return "Agustín, Robert, Silvina"
            if 20 <= slot_idx <= 21:  return "Agustín, Cecilia, Valentina"
            if 22 <= slot_idx <= 23:  return "Cecilia, Valentina"
            if 24 <= slot_idx <= 29:  return "Cecilia, Luna, Valentina"
            if 30 <= slot_idx <= 31:  return "Luna"
        elif weekday_idx == 3:  # JUEVES
            if 0 <= slot_idx <= 9:    return "Agustín, Robert, Silvina"
            if 10 <= slot_idx <= 15:  return "Agustín, Jessi"
            if 16 <= slot_idx <= 18:  return "Irina, Jessi"
            if slot_idx == 19:        return "Cecilia, Irina, Jessi"
            if 20 <= slot_idx <= 21:  return "Cecilia, Irina, Valentina"
            if 22 <= slot_idx <= 25:  return "Cecilia, Valentina"
            if 26 <= slot_idx <= 29:  return "Cecilia, Irina, Valentina"
            if 30 <= slot_idx <= 31:  return "Irina"
        elif weekday_idx == 4:  # VIERNES
            if 0 <= slot_idx <= 9:    return "Robert, Valentina"
            if 10 <= slot_idx <= 11:  return "Robert, Jessi"
            if 12 <= slot_idx <= 19:  return "Agustín, Jessi"
            if 20 <= slot_idx <= 21:  return "Agustín, Cecilia, Irina"
            if 22 <= slot_idx <= 29:  return "Cecilia, Luna, Irina"
            if 30 <= slot_idx <= 31:  return "Luna"
        elif weekday_idx == 5:  # SÁBADO
            if 6 <= slot_idx <= 17:   return "Luna, Irina"  # 10:00 a 16:00
        elif weekday_idx == 6:  # DOMINGO
            if 6 <= slot_idx <= 11:   return "Luna"         # 10:00 a 13:00 (Domingo x2)

    elif unit_name == "Proteína Box":
        if weekday_idx == 0:  # LUNES
            if 22 <= slot_idx <= 27:  return "Yami"         # 18:00 a 21:00
        elif weekday_idx == 1:  # MARTES
            if 1 <= slot_idx <= 4:    return "Yami"         # 07:30 a 09:30
            if 17 <= slot_idx <= 18:  return "Lucas"        # 15:30 a 16:30
            if 22 <= slot_idx <= 27:  return "Lucas"        # 18:00 a 21:00
        elif weekday_idx == 2:  # MIÉRCOLES
            if 1 <= slot_idx <= 4:    return "Yami"         # 07:30 a 09:30
            if 17 <= slot_idx <= 18:  return "Fede"         # 15:30 a 16:30
            if 22 <= slot_idx <= 27:  return "Yami"         # 18:00 a 21:00
        elif weekday_idx == 3:  # JUEVES
            if 1 <= slot_idx <= 4:    return "Yami"         # 07:30 a 09:30
            if 17 <= slot_idx <= 18:  return "Lucas"        # 15:30 a 16:30
            if 22 <= slot_idx <= 27:  return "Lucas"        # 18:00 a 21:00
        elif weekday_idx == 4:  # VIERNES
            if 1 <= slot_idx <= 4:    return "Yami"         # 07:30 a 09:30
            if 17 <= slot_idx <= 18:  return "Fede"         # 15:30 a 16:30
            if 22 <= slot_idx <= 27:  return "Yami"         # 18:00 a 21:00

    elif unit_name == "Boxeo":
        if weekday_idx in (0, 2, 4):  # LUNES, MIÉRCOLES, VIERNES
            if 5 <= slot_idx <= 7:    return "Valentina"    # 09:30 a 11:00
            if 24 <= slot_idx <= 27:  return "Martín"       # 19:00 a 21:00
        elif weekday_idx in (1, 3):   # MARTES, JUEVES
            if 6 <= slot_idx <= 8:    return "Martín"       # 10:00 a 11:30
            if 25 <= slot_idx <= 28:  return "Fede"         # 19:30 a 21:30
        elif weekday_idx == 5:        # SÁBADO
            if 8 <= slot_idx <= 11:   return "Martín"       # 11:00 a 13:00

    return ""


# ==============================================================================
# 3. ESTILOS VISUALES (SECCIÓN 15: VERDE=EDITABLE, GRIS=PROTEGIDO, AZUL=CONFIG)
# ==============================================================================

FONT_FAMILY = "Segoe UI"

f_title = Font(name=FONT_FAMILY, size=15, bold=True, color="FFFFFF")
f_subtitle = Font(name=FONT_FAMILY, size=11, bold=True, color="1E293B")
f_sec_hdr = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
f_tbl_hdr = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
f_sub_hdr = Font(name=FONT_FAMILY, size=9, bold=True, color="1E293B")
f_data = Font(name=FONT_FAMILY, size=10, bold=False, color="0F172A")
f_data_bold = Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A")
f_small = Font(name=FONT_FAMILY, size=9, bold=False, color="334155")
f_small_bold = Font(name=FONT_FAMILY, size=9, bold=True, color="0F172A")
f_link = Font(name=FONT_FAMILY, size=10, bold=True, color="1D4ED8", underline="single")
f_double_badge = Font(name=FONT_FAMILY, size=8, bold=True, color="991B1B")
f_normal_badge = Font(name=FONT_FAMILY, size=8, bold=False, color="475569")

# Fills según Sección 15.2:
# Verde: celdas editables por el usuario
# Gris: información automática o protegida
# Azul: información de configuración
fill_navy = PatternFill("solid", fgColor="0F172A")
fill_slate_hdr = PatternFill("solid", fgColor="1E293B")
fill_unit_gym = PatternFill("solid", fgColor="1E3A8A")
fill_unit_box = PatternFill("solid", fgColor="065F46")
fill_unit_boxeo = PatternFill("solid", fgColor="7C2D12")

fill_editable_green = PatternFill("solid", fgColor="ECFDF5")   # Verde suave editable
fill_editable_strong = PatternFill("solid", fgColor="D1FAE5")  # Verde destacado selector
fill_auto_gray = PatternFill("solid", fgColor="F1F5F9")        # Gris automático/protegido
fill_total_gray = PatternFill("solid", fgColor="E2E8F0")       # Gris totales
fill_config_blue = PatternFill("solid", fgColor="DBEAFE")      # Azul configuración
fill_double_day = PatternFill("solid", fgColor="FEE2E2")       # Rojo suave para Domingo/Feriado (x2)
fill_normal_day = PatternFill("solid", fgColor="F8FAFC")       # Blanco/Gris muy claro día normal
fill_white = PatternFill("solid", fgColor="FFFFFF")
fill_warn_amber = PatternFill("solid", fgColor="FEF3C7")

thin_side = Side(style="thin", color="CBD5E1")
med_navy = Side(style="medium", color="1E293B")
double_bottom = Side(style="double", color="0F172A")

border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_total = Border(top=thin_side, bottom=double_bottom, left=thin_side, right=thin_side)
border_box = Border(left=med_navy, right=med_navy, top=med_navy, bottom=med_navy)

al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
al_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

CURRENCY_FMT = '$#,##0'
HOURS_FMT = '0.0'
DATE_FMT = 'dd/mm/yyyy'
SHORT_DATE_FMT = 'dd/mm'


def style_range(ws, min_col, min_row, max_col, max_row, font=None, fill=None, border=None, alignment=None, num_format=None):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if alignment: cell.alignment = alignment
            if num_format: cell.number_format = num_format


def add_banner(ws, title, subtitle, max_col=12):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c1 = ws.cell(1, 1, title)
    c1.font = f_title
    c1.fill = fill_navy
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2 = ws.cell(2, 1, subtitle)
    c2.font = Font(name=FONT_FAMILY, size=10, italic=True, color="E2E8F0")
    c2.fill = fill_slate_hdr
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 10


# ==============================================================================
# 4. CONSTRUCCIÓN DE CADA PESTAÑA
# ==============================================================================

def build_workbook():
    wb = openpyxl.Workbook()
    # Renombrar hoja inicial a INICIO
    ws_inicio = wb.active
    ws_inicio.title = "INICIO"

    # Crear las pestañas en el orden exacto de la especificación (Sección 3)
    ws_profes = wb.create_sheet("PROFES")
    ws_config = wb.create_sheet("CONFIGURACIÓN")
    ws_feriados = wb.create_sheet("FERIADOS")
    ws_plantilla = wb.create_sheet("PLANTILLA SEMANAL")

    month_sheets = {}
    for m_name, yr, mo in MONTHS:
        month_sheets[m_name] = wb.create_sheet(m_name)

    ws_res_mensual = wb.create_sheet("RESUMEN MENSUAL")
    ws_res_anual = wb.create_sheet("RESUMEN ANUAL")
    ws_rep_unidad = wb.create_sheet("REPORTES POR UNIDAD")
    ws_reg_horas = wb.create_sheet("REGISTRO DE HORAS")

    # Colores de pestañas para navegación visual rápida
    ws_inicio.sheet_properties.tabColor = "0F172A"
    ws_profes.sheet_properties.tabColor = "2563EB"
    ws_config.sheet_properties.tabColor = "0284C7"
    ws_feriados.sheet_properties.tabColor = "DC2626"
    ws_plantilla.sheet_properties.tabColor = "4F46E5"
    for m_name in month_sheets:
        month_sheets[m_name].sheet_properties.tabColor = "10B981"
    ws_res_mensual.sheet_properties.tabColor = "F59E0B"
    ws_res_anual.sheet_properties.tabColor = "D97706"
    ws_rep_unidad.sheet_properties.tabColor = "7C3AED"
    ws_reg_horas.sheet_properties.tabColor = "64748B"

    # --------------------------------------------------------------------------
    # PESTAÑA: PROFES (Sección 5)
    # --------------------------------------------------------------------------
    add_banner(ws_profes, "ADMINISTRACIÓN CENTRALIZADA DE PROFESORES (STAFF ÚNICO)",
               "Proteína Gym · Proteína Box · Boxeo | Celdas Verdes = Editables | Celdas Grises = Cálculo Automático", 10)

    headers_profes = [
        "Nombre del Profesor", "Estado", "Color Identificador", "Código Color",
        "Unidades Habituales", "Rol / Función", "Horas Acum. Período", "Total Liquidado Período",
        "Observaciones", "Lista Dinámica Activos (Automática)"
    ]
    ws_profes.row_dimensions[4].height = 26
    for col_idx, text in enumerate(headers_profes, 1):
        cell = ws_profes.cell(4, col_idx, text)
        cell.font = f_tbl_hdr
        cell.fill = fill_slate_hdr
        cell.alignment = al_center
        cell.border = border_cell

    # Validación de Estado: Activo / Inactivo
    dv_estado = DataValidation(type="list", formula1='"Activo,Inactivo"', allow_blank=False)
    ws_profes.add_data_validation(dv_estado)
    dv_estado.add(f"B5:B{4+MAX_PROF_ROWS}")

    for idx in range(MAX_PROF_ROWS):
        r = 5 + idx
        ws_profes.row_dimensions[r].height = 22
        if idx < len(PROFESORES):
            nombre, estado, bg_hex, fg_hex, unidades_hab, rol = PROFESORES[idx]
            c_nom = ws_profes.cell(r, 1, nombre)
            c_est = ws_profes.cell(r, 2, estado)
            c_col = ws_profes.cell(r, 3, nombre)
            c_hex = ws_profes.cell(r, 4, f"#{bg_hex}")
            c_uni = ws_profes.cell(r, 5, unidades_hab)
            c_rol = ws_profes.cell(r, 6, rol)
            c_obs = ws_profes.cell(r, 9, "Staff inicial activo")

            c_col.fill = PatternFill("solid", fgColor=bg_hex)
            c_col.font = Font(name=FONT_FAMILY, size=10, bold=True, color=fg_hex)
        else:
            c_nom = ws_profes.cell(r, 1, "")
            c_est = ws_profes.cell(r, 2, "")
            c_col = ws_profes.cell(r, 3, "")
            c_hex = ws_profes.cell(r, 4, "")
            c_uni = ws_profes.cell(r, 5, "")
            c_rol = ws_profes.cell(r, 6, "")
            c_obs = ws_profes.cell(r, 9, "")
            c_col.fill = fill_editable_green
            c_col.font = f_data

        for col_i in (1, 2, 5, 6, 9):
            c = ws_profes.cell(r, col_i)
            c.fill = fill_editable_green
            c.font = f_data_bold if col_i == 1 else f_data
            c.border = border_cell
            c.alignment = al_left if col_i in (1, 5, 6, 9) else al_center

        c_col.alignment = al_center
        c_col.border = border_cell
        c_hex.fill = fill_config_blue
        c_hex.font = f_small
        c_hex.alignment = al_center
        c_hex.border = border_cell

        # Horas Acumuladas y Total Liquidado en el período (referencia a RESUMEN ANUAL)
        c_hs = ws_profes.cell(r, 7, f'=IF(A{r}="","",\'RESUMEN ANUAL\'!E{r+3})')
        c_hs.fill = fill_auto_gray
        c_hs.font = f_data_bold
        c_hs.alignment = al_right
        c_hs.border = border_cell
        c_hs.number_format = HOURS_FMT

        c_pay = ws_profes.cell(r, 8, f'=IF(A{r}="","",\'RESUMEN ANUAL\'!F{r+3})')
        c_pay.fill = fill_auto_gray
        c_pay.font = f_data_bold
        c_pay.alignment = al_right
        c_pay.border = border_cell
        c_pay.number_format = CURRENCY_FMT

        # Columna J (10): Lista limpia de profesores activos para alimentar las listas desplegables (R10, R11)
        c_act = ws_profes.cell(r, 10, f'=IF(AND(A{r}<>"",B{r}="Activo"),A{r},"")')
        c_act.fill = fill_auto_gray
        c_act.font = f_small
        c_act.alignment = al_center
        c_act.border = border_cell

    prof_widths = [22, 14, 20, 14, 34, 26, 18, 20, 28, 24]
    for i, w in enumerate(prof_widths, 1):
        ws_profes.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: CONFIGURACIÓN (Sección 8 y Cierres Mensuales Sección 11.5)
    # --------------------------------------------------------------------------
    add_banner(ws_config, "CONFIGURACIÓN GENERAL, VALOR HORA, VIGENCIAS Y CIERRES MENSUALES",
               "Tarifa única para todos los profesores (R09) | Domingos y Feriados pagan doble ×2 (R06-R08)", 9)

    # Bloque 1: Parámetros Principales Vigentes
    ws_config.merge_cells("A4:D4")
    ws_config.cell(4, 1, "1. TARIFA HORARIA VIGENTE (APLICABLE A TODAS LAS UNIDADES)").font = f_sec_hdr
    ws_config.cell(4, 1).fill = fill_unit_gym
    ws_config.cell(4, 1).alignment = al_left
    ws_config.row_dimensions[4].height = 24

    params = [
        ("Valor de la Hora Normal (Lunes a Sábado)", 5000, fill_config_blue, CURRENCY_FMT, "Editable: modifica los meses ABIERTOS"),
        ("Valor de la Hora Doble (Domingos y Feriados ×2)", "=B6*2", fill_auto_gray, CURRENCY_FMT, "Cálculo automático (×2, nunca ×4)"),
        ("Duración de cada Franja Horaria (en horas)", 0.5, fill_config_blue, HOURS_FMT, "Equivalente a bloques de 30 minutos (R01)"),
        ("Cantidad de Profesores Activos", f'=COUNTIF(PROFES!B5:B{4+MAX_PROF_ROWS},"Activo")', fill_auto_gray, '0', "Actualización automática desde PROFES"),
    ]
    for idx, (label, val, fl, numf, note) in enumerate(params, start=6):
        ws_config.row_dimensions[idx].height = 22
        c_lbl = ws_config.cell(idx, 1, label)
        c_lbl.font = f_data_bold
        c_lbl.fill = fill_auto_gray
        c_lbl.border = border_cell
        c_val = ws_config.cell(idx, 2, val)
        c_val.font = Font(name=FONT_FAMILY, size=11, bold=True, color="0F172A")
        c_val.fill = fl
        c_val.border = border_cell
        c_val.alignment = al_right
        c_val.number_format = numf
        ws_config.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=4)
        c_not = ws_config.cell(idx, 3, note)
        c_not.font = f_small
        c_not.fill = fill_white
        c_not.border = border_cell

    # Bloque 2: Histórico de Vigencia de Tarifas (Sección 8.2)
    ws_config.merge_cells("F4:I4")
    ws_config.cell(4, 6, "2. HISTÓRICO DE VIGENCIA DE TARIFAS (SECCIÓN 8.2)").font = f_sec_hdr
    ws_config.cell(4, 6).fill = fill_slate_hdr
    ws_config.cell(4, 6).alignment = al_left

    hist_headers = ["Vigencia Desde", "Vigencia Hasta", "Valor Hora Normal", "Valor Hora Doble (×2)"]
    for idx, h in enumerate(hist_headers, 6):
        c = ws_config.cell(5, idx, h)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    ws_config.cell(6, 6, datetime.date(2026, 11, 1)).number_format = DATE_FMT
    ws_config.cell(6, 7, datetime.date(2027, 12, 31)).number_format = DATE_FMT
    ws_config.cell(6, 8, "=B6").number_format = CURRENCY_FMT
    ws_config.cell(6, 9, "=H6*2").number_format = CURRENCY_FMT
    for col_i in range(6, 10):
        c = ws_config.cell(6, col_i)
        c.fill = fill_editable_green if col_i < 9 else fill_auto_gray
        c.font = f_data
        c.border = border_cell
        c.alignment = al_center if col_i < 8 else al_right

    for r_empty in range(7, 10):
        for col_i in range(6, 10):
            c = ws_config.cell(r_empty, col_i, "" if col_i < 9 else f'=IF(H{r_empty}="","",H{r_empty}*2)')
            c.fill = fill_editable_green if col_i < 9 else fill_auto_gray
            c.font = f_data
            c.border = border_cell
            c.number_format = DATE_FMT if col_i < 8 else CURRENCY_FMT

    # Bloque 3: Control de Períodos Mensuales y Cierre de Liquidaciones (Sección 11.5 - Regla R13)
    ws_config.merge_cells("A12:I12")
    c_b3 = ws_config.cell(12, 1, "3. CONTROL DE PERÍODOS MENSUALES, TARIFA APLICABLE Y CIERRE DE LIQUIDACIONES (R13)")
    c_b3.font = f_sec_hdr
    c_b3.fill = fill_navy
    c_b3.alignment = al_left
    ws_config.row_dimensions[12].height = 24

    ctrl_headers = [
        "Período Mensual", "Estado Liquidación", "Valor Hora al Cerrar",
        "Valor Hora Aplicable", "Valor Doble Aplicable", "Total Horas Mes",
        "Costo Total Mes", "Fecha de Cierre", "Observaciones / Autorización"
    ]
    ws_config.row_dimensions[13].height = 24
    for idx, h in enumerate(ctrl_headers, 1):
        c = ws_config.cell(13, idx, h)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    dv_cierre = DataValidation(type="list", formula1='"ABIERTO,LIQUIDADO"', allow_blank=False)
    ws_config.add_data_validation(dv_cierre)
    dv_cierre.add("B14:B27")

    for idx, (m_name, yr, mo) in enumerate(MONTHS):
        r = 14 + idx
        ws_config.row_dimensions[r].height = 21
        c_per = ws_config.cell(r, 1, m_name)
        c_per.font = f_data_bold
        c_per.fill = fill_auto_gray
        c_per.alignment = al_left
        c_per.border = border_cell

        c_est = ws_config.cell(r, 2, "ABIERTO")
        c_est.font = f_data_bold
        c_est.fill = fill_editable_green
        c_est.alignment = al_center
        c_est.border = border_cell

        c_cong = ws_config.cell(r, 3, 5000)
        c_cong.font = f_data
        c_cong.fill = fill_config_blue
        c_cong.alignment = al_right
        c_cong.border = border_cell
        c_cong.number_format = CURRENCY_FMT

        # Si el mes está LIQUIDADO, conserva la tarifa de cierre (Col C); si está ABIERTO, usa la tarifa vigente ($B$6)
        c_aplic = ws_config.cell(r, 4, f'=IF(AND(B{r}="LIQUIDADO",C{r}>0),C{r},$B$6)')
        c_aplic.font = f_data_bold
        c_aplic.fill = fill_auto_gray
        c_aplic.alignment = al_right
        c_aplic.border = border_cell
        c_aplic.number_format = CURRENCY_FMT

        c_dob = ws_config.cell(r, 5, f'=D{r}*2')
        c_dob.font = f_data
        c_dob.fill = fill_auto_gray
        c_dob.alignment = al_right
        c_dob.border = border_cell
        c_dob.number_format = CURRENCY_FMT

        # Fila 151 de cada hoja mensual contiene los totales del mes
        c_ths = ws_config.cell(r, 6, f"='{m_name}'!N151")
        c_ths.font = f_data_bold
        c_ths.fill = fill_auto_gray
        c_ths.alignment = al_right
        c_ths.border = border_cell
        c_ths.number_format = HOURS_FMT

        c_tco = ws_config.cell(r, 7, f"='{m_name}'!Q151")
        c_tco.font = f_data_bold
        c_tco.fill = fill_auto_gray
        c_tco.alignment = al_right
        c_tco.border = border_cell
        c_tco.number_format = CURRENCY_FMT

        c_fec = ws_config.cell(r, 8, "")
        c_fec.font = f_small
        c_fec.fill = fill_editable_green
        c_fec.alignment = al_center
        c_fec.border = border_cell
        c_fec.number_format = DATE_FMT

        c_obs = ws_config.cell(r, 9, "Período activo para carga" if idx < 2 else "Período planificado")
        c_obs.font = f_small
        c_obs.fill = fill_editable_green
        c_obs.alignment = al_left
        c_obs.border = border_cell

    cfg_widths = [26, 18, 20, 20, 20, 18, 20, 18, 30]
    for i, w in enumerate(cfg_widths, 1):
        ws_config.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: FERIADOS (Sección 9.3)
    # --------------------------------------------------------------------------
    add_banner(ws_feriados, "CALENDARIO DE FERIADOS Y FECHAS ESPECIALES (REMUNERACIÓN DOBLE ×2)",
               "Agregue cualquier feriado en las filas verdes. El sistema lo detecta automáticamente en los 14 meses y en las 3 unidades.", 6)

    fer_headers = ["Fecha", "Día de la Semana", "Descripción del Feriado", "Tipo", "Multiplicador", "Período"]
    ws_feriados.row_dimensions[4].height = 25
    for idx, h in enumerate(fer_headers, 1):
        c = ws_feriados.cell(4, idx, h)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    for idx in range(45):
        r = 5 + idx
        ws_feriados.row_dimensions[r].height = 21
        if idx < len(FERIADOS_LIST):
            f_date, f_desc, f_tipo = FERIADOS_LIST[idx]
            c_dt = ws_feriados.cell(r, 1, f_date)
            c_ds = ws_feriados.cell(r, 2, DAYS_ES[f_date.weekday()])
            c_de = ws_feriados.cell(r, 3, f_desc)
            c_ti = ws_feriados.cell(r, 4, f_tipo)
        else:
            c_dt = ws_feriados.cell(r, 1, "")
            c_ds = ws_feriados.cell(r, 2, f'=IF(A{r}="","",CHOOSE(WEEKDAY(A{r},2),"LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO","DOMINGO"))')
            c_de = ws_feriados.cell(r, 3, "")
            c_ti = ws_feriados.cell(r, 4, "")

        c_dt.number_format = DATE_FMT
        c_dt.fill = fill_editable_green
        c_dt.font = f_data_bold
        c_dt.alignment = al_center
        c_dt.border = border_cell

        c_ds.fill = fill_auto_gray
        c_ds.font = f_small_bold
        c_ds.alignment = al_center
        c_ds.border = border_cell

        c_de.fill = fill_editable_green
        c_de.font = f_data
        c_de.alignment = al_left
        c_de.border = border_cell

        c_ti.fill = fill_editable_green
        c_ti.font = f_small
        c_ti.alignment = al_center
        c_ti.border = border_cell

        c_mul = ws_feriados.cell(r, 5, f'=IF(A{r}="","","×2 (Doble)")')
        c_mul.fill = fill_double_day if idx < len(FERIADOS_LIST) else fill_auto_gray
        c_mul.font = f_double_badge
        c_mul.alignment = al_center
        c_mul.border = border_cell

        c_per = ws_feriados.cell(r, 6, f'=IF(A{r}="","",TEXT(A{r},"mm/yyyy"))')
        c_per.fill = fill_auto_gray
        c_per.font = f_small
        c_per.alignment = al_center
        c_per.border = border_cell

    for i, w in enumerate([16, 18, 48, 22, 16, 16], 1):
        ws_feriados.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: PLANTILLA SEMANAL (Horario Base de las 3 Unidades)
    # --------------------------------------------------------------------------
    add_banner(ws_plantilla, "PLANTILLA SEMANAL BASE — PROTEÍNA GYM · PROTEÍNA BOX · BOXEO",
               "Contiene la estructura habitual de Lunes a Domingo de la planilla original + Boxeo para pre-cargar cualquier mes.", 8)

    unit_fills = [fill_unit_gym, fill_unit_box, fill_unit_boxeo]
    curr_row = 4
    for u_idx, (u_name, u_sub, _, _) in enumerate(UNITS):
        ws_plantilla.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=8)
        c_uh = ws_plantilla.cell(curr_row, 1, f"SECCIÓN {u_idx+1}: {u_name.upper()} — {u_sub.upper()}")
        c_uh.font = f_sec_hdr
        c_uh.fill = unit_fills[u_idx]
        c_uh.alignment = al_left
        ws_plantilla.row_dimensions[curr_row].height = 24
        curr_row += 1

        ws_plantilla.cell(curr_row, 1, "HORARIO").font = f_tbl_hdr
        ws_plantilla.cell(curr_row, 1).fill = fill_slate_hdr
        ws_plantilla.cell(curr_row, 1).alignment = al_center
        ws_plantilla.cell(curr_row, 1).border = border_cell
        for d_idx, d_name in enumerate(DAYS_ES):
            c_d = ws_plantilla.cell(curr_row, 2 + d_idx, d_name)
            c_d.font = f_tbl_hdr
            c_d.fill = fill_slate_hdr
            c_d.alignment = al_center
            c_d.border = border_cell
        ws_plantilla.row_dimensions[curr_row].height = 22
        curr_row += 1

        for s_idx, slot_str in enumerate(TIME_SLOTS):
            ws_plantilla.row_dimensions[curr_row].height = 19
            c_tm = ws_plantilla.cell(curr_row, 1, slot_str)
            c_tm.font = f_small_bold
            c_tm.fill = fill_auto_gray
            c_tm.alignment = al_center
            c_tm.border = border_cell
            for d_idx in range(7):
                val = get_weekly_assignment(u_name, d_idx, s_idx)
                c_v = ws_plantilla.cell(curr_row, 2 + d_idx, val)
                c_v.font = f_small
                c_v.fill = fill_editable_green
                c_v.alignment = al_center
                c_v.border = border_cell
            curr_row += 1
        curr_row += 2

    ws_plantilla.column_dimensions["A"].width = 16
    for col_i in range(2, 9):
        ws_plantilla.column_dimensions[get_column_letter(col_i)].width = 25

    # --------------------------------------------------------------------------
    # PESTAÑAS MENSUALES: NOVIEMBRE 2026 A DICIEMBRE 2027 (Secciones 6, 7, 9, 14)
    # --------------------------------------------------------------------------
    # En cada pestaña mensual:
    # - Fila 4: DÍA DE LA SEMANA (DOMINGO, LUNES...)
    # - Fila 5: FECHA REAL (01/11/2026 .. 30/11/2026)
    # - Fila 6: TIPO DE JORNADA AUTOMÁTICO ("DOBLE (x2)" si es Domingo o está en FERIADOS, sino "NORMAL (x1)")
    # - Filas 8..40:   UNIDAD 1: PROTEÍNA GYM (Encabezado en fila 8, 32 franjas en filas 9..40)
    # - Filas 43..75:  UNIDAD 2: PROTEÍNA BOX (Encabezado en fila 43, 32 franjas en filas 44..75)
    # - Filas 78..110: UNIDAD 3: BOXEO (Encabezado en fila 78, 32 franjas en filas 79..110)
    # - Filas 114..151: TABLA INTERNA DE LIQUIDACIÓN MENSUAL POR PROFESOR Y UNIDAD (Filas 116..150 = 35 profes, Fila 151 = TOTAL)

    feriados_set = {f[0] for f in FERIADOS_LIST}
    registros_bd = []  # Para poblar REGISTRO DE HORAS con trazabilidad completa (Sección 12)
    reg_id_counter = 1

    for m_idx, (m_name, yr, mo) in enumerate(MONTHS):
        ws_m = month_sheets[m_name]
        num_days = calendar.monthrange(yr, mo)[1]
        last_day_col = get_column_letter(1 + num_days)
        cfg_row = 14 + m_idx  # Fila de este mes en CONFIGURACIÓN

        add_banner(
            ws_m,
            f"CRONOGRAMA MENSUAL — {m_name} | PROTEÍNA GYM · PROTEÍNA BOX · BOXEO",
            "Cada celda = 0,5 hs (30 min) por cada profesor asignado (R01-R03). Para múltiples profes o reemplazos, separar con coma (ej: Robert, Agustín).",
            1 + num_days
        )

        # Encabezados de fila 4, 5, 6
        ws_m.cell(4, 1, "DÍA").font = f_tbl_hdr
        ws_m.cell(4, 1).fill = fill_slate_hdr
        ws_m.cell(4, 1).alignment = al_center
        ws_m.cell(4, 1).border = border_cell

        ws_m.cell(5, 1, "FECHA").font = f_tbl_hdr
        ws_m.cell(5, 1).fill = fill_slate_hdr
        ws_m.cell(5, 1).alignment = al_center
        ws_m.cell(5, 1).border = border_cell

        ws_m.cell(6, 1, "TARIFA DÍA").font = f_tbl_hdr
        ws_m.cell(6, 1).fill = fill_slate_hdr
        ws_m.cell(6, 1).alignment = al_center
        ws_m.cell(6, 1).border = border_cell

        ws_m.row_dimensions[4].height = 20
        ws_m.row_dimensions[5].height = 20
        ws_m.row_dimensions[6].height = 20

        ws_m.column_dimensions["A"].width = 18

        for d in range(1, num_days + 1):
            col_i = 1 + d
            col_let = get_column_letter(col_i)
            ws_m.column_dimensions[col_let].width = 19
            dt = datetime.date(yr, mo, d)
            wk_idx = dt.weekday()  # 0=Lunes .. 6=Domingo
            is_double = (wk_idx == 6) or (dt in feriados_set)

            c_day = ws_m.cell(4, col_i, DAYS_ES[wk_idx])
            c_day.font = f_tbl_hdr
            c_day.fill = fill_unit_boxeo if is_double else fill_slate_hdr
            c_day.alignment = al_center
            c_day.border = border_cell

            c_dt = ws_m.cell(5, col_i, dt)
            c_dt.font = f_sub_hdr
            c_dt.fill = fill_double_day if is_double else fill_total_gray
            c_dt.alignment = al_center
            c_dt.border = border_cell
            c_dt.number_format = SHORT_DATE_FMT

            # Fórmula que detecta automáticamente Domingo o Feriado (R05, R06, R07, R08)
            c_jor = ws_m.cell(
                6, col_i,
                f'=IF(OR(WEEKDAY({col_let}$5,2)=7,COUNTIF(FERIADOS!$A$5:$A$49,{col_let}$5)>0),"DOBLE (x2)","NORMAL (x1)")'
            )
            c_jor.font = f_double_badge if is_double else f_normal_badge
            c_jor.fill = fill_double_day if is_double else fill_auto_gray
            c_jor.alignment = al_center
            c_jor.border = border_cell

        # Bloques de las 3 Unidades
        unit_start_rows = [
            ("Proteína Gym", "SECCIÓN 1: PROTEÍNA GYM (CRONOGRAMA GENERAL)", 8, fill_unit_gym),
            ("Proteína Box", "SECCIÓN 2: PROTEÍNA BOX (HORARIOS BOX)", 43, fill_unit_box),
            ("Boxeo", "SECCIÓN 3: BOXEO (ENTRENAMIENTO Y ESCUELA DE BOXEO)", 78, fill_unit_boxeo),
        ]

        # Validación de datos desplegable desde la lista dinámica de profesores activos
        dv_prof = DataValidation(
            type="list",
            formula1=f"=PROFES!$J$5:$J${4+MAX_PROF_ROWS}",
            allow_blank=True
        )
        dv_prof.promptTitle = "Asignación de Profesor"
        dv_prof.prompt = "Seleccione un profesor activo o escriba varios separados por coma (ej: Robert, Agustín)."
        # En Excel/Google Sheets dejar showErrorMessage=False permite escribir múltiples nombres separados por coma
        dv_prof.showErrorMessage = False
        ws_m.add_data_validation(dv_prof)

        # Pre-cargamos datos completos en NOVIEMBRE 2026 y DICIEMBRE 2026 (y primera semana de ENERO 2027)
        preload_month = m_name in ("NOVIEMBRE 2026", "DICIEMBRE 2026")

        for u_name, u_banner_text, hdr_row, u_fill in unit_start_rows:
            ws_m.merge_cells(start_row=hdr_row, start_column=1, end_row=hdr_row, end_column=1 + num_days)
            c_uh = ws_m.cell(hdr_row, 1, f"{u_banner_text} — {m_name}")
            c_uh.font = f_sec_hdr
            c_uh.fill = u_fill
            c_uh.alignment = al_left
            ws_m.row_dimensions[hdr_row].height = 24

            first_slot_row = hdr_row + 1
            last_slot_row = hdr_row + 32
            dv_prof.add(f"B{first_slot_row}:{last_day_col}{last_slot_row}")

            for s_idx, slot_str in enumerate(TIME_SLOTS):
                r = first_slot_row + s_idx
                ws_m.row_dimensions[r].height = 18
                c_tm = ws_m.cell(r, 1, slot_str)
                c_tm.font = f_small_bold
                c_tm.fill = fill_auto_gray
                c_tm.alignment = al_center
                c_tm.border = border_cell

                for d in range(1, num_days + 1):
                    col_i = 1 + d
                    dt = datetime.date(yr, mo, d)
                    wk_idx = dt.weekday()
                    is_double = (wk_idx == 6) or (dt in feriados_set)

                    val = ""
                    tipo_asig = "Titular"
                    if preload_month:
                        val = get_weekly_assignment(u_name, wk_idx, s_idx)
                        # Incluir casos de reemplazo explícitos (Sección 7.1: Lunes 02/11 18:00 en Proteína Box: Robert + Agustín)
                        if dt == datetime.date(2026, 11, 2) and u_name == "Proteína Box" and s_idx in (22, 23):
                            val = "Robert, Agustín"
                            tipo_asig = "Titular + Reemplazo"
                        elif dt == datetime.date(2026, 12, 1) and u_name == "Proteína Box" and s_idx == 22:
                            val = "Robert, Agustín"
                            tipo_asig = "Titular + Reemplazo"

                    c_cell = ws_m.cell(r, col_i, val)
                    c_cell.font = f_small
                    c_cell.fill = fill_double_day if is_double else fill_editable_green
                    c_cell.alignment = al_center
                    c_cell.border = border_cell

                    # Si hay asignación precargada, generamos los registros individuales normalizados para REGISTRO DE HORAS (Sección 12)
                    if val:
                        h_ini, h_fin = slot_str.split(" a ")
                        nombres = [n.strip() for n in val.split(",") if n.strip()]
                        for idx_n, nom_prof in enumerate(nombres):
                            rol_asig = "Reemplazo" if (tipo_asig == "Titular + Reemplazo" and idx_n == len(nombres) - 1) else "Titular"
                            registros_bd.append({
                                "id": f"ASIG-{reg_id_counter:05d}",
                                "mes": m_name,
                                "fecha": dt,
                                "dia": DAYS_ES[wk_idx],
                                "h_ini": h_ini,
                                "h_fin": h_fin,
                                "duracion": 0.5,
                                "profesor": nom_prof,
                                "unidad": u_name,
                                "seccion": "General" if u_name == "Proteína Gym" else ("Box" if u_name == "Proteína Box" else "Boxeo"),
                                "tipo_asig": rol_asig,
                                "jornada": "Doble (×2)" if is_double else "Normal (×1)",
                                "mult": 2 if is_double else 1,
                                "cfg_row": cfg_row,
                            })
                            reg_id_counter += 1

        # ----------------------------------------------------------------------
        # BLOQUE 4 (FILAS 114 A 151): RESUMEN AUTOMÁTICO DEL MES POR PROFESOR Y UNIDAD
        # ----------------------------------------------------------------------
        ws_m.merge_cells(f"A114:Q114")
        c_res_hdr = ws_m.cell(114, 1, f"CONTROL Y LIQUIDACIÓN AUTOMÁTICA DEL MES — {m_name} (CÁLCULO PROTEGIDO EN TIEMPO REAL)")
        c_res_hdr.font = f_sec_hdr
        c_res_hdr.fill = fill_navy
        c_res_hdr.alignment = al_left
        ws_m.row_dimensions[114].height = 24

        calc_headers = [
            "Profesor", "Estado",
            "Gym Hs Norm", "Gym Hs Dob", "Gym Total Hs",
            "Box Hs Norm", "Box Hs Dob", "Box Total Hs",
            "Boxeo Hs Norm", "Boxeo Hs Dob", "Boxeo Total Hs",
            "Total Hs Normales", "Total Hs Dobles", "Total Horas Mes",
            "Importe Normal", "Importe Doble", "Total a Liquidar ($)"
        ]
        ws_m.row_dimensions[115].height = 24
        for idx_c, h_text in enumerate(calc_headers, 1):
            c = ws_m.cell(115, idx_c, h_text)
            c.font = f_tbl_hdr
            c.fill = fill_slate_hdr
            c.alignment = al_center
            c.border = border_cell

        # Columnas auxiliares (columnas S en adelante) para calcular Horas Dobles día por día sin fórmulas matriciales CSE
        # Para cada día d (1..num_days):
        # - Columna 18 + d: Horas dobles del día d en Proteína Gym
        # - Columna 50 + d: Horas dobles del día d en Proteína Box
        # - Columna 82 + d: Horas dobles del día d en Boxeo
        gym_dbl_start_col = 19
        box_dbl_start_col = 51
        boxeo_dbl_start_col = 83

        for idx_p in range(MAX_PROF_ROWS):
            r = 116 + idx_p
            p_row = 5 + idx_p
            ws_m.row_dimensions[r].height = 19

            # A: Nombre del Profesor (vinculado dinámicamente a PROFES)
            c_pnom = ws_m.cell(r, 1, f'=IF(PROFES!A{p_row}="","",PROFES!A{p_row})')
            c_pnom.font = f_data_bold
            c_pnom.fill = fill_auto_gray
            c_pnom.border = border_cell

            # B: Estado
            c_pest = ws_m.cell(r, 2, f'=IF(A{r}="","",PROFES!B{p_row})')
            c_pest.font = f_small
            c_pest.fill = fill_auto_gray
            c_pest.alignment = al_center
            c_pest.border = border_cell

            # Celdas auxiliares por día para contar horas en días DOBLES (Domingos y Feriados)
            for d in range(1, num_days + 1):
                day_col_let = get_column_letter(1 + d)
                # Gym (filas 9..40)
                c_gd = ws_m.cell(
                    r, gym_dbl_start_col + d - 1,
                    f'=IF($A{r}="",0,IF({day_col_let}$6="DOBLE (x2)",COUNTIF({day_col_let}$9:{day_col_let}$40,"*"&$A{r}&"*")*0.5,0))'
                )
                # Box (filas 44..75)
                c_bd = ws_m.cell(
                    r, box_dbl_start_col + d - 1,
                    f'=IF($A{r}="",0,IF({day_col_let}$6="DOBLE (x2)",COUNTIF({day_col_let}$44:{day_col_let}$75,"*"&$A{r}&"*")*0.5,0))'
                )
                # Boxeo (filas 79..110)
                c_xd = ws_m.cell(
                    r, boxeo_dbl_start_col + d - 1,
                    f'=IF($A{r}="",0,IF({day_col_let}$6="DOBLE (x2)",COUNTIF({day_col_let}$79:{day_col_let}$110,"*"&$A{r}&"*")*0.5,0))'
                )

            g_first, g_last = get_column_letter(gym_dbl_start_col), get_column_letter(gym_dbl_start_col + num_days - 1)
            b_first, b_last = get_column_letter(box_dbl_start_col), get_column_letter(box_dbl_start_col + num_days - 1)
            x_first, x_last = get_column_letter(boxeo_dbl_start_col), get_column_letter(boxeo_dbl_start_col + num_days - 1)

            # Proteína Gym: C=Normales, D=Dobles, E=Total
            ws_m.cell(r, 5, f'=IF($A{r}="",0,COUNTIF($B$9:${last_day_col}$40,"*"&$A{r}&"*")*0.5)')
            ws_m.cell(r, 4, f'=SUM({g_first}{r}:{g_last}{r})')
            ws_m.cell(r, 3, f'=E{r}-D{r}')

            # Proteína Box: F=Normales, G=Dobles, H=Total
            ws_m.cell(r, 8, f'=IF($A{r}="",0,COUNTIF($B$44:${last_day_col}$75,"*"&$A{r}&"*")*0.5)')
            ws_m.cell(r, 7, f'=SUM({b_first}{r}:{b_last}{r})')
            ws_m.cell(r, 6, f'=H{r}-G{r}')

            # Boxeo: I=Normales, J=Dobles, K=Total
            ws_m.cell(r, 11, f'=IF($A{r}="",0,COUNTIF($B$79:${last_day_col}$110,"*"&$A{r}&"*")*0.5)')
            ws_m.cell(r, 10, f'=SUM({x_first}{r}:{x_last}{r})')
            ws_m.cell(r, 9, f'=K{r}-J{r}')

            # Consolidado del Profesor en el Mes: L=Normales, M=Dobles, N=Total Horas
            ws_m.cell(r, 12, f'=C{r}+F{r}+I{r}')
            ws_m.cell(r, 13, f'=D{r}+G{r}+J{r}')
            ws_m.cell(r, 14, f'=L{r}+M{r}')

            # Importes en $ usando la tarifa aplicable del período en CONFIGURACIÓN!D{cfg_row}
            ws_m.cell(r, 15, f'=L{r}*CONFIGURACIÓN!$D${cfg_row}')
            ws_m.cell(r, 16, f'=M{r}*CONFIGURACIÓN!$E${cfg_row}')
            ws_m.cell(r, 17, f'=O{r}+P{r}')

            for col_c in range(3, 18):
                cell = ws_m.cell(r, col_c)
                cell.font = f_data_bold if col_c in (5, 8, 11, 14, 17) else f_data
                cell.fill = fill_auto_gray
                cell.alignment = al_right
                cell.border = border_cell
                cell.number_format = CURRENCY_FMT if col_c >= 15 else HOURS_FMT

        # Ocultar columnas auxiliares de cálculo diario (S en adelante)
        for col_hide in range(gym_dbl_start_col, boxeo_dbl_start_col + 32):
            ws_m.column_dimensions[get_column_letter(col_hide)].hidden = True

        # Fila 151: TOTALES DEL MES
        ws_m.row_dimensions[151].height = 23
        ws_m.cell(151, 1, "TOTALES DEL MES").font = f_data_bold
        ws_m.cell(151, 1).fill = fill_total_gray
        ws_m.cell(151, 1).border = border_total
        ws_m.cell(151, 2, f'=COUNTIF(N116:N150,">0") & " profes"').font = f_small_bold
        ws_m.cell(151, 2).fill = fill_total_gray
        ws_m.cell(151, 2).alignment = al_center
        ws_m.cell(151, 2).border = border_total

        for col_c in range(3, 18):
            col_let = get_column_letter(col_c)
            c_tot = ws_m.cell(151, col_c, f'=SUM({col_let}116:{col_let}150)')
            c_tot.font = f_data_bold
            c_tot.fill = fill_total_gray
            c_tot.alignment = al_right
            c_tot.border = border_total
            c_tot.number_format = CURRENCY_FMT if col_c >= 15 else HOURS_FMT

        # Inmovilizar paneles para ver siempre la columna de horarios y las fechas al desplazar
        ws_m.freeze_panes = "B7"

    # --------------------------------------------------------------------------
    # PESTAÑA: RESUMEN MENSUAL (Sección 10)
    # --------------------------------------------------------------------------
    add_banner(ws_res_mensual, "RESUMEN MENSUAL DE HORAS, LIQUIDACIONES Y COSTOS POR UNIDAD",
               "Seleccione el mes en la celda verde C4. Consolida automáticamente las 3 unidades sin duplicar pagos (R14-R16).", 12)

    # Selector desplegable de mes (Sección 10.1)
    ws_res_mensual.row_dimensions[4].height = 28
    ws_res_mensual.merge_cells("A4:B4")
    c_sel_lbl = ws_res_mensual.cell(4, 1, "SELECCIONAR PERÍODO A CONSULTAR:")
    c_sel_lbl.font = f_data_bold
    c_sel_lbl.fill = fill_slate_hdr
    c_sel_lbl.font = f_tbl_hdr
    c_sel_lbl.alignment = al_center

    c_sel_val = ws_res_mensual.cell(4, 3, "NOVIEMBRE 2026")
    c_sel_val.font = Font(name=FONT_FAMILY, size=12, bold=True, color="065F46")
    c_sel_val.fill = fill_editable_strong
    c_sel_val.alignment = al_center
    c_sel_val.border = border_box

    dv_mes_sel = DataValidation(type="list", formula1="=CONFIGURACIÓN!$A$14:$A$27", allow_blank=False)
    ws_res_mensual.add_data_validation(dv_mes_sel)
    dv_mes_sel.add("C4")

    # KPIs Superiores del Mes Seleccionado (Sección 10.1)
    kpis_mensual = [
        (5, "Estado del Período", '=INDEX(CONFIGURACIÓN!$B$14:$B$27,MATCH($C$4,CONFIGURACIÓN!$A$14:$A$27,0))', '@'),
        (7, "Valor Hora Aplicable", '=INDEX(CONFIGURACIÓN!$D$14:$D$27,MATCH($C$4,CONFIGURACIÓN!$A$14:$A$27,0))', CURRENCY_FMT),
        (9, "Total Horas del Mes", '=H46', HOURS_FMT),
        (11, "Total General a Liquidar", '=L46', CURRENCY_FMT),
    ]
    for col_k, k_lbl, k_form, k_fmt in kpis_mensual:
        ws_res_mensual.cell(4, col_k, k_lbl).font = f_small_bold
        ws_res_mensual.cell(4, col_k).fill = fill_auto_gray
        ws_res_mensual.cell(4, col_k).alignment = al_right
        ws_res_mensual.cell(4, col_k).border = border_cell

        c_kv = ws_res_mensual.cell(4, col_k + 1, k_form)
        c_kv.font = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
        c_kv.fill = fill_config_blue
        c_kv.alignment = al_center if k_fmt == '@' else al_right
        c_kv.border = border_cell
        c_kv.number_format = k_fmt

    # Tabla 1: Liquidación General Consolidada por Profesor (Sección 10.2) + Desglose por Unidad
    ws_res_mensual.merge_cells("A6:L6")
    c_t1 = ws_res_mensual.cell(6, 1, "1. LIQUIDACIÓN GENERAL CONSOLIDADA POR PROFESOR (SUMA DE LAS 3 UNIDADES SIN DUPLICACIÓN)")
    c_t1.font = f_sec_hdr
    c_t1.fill = fill_navy
    c_t1.alignment = al_left
    ws_res_mensual.row_dimensions[6].height = 24

    headers_res_m = [
        "Profesor", "Estado",
        "Hs Proteína Gym", "Hs Proteína Box", "Hs Boxeo",
        "Horas Normales", "Horas Dobles (×2)", "Total de Horas",
        "Valor Hora", "Importe Normal", "Importe Doble", "Total a Pagar ($)"
    ]
    ws_res_mensual.row_dimensions[7].height = 25
    for idx_c, h_text in enumerate(headers_res_m, 1):
        c = ws_res_mensual.cell(7, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    # Función helper para elegir el valor de la hoja mensual seleccionada en $C$4 sin usar INDIRECT (100% compatible con Excel y Google Sheets)
    def choose_month_cell(cell_coord):
        terms = [f"'{m_name}'!{cell_coord}" for m_name, _, _ in MONTHS]
        return f"CHOOSE(MATCH($C$4,CONFIGURACIÓN!$A$14:$A$27,0),{','.join(terms)})"

    for idx_p in range(MAX_PROF_ROWS):
        r = 8 + idx_p
        m_row = 116 + idx_p
        p_row = 5 + idx_p
        ws_res_mensual.row_dimensions[r].height = 20

        ws_res_mensual.cell(r, 1, f'=IF(PROFES!A{p_row}="","",PROFES!A{p_row})')
        ws_res_mensual.cell(r, 2, f'=IF(A{r}="","",PROFES!B{p_row})')
        ws_res_mensual.cell(r, 3, f'=IF(A{r}="",0,{choose_month_cell(f"E{m_row}")})')   # Hs Gym
        ws_res_mensual.cell(r, 4, f'=IF(A{r}="",0,{choose_month_cell(f"H{m_row}")})')   # Hs Box
        ws_res_mensual.cell(r, 5, f'=IF(A{r}="",0,{choose_month_cell(f"K{m_row}")})')   # Hs Boxeo
        ws_res_mensual.cell(r, 6, f'=IF(A{r}="",0,{choose_month_cell(f"L{m_row}")})')   # Horas Normales
        ws_res_mensual.cell(r, 7, f'=IF(A{r}="",0,{choose_month_cell(f"M{m_row}")})')   # Horas Dobles
        ws_res_mensual.cell(r, 8, f'=IF(A{r}="",0,F{r}+G{r})')                          # Total Horas
        ws_res_mensual.cell(r, 9, f'=IF(A{r}="",0,$H$4)')                               # Valor Hora
        ws_res_mensual.cell(r, 10, f'=IF(A{r}="",0,F{r}*I{r})')                         # Importe Normal
        ws_res_mensual.cell(r, 11, f'=IF(A{r}="",0,G{r}*I{r}*2)')                       # Importe Doble
        ws_res_mensual.cell(r, 12, f'=IF(A{r}="",0,J{r}+K{r})')                         # Total a Pagar

        for col_c in range(1, 13):
            cell = ws_res_mensual.cell(r, col_c)
            cell.font = f_data_bold if col_c in (1, 8, 12) else f_data
            cell.fill = fill_editable_strong if col_c == 12 else fill_auto_gray
            cell.alignment = al_left if col_c == 1 else (al_center if col_c == 2 else al_right)
            cell.border = border_cell
            if col_c in (3, 4, 5, 6, 7, 8):
                cell.number_format = HOURS_FMT
            elif col_c in (9, 10, 11, 12):
                cell.number_format = CURRENCY_FMT

        if idx_p >= len(PROFESORES):
            ws_res_mensual.row_dimensions[r].hidden = True

    # Fila 46 (8 + MAX_PROF_ROWS + 3 -> ajustemos exactamente en fila 43 = 8 + 35)
    tot_row_m = 8 + MAX_PROF_ROWS  # Fila 43
    ws_res_mensual.row_dimensions[tot_row_m].height = 24
    ws_res_mensual.cell(tot_row_m, 1, "TOTAL GENERAL DEL STAFF").font = f_data_bold
    ws_res_mensual.cell(tot_row_m, 1).fill = fill_total_gray
    ws_res_mensual.cell(tot_row_m, 1).border = border_total
    ws_res_mensual.cell(tot_row_m, 2, f'=COUNTIF(H8:H{tot_row_m-1},">0") & " activos c/hs"').font = f_small_bold
    ws_res_mensual.cell(tot_row_m, 2).fill = fill_total_gray
    ws_res_mensual.cell(tot_row_m, 2).alignment = al_center
    ws_res_mensual.cell(tot_row_m, 2).border = border_total

    for col_c in range(3, 13):
        col_let = get_column_letter(col_c)
        c_t = ws_res_mensual.cell(
            tot_row_m, col_c,
            f'=AVERAGE(I8:I{tot_row_m-1})' if col_c == 9 else f'=SUM({col_let}8:{col_let}{tot_row_m-1})'
        )
        c_t.font = f_data_bold
        c_t.fill = fill_total_gray
        c_t.alignment = al_right
        c_t.border = border_total
        c_t.number_format = CURRENCY_FMT if col_c >= 9 else HOURS_FMT

    # Actualizar referencias de KPIs en fila 4 a tot_row_m
    ws_res_mensual.cell(4, 10, f'=H{tot_row_m}')
    ws_res_mensual.cell(4, 12, f'=L{tot_row_m}')

    # Tabla 2: Desglose de Liquidaciones y Costos por Unidad (Sección 10.3 y 10.5)
    u_tbl_row = tot_row_m + 2  # Fila 45
    ws_res_mensual.merge_cells(f"A{u_tbl_row}:H{u_tbl_row}")
    c_u_hdr = ws_res_mensual.cell(u_tbl_row, 1, "2. DESGLOSE DE HORAS Y COSTOS DE PERSONAL POR UNIDAD DE NEGOCIO (SECCIÓN 10.3)")
    c_u_hdr.font = f_sec_hdr
    c_u_hdr.fill = fill_unit_gym
    c_u_hdr.alignment = al_left
    ws_res_mensual.row_dimensions[u_tbl_row].height = 24

    u_cols = ["Unidad de Trabajo", "Horas Normales", "Horas Dobles (×2)", "Total de Horas", "Importe Normal", "Importe Doble", "Costo Total Personal", "% Participación"]
    ws_res_mensual.row_dimensions[u_tbl_row + 1].height = 23
    for idx_c, h_text in enumerate(u_cols, 1):
        c = ws_res_mensual.cell(u_tbl_row + 1, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    unit_m_coords = [
        ("Proteína Gym", "C151", "D151", "E151"),
        ("Proteína Box", "F151", "G151", "H151"),
        ("Boxeo", "I151", "J151", "K151"),
    ]
    tot_u_row = u_tbl_row + 5
    for idx_u, (u_name, c_norm, c_dob, c_tot) in enumerate(unit_m_coords):
        ru = u_tbl_row + 2 + idx_u
        ws_res_mensual.row_dimensions[ru].height = 21
        ws_res_mensual.cell(ru, 1, u_name).font = f_data_bold
        ws_res_mensual.cell(ru, 1).fill = fill_auto_gray
        ws_res_mensual.cell(ru, 1).border = border_cell

        ws_res_mensual.cell(ru, 2, f'={choose_month_cell(c_norm)}')
        ws_res_mensual.cell(ru, 3, f'={choose_month_cell(c_dob)}')
        ws_res_mensual.cell(ru, 4, f'=B{ru}+C{ru}')
        ws_res_mensual.cell(ru, 5, f'=B{ru}*$H$4')
        ws_res_mensual.cell(ru, 6, f'=C{ru}*$H$4*2')
        ws_res_mensual.cell(ru, 7, f'=E{ru}+F{ru}')
        ws_res_mensual.cell(ru, 8, f'=IF($G${tot_u_row}=0,0,G{ru}/$G${tot_u_row})')

        for col_c in range(2, 9):
            c = ws_res_mensual.cell(ru, col_c)
            c.font = f_data_bold if col_c in (4, 7) else f_data
            c.fill = fill_auto_gray
            c.alignment = al_right
            c.border = border_cell
            c.number_format = '0.0%' if col_c == 8 else (CURRENCY_FMT if col_c >= 5 else HOURS_FMT)

    ws_res_mensual.row_dimensions[tot_u_row].height = 23
    ws_res_mensual.cell(tot_u_row, 1, "COSTO CONSOLIDADO (3 UNIDADES)").font = f_data_bold
    ws_res_mensual.cell(tot_u_row, 1).fill = fill_total_gray
    ws_res_mensual.cell(tot_u_row, 1).border = border_total
    for col_c in range(2, 9):
        col_let = get_column_letter(col_c)
        c = ws_res_mensual.cell(tot_u_row, col_c, f'=SUM({col_let}{u_tbl_row+2}:{col_let}{tot_u_row-1})')
        c.font = f_data_bold
        c.fill = fill_total_gray
        c.alignment = al_right
        c.border = border_total
        c.number_format = '0.0%' if col_c == 8 else (CURRENCY_FMT if col_c >= 5 else HOURS_FMT)

    # Tabla 3: Consulta Individual por Profesor (Recibo / Detalle por Unidad - Sección 10.4)
    ind_row = tot_u_row + 2
    ws_res_mensual.merge_cells(f"A{ind_row}:G{ind_row}")
    c_ind_hdr = ws_res_mensual.cell(ind_row, 1, "3. DESGLOSE INDIVIDUAL POR PROFESOR (DETALLE POR UNIDAD PARA RECIBO / EXPORTACIÓN PDF)")
    c_ind_hdr.font = f_sec_hdr
    c_ind_hdr.fill = fill_unit_box
    c_ind_hdr.alignment = al_left
    ws_res_mensual.row_dimensions[ind_row].height = 24

    ws_res_mensual.cell(ind_row + 1, 1, "Seleccionar Profesor:").font = f_data_bold
    ws_res_mensual.cell(ind_row + 1, 1).fill = fill_auto_gray
    ws_res_mensual.cell(ind_row + 1, 1).border = border_cell
    c_psel = ws_res_mensual.cell(ind_row + 1, 2, "Robert")
    c_psel.font = Font(name=FONT_FAMILY, size=11, bold=True, color="0E6027")
    c_psel.fill = fill_editable_strong
    c_psel.alignment = al_center
    c_psel.border = border_box
    dv_prof_ind = DataValidation(type="list", formula1=f"=PROFES!$A$5:$A${4+MAX_PROF_ROWS}", allow_blank=False)
    ws_res_mensual.add_data_validation(dv_prof_ind)
    dv_prof_ind.add(f"B{ind_row+1}")

    ind_cols = ["Unidad", "Horas Normales", "Horas Dobles (×2)", "Total Horas", "Importe Normal", "Importe Doble", "Total a Pagar ($)"]
    ws_res_mensual.row_dimensions[ind_row + 2].height = 22
    for idx_c, h_text in enumerate(ind_cols, 1):
        c = ws_res_mensual.cell(ind_row + 2, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    for idx_u, (u_name, _, _, _) in enumerate(UNITS):
        ri = ind_row + 3 + idx_u
        ws_res_mensual.row_dimensions[ri].height = 20
        ws_res_mensual.cell(ri, 1, u_name).font = f_data_bold
        ws_res_mensual.cell(ri, 1).fill = fill_auto_gray
        ws_res_mensual.cell(ri, 1).border = border_cell

        # En REPORTES POR UNIDAD las filas 8..42 contienen el desglose del mes seleccionado por profesor y unidad
        # Pero podemos obtenerlo directamente de la fila del profesor en la hoja mensual o desde REPORTES POR UNIDAD!
        # En REPORTES POR UNIDAD:
        # Col C, D, E = Gym Norm, Dob, Tot | Col G, H, I = Box Norm, Dob, Tot | Col K, L, M = Boxeo Norm, Dob, Tot
        rep_norm_col = "C" if idx_u == 0 else ("G" if idx_u == 1 else "K")
        rep_dob_col = "D" if idx_u == 0 else ("H" if idx_u == 1 else "L")

        ws_res_mensual.cell(ri, 2, f'=IFERROR(INDEX(\'REPORTES POR UNIDAD\'!${rep_norm_col}$8:${rep_norm_col}${7+MAX_PROF_ROWS},MATCH($B${ind_row+1},\'REPORTES POR UNIDAD\'!$A$8:$A${7+MAX_PROF_ROWS},0)),0)')
        ws_res_mensual.cell(ri, 3, f'=IFERROR(INDEX(\'REPORTES POR UNIDAD\'!${rep_dob_col}$8:${rep_dob_col}${7+MAX_PROF_ROWS},MATCH($B${ind_row+1},\'REPORTES POR UNIDAD\'!$A$8:$A${7+MAX_PROF_ROWS},0)),0)')
        ws_res_mensual.cell(ri, 4, f'=B{ri}+C{ri}')
        ws_res_mensual.cell(ri, 5, f'=B{ri}*$H$4')
        ws_res_mensual.cell(ri, 6, f'=C{ri}*$H$4*2')
        ws_res_mensual.cell(ri, 7, f'=E{ri}+F{ri}')

        for col_c in range(2, 8):
            c = ws_res_mensual.cell(ri, col_c)
            c.font = f_data_bold if col_c in (4, 7) else f_data
            c.fill = fill_auto_gray
            c.alignment = al_right
            c.border = border_cell
            c.number_format = CURRENCY_FMT if col_c >= 5 else HOURS_FMT

    ri_tot = ind_row + 6
    ws_res_mensual.row_dimensions[ri_tot].height = 22
    ws_res_mensual.cell(ri_tot, 1, "TOTAL PROFESOR EN EL MES").font = f_data_bold
    ws_res_mensual.cell(ri_tot, 1).fill = fill_total_gray
    ws_res_mensual.cell(ri_tot, 1).border = border_total
    for col_c in range(2, 8):
        col_let = get_column_letter(col_c)
        c = ws_res_mensual.cell(ri_tot, col_c, f'=SUM({col_let}{ind_row+3}:{col_let}{ri_tot-1})')
        c.font = f_data_bold
        c.fill = fill_total_gray
        c.alignment = al_right
        c.border = border_total
        c.number_format = CURRENCY_FMT if col_c >= 5 else HOURS_FMT

    for i, w in enumerate([24, 15, 17, 17, 15, 16, 17, 16, 15, 17, 17, 20], 1):
        ws_res_mensual.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: REPORTES POR UNIDAD (Sección 4.2, 10.3 y 14 Regla R15)
    # --------------------------------------------------------------------------
    add_banner(ws_rep_unidad, "REPORTES DETALLADOS POR UNIDAD DE TRABAJO (PROTEÍNA GYM · PROTEÍNA BOX · BOXEO)",
               "Desglose comparativo de horas normales, horas dobles y costos por profesor en cada sector para el período seleccionado en RESUMEN MENSUAL.", 15)

    ws_rep_unidad.row_dimensions[4].height = 24
    ws_rep_unidad.cell(4, 1, "Período en Consulta:").font = f_data_bold
    ws_rep_unidad.cell(4, 2, "='RESUMEN MENSUAL'!C4").font = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
    ws_rep_unidad.cell(4, 2).fill = fill_config_blue
    ws_rep_unidad.cell(4, 2).alignment = al_center
    ws_rep_unidad.cell(4, 2).border = border_cell

    # Encabezados agrupados por unidad
    ws_rep_unidad.merge_cells("A6:B6")
    ws_rep_unidad.cell(6, 1, "PROFESOR").font = f_sec_hdr
    ws_rep_unidad.cell(6, 1).fill = fill_slate_hdr
    ws_rep_unidad.cell(6, 1).alignment = al_center

    ws_rep_unidad.merge_cells("C6:F6")
    ws_rep_unidad.cell(6, 3, "UNIDAD 1: PROTEÍNA GYM").font = f_sec_hdr
    ws_rep_unidad.cell(6, 3).fill = fill_unit_gym
    ws_rep_unidad.cell(6, 3).alignment = al_center

    ws_rep_unidad.merge_cells("G6:J6")
    ws_rep_unidad.cell(6, 7, "UNIDAD 2: PROTEÍNA BOX").font = f_sec_hdr
    ws_rep_unidad.cell(6, 7).fill = fill_unit_box
    ws_rep_unidad.cell(6, 7).alignment = al_center

    ws_rep_unidad.merge_cells("K6:N6")
    ws_rep_unidad.cell(6, 11, "UNIDAD 3: BOXEO").font = f_sec_hdr
    ws_rep_unidad.cell(6, 11).fill = fill_unit_boxeo
    ws_rep_unidad.cell(6, 11).alignment = al_center

    ws_rep_unidad.cell(6, 15, "CONSOLIDADO").font = f_sec_hdr
    ws_rep_unidad.cell(6, 15).fill = fill_navy
    ws_rep_unidad.cell(6, 15).alignment = al_center
    ws_rep_unidad.row_dimensions[6].height = 24

    rep_subheaders = [
        "Profesor", "Estado",
        "Hs Norm", "Hs Dob", "Total Hs", "Costo Gym ($)",
        "Hs Norm", "Hs Dob", "Total Hs", "Costo Box ($)",
        "Hs Norm", "Hs Dob", "Total Hs", "Costo Boxeo ($)",
        "Costo Total ($)"
    ]
    ws_rep_unidad.row_dimensions[7].height = 23
    for idx_c, h_text in enumerate(rep_subheaders, 1):
        c = ws_rep_unidad.cell(7, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    def choose_rep_cell(cell_coord):
        terms = [f"'{m_name}'!{cell_coord}" for m_name, _, _ in MONTHS]
        return f"CHOOSE(MATCH('RESUMEN MENSUAL'!$C$4,CONFIGURACIÓN!$A$14:$A$27,0),{','.join(terms)})"

    for idx_p in range(MAX_PROF_ROWS):
        r = 8 + idx_p
        m_row = 116 + idx_p
        p_row = 5 + idx_p
        ws_rep_unidad.row_dimensions[r].height = 20

        ws_rep_unidad.cell(r, 1, f'=IF(PROFES!A{p_row}="","",PROFES!A{p_row})')
        ws_rep_unidad.cell(r, 2, f'=IF(A{r}="","",PROFES!B{p_row})')

        # Gym
        ws_rep_unidad.cell(r, 3, f'=IF(A{r}="",0,{choose_rep_cell(f"C{m_row}")})')
        ws_rep_unidad.cell(r, 4, f'=IF(A{r}="",0,{choose_rep_cell(f"D{m_row}")})')
        ws_rep_unidad.cell(r, 5, f'=C{r}+D{r}')
        ws_rep_unidad.cell(r, 6, f'=(C{r}+D{r}*2)*\'RESUMEN MENSUAL\'!$H$4')

        # Box
        ws_rep_unidad.cell(r, 7, f'=IF(A{r}="",0,{choose_rep_cell(f"F{m_row}")})')
        ws_rep_unidad.cell(r, 8, f'=IF(A{r}="",0,{choose_rep_cell(f"G{m_row}")})')
        ws_rep_unidad.cell(r, 9, f'=G{r}+H{r}')
        ws_rep_unidad.cell(r, 10, f'=(G{r}+H{r}*2)*\'RESUMEN MENSUAL\'!$H$4')

        # Boxeo
        ws_rep_unidad.cell(r, 11, f'=IF(A{r}="",0,{choose_rep_cell(f"I{m_row}")})')
        ws_rep_unidad.cell(r, 12, f'=IF(A{r}="",0,{choose_rep_cell(f"J{m_row}")})')
        ws_rep_unidad.cell(r, 13, f'=K{r}+L{r}')
        ws_rep_unidad.cell(r, 14, f'=(K{r}+L{r}*2)*\'RESUMEN MENSUAL\'!$H$4')

        # Total Consolidado Profesor
        ws_rep_unidad.cell(r, 15, f'=F{r}+J{r}+N{r}')

        for col_c in range(1, 16):
            cell = ws_rep_unidad.cell(r, col_c)
            cell.font = f_data_bold if col_c in (1, 6, 10, 14, 15) else f_data
            cell.fill = fill_auto_gray
            cell.alignment = al_left if col_c == 1 else (al_center if col_c == 2 else al_right)
            cell.border = border_cell
            if col_c >= 3:
                cell.number_format = CURRENCY_FMT if col_c in (6, 10, 14, 15) else HOURS_FMT

        if idx_p >= len(PROFESORES):
            ws_rep_unidad.row_dimensions[r].hidden = True

    tot_rep_row = 8 + MAX_PROF_ROWS  # Fila 43
    ws_rep_unidad.row_dimensions[tot_rep_row].height = 24
    ws_rep_unidad.cell(tot_rep_row, 1, "TOTALES POR UNIDAD").font = f_data_bold
    ws_rep_unidad.cell(tot_rep_row, 1).fill = fill_total_gray
    ws_rep_unidad.cell(tot_rep_row, 1).border = border_total
    ws_rep_unidad.cell(tot_rep_row, 2, "").fill = fill_total_gray
    ws_rep_unidad.cell(tot_rep_row, 2).border = border_total

    for col_c in range(3, 16):
        col_let = get_column_letter(col_c)
        c = ws_rep_unidad.cell(tot_rep_row, col_c, f'=SUM({col_let}8:{col_let}{tot_rep_row-1})')
        c.font = f_data_bold
        c.fill = fill_total_gray
        c.alignment = al_right
        c.border = border_total
        c.number_format = CURRENCY_FMT if col_c in (6, 10, 14, 15) else HOURS_FMT

    for i, w in enumerate([20, 12, 11, 11, 12, 16, 11, 11, 12, 16, 11, 11, 12, 16, 18], 1):
        ws_rep_unidad.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: RESUMEN ANUAL (Sección 11: Acumulado Nov 2026 - Dic 2027)
    # --------------------------------------------------------------------------
    add_banner(ws_res_anual, "RESUMEN ANUAL Y EVOLUCIÓN DEL PERÍODO (NOVIEMBRE 2026 A DICIEMBRE 2027)",
               "Acumulación automática de los 14 períodos mensuales por profesor y por unidad de trabajo.", 20)

    # Tabla 1: Acumulado Anual por Profesor + Evolución Mensual (Sección 11.1 y 11.2)
    ws_res_anual.merge_cells("A6:T6")
    c_an_hdr = ws_res_anual.cell(6, 1, "1. TABLA ANUAL ACUMULADA POR PROFESOR Y EVOLUCIÓN DE HORAS MES A MES (NOV 2026 – DIC 2027)")
    c_an_hdr.font = f_sec_hdr
    c_an_hdr.fill = fill_navy
    c_an_hdr.alignment = al_left
    ws_res_anual.row_dimensions[6].height = 24

    anual_headers = [
        "Profesor", "Estado", "Horas Normales", "Horas Dobles", "Total Horas", "Total Cobrado ($)"
    ] + [m[0][:3] + " " + str(m[1])[2:] for m in MONTHS]

    ws_res_anual.row_dimensions[7].height = 24
    for idx_c, h_text in enumerate(anual_headers, 1):
        c = ws_res_anual.cell(7, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    for idx_p in range(MAX_PROF_ROWS):
        r = 8 + idx_p
        m_row = 116 + idx_p
        p_row = 5 + idx_p
        ws_res_anual.row_dimensions[r].height = 20

        ws_res_anual.cell(r, 1, f'=IF(PROFES!A{p_row}="","",PROFES!A{p_row})')
        ws_res_anual.cell(r, 2, f'=IF(A{r}="","",PROFES!B{p_row})')

        norm_terms = "+".join([f"'{m_name}'!L{m_row}" for m_name, _, _ in MONTHS])
        dob_terms = "+".join([f"'{m_name}'!M{m_row}" for m_name, _, _ in MONTHS])
        pay_terms = "+".join([f"'{m_name}'!Q{m_row}" for m_name, _, _ in MONTHS])

        ws_res_anual.cell(r, 3, f'=IF(A{r}="",0,{norm_terms})')
        ws_res_anual.cell(r, 4, f'=IF(A{r}="",0,{dob_terms})')
        ws_res_anual.cell(r, 5, f'=C{r}+D{r}')
        ws_res_anual.cell(r, 6, f'=IF(A{r}="",0,{pay_terms})')

        for m_i, (m_name, _, _) in enumerate(MONTHS):
            col_m = 7 + m_i
            ws_res_anual.cell(r, col_m, f'=IF($A{r}="",0,\'{m_name}\'!N{m_row})')

        for col_c in range(1, 21):
            cell = ws_res_anual.cell(r, col_c)
            cell.font = f_data_bold if col_c in (1, 5, 6) else f_small
            cell.fill = fill_editable_strong if col_c == 6 else fill_auto_gray
            cell.alignment = al_left if col_c == 1 else (al_center if col_c == 2 else al_right)
            cell.border = border_cell
            if col_c >= 3:
                cell.number_format = CURRENCY_FMT if col_c == 6 else HOURS_FMT

        if idx_p >= len(PROFESORES):
            ws_res_anual.row_dimensions[r].hidden = True

    tot_an_row = 8 + MAX_PROF_ROWS  # Fila 43
    ws_res_anual.row_dimensions[tot_an_row].height = 24
    ws_res_anual.cell(tot_an_row, 1, "TOTAL ACUMULADO DEL PERÍODO").font = f_data_bold
    ws_res_anual.cell(tot_an_row, 1).fill = fill_total_gray
    ws_res_anual.cell(tot_an_row, 1).border = border_total
    ws_res_anual.cell(tot_an_row, 2, "").fill = fill_total_gray
    ws_res_anual.cell(tot_an_row, 2).border = border_total

    for col_c in range(3, 21):
        col_let = get_column_letter(col_c)
        c = ws_res_anual.cell(tot_an_row, col_c, f'=SUM({col_let}8:{col_let}{tot_an_row-1})')
        c.font = f_data_bold
        c.fill = fill_total_gray
        c.alignment = al_right
        c.border = border_total
        c.number_format = CURRENCY_FMT if col_c == 6 else HOURS_FMT

    # Tabla 2: Resumen Anual por Unidad (Sección 11.3)
    ua_row = tot_an_row + 2  # Fila 45
    ws_res_anual.merge_cells(f"A{ua_row}:F{ua_row}")
    c_ua = ws_res_anual.cell(ua_row, 1, "2. RESUMEN ANUAL ACUMULADO POR UNIDAD DE TRABAJO (SECCIÓN 11.3)")
    c_ua.font = f_sec_hdr
    c_ua.fill = fill_unit_gym
    c_ua.alignment = al_left
    ws_res_anual.row_dimensions[ua_row].height = 24

    ua_headers = ["Unidad", "Horas Normales", "Horas Dobles", "Total Horas", "Costo Anual Acumulado ($)", "% del Total"]
    ws_res_anual.row_dimensions[ua_row + 1].height = 23
    for idx_c, h_text in enumerate(ua_headers, 1):
        c = ws_res_anual.cell(ua_row + 1, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    unit_an_cols = [
        ("Proteína Gym", "C151", "D151"),
        ("Proteína Box", "F151", "G151"),
        ("Boxeo", "I151", "J151"),
    ]
    tot_ua_row = ua_row + 5
    for idx_u, (u_name, c_norm, c_dob) in enumerate(unit_an_cols):
        ru = ua_row + 2 + idx_u
        ws_res_anual.row_dimensions[ru].height = 21
        ws_res_anual.cell(ru, 1, u_name).font = f_data_bold
        ws_res_anual.cell(ru, 1).fill = fill_auto_gray
        ws_res_anual.cell(ru, 1).border = border_cell

        u_norm_sum = "+".join([f"'{m_name}'!{c_norm}" for m_name, _, _ in MONTHS])
        u_dob_sum = "+".join([f"'{m_name}'!{c_dob}" for m_name, _, _ in MONTHS])
        u_cost_sum = "+".join([
            f"(('{m_name}'!{c_norm}*CONFIGURACIÓN!D{14+mi})+('{m_name}'!{c_dob}*CONFIGURACIÓN!E{14+mi}))"
            for mi, (m_name, _, _) in enumerate(MONTHS)
        ])

        ws_res_anual.cell(ru, 2, f'={u_norm_sum}')
        ws_res_anual.cell(ru, 3, f'={u_dob_sum}')
        ws_res_anual.cell(ru, 4, f'=B{ru}+C{ru}')
        ws_res_anual.cell(ru, 5, f'={u_cost_sum}')
        ws_res_anual.cell(ru, 6, f'=IF($E${tot_ua_row}=0,0,E{ru}/$E${tot_ua_row})')

        for col_c in range(2, 7):
            c = ws_res_anual.cell(ru, col_c)
            c.font = f_data_bold if col_c in (4, 5) else f_data
            c.fill = fill_auto_gray
            c.alignment = al_right
            c.border = border_cell
            c.number_format = '0.0%' if col_c == 6 else (CURRENCY_FMT if col_c == 5 else HOURS_FMT)

    ws_res_anual.row_dimensions[tot_ua_row].height = 23
    ws_res_anual.cell(tot_ua_row, 1, "TOTAL GENERAL CONSOLIDADO").font = f_data_bold
    ws_res_anual.cell(tot_ua_row, 1).fill = fill_total_gray
    ws_res_anual.cell(tot_ua_row, 1).border = border_total
    for col_c in range(2, 7):
        col_let = get_column_letter(col_c)
        c = ws_res_anual.cell(tot_ua_row, col_c, f'=SUM({col_let}{ua_row+2}:{col_let}{tot_ua_row-1})')
        c.font = f_data_bold
        c.fill = fill_total_gray
        c.alignment = al_right
        c.border = border_total
        c.number_format = '0.0%' if col_c == 6 else (CURRENCY_FMT if col_c == 5 else HOURS_FMT)

    # Indicadores Generales Anuales (Sección 11.4)
    ws_res_anual.merge_cells(f"H{ua_row}:M{ua_row}")
    c_ind_a = ws_res_anual.cell(ua_row, 8, "3. INDICADORES DESTACADOS DEL PERÍODO (SECCIÓN 11.4)")
    c_ind_a.font = f_sec_hdr
    c_ind_a.fill = fill_unit_box
    c_ind_a.alignment = al_left

    ind_anuales = [
        ("Profesor con Mayor Cantidad de Horas", f'=INDEX(A8:A{tot_an_row-1},MATCH(MAX(E8:E{tot_an_row-1}),E8:E{tot_an_row-1},0))', '@'),
        ("Máximo de Horas de un Profesor", f'=MAX(E8:E{tot_an_row-1})', HOURS_FMT),
        ("Mes con Mayor Costo de Personal", '=INDEX(CONFIGURACIÓN!A14:A27,MATCH(MAX(CONFIGURACIÓN!G14:G27),CONFIGURACIÓN!G14:G27,0))', '@'),
        ("Promedio Mensual de Horas (Meses Activos)", '=AVERAGEIF(CONFIGURACIÓN!F14:F27,">0")', HOURS_FMT),
        ("Costo Total Acumulado del Staff", f'=F{tot_an_row}', CURRENCY_FMT),
    ]
    for idx_i, (ilbl, iform, ifmt) in enumerate(ind_anuales):
        ri = ua_row + 1 + idx_i
        ws_res_anual.merge_cells(f"H{ri}:K{ri}")
        ws_res_anual.cell(ri, 8, ilbl).font = f_data_bold
        ws_res_anual.cell(ri, 8).fill = fill_auto_gray
        ws_res_anual.cell(ri, 8).border = border_cell

        ws_res_anual.merge_cells(f"L{ri}:M{ri}")
        cv = ws_res_anual.cell(ri, 12, iform)
        cv.font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E3A8A")
        cv.fill = fill_config_blue
        cv.alignment = al_center if ifmt == '@' else al_right
        cv.border = border_cell
        cv.number_format = ifmt

    an_widths = [22, 12, 15, 15, 15, 18] + [11] * 14
    for i, w in enumerate(an_widths, 1):
        ws_res_anual.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: REGISTRO DE HORAS (Base de Datos Interna - Sección 12)
    # --------------------------------------------------------------------------
    add_banner(ws_reg_horas, "REGISTRO DE HORAS — BASE DE DATOS INTERNA NORMALIZADA (SECCIÓN 12)",
               "Cada fila representa 0,5 hs de un profesor individual (desglosando selecciones múltiples y reemplazos). Pestaña técnica protegida.", 15)

    bd_headers = [
        "ID Asignación", "Período", "Fecha", "Día de la Semana",
        "Hora Inicio", "Hora Fin", "Duración (Hs)", "Profesor",
        "Unidad", "Sección", "Tipo Asignación", "Tipo de Jornada",
        "Valor Hora Aplicable", "Importe Calculado ($)", "Estado Liquidación"
    ]
    ws_reg_horas.row_dimensions[4].height = 25
    for idx_c, h_text in enumerate(bd_headers, 1):
        c = ws_reg_horas.cell(4, idx_c, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_slate_hdr
        c.alignment = al_center
        c.border = border_cell

    # Cargamos los primeros 600 registros normalizados de muestra en REGISTRO DE HORAS (y el script .gs sincroniza el 100% bajo demanda)
    sample_records = registros_bd[:600]
    for idx_r, rec in enumerate(sample_records):
        r = 5 + idx_r
        cfg_r = rec["cfg_row"]
        ws_reg_horas.cell(r, 1, rec["id"])
        ws_reg_horas.cell(r, 2, rec["mes"])
        c_dt = ws_reg_horas.cell(r, 3, rec["fecha"])
        c_dt.number_format = DATE_FMT
        ws_reg_horas.cell(r, 4, rec["dia"])
        ws_reg_horas.cell(r, 5, rec["h_ini"])
        ws_reg_horas.cell(r, 6, rec["h_fin"])
        c_du = ws_reg_horas.cell(r, 7, rec["duracion"])
        c_du.number_format = HOURS_FMT
        ws_reg_horas.cell(r, 8, rec["profesor"])
        ws_reg_horas.cell(r, 9, rec["unidad"])
        ws_reg_horas.cell(r, 10, rec["seccion"])
        ws_reg_horas.cell(r, 11, rec["tipo_asig"])
        ws_reg_horas.cell(r, 12, rec["jornada"])
        c_vh = ws_reg_horas.cell(r, 13, f'=CONFIGURACIÓN!$D${cfg_r}')
        c_vh.number_format = CURRENCY_FMT
        c_im = ws_reg_horas.cell(r, 14, f'=G{r}*M{r}*{rec["mult"]}')
        c_im.number_format = CURRENCY_FMT
        ws_reg_horas.cell(r, 15, f'=CONFIGURACIÓN!$B${cfg_r}')

    bd_widths = [15, 16, 13, 15, 12, 12, 13, 18, 16, 14, 15, 15, 18, 18, 16]
    for i, w in enumerate(bd_widths, 1):
        ws_reg_horas.column_dimensions[get_column_letter(i)].width = w

    # --------------------------------------------------------------------------
    # PESTAÑA: INICIO (Panel Principal y Accesos Directos - Sección 13)
    # --------------------------------------------------------------------------
    add_banner(ws_inicio, "SISTEMA INTEGRAL DE GESTIÓN DEL STAFF — PROTEÍNA GYM · PROTEÍNA BOX · BOXEO",
               "Panel Principal de Administración de Horarios, Profesores, Reemplazos y Liquidaciones (Noviembre 2026 – Diciembre 2027)", 10)

    # Tarjetas de Indicadores en Vivo (Sección 13)
    ws_inicio.merge_cells("A4:J4")
    c_kpi_t = ws_inicio.cell(4, 1, "INDICADORES EN TIEMPO REAL DEL PERÍODO SELECCIONADO EN RESUMEN MENSUAL")
    c_kpi_t.font = f_sec_hdr
    c_kpi_t.fill = fill_navy
    c_kpi_t.alignment = al_left
    ws_inicio.row_dimensions[4].height = 24

    home_kpis = [
        (1, 2, "MES EN CONSULTA", "='RESUMEN MENSUAL'!C4", "@", fill_config_blue),
        (3, 4, "VALOR HORA NORMAL VIGENTE", "=CONFIGURACIÓN!B6", CURRENCY_FMT, fill_config_blue),
        (5, 6, "PROFESORES ACTIVOS", "=CONFIGURACIÓN!B9", "0", fill_config_blue),
        (7, 8, "HORAS MENSUALES STAFF", "='RESUMEN MENSUAL'!J4", HOURS_FMT, fill_editable_strong),
        (9, 10, "COSTO MENSUAL CONSOLIDADO", "='RESUMEN MENSUAL'!L4", CURRENCY_FMT, fill_editable_strong),
    ]
    ws_inicio.row_dimensions[5].height = 18
    ws_inicio.row_dimensions[6].height = 28
    for c_start, c_end, lbl, formula, fmt, fl in home_kpis:
        ws_inicio.merge_cells(start_row=5, start_column=c_start, end_row=5, end_column=c_end)
        ws_inicio.merge_cells(start_row=6, start_column=c_start, end_row=6, end_column=c_end)
        cl = ws_inicio.cell(5, c_start, lbl)
        cl.font = f_small_bold
        cl.fill = fill_auto_gray
        cl.alignment = al_center
        cl.border = border_cell

        cv = ws_inicio.cell(6, c_start, formula)
        cv.font = Font(name=FONT_FAMILY, size=12, bold=True, color="0F172A")
        cv.fill = fl
        cv.alignment = al_center
        cv.border = border_cell
        cv.number_format = fmt

    # Desglose rápido por Unidad en Inicio
    ws_inicio.row_dimensions[8].height = 24
    ws_inicio.merge_cells("A8:J8")
    c_u_home = ws_inicio.cell(8, 1, "COSTO MENSUAL Y HORAS POR UNIDAD DE TRABAJO (MES EN CONSULTA)")
    c_u_home.font = f_sec_hdr
    c_u_home.fill = fill_slate_hdr
    c_u_home.alignment = al_left

    u_home_cards = [
        (1, 3, "UNIDAD 1: PROTEÍNA GYM", f"='RESUMEN MENSUAL'!D{u_tbl_row+2}", f"='RESUMEN MENSUAL'!G{u_tbl_row+2}", fill_unit_gym),
        (4, 6, "UNIDAD 2: PROTEÍNA BOX", f"='RESUMEN MENSUAL'!D{u_tbl_row+3}", f"='RESUMEN MENSUAL'!G{u_tbl_row+3}", fill_unit_box),
        (7, 10, "UNIDAD 3: BOXEO", f"='RESUMEN MENSUAL'!D{u_tbl_row+4}", f"='RESUMEN MENSUAL'!G{u_tbl_row+4}", fill_unit_boxeo),
    ]
    ws_inicio.row_dimensions[9].height = 22
    ws_inicio.row_dimensions[10].height = 24
    for c_s, c_e, u_title, f_hs, f_cost, u_fl in u_home_cards:
        ws_inicio.merge_cells(start_row=9, start_column=c_s, end_row=9, end_column=c_e)
        ch = ws_inicio.cell(9, c_s, u_title)
        ch.font = f_tbl_hdr
        ch.fill = u_fl
        ch.alignment = al_center
        ch.border = border_cell

        c1 = ws_inicio.cell(10, c_s, f_hs)
        c1.font = f_data_bold
        c1.fill = fill_auto_gray
        c1.alignment = al_right
        c1.border = border_cell
        c1.number_format = '0.0 "hs"'

        ws_inicio.merge_cells(start_row=10, start_column=c_s+1, end_row=10, end_column=c_e)
        c2 = ws_inicio.cell(10, c_s+1, f_cost)
        c2.font = Font(name=FONT_FAMILY, size=11, bold=True, color="0F172A")
        c2.fill = fill_editable_green
        c2.alignment = al_right
        c2.border = border_cell
        c2.number_format = CURRENCY_FMT

    # Accesos directos interactivos (Sección 13: ¿Qué necesitás hacer?)
    ws_inicio.merge_cells("A12:J12")
    c_nav = ws_inicio.cell(12, 1, "¿QUÉ NECESITÁS HACER? — ACCESOS DIRECTOS A LAS SECCIONES DEL SISTEMA")
    c_nav.font = f_sec_hdr
    c_nav.fill = fill_navy
    c_nav.alignment = al_left
    ws_inicio.row_dimensions[12].height = 24

    nav_links = [
        ("1. Administrar Profesores", "PROFES", "Agregar nuevos integrantes, activar/desactivar bajas y consultar colores asignados."),
        ("2. Configurar Valor Hora y Cierres", "CONFIGURACIÓN", "Modificar el valor único de la hora ($5.000), ver histórico de tarifas y cerrar meses liquidados."),
        ("3. Calendario de Feriados (×2)", "FERIADOS", "Consultar o agregar feriados nacionales y fechas especiales con pago doble automático."),
        ("4. Plantilla Semanal Base", "PLANTILLA SEMANAL", "Ver y editar la grilla modelo de Lunes a Domingo para Proteína Gym, Proteína Box y Boxeo."),
        ("5. Liquidación y Resumen Mensual", "RESUMEN MENSUAL", "Consultar cuánto cobrará cada profesor en el mes, el costo por unidad y el recibo individual."),
        ("6. Resumen Anual Acumulado", "RESUMEN ANUAL", "Analizar el acumulado de horas y pagos desde Noviembre 2026 hasta Diciembre 2027."),
        ("7. Reportes Comparativos por Unidad", "REPORTES POR UNIDAD", "Comparar horas y costos de personal entre Proteína Gym, Proteína Box y Boxeo."),
        ("8. Base de Datos Registro de Horas", "REGISTRO DE HORAS", "Auditar los registros individuales normalizados de 30 minutos por fecha, unidad y profesor."),
    ]
    for idx_n, (titulo, sheet_target, desc) in enumerate(nav_links, start=13):
        ws_inicio.row_dimensions[idx_n].height = 22
        ws_inicio.merge_cells(f"A{idx_n}:C{idx_n}")
        c_lk = ws_inicio.cell(idx_n, 1, f'=HYPERLINK("#\'{sheet_target}\'!A1", "👉 {titulo}")')
        c_lk.font = f_link
        c_lk.fill = fill_editable_green
        c_lk.alignment = al_left
        c_lk.border = border_cell

        ws_inicio.merge_cells(f"D{idx_n}:J{idx_n}")
        c_ds = ws_inicio.cell(idx_n, 4, desc)
        c_ds.font = f_data
        c_ds.fill = fill_white
        c_ds.alignment = al_left
        c_ds.border = border_cell

    # Accesos directos a los 14 Cronogramas Mensuales
    ws_inicio.merge_cells("A22:J22")
    c_mnav = ws_inicio.cell(22, 1, "ACCESO RÁPIDO A LOS 14 CRONOGRAMAS MENSUALES (NOVIEMBRE 2026 – DICIEMBRE 2027)")
    c_mnav.font = f_sec_hdr
    c_mnav.fill = fill_unit_box
    c_mnav.alignment = al_left
    ws_inicio.row_dimensions[22].height = 24

    for idx_m, (m_name, _, _) in enumerate(MONTHS):
        r_m = 23 + (idx_m // 5)
        ws_inicio.row_dimensions[r_m].height = 24
        col_m = 1 + (idx_m % 5) * 2
        ws_inicio.merge_cells(start_row=r_m, start_column=col_m, end_row=r_m, end_column=col_m + 1)
        cm = ws_inicio.cell(r_m, col_m, f'=HYPERLINK("#\'{m_name}\'!A1", "📅 {m_name}")')
        cm.font = f_link
        cm.fill = fill_config_blue
        cm.alignment = al_center
        cm.border = border_cell

    # Guía visual de colores de celdas (Sección 15.2)
    ws_inicio.merge_cells("A27:J27")
    c_guia = ws_inicio.cell(27, 1, "GUÍA RÁPIDA DE COLORES DEL SISTEMA (EXPERIENCIA DE USUARIO - SECCIÓN 15.2)")
    c_guia.font = f_sec_hdr
    c_guia.fill = fill_slate_hdr
    c_guia.alignment = al_left
    ws_inicio.row_dimensions[27].height = 24

    legend_rows = [
        ("CELDAS VERDES", fill_editable_green, "Celdas editables por el usuario: asignación de profesores en cronogramas, altas en PROFES, feriados y selectores."),
        ("CELDAS AZULES", fill_config_blue, "Celdas de configuración y parámetros: valor de la hora vigente, vigencias históricas y cierres."),
        ("CELDAS GRISES", fill_auto_gray, "Celdas automáticas y protegidas: fórmulas de horas, liquidaciones, resúmenes y base de datos. No requieren edición manual."),
        ("CELDAS ROJAS SUAVES", fill_double_day, "Domingos y Feriados detectados automáticamente: todas las horas asignadas en estas columnas se liquidan al DOBLE (×2)."),
    ]
    for idx_l, (l_name, l_fill, l_desc) in enumerate(legend_rows, start=28):
        ws_inicio.row_dimensions[idx_l].height = 21
        ws_inicio.merge_cells(f"A{idx_l}:B{idx_l}")
        c_ln = ws_inicio.cell(idx_l, 1, l_name)
        c_ln.font = f_data_bold
        c_ln.fill = l_fill
        c_ln.alignment = al_center
        c_ln.border = border_cell

        ws_inicio.merge_cells(f"C{idx_l}:J{idx_l}")
        c_ld = ws_inicio.cell(idx_l, 3, l_desc)
        c_ld.font = f_data
        c_ld.fill = fill_white
        c_ld.alignment = al_left
        c_ld.border = border_cell

    for i in range(1, 11):
        ws_inicio.column_dimensions[get_column_letter(i)].width = 16

    wb.save(OUTPUT_XLSX)
    print(f"Workbook guardado exitosamente en: {OUTPUT_XLSX}")
    print(f"Total pestañas creadas: {len(wb.sheetnames)}")
    print(f"Total registros individuales de 30 min generados en muestra: {len(registros_bd)}")


if __name__ == "__main__":
    build_workbook()
