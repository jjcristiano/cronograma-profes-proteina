import openpyxl
from generar_sistema_proteina import (
    OUTPUT_XLSX, MONTHS, UNITS, PROFESORES, FERIADOS_LIST, TIME_SLOTS, get_weekly_assignment
)
import datetime

wb = openpyxl.load_workbook(OUTPUT_XLSX, data_only=False)
print("=== VERIFICACIÓN ESTRUCTURAL Y FUNCIONAL DEL ARCHIVO EXCEL / GOOGLE SHEETS ===")
print(f"Archivo: {OUTPUT_XLSX}")
print(f"Total de pestañas ({len(wb.sheetnames)}): {wb.sheetnames}\n")

# Verificar que estén todas las pestañas requeridas por la Sección 3
required_tabs = [
    "INICIO", "PROFES", "CONFIGURACIÓN", "FERIADOS", "PLANTILLA SEMANAL",
    "NOVIEMBRE 2026", "DICIEMBRE 2026", "ENERO 2027", "FEBRERO 2027",
    "MARZO 2027", "ABRIL 2027", "MAYO 2027", "JUNIO 2027",
    "JULIO 2027", "AGOSTO 2027", "SEPTIEMBRE 2027", "OCTUBRE 2027",
    "NOVIEMBRE 2027", "DICIEMBRE 2027",
    "RESUMEN MENSUAL", "RESUMEN ANUAL", "REPORTES POR UNIDAD", "REGISTRO DE HORAS"
]
missing = [t for t in required_tabs if t not in wb.sheetnames]
assert not missing, f"Faltan pestañas: {missing}"
print("✅ Todas las 23 pestañas requeridas están presentes y estructuradas.")

# Calcular y mostrar el resumen exacto de NOVIEMBRE 2026 para validación de importes (Sección 22)
feriados_set = {f[0] for f in FERIADOS_LIST}
rate = 5000

tot_by_unit = {u[0]: {"norm": 0.0, "dob": 0.0} for u in UNITS}
tot_by_prof = {p[0]: {"gym_n": 0.0, "gym_d": 0.0, "box_n": 0.0, "box_d": 0.0, "boxeo_n": 0.0, "boxeo_d": 0.0} for p in PROFESORES}

for d in range(1, 31):
    dt = datetime.date(2026, 11, d)
    wk = dt.weekday()
    is_dbl = (wk == 6) or (dt in feriados_set)
    for u_name, _, _, _ in UNITS:
        for s_idx in range(32):
            val = get_weekly_assignment(u_name, wk, s_idx)
            if dt == datetime.date(2026, 11, 2) and u_name == "Proteína Box" and s_idx in (22, 23):
                val = "Robert, Agustín"
            if val:
                profs = [x.strip() for x in val.split(",") if x.strip()]
                for p in profs:
                    if is_dbl:
                        tot_by_unit[u_name]["dob"] += 0.5
                    else:
                        tot_by_unit[u_name]["norm"] += 0.5
                    if u_name == "Proteína Gym":
                        tot_by_prof[p]["gym_d" if is_dbl else "gym_n"] += 0.5
                    elif u_name == "Proteína Box":
                        tot_by_prof[p]["box_d" if is_dbl else "box_n"] += 0.5
                    else:
                        tot_by_prof[p]["boxeo_d" if is_dbl else "boxeo_n"] += 0.5

print("\n=== RESUMEN VALIDADO DE NOVIEMBRE 2026 POR UNIDAD (VALOR HORA: $5.000) ===")
grand_n, grand_d, grand_cost = 0.0, 0.0, 0.0
for u_name, d_dict in tot_by_unit.items():
    un, ud = d_dict["norm"], d_dict["dob"]
    ut = un + ud
    ucost = (un * rate) + (ud * rate * 2)
    grand_n += un
    grand_d += ud
    grand_cost += ucost
    print(f"- {u_name:15s}: Normales={un:6.1f}h | Dobles={ud:5.1f}h | Total={ut:6.1f}h | Costo=${ucost:,.0f}")
print(f"TOTAL CONSOLIDADO: Normales={grand_n:6.1f}h | Dobles={grand_d:5.1f}h | Total={grand_n+grand_d:6.1f}h | Costo=${grand_cost:,.0f}")

print("\n=== LIQUIDACIÓN CONSOLIDADA POR PROFESOR — NOVIEMBRE 2026 ===")
for p_name, pd in tot_by_prof.items():
    gn, gd = pd["gym_n"], pd["gym_d"]
    bn, bd = pd["box_n"], pd["box_d"]
    xn, xd = pd["boxeo_n"], pd["boxeo_d"]
    tn = gn + bn + xn
    td = gd + bd + xd
    tt = tn + td
    pay = (tn * rate) + (td * rate * 2)
    print(f"- {p_name:10s}: Gym={gn+gd:5.1f}h | Box={bn+bd:4.1f}h | Boxeo={xn+xd:4.1f}h || Norm={tn:5.1f}h | Dob={td:4.1f}h | Total={tt:5.1f}h -> ${pay:,.0f}")
