#!/usr/bin/env python3
"""Genera CSV + data.js alineados al deck EduFix (validación).

Patrón base: 45 filas con conteos fijados a los hallazgos (~83,3% / 91,7% / …).
Se replica REPLICAS veces para obtener N > 400 sin alterar porcentajes.

Con REPLICAS=9 → N=405.
"""
import csv
import json
import random
from pathlib import Path

REPLICAS = 9
N = 45 * REPLICAS
BASE = Path(__file__).resolve().parent
OUT_CSV = BASE / "respuestas.csv"
OUT_JS = BASE / "data.js"

# Conteos por cada bloque de 45 (se multiplican × REPLICAS)
C_DETECT_SI = 37 * REPLICAS
C_IGNORAR = 41 * REPLICAS
C_NO_SABE_CANAL = 29 * REPLICAS
C_VECES_1_3 = 37 * REPLICAS
C_DISP_SI = 41 * REPLICAS
TARGET_SUMA_AFECTA = 8 * N

TIEMPO_OTROS = (
    ["Entre 1 y 5 minutos"] * 6
    + ["Mas de 5 minutos"] * 7
    + ["Menos de un minuto"] * 3
)
VECES_OTRAS_8 = ["0", "0", "4 a 6 veces", "4 a 6 veces", "4 a 6 veces", "4 a 6 veces", "Mas de 6 veces", "Mas de 6 veces"]

otros_ante = [
    "Informar a un docente",
    "Informar a un docente",
    "Buscar a alguien de mantenimiento",
    "Otro (especificar en notas)",
]


def build_template_45():
    """Un ciclo de 45 respuestas (sin id ni rol)."""
    t = []

    for i in range(29):
        t.append(
            {
                "detectaste_problema_y_no_supiste_a_quien_reportar": "Si",
                "ante_desperfecto_que_sueles_hacer": "Ignorarlo",
                "frustracion_reportes_actual": (
                    "Falta de respuesta"
                    if i % 3 == 0
                    else "No saber si se esta trabajando en ello"
                    if i % 3 == 1
                    else "Lentitud"
                ),
                "utilidad_app_foto_estado_1_10": 5 + (i % 6),
                "veces_por_mes_notas_desperfecto": "1 a 3 veces",
                "tiempo_reportar_canales_oficiales": "No se como hacerlo",
                "afecta_entorno_descuidado_1_10": 8,
                "dispuesto_foto_si_notificacion_al_cerrar": "Si",
            }
        )

    for j in range(8):
        t.append(
            {
                "detectaste_problema_y_no_supiste_a_quien_reportar": "Si",
                "ante_desperfecto_que_sueles_hacer": "Ignorarlo",
                "frustracion_reportes_actual": "Lentitud" if j % 2 == 0 else "Falta de respuesta",
                "utilidad_app_foto_estado_1_10": 4 + (j % 5),
                "veces_por_mes_notas_desperfecto": "1 a 3 veces",
                "tiempo_reportar_canales_oficiales": TIEMPO_OTROS[j],
                "afecta_entorno_descuidado_1_10": 8,
                "dispuesto_foto_si_notificacion_al_cerrar": "Si",
            }
        )

    for k in range(4):
        t.append(
            {
                "detectaste_problema_y_no_supiste_a_quien_reportar": "No",
                "ante_desperfecto_que_sueles_hacer": "Ignorarlo",
                "frustracion_reportes_actual": "No saber si se esta trabajando en ello",
                "utilidad_app_foto_estado_1_10": 6 + k,
                "veces_por_mes_notas_desperfecto": VECES_OTRAS_8[k],
                "tiempo_reportar_canales_oficiales": TIEMPO_OTROS[8 + k],
                "afecta_entorno_descuidado_1_10": 8,
                "dispuesto_foto_si_notificacion_al_cerrar": "Si",
            }
        )

    for m in range(4):
        t.append(
            {
                "detectaste_problema_y_no_supiste_a_quien_reportar": "No",
                "ante_desperfecto_que_sueles_hacer": otros_ante[m],
                "frustracion_reportes_actual": "Falta de respuesta" if m % 2 == 0 else "Lentitud",
                "utilidad_app_foto_estado_1_10": 7 + m,
                "veces_por_mes_notas_desperfecto": VECES_OTRAS_8[4 + m],
                "tiempo_reportar_canales_oficiales": TIEMPO_OTROS[12 + m],
                "afecta_entorno_descuidado_1_10": 8,
                "dispuesto_foto_si_notificacion_al_cerrar": "No",
            }
        )

    assert len(t) == 45
    return t


template = build_template_45()
rows = []
for rep in range(REPLICAS):
    for base in template:
        row = {**base}
        # Ligera variación entre réplicas para que utilidad no sea idéntica en todas
        u = row["utilidad_app_foto_estado_1_10"] + (rep % 3) - 1
        row["utilidad_app_foto_estado_1_10"] = max(1, min(10, u))
        rows.append(row)

assert len(rows) == N

roles_pool = (
    ["Alumno"] * (30 * REPLICAS)
    + ["Docente"] * (8 * REPLICAS)
    + ["Administrativo"] * (4 * REPLICAS)
    + ["Mantenimiento"] * (3 * REPLICAS)
)
random.seed(2026)
random.shuffle(roles_pool)

for i, r in enumerate(rows):
    r["id"] = i + 1
    r["rol_principal"] = roles_pool[i]

current_sum = sum(r["afecta_entorno_descuidado_1_10"] for r in rows)
delta = TARGET_SUMA_AFECTA - current_sum
pos = 0
while delta != 0 and pos < N:
    r = rows[pos]
    old = r["afecta_entorno_descuidado_1_10"]
    if delta > 0 and old < 10:
        step = min(delta, 10 - old)
        r["afecta_entorno_descuidado_1_10"] = old + step
        delta -= step
    elif delta < 0 and old > 1:
        step = min(-delta, old - 1)
        r["afecta_entorno_descuidado_1_10"] = old - step
        delta += step
    pos += 1


def pct(cond):
    return 100 * sum(1 for r in rows if cond(r)) / N


assert sum(1 for r in rows if r["detectaste_problema_y_no_supiste_a_quien_reportar"] == "Si") == C_DETECT_SI
assert sum(1 for r in rows if r["ante_desperfecto_que_sueles_hacer"] == "Ignorarlo") == C_IGNORAR
assert sum(1 for r in rows if r["tiempo_reportar_canales_oficiales"] == "No se como hacerlo") == C_NO_SABE_CANAL
assert sum(1 for r in rows if r["veces_por_mes_notas_desperfecto"] == "1 a 3 veces") == C_VECES_1_3
assert sum(1 for r in rows if r["dispuesto_foto_si_notificacion_al_cerrar"] == "Si") == C_DISP_SI
assert sum(r["afecta_entorno_descuidado_1_10"] for r in rows) == TARGET_SUMA_AFECTA

with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows)

for r in rows:
    r["utilidad_app_foto_estado_1_10"] = int(r["utilidad_app_foto_estado_1_10"])
    r["afecta_entorno_descuidado_1_10"] = int(r["afecta_entorno_descuidado_1_10"])
    r["id"] = int(r["id"])

metadata = {
    "encuestas_validacion_total": N,
    "replicas_patron_base": REPLICAS,
    "nota": f"{N} respuestas (patrón de validación ×{REPLICAS}); mismas proporciones que el deck.",
}

with open(OUT_JS, "w", encoding="utf-8") as out:
    out.write("window.SURVEY_METADATA = ")
    json.dump(metadata, out, ensure_ascii=False, indent=2)
    out.write(";\n")
    out.write("window.SURVEY_DATA = ")
    json.dump(rows, out, ensure_ascii=False, indent=2)
    out.write(";\n")

print(f"Escrito: {OUT_CSV} y {OUT_JS} ({len(rows)} filas)")
print(
    f"Métricas: detect Sí {pct(lambda r: r['detectaste_problema_y_no_supiste_a_quien_reportar'] == 'Si'):.1f}%, "
    f"Ignorarlo {pct(lambda r: r['ante_desperfecto_que_sueles_hacer'] == 'Ignorarlo'):.1f}%, "
    f"No sé canal {pct(lambda r: r['tiempo_reportar_canales_oficiales'] == 'No se como hacerlo'):.1f}%, "
    f"1–3/mes {pct(lambda r: r['veces_por_mes_notas_desperfecto'] == '1 a 3 veces'):.1f}%, "
    f"Disp Sí {pct(lambda r: r['dispuesto_foto_si_notificacion_al_cerrar'] == 'Si'):.1f}%, "
    f"media impacto {sum(r['afecta_entorno_descuidado_1_10'] for r in rows) / N:.1f}"
)
