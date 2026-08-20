#!/usr/bin/env python3
"""Da de alta en la BD los distribuidores que solo trae Porcelanite.

Uso:
    python3 altas_porcelanite.py            # muestra que agregaria
    python3 altas_porcelanite.py --aplicar  # escribe en Distribuidores.xlsx

Regla, la misma que se uso con Vitromex: si el distribuidor ya existe en la BD
como grupo, todas sus sucursales ya estan y no se toca nada. Solo entran
empresas que la BD no tiene.

Antes de dar de alta, cada sucursal se compara contra las coordenadas de la BD:
si cae a menos de 100 m de una tienda existente, se descarta, porque casi
siempre significa que es la misma tienda registrada con otro nombre.
"""

import math
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl

SERVIDOR = Path(
    "/Volumes/MK/MK Server/Finanzas/Strategy 2026/Distribuidores/Data/Sucursales"
)
BD = SERVIDOR / "Distribuidores.xlsx"
PORCELANITE = SERVIDOR / "Marcas" / "Porcelanite.xlsx"

# Comparten direccion exacta con una empresa de la BD, pero no alcanzo para
# confirmar que sean la misma razon social. Se quedan fuera: es preferible que
# falte un distribuidor a que se duplique uno.
DUDOSOS = {"GRUPO JCC", "VIZUET", "GUTIERREZ", "PALENCIA", "BERNARDO MENESES"}

RUIDO = {"grupo", "comercializadora", "distribuidora", "distribuidor",
         "corporativo", "sa", "de", "cv", "s", "rl", "sapi", "gpo"}

METROS = 100  # mas cerca que esto de una tienda existente = es la misma


def na(s):
    s = "".join(c for c in unicodedata.normalize("NFD", str(s or "").lower())
                if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def clave(s):
    """Nombre reducido a sus palabras distintivas, para comparar razones sociales."""
    return " ".join(sorted(w for w in re.sub(r"[^\w\s]", " ", na(s)).split()
                           if w not in RUIDO and len(w) > 2))


def km(a, b, c, d):
    return (((a - c) * 110.57) ** 2
            + ((b - d) * 111.32 * math.cos(math.radians(a))) ** 2) ** 0.5


def main():
    aplicar = "--aplicar" in sys.argv

    libro = openpyxl.load_workbook(BD)  # sin data_only: conserva las formulas
    hoja = libro["BD"]
    existentes, puntos = set(), defaultdict(list)
    for fila in hoja.iter_rows(min_row=2):
        if not fila[0].value:
            continue
        existentes.add(clave(fila[0].value))
        try:
            lat, lon = float(fila[7].value), float(fila[8].value)
        except (TypeError, ValueError):
            continue
        puntos[(round(lat, 2), round(lon, 2))].append((lat, lon))

    wb = openpyxl.load_workbook(PORCELANITE, read_only=True, data_only=True)
    porcelanite = [list(f) for f in wb[wb.sheetnames[0]].iter_rows(min_row=2, values_only=True)]
    wb.close()

    candidatas = [f for f in porcelanite
                  if clave(f[0]) not in existentes and f[0] not in DUDOSOS]

    altas, encimadas = [], 0
    for f in candidatas:
        lat, lon = f[6], f[7]
        pegada = any(
            km(lat, lon, x, y) * 1000 < METROS
            for dla in (-0.01, 0, 0.01) for dlo in (-0.01, 0, 0.01)
            for x, y in puntos.get((round(lat + dla, 2), round(lon + dlo, 2)), [])
        )
        if pegada:
            encimadas += 1
        else:
            altas.append(f)

    ya = len({f[0] for f in porcelanite}) - len({f[0] for f in candidatas}) - len(
        {f[0] for f in porcelanite if f[0] in DUDOSOS})
    print(f"  distribuidores que la BD ya tiene:  {ya:>3}  (no se tocan)")
    print(f"  dudosos, se dejan fuera:            {len(DUDOSOS):>3}")
    print(f"  sucursales encimadas a una de la BD:{encimadas:>4}  (se descartan)")
    print(f"  ALTAS: {len(altas)} sucursales de {len({f[0] for f in altas})} distribuidores")

    if not aplicar:
        print("\n  Nada se escribio. Para aplicarlo:  python3 altas_porcelanite.py --aplicar")
        return

    n = hoja.max_row
    for f in altas:
        n += 1
        hoja.cell(n, 1, f[0])   # DISTRIBUIDOR
        hoja.cell(n, 2, f[2])   # NOMBRE
        hoja.cell(n, 4, f[1])   # PAÍS
        hoja.cell(n, 5, f[3])   # ESTADO
        hoja.cell(n, 6, f[4])   # CIUDAD
        hoja.cell(n, 7, f[5])   # DIRECCIÓN
        hoja.cell(n, 8, f[6])   # LATITUD
        hoja.cell(n, 9, f[7])   # LONGITUD
        hoja.cell(n, 10, f'=IF(H{n}="","","https://www.google.com/maps?q="&H{n}&","&I{n})')
    libro.save(BD)
    print(f"\n  BD queda en {hoja.max_row - 1} filas")


if __name__ == "__main__":
    main()
