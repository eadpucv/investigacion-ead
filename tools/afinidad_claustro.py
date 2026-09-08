#!/usr/bin/env python3
"""
Afinidad temática entre investigadores, calculada desde mad-map-data-v2.xlsx.

Se usa en el documento propuesta-dos-lineas.md para justificar la composición
del claustro y de los grupos de estudio. Correr desde la raíz del repositorio:

    python3 tools/afinidad_claustro.py

Imprime la matriz de afinidad entre el claustro y los "en vía", el detalle de
cada par del claustro y las sublíneas, laboratorios y modos de cada persona.
"""
import collections
import itertools

import openpyxl

XLSX = "mad-map-data-v2.xlsx"

# Quiénes cumplen hoy el criterio de productividad y tienen grado de doctor.
CLAUSTRO = [
    "Anna Braghini", "Emanuela Di Felice", "Katherine Exss Cid",
    "Adriana Marín Toro", "Álvaro Mercado Jara", "Jaime Reyes Gil",
    "Daniela Salgado Cofré",
]
# Quiénes están a una o dos publicaciones del umbral, más Spencer (sin grado aún).
EN_VIA = [
    "Ursula Exss Cid", "Iván Ivelic Yanes", "Arturo Chicano Jiménez",
    "Andrés Garcés Alzamora", "Juan Carlos Jeldes Pontio",
    "Herbert Spencer González",
]

# Pesos de cada señal. Son ordinales: sirven para comparar pares, no como métrica absoluta.
PESO_SUBLINEA_COMPARTIDA = 1.0
PESO_PROXIMIDAD = 1.0        # multiplica la afinidad 0..1 declarada en 18_Proximidad_Tematica
PESO_LAB_COMPARTIDO = 0.5
PESO_MODO_COMPARTIDO = 0.25


def leer_hoja(wb, nombre):
    """Devuelve las filas de una hoja como diccionarios, usando como cabecera
    la primera fila que contiene 'investigador', 'nombre' o 'sublínea_a'.
    Se usa porque cada hoja del xlsx tiene una o dos filas de título antes de la cabecera."""
    filas = [f for f in wb[nombre].iter_rows(values_only=True) if any(c is not None for c in f)]
    for i, fila in enumerate(filas):
        celdas = [str(c) for c in fila]
        if any(k in celdas for k in ("investigador", "nombre", "sublínea_a")):
            cabecera = celdas
            break
    return [dict(zip(cabecera, f)) for f in filas[i + 1:]]


def cargar(xlsx=XLSX):
    """Carga las cuatro relaciones que alimentan la afinidad: temas (investigador-sublínea),
    laboratorios, modos y la matriz de proximidad entre sublíneas.
    Se usa una sola vez al inicio; devuelve estructuras indexadas por investigador."""
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    inv_sub = collections.defaultdict(set)
    for t in leer_hoja(wb, "08_Temas"):
        inv_sub[t["investigador"]].add(t["sublínea"])
    inv_lab = collections.defaultdict(set)
    for l in leer_hoja(wb, "12_Investigador_Lab"):
        inv_lab[l["investigador"]].add(l["laboratorio"])
    inv_mod = collections.defaultdict(set)
    for m in leer_hoja(wb, "13_Investigador_Modo"):
        inv_mod[m["investigador"]].add(m["modo"])
    prox = {}
    for p in leer_hoja(wb, "18_Proximidad_Tematica"):
        a, b, w = p["sublínea_a"], p["sublínea_b"], float(p["afinidad"])
        prox[(a, b)] = w
        prox[(b, a)] = w
    return inv_sub, inv_lab, inv_mod, prox


def afinidad(a, b, inv_sub, inv_lab, inv_mod, prox):
    """Puntaje de afinidad entre dos investigadores y sus componentes.
    Se usa para cada par de la matriz. Devuelve (total, sublíneas compartidas,
    suma de proximidades, labs compartidos, modos compartidos)."""
    sa, sb = inv_sub[a], inv_sub[b]
    compartidas = len(sa & sb)
    proximidad = sum(prox.get((x, y), 0) for x in sa for y in sb if x != y)
    labs = len(inv_lab[a] & inv_lab[b])
    modos = len(inv_mod[a] & inv_mod[b])
    total = (compartidas * PESO_SUBLINEA_COMPARTIDA + proximidad * PESO_PROXIMIDAD
             + labs * PESO_LAB_COMPARTIDO + modos * PESO_MODO_COMPARTIDO)
    return round(total, 2), compartidas, round(proximidad, 2), labs, modos


def abreviar(nombre):
    """Etiqueta corta (inicial más apellido truncado) para que la matriz quepa en consola."""
    partes = nombre.split()
    return partes[0][:1] + "." + partes[1][:6]


def main():
    datos = cargar()
    personas = CLAUSTRO + EN_VIA
    print(" " * 12 + " ".join(f"{abreviar(n):>9}" for n in personas))
    for a in personas:
        fila = [f"{afinidad(a, b, *datos)[0]:>9.2f}" if a != b else f"{'-':>9}" for b in personas]
        print(f"{abreviar(a):12}" + " ".join(fila))
    print("\nPares del claustro (total, sublíneas compartidas, proximidad, labs, modos):")
    for a, b in itertools.combinations(CLAUSTRO, 2):
        print(f"{a:22} {b:22} {afinidad(a, b, *datos)}")
    inv_sub, inv_lab, inv_mod, _ = datos
    print("\nSublíneas, laboratorios y modos por persona:")
    for n in personas:
        print(n, "|", "; ".join(sorted(inv_sub[n])), "| labs:", sorted(inv_lab[n]), "| modos:", sorted(inv_mod[n]))


if __name__ == "__main__":
    main()
