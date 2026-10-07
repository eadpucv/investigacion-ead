#!/usr/bin/env python3
"""Genera lineas-investigacion.md desde investigacion-ead.xlsx.

Uso:  python3 tools/build_doc.py
"""
from datos import cargar, ROOT

OUT = ROOT / "lineas-investigacion.md"


def main():
    d = cargar()
    o = ["# Líneas de investigación\n", "**Doctorado en Arquitectura y Diseño**",
         "*Escuela de Arquitectura y Diseño · Pontificia Universidad Católica de Valparaíso*\n",
         "## Marco general\n",
         d["programa"].get("sello (texto)", "") + "\n"]
    o.append("## Resumen\n")
    o.append("| Línea | Modo de investigar | Prolonga el área del Magíster | Sublíneas (teóricas / proyectuales) | Profesores | Claustro | Proyectos 2014-2026 |")
    o.append("|---|---|---|---|---:|---:|---:|")
    info = {}
    for l in d["lineas"]:
        n = l["nombre"]
        S = [s for s in d["subs"] if s["línea"] == n]
        P = [p for p in d["profs"] if p.get("línea") == n]
        C = [p for p in P if p["claustro"]]
        Y = [p for p in d["proyectos"] if p["línea"] == n]
        info[n] = (S, P, C, Y)
        o.append(f"| {n} | {l.get('modo','')} | {l.get('prolonga (área del Magíster)','')} | {len(S)} ({sum(s['polo']=='teórico' for s in S)} / {sum(s['polo']=='proyectual' for s in S)}) | {len(P)} | {len(C)} | {len(Y)} |")
    o.append("")
    for l in d["lineas"]:
        n = l["nombre"]; S, P, C, Y = info[n]
        o.append(f"## {n}\n")
        o.append(f"*{l.get('modo','')}.* {l.get('definición breve','')}\n")
        if l.get("pregunta"):
            o.append(f"**Pregunta:** {l['pregunta']}\n")
        if l.get("descripción"):
            o.append(f"### Alcance\n\n{l['descripción']}\n")
        for polo in ("teórico", "proyectual"):
            ss = sorted([s["nombre"] for s in S if s["polo"] == polo])
            o.append(f"### Polo {polo} ({len(ss)} sublíneas)\n")
            o.append("; ".join(ss) + ".\n" if ss else "Sin sublíneas.\n")
        o.append("### Cuerpo académico\n")
        if C:
            o.append("**Claustro:** " + ", ".join(p["nombre"] for p in C) + ".\n")
        otros = [p["nombre"] for p in P if not p["claustro"]]
        if otros:
            o.append("**Profesores que aportan a la línea:** " + ", ".join(otros) + ".\n")
        o.append("### Consolidación\n")
        recientes = sum((p.get("año") or 0) >= 2022 for p in Y)
        o.append(f"La línea reúne {len(S)} sublíneas, {len(P)} profesores (de ellos {len(C)} en el claustro) y {len(Y)} proyectos de investigación y creación entre 2014 y 2026, {recientes} de ellos desde 2022."
                 + (f" La sostienen los laboratorios {l['laboratorios']}." if l.get("laboratorios") else "") + "\n")
    o.append("## Procedencia\n")
    o.append("Documento generado desde `investigacion-ead.xlsx` con `python3 tools/build_doc.py`. La línea de cada profesor se asigna por el tipo de su contribución (teórico-disciplinar o proyectual).\n")
    OUT.write_text("\n".join(o))
    print("OK:", OUT.name)


if __name__ == "__main__":
    main()
