#!/usr/bin/env python3
"""Genera comparacion-lineas.md desde investigacion-ead.xlsx (+ data/productividad.csv).

Combina texto editorial (aquí) con datos calculados: sublíneas por polo,
profesores, claustro, proyectos y productividad.
Uso:  python3 tools/build_comparacion.py
"""
from collections import defaultdict
from datos import cargar, ROOT

OUT = ROOT / "comparacion-lineas.md"
L1, L2 = "Fundamentos disciplinares", "Prácticas proyectuales"

EDIT = {
    "Relación con el proyecto": ("Investiga *acerca del* proyecto: el proyecto es su **objeto**", "Investiga *a través del* proyecto: el proyecto es su **método**"),
    "Tipo de contribución": ("Teórico-disciplinar: historia, teoría, crítica, métodos y su transmisión en la formación", "Proyectual: obra nueva (un edificio, un espacio público, un objeto, un servicio, un sistema, una herramienta) que se inscribe en un contexto y lo modifica"),
    "Focos del polo teórico": ("Acervo de la e[ad] y Ciudad Abierta; historia y crítica de la arquitectura moderna y latinoamericana; teoría del proyecto y sus categorías (hospitalidad, vacío, palabra poética); poesía y oficio; techné; reforma escolar", "Teoría urbana; urbanización y ecología política; vivienda y financiarización; comunes y resiliencias socioecológicas; perspectivas decoloniales; evaluación de políticas públicas"),
    "Focos del polo proyectual": ("Formación y enseñanza del proyecto: espacios educativos, arquitectura como medio didáctico, medios de aprendizaje, métodos de diseño, formación en pensamiento y acción creativa", "Ciudad y territorio (urbanismo afectivo, movilidad, equipamiento, patrimonio, ciudad-teatro, diseño social y territorial) y herramientas y sistemas (accesibilidad, interacción, comunicación aumentativa, IA, fabricación digital, comunicación visual)"),
    "Formas de investigar": ("Historiografía (archivo, fuentes, genealogías de obra); teoría crítica; investigación proyectual aplicada a la enseñanza", "Investigación proyectual; investigación-acción; creación de obra; desarrollo de herramientas y servicios; teoría crítica con trabajo de campo"),
    "Tipo de tesis": ("Tesis de historia, crítica o teoría del proyecto a partir de archivos y obras; investigación sobre la enseñanza del proyecto con obra docente como evidencia", "Exégesis de una obra, intervención, objeto, servicio o herramienta puesta a prueba; tesis de teoría urbana o vivienda con trabajo de campo; investigación-acción con comunidades"),
    "Competencias con más peso": ("C1 (posicionamiento frente a una tradición) y C3 (conceptualización del aporte)", "C4 (configuración de la obra) y C5 (rendición de cuenta ante quienes la reciben)"),
    "Interlocutores externos": ("Escuelas de arquitectura y diseño, redes de historia y teoría, archivos", "Estado, gobiernos regionales, municipios, organizaciones sociales, comunidades, industria"),
    "Incorporaciones próximas al claustro": ("Ursula Exss (a un capítulo con referato del umbral); Óscar Andrade (Fondecyt IR vigente)", "Iván Ivelic y Andrés Garcés (a un capítulo del umbral); Herbert Spencer al obtener el grado"),
    "Riesgo o brecha": ("Claustro en el mínimo; Arturo Chicano necesita un capítulo con referato más para cumplir el criterio individual; el polo proyectual (formación) no tiene académico en el claustro", "Concentra la mayoría de las sublíneas y proyectos (riesgo de dispersión); planificación urbana sin académico elegible; Mercado y Salgado necesitan publicaciones nuevas antes de 2025-2029"),
}


def main():
    d = cargar()
    subs, pry, prod = d["subs"], d["proyectos"], d["prod"]
    lin = {l["nombre"]: l for l in d["lineas"]}
    pry_sub = defaultdict(list)
    for p in pry:
        pry_sub[p["sublínea"]].append(p)

    def st(n):
        S = [s for s in subs if s["línea"] == n]
        P = [p for p in pry if p["línea"] == n]
        profs = [p for p in d["profs"] if p.get("línea") == n]
        cl = [p["nombre"] for p in profs if p["claustro"]]
        return dict(S=S, P=P, profs=profs, cl=cl,
                    cl_pub=sum(prod.get(x, {}).get("publicaciones_2022_2026", 0) for x in cl),
                    cl_pry=sum(prod.get(x, {}).get("proyectos_2022_2026", 0) for x in cl))
    A, B = st(L1), st(L2)
    row = lambda k, a, b: o.append(f"| **{k}** | {a} | {b} |")
    o = ["# Comparación de las líneas de investigación\n", "**Doctorado en Arquitectura y Diseño**",
         "*Escuela de Arquitectura y Diseño · Pontificia Universidad Católica de Valparaíso*\n",
         f"Las dos líneas se distinguen por el **tipo de contribución**: **{L1}** produce conocimiento teórico-disciplinar *acerca del* proyecto; **{L2}** produce obra *a través del* proyecto. Cada una tiene un polo teórico y un polo proyectual. La línea de cada sublínea y de cada profesor se asigna por ese criterio.\n",
         "> Generado por `tools/build_comparacion.py` desde `investigacion-ead.xlsx`. Los proyectos provienen del levantamiento 2014-2026 y se asignaron a una sublínea a partir del título; los marcados *por revisar* tienen asignación dudosa.\n",
         "## Tabla comparativa\n", f"| | **{L1}** | **{L2}** |", "|---|---|---|"]
    row("Relación con el proyecto", *EDIT["Relación con el proyecto"])
    row("Definición", lin[L1].get("definición breve", ""), lin[L2].get("definición breve", ""))
    row("Pregunta", lin[L1].get("pregunta", ""), lin[L2].get("pregunta", ""))
    row("Tipo de contribución", *EDIT["Tipo de contribución"])
    row("Prolonga (área del Magíster)", lin[L1].get("prolonga (área del Magíster)", ""), lin[L2].get("prolonga (área del Magíster)", ""))
    row("Focos del polo teórico", *EDIT["Focos del polo teórico"])
    row("Focos del polo proyectual", *EDIT["Focos del polo proyectual"])
    for X, k in ((A, "a"), (B, "b")):
        X["sub_txt"] = f"{len(X['S'])} ({sum(s['polo']=='teórico' for s in X['S'])} teóricas, {sum(s['polo']=='proyectual' for s in X['S'])} proyectuales)"
        X["pry_txt"] = f"{len(X['P'])} ({sum((p.get('año') or 0) >= 2022 for p in X['P'])} desde 2022)"
    row("Sublíneas", A["sub_txt"], B["sub_txt"])
    row("Proyectos 2014-2026", A["pry_txt"], B["pry_txt"])
    row("Profesores que aportan", len(A["profs"]), len(B["profs"]))
    row("Claustro", ", ".join(A["cl"]) or "—", ", ".join(B["cl"]) or "—")
    row("Productividad del claustro 2022-2026", f"{A['cl_pub']} publicaciones y {A['cl_pry']} proyectos", f"{B['cl_pub']} publicaciones y {B['cl_pry']} proyectos")
    for k in ("Formas de investigar", "Tipo de tesis", "Competencias con más peso", "Interlocutores externos"):
        row(k, *EDIT[k])
    row("Laboratorios", lin[L1].get("laboratorios", ""), lin[L2].get("laboratorios", ""))
    for k in ("Incorporaciones próximas al claustro", "Riesgo o brecha"):
        row(k, *EDIT[k])
    o.append("\nLa productividad cuenta publicaciones de cualquier tipo declaradas en el levantamiento; el cotejo contra los criterios del área (WoS/Scopus o capítulos con referato) está en `propuesta-dos-lineas.md`.\n")

    for n in (L1, L2):
        o.append(f"## {n}: sublíneas y proyectos\n")
        for polo in ("teórico", "proyectual"):
            S = sorted([s for s in subs if s["línea"] == n and s["polo"] == polo], key=lambda s: (-len(pry_sub[s["nombre"]]), s["nombre"]))
            o.append(f"### Polo {polo}\n")
            o.append("| Sublínea | Profesores | Proyectos | Desde 2022 |\n|---|---|---:|---:|")
            for s in S:
                ps = ", ".join(sorted(d["profs_sub"].get(s["nombre"], {}))) or "—"
                P = pry_sub[s["nombre"]]
                o.append(f"| {s['nombre']} | {ps} | {len(P)} | {sum((p.get('año') or 0) >= 2022 for p in P)} |")
            o.append("")
            for s in S:
                P = sorted(pry_sub[s["nombre"]], key=lambda p: -(p.get("año") or 0))
                if not P:
                    continue
                o.append(f"**{s['nombre']}**\n")
                for p in P:
                    o.append(f"- {p.get('año')} · {p['título']} — {', '.join(p['profesores'])}{' *(por revisar)*' if p['revisar'] else ''}")
                o.append("")

    o.append("## Profesores\n")
    o.append("| Profesor | Línea | Claustro | Publicaciones 2014-2026 | Publicaciones 2022-2026 | Proyectos 2014-2026 | Proyectos 2022-2026 |\n|---|---|:-:|---:|---:|---:|---:|")
    for p in sorted(d["profs"], key=lambda p: (p.get("línea") or "", p["nombre"])):
        q = prod.get(p["nombre"], {})
        o.append(f"| {p['nombre']} | {p.get('línea') or '—'} | {'sí' if p['claustro'] else ''} | {q.get('publicaciones_2014_2026','')} | {q.get('publicaciones_2022_2026','')} | {q.get('proyectos_2014_2026','')} | {q.get('proyectos_2022_2026','')} |")
    o.append("")
    OUT.write_text("\n".join(o))
    print("OK:", OUT.name, {L1: (len(A['S']), len(A['P'])), L2: (len(B['S']), len(B['P']))})


if __name__ == "__main__":
    main()
