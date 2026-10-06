#!/usr/bin/env python3
"""Genera lineas-investigacion.md leyendo mad-map-data-v2.xlsx directamente.

Documento institucional formal para presentar y fundamentar las dos
líneas de investigación del Doctorado en Arquitectura y Diseño. No expone
codificaciones internas (LIN-XX, SUB-XX, INV-XX) ni el mapeo específico
profesor↔sublínea: documenta y justifica la consolidación y sostenibilidad
de cada línea.

Usa tools/xlsx_loader.py — el equivalente Python del loader del navegador —
para evitar duplicación de lógica de resolución por nombre y cómputo de
aristas.

Uso:
  python3 tools/build_doc.py
"""

import sys
from collections import defaultdict
from pathlib import Path

# Permitir importar xlsx_loader desde el mismo directorio
sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsx_loader  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "mad-map-data-v2.xlsx"
OUT = ROOT / "lineas-investigacion.md"


# Contexto narrativo por línea: lo que el documento institucional añade
# por encima de los datos brutos del .xlsx (preguntas nucleares, alcance
# en prosa, condición que la línea aborda). Mantener acá y no en la
# planilla porque es texto editorial cuidado.
LINE_CONTEXT = {
    "LIN-01": {
        "condicion": "investigación acerca del proyecto: los fundamentos de la disciplina y su transmisión",
        "origen": "Educación, Espacio y Aprendizaje",
        "pregunta_nuclear": (
            "Qué tradición, qué oficio y qué teoría sostienen el proyecto de "
            "arquitectura y de diseño, cómo se actualizan sus categorías, y con "
            "qué espacios, métodos y medios se enseña y se aprende a proyectar."
        ),
        "alcance_prosa": (
            "Esta línea prolonga el área Educación, Espacio y Aprendizaje del "
            "Magíster y se hace cargo de la investigación acerca del proyecto de "
            "arquitectura y de diseño: los fundamentos sobre los que se funda la "
            "disciplina (su historia, su teoría, su filosofía y sus métodos) y el "
            "modo en que se dominan y se transmiten en la formación. Su polo "
            "teórico e histórico acoge el acervo de Ciudad Abierta y de la "
            "Escuela, la historia y la crítica de la arquitectura moderna y "
            "latinoamericana y la historia del diseño, la teoría del proyecto y "
            "sus categorías propias (la hospitalidad, el vacío, la palabra "
            "poética y la poética del oficio), la relación entre arte, "
            "arquitectura y diseño, la reforma escolar y la enseñanza de la "
            "arquitectura, y la techné como reflexión epistemológica del diseño. "
            "Su polo proyectual acoge los espacios del aprendizaje (la "
            "arquitectura como medio didáctico, los espacios educativos y "
            "escolares, la estimulación temprana y los contextos vulnerables), "
            "los métodos y medios del diseño (comunicación visual, fabricación "
            "digital, máquinas expresivas, saberes técnicos, transferencia "
            "tecnológica) y el diseño de interacción, la accesibilidad y los "
            "sistemas inteligentes como mediación entre las personas y sus "
            "entornos. Es la línea que forma a quienes enseñarán arquitectura y "
            "diseño, desde el habitar poético que funda a la Escuela."
        ),
        "polo_teorico": "historia, crítica, acervo y teoría del proyecto en arquitectura y diseño",
        "polo_proyectual": "espacios, métodos y medios del aprendizaje del proyecto en arquitectura y diseño",
    },
    "LIN-02": {
        "condicion": "investigación a través del proyecto: obra nueva que se inscribe en un contexto y lo modifica",
        "origen": "Extensión, Ciudad y Habitabilidad",
        "pregunta_nuclear": (
            "Cómo la ciudad y el territorio se construyen, se habitan, se "
            "sostienen y se piensan políticamente, y qué proyectos de "
            "arquitectura y de diseño (urbanos, de equipamiento, de objetos y "
            "servicios) los transforman."
        ),
        "alcance_prosa": (
            "Esta línea prolonga el área Extensión, Ciudad y Habitabilidad del "
            "Magíster y se hace cargo de la investigación a través del proyecto de "
            "arquitectura y de diseño: obra nueva (un edificio, un espacio "
            "público, un objeto, un servicio, un sistema) que se inscribe en un "
            "contexto y lo modifica, con la ciudad, el territorio y sus ecologías "
            "como campo principal. Su polo teórico acoge la teoría urbana, la "
            "urbanización y la ecología política, la vivienda y sus crisis "
            "contemporáneas (financiarización, acceso, políticas habitacionales), "
            "los comunes y las resiliencias socioecológicas, las perspectivas "
            "decoloniales y la evaluación social de las políticas públicas de "
            "inversión. Su polo proyectual acoge la creación de obra sobre el "
            "territorio: la adaptación ante riesgos costeros y desastres, la "
            "infraestructura, la movilidad y el equipamiento, el patrimonio "
            "arquitectónico y natural y su rehabilitación, el urbanismo afectivo, "
            "la deriva y la investigación-acción, las prácticas colectivas y "
            "escénicas sobre el espacio público, el mobiliario y la materialidad "
            "de la obra, el diseño social y territorial, y el confort y el "
            "espacio habitable de las personas. Toma la región de Valparaíso y "
            "el continente americano como laboratorio y prolonga el legado de "
            "las travesías, donde la obra de arquitectura y de diseño es el modo "
            "de conocer el territorio."
        ),
        "polo_teorico": "teoría urbana, ecología política y vivienda",
        "polo_proyectual": "creación de obra urbana, arquitectónica y de diseño sobre el territorio",
    },
}

AREA_NAMES = {
    "ECH": "Extensión, Ciudad y Habitabilidad",
    "EAA": "Educación, Espacio y Aprendizaje",
    "FCT": "Forma, Cultura y Tecnología",
}


def professor_descriptor(inv, area_by_id):
    """Devuelve el nombre del profesor con su área principal entre paréntesis.
    Llamado por la sección "Cuerpo académico que la sostiene"."""
    area_id = inv.get("area_principal", "")
    area_name = area_by_id.get(area_id, area_id) if area_id else ""
    if area_name:
        return f"{inv['nombre']} ({area_name})"
    return inv["nombre"]


def main():
    data = xlsx_loader.load(XLSX)

    lineas = data["lineas"]
    sublineas = data["sublineas"]
    investigadores = data["investigadores"]
    laboratorios = data["laboratorios"]

    inv_by_id = {i["id"]: i for i in investigadores}
    lab_by_id = {l["id"]: l for l in laboratorios}
    area_by_id = {a["id"]: a["nombre"] for a in data["areas"]}

    # investigador → sublíneas (vía coautoría)
    inv_to_subs = defaultdict(set)
    sub_to_invs = defaultdict(set)
    for e in data["edges"]["coautoria"]:
        inv_to_subs[e["source"]].add(e["target"])
        sub_to_invs[e["target"]].add(e["source"])

    # línea → laboratorios (vía sostén-lab)
    linea_to_labs = defaultdict(set)
    for e in data["edges"]["sosten_lab"]:
        linea_to_labs[e["target"]].add(e["source"])

    out = []

    out.append("# Líneas de investigación\n")
    out.append("**Doctorado en Arquitectura y Diseño**")
    out.append("*Escuela de Arquitectura y Diseño · Pontificia Universidad Católica de Valparaíso*\n")

    out.append("## Marco general\n")
    out.append(
        "El programa forma investigadores para quienes la obra es origen y "
        "prueba de la tesis. La pregunta común que esa obra está llamada a "
        "argumentar es **cómo reinventar el habitar humano**. Cada una de las "
        "dos líneas de investigación del doctorado responde a esa pregunta desde "
        "el habitar poético que funda a la Escuela, y en ambas caben por igual la "
        "arquitectura y el diseño: la línea Fundamentos disciplinares investiga "
        "*acerca del proyecto* (su historia, su teoría, sus métodos y su "
        "transmisión); la línea Prácticas proyectuales investiga *a través del "
        "proyecto* (obra nueva que se inscribe en un contexto y lo modifica). Cada línea lleva "
        "dentro su contraparte: un polo teórico y un polo proyectual.\n"
    )
    out.append(
        "Las dos líneas prolongan las áreas Educación, Espacio y Aprendizaje y "
        "Extensión, Ciudad y Habitabilidad del Magíster en Arquitectura y "
        "Diseño, de modo que el tránsito entre ambos niveles sea legible. El "
        "área Forma, Cultura y Tecnología no origina línea: sus temas se "
        "distribuyen entre los polos proyectuales de ambas. Las áreas se "
        "conservan como marco amplio de afiliación de los profesores.\n"
    )

    out.append("## Resumen de las dos líneas\n")
    out.append("| Línea | Prolonga el área | Modo de investigar | Sublíneas | Profesores |")
    out.append("|---|---|---|---:|---:|")
    for l in lineas:
        ctx = LINE_CONTEXT.get(l["id"], {})
        subs_de_linea = [s for s in sublineas if s["linea"] == l["id"]]
        invs_de_linea = set()
        for s in subs_de_linea:
            invs_de_linea |= sub_to_invs[s["id"]]
        out.append(
            f"| {l['nombre']} | {ctx.get('origen', '')} | *{ctx.get('condicion', '')}* "
            f"| {len(subs_de_linea)} | {len(invs_de_linea)} |"
        )
    out.append("")

    for l in lineas:
        ctx = LINE_CONTEXT.get(l["id"], {})
        subs_de_linea = [s for s in sublineas if s["linea"] == l["id"]]

        invs_count = defaultdict(int)
        for s in subs_de_linea:
            for inv_id in sub_to_invs[s["id"]]:
                invs_count[inv_id] += 1
        invs_sorted = sorted(
            invs_count.items(),
            key=lambda kv: (-kv[1], inv_by_id[kv[0]]["nombre"]),
        )

        areas_in_line = set()
        for inv_id, _ in invs_sorted:
            ap = inv_by_id[inv_id].get("area_principal")
            if ap:
                areas_in_line.add(ap)

        labs_de_linea = sorted(
            lab_by_id[lab_id]["nombre"]
            for lab_id in linea_to_labs.get(l["id"], set())
        )

        out.append(f"## {l['nombre']}\n")
        out.append(f"*Modo de investigar:* {ctx.get('condicion', '')}.\n")
        out.append(f"*Prolonga el área del Magíster:* {ctx.get('origen', '')}.\n")
        out.append(f"*Polo teórico:* {ctx.get('polo_teorico', '')}. *Polo proyectual:* {ctx.get('polo_proyectual', '')}.\n")

        out.append("### Alcance\n")
        out.append(ctx.get("alcance_prosa", l["descripcion"]))
        out.append("")
        out.append(f"**Pregunta nuclear:** {ctx.get('pregunta_nuclear', '')}\n")

        out.append("### Cuerpo académico que la sostiene\n")
        out.append(
            f"Esta línea es cultivada por **{len(invs_sorted)} profesores** del "
            f"cuerpo académico de la Escuela de Arquitectura y Diseño:\n"
        )
        for inv_id, _ in invs_sorted:
            inv = inv_by_id[inv_id]
            out.append(f"- {professor_descriptor(inv, area_by_id)}")
        out.append("")

        out.append("### Consolidación y sostenibilidad\n")
        n_subs = len(subs_de_linea)
        n_invs = len(invs_sorted)
        n_areas = len(areas_in_line)
        argumentos = []
        argumentos.append(
            f"La línea está consolidada por la convergencia de **{n_invs} "
            f"profesores** activos que cultivan **{n_subs} sublíneas** "
            f"diferenciadas, lo que asegura masa crítica e indica una "
            f"distribución temática suficientemente amplia para acoger nuevas "
            f"tesis sin colapsar en un único objeto de estudio."
        )
        if n_areas == 1:
            area_label = AREA_NAMES.get(list(areas_in_line)[0], list(areas_in_line)[0])
            argumentos.append(
                f"La afiliación principal del cuerpo académico se concentra en "
                f"el área **{area_label}**, lo que da continuidad institucional "
                f"con la estructura del postgrado y profundidad disciplinar."
            )
        else:
            areas_label = ", ".join(
                AREA_NAMES.get(a, a) for a in sorted(areas_in_line)
            )
            argumentos.append(
                f"El cuerpo académico se distribuye entre **{n_areas} áreas** "
                f"del postgrado ({areas_label}), lo que da soporte transversal "
                f"a la línea y abre puentes con otras líneas del programa."
            )
        if labs_de_linea:
            argumentos.append(
                f"La línea cuenta con vínculos directos a "
                f"**{len(labs_de_linea)} laboratorio"
                f"{'s' if len(labs_de_linea) != 1 else ''}** "
                f"({', '.join(labs_de_linea)}), que operacionalizan la "
                f"investigación, la transferencia y la formación, garantizando "
                f"continuidad y proyección institucional."
            )
        for arg in argumentos:
            out.append(arg + "\n")

    out.append("## Cobertura del cuerpo académico\n")
    invs_total = len(investigadores)
    invs_con_mapeo = len([i for i in investigadores if i["id"] in inv_to_subs])
    invs_por_linea = defaultdict(set)
    for s in sublineas:
        for inv_id in sub_to_invs[s["id"]]:
            invs_por_linea[s["linea"]].add(inv_id)
    out.append(
        f"De los **{invs_total} profesores** del cuerpo académico, "
        f"**{invs_con_mapeo}** tienen al menos una sublínea de investigación "
        f"explícitamente declarada. El conjunto cubre las dos líneas "
        f"de investigación del doctorado, con la siguiente distribución de "
        f"profesores por línea (las afiliaciones pueden cruzarse: un mismo "
        f"profesor puede sostener sublíneas en más de una línea):\n"
    )
    out.append("| Línea | Profesores que la sostienen |")
    out.append("|---|---:|")
    for l in lineas:
        out.append(f"| {l['nombre']} | {len(invs_por_linea[l['id']])} |")
    out.append("")

    out.append("## Procedencia\n")
    out.append(
        "Documento generado automáticamente desde mad-map-data-v2.xlsx, la "
        "fuente única de verdad del programa. Las relaciones investigador↔sublínea "
        "provienen de los temas declarados por cada profesor en su perfil Casiopea "
        "o ANID, consolidados y curados por el equipo del doctorado. Para regenerar "
        "este documento tras editar la planilla, ejecutar `python3 tools/build_doc.py`."
    )

    OUT.write_text("\n".join(out))
    print(f"OK: {OUT}")
    print(f"  Líneas: {len(lineas)}")
    print(f"  Sublíneas: {len(sublineas)}")
    print(f"  Profesores con mapeo: {invs_con_mapeo}/{invs_total}")


if __name__ == "__main__":
    main()
