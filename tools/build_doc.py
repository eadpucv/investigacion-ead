#!/usr/bin/env python3
"""Genera lineas-investigacion.md leyendo mad-map-data-v2.xlsx directamente.

Documento institucional formal para presentar y fundamentar las cuatro
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
        "condicion": "de quién habita",
        "pregunta_nuclear": (
            "Cómo el diseño puede sostener la vida independiente, la "
            "comunicación, la participación y la autonomía de personas con "
            "condiciones diversas, incluida la pregunta cívica por la democracia "
            "y la comunicación ciudadana."
        ),
        "alcance_prosa": (
            "Esta línea concentra la investigación que se ocupa del sujeto que "
            "habita y de los sistemas que lo sostienen. Acoge la accesibilidad "
            "—cognitiva, sensorial, comunicacional— como condición de igualdad; "
            "la inclusión como horizonte ético y operativo del diseño; la "
            "comunicación aumentativa y alternativa como vía de participación; "
            "el diseño de interacción, los servicios y la experiencia de usuario "
            "como mediación cotidiana entre personas y sistemas; la vida "
            "independiente como capacidad cuya autonomía el diseño puede "
            "ampliar; y, en escala cívica, el diseño para la democracia y las "
            "plataformas de comunicación ciudadana."
        ),
    },
    "LIN-02": {
        "condicion": "de dónde se habita",
        "pregunta_nuclear": (
            "Cómo el territorio se construye, se habita, se sostiene y se "
            "piensa políticamente, en sus dimensiones urbana, ecológica, "
            "patrimonial y comunitaria."
        ),
        "alcance_prosa": (
            "Esta línea articula la investigación sobre el lugar del habitar. "
            "Cubre la ciudad y el territorio, las dinámicas de urbanización, la "
            "ecología política y los modos de adaptación frente a la "
            "vulnerabilidad —especialmente costera—, los riesgos y los desastres. "
            "Incluye la habitabilidad de la vivienda y sus crisis contemporáneas "
            "—financiarización, acceso, reuso—, el patrimonio arquitectónico y "
            "natural junto a sus prácticas de rehabilitación, la infraestructura "
            "urbana y la movilidad, y las prácticas colectivas, decoloniales y "
            "afectivas que operan sobre el territorio: urbanismo afectivo, deriva, "
            "geopoética, investigación-acción. Aborda también la evaluación social "
            "de políticas públicas de inversión y el confort, el bienestar y la "
            "habitabilidad personal como condiciones del habitar."
        ),
    },
    "LIN-03": {
        "condicion": "desde dónde se piensa",
        "pregunta_nuclear": (
            "Qué tradición y qué pensamiento sostienen la disciplina, cómo se "
            "actualizan, y qué categorías —del oficio, de la teoría, de la "
            "historia— articulan el proyecto contemporáneo."
        ),
        "alcance_prosa": (
            "Esta línea sostiene la investigación que actualiza la tradición "
            "disciplinar y opera sobre los pensamientos y las categorías del "
            "proyecto. Acoge el acervo histórico de Ciudad Abierta y de la "
            "Escuela de Arquitectura y Diseño, y dialoga con la teoría y la "
            "historia de la arquitectura —particularmente la moderna y "
            "latinoamericana— y del diseño. Trabaja en torno a categorías propias "
            "del oficio: el sentido de la hospitalidad, el vacío arquitectónico, "
            "la palabra poética como fundamento del proyecto y la poética del "
            "oficio. Aborda la relación entre arte y arquitectura, entre "
            "tecnología y sociedad, y la techné como reflexión epistemológica "
            "del proyecto."
        ),
    },
    "LIN-04": {
        "condicion": "con qué se hace y cómo se transmite",
        "pregunta_nuclear": (
            "Cómo la disciplina se hace, se enseña y se reproduce: con qué "
            "métodos, qué medios técnicos, qué oficio, y a través de qué "
            "espacios y dispositivos de enseñanza-aprendizaje."
        ),
        "alcance_prosa": (
            "Esta línea reúne las investigaciones que indagan en cómo la "
            "disciplina se hace, se transmite y se reproduce. Acoge los métodos "
            "del diseño y los saberes técnicos análogos junto a la fabricación "
            "digital, el modelado paramétrico y las prácticas de fablab; los "
            "medios de la comunicación visual, el diseño editorial y la "
            "exposición material; las máquinas expresivas, los algoritmos y el "
            "arte tecnológico; el mobiliario y la materialidad de la obra; la "
            "transferencia tecnológica orientada al emprendimiento local. Cubre "
            "también la enseñanza-aprendizaje del proyecto en sus distintas "
            "escalas: la arquitectura como medio didáctico, el diseño de "
            "espacios educativos —incluyendo contextos vulnerables y de "
            "estimulación temprana—, la formación en pensamiento y acción "
            "creativa, los programas y proyectos de arquitectura escolar, y la "
            "reforma de la enseñanza disciplinar."
        ),
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
        "cuatro líneas troncales del doctorado responde a una de las "
        "condiciones que esa pregunta convoca: la de *quién habita*, la de "
        "*dónde se habita*, la de *desde dónde se piensa* la disciplina, y la "
        "de *con qué medios se hace y cómo se transmite* el oficio.\n"
    )
    out.append(
        "La estructura preserva continuidad institucional con las áreas del "
        "postgrado —Extensión, Ciudad y Habitabilidad; Educación, Espacio y "
        "Aprendizaje; Forma, Cultura y Tecnología— y las articula como ejes "
        "ortogonales: los profesores afilian a una de las áreas como marco "
        "amplio, y las líneas concentran la pregunta de investigación. La "
        "interpretación de cada área se actualiza para acoger sin violencia "
        "los perfiles emergentes de las nuevas generaciones académicas.\n"
    )

    out.append("## Resumen de las cuatro líneas\n")
    out.append("| Línea | Condición que aborda | Sublíneas | Profesores |")
    out.append("|---|---|---:|---:|")
    for l in lineas:
        ctx = LINE_CONTEXT.get(l["id"], {})
        subs_de_linea = [s for s in sublineas if s["linea"] == l["id"]]
        invs_de_linea = set()
        for s in subs_de_linea:
            invs_de_linea |= sub_to_invs[s["id"]]
        out.append(
            f"| {l['nombre']} | *{ctx.get('condicion', '')}* "
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
        out.append(f"*Condición que aborda:* {ctx.get('condicion', '')}.\n")

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
        f"explícitamente declarada. El conjunto cubre las cuatro líneas "
        f"troncales del doctorado, con la siguiente distribución de "
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
