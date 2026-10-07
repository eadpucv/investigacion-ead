# Especificaciones Allium — investigación e[ad]

Especificaciones conductuales del sistema de mapa de investigación del Doctorado en Arquitectura y Diseño (PUCV), escritas en [Allium v3](https://allium.sh).

## Archivos

| Archivo | Contenido |
|---------|-----------|
| `datos.allium` | Modelo de datos: entidades, relaciones, invariantes |
| `carga.allium` | Carga del XLSX en el navegador y derivación de relaciones calculadas |
| `visualizacion.allium` | Motor de visualización, estado del grafo, filtros, búsqueda, superficies |
| `generacion-documentos.allium` | Generación de los documentos `lineas-investigacion.md` y `comparacion-lineas.md` |

## Relación con el código

```
datos.allium          ←→  mad-map-data-v2.xlsx (20 hojas)
carga.allium          ←→  xlsx-loader.js · tools/xlsx_loader.py
visualizacion.allium  ←→  graph.js · cartografia.html · narrativa.html · exploracion.html
generacion-documentos ←→  tools/build_doc.py · tools/build_comparacion.py
```

## Las tres superficies

| Superficie | Audiencia | Perfil |
|------------|-----------|--------|
| Cartografía | Postulantes (público) | Solo estructura: las dos líneas con su definición, sus polos y sublíneas, sin investigadores |
| Narrativa | Evaluadores / CNA | Presets + controles acotados |
| Exploración | Equipo interno | Todos los controles disponibles |

## Pregunta de investigación

> ¿Cómo reinventar el habitar humano?

Las dos líneas de investigación responden a esa pregunta de dos maneras simétricas. **Fundamentos disciplinares** investiga *acerca del* proyecto: lo que sostiene a la disciplina (su historia, su teoría, su filosofía y sus métodos) y cómo se domina y se transmite en la formación. **Prácticas proyectuales** investiga *a través del* proyecto: obra nueva (un edificio, un espacio público, un objeto, un servicio, un sistema) que se inscribe en un contexto y lo modifica. Cada línea lleva dentro su contraparte: un polo teórico y un polo proyectual.
