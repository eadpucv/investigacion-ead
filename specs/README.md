# Especificaciones Allium — investigación e[ad]

Especificaciones conductuales del sistema de mapa de investigación del Doctorado en Arquitectura y Diseño (PUCV), escritas en [Allium v3](https://allium.sh).

## Archivos

| Archivo | Contenido |
|---------|-----------|
| `datos.allium` | Modelo de datos: entidades, relaciones, invariantes |
| `carga.allium` | Carga del XLSX en el navegador y derivación de relaciones calculadas |
| `visualizacion.allium` | Motor de visualización, estado del grafo, filtros, búsqueda, superficies |
| `generacion-documentos.allium` | Generación del documento institucional `lineas-investigacion.md` |

## Relación con el código

```
datos.allium          ←→  mad-map-data-v2.xlsx (18 hojas)
carga.allium          ←→  xlsx-loader.js · tools/xlsx_loader.py
visualizacion.allium  ←→  graph.js · cartografia.html · narrativa.html · exploracion.html
generacion-documentos ←→  tools/build_doc.py
```

## Las tres superficies

| Superficie | Audiencia | Perfil |
|------------|-----------|--------|
| Cartografía | Postulantes (público) | Solo estructura: líneas y sublíneas, sin investigadores |
| Narrativa | Evaluadores / CNA | Presets + controles acotados |
| Exploración | Equipo interno | Todos los controles disponibles |

## Pregunta de investigación

> ¿Cómo reinventar el habitar humano?

Las 4 líneas troncales responden a esa pregunta desde cuatro dimensiones: quién habita, dónde habitar, cómo pensarlo, con qué medios y cómo enseñarlo.
