# Roadmap

Documento de trabajo del proyecto **Investigación e[ad]**. Reemplaza al spec formal `mad-map.allium` y se mantiene en Markdown para edición cotidiana. Combina dos cosas: un estado del sistema (qué hace hoy) y una agenda accionable (qué falta, qué está descartado, qué decisiones siguen abiertas).

## Estado actual del sistema

El sistema produce un mapa interactivo del cuerpo investigativo del Doctorado en Arquitectura y Diseño. La fuente única de verdad es `mad-map-data-v2.xlsx`, commiteado en este repositorio. La visualización lo lee directamente en el navegador con SheetJS y dibuja el grafo con D3 force-directed. No hay servicios externos en tiempo de visualización ni scripts intermedios entre la planilla y la pantalla.

Tres superficies comparten el mismo motor pero exponen controles distintos según su audiencia. *Cartografía* es pública (postulantes), muestra solo líneas y sublíneas. *Narrativa* es para evaluadores y CNA, agrega la capa de profesores y dos vistas predefinidas. *Exploración* es la herramienta interna del equipo, con todos los controles disponibles. Una portada (`index.html`) enruta a las tres y muestra el sello formativo.

El modelo de datos cubre nueve entidades primarias[^1] más siete relaciones m:n[^2], todas referenciadas por nombre legible en las hojas de relación gracias a dropdowns dinámicos en Excel. Los IDs internos se mantienen pero el editor humano no los necesita para construir vínculos.

[^1]: Línea, Sublínea, Área, Modo, Salida, Laboratorio, Investigador, Tema, Sello formativo.
[^2]: Sublínea-Tema, Lab-Línea, Lab-Salida, Investigador-Lab, Investigador-Modo, Línea-Modo, Proximidad temática Sublínea-Sublínea.

## Roadmap accionable

### Listo

Migración a xlsx como fuente única de verdad. Visualización lee el .xlsx directo en el navegador. Cero servicios externos. Cero scripts intermedios. Layout force-directed orgánico.

Selectores desplegables dinámicos en columnas referenciales. Named ranges con `OFFSET`+`COUNTA` que crecen automáticamente con los datos. Editor humano nunca tipea un código a mano.

IDs de investigadores con iniciales (`INV-HSG`, `INV-MWU`...). Reconocibles a primera vista en cualquier hoja donde aparezcan. Script idempotente para regenerar.

Hojas de relación referenciadas por nombre. Las columnas `*_id` se renombraron a forma legible (`línea`, `área`, `laboratorio`, etc.) y guardan nombres en vez de códigos. El loader resuelve nombre→id internamente y emite warnings ante drift.

Fusión de `08_Temas` + `09_Sublinea_Tema` en una sola hoja con `(sublínea, investigador, tema)`. Eliminó la indirección por `tema_id`.

Limpieza de columnas espejo redundantes en `18_Proximidad_Tematica` (`sublinea_a_nombre`, `sublinea_b_nombre` ya no existen).

Documento institucional `lineas-investigacion.md` regenerable con `python3 tools/build_doc.py`, que ahora lee el xlsx directo a través de `tools/xlsx_loader.py`.

### Próximos pasos prioritarios

Estos son items vivos. Cada uno tiene estado, motivación, y un primer paso concreto si se decide abordar.

**Búsqueda extendida**[^q1]. Hoy la búsqueda matchea solo `nombre` y `descripcion`. Las sublíneas tienen `notas` en lugar de `descripcion` y la búsqueda no las cubre; los temas declarados (texto crudo) tampoco son indexados. Decisión pendiente: ¿extendemos el match a `notas` y al texto de los temas? Costo: trivial en `graph.js#_matchesSearch`. Beneficio: el equipo busca por palabras del tema sin tener que ir al xlsx.

[^q1]: Pregunta abierta heredada del spec original.

**URL state y vistas guardadas**. Cualquier estado de la viz (filtros, presets, nodo seleccionado, búsqueda) hoy se pierde al recargar. Implementar: serializar el `state` del `MadMapGraph` en el query string al cambiar y restaurarlo en `_initWithData`. Habilita compartir un link a una vista específica. Costo medio: ~80 líneas en `graph.js`. Beneficio: presentaciones reproducibles, debugging.

**Exportación SVG/PNG/JSON**. Botón en cada superficie que descarga el grafo actual. SVG es el más útil (vectorial editable). PNG para presentaciones. JSON para análisis externo. Costo: bajo para SVG (el grafo ya está en SVG, solo hay que serializar y descargar); medio para PNG (requiere `canvg` o equivalente).

**Preset "huecos del programa" en Narrativa**[^q1]. Vista que destaque sublíneas sin investigadores declarados o investigadores sin sublíneas mapeadas. Útil para curaduría editorial: identifica dónde falta poblar el mapa. Costo: bajo, agregar caso a `applyPreset` con highlight especial en lugar de filtrar.

**Aristas atenuadas o ocultas durante búsqueda**[^q1]. Hoy las aristas se atenúan junto con sus nodos cuando no calzan con la búsqueda. ¿Deberían desaparecer en cambio? Decisión visual; afecta la legibilidad del grafo en búsquedas con muchos resultados. Probar ambas y elegir.

**Sello formativo desde el xlsx en vivo**[^q1]. La portada ya carga el sello desde el xlsx (post-migración) — esta open question del spec original quedó resuelta como efecto secundario del paso 1.

### Descartado

**Modelo IV: embedding semántico + proyección 2D/3D**. Estaba marcado como `deferred` en el spec original. Decisión de mayo de 2026: descartado. El layout force-directed bien calibrado produce lecturas equivalentes o mejores que un PCA estructural, sin la complejidad inorgánica del pipeline de embeddings (matriz de features, SVD, normalización, fallback cuando faltan dimensiones, selector UI con tres algoritmos que no estaban implementados). El selector "UMAP / PCA / t-SNE" se removió de Exploración. La interfaz queda más limpia y honesta sobre lo que el sistema realmente hace.

**Autenticación / control de acceso**. Excluido explícitamente del scope. La visualización es pública, sin perfiles ni roles. La planilla es la pieza editable y vive bajo el control de acceso de git.

### Decisiones pendientes

**Semántica del peso en proximidad temática**[^q1]. La columna `afinidad` en `18_Proximidad_Tematica` guarda valores en 0..1, pero el peso no afecta hoy el layout (todas las aristas de proximidad pesan igual en el force-directed). Tres opciones: (a) el peso modula la fuerza de atracción del resorte; (b) el peso modula el grosor visual de la arista; (c) descartar el campo y usar solo presencia/ausencia. Recomendación: (a) y (b) son baratas y dan más expresividad; (c) simplifica el modelo. Conversación abierta.

**Política de errores ante referencias rotas**. El loader emite warnings en consola y omite filas con referencias colgadas. Para un editor en vivo conviene que estos warnings sean visibles en la propia UI (un toast o panel lateral), no solo en la consola del navegador. Costo: bajo. Beneficio: feedback inmediato al editor.

**Tests automatizados**. No existen. La idea: una suite mínima en Python que (i) valida invariantes del xlsx (todos los nombres únicos por entidad, cero referencias colgadas, proximidad simétrica, sublínea pertenece a una línea, modos predominantes consistentes); (ii) compara el output de `xlsx_loader.py` contra un snapshot esperado para detectar regresiones. Costo medio. Beneficio alto: la migración del paso 2 hubiera tenido red de seguridad.

## Anexo: contrato del comportamiento

Esta sección reemplaza al spec `mad-map.allium`. Documenta las invariantes y reglas que el sistema garantiza, en prosa Markdown. Si en el futuro se quiere volver a un lenguaje formal, este es el contrato a re-traducir.

### Invariantes del modelo

Toda sublínea pertenece a exactamente una línea troncal. Las sublíneas sin línea madre no existen en el modelo y serían filtradas por el loader.

La proximidad temática es simétrica: para cada par (A, B) con afinidad x debe existir el par (B, A) con la misma afinidad. Esto es responsabilidad del editor; el loader respeta lo que encuentre.

Las entradas de `18_Proximidad_Tematica` con `estado = DESCARTADO` se ignoran por completo (no generan aristas).

Toda envolvente (convex hull) requiere al menos 3 nodos visibles para renderizarse, porque un polígono necesita al menos 3 vértices.

Los nombres son únicos por entidad. Si dos entidades del mismo tipo comparten nombre, el loader emite warning y usa la primera ocurrencia. La unicidad la garantiza el editor humano (no hay constraint en Excel, solo convención).

### Invariantes de las superficies

*Cartografía* nunca expone investigadores como nodos. La capa de perfiles no es activable en esta superficie.

Las aristas visibles en *Cartografía* se restringen a `jerarquica` y `proximidad_semantica`.

Las aristas visibles en *Narrativa* se restringen a `jerarquica`, `coautoria`, `coinvestigacion` y `sosten_lab`.

Una arista solo es visible si ambos extremos están visibles. Filtrar un nodo oculta sus aristas como efecto secundario.

### Reglas de comportamiento

Al cargar cualquier superficie, el loader descarga `mad-map-data-v2.xlsx`, lo parsea y deriva las relaciones. Mientras carga, se muestra el texto del sello como indicador.

Al activar la capa de perfiles en *Narrativa* o *Exploración*, las aristas de coautoría se encienden automáticamente si la superficie las permite. Un investigador sin aristas a sus sublíneas quedaría flotando sin contexto.

Click sobre un nodo abre el panel de detalle. Click fuera de cualquier nodo lo cierra.

La búsqueda atenúa nodos no coincidentes pero no los oculta: el grafo sigue siendo legible globalmente.

Los presets de *Narrativa* (cobertura por línea, perfiles por área) configuran un estado fijo de aristas, envolventes y capa de perfiles, descartando el estado actual.

Cuando los filtros vacían el grafo (cero nodos visibles), aparece un estado vacío con un botón para limpiar filtros.

### Defaults por superficie

| Superficie | Aristas activas | Envolventes | Perfiles |
|---|---|---|---|
| Cartografía | jerárquica, proximidad semántica | área | no |
| Narrativa | jerárquica, coinvestigación | área | no |
| Exploración | jerárquica | área | no |
