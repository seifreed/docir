# ROADMAP - AST Completo para Office (OOXML + ODF)

Este roadmap define el camino para lograr un AST (IR) completo, estable y determinista que cubra **OOXML (DOCX/XLSX/PPTX)** y **ODF/OpenOffice (ODT/ODS/ODP)**. El objetivo es que **cualquier** documento Office quede representado en el IR sin perdida semantica relevante (aunque no sea un espejo 1:1 del XML).

Principios
- IR sobre XML: no exponer XML como API principal.
- Determinismo: misma entrada => mismo IR.
- Seguridad primero: todo contenido activo y referencias remotas modeladas explicitamente.
- Unificado: nodos comunes entre formatos, detalles especificos en nodos especializados.

---

## Fase 0 - Base y estabilidad [X]
Objetivo: base solida y segura para crecer sin cambiar el modelo.

Subpuntos
- Determinismo
  - [X] Normalizar orden de nodos en todo el pipeline
  - [X] Canonicalizar JSON (orden/valores)
- Trazabilidad
  - [X] Expandir SourceSpan (line/column)
  - [X] relationship_id en mas nodos
- Calidad
  - [X] Fixtures reales (DOCX/XLSX/PPTX)
  - [X] Validacion estructural minima por formato
- Performance
  - [X] Metricas de performance
  - [X] Limites por defecto configurables

---

## Fase 2 - Word (DOCX) completo [X]
Objetivo: cobertura completa de WordprocessingML.

Subpuntos
1) Body model y bloques
   - [X] Section properties completas
   - [X] Paragraph properties completas
   - [X] Run properties avanzadas
   - [X] Hyperlinks internos/externos con anchoring
   - [X] Tables con merges y grid completo
2) Parts estructurales
   - Headers/Footers por seccion
   - [X] Footnotes/Endnotes con anclajes
   - [X] Comments/Annotations completos
   - Bookmarks y fields de referencia
3) Content controls (SDT)
   - [X] SDT block/inline + data bindings
4) DrawingML/VML
   - [X] Shapes, textboxes, inline/floating
   - [X] Pictures/media con sizing/crop
   - [X] VML legacy completo
5) SmartArt y Charts
   - [X] Diagram data/layout/colors/style
   - [X] Charts embebidos o vinculados
6) Styles y numbering
   - [X] Styles.xml completo
   - [X] Numbering.xml completo
7) Fields y automatizacion
   - [X] Field instructions (DDE, INCLUDETEXT, HYPERLINK)
   - [X] AutoText/AutoCorrect markers
8) Revisions y tracking
   - [X] Track changes completo
9) Seguridad Word
   - [X] Macros docm, ActiveX, DDE
10) Relaciones externas
   - [X] attachedTemplate, links, OLE

---

## Fase 3 - PartRegistry y coverage detallado [X]
Objetivo: inventario mas preciso para coverage verificable con content-types reales.

Subpuntos
- Content-types completos
  - Word: numbering/comments/notes/header/footer/webSettings + macros
  - Excel: pivotCacheRecords + macros
  - PowerPoint: macro-enabled + OLE embeddings
- PartRegistry extendido
  - Patrones especificos para pivotCacheDefinition/pivotCacheRecords
  - Embeddings por formato (OLE)
  - .rels clave (comments/notes) con content-type relaciones

---

## Fase 4 - PPTX media/embeds expansion [X]
Objetivo: completar referencias a media embebida/enlazada en shapes.

Subpuntos
- [X] Soporte r:link en a:blip (audio/video enlazado)
- [X] Resolucion de targets externos e internos para media en shapes

---

## Fase 5 - DOCX DrawingML SmartArt/Chart mapping [X]
Objetivo: mejorar mapeo de DrawingML (Word) con targets normalizados y SmartArt relIds.

Subpuntos
- [X] Normalizar targets de DrawingML a paths de Word (word/...)
- [X] Capturar dgm:relIds como related_targets en shapes

---

## Fase 6 - Inventario package-level + relaciones [X]
Objetivo: coverage verifiable para properties y relaciones OOXML comunes.

Subpuntos
- [X] docProps core/app/custom en PartRegistry
- [X] customXml/*.xml como partes semanticas
- [X] relaciones (_rels/.rels y */_rels/*.rels) incluidas en coverage

---

## Fase 7 - Firmas digitales y customXml properties [X]
Objetivo: registrar partes de firmas y propiedades de customXml en coverage.

Subpuntos
- [X] _xmlsignatures/*.xml + origin.sigs en PartRegistry
- [X] customXml/itemProps*.xml con content-type correcto

---

## Fase 8 - OOXML coverage 100% verificable [X]
Objetivo: inventario exhaustivo por formato y coverage verificable con export.

Subpuntos
- [X] Inventario total por formato con paths reales y content-types
- [X] Matriz coverage (Part -> IR nodes -> tests)
- [X] Export de coverage (JSON/CSV)

---

## Fase 9 - ODF/OOo inventario y base de parsing [X]
Objetivo: habilitar soporte ODF con inventario completo y parsing seguro.

Subpuntos
1) Inventario ODF (ODT/ODS/ODP)
   - [X] ZIP + META-INF/manifest.xml
   - [X] content.xml, styles.xml, meta.xml, settings.xml
   - [X] Thumbnails/thumbnail.png
   - [X] mimetype (stored, no compression)
   - [X] Scripts y objetos embebidos (ObjectReplacements, Objects)
   - [X] RDF/metadata (meta.rdf si aplica)
2) PartRegistry ODF
   - [X] Content-type/mediatype segun manifest
   - [X] Clasificador por path + mediatype
3) Parser base ODF
   - [X] Lector ZIP con protecciones (zip bomb, path traversal)
   - [X] XML parsing con quick-xml, sin panics
   - [X] Normalizacion inicial a nodos comunes (Document/Section/Paragraph/Run/Table/Cell/Slide/Shape)

---

## Fase 10 - ODT completo (Writer) [X]
Objetivo: cobertura semantica completa de ODT.

Subpuntos
- [X] Text model: office:text, text:section, text:p, text:span
- [X] Lists y numbering (text:list, text:list-style)
- [X] Tables (table:table, table:table-row/cell)
- [X] Headers/footers (style:header/footer en page layouts)
- [X] Notes/comments (text:note, office:annotation)
- [X] Bookmarks/fields (text:bookmark, text:reference, text:date/time)
- [X] Drawings (draw:frame, draw:text-box, draw:image)
- [X] Styles (style:style, automatic styles, page layouts)
- [X] Track changes (office:change-info, text:tracked-changes)

---

## Fase 11 - ODS completo (Calc) [X]
Objetivo: cobertura semantica completa de ODS.

Subpuntos
- [X] Sheets y celdas (table:table, table:table-row/cell)
- [X] Formulas (table:formula) + referencias
- [X] Styles de celda/columna/fila
- [X] Validations y data filters
- [X] Charts (chart:chart) y drawing frames
- [X] External links y data pilot (pivot)

---

## Fase 12 - ODP completo (Impress) [X]
Objetivo: cobertura semantica completa de ODP.

Subpuntos
- [X] Slides y masters (draw:page, style:master-page)
- [X] Shapes/text (draw:frame, draw:text-box)
- [X] Images/media embebida/enlazada
- [X] Transitions/animations
- [X] Notes

---

## Fase 13 - Seguridad ODF [X]
Objetivo: modelar contenido activo y referencias remotas ODF.

Subpuntos
- [X] Scripts/macros (OpenOffice Basic, Python, Java)
- [X] External links y embedded objects

---

## Fase 14 - Cobertura observada (dataset con limites ajustados) [X]
Objetivo: documentar lo que ya parsea el parser con el dataset de OpenOffice.

Subpuntos
- ODT: texto/parrafos, listas basicas, tablas basicas, bookmarks, anotaciones, notas; estilos basicos (id/familia + propiedades text/paragraph); deteccion de objetos/links; firmas si existen
- ODS: hojas, celdas (valores), formulas como texto y evaluacion simple (refs/rangos/funciones basicas), merges, validaciones basicas, conditional formatting basico + prioridad/operador simple, pivots basicos + cache/records (conteo de campos), named ranges
- ODP: slides, shapes, texto en cuadros, imagenes, charts como shape, audio/video detectados, animaciones basicas, referencias externas
- Partes generales: meta.xml, styles.xml, settings.xml, manifest.xml, listados de partes, assets media, OLE links detectados

## Fase 15 - Cobertura parcial / limitada (gap parcial) [ ]
Objetivo: completar parsing parcial identificado en ODF.

Subpuntos
- Styles: se parsean IDs/familias y propiedades basicas de text/paragraph; faltan estilos completos de table/list y atributos avanzados (`crates/docir-parser/src/odf/mod.rs`)
- ODS formulas: evaluacion limitada a refs/rangos y funciones basicas (SUM/AVERAGE/MIN/MAX/COUNT); no hay motor completo (`crates/docir-parser/src/odf/mod.rs`)
- Conditional formatting: condicion + apply-style; operadores/prioridad simples, sin estilos detallados ni reglas avanzadas (`crates/docir-parser/src/odf/mod.rs`)
- Pivots: se detecta tabla y cache basica con fuente; sin campos detallados ni cache real (`crates/docir-parser/src/odf/mod.rs`)
- ODP layout: detecta shapes/texto/medios; sin layout avanzado ni render (`crates/docir-parser/src/odf/mod.rs`)
- Charts: detecta y extrae tipo/titulo; sin parse de series/datos (`crates/docir-parser/src/odf/mod.rs`)

## Fase 16 - Gaps claros (no soportado) [X]
Objetivo: cubrir los puntos que hoy no estan soportados en ODF.

Subpuntos
- [X] ODF encryption: decrypt solo si hay password y metadatos completos; no cubre variantes ni casos sin password (`crates/docir-parser/src/odf/mod.rs`)
- [X] Formulas complejas ODS (funciones avanzadas, matrices, referencias 3D completas, fechas/formatos): no evaluadas
- [X] Conditional formatting avanzado (prioridades complejas, estilos completos, multiples reglas/rangos detallados)
- [X] Pivot avanzado (cache real, campos, filtros, layouts)
- [X] Layout/multimedia avanzado en ODP (timings complejos, video/audio embebido con metadata detallada)
- [X] OLE en objetos embebidos
- [X] DDE/links en formulas (si aplica)
- [X] Digital signatures (META-INF/documentsignatures.xml)
- [X] Encryption/protected content (flag + metadatos)

---

## Fase 14 - Paridad OOXML/ODF en IR [X]
Objetivo: asegurar paridad semantica entre OOXML y ODF.

Subpuntos
- Mapeo consistente de nodos comunes
- Normalizacion de estilos y tablas
- Equivalencias de comentarios, notas, cambios, links
- Tests de equivalencia (mismo documento convertido OOXML/ODF)

---

## Fase 15 - Diff y rule engine (ambos formatos) [X]
Objetivo: diff semantico y reglas sobre IR unificado.

Subpuntos
- [X] Diff estructural
- [X] Diff de estilos
- [X] Reglas de seguridad
- [X] Export reportes

---

## Fase 16 - HWP/HWPX inventario y base de parsing [X]
Objetivo: habilitar soporte Hangul Word Processor (Hancom) con inventario completo y parsing seguro.

Subpuntos
1) Inventario HWP (legacy) y HWPX
   - HWP: contenedor OLE/CFB, stream inventory
   - HWPX: ZIP con XMLs (content/style/meta)
   - Recursos embebidos (imagenes, ole, binarios)
2) PartRegistry HWP/HWPX
   - Clasificador por stream/path + tipo
   - Mapeo de content-types/streams esperados
3) Parser base HWP/HWPX
   - Lector CFB (HWP) con protecciones
   - Lector ZIP (HWPX) con protecciones
   - XML parsing (HWPX) + normalizacion inicial a nodos comunes

---

## Fase 17 - HWP (legacy) completo [X]
Objetivo: cobertura semantica completa de HWP binario.

Subpuntos
- Estructura de documentos/sections/paragraphs/runs
- Tables, fields y estilos
- Embedded objects y OLE
- Seguridad: scripts/links/embedded content

---

## Fase 18 - HWPX completo [X]
Objetivo: cobertura semantica completa de HWPX (XML).

Subpuntos
- [X] Text model, sections, paragraphs, runs
- [X] Tables y styles
- [X] Drawings/images
- [X] Notes/comments
- [X] Track changes (si aplica)

---

## Fase 19 - Seguridad HWP/HWPX [X]
Objetivo: modelar contenido activo y referencias remotas en HWP/HWPX.

Subpuntos
- [X] Scripts/macros y autoexec
- [X] External links
- [X] OLE/embedded objects
- [X] Encrypted/protected content (flags y metadatos)

---

## Fase 20 - RTF inventario y parsing completo [X]
Objetivo: soporte completo para RTF con AST normalizado y seguro.

Subpuntos
- Inventario RTF
  - [X] Diccionario de control words (RTF spec) y grupos basicos
  - [X] Soporte de encodings (ANSI/Unicode/CodePage)
- Parser seguro
  - [X] Tokenizador RTF (grupos, control words, hex/escaped)
  - [X] Limites de profundidad y size (proteccion DoS)
- Normalizacion a IR
  - [X] Texto/paragraph/run
  - [X] Tables (\\trowd/\\cell)
  - [X] Fields (\\field/\\fldinst)
  - [X] Images (\\pict) y objetos (\\object)
- Seguridad RTF
  - [X] OLE/objdata
  - [X] External links y fields potencialmente activos
  - [X] Hyperlinks y autoexec conocidos
