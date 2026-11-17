# PROMPT – Tablero vivo de indicadores del proceso Jurídico · HESEGO INGENIERÍA S.A.S.

> Copia todo este documento en el chat del asistente de VS Code (Claude Code, Copilot Chat en modo agente o Cursor), con los archivos de la carpeta `datos/` y `assets/` ya ubicados en el proyecto. Pídele que trabaje **por fases** y que no avance de fase sin mostrarte el resultado.

---

## 0. Rol y objetivo

Actúa como desarrollador full-stack senior y diseñador de visualización de datos. Vas a construir una **aplicación web local de presentación ejecutiva** para el proceso Jurídico de HESEGO INGENIERÍA S.A.S. que:

1. Lee directamente mis matrices de Excel (FO-GH-47 y FO-IG-136) y **se actualiza sola en tiempo real** cada vez que guardo uno de esos archivos, sin recargar la página.
2. Calcula los indicadores en cascada (operativo → táctico → estratégico) **para cualquier mes o año**, y también el acumulado del año.
3. Se presenta como **diapositivas** (modo presentación con flechas y pantalla completa) y también como **tablero navegable**.
4. Genera un **texto narrativo automático** por diapositiva que cambia según los resultados de cada corte (qué subió, qué bajó, qué incumple, qué requiere decisión).
5. Supera la presentación de referencia (informe de gestión en PowerPoint con donas de cumplimiento, META, BRECHA en p.p., N.º de operativos y estado Cumple / No cumple): debe tener todo eso, más tendencia, comparación con el mes anterior, profundización hasta el caso y narrativa automática.

El público es la Gerencia y la Junta Directiva: debe ser **muy entendible en 10 segundos por diapositiva**, sobrio, ejecutivo y con la identidad de Hesego.

---

## 1. Stack técnico (obligatorio)

- **Vite + React + TypeScript**.
- **Apache ECharts** (`echarts-for-react`) para gráficos: donas de cumplimiento, barras vs. meta, líneas de tendencia, cascada (waterfall) de contingencia, semáforos.
- **SheetJS (`xlsx`)** para leer los Excel. Leer **valores de las celdas fuente**, no depender de fórmulas: la app recalcula todo por su cuenta.
- **Servidor local Node** (Express + `ws` WebSocket) con **chokidar** vigilando la carpeta de datos. Al detectar un guardado del Excel, espera 1,5 s (debounce, porque Excel escribe en varias pasadas y crea archivos temporales `~$…`, que deben ignorarse), vuelve a leer, recalcula y envía el nuevo modelo a la página por WebSocket.
- **Zustand** para el estado (periodo seleccionado, diapositiva actual, modo presentación).
- `date-fns` con locale `es`.
- Pruebas con **Vitest** para el motor de cálculo.
- Un solo comando para arrancar: `npm run dev` (servidor + front) y otro para construir: `npm run build`.
- Configuración en `config/app.config.json` (rutas de archivos, ANS, colores, umbrales del semáforo), para que yo pueda cambiar parámetros sin tocar código.

La ruta de la carpeta de datos será la carpeta de OneDrive/SharePoint sincronizada en mi equipo, para que también se actualice cuando otra persona modifique el archivo en SharePoint.

---

## 2. Fuentes de datos

### 2.1 FO-GH-47 – Matriz de seguimiento jurídico (fuente principal, un caso por fila)

Archivo: `FO-GH-47_Matriz_de_Seguimiento_Juridico_V.122_automatica.xlsx`, hoja **`Proceso Jurídico`**, encabezados en la fila 1. Columnas (localízalas por nombre de encabezado normalizado, no por letra):

| Columna | Uso |
|---|---|
| No. | Id del caso |
| CLASIFICACIÓN DEL PROCESO JURÍDICO | Tipo de caso (ver 3.1) |
| NOMBRE / ASUNTO DE LA PETICIÓN / PROCESO QUE REMITE-SEDE | Detalle para la vista de casos |
| FECHA DE RECIBIDO | Inicio del término |
| TIPO DE TÉRMINO (HÁBIL/CALENDARIO), FECHA LÍMITE RESPUESTA | Términos legales |
| FECHA EN QUE DI RESPUESTA | Cierre. Define el mes de medición |
| ESTADO | CERRADO / EN GESTIÓN / PENDIENTE |
| RIESGO JURÍDICO | BAJO / MEDIO / ALTO (semáforo de alerta) |
| VALOR ESTIMADO DEL RIESGO | Exposición estimada (puede venir como texto: tratar como no numérico) |
| PENDIENTE DE REVISIÓN POR PROCESO ENCARGADO | SI / NO |
| POSICIÓN DE HESEGO | EN CONTRA (o vacío) / A FAVOR |
| VALOR PRETENDIDO | Lo que reclama la contraparte |
| SOLIDEZ DE LA DEMANDA, PRUEBAS EN CONTRA, JURISPRUDENCIA DESFAVORABLE, CONDUCTA DE LA CONTRAPARTE | Factores 1–5 (metodología FO-GH-48) |
| VALOR FINAL (a cargo de Hesego o recuperado) | Registrado al cierre |

La hoja `INDICADORES RIESGO` existe, pero **no la uses como fuente**: recalcula desde las columnas base para poder filtrar por cualquier periodo.

Limpieza obligatoria: `trim` + mayúsculas + quitar tildes en clasificación y estado; fechas que lleguen como texto (`dd/mm/aaaa`) o como número serial de Excel; filas sin clasificación se ignoran; registrar en un panel de **"Calidad de datos"** las filas con problemas (fecha de respuesta anterior a la de recibido, valor estimado en texto, abierto sin riesgo valorado, etc.).

### 2.2 FO-IG-136 – Matriz de indicadores en cascada (parámetros)

Archivo: `FO-IG-136_Indicadores_Juridico_*.xlsx` (tomar el más reciente de la carpeta), hoja `Mapa en cascada`, filas desde la 8. De aquí se leen **los nombres, pesos, metas, sentido (Mayor / Menor / Igual), frecuencia y fichas técnicas** de cada indicador. Estructura: columna G táctico (con su peso en F), L operativo, M peso operativo, O meta, P sentido, R–W ficha técnica. Los tácticos ocupan bloques de filas combinadas. **Nunca dejes pesos ni metas fijos en el código**: si yo cambio un peso en el Excel, la app debe reflejarlo al guardar.

### 2.3 Registro mensual (nuevo, lo creas tú)

Algunos operativos no salen de la FO-GH-47 (informes a gerencia, matriz legal FO-GC-35, plan de prevención). Crea `datos/registro_mensual.xlsx` con una hoja `REGISTRO` y columnas: `Año | Mes | Indicador operativo | Numerador | Denominador | Valor | Comentario del responsable`. La app lee de ahí esos operativos por periodo. El campo *Comentario del responsable*, si existe, se muestra junto al texto automático.

---

## 3. Reglas de cálculo (motor en `src/engine/`, con pruebas)

### 3.1 Agrupación por táctico

- **Acciones constitucionales y legales**: TUTELAS, DERECHO DE PETICION, DEMANDAS, EMBARGOS, GESTION JURIDICA, SINIESTROS, RENUNCIA MOTIVADA.
- **Revisión documental y disciplinarios**: REVISION DOCUMENTAL Y CONCEPTO JURIDICO, PROCESOS DISCIPLINARIOS.
- **Control del riesgo y la contingencia jurídica**: todos los casos, con el bloque de riesgo.
- **Cumplimiento normativo y prevención del daño jurídico**: registro mensual.

### 3.2 ANS internos (días hábiles), editables en config

DP 5 · Tutela 2 · Gestión jurídica 10 · Revisión documental y concepto 7 · Embargos 5 · Siniestros 8 · Demandas 15 · Renuncia motivada 2 · Procesos disciplinarios 5.

### 3.3 Días hábiles

`días = NETWORKDAYS(recibido, respuesta) − 1`, excluyendo sábados, domingos y **festivos de Colombia**. Implementa el cálculo de festivos para cualquier año (fijos, los trasladables al lunes por Ley 51 de 1983 y los que dependen de la Pascua), no una lista quemada. Prueba: los festivos de 2026 deben ser 1-ene, 12-ene, 23-mar, 2-abr, 3-abr, 1-may, 18-may, 8-jun, 15-jun, 29-jun, 20-jul, 7-ago, 17-ago, 12-oct, 2-nov, 16-nov, 8-dic, 25-dic.

### 3.4 Periodo

Selector de **Año** y **Mes** (o "Acumulado del año"). Un caso **cerrado** pertenece al mes de su FECHA EN QUE DI RESPUESTA. Los indicadores de **casos abiertos** (cobertura, vencidos, exposición) se calculan **al corte del último día del periodo**: abiertos = recibidos hasta esa fecha y sin respuesta a esa fecha (así puedo reconstruir cómo estaba cualquier mes pasado).

### 3.5 Operativos

1. **Cerradas dentro del ANS** (por táctico): casos cerrados en el periodo con días ≤ ANS de su clasificación ÷ cerrados en el periodo.
2. **Términos legales cumplidos** (tutelas y DP): respuesta ≤ fecha límite (o 15 días hábiles en DP y 2 en tutela si falta la fecha límite). Sin casos en el periodo = 100 % con etiqueta "sin casos".
3. **Tiempo promedio de respuesta**: promedio de días hábiles de los cerrados en el periodo (sentido Menor).
4. **Cobertura de la valoración del riesgo**: abiertos al corte con probabilidad ≠ "POR VALORAR" ÷ abiertos al corte. Probabilidad: si no hay valor estimado ni pretendido numéricos → POR VALORAR; si ambos son 0 → SIN CONTINGENCIA ECONÓMICA (cuenta como valorado); si hay valor y faltan factores → POR VALORAR; si no, % = suma de los 4 factores / 20 → < 50 % REMOTO, < 70 % POSIBLE, ≥ 70 % PROBABLE.
5. **Efectividad económica** (acumulado del año hasta el mes elegido): casos EN CONTRA cerrados con pretendido > 0 y valor final registrado → (Σ pretendido − Σ final) / Σ pretendido.
6. **Exposición y reserva** (seguimiento, peso 0): total pretendido de abiertos en contra; reserva = valor estimado (o pretendido si el estimado es 0) × % para PROBABLE (provisionar) y POSIBLE (nota de advertencia); REMOTO sin registro.
7. **Recuperación a favor** (seguimiento): casos A FAVOR cerrados, recuperado / reclamado.
8. Los demás operativos, del registro mensual.

### 3.6 Cascada

- Cumplimiento operativo: Mayor → `min(1, resultado/meta)`; Menor → `min(1, meta/resultado)` (si resultado = 0 → 1; si meta = 0 → 1 solo cuando resultado = 0); Igual → 1 o 0.
- Táctico = Σ (peso operativo × cumplimiento), **solo si** los pesos suman 100 % y todos los operativos con peso tienen dato; si no, mostrar "Sin medición completa" e indicar qué falta.
- Estratégico (Índice de cumplimiento ponderado de la gestión jurídica) = Σ (peso táctico × resultado táctico), con la misma regla.
- Opción en config `modoParcial: true` para recalcular el estratégico repartiendo el peso solo entre los tácticos con medición (marcarlo visiblemente como "parcial").

### 3.7 Pruebas de aceptación (deben pasar con los datos actuales, periodo octubre 2026)

- Acciones constitucionales y legales: 32 de 34 dentro del ANS (94,12 %); términos legales 7 de 7; tiempo promedio 2,15 días hábiles.
- Revisión documental y conceptos: 18 de 19 (94,74 %); disciplinarios 15 de 16 (93,75 %); tiempo promedio 1,74 días hábiles.
- Corte 05/11/2026: 15 procesos abiertos, cobertura 100 %, exposición en contra $3.724.044.271, a provisionar $24.500.000, en nota $15.000.000.

Si un resultado no cuadra, muéstrame la lista de casos que lo explican antes de cambiar la regla.

---

## 4. Diapositivas / vistas

Navegación con ← →, barra de miniaturas, `F` pantalla completa, `P` modo presentación (oculta controles), selector de periodo siempre visible arriba a la derecha. Cada diapositiva tiene: título, **mensaje clave** (una frase grande generada automáticamente), visual principal y texto narrativo breve.

1. **Portada**: proceso Jurídico, periodo, fecha y hora de la última actualización de los datos, y un indicador "EN VIVO" (punto verde que parpadea al recibir datos nuevos).
2. **Resumen estratégico**: dona grande del índice global con META, BRECHA (p.p.), estado (Cumple / En alerta / No cumple) y variación frente al mes anterior; debajo, las 4 tarjetas tácticas con dona, peso, N.º de operativos y estado; tendencia del índice en los 12 meses del año.
3. **Una diapositiva por táctico**: dona del táctico + tarjetas de cada operativo (VALOR vs META, sentido, mini barra de tendencia mensual con línea de meta, flecha de variación). Al hacer clic en un operativo, se abre un **panel lateral con los casos** que lo componen (los que incumplen, primero).
4. **Contingencia jurídica**: tarjetas de exposición, provisión y nota; **cascada** pretendido → ahorro → valor final; matriz de calor riesgo (BAJO/MEDIO/ALTO) × probabilidad (REMOTO/POSIBLE/PROBABLE); tabla de los procesos con valor, ordenada por reserva.
5. **Casos que requieren atención**: abiertos vencidos frente al ANS sin pendiente de otro proceso, casos POR VALORAR, términos legales próximos a vencer (≤ 2 días hábiles).
6. **Volumen y carga**: casos recibidos vs. cerrados por mes y por clasificación, por sede.
7. **Calidad de datos** (solo en modo tablero, no en presentación).
8. **Conclusiones y decisiones**: compila automáticamente los mensajes de alerta de todas las diapositivas + el comentario del responsable.

---

## 5. Narrativa automática (motor en `src/narrative/`)

- Basada en **reglas y plantillas determinísticas** (no IA externa), en español formal, impersonal y técnico; sin adjetivos innecesarios; cifras con formato colombiano (`$ 1.234.567`, `94,1 %`, `p.p.`).
- Para cada indicador calcula: estado vs. meta, brecha, variación vs. periodo anterior, tendencia de 3 meses (mejora / deterioro / estable) y causa principal (los casos que explican el incumplimiento).
- Prioriza: 1) incumplimientos, 2) deterioros, 3) riesgos económicos nuevos, 4) logros. Máximo 3 frases por diapositiva.
- Ejemplos de salida esperada:
  - "En octubre de 2026 el índice de gestión jurídica alcanzó 96,2 %, 3,8 p.p. por debajo de la meta, explicado por dos solicitudes de gestión jurídica que superaron el ANS de 10 días hábiles."
  - "La exposición en procesos abiertos asciende a $3.724 millones; el 92,6 % corresponde a un solo proceso laboral calificado como remoto, por lo que no exige provisión."
  - "Se requiere provisionar $24,5 millones por el siniestro de Natagaima, cuya probabilidad se calificó como probable (70 %)."
- Cuando falte información: "El táctico de cumplimiento normativo no se calcula este mes porque falta la evaluación semestral de la matriz legal."
- Botón "Editar texto" que permite sobrescribir la narrativa de una diapositiva para un periodo; se guarda en `datos/narrativa_manual.json` y se marca como "texto ajustado por el responsable".

---

## 6. Diseño visual

- Fondo de diapositiva: `assets/fondo_hesego.jpg` (identidad Hesego), con una capa blanca semitransparente para que el contenido se lea.
- Paleta: azul institucional `#0B1D2F` (títulos, barras principales), azul claro `#DDEBF7` (fondos de tarjetas), amarillo `#FFF2CC` (datos manuales), semáforo: verde `#2E9E5B` (cumple, ≥ meta), ámbar `#F2B134` (en alerta, hasta 5 p.p. por debajo), rojo `#D64545` (no cumple). Umbrales en config.
- Tipografía sans-serif limpia (Inter o Segoe UI), cifras grandes y tabulares.
- Formato 16:9, pensado para proyectar; también responsive para verlo en el celular.
- Animaciones suaves al cambiar de periodo (las donas y barras se mueven hasta su nuevo valor).
- Exportar: botón "PDF" (impresión con CSS por diapositiva, una página por diapositiva) y "PNG" de la diapositiva actual.

---

## 7. Estructura del proyecto

```
tablero-juridico/
  config/app.config.json
  datos/            (o ruta de OneDrive configurada)
  assets/fondo_hesego.jpg
  server/ watcher.ts, reader.ts, ws.ts
  src/
    engine/  (parser FO-GH-47, parser FO-IG-136, festivos, días hábiles, operativos, cascada) + tests
    narrative/
    slides/
    components/ (Dona, TarjetaOperativo, Tendencia, Cascada, MatrizCalor, PanelCasos, SelectorPeriodo)
    store/
```

---

## 8. Fases de trabajo

1. **Motor de cálculo + pruebas** (sin interfaz). Mostrar la tabla de resultados de octubre 2026 y que pasen las pruebas de aceptación.
2. **Servidor con lectura en vivo** (chokidar + WebSocket). Prueba: modifico y guardo el Excel y en menos de 3 s la página cambia.
3. **Diapositivas 1–3** con el diseño.
4. **Contingencia, casos de atención, volumen, calidad de datos**.
5. **Narrativa automática** y edición manual.
6. **Exportar PDF/PNG**, modo presentación, pulido visual.
7. (Opcional) **Publicación**: versión para que la Gerencia la vea sin mi equipo, desplegada como Azure Static Web App leyendo los Excel de SharePoint con Microsoft Graph e inicio de sesión de Microsoft 365.

Reglas generales: no inventes datos ni metas; si una columna no existe o viene vacía, muéstralo en "Calidad de datos" y sigue. Comenta el código en español. Antes de cada fase, dime qué vas a hacer en 3 líneas.
