# NFL Fantasy Analysis System — Implementation Plan

> **Para ejecutar este plan:** sigue las tareas en orden. Cada tarea tiene pasos concretos y un resultado esperado al final para que puedas verificar que quedó bien.

**Goal:** Construir el sistema completo de análisis fantasy (Google Sheets + Make.com + Canva) listo para usarse en el draft 2025 y alimentar contenido de redes sociales.

**Architecture:** Google Sheets como cerebro central con 8 tabs interconectados, Make.com para alertas automáticas de fuentes expertas, y Canva con 3 templates de posts listos para publicar.

**Tech Stack:** Google Sheets (IMPORTHTML, VLOOKUP, conditional formatting), Make.com (free tier), Canva (free tier)

---

## Fase 1 — Google Sheets: Estructura Base

### Tarea 1: Crear el Spreadsheet y todos los tabs

**Archivos:**
- Crear: Google Sheet llamado `NFL Fantasy 2025 — Analysis Hub`

- [ ] **Paso 1: Crear el spreadsheet**
  - Ir a [sheets.new](https://sheets.new) para crear un nuevo Google Sheet
  - Renombrar el archivo: clic en "Untitled spreadsheet" arriba a la izquierda → escribir `NFL Fantasy 2025 — Analysis Hub`

- [ ] **Paso 2: Crear los 8 tabs**
  - Clic en el `+` abajo a la izquierda para agregar tabs
  - Renombrar cada tab (doble clic en el nombre): crear estos 8 exactos:
    1. `Players DB`
    2. `Scoring Calc`
    3. `SoS`
    4. `OL Grades`
    5. `Schedule Flags`
    6. `OC Changes`
    7. `Big Board`
    8. `Content Hub`
  - Eliminar el "Sheet1" vacío si quedó

- [ ] **Paso 3: Verificar**

  Resultado esperado: en la barra de tabs de abajo deben aparecer los 8 tabs en ese orden exacto.

---

### Tarea 2: Tab `Players DB` — Base maestra de jugadores

**Archivos:**
- Modificar: tab `Players DB`

- [ ] **Paso 1: Crear headers en la fila 1**

  En el tab `Players DB`, clic en celda A1 y escribe cada header en las columnas siguientes:

  | Col | Header |
  |-----|--------|
  | A | Player |
  | B | Position |
  | C | Team |
  | D | Tier |
  | E | ADP_FantasyPros |
  | F | ADP_Underdog |
  | G | Expert_Notes |
  | H | OC_Change |
  | I | B2B_Weeks |
  | J | Primetime_Weeks |
  | K | Playoff_SoS |
  | L | OL_Grade |

- [ ] **Paso 2: Agregar validación de datos para columna B (Position)**
  - Seleccionar toda la columna B (clic en la letra B arriba)
  - Menú: `Data → Data validation`
  - Criteria: `List of items` → escribir: `QB,RB,WR,TE,K,D`
  - Guardar

- [ ] **Paso 3: Agregar validación de datos para columna D (Tier)**
  - Seleccionar columna D
  - `Data → Data validation` → `List of items` → `1,2,3,4,5`

- [ ] **Paso 4: Agregar validación para columna H (OC_Change)**
  - Seleccionar columna H
  - `Data → Data validation` → `List of items` → `Yes,No`

- [ ] **Paso 5: Agregar validación para columna K (Playoff_SoS)**
  - Seleccionar columna K
  - `Data → Data validation` → `List of items` → `Easy,Neutral,Hard`

- [ ] **Paso 6: Congelar fila de headers**
  - Clic en la fila 1 (número 1 a la izquierda)
  - Menú: `View → Freeze → 1 row`

- [ ] **Paso 7: Ingresar 5 jugadores de prueba para verificar**

  Ingresar estos datos en filas 2–6:

  | Player | Position | Team | Tier | ADP_FantasyPros | ADP_Underdog | OC_Change | Playoff_SoS |
  |--------|----------|------|------|-----------------|--------------|-----------|-------------|
  | Ja'Marr Chase | WR | CIN | 1 | 3.2 | 3.5 | No | Easy |
  | Christian McCaffrey | RB | SF | 1 | 1.1 | 1.0 | No | Neutral |
  | Lamar Jackson | QB | BAL | 1 | 8.5 | 8.1 | No | Hard |
  | Bijan Robinson | RB | ATL | 1 | 5.0 | 5.3 | No | Easy |
  | Amon-Ra St. Brown | WR | DET | 2 | 12.1 | 11.8 | No | Neutral |

- [ ] **Paso 8: Verificar**

  Resultado esperado: los dropdowns funcionan en las columnas B, D, H y K. Si escribes algo que no está en la lista, Sheets muestra una advertencia.

---

### Tarea 3: Tab `Scoring Calc` — Calculadora de 4 formatos

**Archivos:**
- Modificar: tab `Scoring Calc`

- [ ] **Paso 1: Crear headers**

  En el tab `Scoring Calc`, fila 1:

  | Col | Header |
  |-----|--------|
  | A | Player |
  | B | Receptions |
  | C | Rec_Yards |
  | D | Rec_TDs |
  | E | Carries |
  | F | Rush_Yards |
  | G | Rush_TDs |
  | H | Completions |
  | I | Pass_Yards |
  | J | Pass_TDs |
  | K | INTs |
  | L | PPR_Points |
  | M | HalfPPR_Points |
  | N | Standard_Points |
  | O | Custom_Points |

- [ ] **Paso 2: Ingresar fórmulas en fila 2**

  En celda L2, escribe esta fórmula (PPR):
  ```
  =(B2*1)+(C2*0.1)+(D2*6)+(F2*0.1)+(G2*6)+(I2*0.04)+(J2*4)+(K2*-2)
  ```

  En celda M2 (Half-PPR):
  ```
  =(B2*0.5)+(C2*0.1)+(D2*6)+(F2*0.1)+(G2*6)+(I2*0.04)+(J2*4)+(K2*-2)
  ```

  En celda N2 (Standard):
  ```
  =(C2*0.1)+(D2*6)+(F2*0.1)+(G2*6)+(I2*0.04)+(J2*4)+(K2*-2)
  ```

  En celda O2 (Custom — 0.8 PPR + 0.2 PPC):
  ```
  =(B2*0.8)+(C2*0.1)+(D2*6)+(E2*0.2)+(F2*0.1)+(G2*6)+(I2*0.04)+(J2*4)+(K2*-2)
  ```

- [ ] **Paso 3: Extender fórmulas hacia abajo**
  - Seleccionar celdas L2:O2
  - Copiar (Ctrl+C)
  - Seleccionar L3:O200
  - Pegar (Ctrl+V)

- [ ] **Paso 4: Ingresar datos de prueba para verificar las fórmulas**

  En fila 2, ingresar stats proyectadas de Ja'Marr Chase (temporada completa aproximada):

  | A | B | C | D | E | F | G | H | I | J | K |
  |---|---|---|---|---|---|---|---|---|---|---|
  | Ja'Marr Chase | 100 | 1400 | 10 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |

- [ ] **Paso 5: Verificar los cálculos manualmente**

  Con esos stats, los resultados esperados son:
  - **L2 (PPR):** `(100×1)+(1400×0.1)+(10×6) = 100+140+60 = 300 pts`
  - **M2 (Half-PPR):** `(100×0.5)+140+60 = 50+140+60 = 250 pts`
  - **N2 (Standard):** `140+60 = 200 pts`
  - **O2 (Custom):** `(100×0.8)+140+60 = 80+140+60 = 280 pts`

  Si los números en tu hoja coinciden → fórmulas correctas.

- [ ] **Paso 6: Probar con QB (Lamar Jackson) en fila 3**

  | A | B | C | D | E | F | G | H | I | J | K |
  |---|---|---|---|---|---|---|---|---|---|---|
  | Lamar Jackson | 0 | 0 | 0 | 60 | 800 | 5 | 380 | 4200 | 35 | 8 |

  Resultado esperado en O3 (Custom):
  `(0×0.8)+(0×0.1)+(0×6)+(60×0.2)+(800×0.1)+(5×6)+(380×0)+(4200×0.04)+(35×4)+(8×-2)`
  `= 0+0+0+12+80+30+0+168+140-16 = 414 pts`

  Esto demuestra el boost del 0.2 PPC en QBs corredores.

---

### Tarea 4: Tab `SoS` — Fuerza de calendario por posición

**Archivos:**
- Modificar: tab `SoS`

- [ ] **Paso 1: Crear estructura de headers**

  En la fila 1 del tab `SoS`:

  | Col | Header |
  |-----|--------|
  | A | Team |
  | B–S | Wk1 a Wk18 (una columna por semana) |
  | T | Playoff_Avg (Wks 15-17) |
  | U | Override_Notes |

  Para las semanas: en B1 escribe `Wk1`, en C1 `Wk2`, ... hasta S1 `Wk18`.

- [ ] **Paso 2: Agregar fórmula del promedio playoff en columna T**

  En T2, escribe:
  ```
  =AVERAGE(R2,S2,T2)
  ```
  Nota: las columnas R, S, T corresponden a Wk15, Wk16, Wk17.

  Ajuste: verifica cuáles columnas son Wk15–17 y usa esas letras. Si Wk1 = B, entonces Wk15 = P, Wk16 = Q, Wk17 = R.

  Fórmula correcta:
  ```
  =AVERAGE(P2,Q2,R2)
  ```

  Extender hacia abajo hasta fila 33 (32 equipos NFL).

- [ ] **Paso 3: Crear 4 sub-secciones por posición**

  Este tab se replica para cada posición. La forma más simple:
  - Fila 1: Headers
  - Filas 2–33: Equipos con SoS para **RB**
  - Dejar 2 filas vacías
  - Repetir para **WR**, **QB**, **TE**

  Agregar una celda fusionada de título antes de cada sección: "RB SoS", "WR SoS", etc.

- [ ] **Paso 4: Ingresar datos de SoS manualmente**

  Fuente: ir a [fantasypros.com/nfl/schedule/rb.php](https://www.fantasypros.com/nfl/schedule/rb.php)

  El color de cada semana en FantasyPros indica dificultad. Usar escala numérica:
  - Verde oscuro = 1 (muy fácil)
  - Verde = 2 (fácil)
  - Amarillo = 3 (neutral)
  - Naranja = 4 (difícil)
  - Rojo = 5 (muy difícil)

  > **Nota sobre IMPORTHTML:** FantasyPros usa JavaScript para renderizar su tabla de SoS, lo que bloquea IMPORTHTML. Los datos se ingresan manualmente una vez por temporada (30–45 min la primera vez, actualizaciones puntuales después).

- [ ] **Paso 5: Agregar color condicional automático**
  - Seleccionar el rango B2:R33 (todas las semanas para todos los equipos)
  - Menú: `Format → Conditional formatting`
  - Agregar 5 reglas:
    - Valor = 1 → fondo verde oscuro (#1a7b47), texto blanco
    - Valor = 2 → fondo verde (#34a853), texto blanco
    - Valor = 3 → fondo amarillo (#fbbc04), texto negro
    - Valor = 4 → fondo naranja (#fa7b17), texto blanco
    - Valor = 5 → fondo rojo (#d93025), texto blanco

- [ ] **Paso 6: Verificar**

  Resultado esperado: al ingresar un número 1–5 en cualquier celda del rango, el color cambia automáticamente.

---

### Tarea 5: Tab `OL Grades` — Calidad de línea ofensiva

**Archivos:**
- Modificar: tab `OL Grades`

- [ ] **Paso 1: Crear headers**

  | Col | Header |
  |-----|--------|
  | A | Team |
  | B | Run_Block_Rating |
  | C | Pass_Block_Rating |
  | D | Pressure_Pct_Allowed |
  | E | Expert_Consensus |
  | F | Notes |
  | G | Last_Updated |

- [ ] **Paso 2: Ingresar los 32 equipos en columna A**

  Lista de equipos (en orden alfabético por ciudad):
  ```
  ARI, ATL, BAL, BUF, CAR, CHI, CIN, CLE, DAL, DEN, DET, GB, HOU, IND,
  JAX, KC, LAC, LAR, LV, MIA, MIN, NE, NO, NYG, NYJ, PHI, PIT, SEA,
  SF, TB, TEN, WAS
  ```

- [ ] **Paso 3: Agregar validación de color para columna B y C**
  - Seleccionar B2:C33
  - `Format → Conditional formatting`
  - Escala de color: valor bajo = rojo, valor medio = amarillo, valor alto = verde
  - Usar escala numérica 1–10 (10 = mejor OL)

- [ ] **Paso 4: Fuente de datos para llenar**

  Ir a [pro-football-reference.com/years/2025/blocking.htm](https://www.pro-football-reference.com/years/2025/blocking.htm)

  Para Pressure_Pct_Allowed: buscar "Offensive Line Stats" en PFR. Menor % = mejor OL de pase.

  Para Run_Block y Pass_Block ratings: escala personal 1–10 basada en:
  - Yards before contact per carry (run block)
  - Sacks allowed / pressure % (pass block)
  - Referencias de Fantasy Footballers y Underdog

- [ ] **Paso 5: Verificar**

  Resultado esperado: la tabla tiene los 32 equipos con al menos las columnas A, D y G llenas.

---

### Tarea 6: Tab `Schedule Flags` — Banderas de calendario

**Archivos:**
- Modificar: tab `Schedule Flags`

- [ ] **Paso 1: Crear headers**

  | Col | Header |
  |-----|--------|
  | A | Team |
  | B | Bye_Week |
  | C | B2B_Away_Weeks |
  | D | Primetime_Count |
  | E | Primetime_Weeks |
  | F | TNF_Weeks |
  | G | Wk15_Opp |
  | H | Wk16_Opp |
  | I | Wk17_Opp |
  | J | Playoff_Schedule_Rating |

- [ ] **Paso 2: Ingresar los 32 equipos**

  Misma lista de equipos que en OL Grades.

- [ ] **Paso 3: Validación para columna J (Playoff_Schedule_Rating)**
  - `Data → Data validation` → `List of items` → `Elite,Good,Neutral,Bad,Brutal`

- [ ] **Paso 4: Fuente de datos**

  Ir a [nfl.com/schedules](https://www.nfl.com/schedules) o ESPN para el calendario completo 2025.

  Para cada equipo, identificar:
  - Semana de bye
  - Semanas con back-to-back partidos fuera de casa
  - Juegos en primetime (SNF/MNF/TNF marcados en el calendario)
  - Rivales en semanas 15, 16 y 17

- [ ] **Paso 5: Verificar**

  Resultado esperado: los 32 equipos tienen Bye_Week lleno y al menos Wk15–17 opponents anotados.

---

### Tarea 7: Tab `OC Changes` — Cambios de coordinador ofensivo

**Archivos:**
- Modificar: tab `OC Changes`

- [ ] **Paso 1: Crear headers**

  | Col | Header |
  |-----|--------|
  | A | Team |
  | B | Previous_OC |
  | C | New_OC |
  | D | Change_Type |
  | E | Affected_Positions |
  | F | Impact_Rating |
  | G | Notes |
  | H | Source |

- [ ] **Paso 2: Agregar validaciones**

  Columna D (Change_Type):
  - `Data → Data validation` → `List of items` → `Fired,Resigned,Promoted,Lateral`

  Columna F (Impact_Rating):
  - `Data → Data validation` → `List of items` → `Upgrade,Neutral,Downgrade`

- [ ] **Paso 3: Fuente de datos**

  Buscar en Google: "NFL offensive coordinator changes 2025" o revisar:
  - [FantasyPros OC changes article] — buscar "OC changes 2025 fantasy impact"
  - Fantasy Footballers episodios de offseason sobre cambios de staff

- [ ] **Paso 4: Verificar**

  Resultado esperado: al menos 5–8 equipos con cambios de OC ingresados con Impact_Rating lleno.

---

### Tarea 8: Tab `Big Board` — Vista principal de análisis

**Archivos:**
- Modificar: tab `Big Board`

- [ ] **Paso 1: Crear headers**

  | Col | Header |
  |-----|--------|
  | A | Player |
  | B | Position |
  | C | Team |
  | D | Tier |
  | E | Custom_Points |
  | F | PPR_Points |
  | G | HalfPPR_Points |
  | H | Standard_Points |
  | I | ADP_FantasyPros |
  | J | ADP_Underdog |
  | K | Playoff_SoS |
  | L | OC_Change |
  | M | Schedule_Rating |
  | N | Overall_Grade |
  | O | Notes |

- [ ] **Paso 2: Vincular columnas desde Players DB con VLOOKUP**

  En B2, fórmula para traer Position desde Players DB:
  ```
  =IFERROR(VLOOKUP(A2,'Players DB'!A:L,2,0),"")
  ```

  En C2, Team:
  ```
  =IFERROR(VLOOKUP(A2,'Players DB'!A:L,3,0),"")
  ```

  En D2, Tier:
  ```
  =IFERROR(VLOOKUP(A2,'Players DB'!A:L,4,0),"")
  ```

  En K2, Playoff_SoS:
  ```
  =IFERROR(VLOOKUP(A2,'Players DB'!A:L,11,0),"")
  ```

  En L2, OC_Change:
  ```
  =IFERROR(VLOOKUP(A2,'Players DB'!A:L,8,0),"")
  ```

- [ ] **Paso 3: Vincular scoring desde Scoring Calc**

  En E2 (Custom_Points):
  ```
  =IFERROR(VLOOKUP(A2,'Scoring Calc'!A:O,15,0),"")
  ```

  En F2 (PPR_Points):
  ```
  =IFERROR(VLOOKUP(A2,'Scoring Calc'!A:O,12,0),"")
  ```

  En G2 (HalfPPR_Points):
  ```
  =IFERROR(VLOOKUP(A2,'Scoring Calc'!A:O,13,0),"")
  ```

  En H2 (Standard_Points):
  ```
  =IFERROR(VLOOKUP(A2,'Scoring Calc'!A:O,14,0),"")
  ```

- [ ] **Paso 4: Vincular Schedule_Rating desde Schedule Flags**

  En M2:
  ```
  =IFERROR(VLOOKUP(C2,'Schedule Flags'!A:J,10,0),"")
  ```

- [ ] **Paso 5: Extender todas las fórmulas**
  - Seleccionar B2:M2
  - Copiar y pegar en B3:M200

- [ ] **Paso 6: Agregar validación para columna N (Overall_Grade)**
  - `Data → Data validation` → `List of items` → `A+,A,A-,B+,B,B-,C+,C,C-,D,F`

- [ ] **Paso 7: Agregar color condicional para Overall_Grade**
  - Seleccionar columna N
  - `Format → Conditional formatting`
  - `A+` o `A` = verde (#34a853)
  - `B+` o `B` = amarillo claro (#fbbc04)
  - `C` = naranja (#fa7b17)
  - `D` o `F` = rojo (#d93025)

- [ ] **Paso 8: Agregar filtros**
  - Seleccionar fila 1 completa
  - Menú: `Data → Create a filter`

- [ ] **Paso 9: Ingresar los 5 jugadores de prueba en columna A**

  Escribir los mismos nombres que en Players DB:
  - Ja'Marr Chase
  - Christian McCaffrey
  - Lamar Jackson
  - Bijan Robinson
  - Amon-Ra St. Brown

- [ ] **Paso 10: Verificar**

  Resultado esperado: las columnas B, C, D, K, L, M se llenan automáticamente desde los otros tabs. E, F, G, H traen los puntos calculados del Scoring Calc.

---

### Tarea 9: Tab `Content Hub` — Pipeline de publicaciones

**Archivos:**
- Modificar: tab `Content Hub`

- [ ] **Paso 1: Crear headers**

  | Col | Header |
  |-----|--------|
  | A | Insight_Title |
  | B | Player |
  | C | Category |
  | D | Analysis_Text |
  | E | Data_Points |
  | F | Source |
  | G | Content_Type |
  | H | Status |
  | I | Platform |
  | J | Date_Created |

- [ ] **Paso 2: Agregar validaciones**

  Columna C (Category):
  - `List of items` → `SoS,OL,Schedule,OC Change,Scoring Format,General`

  Columna F (Source):
  - `List of items` → `Fantasy Footballers,Underdog Fantasy,Pro Football Ref,Personal`

  Columna G (Content_Type):
  - `List of items` → `Rankings,Spotlight,Alert`

  Columna H (Status):
  - `List of items` → `Review,Draft,Ready,Published`

  Columna I (Platform):
  - `List of items` → `Instagram,TikTok,X,All`

- [ ] **Paso 3: Agregar color condicional para columna H (Status)**
  - `Review` = gris claro (#e8eaed)
  - `Draft` = amarillo (#fbbc04)
  - `Ready` = verde (#34a853)
  - `Published` = azul claro (#4285f4), texto blanco

- [ ] **Paso 4: Ingresar 2 insights de prueba**

  Fila 2:
  | Campo | Valor |
  |-------|-------|
  | Insight_Title | Chase — WR1 élite en playoff weeks |
  | Player | Ja'Marr Chase |
  | Category | SoS |
  | Analysis_Text | Chase enfrenta matchups de élite en semanas 15-17. Con .8 PPR, su alto target share lo convierte en un ancla de tu roster en playoffs. |
  | Data_Points | Wk15: vs MIA (fácil), Wk16: vs NE (fácil), Wk17: vs PIT (neutral) |
  | Source | Personal |
  | Content_Type | Spotlight |
  | Status | Draft |
  | Platform | All |
  | Date_Created | 2026-05-26 |

- [ ] **Paso 5: Verificar**

  Resultado esperado: los dropdowns funcionan en C, F, G, H, I. El color de la fila cambia según el Status.

---

## Fase 2 — Make.com: Alertas automáticas

### Tarea 10: Configurar Make.com y Flow de Fantasy Footballers

**Archivos:**
- Crear: Make.com scenario "FF New Content Alert"

- [ ] **Paso 1: Crear cuenta en Make.com**
  - Ir a [make.com](https://www.make.com) → Sign up gratis
  - Plan Free incluye 1,000 operaciones/mes — suficiente para este uso

- [ ] **Paso 2: Compartir el Google Sheet con Make**
  - En tu Google Sheet: `File → Share → Share with others`
  - Agregar el email de tu cuenta Make o dar acceso general (para el paso de conexión)

- [ ] **Paso 3: Crear nuevo Scenario**
  - En Make dashboard: clic en `Create a new scenario`
  - Nombre: `Fantasy Footballers — New Content Alert`

- [ ] **Paso 4: Agregar trigger RSS**
  - Clic en el círculo de inicio `+`
  - Buscar `RSS` → seleccionar `RSS Feed → Watch RSS feed items`
  - URL del feed Fantasy Footballers:
    ```
    https://feeds.megaphone.fm/GLT1412515089
    ```
  - Maximum number of results: `5`
  - Interval: `1 day` (revisa una vez al día)

- [ ] **Paso 5: Agregar acción Google Sheets**
  - Clic en `+` después del módulo RSS
  - Buscar `Google Sheets` → `Add a Row`
  - Conectar tu cuenta de Google
  - Spreadsheet: `NFL Fantasy 2025 — Analysis Hub`
  - Sheet: `Content Hub`
  - Mapear campos:
    - Columna A (Insight_Title): `{{RSS.title}}`
    - Columna B (Player): *(dejar vacío — llenar manualmente)*
    - Columna F (Source): `Fantasy Footballers`
    - Columna G (Content_Type): `Alert`
    - Columna H (Status): `Review`
    - Columna J (Date_Created): `{{now}}`

- [ ] **Paso 6: Activar el scenario**
  - Clic en el toggle de ON/OFF → activar
  - Clic en `Run once` para hacer una prueba manual

- [ ] **Paso 7: Verificar**

  Resultado esperado: aparecen filas nuevas en el tab `Content Hub` con episodios recientes de Fantasy Footballers, Status = "Review".

---

### Tarea 11: Flow de Underdog Fantasy

**Archivos:**
- Crear: Make.com scenario "Underdog New Content Alert"

- [ ] **Paso 1: Crear nuevo Scenario en Make**
  - Nombre: `Underdog Fantasy — New Content Alert`

- [ ] **Paso 2: Agregar trigger RSS**
  - Módulo: `RSS Feed → Watch RSS feed items`
  - URL del blog Underdog Fantasy:
    ```
    https://underdogfantasy.com/blog/rss.xml
    ```
  - Maximum: `5`, Interval: `1 day`

  > Si ese RSS no funciona, usar: `https://underdogfantasy.com/feed` o buscar "Underdog Fantasy RSS" para la URL actualizada.

- [ ] **Paso 3: Agregar acción Google Sheets (misma configuración)**
  - `Google Sheets → Add a Row → Content Hub`
  - Columna A: `{{RSS.title}}`
  - Columna F: `Underdog Fantasy`
  - Columna G: `Alert`
  - Columna H: `Review`
  - Columna J: `{{now}}`

- [ ] **Paso 4: Activar y probar**
  - Activar → `Run once`

- [ ] **Paso 5: Verificar**

  Resultado esperado: filas nuevas en Content Hub con artículos de Underdog, Source = "Underdog Fantasy".

---

## Fase 3 — Canva: Templates de publicaciones

### Tarea 12: Template A — Rankings / Listas

**Archivos:**
- Crear: Canva template "Rankings NFL Fantasy"

- [ ] **Paso 1: Crear cuenta Canva**
  - Ir a [canva.com](https://www.canva.com) → Sign up gratis con tu Gmail

- [ ] **Paso 2: Crear diseño nuevo**
  - `Create a design → Instagram Post (1080 × 1080 px)`
  - Nombre: `Template A — Rankings`

- [ ] **Paso 3: Diseño del template**

  Estructura visual:
  ```
  ┌─────────────────────────────┐
  │  [LOGO/NOMBRE CUENTA]  arriba izquierda │
  │                             │
  │  TOP 5 RBs                  │
  │  Mejor SoS Semanas 15-17    │
  │                             │
  │  1. [Nombre jugador]  🟢    │
  │  2. [Nombre jugador]  🟢    │
  │  3. [Nombre jugador]  🟡    │
  │  4. [Nombre jugador]  🟡    │
  │  5. [Nombre jugador]  🔴    │
  │                             │
  │  [Formato de scoring]       │
  └─────────────────────────────┘
  ```

  - Fondo: negro o azul marino oscuro (#0d1117 o #1a1f36)
  - Texto: blanco (#ffffff) para títulos, gris claro (#9ca3af) para subtítulos
  - Números en verde (#34a853), amarillo (#fbbc04), rojo (#d93025) según SoS

- [ ] **Paso 4: Marcar los textos como editables**
  - Cada nombre de jugador va en un Text Box separado
  - Etiquetar con comentarios en Canva: "Editar aquí"

- [ ] **Paso 5: Crear versión vertical (Stories/TikTok)**
  - `Resize → Instagram Story (1080 × 1920 px)`
  - Ajustar layout para formato vertical

- [ ] **Paso 6: Verificar**

  Resultado esperado: el template tiene placeholders claros para 5 jugadores, logo, título y formato. Al hacer clic en cada texto, se puede editar rápidamente.

---

### Tarea 13: Template B — Player Spotlight

**Archivos:**
- Crear: Canva template "Player Spotlight NFL Fantasy"

- [ ] **Paso 1: Crear nuevo diseño**
  - `Create a design → Instagram Post (1080 × 1080 px)`
  - Nombre: `Template B — Spotlight`

- [ ] **Paso 2: Diseño del template**

  ```
  ┌─────────────────────────────┐
  │  [LOGO]          [POSICIÓN] │
  │                             │
  │  [NOMBRE JUGADOR]           │
  │  [EQUIPO]                   │
  │                             │
  │  ✅ SoS: Easy (Wks 15-17)   │
  │  ✅ OL: Top 5 Run Block      │
  │  ✅ OC: Mismo sistema        │
  │  ⚠️  ADP: Sobrevalorado?     │
  │                             │
  │  "Análisis en 1 línea"      │
  │                             │
  │  [Custom: 280 pts] [PPR: 300]│
  └─────────────────────────────┘
  ```

  - Espacio para foto del jugador (lado derecho o fondo con opacidad reducida)
  - 3–4 bullets con iconos ✅ / ⚠️ / ❌
  - Strip inferior con puntos proyectados por formato

- [ ] **Paso 3: Verificar**

  Resultado esperado: el template tiene secciones claramente separadas para nombre, bullets de análisis y strip de scoring.

---

### Tarea 14: Template C — Alert / Breaking News

**Archivos:**
- Crear: Canva template "Alert NFL Fantasy"

- [ ] **Paso 1: Crear nuevo diseño**
  - `Create a design → Instagram Post (1080 × 1080 px)`
  - Nombre: `Template C — Alert`

- [ ] **Paso 2: Diseño del template**

  ```
  ┌─────────────────────────────┐
  │  🚨 BREAKING                │
  │                             │
  │  [TÍTULO DE LA NOTICIA]     │
  │  En mayúsculas, bold        │
  │                             │
  │  Qué significa para         │
  │  tu fantasy:                │
  │                             │
  │  → [Impacto 1]              │
  │  → [Impacto 2]              │
  │                             │
  │  [LOGO]    [FECHA]          │
  └─────────────────────────────┘
  ```

  - Fondo rojo oscuro (#7f1d1d) o amarillo de alerta (#b45309)
  - Texto grande, alto contraste
  - Diseño para lectura rápida en feed

- [ ] **Paso 3: Verificar**

  Resultado esperado: el template transmite urgencia, tiene placeholder para título y 2 bullets de impacto.

---

## Verificación Final del Sistema Completo

- [ ] **Check 1:** Ingresar un jugador nuevo en `Players DB` → aparece automáticamente en `Big Board` con todos sus datos

- [ ] **Check 2:** Cambiar una stat en `Scoring Calc` → el punto cambia en `Big Board` en las 4 columnas de scoring

- [ ] **Check 3:** Make.com corre una vez → aparece al menos 1 fila nueva en `Content Hub` con Status = "Review"

- [ ] **Check 4:** Tomar un insight del `Content Hub` con Status "Ready" → abrirlo en Template B de Canva → exportar como JPG → subir a Instagram (prueba manual del workflow completo)

---

## Referencia Rápida — Fórmulas Scoring

```
PPR     = (Rec×1.0) + (RecYds×0.1) + (RecTD×6) + (RushYds×0.1) + (RushTD×6) + (PassYds×0.04) + (PassTD×4) + (INT×-2)
Half    = (Rec×0.5) + (RecYds×0.1) + (RecTD×6) + (RushYds×0.1) + (RushTD×6) + (PassYds×0.04) + (PassTD×4) + (INT×-2)
STD     =             (RecYds×0.1) + (RecTD×6) + (RushYds×0.1) + (RushTD×6) + (PassYds×0.04) + (PassTD×4) + (INT×-2)
Custom  = (Rec×0.8) + (RecYds×0.1) + (RecTD×6) + (Car×0.2) + (RushYds×0.1) + (RushTD×6) + (PassYds×0.04) + (PassTD×4) + (INT×-2)
```

---

## Upgrades Futuros (cuando la cuenta tenga tracción)

1. **PFF $8/mo** → reemplazar OL grades manuales con datos de PFF (Pass Block Grade, Run Block Grade, ESPN grades)
2. **Airtable** → migrar Players DB para mejor UX de filtrado
3. **Softr** → convertir Big Board en web app pública para tu audiencia
4. **Sleeper API** → actualizar ADP automáticamente mid-season
