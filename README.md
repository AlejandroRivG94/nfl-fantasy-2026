# NFL Fantasy 2026 — Google Sheets Analysis Hub

Sistema de análisis para draft de fantasy football 2026.  
Formato: **0.8 PPR + 0.2 PPC (custom)** | Liga 12 equipos | 14 rondas | Playoffs Wks 15-17

---

## Flujo de trabajo y dependencias entre tabs

```
FUENTES EXTERNAS (usuario pega datos)
  FP_Raw          ← CSVs de FantasyPros (QB/RB/WR/TE)
  ESPN_Import     ← proyecciones ESPN (formato estándar)
  Yahoo_Import    ← proyecciones Yahoo (formato estándar)
  Vegas_Import    ← season props (disponible agosto)
        ↓
  parseFPProjections_2026()
        ↓
  FP_Import       ← stats de FP parseadas y normalizadas
        ↓
  updateScoringCalcFromConsensus_2026()
        ↓
┌─────────────────────────────────────┐
│  SCORING CALC  ← FUENTE DE VERDAD  │  stats raw + fórmulas Custom/PPR/Half/Std
└─────────────────────────────────────┘
        ↓                    ↓
  setupBigBoardFormulas()   buildProjectionComparison_2026()
  sortBigBoardByCustom()           ↓
        ↓              PROJECTION HUB ← FP vs ESPN vs Yahoo vs "Nuestro"
   BIG BOARD
        ↓
  addFullDraftAnalysis()
        ↓
  VOR + DRAFT STRATEGY
```

### ⚠️ Regla crítica
Cuando actualizas **Scoring Calc**, todos los tabs downstream necesitan refresh manual.  
Usa **`masterRefresh_2026()`** (en PROJECTIONS.gs) para sincronizar todo de una sola vez.

---

## Workflows

### A. Setup inicial (una sola vez)
```
1. updatePlayersDB_2026_VERIFIED()     → Players DB
2. updateScoringCalc_2026_VERIFIED()   → Scoring Calc (estimados manuales)
3. runAllVerified_2026()               → Big Board completo
4. addFullDraftAnalysis()              → VOR + Draft Strategy
5. runRemainingTabs_2026()             → SoS, OL Grades, Schedule Flags, OC Changes
6. setupProjectionHub_2026()           → crea tabs de importación
7. createDraftSimulator_2026()         → Draft Simulator
```

### B. Importar proyecciones de FantasyPros (cuando hay CSVs nuevos)
```
1. Pegar CSV QB/RB/WR/TE en FP_Raw (uno debajo del otro)
2. parseFPProjections_2026()           → FP_Import (datos normalizados)
3. updateScoringCalcFromConsensus_2026() → Scoring Calc actualizado con FP
4. masterRefresh_2026()                → sincroniza Projection Hub + Big Board + VOR
```

### C. Actualización rápida (solo cambió Scoring Calc)
```
masterRefresh_2026()  ← hace todo: Projection Hub + Big Board + VOR
```

### D. Actualización agosto (proyecciones pre-temporada)
```
1. Descargar CSVs de FP, ESPN, Yahoo
2. Repetir Workflow B completo
3. Agregar Vegas season props en Vegas_Import
4. masterRefresh_2026()
```

---

## Scripts

| # | Archivo | Funciones principales | Actualiza |
|---|---------|----------------------|-----------|
| 1 | `nfl_fantasy_2026_PLAYERS.gs` | `updatePlayersDB_2026_VERIFIED()` | Players DB |
| 2 | `nfl_fantasy_2026_SCORING.gs` | `updateScoringCalc_2026_VERIFIED()` | Scoring Calc |
| 3 | `nfl_fantasy_2026_BIGBOARD.gs` | `runAllVerified_2026()` · `setupBigBoardFormulas_2026()` · `sortBigBoardByCustom_2026()` | Big Board |
| 4 | `nfl_fantasy_2026_VOR.gs` | `addFullDraftAnalysis()` | VOR + Draft Strategy |
| 5 | `nfl_fantasy_2026_TABS.gs` | `runRemainingTabs_2026()` | SoS · OL Grades · Schedule Flags · OC Changes |
| 6 | `nfl_fantasy_2026_DRAFT_SIM.gs` | `createDraftSimulator_2026()` · `updateSimulatorForPosition_2026()` | Draft Simulator |
| 7 | `nfl_fantasy_2026_PROJECTIONS.gs` | `parseFPProjections_2026()` · `buildProjectionComparison_2026()` · `updateScoringCalcFromConsensus_2026()` · **`masterRefresh_2026()`** | Projection Hub · FP_Import · Scoring Calc |

---

## Estructura de tabs y su fuente de datos

| Tab | Fuente | Actualizado por |
|-----|--------|-----------------|
| Players DB | Manual / PLAYERS.gs | `updatePlayersDB_2026_VERIFIED()` |
| Scoring Calc | Consenso externo o SCORING.gs | `updateScoringCalcFromConsensus_2026()` o `updateScoringCalc_2026_VERIFIED()` |
| FP_Raw | Usuario pega CSV | Manual |
| FP_Import | FP_Raw | `parseFPProjections_2026()` |
| ESPN/Yahoo/Underdog_Import | Usuario pega datos | Manual |
| Vegas_Import | Sportsbooks (agosto) | Manual |
| **Projection Hub** | FP_Import + Scoring Calc | `buildProjectionComparison_2026()` |
| **Big Board** | Scoring Calc + Players DB | `setupBigBoardFormulas_2026()` + `sortBigBoardByCustom_2026()` |
| VOR / Draft Strategy | Big Board | `addFullDraftAnalysis()` |
| SoS / OL Grades / Schedule Flags / OC Changes | TABS.gs | `runRemainingTabs_2026()` |
| Draft Simulator | Big Board (ADP) | `createDraftSimulator_2026()` |

---

## Fórmula Custom Points

```
Custom = (Rec × 0.8) + (RecYds × 0.1) + (RecTDs × 6)
       + (Carries × 0.2) + (RushYds × 0.1) + (RushTDs × 6)
       + (PassYds × 0.04) + (PassTDs × 4) + (INTs × -2)
```

## VOR — Replacement Levels (con proyecciones FP mayo 2026)

| Posición | Jugador referencia | Custom Pts |
|----------|-------------------|-----------|
| QB #13 | Jordan Love / Baker Mayfield | ~280 pts |
| RB #37 | Último RB de flex | ~190 pts |
| WR #37 | Última WR de flex | ~165 pts |
| TE #13 | Dallas Goedert / Tucker Kraft | ~158 pts |

## Key Insights (con datos FP mayo 2026)

- **Bijan Robinson** (ADP 1.1) = 423 Custom Pts = #1 overall en formato .2PPC
- **Patrick Mahomes** (ADP 88.7) = 364 Custom Pts en R8 = mayor VOR por pick del draft
- **Drake Maye** (ADP 57) = breakout 2025 confirmado (4394py/31TD) = valor en R5
- **Lamar Jackson** (ADP 43) = FP proyecta conservador por lesión 2025 = riesgo/recompensa
- **TE cliff** después del pick 2-3 — McBride/Bowers en R2 = ventaja sostenida
- **Schedule Flags** (Wk15-17): verificar nfl.com/schedules/2026 antes del draft

## Cola de mejoras futuras

- **Coaching Hub** — expandir OC Changes a HC + DC + OC completo

---

*Proyecciones: FantasyPros Consensus (2 fuentes) — Mayo 2026*  
*Actualizar con proyecciones completas en agosto vía `masterRefresh_2026()`*
