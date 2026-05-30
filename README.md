# NFL Fantasy 2026 — Google Sheets Analysis Hub

Sistema de análisis para draft de fantasy football 2026.  
Formato: **0.8 PPR + 0.2 PPC (custom)** | Liga 12 equipos | 14 rondas | Playoffs Wks 15-17

## Scripts (Google Apps Script)

Correr en este orden desde **Extensions → Apps Script** en el Google Sheet:

| # | Archivo | Función | Tab |
|---|---------|---------|-----|
| 1 | `output/nfl_fantasy_2026_PLAYERS.gs` | `updatePlayersDB_2026_VERIFIED()` | Players DB |
| 2 | `output/nfl_fantasy_2026_SCORING.gs` | `updateScoringCalc_2026_VERIFIED()` | Scoring Calc |
| 3 | `output/nfl_fantasy_2026_BIGBOARD.gs` | `runAllVerified_2026()` | Big Board (nombres + grades + **VLOOKUPs B-M** + sort) |
| 4 | `output/nfl_fantasy_2026_VOR.gs` | `addFullDraftAnalysis()` | Big Board (VOR) + Draft Strategy |
| 5 | `output/nfl_fantasy_2026_TABS.gs` | `runRemainingTabs_2026()` | SoS + OL Grades + Schedule Flags + OC Changes |
| 6 | `output/nfl_fantasy_2026_DRAFT_SIM.gs` | `createDraftSimulator_2026()` | Draft Simulator (picks + builds + scenarios + stacks) |

> **Si el Big Board ya tiene nombres pero columnas B-M vacías:** corre solo `setupBigBoardFormulas_2026()` (en BIGBOARD.gs) para instalar los VLOOKUPs sin borrar los datos existentes.
>
> **Para cambiar la posición en el simulador:** edita la celda amarilla (B5) en el tab y corre `updateSimulatorForPosition_2026()`.

## Cola de mejoras futuras

- **Coaching Hub** — expandir OC Changes a HC + DC + OC. Impacto completo del coaching staff en cada equipo.

## Estructura

```
output/
├── nfl_fantasy_2026_PLAYERS.gs    ← 175+ jugadores, equipos verificados (FantasyPros)
├── nfl_fantasy_2026_SCORING.gs    ← Proyecciones custom .8PPR+.2PPC
├── nfl_fantasy_2026_BIGBOARD.gs   ← Big Board ordenado por Custom Points
├── nfl_fantasy_2026_VOR.gs        ← VOR, Draft Round Ideal, Value vs ADP
├── nfl_fantasy_2026_TABS.gs       ← SoS, OL Grades, Schedule Flags, OC Changes
└── nfl_fantasy_2026_FINAL.gs      ← Referencia (versión anterior, no usar)

docs/
└── superpowers/
    ├── plans/    ← Plan de implementación
    └── specs/    ← Especificaciones del sistema
```

## Fórmula Custom Points

```
Custom = (Rec × 0.8) + (RecYds × 0.1) + (RecTDs × 6)
       + (Carries × 0.2) + (RushYds × 0.1) + (RushTDs × 6)
       + (PassYds × 0.04) + (PassTDs × 4) + (INTs × -2)
```

## VOR — Replacement Levels

| Posición | Jugador | Custom Pts |
|----------|---------|-----------|
| QB #13 | Jordan Love / Baker Mayfield | ~308 pts |
| RB #37 | Último RB de flex (liga 12 eq) | ~190 pts |
| WR #37 | Última WR de flex | ~165 pts |
| TE #13 | Dallas Goedert / Tucker Kraft | ~158 pts |

## Key Insights

- **Josh Allen** (ADP 21) = ~+174 VOR sobre replacement QB = robo de ronda 2
- **Jalen Hurts** (ADP 66) = Ronda 6 con valor de Ronda 2 en formato .2PPC
- **TE cliff** después del pick 2-3 — McBride/Bowers en R2 = ventaja sostenida
- **Schedule Flags** (Wk15-17): verificar en nfl.com/schedules/2026 antes del draft

---

*Datos verificados vía FantasyPros Half-PPR ADP — Mayo 2026*
