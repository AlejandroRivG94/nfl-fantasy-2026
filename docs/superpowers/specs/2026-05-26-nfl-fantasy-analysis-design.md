# NFL Fantasy Analysis System — Design Spec
Date: 2026-05-26

## Overview

A no-code, low-cost system for analyzing NFL players in a custom fantasy format, designed to support both personal league decisions and social media content creation (Instagram, TikTok, X).

---

## Context & Goals

**League format:**
- Scoring: 0.8 PPR + 0.2 PPC (points per carry)
- Roster: 1 QB, 2 RB, 2 WR, 1 TE, 2 FLEX, 1 D, 1 K
- 12 teams, redraft (no keepers/dynasty)
- Fantasy playoffs: Weeks 15–17

**System goals:**
1. Centralize player analysis across 4 scoring formats: PPR, Half-PPR, Standard, Custom (0.8 PPR + 0.2 PPC)
2. Track key analysis factors: SoS, OL grades, schedule flags, OC changes
3. Generate ready-to-publish insights for social media

**Constraints:**
- No coding required — user learns as they go
- Low cost (free tools prioritized)
- Must support frequent manual updates as season info evolves

---

## Stack

| Tool | Role | Cost |
|------|------|------|
| Google Sheets | Analysis brain, scoring calculator, data hub | Free |
| Make.com | Alerts from Fantasy Footballers, Underdog Fantasy | Free tier |
| Canva | Social media templates and content export | Free tier |

---

## Data Sources

| Source | Data | Method |
|--------|------|--------|
| Pro Football Reference | OL stats, pressure rates, yards before contact | IMPORTHTML (auto) |
| FantasyPros | Expert rankings, SoS by position | IMPORTHTML (auto) |
| Underdog Fantasy ADP | Redraft ADP | IMPORTHTML (auto) |
| Sleeper ADP | Redraft ADP | IMPORTDATA (auto) |
| NFL.com | Schedule, bye weeks, primetime | IMPORTHTML (auto) |
| Fantasy Footballers | Expert opinion, rankings, OC analysis | Make.com alert → manual note |
| Underdog Fantasy content | Expert opinion, sleepers, busts | Make.com alert → manual note |

PFF OL grades: deferred until social media generates revenue (~$8/mo subscription).

---

## Google Sheets Architecture

### Tab: `Players DB`
Master player list. One row per player. Columns:

- Player name
- Position (QB / RB / WR / TE / K / D)
- Team
- Tier (1–5)
- ADP FantasyPros
- ADP Underdog
- Expert notes (free text — snippets from Footballers / Underdog)
- OC change flag (Yes / No)
- B2B week flag
- Primetime weeks
- Weeks 15–17 SoS rating (Easy / Neutral / Hard)
- OL grade (free text, from PFR until PFF is available)

### Tab: `Scoring Calc`
Input projected season stats once; formulas auto-calculate points in all 4 formats.

Inputs per player: receptions, receiving yards, rec TDs, carries, rushing yards, rush TDs, (QB: completions, passing yards, pass TDs, INTs)

Outputs: PPR points / Half-PPR points / Standard points / Custom points

Changing one stat updates all 4 formats automatically.

### Tab: `SoS`
Strength of schedule by position and week. Pulled from FantasyPros via IMPORTHTML.
- Columns: Team, Weeks 1–18 (color-coded: green = easy, red = hard)
- Weeks 15–17 highlighted separately for playoff analysis
- Manual override column for personal adjustments

### Tab: `OL Grades`
One row per NFL team. Tracks offensive line quality.

Columns: Team, Run Block Grade (PFR), Pass Block Grade (PFR), Pressure % allowed, Expert consensus (Footballers/Underdog notes), Last updated date

Updated manually each week or when significant news breaks.

### Tab: `Schedule Flags`
One row per team. Tracks schedule-based risk/opportunity.

Columns: Team, Bye Week, B2B weeks, Primetime games (count + weeks), Weeks 15–17 opponents, Overall schedule difficulty rating

### Tab: `OC Changes`
Tracks offensive coordinator changes and their projected impact.

Columns: Team, Previous OC, New OC, Change type (fired / resigned / promoted), Affected positions, Impact rating (Upgrade / Neutral / Downgrade), Notes, Source

### Tab: `Big Board`
The primary working view. Aggregates all analysis into a single ranked list.

- Filterable by: position, scoring format, tier, SoS rating, OC change flag
- Sortable by: any scoring format, ADP, custom rank
- Color-coded: green = buy, yellow = watch, red = avoid
- Auto-pulls from Players DB, SoS, Schedule Flags, OC Changes

### Tab: `Content Hub`
Pipeline for social media content. One row per publishable insight.

Columns:
- Insight title (e.g., "Chase — elite WR1 in playoff weeks")
- Player(s)
- Category (SoS / OL / Schedule / OC / Scoring Format / General)
- Analysis text (2–4 sentences, ready to use)
- Supporting data points (bullet list)
- Source (Footballers / Underdog / PFR / Personal)
- Content type (Rankings / Spotlight / Alert)
- Status (Draft / Ready / Published)
- Target platform (IG / TikTok / X / All)

---

## Make.com Automations

**Flow 1 — Fantasy Footballers new content alert**
Trigger: RSS feed or new YouTube video published
Action: Append row to Content Hub with title, link, date, status = "Review"

**Flow 2 — Underdog Fantasy new content alert**
Trigger: RSS feed or new article published
Action: Same as Flow 1

Both flows are read-only — they flag new content for manual review, not auto-publish.

---

## Canva Templates

### Template A — Rankings / Lists
Use case: "Top 5 RBs with best SoS weeks 15–17"
Format: Square (Instagram feed) + Vertical (Stories / TikTok)
Content source: Big Board filtered view → copy into template

### Template B — Player Spotlight
Use case: "Why X is a sleeper in your format"
Format: Square with player photo + 3–4 bullet analysis
Content source: Content Hub row (Spotlight type)

### Template C — Alert / Breaking News
Use case: "OC change — what it means for Y"
Format: Bold text-heavy, urgent feel
Content source: Content Hub row (Alert type), triggered by Make.com flag

All templates export directly from Canva to Instagram, TikTok, and X (free tier).

---

## Content Publishing Workflow

```
External source (Footballers / Underdog / PFR / personal analysis)
    ↓
Google Sheets — Players DB, SoS, OL Grades, Schedule Flags, OC Changes
    ↓
Big Board — filtered ranking view
    ↓
Content Hub — insight row marked "Ready"
    ↓
Canva — select template (A / B / C), paste content
    ↓
Export → Instagram / TikTok / X
```

---

## Scoring Format Reference

| Format | Reception | Carry | Notes |
|--------|-----------|-------|-------|
| PPR | 1.0 | 0 | Standard |
| Half-PPR | 0.5 | 0 | Most common |
| Standard | 0 | 0 | No reception bonus |
| Custom | 0.8 | 0.2 | User's league — rewards workhorse RBs + dual-threat QBs |

The Custom format particularly elevates: high-volume RBs (carries + receptions), dual-threat QBs (Lamar Jackson, Jalen Hurts, Josh Allen), pass-catching backs in PPR-adjacent roles.

---

## Key Analysis Factors

**SoS (Strength of Schedule)**
- Sourced from FantasyPros, filtered by position
- Weeks 15–17 are the primary lens (fantasy playoffs)
- Secondary lens: early-season matchups for fast starters

**OL Grades**
- Free: Pro Football Reference pressure %, run block success rate
- Future: PFF when budget allows
- Key for: RB value, QB protection, run game viability

**B2B & Primetime**
- B2B (back-to-back away games): fatigue risk, slight negative
- Primetime (MNF/SNF/TNF): historically conservative playcalling in prime games for some teams; specific TNF risk due to short week
- Tracked in Schedule Flags tab

**Fantasy Playoff Schedule (Wks 15–17)**
- Highest-weight factor for draft value
- Players with 3 easy matchups in wks 15–17 get significant bump
- Tracked in SoS tab with dedicated playoff column

**OC Changes**
- New OC = potential scheme shift → target share redistribution
- Tracked in OC Changes tab with affected positions noted
- Cross-referenced in Players DB with flag column

---

## Future Upgrades (when social media gains traction)

- PFF subscription ($8/mo) — unlock OL grades, route participation, target quality
- Airtable migration — better filtering UI as player database grows
- Softr web app — turn the Big Board into a public-facing tool
- Automated scoring updates mid-season via Sleeper API
