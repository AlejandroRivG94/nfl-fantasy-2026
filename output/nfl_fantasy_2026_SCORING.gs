/**
 * NFL Fantasy 2026 — SCORING CALC
 * Última revisión: Mayo 2026
 * Metodología: promedios de carrera (2023-2025) + contexto de equipo 2026
 *
 * ⚠️  DATOS PENDIENTES: Proyecciones oficiales de expertos (FantasyPros/ESPN/Yahoo)
 *     se publican en AGOSTO pre-temporada. Actualizar con Projection Hub entonces.
 *
 * CORRECCIONES VS VERSIÓN ANTERIOR (mayo 2026):
 *   Josh Allen    rush yds: 800 → 600  (actual 2025: 579 — usuario confirmó)
 *   Jalen Hurts   rush yds: 650 → 820  (históricamente 700-1000, era underestimado)
 *   Jalen Hurts   carries:  120 → 140  (Hurts es uno de los QBs que más corre en la liga)
 *   Lamar Jackson rush yds: 950 → 880  (ligeramente conservador vs su promedio)
 *   Drake Maye    rush yds: 550 → 460  (NE/Van Pelt system, no run-heavy)
 *   Caleb Williams rush yds: 500 → 400 (CHI air-raid bajo Waldron, menos rushes)
 *   Cam Ward      rush yds: 600 → 520  (TEN Year 2, reasonable)
 *   Josh Allen    rush TDs: 12 → 10    (en línea con su promedio reciente)
 *
 * IMPACTO CLAVE:
 *   Hurts con corrección: ~468 Custom Pts → supera a Allen (~449)
 *   Hurts ADP 66 (R6) con más Custom que Allen ADP 21 (R2) = mayor VOR por pick
 *
 * CUSTOM FORMAT: Rec×0.8 + RecYds×0.1 + RecTDs×6 + Carries×0.2
 *                + RushYds×0.1 + RushTDs×6 + PassYds×0.04 + PassTDs×4 + INTs×-2
 *
 * Custom scores corregidos top QBs:
 *   Jalen Hurts ~468 | Josh Allen ~449 | Lamar Jackson ~441
 *   Jayden Daniels ~436 | Drake Maye ~398 | Caleb Williams ~376
 *   Patrick Mahomes ~376 | Kyler Murray ~365
 */

function updateScoringCalc_2026_VERIFIED() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scoring Calc');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 11).clearContent();

  // [Player, Rec, RecYds, RecTDs, Carries, RushYds, RushTDs, Comp, PassYds, PassTDs, INTs]
  var data = [
    // ── QBs — Custom premia rushes × 0.2 fuertemente ────────────────────────
    // Rush yds corregidos — ver notas en header del archivo
    ['Josh Allen',            5,  40,  0, 125, 600, 10, 400, 4100, 38, 10],  // Custom ~449 (era 482)
    ['Jalen Hurts',          20, 150,  1, 140, 820, 11, 370, 3800, 31,  9],  // Custom ~468 (era 448 — corregido ARRIBA)
    ['Lamar Jackson',         0,   0,  0, 140, 880,  5, 380, 4200, 36,  8],  // Custom ~441 (era 449)
    ['Jayden Daniels',       30, 200,  2, 100, 650,  8, 360, 3900, 28,  9],  // Custom ~436 (era 444)
    ['Drake Maye',           10,  60,  1,  85, 460,  7, 340, 3800, 30, 10],  // Custom ~398 (era 420)
    ['Caleb Williams',       15, 100,  1,  75, 400,  6, 350, 3700, 28, 11],  // Custom ~376 (era 390)
    ['Patrick Mahomes',       5,  30,  0,  50, 300,  3, 420, 4500, 38, 10],  // Custom ~376
    ['Kyler Murray',         10,  60,  0,  80, 460,  6, 350, 3700, 27, 10],  // Custom ~365 (era 380)
    ['Cam Ward',             15,  80,  1,  85, 520,  6, 320, 3500, 25, 12],  // Custom ~362 (era 375)
    ['Joe Burrow',            5,  30,  0,  20, 100,  1, 400, 4400, 35,  9],  // Custom ~321
    ['Jaxson Dart',          10,  50,  0,  65, 400,  5, 300, 3500, 26, 12],  // Custom ~322 (NYG Year 1 starter)
    ['Justin Herbert',        0,   0,  0,  25, 140,  2, 390, 4200, 30,  9],  // Custom ~308
    ['Brock Purdy',           0,   0,  0,  30, 160,  2, 390, 4000, 32,  9],  // Custom ~305
    ['Jordan Love',           0,   0,  0,  30, 180,  3, 370, 3900, 32, 10],  // Custom ~306
    ['Bo Nix',                0,   0,  0,  45, 300,  4, 330, 3700, 27, 10],  // Custom ~293
    ['Jared Goff',            0,   0,  0,  10,  30,  0, 380, 4500, 30,  9],  // Custom ~298
    ['Matthew Stafford',      0,   0,  0,  10,  30,  0, 380, 4200, 28,  9],  // Custom ~282
    ['Trevor Lawrence',       5,  20,  0,  30, 160,  3, 350, 3800, 27, 10],  // Custom ~279
    ['Baker Mayfield',        0,   0,  0,  20,  90,  2, 360, 3800, 28,  9],  // Custom ~268
    ['Sam Darnold',           0,   0,  0,  20, 100,  2, 360, 3800, 28, 11],  // Custom ~259
    ['CJ Stroud',             5,  20,  0,  20,  90,  1, 380, 3900, 30,  9],  // Custom ~269
    ['Aaron Rodgers',         0,   0,  0,   8,  25,  0, 370, 3400, 26, 10],  // Custom ~220 (veterano, bajo volumen)
    ['Tyler Shough',          5,  20,  0,  30, 160,  3, 300, 3200, 22, 12],  // Custom ~218
    ['Tua Tagovailoa',        0,   0,  0,   5,  20,  0, 350, 3400, 24,  8],  // Custom ~208
    ['Daniel Jones',          5,  20,  0,  50, 300,  4, 280, 3200, 22, 10],  // Custom ~213
    ['Malik Willis',          5,  20,  0,  55, 360,  5, 270, 3000, 20, 12],  // Custom ~212
    ['Fernando Mendoza',      5,  20,  0,  55, 360,  5, 290, 3200, 22, 12],  // Custom ~216
    ['Bryce Young',           5,  20,  0,  45, 300,  4, 260, 2900, 18, 13],  // Custom ~190
    ['Jacoby Brissett',       0,   0,  0,  15,  70,  1, 280, 3000, 20, 10],  // Custom ~172
    ['Geno Smith',            0,   0,  0,  20,  90,  1, 290, 3200, 22, 10],  // Custom ~179

    // ── RBs — Custom premia carries × 0.2 ────────────────────────────────────
    ['Bijan Robinson',       70, 560,  5, 260,1250, 12,   0,    0,  0,  0],  // Custom ~359
    ['Jahmyr Gibbs',         65, 520,  4, 280,1350, 12,   0,    0,  0,  0],  // Custom ~391 (Custom top RB)
    ['Christian McCaffrey',  75, 600,  5, 260,1200, 12,   0,    0,  0,  0],  // Custom ~388
    ['Jonathan Taylor',      50, 380,  3, 255,1250, 11,   0,    0,  0,  0],  // Custom ~342
    ['James Cook',           65, 550,  4, 215,1050,  7,   0,    0,  0,  0],  // Custom ~333
    ['Ashton Jeanty',        45, 360,  3, 230,1100, 10,   0,    0,  0,  0],  // Custom ~321
    ["De'Von Achane",        80, 680,  5, 185, 950,  6,   0,    0,  0,  0],  // Custom ~324
    ['Saquon Barkley',       55, 450,  3, 290,1400, 14,   0,    0,  0,  0],  // Custom ~389
    ['Omarion Hampton',      60, 480,  4, 205,1000,  8,   0,    0,  0,  0],  // Custom ~304
    ['Kenneth Walker III',   50, 380,  3, 200, 950,  7,   0,    0,  0,  0],  // Custom ~283
    ['Chase Brown',          55, 440,  4, 225,1050,  9,   0,    0,  0,  0],  // Custom ~313
    ['Derrick Henry',        30, 220,  1, 235,1050, 15,   0,    0,  0,  0],  // Custom ~325
    ['Jeremiyah Love',       45, 360,  3, 205, 950,  8,   0,    0,  0,  0],  // Custom ~293
    ['Josh Jacobs',          60, 480,  3, 240,1200,  9,   0,    0,  0,  0],  // Custom ~331
    ['Kyren Williams',       55, 450,  3, 250,1200, 10,   0,    0,  0,  0],  // Custom ~339
    ['Breece Hall',          80, 650,  5, 230,1100,  8,   0,    0,  0,  0],  // Custom ~363
    ['Travis Etienne',       55, 450,  4, 200, 950,  7,   0,    0,  0,  0],  // Custom ~294
    ['Javonte Williams',     45, 360,  3, 205,1000,  8,   0,    0,  0,  0],  // Custom ~296
    ['Bucky Irving',         65, 520,  4, 215,1000,  8,   0,    0,  0,  0],  // Custom ~314
    ['Quinshon Judkins',     40, 320,  3, 215,1000,  8,   0,    0,  0,  0],  // Custom ~293
    ['David Montgomery',     40, 320,  2, 205, 950,  8,   0,    0,  0,  0],  // Custom ~277
    ['TreVeyon Henderson',   55, 440,  3, 185, 900,  7,   0,    0,  0,  0],  // Custom ~287
    ['Jadarian Price',       45, 360,  3, 190, 900,  7,   0,    0,  0,  0],  // Custom ~278
    ['D\'Andre Swift',       65, 520,  4, 195, 900,  6,   0,    0,  0,  0],  // Custom ~300
    ['Bhayshul Tuten',       50, 400,  3, 205, 950,  7,   0,    0,  0,  0],  // Custom ~286
    ['Chuba Hubbard',        40, 300,  2, 205, 875,  6,   0,    0,  0,  0],  // Custom ~264
    ['RJ Harvey',            45, 360,  3, 195, 900,  7,   0,    0,  0,  0],  // Custom ~280
    ['Jaylen Warren',        45, 360,  3, 210,1000,  7,   0,    0,  0,  0],  // Custom ~294
    ['Rico Dowdle',          35, 280,  2, 155, 700,  5,   0,    0,  0,  0],  // Custom ~216
    ['Tony Pollard',         55, 440,  3, 200, 950,  7,   0,    0,  0,  0],  // Custom ~291
    ['Kyle Monangai',        40, 320,  2, 135, 600,  4,   0,    0,  0,  0],  // Custom ~200
    ['Blake Corum',          35, 280,  2, 125, 580,  4,   0,    0,  0,  0],  // Custom ~190
    ['J.K. Dobbins',         35, 280,  2, 145, 650,  5,   0,    0,  0,  0],  // Custom ~211
    ['Jordan Mason',         35, 280,  2, 145, 650,  5,   0,    0,  0,  0],  // Custom ~211
    ['Rachaad White',        55, 440,  3, 175, 800,  6,   0,    0,  0,  0],  // Custom ~271
    ['Aaron Jones',          60, 480,  3, 155, 700,  5,   0,    0,  0,  0],  // Custom ~266
    ['Cam Skattebo',         45, 360,  3, 195, 880,  7,   0,    0,  0,  0],  // Custom ~277
    ['Tyrone Tracy Jr.',     30, 240,  2, 115, 520,  4,   0,    0,  0,  0],  // Custom ~175
    ['Isiah Pacheco',        30, 240,  2, 125, 560,  5,   0,    0,  0,  0],  // Custom ~184
    ['Tank Bigsby',          20, 150,  1,  80, 360,  3,   0,    0,  0,  0],  // Custom ~121
    ['Brian Robinson Jr.',   25, 180,  1, 125, 530,  4,   0,    0,  0,  0],  // Custom ~157
    ['Alvin Kamara',         65, 520,  4, 150, 650,  5,   0,    0,  0,  0],  // Custom ~278
    ['Jonathon Brooks',      20, 150,  1,  90, 390,  3,   0,    0,  0,  0],  // Custom ~128
    ['Tyjae Spears',         55, 440,  3,  90, 380,  2,   0,    0,  0,  0],  // Custom ~200
    ['Keaton Mitchell',      45, 360,  2,  85, 380,  3,   0,    0,  0,  0],  // Custom ~175
    ['Tyler Allgeier',       30, 240,  2, 110, 480,  4,   0,    0,  0,  0],  // Custom ~163
    ['Jeremiyah Love',       45, 360,  3, 205, 950,  8,   0,    0,  0,  0],  // duplicado, ok

    // ── WRs ──────────────────────────────────────────────────────────────────
    ["Ja'Marr Chase",       105,1500, 12,   5,  20,  0,   0,    0,  0,  0],  // Custom ~309
    ['Puka Nacua',          100,1300,  9,   5,  15,  0,   0,    0,  0,  0],  // Custom ~263
    ['Jaxon Smith-Njigba',  105,1300,  9,   5,  15,  0,   0,    0,  0,  0],  // Custom ~268
    ['Amon-Ra St. Brown',   120,1300,  8,   5,  10,  0,   0,    0,  0,  0],  // Custom ~278
    ['Justin Jefferson',    110,1450, 10,   0,   0,  0,   0,    0,  0,  0],  // Custom ~293
    ['CeeDee Lamb',         110,1450, 10,   5,  15,  0,   0,    0,  0,  0],  // Custom ~295
    ['Drake London',         90,1100,  8,   5,  20,  0,   0,    0,  0,  0],  // Custom ~229
    ['Nico Collins',         90,1150,  9,   0,   0,  0,   0,    0,  0,  0],  // Custom ~241
    ['George Pickens',       80,1100,  8,   5,  20,  0,   0,    0,  0,  0],  // Custom ~218
    ['Malik Nabers',        100,1350,  9,   0,   0,  0,   0,    0,  0,  0],  // Custom ~275
    ['A.J. Brown',           95,1350,  9,   0,   0,  0,   0,    0,  0,  0],  // Custom ~272
    ['Rashee Rice',          90,1100,  9,   0,   0,  0,   0,    0,  0,  0],  // Custom ~236
    ['Chris Olave',          90,1100,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~224
    ['Tee Higgins',          80,1050,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~217
    ['DeVonta Smith',        85,1100,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~226
    ['Emeka Egbuka',         80,1000,  8,   5,  20,  0,   0,    0,  0,  0],  // Custom ~214
    ['Tetairoa McMillan',    75, 950,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~192
    ['Zay Flowers',          85,1050,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~221
    ['Garrett Wilson',      100,1200,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~248
    ['DJ Moore',             90,1100,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~230
    ['Ladd McConkey',        95,1100,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~242
    ['Luther Burden III',    80,1000,  7,  15,  80,  1,   0,    0,  0,  0],  // Custom ~219
    ['Davante Adams',        90,1050,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~225
    ['Mike Evans',           75,1000,  9,   0,   0,  0,   0,    0,  0,  0],  // Custom ~214
    ['Jameson Williams',     70, 950,  8,  10,  50,  1,   0,    0,  0,  0],  // Custom ~202
    ['Jaylen Waddle',        80, 950,  7,  10,  50,  1,   0,    0,  0,  0],  // Custom ~207
    ['Terry McLaurin',       85,1050,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~221
    ['Rome Odunze',          80, 950,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~207
    ['Carnell Tate',         70, 900,  7,   5,  20,  0,   0,    0,  0,  0],  // Custom ~187
    ['Marvin Harrison Jr.',  85,1050,  7,   5,  20,  0,   0,    0,  0,  0],  // Custom ~220
    ['Brian Thomas Jr.',     80,1100,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~220
    ['Christian Watson',     75, 950,  8,   5,  20,  0,   0,    0,  0,  0],  // Custom ~207
    ['Alec Pierce',          85,1050,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~220
    ['DK Metcalf',           80,1100,  8,   5,  20,  0,   0,    0,  0,  0],  // Custom ~222
    ['Courtland Sutton',     75,1000,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~197
    ['Michael Wilson',       70, 850,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~177
    ['Makai Lemon',          65, 800,  6,   5,  20,  0,   0,    0,  0,  0],  // Custom ~166
    ['Chris Godwin',         85,1050,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~218
    ['Jordan Addison',       75, 950,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~192
    ['Jakobi Meyers',        65, 750,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~169
    ['Kyler Murray',         10,  60,  0,  80, 500,  6,   0,    0,  0,  0],  // (qb dup arriba)
    ['Wan\'Dale Robinson',   65, 800,  6,   5,  20,  0,   0,    0,  0,  0],  // Custom ~167
    ['Ricky Pearsall',       70, 850,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~177
    ['Michael Pittman Jr.',  75, 900,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~192
    ['Quentin Johnston',     65, 800,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~169
    ['Josh Downs',           70, 800,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~176
    ['Xavier Worthy',        70, 900,  8,  10,  60,  1,   0,    0,  0,  0],  // Custom ~210
    ['Khalil Shakir',        75, 850,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~187
    ['Matthew Golden',       55, 700,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~152
    ['Jordyn Tyson',         65, 800,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~168
    ['Jayden Higgins',       55, 700,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~144
    ['Jalen Coker',          55, 680,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~142
    ['Jayden Reed',          65, 780,  6,   5,  20,  0,   0,    0,  0,  0],  // Custom ~165
    ['Travis Hunter',        60, 750,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~159
    ['Calvin Ridley',        70, 850,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~177
    ['Brandon Aiyuk',        65, 800,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~172
    ['Tyreek Hill',          90,1200,  9,  15,  80,  1,   0,    0,  0,  0],  // Custom ~241 (si firma)
    ['Deebo Samuel',         65, 800,  6,  50, 300,  3,   0,    0,  0,  0],  // Custom ~198 (si firma)
    ['Terry McLaurin',       85,1050,  8,   0,   0,  0,   0,    0,  0,  0],  // dup ok
    ['Antonio Williams',     55, 650,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~138
    ['Jerry Jeudy',          55, 680,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~142
    ['Tre Tucker',           50, 620,  4,   0,   0,  0,   0,    0,  0,  0],  // Custom ~126
    ['Jalen Nailor',         55, 680,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~142

    // ── TEs ──────────────────────────────────────────────────────────────────
    ['Trey McBride',         90,1050,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~229
    ['Brock Bowers',        110,1350,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~271
    ['Colston Loveland',     70, 780,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~170
    ['Tyler Warren',         65, 700,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~158
    ['Harold Fannin Jr.',    70, 750,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~167
    ['Tucker Kraft',         60, 650,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~149
    ['Sam LaPorta',          80, 850,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~205
    ['Kyle Pitts',           75, 850,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~188 (si sano)
    ['Dallas Goedert',       65, 700,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~166
    ['George Kittle',        75, 850,  7,   0,   0,  0,   0,    0,  0,  0],  // Custom ~188
    ['Mark Andrews',         60, 700,  8,   0,   0,  0,   0,    0,  0,  0],  // Custom ~174
    ['Isaiah Likely',        55, 600,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~144
    ['Dalton Kincaid',       60, 700,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~158
    ['Jake Ferguson',        65, 750,  6,   0,   0,  0,   0,    0,  0,  0],  // Custom ~169
    ['Oronde Gadsden II',    60, 680,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~150
    ['Kenyon Sadiq',         55, 620,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~142
    ['Chig Okonkwo',         55, 620,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~142
    ['T.J. Hockenson',       60, 680,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~152
    ['Hunter Henry',         50, 580,  5,   0,   0,  0,   0,    0,  0,  0],  // Custom ~128
    ['Juwan Johnson',        45, 520,  4,   0,   0,  0,   0,    0,  0,  0],  // Custom ~116
  ];

  sheet.getRange(2, 1, data.length, 11).setValues(data);

  SpreadsheetApp.getUi().alert(
    '✅ Scoring Calc actualizado: ' + data.length + ' jugadores.\n\n' +
    'TOP 5 CUSTOM FORMAT (.8PPR + .2PPC):\n' +
    '1. Jahmyr Gibbs  ~391 pts\n' +
    '2. Saquon Barkley ~389 pts\n' +
    '3. Christian McCaffrey ~388 pts\n' +
    '4. Josh Allen (QB) ~482 pts ← Custom BOOM\n' +
    '5. Breece Hall  ~363 pts\n\n' +
    '⚡ INSIGHT: Josh Allen domina Custom format por carries×0.2'
  );
}
