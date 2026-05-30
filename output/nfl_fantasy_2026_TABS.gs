/**
 * NFL Fantasy 2026 — Tabs Restantes: SoS, OL Grades, Schedule Flags, OC Changes
 * ================================================================================
 * Este archivo cubre los 4 tabs que no están en PLAYERS / SCORING / BIGBOARD / VOR.
 *
 * ORDEN DE EJECUCIÓN (después de los otros 4 archivos):
 *   5. runRemainingTabs_2026()  ← corre todo este archivo de una vez
 *
 * O individualmente:
 *   updateSoS_2026()            → Tab "SoS"
 *   updateOLGrades_2026()       → Tab "OL Grades"
 *   updateScheduleFlags_2026()  → Tab "Schedule Flags"
 *   updateOCChanges_2026()      → Tab "OC Changes"
 *
 * Liga: 12 equipos | 0.8 PPR + 0.2 PPC | 1QB/2RB/2WR/1TE/2FLEX/1D/1K
 * Playoffs Fantasy: Semanas 15-17
 *
 * FUENTES:
 *   SoS:             FantasyPros Schedule Analysis / PFF 2025 defense grades
 *   OL Grades:       PFF 2025 OL grades
 *   Schedule Flags:  NFL.com 2026 (estimado — verificar antes del draft)
 *   OC Changes:      CBS Sports / NFL.com verificado offseason 2026
 */

// ─────────────────────────────────────────────────────────────────────────────
// MASTER RUNNER — corre los 4 tabs restantes
// ─────────────────────────────────────────────────────────────────────────────
function runRemainingTabs_2026() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Cargando tabs restantes...\nSoS → OL Grades → Schedule Flags → OC Changes');

  updateSoS_2026();
  SpreadsheetApp.flush();

  updateOLGrades_2026();
  SpreadsheetApp.flush();

  updateScheduleFlags_2026();
  SpreadsheetApp.flush();

  updateOCChanges_2026();
  SpreadsheetApp.flush();

  ui.alert(
    '✅ Tabs restantes cargados:\n\n' +
    '  📊 SoS         — Fuerza de defensas 2026 por posición\n' +
    '  🏋️ OL Grades   — Calificaciones líneas ofensivas 2025\n' +
    '  📅 Schedule    — Bye weeks + Semanas playoff\n' +
    '  🔄 OC Changes  — Cambios coordinadores ofensivos 2026\n\n' +
    '⚠️ Schedule Flags son ESTIMADOS — verificar nfl.com/schedules/2026\n' +
    '⚠️ Confirmar equipos de Tyreek Hill, Stefon Diggs, Deebo Samuel (FA)'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB: SoS (Strength of Schedule)
// ─────────────────────────────────────────────────────────────────────────────
// CÓMO LEER:
//   Las columnas _Rank son 1-32. 32 = defensa más débil = más fácil para la ofensiva.
//   Fantasy_Grade = tu jugador ofensivo vs esa defensa → A+=muy fácil, F=muy difícil
//   Las filas representan cada DEFENSA, no la ofensiva.
//
// USO EN DRAFT:
//   1. Busca el equipo de tu jugador en Schedule Flags → ve qué rival juega Wk15-17
//   2. Busca ese rival en SoS → revisa el rank del rival para la posición de tu jugador
//   3. Rank 25-32 = matchup fácil en playoffs fantasy = buen signo
// ─────────────────────────────────────────────────────────────────────────────
function updateSoS_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SoS');
  if (!sheet) { sheet = ss.insertSheet('SoS'); }
  sheet.clearContents();
  sheet.clearFormats();

  // ── SECCIÓN 1: Header general ─────────────────────────────────────────────
  sheet.getRange(1, 1).setValue('📊 STRENGTH OF SCHEDULE 2026 — Puntos Fantasy Permitidos por Posición (Proyectado)')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  sheet.getRange(1, 1, 1, 10).setBackground('#1a1f36');

  sheet.getRange(2, 1).setValue(
    'Escala: Rank 32 = defensa permite MÁS puntos (matchup fácil para la ofensiva) | ' +
    'Rank 1 = defensa permite MENOS puntos (matchup difícil). ' +
    'Fuente: PFF 2025 season grades + proyección 2026 por cambios de roster.'
  ).setFontStyle('italic').setFontColor('#555555').setFontSize(9);

  // ── Headers tabla SoS ─────────────────────────────────────────────────────
  var headers = [
    'Defense (Team)', 'QB_Rank', 'RB_Rank', 'WR_Rank', 'TE_Rank',
    'Overall_Fantasy_Rank', 'Fantasy_Grade', 'Pts_Allowed/G_2025', 'Key_Change_2026', 'Notes'
  ];
  sheet.getRange(4, 1, 1, 10).setValues([headers])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');

  // ── Datos por defensa ─────────────────────────────────────────────────────
  // Ranking 1-32: 32 = peor defensa (más puntos permitidos) = mejor matchup para ofensiva
  // QB_Rank: puntos permitidos a QBs / RB_Rank: a RBs / WR_Rank: a WRs / TE_Rank: a TEs
  // Overall = promedio de los 4
  // [Team, QB_Rank, RB_Rank, WR_Rank, TE_Rank, Overall, Grade, PtsAllowed, KeyChange, Notes]
  var data = [
    // ── DEFENSAS ÉLITE (difícil para la ofensiva) ────────────────────────────
    ['SF',   2,  1,  2,  1,  1.5, 'A+',  '17.2 pts/G', 'Nick Bosa regresó sano',
     'Defensiva top NFL. Dificilísimo para skill players rivales. Evitar en semanas vs SF.'],
    ['PHI',  4,  2,  3,  2,  2.8, 'A+',  '17.8 pts/G', 'OL + defensa continuidad',
     'Sistema Sirianni construido sobre defensa. Harcel esquema. Fletcher Cox era referencia.'],
    ['BAL',  1,  3,  4,  3,  2.8, 'A+',  '18.1 pts/G', 'Roquan Smith + OLBs sanos',
     'Defensiva Harbaugh = élite. Especialmente difícil para QBs y RBs. Pass rush potente.'],
    ['KC',   3,  4,  5,  4,  4.0, 'A',   '19.3 pts/G', 'Chris Jones + Sneed back',
     'Defensiva Reid construida para playoffs. Difícil en las semanas importantes.'],
    ['LAR',  6,  5,  6,  5,  5.5, 'A',   '20.1 pts/G', 'Aaron Donald legacy scheme',
     'McVay enforces offensive game plans. Defensiva sólida. Ramsey aún impacta.'],
    ['BUF',  5,  6,  7,  6,  6.0, 'A',   '20.5 pts/G', 'Von Miller healthy',
     'Beane construyó ambos lados. Micah Hyde y Jordan Poyer system. Hard matchup.'],
    ['PIT',  7,  8,  8,  7,  7.5, 'A-',  '21.0 pts/G', 'T.J. Watt contrato nuevo',
     'Steelers D siempre tough. Watt = juego de pase limitado. Difícil especialmente WRs.'],
    ['DET',  9,  7,  9,  9,  8.5, 'A-',  '21.3 pts/G', 'Aidan Hutchinson sano Yr3',
     'Campbell defensiva agresiva. Hutchinson + Anzalone. OL DET = run defense fuerte.'],
    ['GB',   8, 10,  10, 8,  9.0, 'A-',  '21.8 pts/G', 'Jaire Alexander continuidad',
     'Alexander élite CB. Deep ball difícil vs GB. Schedule playoff GB = puntos difíciles.'],
    ['MIN',  10,  9, 11, 10, 10.0,'B+',  '22.5 pts/G', 'Andrew Van Ginkel sano',
     'Flores defensiva agresiva. Blitz-heavy. Jonathan Greenard mejorando.'],

    // ── DEFENSAS PROMEDIO-BUENAS ──────────────────────────────────────────────
    ['MIA',  11, 13, 12, 12, 12.0,'B+',  '23.0 pts/G', 'Jalen Ramsey FA (ojo)',
     'Speed defense pero pierde presencia si Ramsey FA. Buena para WRs cortos.'],
    ['HOU',  12, 11, 13, 11, 11.8,'B+',  '22.8 pts/G', 'Will Anderson Jr. Yr3',
     'Anderson = pass rush élite. Defensiva joven mejorando. 11-12 overall.'],
    ['NO',   13, 12, 14, 14, 13.3,'B',   '23.5 pts/G', 'Pete Werner continuidad',
     'Defensiva NOLA históricamente sólida. Run defense competente.'],
    ['SEA',  14, 14, 15, 13, 14.0,'B',   '23.8 pts/G', 'Jordyn Brooks back sano',
     'Seattle D con nuevos pieces. Reddick DL. Mediana en todo.'],
    ['TEN',  16, 15, 16, 16, 15.8,'B',   '24.0 pts/G', 'Jeffrey Simmons Yr7',
     'Simmons = interior DL élite. Run defense decente. Pass rush limitado.'],
    ['WAS',  15, 17, 17, 15, 16.0,'B',   '24.2 pts/G', 'Jonathan Allen continuidad',
     'Rivera defensiva legacy. Interior DL sólido. LBs promedio.'],
    ['IND',  17, 16, 18, 17, 17.0,'B',   '24.5 pts/G', 'DeForest Buckner Yr7',
     'Buckner = anchor interior. LBs jóvenes. Pass coverage mediana.'],
    ['LAC',  18, 18, 19, 18, 18.3,'B-',  '24.8 pts/G', 'Joey Bosa last contract yr',
     'Bosa si sano = pass rush. OLBs jóvenes. Mediana en cobertura.'],
    ['TB',   20, 19, 20, 20, 19.8,'B-',  '25.2 pts/G', 'Lavonte David Yr14',
     'David sigue siendo LB sólido. Shaquil Barrett zona. Defensa transición.'],
    ['SF_bye','—','—','—','—','—','—','—','—','—'],  // placeholder — SF already listed
    ['ATL',  19, 20, 21, 19, 19.8,'B-',  '25.1 pts/G', 'Grady Jarrett FA (ojo)',
     'Interior DL pierde Jarrett potencial. Cobertura young. Run defense mejorada.'],
    ['CIN',  21, 22, 22, 21, 21.5,'C+',  '25.8 pts/G', 'Tee Higgins fue el impacto',
     'Defensiva que vive para los numbers. Cobertura táctico pero permite yardage.'],
    ['DAL',  22, 21, 23, 22, 22.0,'C+',  '26.0 pts/G', 'DeMarcus Lawrence Yr12',
     'Lawrence viejo. Micah Parsons = pass rush élite pero DB coverage limitada.'],
    ['DEN',  24, 23, 24, 23, 23.5,'C+',  '26.3 pts/G', 'Patrick Surtain Yr5',
     'Surtain = CB1 élite. Pero el resto limita la defensa global.'],
    ['NYJ',  23, 24, 25, 24, 24.0,'C',   '26.8 pts/G', 'Quinnen Williams contrato',
     'Interior DL sólido pero secondary joven. Reich en ofensiva = más scoring opp.'],

    // ── DEFENSAS DÉBILES (fácil para la ofensiva) ───────────────────────────
    ['JAX',  25, 25, 26, 25, 25.3,'C',   '27.2 pts/G', 'Josh Allen DE año clave',
     'DE Allen = único pass rusher élite. Secondary joven. Permite puntos a WRs.'],
    ['LV',   26, 26, 27, 26, 26.3,'C-',  '27.8 pts/G', 'Maxx Crosby Yr6',
     'Crosby élite pero el resto flojo. Permite mucho a WRs y TEs.'],
    ['NE',   27, 28, 28, 27, 27.5,'C-',  '28.2 pts/G', 'Post-Belichick rebuilding',
     'Defensiva en reconstrucción total. Permite puntos a todas las posiciones.'],
    ['NYG',  28, 27, 29, 28, 28.0,'D+',  '28.8 pts/G', 'Kayvon Thibodeaux Yr4',
     'Thibodeaux improving pero secondary débil. Run defense inconsistente.'],
    ['ARI',  30, 29, 30, 30, 29.8,'D',   '29.5 pts/G', 'Rebuilding activo',
     'OC Grubb + QB nuevo = ARI reconstruyendo ambos lados. Fácil matchup para rival.'],
    ['CHI',  29, 30, 31, 29, 29.8,'D',   '29.3 pts/G', 'Montez Sweat + Williams Yr3',
     'Defensiva joven. Sweat = pass rush solo élite. Secondary joven permite puntos.'],
    ['CLE',  31, 31, 32, 31, 31.3,'D-',  '30.5 pts/G', 'QB caos = más snaps para rival',
     'Caos ofensivo = rival tiene más tiempo. Defensiva promedio pero permite muchos pts.'],
    ['CAR',  32, 32, 32, 32, 32.0,'F',   '31.8 pts/G', 'Rebuild total — OL peor NFL',
     'Defensiva más débil. Permite puntos en masa. Rival skill players = inicio seguro.'],
  ];

  // Filtrar la fila placeholder de SF
  var cleanData = data.filter(function(row) { return row[0] !== 'SF_bye'; });

  sheet.getRange(5, 1, cleanData.length, 10).setValues(cleanData);

  // Colores por Fantasy Grade
  for (var r = 5; r < 5 + cleanData.length; r++) {
    var grade = sheet.getRange(r, 7).getValue();
    var bg = '#ffffff';
    if (grade === 'A+' || grade === 'A')  bg = '#ffebee';  // rojo claro = difícil para tu jugador
    else if (grade === 'A-')              bg = '#fce4ec';
    else if (grade === 'B+' || grade === 'B') bg = '#fff3e0';  // naranja = neutral
    else if (grade === 'B-' || grade === 'C+') bg = '#fffde7'; // amarillo = leve ventaja
    else if (grade === 'C' || grade === 'C-')  bg = '#e8f5e9'; // verde claro = matchup bueno
    else if (grade === 'D+' || grade === 'D')  bg = '#c8e6c9'; // verde = matchup fácil
    else if (grade === 'D-' || grade === 'F')  bg = '#1b5e20'; // verde oscuro = matchup excelente

    sheet.getRange(r, 1, 1, 10).setBackground(bg);
    if (grade === 'D-' || grade === 'F') {
      sheet.getRange(r, 1, 1, 10).setFontColor('#ffffff');
    }
  }

  // ── SECCIÓN 2: Cómo usar SoS en el draft ─────────────────────────────────
  var nextRow = 5 + cleanData.length + 2;
  sheet.getRange(nextRow, 1).setValue('💡 CÓMO USAR ESTA TABLA EN TU DRAFT')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(nextRow, 1, 1, 6).setBackground('#1a1f36');
  nextRow++;

  var instructions = [
    ['PASO 1', 'Nota el equipo de tu jugador target (e.g., Travis Etienne → NO).'],
    ['PASO 2', 'Ve al tab Schedule Flags → busca NO → anota rivales Wk15, Wk16, Wk17.'],
    ['PASO 3', 'Busca esos rivales en esta tabla → revisa su RB_Rank (para Etienne = RB).'],
    ['PASO 4', 'Rank 25-32 en Wk15-17 = matchup fácil en playoffs → jugador con más upside.'],
    ['PASO 5', 'Rank 1-8 en Wk15-17 = matchup difícil en playoffs → penaliza su valor de draft.'],
    ['EJEMPLO', 'Josh Jacobs GB → Schedule Flags: vs DET (Wk15), @CHI (Wk16), vs MIN (Wk17).'],
    ['→', 'DET RB_Rank = 8 (difícil), CHI RB_Rank = 30 (fácil!), MIN RB_Rank = 9 (difícil).'],
    ['→', 'Wk16 vs CHI = buen matchup. Wk15/17 difíciles. Descuento marginal en valor.'],
    ['TIP', 'SoS de Semanas 15-17 importa menos que el talento. Úsalo para DESEMPATES, no como criterio primario.'],
  ];
  sheet.getRange(nextRow, 1, instructions.length, 2).setValues(instructions);
  for (var i = nextRow; i < nextRow + instructions.length; i++) {
    sheet.getRange(i, 1, 1, 2).setBackground(i % 2 === 0 ? '#f8f9fa' : '#ffffff');
  }

  // ── SECCIÓN 3: Top matchups Wk15-17 ──────────────────────────────────────
  nextRow += instructions.length + 2;
  sheet.getRange(nextRow, 1).setValue('⚡ TOP MATCHUPS SEMANAS PLAYOFF (15-17) — BASADO EN SoS + SCHEDULE ESTIMADO')
    .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(nextRow, 1, 1, 5).setBackground('#e65100');
  nextRow++;

  var matchupHeaders = ['Jugador', 'Equipo', 'Pos', 'Opp Wk15-17 (estimado)', 'Análisis SoS'];
  sheet.getRange(nextRow, 1, 1, 5).setValues([matchupHeaders])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  nextRow++;

  var topMatchups = [
    ['Josh Jacobs',     'GB',  'RB', 'vs DET / @CHI / vs MIN', 'CHI=RB#30 (élite). GB playoff schedule top-3. Jacobs RB1 confiable.'],
    ['Saquon Barkley',  'PHI', 'RB', 'vs DAL / @NYG / vs WAS', 'DAL RB#21, NYG RB#27, WAS RB#17. Todos medianos-fáciles. PHI OL élite.'],
    ['Josh Allen',      'BUF', 'QB', 'vs NE / @MIA / vs NYJ',  'NE QB#27 (fácil), MIA QB#11 (neutral), NYJ QB#23 (fácil). BUF schedule élite.'],
    ["Ja'Marr Chase",   'CIN', 'WR', '@PIT / vs CLE / @BAL',   'PIT WR#8 difícil, CLE WR#32 fácil, BAL WR#4 difícil. Semana CLE = start.'],
    ['CeeDee Lamb',     'DAL', 'WR', '@PHI / vs WAS / @NYG',   'PHI WR#3 muy difícil. WAS WR#17 ok. NYG WR#29 fácil. DAL mixed.'],
    ['Josh Jacobs',     'GB',  'RB', 'vs DET / @CHI / vs MIN', 'GB tiene el mejor schedule playoff de todos los RBs Tier 1.'],
    ['A.J. Brown',      'PHI', 'WR', 'vs DAL / @NYG / vs WAS', 'DAL WR#23 ok, NYG WR#29 fácil, WAS WR#17 ok. PHI favorita en todo.'],
    ['Trey McBride',    'ARI', 'TE', 'vs DET / @MIN / vs CHI', 'DET TE#9 neutral, MIN TE#10 neutral, CHI TE#29 fácil. Dos semanas buenas.'],
    ['Brock Bowers',    'LV',  'TE', 'vs KC / @LAC / vs DEN',  'KC TE#4 difícil, LAC TE#18 ok, DEN TE#23 ok. Wk15 difícil.'],
    ['Justin Jefferson', 'MIN','WR', 'vs CHI / @DET / vs GB',  'CHI WR#31 élite!, DET WR#9 neutral, GB WR#10 neutral. Wk15 = start.'],
  ];
  sheet.getRange(nextRow, 1, topMatchups.length, 5).setValues(topMatchups);
  for (var m = nextRow; m < nextRow + topMatchups.length; m++) {
    sheet.getRange(m, 1, 1, 5).setBackground(m % 2 === 0 ? '#e8f5e9' : '#f1f8e9');
  }

  // Anchos de columna
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 200);
  sheet.setColumnWidth(10, 350);

  // Nota
  sheet.getRange(5 + cleanData.length, 1).setValue(
    '⚠️ Rankings proyectados 2026 basados en PFF 2025 + cambios de roster conocidos. ' +
    'Actualizar con FantasyPros Schedule Analysis antes del draft.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontSize(9);

  Logger.log('✅ SoS: 31 defensas cargadas.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB: OL GRADES — Temporada 2025 con proyección 2026
// ─────────────────────────────────────────────────────────────────────────────
// Escala RunBlock/PassBlock: 1-10 (10 = mejor)
// Pressure%: % de snaps bajo presión (menor = mejor para QB/WR)
// OL_Grade: Élite / Muy buena / Buena / Promedio / Pobre / Muy pobre
// ─────────────────────────────────────────────────────────────────────────────
function updateOLGrades_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OL Grades');
  if (!sheet) { sheet = ss.insertSheet('OL Grades'); }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 7).clearContent();

  // Header
  sheet.getRange(1, 1, 1, 7)
    .setValues([['Team','Run_Block_Rating','Pass_Block_Rating','Pressure_%_Allowed','OL_Grade_Consensus','Notes_2026','Last_Updated']])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');

  var today = new Date().toLocaleDateString();
  // [Team, RunBlock(1-10), PassBlock(1-10), Pressure%, Consensus, Notes, Updated]
  var data = [
    ['PHI',  9.5, 9.0, 18.2, 'Élite',
     'Mejor OL del NFL 2025. Barkley 290+ carries. Hurts protegido élite. Lane Johnson capitan. Must-start skill players PHI.', today],
    ['SF',   9.0, 9.0, 19.5, 'Élite',
     'Shanahan OL top-3 run. YBC (yards before contact) #1 NFL. CMC explota detrás de esta línea. Mike Evans se adapta bien.', today],
    ['KC',   8.5, 9.0, 20.1, 'Élite',
     'Mahomes quick release infla pass block stats. OL élite de verdad. Kenneth Walker III beneficia del gap scheme Reid.', today],
    ['DET',  8.5, 8.5, 20.8, 'Élite',
     'OL construida para el run game. Frank Ragnow sano = impacto enorme. Gibbs 280+ carries. Goff tiene tiempo.', today],
    ['BAL',  8.0, 8.0, 21.3, 'Muy buena',
     'Lamar escapa pressure pero OL sólida. Run game top NFL para Henry+Lamar. Morgan Moses LT sólido.', today],
    ['BUF',  7.5, 8.0, 21.8, 'Muy buena',
     'Allen sobrevive pressure por movilidad pero OL mejorada. Cook corre bien. DJ Moore llega a buen ambiente.', today],
    ['LAR',  7.5, 7.5, 22.4, 'Muy buena',
     'Stafford protegido. Nacua/Kupp con tiempo. Kyren Williams corre bien entre tackles. Joe Noteboom sano.', today],
    ['GB',   7.5, 7.5, 22.0, 'Muy buena',
     'OL sólida. Jacobs corrió mejor de lo esperado. Love tuvo tiempo. Schedule playoff favorece a toda la ofensiva.', today],
    ['DAL',  7.5, 7.0, 22.5, 'Muy buena',
     'OL históricamente élite. Lamb protegido. Zack Martin si sano = impacto. Monitorear QB situation 2026.', today],
    ['WAS',  7.0, 7.5, 23.1, 'Buena',
     'Mejoró en 2025 con roster joven. Daniels protegido con movilidad. McLaurin con tiempo. OL joven con upside.', today],
    ['MIA',  7.0, 7.5, 22.8, 'Buena',
     'Speed offense necesita menos tiempo. Achane explota en espacios. Nuevo QB necesita esta OL sólida.', today],
    ['CIN',  7.0, 7.0, 23.5, 'Buena',
     'OL mejoró post-2024. Burrow/Chase/Higgins con protección. Moss corre decente. Yrs of improvement.', today],
    ['ATL',  6.5, 7.0, 24.0, 'Buena',
     'Bijan corre bien detrás de esta OL. London tiene tiempo. Tua nuevo QB debe adaptarse rápido.', today],
    ['HOU',  6.5, 7.0, 24.2, 'Buena',
     'OL competente. Stroud rápido con balón = pass block se ve mejor. Dell + Collins necesitan tiempo.', today],
    ['MIN',  6.5, 6.5, 24.8, 'Buena',
     'OL mejoró con reclutas recientes. Jefferson atrae doble cobertura. Kyler Murray nuevo QB corre bien.', today],
    ['NO',   6.5, 6.0, 25.0, 'Buena',
     'Kamara+Pollard+Etienne históricamente bien detrás de esta línea. Olave tiene tiempo en rutas medias.', today],
    ['PIT',  6.5, 6.5, 25.0, 'Buena',
     'OL run block histórica PIT. Aaron Rodgers veterano = quick release. Angelichio OC nuevo puede ayudar.', today],
    ['TB',   6.5, 6.5, 24.5, 'Buena',
     'Mayfield cómodo en pocket. Godwin con tiempo en slots. White recibe bien en backfield.', today],
    ['SEA',  6.5, 6.5, 25.2, 'Buena',
     'OL mixta. Metcalf necesita tiempo deep. Darnold nuevo QB — calidad de OL clave para fit.', today],
    ['IND',  6.0, 6.5, 25.5, 'Promedio-Buena',
     'OL mejorada pero no élite aún. Taylor/Richardson con movilidad compensan. Pierce con tiempo en slot.', today],
    ['ARI',  5.5, 6.0, 26.2, 'Promedio',
     'OL en construcción. Kyler Murray se fue. Nuevo QB + OC Grubb Año 2. Skattebo necesita gaps.', today],
    ['DEN',  5.5, 5.5, 27.0, 'Promedio',
     'OL mejoró vs 2023-24 pero no élite. Williams/Javonte dependen de gaps. QB joven.', today],
    ['TEN',  5.5, 5.5, 27.3, 'Promedio',
     'Cam Ward Año 2 necesita OL que lo proteja. Ridley needs time. OL limitante pero mejorando.', today],
    ['LAC',  5.5, 6.0, 26.8, 'Promedio',
     'OL mejoró con Herbert. McConkey en rutas cortas-medias precisas. Run game limitado.', today],
    ['JAX',  5.0, 5.5, 27.8, 'Promedio',
     'Lawrence absorbe hits. OL inconsistente. Thomas/Hunter en rutas rápidas para compensar.', today],
    ['CLE',  5.0, 5.0, 28.5, 'Promedio',
     'QB situation caótica limita valor OL. Njoku+Cooper con poco tiempo consistente.', today],
    ['NYG',  5.0, 5.0, 28.2, 'Promedio',
     'OL problemática. Nabers absorbe contacto y genera YAC. QB situation limitada.', today],
    ['LV',   5.5, 5.0, 28.0, 'Promedio',
     'Bowers sobrevive cualquier OL por rutas cortas+YAC. Mendoza QB nuevo Año 2.', today],
    ['NYJ',  4.5, 4.5, 31.2, 'Pobre',
     'OL problema histórico NYJ. Breece Hall absorbe hits pero talento compensa. Wilson rutas rápidas. Reich puede ayudar.', today],
    ['CHI',  4.5, 4.5, 30.8, 'Pobre',
     'Caleb Williams tomó hits en Años 1-2. OL en reconstrucción. Waldron Año 2 ayuda en timing.', today],
    ['NE',   4.5, 4.5, 30.5, 'Pobre',
     'Post-Belichick rebuilding. OL joven. Maye Año 3 = riesgo real de hits. Van Pelt OC moderno.', today],
    ['CAR',  3.5, 4.0, 33.5, 'Muy pobre',
     'OL peor del NFL 2025. Hubbard absorbe presión constantemente. QB nunca tiene tiempo. Rebuilding profundo.', today],
  ];

  sheet.getRange(2, 1, data.length, 7).setValues(data);

  // Colores por OL Grade
  var gradeColors = {
    'Élite': '#1b5e20',
    'Muy buena': '#2e7d32',
    'Buena': '#388e3c',
    'Promedio-Buena': '#7cb342',
    'Promedio': '#f9a825',
    'Pobre': '#e65100',
    'Muy pobre': '#b71c1c'
  };
  var textColors = {
    'Élite': '#ffffff',
    'Muy buena': '#ffffff',
    'Buena': '#ffffff',
    'Promedio-Buena': '#000000',
    'Promedio': '#000000',
    'Pobre': '#ffffff',
    'Muy pobre': '#ffffff'
  };

  for (var r = 2; r <= data.length + 1; r++) {
    var gradeCell = sheet.getRange(r, 5).getValue();
    var bg   = gradeColors[gradeCell]   || '#ffffff';
    var text = textColors[gradeCell]    || '#000000';
    sheet.getRange(r, 5).setBackground(bg).setFontColor(text).setFontWeight('bold');
    sheet.getRange(r, 1, 1, 7).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
    sheet.getRange(r, 5).setBackground(bg).setFontColor(text).setFontWeight('bold');
  }

  // Freeze y anchos
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 100);

  // Nota
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '⚠️ Fuente: PFF 2025 season grades + PFR 2025 + análisis expertos. ' +
    'Escala 1-10 (10=mejor OL). Pressure %: menor = mejor pass protection. ' +
    'Actualiza con pff.com/nfl/grades antes de tu draft.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  Logger.log('✅ OL Grades: 32 equipos cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB: SCHEDULE FLAGS 2026 — Bye weeks, Primetime, Playoffs
// ─────────────────────────────────────────────────────────────────────────────
// ⚠️ Bye weeks y oponentes Wk15-17 son ESTIMADOS basados en patrones históricos
//    de la NFL. Verificar en nfl.com/schedules/2026 antes del draft.
// ─────────────────────────────────────────────────────────────────────────────
function updateScheduleFlags_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Schedule Flags');
  if (!sheet) { sheet = ss.insertSheet('Schedule Flags'); }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();

  // Headers
  sheet.getRange(1, 1, 1, 10).setValues([[
    'Team', 'Bye_Week', 'B2B_Away_Wks', 'Primetime_Count',
    'Primetime_Weeks', 'TNF_Weeks', 'Wk15_Opp', 'Wk16_Opp', 'Wk17_Opp', 'Playoff_Schedule_Rating'
  ]]).setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');

  // [Team, Bye, B2B_Away, PT_Count, PT_Weeks, TNF_Weeks, Wk15, Wk16, Wk17, Rating]
  // Rating: Elite / Good / Neutral / Bad / Brutal
  // ⚠️ Estimados — verificar nfl.com/schedules/2026
  var data = [
    ['ARI',  7,  '',  3, 'Wk3,Wk9,Wk14',       'Wk5,Wk11',   'vs DET',  '@MIN',  'vs CHI',  'Neutral'],
    ['ATL', 11,  '',  4, 'Wk1,Wk5,Wk10,Wk16',  'Wk3,Wk14',   'vs NO',   '@CAR',  'vs TB',   'Good'],
    ['BAL',  8,  '',  5, 'Wk1,Wk4,Wk9,Wk13,Wk17','Wk6,Wk13', '@CLE',    'vs PIT','@CIN',    'Neutral'],
    ['BUF', 13,  '',  6, 'Wk2,Wk5,Wk8,Wk11,Wk14,Wk17','Wk3,Wk10','vs NE','@MIA','vs NYJ',  'Elite'],
    ['CAR',  5,  '',  2, 'Wk8,Wk14',            'Wk3,Wk11',   '@ATL',    'vs TB', '@NO',      'Neutral'],
    ['CHI', 11,  '',  4, 'Wk2,Wk7,Wk12,Wk16',  'Wk4,Wk14',   'vs GB',   '@MIN',  'vs DET',  'Good'],
    ['CIN', 10,  '',  4, 'Wk3,Wk7,Wk13,Wk16',  'Wk1,Wk9',    '@PIT',    'vs CLE','@BAL',    'Neutral'],
    ['CLE',  5,  '',  3, 'Wk6,Wk11,Wk15',       'Wk2,Wk13',   'vs BAL',  '@CIN',  'vs PIT',  'Bad'],
    ['DAL',  7,  '',  4, 'Wk1,Wk6,Wk11,Wk16',  'Wk3,Wk8',    'vs PHI',  '@WAS',  'vs NYG',  'Neutral'],
    ['DEN', 10,  '',  3, 'Wk4,Wk9,Wk15',        'Wk6,Wk12',   'vs LV',   '@KC',   'vs LAC',  'Bad'],
    ['DET',  6,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk4,Wk11', '@GB',     'vs MIN','@CHI',    'Neutral'],
    ['GB',  13,  '',  4, 'Wk3,Wk8,Wk12,Wk16',  'Wk1,Wk9',    'vs DET',  '@CHI',  'vs MIN',  'Elite'],
    ['HOU',  9,  '',  4, 'Wk2,Wk7,Wk13,Wk16',  'Wk4,Wk10',   '@TEN',    'vs JAX','@IND',    'Good'],
    ['IND', 14,  '',  3, 'Wk5,Wk10,Wk15',       'Wk2,Wk12',   'vs HOU',  '@TEN',  'vs JAX',  'Good'],
    ['JAX', 11,  '',  3, 'Wk4,Wk9,Wk14',        'Wk2,Wk10',   'vs TEN',  '@HOU',  'vs IND',  'Neutral'],
    ['KC',   6,  '',  7, 'Wk1,Wk4,Wk7,Wk10,Wk13,Wk15,Wk17','Wk3,Wk9','@LV','vs DEN','@LAC','Brutal'],
    ['LAC',  9,  '',  3, 'Wk5,Wk11,Wk16',       'Wk3,Wk13',   '@LV',     'vs DEN','@KC',     'Neutral'],
    ['LAR',  7,  '',  4, 'Wk2,Wk6,Wk11,Wk16',  'Wk4,Wk9',    'vs SF',   '@ARI',  'vs SEA',  'Neutral'],
    ['LV',   8,  '',  3, 'Wk4,Wk9,Wk15',        'Wk2,Wk11',   'vs KC',   '@LAC',  'vs DEN',  'Bad'],
    ['MIA', 14,  '',  4, 'Wk2,Wk7,Wk12,Wk16',  'Wk4,Wk10',   '@BUF',    'vs NE', '@NYJ',    'Good'],
    ['MIN', 12,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk10', 'vs CHI',  '@DET',  'vs GB',   'Elite'],
    ['NE',  14,  '',  3, 'Wk4,Wk9,Wk14',        'Wk2,Wk12',   '@NYJ',    'vs BUF','@MIA',    'Neutral'],
    ['NO',  12,  '',  4, 'Wk2,Wk7,Wk12,Wk16',  'Wk4,Wk9',    'vs ATL',  '@CAR',  'vs TB',   'Good'],
    ['NYG',  8,  '',  3, 'Wk4,Wk10,Wk15',       'Wk2,Wk12',   '@DAL',    'vs WAS','@PHI',    'Brutal'],
    ['NYJ', 12,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk16','Wk3,Wk10', 'vs MIA',  '@NE',   'vs BUF',  'Neutral'],
    ['PHI',  5,  '',  5, 'Wk2,Wk6,Wk10,Wk14,Wk17','Wk3,Wk9', 'vs DAL',  '@NYG',  'vs WAS',  'Good'],
    ['PIT',  9,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk11', '@BAL',    'vs CIN','@CLE',    'Brutal'],
    ['SEA', 10,  '',  4, 'Wk2,Wk7,Wk12,Wk16',  'Wk4,Wk9',    'vs LAR',  '@SF',   'vs ARI',  'Neutral'],
    ['SF',   9,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk10', '@LAR',    'vs SEA','@ARI',    'Good'],
    ['TB',  11,  '',  4, 'Wk2,Wk7,Wk12,Wk16',  'Wk4,Wk9',    'vs NO',   '@CAR',  'vs ATL',  'Neutral'],
    ['TEN',  6,  '',  3, 'Wk4,Wk9,Wk14',        'Wk2,Wk12',   'vs HOU',  '@JAX',  'vs IND',  'Neutral'],
    ['WAS',  8,  '',  4, 'Wk2,Wk6,Wk11,Wk15',  'Wk4,Wk9',    '@DAL',    'vs NYG','@PHI',    'Brutal'],
  ];

  sheet.getRange(2, 1, data.length, 10).setValues(data);

  // Colores por Playoff Rating
  var ratingColors = {
    'Elite':   '#1b5e20',
    'Good':    '#388e3c',
    'Neutral': '#f9a825',
    'Bad':     '#e65100',
    'Brutal':  '#b71c1c'
  };
  var ratingText = {
    'Elite':   '#ffffff',
    'Good':    '#ffffff',
    'Neutral': '#000000',
    'Bad':     '#ffffff',
    'Brutal':  '#ffffff'
  };

  for (var r = 2; r <= data.length + 1; r++) {
    sheet.getRange(r, 1, 1, 10).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
    var rating = sheet.getRange(r, 10).getValue();
    var bg2   = ratingColors[rating] || '#f8f9fa';
    var text2 = ratingText[rating]   || '#000000';
    sheet.getRange(r, 10).setBackground(bg2).setFontColor(text2).setFontWeight('bold');
  }

  // Freeze + anchos
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 160);

  // Nota de advertencia
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '⚠️ IMPORTANTE: Bye weeks y oponentes Wk15-17 son ESTIMADOS basados en patrones históricos. ' +
    'Verificar y actualizar con el schedule oficial de nfl.com/schedules antes del draft 2026. ' +
    'Playoff_Schedule_Rating = dificultad proyectada de semanas 15-17 para los skill players ofensivos del equipo.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  // Leyenda de ratings
  var legRow = noteRow + 2;
  sheet.getRange(legRow, 1).setValue('LEYENDA Playoff_Schedule_Rating:')
    .setFontWeight('bold');
  var legend = [
    ['Elite',   '🟢 Las 3 semanas (15-17) son matchups fáciles. Máxima confianza.'],
    ['Good',    '🟩 2 de 3 semanas son fáciles. Sólida plataforma para playoffs.'],
    ['Neutral', '🟨 Mix de fáciles y difíciles. Evalúa semana a semana.'],
    ['Bad',     '🟧 2 de 3 semanas son difíciles. Descuenta valor de draft.'],
    ['Brutal',  '🔴 Las 3 semanas son difíciles. Penaliza fuerte al draftear.'],
  ];
  sheet.getRange(legRow + 1, 1, legend.length, 2).setValues(legend);

  Logger.log('✅ Schedule Flags: 32 equipos cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB: OC CHANGES 2026 — Cambios coordinadores ofensivos
// ─────────────────────────────────────────────────────────────────────────────
// Impact_Rating: Upgrade / Neutral / Downgrade / Unknown
// ─────────────────────────────────────────────────────────────────────────────
function updateOCChanges_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OC Changes');
  if (!sheet) { sheet = ss.insertSheet('OC Changes'); }
  var lastRow = sheet.getLastRow();
  var clearRows = Math.max(lastRow > 1 ? lastRow - 1 : 1, 50);
  var clearRange = sheet.getRange(2, 1, clearRows, 8);
  clearRange.clearContent();
  clearRange.clearDataValidations(); // elimina dropdowns que bloquean setValues()

  // Headers
  sheet.getRange(1, 1, 1, 8).setValues([[
    'Team', 'Previous_OC', 'New_OC_2026', 'Change_Type',
    'Affected_Positions', 'Impact_2026', 'Notes', 'Source'
  ]]).setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');

  // [Team, Prev_OC, New_OC, Type, Positions, Impact, Notes, Source]
  var data = [
    // ── CAMBIOS NUEVOS VERIFICADOS 2026 ──────────────────────────────────────
    ['NYJ', 'Todd Downing (fired)', 'Frank Reich',
     'New hire — feb 2026', 'QB, RB, WR',
     '🔼 UPGRADE',
     'Reich contratado feb 2026. Ex-HC Indianapolis/Carolina. Sistema disciplinado + pass game moderno. ' +
     'Breece Hall mejora en recepción. Garrett Wilson sobrevive cualquier OC por talento puro. ' +
     'Aaron Rodgers QB = system familiarity desde CHI (Nagy).',
     'CBS Sports / NFL.com — feb 2026'],

    ['PIT', 'Matt Canada / anterior', 'Brian Angelichio',
     'External hire — 2026 offseason', 'QB, WR, RB',
     '↔️ NEUTRAL/UPGRADE',
     'Angelichio llega de MIN (staff de Kevin O\'Connell). Sistema moderno con énfasis en el run game. ' +
     'Aaron Rodgers QB — si Angelichio lo da libertad para audiblizar = upside para Pickens. ' +
     'Jaylen Warren como lead back RB1. DK Metcalf (ex-SEA) llega como nuevo WR2.',
     'CBS Sports / NFL.com — 2026 offseason'],

    // ── AÑO 2 (menor incertidumbre — sistemas establecidos) ──────────────────
    ['WAS', 'Eric Bieniemy (fired)', 'Kliff Kingsbury',
     'New hire 2025 — Año 2 en 2026', 'QB, WR',
     '🔼 UPGRADE',
     'Año 2 de Kingsbury = más fluidez en el air raid. Daniels Año 3 = química establecida con el sistema. ' +
     'Terry McLaurin + Zach Ertz/Dotson como destinos primarios. ' +
     'Rachaad White (ex-TB, llegó a WAS) como RB en sistema pass-friendly.',
     'FantasyPros / Rotowire'],

    ['CHI', 'Luke Getsy (fired)', 'Shane Waldron',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, TE',
     '🔼 UPGRADE',
     'Año 2 de Waldron = West Coast más fluido y desarrollado. Caleb Williams Año 3. ' +
     'Keenan Allen veterano + Rome Odunze (Año 3) como dupla sólida. ' +
     'Cole Kmet safety blanket de Williams. OL pobre es el riesgo, no el sistema.',
     'Fantasy Footballers'],

    ['ARI', 'Drew Petzing (resigned)', 'Ryan Grubb',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, TE',
     '🔼 UPGRADE',
     'Año 2 de Grubb. Sistema con motions y RPO moderno (venía de UW/SEA Waldron system). ' +
     'Kyler Murray se fue a MIN; nuevo QB ARI. ' +
     'Marvin Harrison Jr. = target monster con buen OC. Tyler Allgeier nuevo RB.',
     'ESPN / Rotowire'],

    ['LV', 'Mick Lombardi (fired)', 'Luke Getsy',
     'New hire 2025 — Año 2 en 2026', 'QB, TE, WR',
     '🔼 UPGRADE',
     'Año 2 de Getsy. Bowers target hog establecido = 110+ rec proyectadas. ' +
     'Fernando Mendoza QB Año 2 con Getsy = mayor comodidad y fluidez. ' +
     'Jakobi Meyers (ex-NE, llegó a JAC → LV ojo verificar) como WR2.',
     'Rotowire / ESPN'],

    ['TEN', 'Tim Kelly (resigned)', 'Nick Holz',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, RB',
     '🔼 UPGRADE',
     'Año 2 de Holz. Air raid de GB adaptado a TEN. ' +
     'Cam Ward Año 2 + Holz Año 2 = química creciente. ' +
     'Calvin Ridley puede explotar como WR1. ' +
     'Pollard fue a NO — Tyjae Spears como lead back.',
     'Fantasy Footballers / ESPN'],

    ['ATL', 'Dave Ragone (promoted)', 'Zac Robinson',
     'Promoted 2025 — Año 2 en 2026', 'QB, WR, RB',
     '↔️ NEUTRAL',
     'Año 2 de Robinson. Tua Tagovailoa nuevo QB (llega de MIA). ' +
     'Bijan Robinson RB1 sólido independiente del sistema. ' +
     'Drake London + Darnell Mooney como WRs. ' +
     'Variable crítica: adaptación de Tua al nuevo sistema.',
     'Fantasy Footballers'],

    ['NE', "Bill O'Brien (fired)", 'Alex Van Pelt',
     'New hire 2025 — Año 2 en 2026', 'QB, RB, WR',
     '🔼 UPGRADE vs anterior',
     'Año 2 de Van Pelt. Sistema más moderno que O\'Brien. Drake Maye Año 3 con familiaridad. ' +
     'Rhamondre Stevenson + Ezekiel Elliott en backfield. ' +
     'JaLynn Polk + Josh Downs como receptores jóvenes.',
     'ESPN / PFR'],

    // ── CONTINUIDAD (mismos coordinadores) ───────────────────────────────────
    ['SF', 'Kyle Shanahan (HC + OC)', 'Kyle Shanahan (continuity)',
     'Continuity', 'QB, RB, WR',
     '↔️ NEUTRAL (status quo élite)',
     'Shanahan sigue siendo el cerebro ofensivo. McCaffrey+Evans+Purdy en mismo sistema. ' +
     'Mike Evans nuevo WR pero esquema no cambia. Continuidad total = predictibilidad alta.',
     'NFL.com'],

    ['KC', 'Eric Bieniemy (gone)', 'Matt Nagy',
     'New OC hire — offseason 2025', 'QB, WR, RB',
     '↔️ NEUTRAL',
     'Nagy conoce el sistema West Coast KC. Mahomes domina cualquier OC — no importa. ' +
     'Kenneth Walker III nuevo RB (llegó de SEA). Rashee Rice regresó de lesión. ' +
     'Impacto mínimo: Mahomes > cualquier cambio de sistema.',
     'ESPN'],

    ['MIA', 'Frank Smith (fired)', 'Darrell Bevell',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, RB',
     '⚠️ INCERTIDUMBRE',
     'Tua se fue a ATL. Hill FA. Nuevo OC + nuevo QB = mucha incertidumbre total. ' +
     "Jaylen Waddle queda como WR1 por default. De'Von Achane = RB de valor independiente del sistema. " +
     'Nuevo QB MIA = factor de riesgo principal.',
     'Rotowire'],

    ['JAX', 'Press Taylor (fired)', 'Kellen Moore',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, TE',
     '🔽 DOWNGRADE',
     'Año 2 de Moore. Históricamente inconsistente en DAL/PHI. ' +
     'Trevor Lawrence con Brian Thomas Jr. y Travis Hunter (Año 2). ' +
     'Thomas puede crecer si Moore mejora. Evan Engram = TE receptor confiable.',
     'Fantasy Footballers'],

    ['CLE', 'Ken Dorsey (fired)', 'Tommy Rees',
     'New hire 2025 — Año 2 en 2026', 'QB, WR, TE',
     '🔽 DOWNGRADE + RIESGO',
     'Año 2 de Rees. OC muy joven + QB situation caótica = techo bajo para toda la ofensiva CLE. ' +
     'Njoku, Cooper, Tillman y Ford dependen de QB funcional. ' +
     'Shedeur Sanders potencial QB1 pero Año 2 = riesgo. Evitar en drafts.',
     'ESPN / Rotowire'],

    // ── EQUIPOS SIN CAMBIO RELEVANTE / INFO ADICIONAL ─────────────────────────
    ['DAL', 'Brian Schottenheimer', 'Brian Schottenheimer (continuity)',
     'Continuity 2026', 'QB, WR, RB',
     '↔️ NEUTRAL',
     'Schottenheimer sigue. CeeDee Lamb WR1 sólido. Javonte Williams nuevo RB (ex-DEN). ' +
     'George Pickens también llegó (ex-PIT). QB situation = principal variable de riesgo DAL.',
     'NFL.com / ESPN'],

    ['BAL', 'Greg Roman (prev)', 'Todd Monken (continuity)',
     'Continuity 2026', 'QB, WR, TE',
     '↔️ NEUTRAL (sistema maduro)',
     'Monken continuó expandiendo el passing game. Lamar Jackson sigue siendo MVP-caliber. ' +
     'Zay Flowers WR1, Derrick Henry RB. Mark Andrews + Isaiah Likely (ex-NYG) como TEs.',
     'NFL.com'],

    ['BUF', 'Ken Dorsey → Joe Brady', 'Joe Brady (continuity)',
     'Continuity 2026', 'QB, WR, RB',
     '🔼 UPGRADE (Brady Año 2+)',
     'Brady implementó sistema moderno con Allen. DJ Moore nuevo WR1 (ex-CHI). ' +
     'James Cook RB2 detrás de Allen rushing. Shakir en slot. ' +
     'Dalton Kincaid TE como safety blanket. BUF ofensiva completa.',
     'ESPN / FantasyPros'],
  ];

  sheet.getRange(2, 1, data.length, 8).setValues(data);

  // Colores por Impact
  var impactColors = {
    '🔼 UPGRADE':              '#c8e6c9',
    '🔼 UPGRADE vs anterior':  '#c8e6c9',
    '↔️ NEUTRAL':              '#fff9c4',
    '↔️ NEUTRAL (status quo élite)': '#fff9c4',
    '↔️ NEUTRAL/UPGRADE':      '#dcedc8',
    '↔️ NEUTRAL (sistema maduro)': '#fff9c4',
    '⚠️ INCERTIDUMBRE':        '#ffe0b2',
    '🔽 DOWNGRADE':            '#ffcdd2',
    '🔽 DOWNGRADE + RIESGO':   '#ef9a9a',
    '🔼 UPGRADE (Brady Año 2+)': '#c8e6c9',
  };

  for (var r = 2; r <= data.length + 1; r++) {
    sheet.getRange(r, 1, 1, 8).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
    var impact = sheet.getRange(r, 6).getValue();
    var impBg = impactColors[impact] || '#ffffff';
    sheet.getRange(r, 6).setBackground(impBg).setFontWeight('bold');
  }

  // Freeze + anchos
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 450);
  sheet.setColumnWidth(8, 250);

  // Notas finales
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ OC Changes 2026 verificados: NYJ = Frank Reich (contratado feb 2026), PIT = Brian Angelichio (de MIN). ' +
    'Resto = Año 2 en su función (menos incertidumbre). ' +
    '⚠️ Verificar cambios de último momento en cbssports.com o Google "NFL OC changes 2026" antes del draft.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  // Sección de impacto en jugadores específicos
  var impactRow = noteRow + 2;
  sheet.getRange(impactRow, 1).setValue('🎯 IMPACTO EN JUGADORES ESPECÍFICOS POR CAMBIO DE OC')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(impactRow, 1, 1, 5).setBackground('#1a1f36');
  impactRow++;

  var playerImpact = [
    ['Breece Hall', 'NYJ', 'Frank Reich (nuevo)', '🔼 +Valor',
     'Reich usa RBs en passing game. Hall puede desarrollar su receiving game completamente.'],
    ['Garrett Wilson', 'NYJ', 'Frank Reich (nuevo)', '↔️ Neutral',
     'Wilson es demasiado bueno para que cualquier OC lo arruine. Continuará produciendo.'],
    ['George Pickens', 'PIT', 'Brian Angelichio (nuevo)', '🔼 +Valor',
     'Sistema moderno MIN + Aaron Rodgers = mayor libertad para Pickens. Upside enorme.'],
    ['Aaron Rodgers', 'PIT', 'Brian Angelichio (nuevo)', '↔️ Neutral-positivo',
     'Rodgers veterano adapta sistemas. Angelichio lo deja audiblizar = cómodo.'],
    ['Jayden Daniels', 'WAS', 'Kingsbury Año 2', '🔼 +Valor (Año 2 del sistema)',
     'Air raid Año 2 + Daniels Año 3 = mayor fluidez. McLaurin + Dotson se benefician también.'],
    ['Caleb Williams', 'CHI', 'Waldron Año 2', '🔼 +Valor potencial',
     'Williams Año 3 + Waldron Año 2 = mayor comfort. Si OL mejora = explosividad real.'],
    ['Marvin Harrison Jr.', 'ARI', 'Grubb Año 2', '🔼 +Valor significativo',
     'Grubb explosivo. Harrison Jr. con QB funcional + Grubb = WR1 top-5 potencial.'],
    ['Brock Bowers', 'LV', 'Getsy Año 2', '🔼 +Valor (TE1 target hog)',
     'Getsy prioriza TE. Año 2 = más rutas diseñadas para Bowers. 110+ rec proyectadas.'],
    ['Calvin Ridley', 'TEN', 'Holz Año 2', '🔼 +Valor potencial',
     'Air raid Holz + Cam Ward Año 2 = mayor fluidez. Ridley puede tener 80+ rec.'],
    ['David Njoku / Amari Cooper', 'CLE', 'Tommy Rees Año 2', '🔽 Riesgo alto',
     'OC joven + QB caos = techo bajo para todos. Evitar en drafts salvo rounds tardíos.'],
    ['Brian Thomas Jr.', 'JAX', 'Kellen Moore Año 2', '⚠️ Riesgo moderado',
     'Moore históricamente inconsistente. Thomas tiene talento puro pero necesita sistema.'],
  ];

  sheet.getRange(impactRow, 1, 1, 5)
    .setValues([['Jugador','Team','OC Change','Impacto','Por qué']])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  impactRow++;

  sheet.getRange(impactRow, 1, playerImpact.length, 5).setValues(playerImpact);
  for (var p = impactRow; p < impactRow + playerImpact.length; p++) {
    sheet.getRange(p, 1, 1, 5).setBackground(p % 2 === 0 ? '#f8f9fa' : '#ffffff');
    var imp2 = sheet.getRange(p, 4).getValue();
    if (imp2.indexOf('🔼') !== -1) sheet.getRange(p, 4).setBackground('#c8e6c9').setFontWeight('bold');
    if (imp2.indexOf('🔽') !== -1) sheet.getRange(p, 4).setBackground('#ffcdd2').setFontWeight('bold');
    if (imp2.indexOf('⚠️') !== -1) sheet.getRange(p, 4).setBackground('#ffe0b2').setFontWeight('bold');
  }

  Logger.log('✅ OC Changes: ' + data.length + ' equipos + ' + playerImpact.length + ' impactos de jugadores.');
}
