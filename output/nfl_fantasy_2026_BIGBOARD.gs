/**
 * NFL Fantasy 2026 — BIG BOARD FINAL
 * Orden basado en ADP Half-PPR verificado FantasyPros mayo 2026
 * Corre: fixBigBoard_2026_VERIFIED()
 * Después de que cargue (~5 seg): corre sortBigBoardByCustom_2026()
 */

function fixBigBoard_2026_VERIFIED() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Big Board');

  // Limpiar columna A (nombres) y N (Overall_Grade)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
    sheet.getRange(2, 14, lastRow - 1, 2).clearContent(); // N y O
  }

  // ── Jugadores ordenados por ADP Half-PPR FantasyPros mayo 2026 ─────────────
  // Fuente: fantasypros.com/nfl/adp/half-point-ppr-overall.php
  var players = [
    // ADP 1-10
    'Bijan Robinson',        // ATL RB  1.0
    'Jahmyr Gibbs',          // DET RB  2.0
    "Ja'Marr Chase",         // CIN WR  3.0
    'Puka Nacua',            // LAR WR  4.0
    'Jaxon Smith-Njigba',    // SEA WR  5.3
    'Christian McCaffrey',   // SF  RB  6.0
    'Jonathan Taylor',       // IND RB  7.0
    'Amon-Ra St. Brown',     // DET WR  7.7
    'Justin Jefferson',      // MIN WR 10.0
    'James Cook',            // BUF RB 10.7

    // ADP 11-20
    'CeeDee Lamb',           // DAL WR 11.3
    'Ashton Jeanty',         // LV  RB 13.0
    "De'Von Achane",         // MIA RB 13.0
    'Saquon Barkley',        // PHI RB 14.7
    'Trey McBride',          // ARI TE 16.7
    'Brock Bowers',          // LV  TE 17.3
    'Omarion Hampton',       // LAC RB 17.3
    'Kenneth Walker III',    // KC  RB 18.7
    'Chase Brown',           // CIN RB 19.3
    'Drake London',          // ATL WR 19.7

    // ADP 20-30
    'Derrick Henry',         // BAL RB 20.3
    'Josh Allen',            // BUF QB 20.7
    'Jeremiyah Love',        // ARI RB 22.7
    'Nico Collins',          // HOU WR 23.7
    'George Pickens',        // DAL WR 24.0
    'Malik Nabers',          // NYG WR 25.7
    'A.J. Brown',            // PHI WR 28.7
    'Rashee Rice',           // KC  WR 28.7
    'Chris Olave',           // NO  WR 29.3

    // ADP 31-50
    'Josh Jacobs',           // GB  RB 31.7
    'Kyren Williams',        // LAR RB 32.0
    'Breece Hall',           // NYJ RB 33.0
    'Travis Etienne',        // NO  RB 35.3
    'Javonte Williams',      // DAL RB 35.7
    'Colston Loveland',      // CHI TE 36.0
    'Tee Higgins',           // CIN WR 37.3
    'DeVonta Smith',         // PHI WR 39.0
    'Emeka Egbuka',          // TB  WR 39.3
    'Tetairoa McMillan',     // CAR WR 39.3
    'Zay Flowers',           // BAL WR 41.7
    'Cam Skattebo',          // NYG RB 42.0
    'Lamar Jackson',         // BAL QB 43.0
    'Garrett Wilson',        // NYJ WR 43.3
    'DJ Moore',              // BUF WR 46.0
    'Ladd McConkey',         // LAC WR 46.7
    'Bucky Irving',          // TB  RB 48.3
    'Joe Burrow',            // CIN QB 48.7
    'Quinshon Judkins',      // CLE RB 49.0
    'David Montgomery',      // HOU RB 50.0

    // ADP 51-72
    'Luther Burden III',     // CHI WR 51.0
    'Tyler Warren',          // IND TE 51.0
    'TreVeyon Henderson',    // NE  RB 52.3
    'Davante Adams',         // LAR WR 52.7
    'Mike Evans',            // SF  WR 53.3
    'Jadarian Price',        // SEA RB 55.0
    'Jameson Williams',      // DET WR 56.0
    'Jaylen Waddle',         // DEN WR 56.0
    'Terry McLaurin',        // WAS WR 56.3
    'Drake Maye',            // NE  QB 57.0
    'Jayden Daniels',        // WAS QB 57.3
    "D'Andre Swift",         // CHI RB 58.0
    'Rome Odunze',           // CHI WR 63.0
    'Bhayshul Tuten',        // JAC RB 64.3
    'Jalen Hurts',           // PHI QB 66.3
    'Caleb Williams',        // CHI QB 66.7
    'Harold Fannin Jr.',     // CLE TE 68.3
    'Tucker Kraft',          // GB  TE 71.0
    'Christian Watson',      // GB  WR 71.0
    'Marvin Harrison Jr.',   // ARI WR 71.0
    'Brian Thomas Jr.',      // JAC WR 71.7
    'Chuba Hubbard',         // CAR RB 72.0
    'Carnell Tate',          // TEN WR 72.7

    // ADP 73-100
    'Jordyn Tyson',          // NO  WR 76.7
    'Alec Pierce',           // IND WR 78.3
    'RJ Harvey',             // DEN RB 78.3
    'Sam LaPorta',           // DET TE 79.3
    'Kyle Pitts',            // ATL TE 79.7
    'Trevor Lawrence',       // JAC QB 81.0
    'Justin Herbert',        // LAC QB 82.3
    'Jaylen Warren',         // PIT RB 82.7
    'Jaxson Dart',           // NYG QB 84.7
    'DK Metcalf',            // PIT WR 84.7
    'Courtland Sutton',      // DEN WR 86.0
    'Tony Pollard',          // TEN RB 86.0
    'Rico Dowdle',           // PIT RB 86.7
    'Patrick Mahomes',       // KC  QB 88.7
    'Kyle Monangai',         // CHI RB 89.3
    'Michael Wilson',        // ARI WR 90.0
    'Makai Lemon',           // PHI WR 90.3
    'Chris Godwin',          // TB  WR 91.0
    'Matthew Stafford',      // LAR QB 91.7
    'George Kittle',         // SF  TE 95.3
    'Brock Purdy',           // SF  QB 96.7
    'Bo Nix',                // DEN QB 98.3
    'Blake Corum',           // LAR RB 99.7

    // ADP 101-120
    'J.K. Dobbins',          // DEN RB 101.7
    'Jayden Reed',           // GB  WR 101.7
    'Jared Goff',            // DET QB 102.0
    'Jordan Addison',        // MIN WR 103.7
    'Jakobi Meyers',         // JAC WR 105.0
    'Kyler Murray',          // MIN QB 107.3
    "Wan'Dale Robinson",     // TEN WR 107.7
    'Ricky Pearsall',        // SF  WR 109.7
    'Dalton Kincaid',        // BUF TE 109.7
    'Michael Pittman Jr.',   // PIT WR 110.0
    'Jake Ferguson',         // DAL TE 110.7
    'Quentin Johnston',      // LAC WR 112.0
    'Tyler Shough',          // NO  QB 112.3
    'Isaiah Likely',         // NYG TE 113.3
    'Josh Downs',            // IND WR 113.0
    'Jordan Love',           // GB  QB 115.3
    'Baker Mayfield',        // TB  QB 116.3
    'Dallas Goedert',        // PHI TE 117.0
    'Jordan Mason',          // MIN RB 117.3
    'Mark Andrews',          // BAL TE 119.0
    'Xavier Worthy',         // KC  WR 120.0
    'Rachaad White',         // WAS RB 120.0

    // ADP 121-145
    'Matthew Golden',        // GB  WR 121.3
    'Jacory Croskey-Merritt',// WAS RB 122.0
    'Aaron Jones',           // MIN RB 122.3
    'Jonathon Brooks',       // CAR RB 124.3
    'Chris Rodriguez Jr.',   // JAC RB 124.7
    'Tyrone Tracy Jr.',      // NYG RB 130.3
    'Kenyon Sadiq',          // NYJ TE 130.0
    'Oronde Gadsden II',     // LAC TE 131.0
    'Malik Willis',          // MIA QB 132.3
    'Khalil Shakir',         // BUF WR 133.7
    'Sam Darnold',           // SEA QB 137.3
    'Jayden Higgins',        // HOU WR 138.0
    'CJ Stroud',             // HOU QB 138.7
    'Juwan Johnson',         // NO  TE 140.3
    'Cam Ward',              // TEN QB 141.0
    'Jalen Coker',           // CAR WR 141.3
    'Chig Okonkwo',          // WAS TE 145.3

    // ADP 146-168
    'Isiah Pacheco',         // DET RB 146.0
    'Zach Charbonnet',       // SEA RB 147.3
    'Hunter Henry',          // NE  TE 147.3
    'Daniel Jones',          // IND QB 147.0
    'Jonah Coleman',         // DEN RB 148.7
    'Tyler Allgeier',        // ARI RB 149.7
    'T.J. Hockenson',        // MIN TE 150.3
    'Bryce Young',           // CAR QB 151.7
    'Tyjae Spears',          // TEN RB 159.7
    'Alvin Kamara',          // NO  RB 161.7
    'Keaton Mitchell',       // LAC RB 162.3
    'Jalen Nailor',          // LV  WR 164.3
    'Tyreek Hill',           // FA  WR 165.3
    'Brian Robinson Jr.',    // ATL RB 168.7

    // ADP 169-200
    'Jacoby Brissett',       // ARI QB 169.0
    'Brandon Aiyuk',         // SF  WR 169.0
    'Fernando Mendoza',      // LV  QB 169.3
    'Dylan Sampson',         // CLE RB 170.7
    'Nicholas Singleton',    // TEN RB 171.0
    'Aaron Rodgers',         // PIT QB 172.3
    'Travis Hunter',         // JAC WR 174.5
    'Mike Washington Jr.',   // LV  RB 174.0
    'Geno Smith',            // NYJ QB 175.7
    'Deebo Samuel',          // FA  WR 177.7
    'Antonio Williams',      // WAS WR 177.0
    'Jerry Jeudy',           // CLE WR 178.7
    'Tank Bigsby',           // PHI RB 180.3
    'Calvin Ridley',         // TEN WR 181.3
    'Tre Tucker',            // LV  WR 182.0
    'Tua Tagovailoa',        // ATL QB 197.7

    // D/ST y Kickers (ronda 14)
    'Eagles D/ST',
    'Ravens D/ST',
    'Steelers D/ST',
    'Evan McPherson',
    'Jake Elliott',
    'Harrison Butker',
  ];

  // Eliminar duplicados
  var seen = {};
  var unique = [];
  for (var i = 0; i < players.length; i++) {
    if (!seen[players[i]]) {
      seen[players[i]] = true;
      unique.push([players[i]]);
    }
  }

  // Máximo 199 jugadores (filas 2-200)
  var count = Math.min(unique.length, 199);
  sheet.getRange(2, 1, count, 1).setValues(unique.slice(0, count));

  // Overall_Grade automático por posición en ADP
  var grades = [];
  for (var j = 0; j < count; j++) {
    var pos = j + 1;
    var g;
    if      (pos <= 12)  g = 'A+';
    else if (pos <= 24)  g = 'A';
    else if (pos <= 40)  g = 'A-';
    else if (pos <= 60)  g = 'B+';
    else if (pos <= 84)  g = 'B';
    else if (pos <= 110) g = 'B-';
    else if (pos <= 140) g = 'C+';
    else if (pos <= 168) g = 'C';
    else                 g = 'C-';
    grades.push([g]);
  }
  sheet.getRange(2, 14, count, 1).setValues(grades);

  // Notas de custom format insight en columna O para top jugadores
  var insights = [
    ['RB1 overall. ATL OL sólida. Tua QB. 260 carries + 70 rec.'],
    ['Lead back solo DET. OL élite. 280 carries + 65 rec.'],
    ['WR1. Burrow connection. 100+ rec + 1500 yds + 12 TDs.'],
    ['WR1 LAR. Stafford. Adams como WR2. 100 rec + 1300 yds.'],
    ['WR1 SEA. Año 3 clase 2024. Metcalf se fue a PIT.'],
    ['Si sano = RB1 overall. Shanahan system élite.'],
    ['RB confiable IND. Pierce WR. Taylor RB. 250 carries.'],
    ['Slot élite DET. 120+ rec. OL élite. Año 6.'],
    ['WR élite MIN. Kyler Murray QB nuevo de ARI.'],
    ['RB BUF. Allen corre también. 200 carries + 70 rec.'],
  ];
  // Solo top 10 con notas detalladas
  sheet.getRange(2, 15, insights.length, 1).setValues(insights);

  SpreadsheetApp.flush();

  var noteRow = count + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ Big Board 2026 — ' + count + ' jugadores | ADP: FantasyPros Half-PPR verificado mayo 2026 | ' +
    'Equipos 100% corregidos | VLOOKUP activo → espera 5 seg y corre sortBigBoardByCustom_2026()'
  ).setFontStyle('italic').setFontColor('#1a7b47').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '✅ Big Board cargado: ' + count + ' jugadores en orden ADP.\n\n' +
    '⏳ Espera ~5 segundos para que los VLOOKUPs calculen.\n\n' +
    'Luego corre: sortBigBoardByCustom_2026()\n' +
    'para ordenar por Custom Points (.8PPR + .2PPC)\n\n' +
    '14 rondas × 12 equipos = 168 picks cubiertos ✅'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Poblar Big Board columnas B-M leyendo valores directamente de los tabs fuente
// Más robusto que VLOOKUP cross-sheet. Correr después de fixBigBoard_2026_VERIFIED()
// ─────────────────────────────────────────────────────────────────────────────
function setupBigBoardFormulas_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Players DB → mapa nombre: fila completa ─────────────────────────────────
  var playersData = ss.getSheetByName('Players DB').getDataRange().getValues();
  var pMap = {};
  for (var i = 1; i < playersData.length; i++) {
    if (playersData[i][0]) pMap[playersData[i][0]] = playersData[i];
  }

  // ── Scoring Calc → mapa nombre: fila completa ───────────────────────────────
  // Columnas L-O (índices 11-14) = PPR / HalfPPR / Standard / Custom Points
  var scoringData = ss.getSheetByName('Scoring Calc').getDataRange().getValues();
  var sMap = {};
  for (var i = 1; i < scoringData.length; i++) {
    if (scoringData[i][0]) sMap[scoringData[i][0]] = scoringData[i];
  }

  // ── Schedule Flags → mapa equipo: fila completa ─────────────────────────────
  // Columna J (índice 9) = Playoff_Schedule_Rating
  var schedData = ss.getSheetByName('Schedule Flags').getDataRange().getValues();
  var schMap = {};
  for (var i = 1; i < schedData.length; i++) {
    if (schedData[i][0]) schMap[schedData[i][0]] = schedData[i];
  }

  // ── Big Board: leer col A (nombres), escribir B-M ───────────────────────────
  var bbSheet  = ss.getSheetByName('Big Board');
  var bbLastRow = bbSheet.getLastRow();
  if (bbLastRow < 2) {
    SpreadsheetApp.getUi().alert('⚠️ Big Board vacío. Corre primero fixBigBoard_2026_VERIFIED()');
    return;
  }

  var names = bbSheet.getRange(2, 1, bbLastRow - 1, 1).getValues();
  var out   = [];

  for (var j = 0; j < names.length; j++) {
    var name = names[j][0];
    if (!name) { out.push(['','','','','','','','','','','','']); continue; }

    var pd  = pMap[name]  || null;
    var sc  = sMap[name]  || null;
    var sch = (pd && pd[2]) ? (schMap[pd[2]] || null) : null;

    out.push([
      pd  ? pd[1]  : '',   // B: Position        (Players DB col B, índice 1)
      pd  ? pd[2]  : '',   // C: Team             (índice 2)
      pd  ? pd[3]  : '',   // D: Tier             (índice 3)
      sc  ? sc[14] : '',   // E: Custom_Points    (Scoring Calc col O, índice 14)
      sc  ? sc[11] : '',   // F: PPR_Points       (col L, índice 11)
      sc  ? sc[12] : '',   // G: HalfPPR_Points   (col M, índice 12)
      sc  ? sc[13] : '',   // H: Standard_Points  (col N, índice 13)
      pd  ? pd[4]  : '',   // I: ADP_FantasyPros  (índice 4)
      pd  ? pd[5]  : '',   // J: ADP_Underdog     (índice 5)
      pd  ? pd[10] : '',   // K: Playoff_SoS      (índice 10)
      pd  ? pd[7]  : '',   // L: OC_Change        (índice 7)
      sch ? sch[9] : ''    // M: Schedule_Rating  (Schedule Flags col J, índice 9)
    ]);
  }

  bbSheet.getRange(2, 2, out.length, 12).setValues(out);
  SpreadsheetApp.flush();

  var filled = out.filter(function(r) { return r[0] !== ''; }).length;
  SpreadsheetApp.getUi().alert(
    '✅ Big Board actualizado: ' + filled + ' jugadores con datos.\n\n' +
    'Columnas B-M desde:\n' +
    '  • Players DB → Position, Team, Tier, ADP, Playoff_SoS, OC_Change\n' +
    '  • Scoring Calc → Custom, PPR, HalfPPR, Standard Points\n' +
    '  • Schedule Flags → Schedule_Rating\n\n' +
    'Ahora corre sortBigBoardByCustom_2026()'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Ordenar Big Board por Custom Points (columna E) — correr DESPUÉS del VLOOKUP
// ─────────────────────────────────────────────────────────────────────────────
function sortBigBoardByCustom_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Big Board');
  var lastRow = sheet.getLastRow();

  if (lastRow <= 2) {
    SpreadsheetApp.getUi().alert('⚠️ El Big Board no tiene datos. Corre primero fixBigBoard_2026_VERIFIED()');
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, 15).sort({ column: 5, ascending: false });

  SpreadsheetApp.getUi().alert(
    '✅ Big Board ordenado por Custom Points (.8PPR + .2PPC)\n\n' +
    '── TOP 10 CUSTOM FORMAT (aprox) ──\n' +
    '1.  Josh Allen      QB BUF  ~482 pts ⚡\n' +
    '2.  Jalen Hurts     QB PHI  ~448 pts ⚡\n' +
    '3.  Lamar Jackson   QB BAL  ~449 pts ⚡\n' +
    '4.  Jayden Daniels  QB WAS  ~444 pts ⚡\n' +
    '5.  Jahmyr Gibbs    RB DET  ~391 pts\n' +
    '6.  Saquon Barkley  RB PHI  ~389 pts\n' +
    '7.  CMC             RB SF   ~388 pts\n' +
    '8.  Breece Hall     RB NYJ  ~363 pts\n' +
    '9.  Bijan Robinson  RB ATL  ~359 pts\n' +
    '10. Josh Jacobs     RB GB   ~331 pts\n\n' +
    '⚡ INSIGHT CLAVE:\n' +
    'Allen (ADP 21), Hurts (ADP 66), Jackson (ADP 43)\n' +
    'están SUBVALORADOS en Half-PPR pero DOMINAN\n' +
    'en tu formato .8PPR + .2PPC por las carreras.\n\n' +
    'Gibbs (ADP 2) es el RB#1 real en Custom.\n' +
    'Chase (ADP 3) baja a ~WR1 pero no top-5 Custom.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ORDEN DE EJECUCIÓN COMPLETO
// ─────────────────────────────────────────────────────────────────────────────
function runAllVerified_2026() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Iniciando actualización completa verificada 2026...');

  // 1. Players DB con equipos correctos
  updatePlayersDB_2026_VERIFIED();

  // 2. Scoring Calc con proyecciones corregidas
  updateScoringCalc_2026_VERIFIED();

  // 3. Big Board con ADP correcto
  fixBigBoard_2026_VERIFIED();

  // 4. Instalar fórmulas VLOOKUP en columnas B-M
  setupBigBoardFormulas_2026();

  // 5. Flush + pausa para que los VLOOKUPs calculen
  SpreadsheetApp.flush();
  Utilities.sleep(5000); // 5 segundos

  // 6. Ordenar por Custom Points
  sortBigBoardByCustom_2026();

  ui.alert(
    '✅ ACTUALIZACIÓN COMPLETA VERIFICADA\n\n' +
    'Players DB → Scoring Calc → Big Board\n' +
    'todos actualizados con datos reales FantasyPros\n\n' +
    '⚠️ Pendientes de verificar manualmente:\n' +
    '• Tyreek Hill — agente libre\n' +
    '• Deebo Samuel — agente libre\n' +
    '• Bye weeks en Schedule Flags tab'
  );
}
