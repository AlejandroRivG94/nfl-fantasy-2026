/**
 * NFL Fantasy 2026 — Projection Comparison Hub
 * ===============================================
 * Compara proyecciones de FantasyPros, ESPN y Yahoo.
 * Aplica tu fórmula Custom (0.8PPR + 0.2PPC) a cada fuente.
 *
 * FLUJO COMPLETO:
 *   1. setupProjectionHub_2026()          → crea tabs e instrucciones
 *   2. [Pega datos en FP_Raw / ESPN_Import / Yahoo_Import]
 *   3. parseFPProjections_2026()          → convierte FP_Raw al formato estándar
 *   4. buildProjectionComparison_2026()   → construye comparativa
 *   5. updateScoringCalcFromConsensus_2026() → actualiza Scoring Calc con promedio
 *
 * FORMATO ESTÁNDAR (todas las fuentes usan las mismas columnas):
 *   Player | Rec | RecYds | RecTDs | Carries | RushYds | RushTDs | Comp | PassYds | PassTDs | INTs
 *
 * FUENTES:
 *   FP        → fantasypros.com/nfl/projections/[pos].php → Export CSV → pega en FP_Raw
 *   ESPN      → fantasy.espn.com → Rankings → Export → pega en ESPN_Import (formato estándar)
 *   Yahoo     → football.fantasysports.yahoo.com → Projections → pega en Yahoo_Import
 *   Underdog  → underdogfantasy.com/blog → "Season Projections" → pega en Underdog_Import
 *   Vegas     → FanDuel/DraftKings season props → pega líneas en Vegas_Import
 *              (disponibles en agosto pre-temporada — una columna por stat)
 *   Sleeper   → fetchSleeperProjections_2026() intenta auto-fetch (API gratuita)
 *              (datos semanales, disponibles desde Week 1 preseason)
 */

// ─────────────────────────────────────────────────────────────────────────────
// FÓRMULA CUSTOM — igual que Scoring Calc
// ─────────────────────────────────────────────────────────────────────────────
function _customPts(rec, recYds, recTDs, carries, rushYds, rushTDs, comp, passYds, passTDs, ints) {
  return (rec    * 0.8) + (recYds  * 0.1) + (recTDs  * 6) +
         (carries* 0.2) + (rushYds * 0.1) + (rushTDs * 6) +
         (passYds* 0.04)+ (passTDs * 4)   + (ints    *-2);
}

function _n(v) { var f = parseFloat(v); return isNaN(f) ? 0 : f; }

// ─────────────────────────────────────────────────────────────────────────────
// 1. SETUP — crea tabs e instrucciones
// ─────────────────────────────────────────────────────────────────────────────
function setupProjectionHub_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tabs = [
    { name: 'Projection Hub',   fn: _setupHubTab    },
    { name: 'FP_Raw',           fn: _setupFPRawTab  },
    { name: 'FP_Import',        fn: _setupImportTab },
    { name: 'ESPN_Import',      fn: _setupImportTab },
    { name: 'Yahoo_Import',     fn: _setupImportTab },
    { name: 'Underdog_Import',  fn: _setupImportTab },
    { name: 'Vegas_Import',     fn: _setupVegasTab  },
    { name: 'Sleeper_Import',   fn: _setupImportTab },
  ];

  tabs.forEach(function(t) {
    var sheet = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
    sheet.clear();
    t.fn(sheet, t.name);
  });

  SpreadsheetApp.getUi().alert(
    '✅ Projection Hub creado.\n\n' +
    'PRÓXIMOS PASOS:\n' +
    '1. Ve a FP_Raw → lee las instrucciones → pega el CSV de FantasyPros\n' +
    '2. Corre parseFPProjections_2026()\n' +
    '3. Ve a ESPN_Import y Yahoo_Import → pega datos en formato estándar\n' +
    '4. Corre buildProjectionComparison_2026()'
  );
}

function _setupHubTab(sheet) {
  var r = 1;
  sheet.getRange(r, 1, 1, 9).merge()
    .setValue('📊  NFL FANTASY 2026 — PROJECTION COMPARISON HUB')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  r++;
  sheet.getRange(r, 1, 1, 9).merge()
    .setValue('Fuentes: FantasyPros Consensus · ESPN · Yahoo  |  Fórmula: 0.8 PPR + 0.2 PPC  |  Corre buildProjectionComparison_2026() para poblar')
    .setBackground('#1a1f36').setFontColor('#9ca3af')
    .setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;
  sheet.getRange(r, 1, 1, 9).merge()
    .setValue('⏳ Pendiente de datos — sigue el flujo en la parte superior de este archivo.')
    .setBackground('#fffbeb').setFontColor('#92400e').setHorizontalAlignment('center');
}

function _setupFPRawTab(sheet) {
  var r = 1;
  var title = '📥  FP_Raw — Pega aquí el CSV de FantasyPros (sin modificar)';
  sheet.getRange(r, 1, 1, 20).merge()
    .setValue(title)
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  var instructions = [
    ['FUENTE:',     'fantasypros.com/nfl/projections/qb.php  (y /rb.php, /wr.php, /te.php)'],
    ['SCORING:',    'Selecciona "Half Point PPR" en el filtro de la página'],
    ['TIMEFRAME:',  'Selecciona "Full Season" (no weekly)'],
    ['EXPORT:',     'Haz clic en "Export" (ícono de descarga) → selecciona CSV → abre el archivo'],
    ['PEGAR:',      'Selecciona TODO el contenido del CSV (Ctrl+A) y pégalo abajo en la fila 6'],
    ['POSICIONES:', 'Repite para QB, RB, WR, TE. Pega uno debajo del otro — el parser detecta el cambio de formato'],
    ['DESPUÉS:',    'Corre parseFPProjections_2026() para convertir al formato estándar'],
  ];
  instructions.forEach(function(row, i) {
    var bg = i % 2 === 0 ? '#eff6ff' : '#ffffff';
    sheet.getRange(r, 1).setValue(row[0]).setFontWeight('bold').setBackground(bg);
    sheet.getRange(r, 2, 1, 19).merge().setValue(row[1]).setBackground(bg);
    r++;
  });

  r++;
  sheet.getRange(r, 1, 1, 20).merge()
    .setValue('▼ PEGA EL CSV ABAJO DE ESTA LÍNEA')
    .setBackground('#374151').setFontColor('#f9fafb').setFontWeight('bold');
}

function _setupVegasTab(sheet) {
  var r = 1;
  sheet.getRange(r, 1, 1, 12).merge()
    .setValue('🎰  Vegas_Import — Líneas de temporada completa (FanDuel / DraftKings / BetMGM)')
    .setBackground('#1a1f36').setFontColor('#f5c518').setFontWeight('bold').setFontSize(13);
  r++;

  var notes = [
    ['⚠️ DISPONIBILIDAD:', 'Las season props se publican en AGOSTO, semanas antes del inicio de temporada regular. No existen antes de preseason.'],
    ['FUENTE FanDuel:',    'sportsbook.fanduel.com → NFL → Player Props → Season → filtra por "Passing/Rushing/Receiving Yards"'],
    ['FUENTE DraftKings:', 'sportsbook.draftkings.com → NFL → Player Props → Season → mismo proceso'],
    ['QUÉ BUSCAR:',        'Season Passing Yards O/U | Season Rushing Yards O/U | Season TDs O/U | Season Receptions O/U | Season Receiving Yards O/U'],
    ['CÓMO USAR LA LÍNEA:','El número de la línea (ej: "Josh Allen 640.5 Rush Yds") ES la proyección del mercado. Úsala como el valor de esa estadística.'],
    ['PARA STATS SIN PROP:','Si no hay prop de completions o carries exactas, usa 0 — el Custom Points se calculará con los stats disponibles.'],
    ['VENTAJA REAL:',      'Las líneas de Vegas son las proyecciones más "honestas" — están respaldadas por dinero real de apostadores profesionales (sharp money).'],
  ];
  notes.forEach(function(row, i) {
    var bg = i % 2 === 0 ? '#0d1117' : '#1a1f36';
    var tc = i === 0 ? '#f5c518' : '#e2e8f0';
    sheet.getRange(r, 1).setValue(row[0]).setBackground(bg).setFontColor(tc).setFontWeight('bold');
    sheet.getRange(r, 2, 1, 11).merge().setValue(row[1]).setBackground(bg).setFontColor('#e2e8f0');
    r++;
  });

  r++;
  sheet.getRange(r, 1, 1, 12).merge()
    .setValue('FORMATO: mismas columnas que las otras fuentes — usa la línea del prop como el valor de cada stat')
    .setBackground('#374151').setFontColor('#d1d5db').setFontStyle('italic');
  r++;

  var headers = ['Player','Rec','RecYds','RecTDs','Carries','RushYds','RushTDs','Comp','PassYds','PassTDs','INTs'];
  sheet.getRange(r, 1, 1, 11).setValues([headers])
    .setBackground('#f5c518').setFontColor('#0d1117').setFontWeight('bold');

  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 11; c++) sheet.setColumnWidth(c, 80);
}

function _setupImportTab(sheet, tabName) {
  var r = 1;
  var sources = {
    'ESPN_Import':     { url: 'fantasy.espn.com → My Team → Scoring → Player Rankings → Export CSV',               color: '#c2185b' },
    'Yahoo_Import':    { url: 'football.fantasysports.yahoo.com → Players → Projections → Export',                  color: '#7b1fa2' },
    'FP_Import':       { url: 'Generado automáticamente por parseFPProjections_2026()',                              color: '#1565c0' },
    'Underdog_Import': { url: 'underdogfantasy.com/blog → busca "Season Projections" → pega en formato estándar',   color: '#b45309' },
    'Sleeper_Import':  { url: 'Auto-populado por fetchSleeperProjections_2026() — o pega manualmente desde sleeper.com/stats', color: '#374151' },
  };
  var src = sources[tabName] || { url: '', color: '#374151' };

  sheet.getRange(r, 1, 1, 12).merge()
    .setValue('📥  ' + tabName + ' — Formato estándar')
    .setBackground(src.color).setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
  r++;
  sheet.getRange(r, 1, 1, 12).merge()
    .setValue('Fuente: ' + src.url)
    .setBackground('#f8fafc').setFontColor('#374151').setFontStyle('italic');
  r++;

  if (tabName !== 'FP_Import') {
    var instrucciones = [
      'INSTRUCCIONES: Descarga las proyecciones de ' + tabName.replace('_Import','') + ' y reorganiza las columnas al formato estándar de abajo.',
      'QBs: llena Comp, PassYds, PassTDs, INTs, Carries, RushYds, RushTDs. Rec=0.',
      'RBs: llena Carries, RushYds, RushTDs, Rec, RecYds, RecTDs. Pass=0.',
      'WRs/TEs: llena Rec, RecYds, RecTDs. Carries/Pass=0 (salvo jugadores que corren).',
      'No necesitas llenar las 11 columnas si no hay datos — deja en 0.',
    ];
    instrucciones.forEach(function(txt, i) {
      var bg = i % 2 === 0 ? '#f0fdf4' : '#ffffff';
      sheet.getRange(r, 1, 1, 12).merge().setValue(txt).setBackground(bg).setFontStyle('italic');
      r++;
    });
  }

  r++;
  // Standard headers
  var headers = ['Player','Rec','RecYds','RecTDs','Carries','RushYds','RushTDs','Comp','PassYds','PassTDs','INTs'];
  sheet.getRange(r, 1, 1, 11).setValues([headers])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');

  // Column widths
  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 11; c++) sheet.setColumnWidth(c, 80);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER — parsea una línea CSV respetando valores entre comillas
// Maneja casos como: Jahmyr Gibbs,"DET",267.6,"1,337.5","14.6"
// ─────────────────────────────────────────────────────────────────────────────
function _parseCSVLine(line) {
  var result = [];
  var current = '';
  var inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if (c === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += c;
    }
  }
  result.push(current.trim());
  return result;
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. PARSER — convierte CSV de FantasyPros al formato estándar en FP_Import
//
// Soporta DOS formatos de pegado:
//   A) Datos separados en columnas (paste normal de Google Sheets)
//   B) Todo en columna A como texto CSV (lo que pasa frecuentemente)
//
// Formato real del export de FantasyPros (simplificado sin Rank/AVG/TGT):
//   RB:    Player | Team | ATT | YDS | TDS | REC | YDS | TDS | FL | FPTS
//   QB:    Player | Team | CMP | ATT | PCT | YDS | TDS | INT | ATT | YDS | TDS | FL | FPTS
//   WR/TE: Player | Team | REC | YDS | TDS | ATT | YDS | TDS | FL | FPTS
// ─────────────────────────────────────────────────────────────────────────────
function parseFPProjections_2026() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet  = ss.getSheetByName('FP_Raw');
  var destSheet = ss.getSheetByName('FP_Import');

  if (!rawSheet || !destSheet) {
    SpreadsheetApp.getUi().alert('Corre primero setupProjectionHub_2026()');
    return;
  }

  var raw  = rawSheet.getDataRange().getValues();
  var rows = [];
  var mode = null;

  for (var i = 0; i < raw.length; i++) {
    var rowRaw = raw[i];

    // ── Detectar si los datos están en columna A como texto CSV ──────────────
    var isRawText = (String(rowRaw[1]).trim() === '') &&
                    (String(rowRaw[0]).indexOf(',') >= 0);

    var cols;
    if (isRawText) {
      cols = _parseCSVLine(String(rowRaw[0]));
    } else {
      cols = rowRaw.map(function(c) { return String(c).trim(); });
    }

    if (!cols || cols.length < 2) continue;

    var p0 = cols[0].trim();
    var p1 = cols[1] ? cols[1].trim() : '';
    var p0up = p0.toUpperCase();

    // ── Detectar fila de headers ──────────────────────────────────────────────
    if (p0up === 'PLAYER') {
      var hdrs = cols.map(function(c) { return c.toUpperCase().replace(/"/g, ''); });
      if (hdrs.indexOf('CMP') >= 0 || hdrs.indexOf('PCT') >= 0) {
        mode = 'QB';
      } else {
        var attIdx = hdrs.indexOf('ATT');
        var recIdx = hdrs.indexOf('REC');
        if (attIdx >= 0 && recIdx >= 0) {
          mode = (attIdx < recIdx) ? 'RB' : 'WR_TE';
        } else if (recIdx >= 0) {
          mode = 'WR_TE';
        }
      }
      continue;
    }

    // ── Saltar filas vacías o de separación ──────────────────────────────────
    if (!p0 || p0 === '' || p0 === '--') continue;
    // Saltar filas donde col0 es un número (rank) — formato viejo FP
    if (!isNaN(parseFloat(p0)) && p0.indexOf('.') === -1 && parseFloat(p0) < 300) continue;

    // ── Nombre del jugador ────────────────────────────────────────────────────
    // En el formato simplificado: Player=col0, Team=col1
    var playerName = p0.replace(/"/g, '').trim();
    // Eliminar sufijo de equipo si está en el nombre: "Josh Allen, BUF" → "Josh Allen"
    playerName = playerName.replace(/,\s*[A-Z]{2,3}$/, '').trim();

    if (!playerName || !mode) continue;

    var rec=0, recYds=0, recTDs=0, carries=0, rushYds=0, rushTDs=0,
        comp=0, passYds=0, passTDs=0, ints=0;

    if (mode === 'RB') {
      // Player(0) | Team(1) | ATT/carries(2) | RushYds(3) | RushTDs(4)
      // | Rec(5) | RecYds(6) | RecTDs(7) | FL(8) | FPTS(9)
      carries = _n(cols[2]);
      rushYds = _n(cols[3]);
      rushTDs = _n(cols[4]);
      rec     = _n(cols[5]);
      recYds  = _n(cols[6]);
      recTDs  = _n(cols[7]);

    } else if (mode === 'QB') {
      // Player(0) | Team(1) | CMP(2) | PassAtt(3) | PCT(4) | PassYds(5)
      // | PassTDs(6) | INT(7) | RushAtt(8) | RushYds(9) | RushTDs(10) | FL(11) | FPTS(12)
      comp    = _n(cols[2]);
      passYds = _n(cols[5]);
      passTDs = _n(cols[6]);
      ints    = _n(cols[7]);
      carries = _n(cols[8]);
      rushYds = _n(cols[9]);
      rushTDs = _n(cols[10]);

    } else if (mode === 'WR_TE') {
      // Player(0) | Team(1) | Rec(2) | RecYds(3) | RecTDs(4)
      // | RushAtt(5) | RushYds(6) | RushTDs(7) | FL(8) | FPTS(9)
      rec     = _n(cols[2]);
      recYds  = _n(cols[3]);
      recTDs  = _n(cols[4]);
      carries = _n(cols[5]);
      rushYds = _n(cols[6]);
      rushTDs = _n(cols[7]);
    }

    rows.push([playerName, rec, recYds, recTDs, carries, rushYds, rushTDs,
               comp, passYds, passTDs, ints]);
  }

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No se encontraron datos en FP_Raw.\n\n' +
      'Verifica que:\n' +
      '1. Pegaste el CSV debajo de la línea "▼ PEGA EL CSV ABAJO"\n' +
      '2. El CSV incluye la fila de headers (Player, Team, ATT...)\n' +
      '3. Pegaste al menos RB — QB, WR, TE pueden agregarse después'
    );
    return;
  }

  // Escribir en FP_Import
  var headerRow = _findHeaderRow(destSheet);
  var writeRow  = headerRow + 1;
  var existing  = destSheet.getLastRow() - writeRow + 1;
  if (existing > 0) destSheet.getRange(writeRow, 1, Math.max(existing, 1), 11).clearContent();
  destSheet.getRange(writeRow, 1, rows.length, 11).setValues(rows);

  SpreadsheetApp.flush();

  // Contar por posición
  var rbCount  = rows.filter(function(r) { return _n(r[4]) > 0 && _n(r[7]) === 0; }).length;
  var qbCount  = rows.filter(function(r) { return _n(r[7]) > 0; }).length;
  var wrteCount = rows.filter(function(r) { return _n(r[4]) === 0 && _n(r[1]) > 0 && _n(r[7]) === 0; }).length;

  SpreadsheetApp.getUi().alert(
    '✅ FantasyPros parseado: ' + rows.length + ' jugadores en FP_Import.\n\n' +
    '  Detección aproximada:\n' +
    '  QBs:    ~' + qbCount + '\n' +
    '  RBs:    ~' + rbCount + '\n' +
    '  WR/TE:  ~' + wrteCount + '\n\n' +
    '⚠️ Revisa FP_Import para confirmar que los datos se parsearon bien.\n' +
    'Si faltan posiciones, pega sus CSVs en FP_Raw y corre de nuevo.'
  );
}

function _findHeaderRow(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'Player') return i + 1; // 1-indexed
  }
  return 7; // fallback
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. BUILD COMPARISON — lee las 3 fuentes + Scoring Calc, construye la vista
// ─────────────────────────────────────────────────────────────────────────────
function buildProjectionComparison_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read each source into a map: playerName → Custom Points
  var sources = [
    { name: 'FP',       tab: 'FP_Import'       },
    { name: 'ESPN',     tab: 'ESPN_Import'      },
    { name: 'Yahoo',    tab: 'Yahoo_Import'     },
    { name: 'Underdog', tab: 'Underdog_Import'  },
    { name: 'Vegas',    tab: 'Vegas_Import'     },
    { name: 'Sleeper',  tab: 'Sleeper_Import'   },
  ];

  var sourceMaps = {};
  var sourceStats = {};
  sources.forEach(function(src) {
    var sheet = ss.getSheetByName(src.tab);
    if (!sheet) return;
    var map   = {};
    var stats = {};
    var hRow  = _findHeaderRow(sheet);
    var data  = sheet.getRange(hRow + 1, 1, Math.max(sheet.getLastRow() - hRow, 1), 11).getValues();
    data.forEach(function(row) {
      var name = String(row[0]).trim();
      if (!name) return;
      var pts = _customPts(_n(row[1]),_n(row[2]),_n(row[3]),_n(row[4]),_n(row[5]),
                           _n(row[6]),_n(row[7]),_n(row[8]),_n(row[9]),_n(row[10]));
      map[name]   = parseFloat(pts.toFixed(1));
      stats[name] = row; // keep raw stats for consensus calculation
    });
    sourceMaps[src.name]  = map;
    sourceStats[src.name] = stats;
  });

  // Read our current Scoring Calc
  var scSheet = ss.getSheetByName('Scoring Calc');
  var scData  = scSheet.getDataRange().getValues();
  var ourMap  = {};
  for (var i = 1; i < scData.length; i++) {
    var r = scData[i];
    if (!r[0]) continue;
    var pts = _customPts(_n(r[1]),_n(r[2]),_n(r[3]),_n(r[4]),_n(r[5]),
                         _n(r[6]),_n(r[7]),_n(r[8]),_n(r[9]),_n(r[10]));
    ourMap[String(r[0]).trim()] = parseFloat(pts.toFixed(1));
  }

  // Build comparison from Big Board player list
  var bbData = ss.getSheetByName('Big Board').getDataRange().getValues();
  var rows   = [];
  for (var b = 1; b < bbData.length; b++) {
    var name = String(bbData[b][0]).trim();
    var pos  = bbData[b][1];
    var team = bbData[b][2];
    if (!name || !pos) continue;

    var fp       = sourceMaps['FP']       ? (sourceMaps['FP'][name]       || '') : '';
    var espn     = sourceMaps['ESPN']     ? (sourceMaps['ESPN'][name]     || '') : '';
    var yahoo    = sourceMaps['Yahoo']    ? (sourceMaps['Yahoo'][name]    || '') : '';
    var underdog = sourceMaps['Underdog'] ? (sourceMaps['Underdog'][name] || '') : '';
    var vegas    = sourceMaps['Vegas']    ? (sourceMaps['Vegas'][name]    || '') : '';
    var sleeper  = sourceMaps['Sleeper']  ? (sourceMaps['Sleeper'][name]  || '') : '';
    var ours     = ourMap[name] || '';

    // Consensus = average of ALL available market sources (excludes blanks)
    var vals = [fp, espn, yahoo, underdog, vegas, sleeper].filter(function(v) { return v !== ''; });
    var consensus = vals.length > 0
      ? parseFloat((vals.reduce(function(a,b){return a+b;}, 0) / vals.length).toFixed(1))
      : '';

    var delta  = (consensus !== '' && ours !== '') ? parseFloat((consensus - ours).toFixed(1)) : '';
    var status = delta === '' ? '' :
                 Math.abs(delta) <= 10  ? '✅ Alineado' :
                 Math.abs(delta) <= 25  ? '⚠️ Revisar'  : '🔴 Corregir';

    rows.push([name, pos, team, fp, espn, yahoo, underdog, vegas, sleeper, consensus, ours, delta, status]);
  }

  // Write to Projection Hub
  var hubSheet = ss.getSheetByName('Projection Hub');
  hubSheet.clear();

  var r = 1;
  hubSheet.getRange(r, 1, 1, 10).merge()
    .setValue('📊  NFL FANTASY 2026 — PROJECTION COMPARISON HUB')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  r++;
  hubSheet.getRange(r, 1, 1, 10).merge()
    .setValue('Custom Points por fuente (0.8PPR + 0.2PPC) · Δ = Consenso − Nuestro actual · Última actualización: ' + new Date().toLocaleDateString())
    .setBackground('#1a1f36').setFontColor('#9ca3af').setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;

  // Table headers (13 columns now)
  var COLS = 13;
  hubSheet.getRange(r, 1, 1, COLS).setValues([[
    'Player', 'Pos', 'Team',
    'FP', 'ESPN', 'Yahoo', 'Underdog', 'Vegas 🎰', 'Sleeper',
    'Consenso', 'Nuestro', 'Δ', 'Status'
  ]]).setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  // Highlight Vegas column
  hubSheet.getRange(r, 8).setBackground('#f5c518').setFontColor('#0d1117');
  r++;

  // Data rows
  rows.forEach(function(row, idx) {
    var bg = idx % 2 === 0 ? '#f8fafc' : '#ffffff';
    hubSheet.getRange(r + idx, 1, 1, COLS).setValues([row]).setBackground(bg);

    // Color status cell (col 13)
    var status = row[12];
    var stBg   = status === '✅ Alineado' ? '#dcfce7' :
                 status === '⚠️ Revisar'  ? '#fef9c3' :
                 status === '🔴 Corregir' ? '#fee2e2' : bg;
    hubSheet.getRange(r + idx, 13).setBackground(stBg);

    // Delta color (col 12)
    var delta = row[11];
    if (delta !== '') {
      hubSheet.getRange(r + idx, 12)
        .setFontColor(delta > 0 ? '#15803d' : delta < 0 ? '#dc2626' : '#374151')
        .setFontWeight('bold');
    }

    // Vegas column highlight when populated
    if (row[7] !== '') {
      hubSheet.getRange(r + idx, 8).setBackground('#fefce8').setFontWeight('bold');
    }
  });
  r += rows.length;

  // Discrepancy summary
  r += 2;
  var bigDiffs = rows.filter(function(row) { return row[12] === '🔴 Corregir'; });
  if (bigDiffs.length > 0) {
    hubSheet.getRange(r, 1, 1, COLS).merge()
      .setValue('🔴 DISCREPANCIAS GRANDES (Δ > 25 pts) — Revisar antes del draft')
      .setBackground('#7f1d1d').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    bigDiffs.forEach(function(row, idx) {
      var bg = idx % 2 === 0 ? '#fef2f2' : '#fff5f5';
      hubSheet.getRange(r + idx, 1, 1, COLS).setValues([row]).setBackground(bg);
    });
  }

  // Column widths
  hubSheet.setColumnWidth(1, 185);
  hubSheet.setColumnWidth(2, 50);
  hubSheet.setColumnWidth(3, 50);
  [4,5,6,7,8,9,10,11].forEach(function(c) { hubSheet.setColumnWidth(c, 75); });
  hubSheet.setColumnWidth(12, 75);
  hubSheet.setColumnWidth(13, 110);
  hubSheet.setFrozenRows(r - rows.length - 1);

  SpreadsheetApp.flush();
  var filled   = rows.filter(function(row){ return row[9] !== ''; }).length;
  var corregir = bigDiffs.length;
  SpreadsheetApp.getUi().alert(
    '✅ Comparativa construida.\n\n' +
    '  Jugadores con consenso: ' + filled + '\n' +
    '  ✅ Alineados:  ' + rows.filter(function(row){return row[12]==='✅ Alineado';}).length + '\n' +
    '  ⚠️ Revisar:   ' + rows.filter(function(row){return row[12]==='⚠️ Revisar';}).length + '\n' +
    '  🔴 Corregir:  ' + corregir + '\n\n' +
    'Vegas 🎰 disponible en agosto — agrega las season props cuando salgan.\n' +
    'Corre updateScoringCalcFromConsensus_2026() para actualizar Scoring Calc.'
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. UPDATE SCORING CALC — promedio de todas las fuentes disponibles
// ─────────────────────────────────────────────────────────────────────────────
function updateScoringCalcFromConsensus_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sources = ['FP_Import', 'ESPN_Import', 'Yahoo_Import', 'Underdog_Import', 'Vegas_Import', 'Sleeper_Import'];
  var allStats = {}; // playerName → array of stat arrays

  sources.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    var hRow = _findHeaderRow(sheet);
    var data = sheet.getRange(hRow + 1, 1, Math.max(sheet.getLastRow() - hRow, 1), 11).getValues();
    data.forEach(function(row) {
      var name = String(row[0]).trim();
      if (!name) return;
      // Check if row has any non-zero stats
      var hasData = false;
      for (var c = 1; c <= 10; c++) { if (_n(row[c]) !== 0) { hasData = true; break; } }
      if (!hasData) return;
      if (!allStats[name]) allStats[name] = [];
      allStats[name].push(row);
    });
  });

  if (Object.keys(allStats).length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No hay datos en los tabs de importación. Pega los datos primero.');
    return;
  }

  // Build consensus stats for each player (average of all sources)
  var consensus = {};
  Object.keys(allStats).forEach(function(name) {
    var rows  = allStats[name];
    var avg   = [0,0,0,0,0,0,0,0,0,0];
    rows.forEach(function(row) {
      for (var c = 1; c <= 10; c++) avg[c-1] += _n(row[c]);
    });
    avg = avg.map(function(v) { return Math.round(v / rows.length); });
    consensus[name] = avg;
  });

  // Update Scoring Calc
  var scSheet  = ss.getSheetByName('Scoring Calc');
  var scData   = scSheet.getDataRange().getValues();
  var updated  = 0;
  var notFound = [];

  for (var i = 1; i < scData.length; i++) {
    var name = String(scData[i][0]).trim();
    if (!name) continue;
    if (consensus[name]) {
      var stats = consensus[name];
      scSheet.getRange(i + 1, 2, 1, 10).setValues([stats]);
      updated++;
    } else {
      notFound.push(name);
    }
  }

  // Also add players in consensus but not in Scoring Calc
  var newPlayers = 0;
  var lastRow    = scSheet.getLastRow();
  var scNames    = scData.slice(1).map(function(r) { return String(r[0]).trim(); });
  Object.keys(consensus).forEach(function(name) {
    if (scNames.indexOf(name) < 0) {
      var statsRow = [name].concat(consensus[name]);
      scSheet.getRange(lastRow + 1 + newPlayers, 1, 1, 11).setValues([statsRow]);
      newPlayers++;
    }
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅ Scoring Calc actualizado con consenso.\n\n' +
    '  Actualizados: ' + updated + '\n' +
    '  Nuevos agregados: ' + newPlayers + '\n' +
    '  Sin datos en fuentes: ' + notFound.length + '\n\n' +
    (notFound.length > 0
      ? '⚠️ Sin consenso: ' + notFound.slice(0,5).join(', ') + (notFound.length>5?' y '+(notFound.length-5)+' más...':'')
      : '✅ Todos los jugadores del Scoring Calc fueron actualizados.') + '\n\n' +
    'Recuerda correr setupBigBoardFormulas_2026() en BIGBOARD.gs\n' +
    'para que el Big Board refleje los nuevos Custom Points.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. SLEEPER AUTO-FETCH — intenta obtener proyecciones de la API pública
//    API gratuita, sin autenticación. Datos semanales disponibles desde preseason.
//    Llama a Sleeper player DB + projections y normaliza al formato estándar.
// ─────────────────────────────────────────────────────────────────────────────
function fetchSleeperProjections_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sleeper_Import');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Corre primero setupProjectionHub_2026()');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Intentando conectar con la API de Sleeper...\nEsto puede tomar 15-30 segundos.');

  try {
    // Step 1: Get Sleeper player DB (maps player_id → name + position)
    var playersResp = UrlFetchApp.fetch('https://api.sleeper.app/v1/players/nfl', {
      muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    if (playersResp.getResponseCode() !== 200) {
      ui.alert('⚠️ No se pudo conectar con Sleeper API.\nCódigo: ' + playersResp.getResponseCode() + '\nUsa el import manual.');
      return;
    }
    var playerDB = JSON.parse(playersResp.getContentText());

    // Build a map: full_name → { player_id, position, team }
    var nameMap = {};
    Object.keys(playerDB).forEach(function(pid) {
      var p = playerDB[pid];
      if (!p.full_name || !p.position) return;
      if (['QB','RB','WR','TE'].indexOf(p.position) < 0) return;
      nameMap[p.full_name] = { id: pid, pos: p.position, team: p.team || '' };
    });

    // Step 2: Get season projections (week 0 = season totals in some seasons)
    // Sleeper uses week 1-18 for regular season; aggregate weeks 1-17 for season totals
    // For pre-season: try the projections endpoint (may return empty until season starts)
    var projResp = UrlFetchApp.fetch(
      'https://api.sleeper.app/v1/projections/nfl/2026/1?season_type=regular&position[]=QB&position[]=RB&position[]=WR&position[]=TE',
      { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } }
    );

    var projData = {};
    if (projResp.getResponseCode() === 200) {
      var raw = JSON.parse(projResp.getContentText());
      // raw is array: [{ player_id, stats: { rec, rec_yd, rec_td, rush_att, rush_yd, rush_td, pass_cmp, pass_yd, pass_td, pass_int } }]
      if (Array.isArray(raw)) {
        raw.forEach(function(entry) {
          if (entry && entry.player_id && entry.stats) projData[entry.player_id] = entry.stats;
        });
      }
    }

    if (Object.keys(projData).length === 0) {
      ui.alert(
        '⚠️ Sleeper no tiene proyecciones para 2026 aún.\n\n' +
        'Las proyecciones de Sleeper se publican cuando comienza la preseason (julio-agosto).\n\n' +
        'Sleeper_Import queda disponible para import manual:\n' +
        'sleeper.com/stats/nfl/player/[id] → season totals'
      );
      return;
    }

    // Step 3: Match our Big Board players with Sleeper IDs and write projections
    var bbData  = ss.getSheetByName('Big Board').getDataRange().getValues();
    var hRow    = _findHeaderRow(sheet);
    var rows    = [];
    var matched = 0;

    for (var b = 1; b < bbData.length; b++) {
      var playerName = String(bbData[b][0]).trim();
      if (!playerName) continue;
      var sleeperInfo = nameMap[playerName];
      if (!sleeperInfo) continue;
      var stats = projData[sleeperInfo.id];
      if (!stats) continue;

      var rec     = _n(stats.rec)       || 0;
      var recYds  = _n(stats.rec_yd)    || 0;
      var recTDs  = _n(stats.rec_td)    || 0;
      var carries = _n(stats.rush_att)  || 0;
      var rushYds = _n(stats.rush_yd)   || 0;
      var rushTDs = _n(stats.rush_td)   || 0;
      var comp    = _n(stats.pass_cmp)  || 0;
      var passYds = _n(stats.pass_yd)   || 0;
      var passTDs = _n(stats.pass_td)   || 0;
      var ints    = _n(stats.pass_int)  || 0;

      rows.push([playerName, rec, recYds, recTDs, carries, rushYds, rushTDs, comp, passYds, passTDs, ints]);
      matched++;
    }

    if (rows.length === 0) {
      ui.alert('⚠️ Sleeper respondió pero no se pudieron matchear jugadores. Verifica los nombres en Big Board.');
      return;
    }

    var writeRow = hRow + 1;
    var existing = sheet.getLastRow() - writeRow + 1;
    if (existing > 0) sheet.getRange(writeRow, 1, Math.max(existing, 1), 11).clearContent();
    sheet.getRange(writeRow, 1, rows.length, 11).setValues(rows);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ Sleeper proyecciones importadas.\n\n' +
      '  Jugadores encontrados: ' + matched + '\n\n' +
      'Nota: datos de Week 1 — para season totals proyectados, Sleeper\n' +
      'consolida mejor en agosto cuando publica sus preseason projections.\n\n' +
      'Corre buildProjectionComparison_2026() para actualizar la comparativa.'
    );

  } catch (e) {
    ui.alert('❌ Error al conectar con Sleeper API:\n' + e.message + '\n\nUsa el import manual.');
  }
}
