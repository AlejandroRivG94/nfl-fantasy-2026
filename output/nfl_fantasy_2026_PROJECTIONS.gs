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
 *   FP   → fantasypros.com/nfl/projections/[pos].php → Export CSV → pega en FP_Raw
 *   ESPN → fantasy.espn.com → Rankings → Export        → pega en ESPN_Import (formato estándar)
 *   Yahoo→ football.fantasysports.yahoo.com → Projections → pega en Yahoo_Import (formato estándar)
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
    { name: 'Projection Hub',  fn: _setupHubTab    },
    { name: 'FP_Raw',          fn: _setupFPRawTab  },
    { name: 'FP_Import',       fn: _setupImportTab },
    { name: 'ESPN_Import',     fn: _setupImportTab },
    { name: 'Yahoo_Import',    fn: _setupImportTab },
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

function _setupImportTab(sheet, tabName) {
  var r = 1;
  var sources = {
    'ESPN_Import':  { url: 'fantasy.espn.com → My Team → Scoring → Player Rankings → Export',  color: '#c2185b' },
    'Yahoo_Import': { url: 'football.fantasysports.yahoo.com → Players → Projections → Export', color: '#7b1fa2' },
    'FP_Import':    { url: 'Generado automáticamente por parseFPProjections_2026()',              color: '#1565c0' },
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
// 2. PARSER — convierte CSV de FantasyPros al formato estándar en FP_Import
// ─────────────────────────────────────────────────────────────────────────────
function parseFPProjections_2026() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet  = ss.getSheetByName('FP_Raw');
  var destSheet = ss.getSheetByName('FP_Import');

  if (!rawSheet || !destSheet) {
    SpreadsheetApp.getUi().alert('Corre primero setupProjectionHub_2026()');
    return;
  }

  var raw   = rawSheet.getDataRange().getValues();
  var rows  = [];
  var mode  = null; // 'QB' | 'RB' | 'WR_TE'

  for (var i = 0; i < raw.length; i++) {
    var row  = raw[i];
    var col0 = String(row[0]).trim().toUpperCase();
    var col1 = String(row[1]).trim();

    // Detect header rows to determine current position group
    if (col1 === 'PLAYER' || col1 === 'Player') {
      // Detect mode by looking at column headers
      var headers = row.map(function(c) { return String(c).trim().toUpperCase(); });
      if (headers.indexOf('CMP') >= 0 || headers.indexOf('PCT') >= 0) {
        mode = 'QB';
      } else if (headers.indexOf('ATT') >= 0 && headers.indexOf('REC') >= 0) {
        // RB: has both ATT (rush) and REC
        // WR/TE: has REC before ATT (or no ATT)
        var attIdx = headers.indexOf('ATT');
        var recIdx = headers.indexOf('REC');
        mode = (attIdx < recIdx) ? 'RB' : 'WR_TE';
      } else if (headers.indexOf('REC') >= 0) {
        mode = 'WR_TE';
      }
      continue; // skip header row
    }

    // Skip empty rows, rank rows, or non-player rows
    if (!col1 || col1 === '' || !isNaN(parseInt(col0)) === false) continue;
    // FP format: row[0]=Rank (number), row[1]=Player Name
    if (isNaN(parseInt(col0))) continue;

    // Normalize player name (FP format: "FirstName LastName, TEAM")
    var playerName = col1.replace(/,\s*[A-Z]{2,3}$/, '').trim();

    var rec=0, recYds=0, recTDs=0, carries=0, rushYds=0, rushTDs=0,
        comp=0, passYds=0, passTDs=0, ints=0;

    if (mode === 'QB') {
      // QB: Rank|Player|Team|CMP|ATT|PCT|YDS|AVG|TD|INT|ATT|YDS|AVG|TD|FL|G|FPTS
      comp    = _n(row[3]);   // CMP
      // passYds = row[6], passTDs = row[8], ints = row[9]
      passYds = _n(row[6]);
      passTDs = _n(row[8]);
      ints    = _n(row[9]);
      carries = _n(row[10]);  // rushing ATT (second ATT column)
      rushYds = _n(row[11]);
      rushTDs = _n(row[13]);

    } else if (mode === 'RB') {
      // RB: Rank|Player|Team|ATT|YDS|AVG|TDS|REC|TGT|YDS|AVG|TDS|FL|G|FPTS
      carries = _n(row[3]);
      rushYds = _n(row[4]);
      rushTDs = _n(row[6]);
      rec     = _n(row[7]);
      recYds  = _n(row[9]);
      recTDs  = _n(row[11]);

    } else if (mode === 'WR_TE') {
      // WR/TE: Rank|Player|Team|REC|TGT|YDS|AVG|TDS|ATT|YDS|AVG|TDS|FL|G|FPTS
      rec     = _n(row[3]);
      recYds  = _n(row[5]);
      recTDs  = _n(row[7]);
      carries = _n(row[8]);   // rush attempts (usually 0 for most WRs)
      rushYds = _n(row[9]);
      rushTDs = _n(row[11]);
    }

    rows.push([playerName, rec, recYds, recTDs, carries, rushYds, rushTDs, comp, passYds, passTDs, ints]);
  }

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No se encontraron datos en FP_Raw. Verifica que pegaste el CSV correctamente (fila 10+).');
    return;
  }

  // Write to FP_Import (after header row at row 7)
  var headerRow = _findHeaderRow(destSheet);
  var writeRow  = headerRow + 1;
  var existing  = destSheet.getLastRow() - writeRow + 1;
  if (existing > 0) destSheet.getRange(writeRow, 1, Math.max(existing, 1), 11).clearContent();
  destSheet.getRange(writeRow, 1, rows.length, 11).setValues(rows);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅ FantasyPros parseado: ' + rows.length + ' jugadores en FP_Import.\n\n' +
    'Ahora llena ESPN_Import y Yahoo_Import,\n' +
    'luego corre buildProjectionComparison_2026()'
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
    { name: 'FP',    tab: 'FP_Import'    },
    { name: 'ESPN',  tab: 'ESPN_Import'  },
    { name: 'Yahoo', tab: 'Yahoo_Import' },
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

    var fp    = sourceMaps['FP']    ? (sourceMaps['FP'][name]    || '') : '';
    var espn  = sourceMaps['ESPN']  ? (sourceMaps['ESPN'][name]  || '') : '';
    var yahoo = sourceMaps['Yahoo'] ? (sourceMaps['Yahoo'][name] || '') : '';
    var ours  = ourMap[name] || '';

    // Consensus = average of available sources (excluding empty)
    var vals = [fp, espn, yahoo].filter(function(v) { return v !== ''; });
    var consensus = vals.length > 0
      ? parseFloat((vals.reduce(function(a,b){return a+b;}, 0) / vals.length).toFixed(1))
      : '';

    var delta  = (consensus !== '' && ours !== '') ? parseFloat((consensus - ours).toFixed(1)) : '';
    var status = delta === '' ? '' :
                 Math.abs(delta) <= 10  ? '✅ Alineado' :
                 Math.abs(delta) <= 25  ? '⚠️ Revisar'  : '🔴 Corregir';

    rows.push([name, pos, team, fp, espn, yahoo, consensus, ours, delta, status]);
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

  // Table headers
  hubSheet.getRange(r, 1, 1, 10).setValues([[
    'Player', 'Pos', 'Team',
    'FP Custom', 'ESPN Custom', 'Yahoo Custom',
    'Consenso', 'Nuestro Actual', 'Δ Consenso', 'Status'
  ]]).setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  // Data rows
  rows.forEach(function(row, idx) {
    var bg = idx % 2 === 0 ? '#f8fafc' : '#ffffff';
    hubSheet.getRange(r + idx, 1, 1, 10).setValues([row]).setBackground(bg);

    // Color status cell
    var status = row[9];
    var stBg   = status === '✅ Alineado' ? '#dcfce7' :
                 status === '⚠️ Revisar'  ? '#fef9c3' :
                 status === '🔴 Corregir' ? '#fee2e2' : bg;
    hubSheet.getRange(r + idx, 10).setBackground(stBg);

    // Color delta: green if consensus > ours, red if < ours
    var delta = row[8];
    if (delta !== '') {
      hubSheet.getRange(r + idx, 9)
        .setFontColor(delta > 0 ? '#15803d' : delta < 0 ? '#dc2626' : '#374151')
        .setFontWeight('bold');
    }
  });
  r += rows.length;

  // Discrepancy summary
  r += 2;
  var bigDiffs = rows.filter(function(row) { return row[9] === '🔴 Corregir'; });
  if (bigDiffs.length > 0) {
    hubSheet.getRange(r, 1, 1, 10).merge()
      .setValue('🔴 DISCREPANCIAS GRANDES (Δ > 25 pts) — Revisar antes del draft')
      .setBackground('#7f1d1d').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    bigDiffs.forEach(function(row, idx) {
      var bg = idx % 2 === 0 ? '#fef2f2' : '#fff5f5';
      hubSheet.getRange(r + idx, 1, 1, 10).setValues([row]).setBackground(bg);
    });
  }

  // Column widths
  hubSheet.setColumnWidth(1, 195);
  hubSheet.setColumnWidth(2, 55);
  hubSheet.setColumnWidth(3, 55);
  [4,5,6,7,8].forEach(function(c) { hubSheet.setColumnWidth(c, 90); });
  hubSheet.setColumnWidth(9, 90);
  hubSheet.setColumnWidth(10, 110);
  hubSheet.setFrozenRows(r - rows.length - 1); // freeze headers

  SpreadsheetApp.flush();
  var filled  = rows.filter(function(row){ return row[6] !== ''; }).length;
  var corregir = bigDiffs.length;
  SpreadsheetApp.getUi().alert(
    '✅ Comparativa construida.\n\n' +
    '  Jugadores con consenso: ' + filled + '\n' +
    '  ✅ Alineados:  ' + rows.filter(function(r){return r[9]==='✅ Alineado';}).length + '\n' +
    '  ⚠️ Revisar:   ' + rows.filter(function(r){return r[9]==='⚠️ Revisar';}).length + '\n' +
    '  🔴 Corregir:  ' + corregir + '\n\n' +
    'Corre updateScoringCalcFromConsensus_2026() para actualizar.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. UPDATE SCORING CALC — promedio de todas las fuentes disponibles
// ─────────────────────────────────────────────────────────────────────────────
function updateScoringCalcFromConsensus_2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sources = ['FP_Import', 'ESPN_Import', 'Yahoo_Import'];
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
