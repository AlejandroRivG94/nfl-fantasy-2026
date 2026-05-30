/**
 * NFL Fantasy 2026 — Draft Simulator
 * ====================================
 * Liga: 12 equipos | Snake | 14 rondas | 1 QB | 0.8PPR + 0.2PPC | Playoffs Wks 15-17
 *
 * INSTRUCCIONES:
 *   1. Corre createDraftSimulator_2026()  → genera el tab completo (posición 1 por defecto)
 *   2. Cambia la celda amarilla (B5) al número de tu posición (1-12)
 *   3. Corre updateSimulatorForPosition_2026() → regenera la tabla de picks
 *
 * SECCIONES:
 *   A. Tabla de picks por ronda (dinámica por posición)
 *   B. 4 builds recomendados según posición
 *   C. Guía de decisiones situacionales
 *   D. Análisis de stacks QB + receptor
 */

var SIM_LEAGUE  = 12;
var SIM_ROUNDS  = 14;

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function _simPickNum(round, position) {
  return round % 2 === 1
    ? (round - 1) * SIM_LEAGUE + position
    : round * SIM_LEAGUE - position + 1;
}

function _playersNearPick(players, pickNum) {
  // Returns up to 4 players whose ADP falls in [pickNum-2, pickNum+5]
  var results = [];
  for (var i = 0; i < players.length; i++) {
    var adp = players[i].adp;
    if (adp >= pickNum - 2 && adp <= pickNum + 5) results.push(players[i]);
    if (results.length >= 4) break;
  }
  return results;
}

function _readPlayers(ss) {
  var bbData = ss.getSheetByName('Big Board').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < bbData.length; i++) {
    var r = bbData[i];
    if (!r[0]) continue;
    var adp = parseFloat(r[8]);
    if (isNaN(adp) || adp <= 0) continue;
    list.push({
      name:   r[0],
      pos:    r[1]  || '',
      team:   r[2]  || '',
      custom: parseFloat(r[4]) || 0,
      adp:    adp,
      sos:    r[10] || '',
      sched:  r[12] || '',
      grade:  r[13] || ''
    });
  }
  list.sort(function(a, b) { return a.adp - b.adp; });
  return list;
}

function _byName(players) {
  var map = {};
  players.forEach(function(p) { map[p.name] = p; });
  return map;
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCIÓN PRINCIPAL
// ─────────────────────────────────────────────────────────────────────────────

function createDraftSimulator_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Draft Simulator') || ss.insertSheet('Draft Simulator');
  sheet.clear();

  var players = _readPlayers(ss);
  var pMap    = _byName(players);

  var nextRow = _writeHeader(sheet);
  nextRow     = _writePickTable(sheet, players, 1, nextRow); // default: position 1
  nextRow     = _writeBuilds(sheet, nextRow);
  nextRow     = _writeScenarios(sheet, nextRow);
  nextRow     = _writeStacks(sheet, players, pMap, nextRow);

  // Column widths
  var widths = [55, 70, 85, 160, 60, 160, 60, 160, 60, 220];
  for (var c = 1; c <= widths.length; c++) sheet.setColumnWidth(c, widths[c-1]);

  SpreadsheetApp.getUi().alert(
    '✅ Draft Simulator generado (Posición 1 por defecto).\n\n' +
    'Para tu posición real:\n' +
    '  1. Cambia la celda amarilla (B5) al número 1-12\n' +
    '  2. Corre updateSimulatorForPosition_2026()'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// REGENERAR SOLO LA TABLA DE PICKS (cambio de posición)
// ─────────────────────────────────────────────────────────────────────────────

function updateSimulatorForPosition_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Draft Simulator');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Corre primero createDraftSimulator_2026()');
    return;
  }

  var pos = parseInt(sheet.getRange(5, 2).getValue());
  if (isNaN(pos) || pos < 1 || pos > 12) {
    SpreadsheetApp.getUi().alert('Ingresa un número del 1 al 12 en la celda amarilla (B5).');
    return;
  }

  var players = _readPlayers(ss);

  // Clear and rewrite pick table section (rows 8 to 8 + ROUNDS + 6 phase labels)
  var TABLE_START = 8;
  var TABLE_ROWS  = SIM_ROUNDS + 6;
  sheet.getRange(TABLE_START, 1, TABLE_ROWS, 10).clearContent().clearFormat();
  _writePickTable(sheet, players, pos, TABLE_START);

  SpreadsheetApp.getUi().alert(
    '✅ Picks actualizados para Posición ' + pos + '\n' +
    '12 equipos | Snake | 14 rondas'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECCIÓN A: HEADER + CONFIG
// ─────────────────────────────────────────────────────────────────────────────

function _writeHeader(sheet) {
  var r = 1;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('🏈  NFL FANTASY 2026 — DRAFT SIMULATOR')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('12 equipos  |  Snake 14 rondas  |  1 QB  |  0.8 PPR + 0.2 PPC  |  Playoffs Semanas 15-17')
    .setBackground('#1a1f36').setFontColor('#9ca3af')
    .setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;

  // Position input
  sheet.getRange(r, 1).setValue('📍  TU POSICIÓN:')
    .setFontWeight('bold').setFontSize(12).setBackground('#f1f5f9');
  sheet.getRange(r, 2).setValue(1)
    .setBackground('#fef9c3').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center');
  sheet.getRange(r, 3, 1, 4).merge()
    .setValue('← Cambia este número (1-12) y corre  updateSimulatorForPosition_2026()')
    .setFontColor('#92400e').setFontStyle('italic').setBackground('#fffbeb');
  r += 2;

  return r; // row 7 — pick table starts here
}

// ─────────────────────────────────────────────────────────────────────────────
// SECCIÓN B: TABLA DE PICKS
// ─────────────────────────────────────────────────────────────────────────────

function _writePickTable(sheet, players, draftPos, startRow) {
  var r = startRow;

  // Header row
  sheet.getRange(r, 1, 1, 10).setValues([[
    'RONDA', 'PICK #', 'FASE', 'JUGADOR 1', 'POS', 'JUGADOR 2', 'POS', 'JUGADOR 3', 'POS', 'CONSEJO RÁPIDO'
  ]]).setBackground('#0d1117').setFontColor('#f5c518').setFontWeight('bold');
  r++;

  var phases = [
    { rounds: [1,2,3],      label: 'FUNDACIÓN',       bg: '#14532d', tip: 'Elite puro — toma el mejor Custom Points disponible sin importar posición' },
    { rounds: [4,5,6],      label: 'DIFERENCIACIÓN',  bg: '#7c2d12', tip: 'Aquí se ganan los drafts — VOR positivo + schedule de playoffs' },
    { rounds: [7,8,9],      label: 'PROFUNDIDAD',      bg: '#1e3a5f', tip: 'Breakout candidates, handcuffs de OL élite, TEs de valor' },
    { rounds: [10,11,12,13,14], label: 'UPSIDE + HANDCUFFS', bg: '#374151', tip: 'Lotería controlada — Yr2 WRs, RBs en situaciones, Kicker/D élite' }
  ];

  for (var rnd = 1; rnd <= SIM_ROUNDS; rnd++) {
    var pickNum = _simPickNum(rnd, draftPos);

    // Phase divider
    for (var ph = 0; ph < phases.length; ph++) {
      if (phases[ph].rounds[0] === rnd) {
        sheet.getRange(r, 1, 1, 10).merge()
          .setValue('── ' + phases[ph].label + '  (R' + phases[ph].rounds[0] +
                    (phases[ph].rounds.length > 1 ? '-R' + phases[ph].rounds[phases[ph].rounds.length-1] : '') +
                    ')  |  ' + phases[ph].tip)
          .setBackground(phases[ph].bg).setFontColor('#e2e8f0')
          .setFontWeight('bold').setFontStyle('italic');
        r++;
        break;
      }
    }

    var avail = _playersNearPick(players, pickNum);
    var rowData = ['R' + rnd, 'Pick ' + pickNum, ''];

    for (var p = 0; p < 3; p++) {
      if (avail[p]) {
        var flag = avail[p].sched === 'Brutal' ? ' ⚠️' :
                   avail[p].sched === 'Elite'  ? ' ✅' : '';
        rowData.push(avail[p].name + flag);
        rowData.push(avail[p].pos);
      } else {
        rowData.push(''); rowData.push('');
      }
    }

    // Quick tip for round
    rowData.push(_roundTip(rnd, draftPos, avail));

    var rowBg = rnd % 2 === 0 ? '#f8fafc' : '#ffffff';
    sheet.getRange(r, 1, 1, 10).setValues([rowData]).setBackground(rowBg);
    sheet.getRange(r, 1).setFontWeight('bold');
    sheet.getRange(r, 2).setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center');

    r++;
  }

  return r + 1;
}

function _roundTip(round, pos, avail) {
  var positions = avail.map(function(p) { return p.pos; });
  var hasQB  = positions.indexOf('QB') >= 0;
  var hasTE  = positions.indexOf('TE') >= 0;
  var hasRB  = positions.indexOf('RB') >= 0;

  if (round === 1) {
    if (pos <= 3)  return 'Bijan/Gibbs/Chase disponibles. RB o WR élite, no QB.';
    if (pos <= 6)  return 'RB top-5 o WR1. Si Chase sigue disponible → tómalo.';
    if (pos <= 9)  return 'Ventana de WR1 y RB sólidos. Evalúa el ADP window.';
    return 'Posición 10-12: mejor ADP window del R1 para WRs de élite.';
  }
  if (round === 2) {
    if (hasQB) return '⚡ QB disponible — en 1QB estándar solo si es Allen/Hurts/Jackson.';
    if (hasTE) return 'TE en R2 solo si es McBride/Bowers. De lo contrario RB/WR.';
    if (hasRB) return 'RB sólido en R2 = fundación completa. Prioriza.';
    return 'WR de alto piso con schedule favorable.';
  }
  if (round === 3) return hasQB ? '⚡ Allen/Hurts en R3 = steal masivo en tu formato. Tómalo.' : 'Completa tu core. Sin RB aún? Prioridad máxima.';
  if (round === 4) return hasQB ? 'Lamar/Daniels en R4: considera si aún no tienes QB.' : 'VOR check — ¿Playoff SoS favorable?';
  if (round === 5) return 'Zona de mayor valor del draft. Jugadores con VOR > 60.';
  if (round === 6) return hasQB ? 'Hurts ADP ~66: si está aquí y no tienes QB, es obligatorio.' : 'Depth RB o WR con ceiling alto.';
  if (round >= 7 && round <= 9) return 'Busca Yr2 jugadores y receptores en OL élite.';
  if (round >= 10) return 'Handcuffs de valor (Bigsby/PHI, Pacheco/DET). Kicker/DST R13-14.';
  return '';
}

// ─────────────────────────────────────────────────────────────────────────────
// SECCIÓN C: 4 BUILDS
// ─────────────────────────────────────────────────────────────────────────────

function _writeBuilds(sheet, startRow) {
  var r = startRow;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('🏗️  BUILDS RECOMENDADOS — Estrategias de draft para tu formato 0.8PPR + 0.2PPC')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  var builds = [
    {
      name:   '⚡ BUILD A — QB ÉLITE TEMPRANO  (mejor para posiciones 5-12)',
      color:  '#7c3aed',
      desc:   'Anclar el roster con un QB top-3 en las primeras 3 rondas. La mayor ventaja sistémica en formato 0.2PPC.',
      picks:  [
        ['R1 (picks 5-12)',  'RB o WR de élite — Gibbs, Chase, Nacua, Barkley, Jefferson'],
        ['R2 (pick cae)',    '⚡ JOSH ALLEN si está disponible. Es el pick más valioso del draft en tu formato (~482 Custom Pts, VOR +174 sobre replacement QB).'],
        ['R2 (si Allen fue)','RB2 sólido — Hall, Etienne, J.Williams. Necesitas 2 RBs en R1-R3.'],
        ['R3',              'Si Allen fue en R2: otro RB/WR. Si NO tienes QB aún: Lamar (ADP 43) o Hurts (si cae a R3 en tu pick).'],
        ['R4-R5',           'WR depth (Flowers, Metcalf, McConkey) o TE si McBride/Bowers sigue aquí.'],
        ['R6-R8',           'Jalen Hurts (ADP 66) si aún no tienes QB y Allen fue. Obligatorio en tu formato.'],
        ['R9-R11',          'RB3 con upside, Yr2 WRs (Emeka Egbuka, Tetairoa McMillan).'],
        ['R12-R14',         'Handcuff RB1 tuyo (Tank Bigsby si tienes Barkley), Kicker BUF/PHI/CIN, D/ST Eagles/Ravens.'],
      ],
      insight: '⭐ POR QUÉ FUNCIONA: En 1 QB estándar, Allen cayendo a R2-R3 es el mayor mismatch entre valor real y ADP. Su ventaja sobre el replacement QB (~174 pts) es mayor que la ventaja de cualquier WR1 o RB1 sobre su replacement. Una vez que tienes un QB top-3, puedes ignorar la posición por 14 rondas.'
    },
    {
      name:   '🏃 BUILD B — RB HEAVY  (mejor para posiciones 1-4)',
      color:  '#15803d',
      desc:   'Asegurar dos RBs de élite en R1-R2. Explotar el cliff de RB que ocurre después del pick 25.',
      picks:  [
        ['R1 (picks 1-4)',  'RB top-4: Bijan, Gibbs, CMC, Barkley. OL élite = piso altísimo toda la temporada.'],
        ['R2 (picks 25-21)', 'WR1 si hay (Chase, Nacua, JSN) O segundo RB si hay drop-off (Cook, Hall, Etienne). Evalúa el ADP.'],
        ['R3',              'Cierra el core: si tienes 2 RBs, WR1 es prioridad. Si tienes 1 RB + WR, segundo RB ahora.'],
        ['R4-R5',           'QB si Allen/Lamar sigue. De lo contrario WR depth con schedule favorable.'],
        ['R5-R6',           'Jalen Hurts (ADP 66) = obligatorio si no tienes QB aún. No esperes a R7+ en tu formato.'],
        ['R7-R9',           'TE si McBride/Bowers cayó (improbable pero posible), o WR2 con ceiling alto.'],
        ['R10-R14',         'Handcuffs de TUS RBs es crítico: Tank Bigsby (Barkley), Pacheco (Gibbs), Chandler (CMC).'],
      ],
      insight: '⭐ POR QUÉ FUNCIONA: Los RBs en OL élite (PHI, DET, SF, BUF) tienen piso de 280 carries garantizados. Bajan el riesgo de tu roster semana a semana. El riesgo de lesión se mitiga teniendo 2 RBs de élite y sus handcuffs. El roster resiste semanas malas mejor que cualquier otro build.'
    },
    {
      name:   '🔗 BUILD C — STACK WR+QB  (mejor para posiciones 3-9)',
      color:  '#0369a1',
      desc:   'Draftear un WR1 élite en R1 y hacer stack con su QB en R4-R6. Máximo ceiling en semanas de playoff.',
      picks:  [
        ['R1',              "WR1 de stack élite: Ja'Marr Chase (CIN) o Amon-Ra St. Brown (DET). Máximo proyección de Custom Pts."],
        ['R2',              'RB sólido — Hall, Cook, Etienne. Necesitas RB1 antes del R3.'],
        ['R3',              'Segundo RB o WR2. Construye la fundación antes del stack.'],
        ['R4-R5',           "STACK: si tienes Chase → Burrow (ADP 49, R4-R5). Si tienes ARSB → Goff (ADP 102, R9 = valor tardío). Si tienes otro WR → evalúa."],
        ['R5-R6',           'Jalen Hurts si no tienes QB aún (o si el stack ya lo resuelve, toma el mejor disponible).'],
        ['R7-R8',           'TE con valor (Tyler Warren IND, Harold Fannin CLE) o WR3 con upside.'],
        ['R9-R11',          'Yr2 WRs: Emeka Egbuka (TB), Tetairoa McMillan (CAR), Carnell Tate (TEN).'],
        ['R12-R14',         'Extensión del stack si el precio es correcto (Tee Higgins = CIN triple stack con Chase+Burrow).'],
      ],
      insight: '⭐ POR QUÉ FUNCIONA: En semanas de championship (Wks 15-17), tener Chase + Burrow significa que cuando CIN tiene el partido de su vida, AMBOS explotan juntos. Ceiling stacking es la estrategia correcta si tu objetivo es ganar el campeonato, no solo llegar. El riesgo: si Burrow se lesiona antes de R15, Chase pierde valor. Mitígalo teniendo RBs de piso alto.'
    },
    {
      name:   '🚀 BUILD D — ZERO RB  (mejor para posiciones 7-12)',
      color:  '#9a3412',
      desc:   'Ignorar RB en R1-R2, acumular WRs de élite y QBs de valor, tomar RBs de alto upside en R5+.',
      picks:  [
        ['R1',              'WR1 absoluto: Chase, Nacua, JSN, Jefferson, Adams. Los mejores WRs de la liga.'],
        ['R2',              'Segundo WR1: Amon-Ra, Malik Nabers, A.J. Brown, Rashee Rice. Dos WRs top-30 = base sólida.'],
        ['R3',              'QB élite si Allen/Hurts está disponible (el "cheat code" de Zero RB es no necesitar RB en R3). O WR3 top.'],
        ['R4-R5',           'PRIMER RB: Chase Brown (CIN, ADP 19.3), Cam Skattebo (NYG, ADP 42), Bucky Irving (TB, ADP 48.3). RBs con volumen garantizado.'],
        ['R5-R6',           'Segundo RB: Javonte Williams (DAL), Travis Etienne (NO), Jaleel McLaughlin (DEN).'],
        ['R7-R8',           'TE: Tyler Warren (IND), Harold Fannin (CLE), Kyle Pitts (ATL). Valor real en R7.'],
        ['R9-R14',          'Handcuffs estratégicos: Isiah Pacheco (DET handcuff de Gibbs), Tank Bigsby (PHI). Lotería de RB.'],
      ],
      insight: '⭐ POR QUÉ FUNCIONA: Los mejores WRs tienen mayor consistencia semana a semana que los RBs (menos lesiones, menos platoon situations). Si acumulas WRs de élite y QB top, tienes alto piso. Los RBs de valor en R5-R7 (Chase Brown, Skattebo, Irving) tienen upside de RB1 si el titular se lesiona. Riesgo: si tus RBs de valor no rinden, estás expuesto en el flex.'
    }
  ];

  for (var b = 0; b < builds.length; b++) {
    var bld = builds[b];
    r++;

    // Build title
    sheet.getRange(r, 1, 1, 10).merge()
      .setValue(bld.name)
      .setBackground(bld.color).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(12);
    r++;

    sheet.getRange(r, 1, 1, 10).merge()
      .setValue(bld.desc)
      .setBackground('#f8fafc').setFontColor('#374151')
      .setFontStyle('italic');
    r++;

    // Picks table
    sheet.getRange(r, 1, 1, 2).setValues([['RONDA / PICK', 'RECOMENDACIÓN']])
      .setBackground('#e2e8f0').setFontWeight('bold');
    sheet.getRange(r, 3, 1, 8).clearContent();
    r++;

    for (var pk = 0; pk < bld.picks.length; pk++) {
      var pkBg = pk % 2 === 0 ? '#f8fafc' : '#ffffff';
      sheet.getRange(r, 1).setValue(bld.picks[pk][0]).setFontWeight('bold').setBackground(pkBg);
      sheet.getRange(r, 2, 1, 9).merge().setValue(bld.picks[pk][1]).setBackground(pkBg);
      r++;
    }

    // Insight
    sheet.getRange(r, 1, 1, 10).merge()
      .setValue(bld.insight)
      .setBackground('#fffbeb').setFontColor('#92400e')
      .setFontStyle('italic').setFontWeight('normal');
    r++;
  }

  return r + 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECCIÓN D: DECISIONES SITUACIONALES
// ─────────────────────────────────────────────────────────────────────────────

function _writeScenarios(sheet, startRow) {
  var r = startRow;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('🎯  GUÍA DE DECISIONES SITUACIONALES — Qué hacer cuando el draft no va según el plan')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  var scenarios = [
    {
      title: '⚡ Josh Allen cae a tu pick de Ronda 2 o 3',
      color: '#7c3aed',
      rows: [
        ['Contexto',          'Allen ADP ~21. En posiciones 5-12, tu R2 pick está entre 13-20. Allen puede caer hasta pick 25-30 en muchos drafts cuando los managers subestiman los QBs corredores en formato 0.2PPC.'],
        ['TÓMALO',            'No lo pienses. Allen ~482 Custom Pts = jugador #1 del draft en TU formato. Su VOR (+174 sobre replacement QB) es mayor que el VOR de cualquier WR1 o RB1 sobre su replacement.'],
        ['Post-Allen',        'Ya RESOLVISTE la posición QB por 14 semanas. En rondas 4-8 pivota DURO a RBs y WRs. No necesitas a Lamar, Hurts ni Daniels. Si alguno cae, deja que otro manager lo agarre.'],
        ['¿Cuándo NO tomarlo?','Si ya tienes a Lamar o Hurts de rondas anteriores (R1-R2). Duplicar QB top-3 desperdicia un pick valioso. Dos QBs de élite = 1 spot ocupado que podría ser un RB/WR ganador.'],
        ['Handcuff Allen',    'DJ Moore (WR1 BUF, ADP 46) en R4 = extensión del ecosistema Allen. No es un handcuff pero sí una correlación: cuando Allen boom, Moore boom.'],
      ]
    },
    {
      title: '📊 Tienes a Ja\'Marr Chase — ¿Vale el stack con Burrow?',
      color: '#0369a1',
      rows: [
        ['El número',         "Chase ADP 3 (R1). Burrow ADP 49 (R4-R5). Si drafteaste Chase en R1, Burrow estará disponible en tu R4-R5 pick en casi todos los drafts."],
        ['Valor del stack',   'Burrow proyecta ~325 Custom Pts. Ambos juntos = 634 Custom Pts de proyección base. El REAL valor es la correlación: en una semana de 30+ pts de Burrow, Chase típicamente tiene 25+ pts también.'],
        ['Sí al stack si...',  'No tienes QB aún en R4-R5 (Burrow resuelve la posición + stack), o si tienes un QB mid-tier (Daniels, Maye) y Burrow está disponible como upgrade en R5.'],
        ['No al stack si...', 'Ya tienes Allen o Hurts. No necesitas Burrow y el spot vale más siendo un RB/WR diferenciador. Burrow sin Chase también funciona bien (35 PassTDs proyectados).'],
        ['Extensión del stack','Tee Higgins (ADP 37.3, R3-R4) + Chase + Burrow = triple stack CIN. Máximo ceiling pero triple correlación negativa. Si CIN pierde por 35-7, los tres se hunden juntos. Úsalo para el roster de playoff, no como base.'],
        ['Resumen ejecutivo',  'Chase primero (R1) → Burrow en R4-R5 si QB slot libre = estrategia A+. Chase solo sin Burrow también funciona porque Chase produce puntos independientemente (rec yards + TDs).'],
      ]
    },
    {
      title: '⚠️ Llegaste a Ronda 4 sin ningún RB top-25',
      color: '#b91c1c',
      rows: [
        ['Cuándo pasa',       'Tomaste WR1 (Chase) + QB élite (Allen) + WR2 (ARSB) en R1-R3. Zero RB intencional o por disponibilidad.'],
        ['No pánico',         'Chase Brown (CIN, ADP 19.3) y Cam Skattebo (NYG, ADP 42) son los RBs de mayor valor en R4-R5. Volumen garantizado con upside de RB1.'],
        ['R4-R5 RB targets',  'Chase Brown (CIN): lead back con Burrow. Skattebo (NYG): ADP 42, Año 2 con workload full. Bucky Irving (TB): Mayfield lo usa en pase. Avery Williams (ATL, R5+): handcuff valioso.'],
        ['Regla de mínimos',  'Necesitas 2 RBs en los primeros 7 rounds para no quedar expuesto en el flex. Si llegas a R7 con 0 RBs, hay problema.'],
        ['Flex solution',     'Tener 3-4 WRs de élite con alto piso (ARSB = 120+ rec) te da opciones semanales en el flex sin necesitar RBs premium. La estrategia Zero RB funciona si tus WRs son top-tier.'],
      ]
    },
    {
      title: '🎯 McBride o Bowers están disponibles en tu pick de R2',
      color: '#047857',
      rows: [
        ['El dilema',         'McBride ADP 16.7, Bowers ADP 17.3. Si estás en posición 7-12, podrías verlos caer a tu R2 pick (13-17). ¿Los tomas sobre un RB/WR de R2?'],
        ['Por qué sí',        'El TE cliff es real: McBride/Bowers producen ~250 Custom Pts. El TE #13 produce ~158 pts. Ventaja de +100 pts sobre replacement que NINGÚN WR/RB de R2 puede igualar.'],
        ['Por qué no',        'Tomando TE en R2, te quedas sin RB2 hasta R3-R4. En esa zona el drop de RB ya es notable. El trade-off es: +100 pts en TE, -20-40 pts en RB (vs. tomar RB de R2).'],
        ['La regla',          'Si el mejor RB disponible en tu pick de R2 tiene ADP > 28 (ronda tardía R3 territory), toma el TE de élite. Si hay RBs ADP 20-28 disponibles, compara VOR antes de decidir.'],
        ['Post-TE élite',     'Ya NO necesitas Tucker Kraft (GB, R6), Harold Fannin (CLE, R6), ni Kyle Pitts (ATL, R7). Ese pick libre lo conviertes en RB/WR diferenciador o stack QB.'],
      ]
    },
    {
      title: '📅 Tu star player tiene Playoff SoS "Brutal" (Wks 15-17)',
      color: '#92400e',
      rows: [
        ['Contexto',          'Ejemplos: KC (Brutal), PIT (Brutal), WAS (Brutal), NYG (Brutal). Si tienes a Rashee Rice (KC) o George Pickens (PIT), tus playoffs están en peligro.'],
        ['No descartes',      'El SoS es un desempate, no un criterio primario. Rice proyecta ~250 Custom Pts EN TEMPORADA REGULAR con Mahomes. Ese valor supera el riesgo de playoff schedule malo.'],
        ['Mitiga el riesgo',  'Complementa con jugadores de playoff schedule Elite/Good en tu lineup. Si tienes Rice (Brutal), asegúrate de tener Barkley (PHI, Good) o Cook (BUF, Elite) como tu otro RB.'],
        ['El peor escenario', 'Dos o más starters con Brutal en Wks 15-17. Eso sí es problema. Un jugador Brutal sobrevivible; dos te dejan sin opciones de bench para las semanas que importan.'],
        ['Estrategia late',   'En rounds 10-12, prioriza jugadores con schedule Elite/Good en playoffs sobre jugadores más famosos con schedule Brutal. La diferencia de matchup en R15 puede valer 10-15 pts.'],
      ]
    },
  ];

  for (var s = 0; s < scenarios.length; s++) {
    var sc = scenarios[s];

    sheet.getRange(r, 1, 1, 10).merge()
      .setValue(sc.title)
      .setBackground(sc.color).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(12);
    r++;

    for (var sr = 0; sr < sc.rows.length; sr++) {
      var bg = sr % 2 === 0 ? '#f8fafc' : '#ffffff';
      sheet.getRange(r, 1).setValue(sc.rows[sr][0]).setFontWeight('bold').setBackground(bg);
      sheet.getRange(r, 2, 1, 9).merge().setValue(sc.rows[sr][1]).setBackground(bg).setWrap(true);
      r++;
    }
    r++;
  }

  return r + 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECCIÓN E: ANÁLISIS DE STACKS
// ─────────────────────────────────────────────────────────────────────────────

function _writeStacks(sheet, players, pMap, startRow) {
  var r = startRow;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('🔗  ANÁLISIS DE STACKS — QB + Receptor del mismo equipo')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('Stack = QB + WR/TE del mismo equipo. Beneficio: correlación positiva en semanas de alto scoring. Riesgo: ambos bajan juntos si el QB tiene mal partido o se lesiona.')
    .setBackground('#1a1f36').setFontColor('#9ca3af')
    .setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;

  // Table headers
  sheet.getRange(r, 1, 1, 10).setValues([[
    'QB', 'Custom QB', 'Receptor', 'Custom Rec', 'Total Stack',
    'ADP QB', 'ADP Rec', 'Ronda Stack', 'Rating', 'Análisis del stack'
  ]]).setBackground('#1a1f36').setFontColor('#f5c518').setFontWeight('bold');
  r++;

  var stacks = [
    { qb: 'Joe Burrow',     rec: "Ja'Marr Chase",      rating: 'A+',  color: '#14532d' },
    { qb: 'Josh Allen',     rec: 'DJ Moore',            rating: 'A+',  color: '#14532d' },
    { qb: 'Jalen Hurts',    rec: 'A.J. Brown',          rating: 'A',   color: '#15803d' },
    { qb: 'Lamar Jackson',  rec: 'Zay Flowers',         rating: 'A',   color: '#15803d' },
    { qb: 'Jayden Daniels', rec: 'Terry McLaurin',      rating: 'B+',  color: '#0369a1' },
    { qb: 'Jared Goff',     rec: 'Amon-Ra St. Brown',   rating: 'B+',  color: '#0369a1' },
    { qb: 'Jordan Love',    rec: 'Davante Adams',       rating: 'B',   color: '#6b21a8' },
    { qb: 'Caleb Williams', rec: 'Rome Odunze',         rating: 'B',   color: '#6b21a8' },
    { qb: 'Patrick Mahomes',rec: 'Rashee Rice',         rating: 'B-',  color: '#92400e' },
  ];

  var stackAnalysis = {
    'Joe Burrow':      "El stack más explosivo. Chase + Burrow = 634 Custom Pts proyectados. Cuando CIN boom, ambos boom. Burrow cae R4-R5 (ADP 49). Tómalo si Chase es tu R1 y tienes slot QB libre.",
    'Josh Allen':      "Allen ya produce ~482 Custom Pts solo (QB#1 custom). Moore WR1 BUF (ADP 46). Stack más alto ceiling del draft: suma ~700 Custom Pts. Ambos pueden caer R1 tardío + R4.",
    'Jalen Hurts':     "Stack establecido. Hurts 448 Custom Pts + Brown 250 = 698. PHI playoff schedule 'Good'. Brown ADP 28 (R2-R3), Hurts ADP 66 (R5-R6). Stack económico en rounds diferentes.",
    'Lamar Jackson':   "Lamar 449 Custom Pts + Flowers 220 = 669. Ambos en R4 (ADPs 43 y 41.7). Stack de mismo round = posible en posiciones 1-4. BAL run-first limita el ceiling de Flowers vs Chase.",
    'Jayden Daniels':  "Daniels 444 Custom Pts + McLaurin 220 = 664. Ambos ADP ~57. WAS playoff schedule 'Brutal' es el principal riesgo. Mejor como stack secundario que como núcleo del roster.",
    'Jared Goff':      "Stack asimétrico de alto valor. ARSB R1 (ADP 7.7) + Goff R9 (ADP 102). Goff como QB2 tardío con ARSB ya en tu roster = stack de bajo costo. DET schedule 'Neutral'.",
    'Jordan Love':     "Stack de playoff schedule. GB tiene schedule Elite en Wks 15-17. Love R10 (ADP 115) + Adams R4-R5 (ADP 52.7). Stack tardío diseñado para explotar en championship.",
    'Caleb Williams':  "Stack de upside. Williams Año 3 + Odunze Año 3 = curva de crecimiento. CHI OL 'Pobre' es el riesgo. Schedule playoff 'Good'. Williams ADP 67, Odunze ADP 63.",
    'Patrick Mahomes': "Stack de piso alto pero ceiling limitado por schedule 'Brutal' en playoffs. Rice regresó de lesión y es WR1 KC. Mahomes siempre produce pero KC Wks 15-17 = difícil. Mejor como stack secundario.",
  };

  for (var st = 0; st < stacks.length; st++) {
    var s   = stacks[st];
    var qbd = pMap[s.qb];
    var rcd = pMap[s.rec];

    var qbCust  = qbd ? qbd.custom.toFixed(0) : 'N/A';
    var rcCust  = rcd ? rcd.custom.toFixed(0) : 'N/A';
    var total   = (qbd && rcd) ? (qbd.custom + rcd.custom).toFixed(0) : 'N/A';
    var qbAdp   = qbd ? qbd.adp.toFixed(1) : 'N/A';
    var rcAdp   = rcd ? rcd.adp.toFixed(1) : 'N/A';
    var maxAdp  = Math.max(parseFloat(qbAdp) || 0, parseFloat(rcAdp) || 0);
    var stackRd = maxAdp > 0 ? 'R' + Math.ceil(maxAdp / SIM_LEAGUE) : 'N/A';

    var analysis = stackAnalysis[s.qb] || '';

    sheet.getRange(r, 1, 1, 10).setValues([[
      s.qb, qbCust, s.rec, rcCust, total, qbAdp, rcAdp, stackRd, s.rating, analysis
    ]]);

    // Rating cell color
    var ratBg = s.rating === 'A+' ? '#14532d' :
                s.rating === 'A'  ? '#15803d' :
                s.rating === 'B+' ? '#0369a1' :
                s.rating === 'B'  ? '#6b21a8' : '#92400e';
    sheet.getRange(r, 9).setBackground(ratBg).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(r, 5).setBackground(s.color).setFontColor('#ffffff').setFontWeight('bold');

    var rowBg = st % 2 === 0 ? '#f8fafc' : '#ffffff';
    sheet.getRange(r, 1).setBackground(rowBg);
    sheet.getRange(r, 2).setBackground(rowBg);
    sheet.getRange(r, 3).setBackground(rowBg);
    sheet.getRange(r, 4).setBackground(rowBg);
    sheet.getRange(r, 6).setBackground(rowBg);
    sheet.getRange(r, 7).setBackground(rowBg);
    sheet.getRange(r, 8).setBackground(rowBg);
    sheet.getRange(r, 10).setBackground(rowBg).setWrap(true);
    r++;
  }

  // Nota final
  r += 2;
  sheet.getRange(r, 1, 1, 10).merge()
    .setValue('⚠️ NOTA: Los Custom Points son proyecciones basadas en FantasyPros ADP mayo 2026. Los stacks se evalúan por correlación de scoring, no por puntos individuales. Un stack A+ = correlación alta + schedule favorable + jugadores en posición de producir. Actualizar proyecciones antes del draft.')
    .setBackground('#fffbeb').setFontColor('#92400e')
    .setFontStyle('italic').setWrap(true);

  return r + 2;
}
