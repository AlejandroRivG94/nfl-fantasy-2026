/**
 * NFL Fantasy 2026 — VOR (Value Over Replacement) + Draft Strategy
 * ================================================================
 * CONCEPTO:
 *   VOR = Custom Points del jugador - Custom Points del "replacement"
 *   Replacement = el último jugador que CUALQUIER equipo puede agarrar
 *   gratis en waiver wire después del draft.
 *
 *   Liga: 12 equipos | 1QB 2RB 2WR 1TE 2FLEX 1D 1K | 14 rondas
 *
 * REPLACEMENT LEVELS (12 equipos, estimado):
 *   QB:  QB #13 en Custom  → ~Jordan Love / Baker Mayfield   ~308 pts
 *   RB:  RB #37 en Custom  → última RB de flex en liga 12-eq ~190 pts
 *        (12×2 titulares + 12×1 flex mínimo = 36 RBs)
 *   WR:  WR #37 en Custom  → última WR de flex               ~165 pts
 *        (12×2 titulares + 12×1 flex = 36 WRs)
 *   TE:  TE #13 en Custom  → ~Dallas Goedert / Tucker Kraft   ~158 pts
 *
 * CÓMO LEER EL BIG BOARD:
 *   • VOR alto = jugador produce MUCHO más que lo gratuito
 *   • Draft_Round = ronda donde su VOR justifica agarrarlo
 *   • Si ADP_Round < Draft_Round = lo estás agarrando ANTES de lo necesario
 *   • Si ADP_Round > Draft_Round = VALOR — lo puedes agarrar TARDE y está infravalorado
 *
 * FUNCIÓN PRINCIPAL: addVOR_toBigBoard()
 */

// ─────────────────────────────────────────────────────────────────────────────
// PASO 1: Agrega columnas P, Q, R al Big Board con VOR, Draft Round, Value
// ─────────────────────────────────────────────────────────────────────────────
function addVOR_toBigBoard() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Big Board');

  // Agregar headers en P16, Q17, R18 (columnas 16, 17, 18)
  var vorHeaders = ['VOR_Custom', 'Draft_Round_Ideal', 'Value_vs_ADP'];
  sheet.getRange(1, 16, 1, 3).setValues([vorHeaders]);

  // Estilo headers
  sheet.getRange(1, 16).setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(1, 17).setBackground('#1a7b47').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(1, 18).setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');

  sheet.setColumnWidth(16, 110);
  sheet.setColumnWidth(17, 160);
  sheet.setColumnWidth(18, 140);

  // Replacement levels (Custom Points)
  // QB13 ≈ 308 | RB37 ≈ 190 | WR37 ≈ 165 | TE13 ≈ 158
  var repQB = 308;
  var repRB = 190;
  var repWR = 165;
  var repTE = 158;

  // Para filas 2-200: calcular VOR, Round Ideal, Value vs ADP
  // Col B = Position (col 2) | Col E = Custom Points (col 5) | Col I = ADP (col 9)
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (var r = 2; r <= lastRow; r++) {
    // VOR = Custom Points - Replacement Level según posición
    // =IF(B2="QB", E2-308, IF(B2="RB", E2-190, IF(B2="WR", E2-165, IF(B2="TE", E2-158, ""))))
    sheet.getRange(r, 16).setFormula(
      '=IF(B'+r+'="QB",IFERROR(E'+r+'-'+repQB+',0),' +
      'IF(B'+r+'="RB",IFERROR(E'+r+'-'+repRB+',0),' +
      'IF(B'+r+'="WR",IFERROR(E'+r+'-'+repWR+',0),' +
      'IF(B'+r+'="TE",IFERROR(E'+r+'-'+repTE+',0),""))))'
    );

    // Draft Round Ideal = la ronda donde su VOR justifica la selección
    // Basado en: VOR > 150 = R1 | > 100 = R2 | > 80 = R3 | > 60 = R4
    //            > 45 = R5 | > 30 = R6 | > 20 = R7 | > 10 = R8 | ≤ 10 = R9+
    sheet.getRange(r, 17).setFormula(
      '=IF(P'+r+'="","",'+
      'IF(P'+r+'>150,"Ronda 1 — Élite",'+
      'IF(P'+r+'>100,"Ronda 2 — Top Pick",'+
      'IF(P'+r+'>80,"Ronda 3 — Sólido",'+
      'IF(P'+r+'>60,"Ronda 4 — Buen valor",'+
      'IF(P'+r+'>45,"Ronda 5 — Flex sólido",'+
      'IF(P'+r+'>30,"Rondas 6-7",'+
      'IF(P'+r+'>20,"Rondas 7-9",'+
      'IF(P'+r+'>10,"Rondas 10-11","Ronda 12+"")'
      + '))))))))'
    );

    // Value vs ADP:
    // ADP_Round = CEILING(I / 12, 1) — en qué ronda cae su ADP
    // Draft_Round_Num = número de ronda ideal (extraído de P16 formula)
    // Diferencia = ADP_Round - Ideal_Round
    // + = puedes agarrarlo MÁS TARDE de lo ideal (valor)
    // - = debes agarrarlo ANTES de lo ideal (costo)
    sheet.getRange(r, 18).setFormula(
      '=IF(OR(P'+r+'="",I'+r+'=""),"",'+
      'LET(adp_rnd, CEILING(I'+r+'/12,1),'+
      'vor_rnd,'+
      'IF(P'+r+'>150,1,IF(P'+r+'>100,2,IF(P'+r+'>80,3,IF(P'+r+'>60,4,'+
      'IF(P'+r+'>45,5,IF(P'+r+'>30,6,IF(P'+r+'>20,8,IF(P'+r+'>10,10,12)))))))),'+
      'IF(adp_rnd-vor_rnd>1,"✅ VALOR +R"&(adp_rnd-vor_rnd),'+
      'IF(adp_rnd-vor_rnd<-1,"⚠️ COSTO -R"&ABS(adp_rnd-vor_rnd),'+
      '"— Precio justo"))))'
    );
  }

  // Formato numérico VOR
  sheet.getRange(2, 16, lastRow - 1, 1).setNumberFormat('0.0');

  // Colores condicionales para Value_vs_ADP (col R = 18)
  var valueRange = sheet.getRange(2, 18, lastRow - 1, 1);
  var valorRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('✅')
    .setBackground('#1a7b47').setFontColor('#ffffff')
    .setRanges([valueRange]).build();
  var costoRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('⚠️')
    .setBackground('#d93025').setFontColor('#ffffff')
    .setRanges([valueRange]).build();
  var justoRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('—')
    .setBackground('#fbbc04').setFontColor('#000000')
    .setRanges([valueRange]).build();
  var existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat([valorRule, costoRule, justoRule]));

  // Colores condicionales para VOR (col P = 16)
  var vorRange = sheet.getRange(2, 16, lastRow - 1, 1);
  var eliteVOR = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(100)
    .setBackground('#1a7b47').setFontColor('#ffffff')
    .setRanges([vorRange]).build();
  var goodVOR = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(50, 100)
    .setBackground('#34a853').setFontColor('#ffffff')
    .setRanges([vorRange]).build();
  var midVOR = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(20, 49)
    .setBackground('#fbbc04').setFontColor('#000000')
    .setRanges([vorRange]).build();
  var lowVOR = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(20)
    .setBackground('#fa7b17').setFontColor('#ffffff')
    .setRanges([vorRange]).build();
  sheet.setConditionalFormatRules(
    sheet.getConditionalFormatRules().concat([eliteVOR, goodVOR, midVOR, lowVOR])
  );

  SpreadsheetApp.getUi().alert(
    '✅ VOR añadido al Big Board.\n\n' +
    'Nuevas columnas:\n' +
    '  P = VOR_Custom   (pts sobre replacement)\n' +
    '  Q = Draft Round Ideal\n' +
    '  R = Value vs ADP ← LA MÁS IMPORTANTE\n\n' +
    '✅ VALOR = lo puedes agarrar más tarde\n' +
    '⚠️ COSTO = tienes que agarrarlo antes\n' +
    '— = precio justo en su ADP\n\n' +
    'Siguiente: corre createDraftStrategyTab()\n' +
    'para ver el análisis por escenario de draft.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PASO 2: Tab de estrategia de draft — responde "¿qué pasa si agarro X en ronda Z?"
// ─────────────────────────────────────────────────────────────────────────────
function createDraftStrategyTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear tab si no existe
  var sheet = ss.getSheetByName('Draft Strategy');
  if (!sheet) {
    sheet = ss.insertSheet('Draft Strategy');
    ss.moveActiveSheet(ss.getSheets().length);
  }
  sheet.clearContents();
  sheet.clearFormats();

  // ── SECCIÓN 1: Escasez posicional ─────────────────────────────────────────
  var row = 1;
  sheet.getRange(row, 1).setValue('📊 ESCASEZ POSICIONAL — Liga 12 equipos | 14 rondas | Custom .8PPR+.2PPC')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 1, 1, 8).setBackground('#1a1f36');
  row += 2;

  // Headers
  var scarcityHeaders = ['Posición','Reempl. #','Pts Reempl.','Top Jugador','Top Pts','VOR Top','Cliff en...','Estrategia'];
  sheet.getRange(row, 1, 1, 8).setValues([scarcityHeaders])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  row++;

  var scarcity = [
    ['QB','QB #13','~308 pts','Josh Allen BUF','~482 pts','~+174','Ronda 2-3','Espera: QB13 es jugable. Allen en R2 = robo. Lamar en R4 = valor enorme.'],
    ['RB','RB #37','~190 pts','Jahmyr Gibbs DET','~391 pts','~+201','Ronda 5-6','Elite RBs se van rápido. Gibbs/Barkley/CMC en R1. Etienne/Jacobs en R3-4.'],
    ['WR','WR #37','~165 pts','Ja\'Marr Chase CIN','~309 pts','~+144','Ronda 7','Más profundidad. JSN/ARSB/Jefferson en R1. Buena WR2 disponible R4-6.'],
    ['TE','TE #13','~158 pts','Trey McBride ARI','~229 pts','~+71','Ronda 4-5','CLIFF en TE3. McBride/Bowers = R2. Loveland/Warren = R4. Después cae mucho.'],
  ];
  sheet.getRange(row, 1, scarcity.length, 8).setValues(scarcity);
  for (var r = row; r < row + scarcity.length; r++) {
    sheet.getRange(r, 1, 1, 8).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
  }
  row += scarcity.length + 2;

  // ── SECCIÓN 2: Análisis de rondas — qué hay disponible en cada ronda ───────
  sheet.getRange(row, 1).setValue('🎯 QUÉ HAY DISPONIBLE POR RONDA (estimado por ADP)')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 1, 1, 6).setBackground('#1a1f36');
  row += 2;

  var roundHeaders = ['Ronda','ADP picks','Mejores QBs disponibles','Mejores RBs','Mejores WRs','Mejores TEs'];
  sheet.getRange(row, 1, 1, 6).setValues([roundHeaders])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  row++;

  var rounds = [
    ['R1', '1-12',  '—',
     'Bijan, Gibbs, CMC, J.Taylor, J.Cook, Jeanty, Achane, Barkley, O.Hampton, K.Walker, C.Brown, D.Henry',
     'Chase, Nacua, JSN, ARSB, Jefferson, Lamb',
     'McBride, Bowers'],

    ['R2', '13-24', 'Josh Allen (ADP 21!) ← ROBO',
     'Jeremiyah Love, Breece Hall, Travis Etienne, Javonte Williams',
     'Drake London, Nico Collins, George Pickens, Nabers, AJ Brown, Rice',
     '—'],

    ['R3', '25-36', 'Lamar Jackson (ADP 43) ← Esperar aquí',
     'Josh Jacobs, Kyren Williams',
     'Tee Higgins, DeVonta Smith, Egbuka, McMillan, Zay Flowers, Skattebo RB',
     'Colston Loveland (CHI)'],

    ['R4', '37-48', 'Joe Burrow (ADP 49)',
     'Cam Skattebo (NYG), Bucky Irving, Quinshon Judkins, David Montgomery',
     'DJ Moore, Ladd McConkey, Luther Burden, Davante Adams, Evans',
     '—'],

    ['R5', '49-60', 'Drake Maye / Jayden Daniels',
     'TreVeyon Henderson, Jadarian Price, D\'Andre Swift',
     'Jameson Williams, Jaylen Waddle, Terry McLaurin',
     'Tyler Warren (IND) ← si no salió antes'],

    ['R6', '61-72', 'Caleb Williams / Jalen Hurts (ADP 66) ← ROBO',
     'Bhayshul Tuten, Chuba Hubbard',
     'Rome Odunze, Christian Watson, Marvin Harrison Jr., Brian Thomas Jr.',
     'Harold Fannin Jr., Tucker Kraft'],

    ['R7', '73-84', 'Trevor Lawrence / Justin Herbert',
     'RJ Harvey (DEN), Jaylen Warren (PIT), Tony Pollard',
     'Alec Pierce, Carnell Tate, Jordyn Tyson',
     'Sam LaPorta, Kyle Pitts'],

    ['R8', '85-96', 'Patrick Mahomes (ADP 89) ← valor aquí',
     'Kyle Monangai, Rico Dowdle, Blake Corum',
     'DK Metcalf (PIT), Courtland Sutton, Michael Wilson, Chris Godwin',
     'George Kittle (si disponible)'],

    ['R9', '97-108','Brock Purdy / Bo Nix / Jared Goff',
     'J.K. Dobbins, Jordan Mason',
     'Ricky Pearsall, Michael Pittman Jr., Quentin Johnston',
     'Dallas Goedert, Mark Andrews'],

    ['R10','109-120','Kyler Murray / Jordan Love / Baker Mayfield',
     'Rachaad White, Aaron Jones',
     'Khalil Shakir, Xavier Worthy, Matthew Golden',
     'Dalton Kincaid, Isaiah Likely'],

    ['R11','121-132','Sam Darnold / CJ Stroud',
     'Jacory Croskey-Merritt, Jonathon Brooks',
     'Jayden Higgins, Jalen Coker, Wan\'Dale Robinson',
     'T.J. Hockenson, Hunter Henry'],

    ['R12','133-144','Cam Ward / Tyler Shough',
     'Isiah Pacheco, Zach Charbonnet',
     'Jaxon Smith-Njigba backup depth',
     'Chig Okonkwo, Juwan Johnson'],

    ['R13','145-156','Aaron Rodgers / Fernando Mendoza',
     'Tyler Allgeier, Tyjae Spears, Alvin Kamara',
     'Tyreek Hill (FA!), Brandon Aiyuk, Calvin Ridley',
     '—'],

    ['R14','157-168','Geno Smith / Bryce Young',
     'Keaton Mitchell, Brian Robinson Jr.',
     'Tre Tucker, Antonio Williams, Jerry Jeudy',
     '— (TEs de valor ya no hay)'],
  ];
  sheet.getRange(row, 1, rounds.length, 6).setValues(rounds);
  for (var r2 = row; r2 < row + rounds.length; r2++) {
    sheet.getRange(r2, 1, 1, 6).setBackground(r2 % 2 === 0 ? '#f8f9fa' : '#ffffff');
  }
  row += rounds.length + 2;

  // ── SECCIÓN 3: Estrategias según posición de draft ──────────────────────────
  sheet.getRange(row, 1).setValue('🏆 ESTRATEGIAS RECOMENDADAS SEGÚN POSICIÓN EN EL DRAFT')
    .setBackground('#1a1f36').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 1, 1, 6).setBackground('#1a1f36');
  row += 2;

  var stratHeaders = ['Pick','Estrategia','R1','R2','R3','R4-6'];
  sheet.getRange(row, 1, 1, 6).setValues([stratHeaders])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  row++;

  var strategies = [
    ['Pick 1-3','RB Élite First',
     'Bijan / Gibbs / CMC',
     'Chase / Nacua / JSN (WR)',
     'Lamar / Allen (QB) si disponible',
     'RB2 sólido + McBride/Bowers TE'],

    ['Pick 4-6','RB + WR Élite',
     'J.Taylor / Jeanty / Barkley',
     'Chase / Jefferson / Lamb (WR)',
     'Allen / Lamar (QB)',
     'Travis Etienne + TE top'],

    ['Pick 7-9','WR Élite + RB',
     'JSN / Nacua / Jefferson / Lamb',
     'J.Cook / O.Hampton / K.Walker (RB)',
     'Joe Burrow / Hurts (QB)',
     'Breece Hall + TE'],

    ['Pick 10-12','WR Stack o QB Tardío',
     'Nabers / AJ Brown / Rashee Rice / Olave',
     'Jacobs / Williams (RB)',
     'Tee Higgins / DeVonta (WR)',
     'Allen R2 si disponible ← ROBO'],
  ];
  sheet.getRange(row, 1, strategies.length, 6).setValues(strategies);
  for (var r3 = row; r3 < row + strategies.length; r3++) {
    sheet.getRange(r3, 1, 1, 6).setBackground(r3 % 2 === 0 ? '#e8f5e9' : '#f1f8e9');
  }
  row += strategies.length + 2;

  // ── SECCIÓN 4: Valores por ADP vs Custom — los mayores discrepancias ─────────
  sheet.getRange(row, 1).setValue('⚡ MAYORES DISCREPANCIAS ADP vs CUSTOM POINTS — OPORTUNIDADES')
    .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 1, 1, 6).setBackground('#e65100');
  row += 2;

  var discHeaders = ['Jugador','Pos','ADP Round','Custom Rank','Diferencia','Por qué'];
  sheet.getRange(row, 1, 1, 6).setValues([discHeaders])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold');
  row++;

  var discrepancies = [
    ['Josh Allen',       'QB BUF','R2 (ADP 21)', 'Custom #1 QB',  '🔥 R1 value en R2',
     'ADP 21 pero ~482 Custom pts. Mejor QB del formato .2PPC. Cargas 130+ en tierra.'],
    ['Jalen Hurts',      'QB PHI','R6 (ADP 66)', 'Custom #3 QB',  '🔥 R2 value en R6',
     'ADP 66 es absurdo. 120 carries + 10 TDs terrestres. OL élite. Robalo en R6.'],
    ['Lamar Jackson',    'QB BAL','R4 (ADP 43)', 'Custom #2 QB',  '🔥 R2 value en R4',
     '140 carries + 950 yds = .2PPC explota. ADP 43 = gran valor en Custom format.'],
    ['Jayden Daniels',   'QB WAS','R5 (ADP 57)', 'Custom #4 QB',  '✅ R3 value en R5',
     'Año 3. Kingsbury air raid. 700 rushYds. Daniels como QB2 confiable con upside.'],
    ['Jahmyr Gibbs',     'RB DET','R1 (ADP 2)',  'Custom #5 RB',  '— Precio justo',
     'ADP correcto para su valor. Lead back solo + OL élite DET. Confirmed must-draft.'],
    ['Brock Bowers',     'TE LV', 'R2 (ADP 17)', 'Custom #2 TE',  '— Precio justo',
     'TE VOR altísimo. Si agarras TE en R2 = ventaja toda la temporada.'],
    ['Trey McBride',     'TE ARI','R2 (ADP 17)', 'Custom #1 TE',  '— Precio justo',
     'TE1 overall. ~229 Custom pts vs replacement ~158. VOR = +71. Vale la R2.'],
    ['Colston Loveland', 'TE CHI','R3 (ADP 36)', 'Custom #4 TE',  '✅ R5 value en R3',
     'Año 2 CHI. Williams Año 3. Williams-Loveland connection. TE upside enorme.'],
    ['Jaxson Smith-Njigba','WR SEA','R1 (ADP 5)','Custom #3 WR', '— Precio justo',
     'ADP 5 refleja su breakout. Metcalf se fue. JSN = WR1 SEA. Darnold lo alimenta.'],
    ['Breece Hall',      'RB NYJ','R3 (ADP 33)', 'Custom #8 RB',  '✅ Valor en R3',
     'Frank Reich nuevo OC. Si OL NYJ mejora = top-5 RB. ADP 33 = riesgo razonable.'],
    ['Patrick Mahomes',  'QB KC', 'R8 (ADP 89)', 'Custom #8 QB',  '✅ R5 value en R8',
     'ADP 89 para QB #1 de la última década. Custom ~376 pts. R8 es un robo.'],
    ['Ashton Jeanty',    'RB LV', 'R2 (ADP 13)', 'Custom #11 RB', '⚠️ R3 value en R2',
     'Año 2 con alto upside pero ADP 13 = R2 de 12 teams. Es caro. Monitorear OL LV.'],
    ['Omarion Hampton',  'RB LAC','R2 (ADP 17)', 'Custom #9 RB',  '✅ R3 value en R2',
     'Año 2. 205 carries proyectados. ADP 17 puede ser valor si Herbert lo usa bien.'],
  ];
  sheet.getRange(row, 1, discrepancies.length, 6).setValues(discrepancies);
  for (var r4 = row; r4 < row + discrepancies.length; r4++) {
    var bg = '#fff3e0';
    if (String(discrepancies[r4-row][4]).indexOf('🔥') !== -1) bg = '#e8f5e9';
    if (String(discrepancies[r4-row][4]).indexOf('⚠️') !== -1) bg = '#ffebee';
    sheet.getRange(r4, 1, 1, 6).setBackground(bg);
  }
  row += discrepancies.length + 2;

  // ── SECCIÓN 5: Anchos de columna ─────────────────────────────────────────
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 350);
  sheet.setColumnWidth(5, 350);
  sheet.setColumnWidth(6, 400);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    '✅ Tab "Draft Strategy" creado.\n\n' +
    'Contiene:\n' +
    '1. Escasez posicional por posición\n' +
    '2. Qué hay disponible en cada ronda (R1-R14)\n' +
    '3. Estrategias según tu pick en el draft\n' +
    '4. Top 13 discrepancias ADP vs Custom\n\n' +
    'Úsalo junto al Big Board para responder:\n' +
    '"¿Qué pasa si agarro X en ronda Z?"'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// CORRER TODO: VOR + Draft Strategy tab
// ─────────────────────────────────────────────────────────────────────────────
function addFullDraftAnalysis() {
  addVOR_toBigBoard();
  createDraftStrategyTab();
}
