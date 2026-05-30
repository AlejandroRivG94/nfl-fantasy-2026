/**
 * NFL Fantasy 2026 — Coaching Hub
 * ================================
 * Expansión de OC Changes: cubre HC + OC + DC
 *
 * CÓMO LEER ESTA TABLA PARA FANTASY:
 *   HC nuevo (ofensivo)  → filosofía pass-heavy → QB/WR ganan, RBs pierden
 *   HC nuevo (defensivo) → filosofía conservadora → RBs ganan, menos puntos por juego
 *   OC nuevo             → sistema de pase cambia → distribución de targets incierta
 *   DC nuevo             → si mejora la defensa → más leads → más rushing → RBs ganan
 *
 * COLUMNAS:
 *   Team | Coach_Role | Previous_Coach | New_Coach_2026 | Change_Type
 *   Offensive_Style | Affected_Positions | Fantasy_Impact | Año_En_Rol | Notes | Source
 *
 * Corre: createCoachingHub_2026()
 * Después de correrlo, puedes eliminar el tab "OC Changes" (ya queda reemplazado)
 */

function createCoachingHub_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Coaching Hub';
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clear();

  var r = 1;

  // ── HEADER ─────────────────────────────────────────────────────────────────
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue('🏈  NFL FANTASY 2026 — COACHING HUB  |  HC + OC + DC')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  sheet.getRange(r, 1, 1, 11).merge()
    .setValue(
      'HC = filosofía ofensiva del equipo  |  OC = sistema de pase/targets  |  ' +
      'DC = game script (leads → más rush)  |  Año 1 = riesgo alto  |  Año 2+ = mayor predictibilidad'
    )
    .setBackground('#1a1f36').setFontColor('#9ca3af')
    .setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;

  // ── COLUMN HEADERS ─────────────────────────────────────────────────────────
  var headers = [
    'Team', 'Role', 'Previous', 'Nuevo 2026', 'Tipo',
    'Estilo Ofensivo', 'Posiciones Afectadas', 'Impacto Fantasy', 'Año en Rol', 'Notas', 'Fuente'
  ];
  sheet.getRange(r, 1, 1, 11).setValues([headers])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  // ─────────────────────────────────────────────────────────────────────────
  // DATA — HC CHANGES
  // ─────────────────────────────────────────────────────────────────────────
  var hcLabel = r;
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue('── HEAD COACHES (HC) — Determinan la filosofía ofensiva completa del equipo ──')
    .setBackground('#7c2d12').setFontColor('#fed7aa').setFontWeight('bold').setFontStyle('italic');
  r++;

  // [Team, Role, Previous, New, ChangeType, OffensiveStyle, Positions, Impact, YearInRole, Notes, Source]
  var hcData = [
    ['NYJ', 'HC', 'Robert Saleh (fired)', 'Aaron Glenn',
     'External hire — 2025', 'Equilibrado — Glenn es DC de origen, contrata OC (Frank Reich) para el pase',
     'QB, RB, WR', '↔️ NEUTRAL',
     'Año 2 — 2026',
     'Glenn viene de DET defensiva élite. Frank Reich como OC maneja el passing game. ' +
     'Breece Hall y Garrett Wilson dependen más de Reich que de Glenn. Sistema más disciplinado en 2026.',
     'NFL.com / ESPN'],

    ['PIT', 'HC', 'Mike Tomlin (continuity)', 'Mike Tomlin',
     'Continuity', 'Run-first históricamente — adaptando a Rodgers en Año 2',
     'QB, RB', '↔️ NEUTRAL',
     'Año 19 — veterano',
     'Tomlin sigue. Con Aaron Rodgers en Año 2 bajo Angelichio OC, PIT busca más balance pass/run. ' +
     'Históricamente run-heavy (Najee/Warren) pero Rodgers puede abrir el passing game.',
     'NFL.com'],

    ['NE', 'HC', 'Jerod Mayo (fired)', 'Mike Vrabel',
     'External hire — 2025', 'Equilibrado — Vrabel HC de origen defensivo pero abre offense con Van Pelt OC',
     'QB, RB, WR', '🔼 UPGRADE',
     'Año 2 — 2026',
     'Vrabel llegó de TEN con historia de hacer funcionar offenses sólidas. ' +
     'Drake Maye Año 3 bajo Vrabel + Van Pelt OC = mayor libertad. ' +
     'Stevenson y Elliott en backfield. NE en proceso de reconstrucción acelerada.',
     'ESPN / PFR'],

    ['WAS', 'HC', 'Ron Rivera (fired)', 'Dan Quinn',
     'External hire — 2025', 'Pass-heavy — Quinn (exDAL) valora el pase, Kingsbury OC air-raid',
     'QB, WR', '🔼 UPGRADE',
     'Año 2 — 2026',
     'Quinn + Kingsbury = combinación pass-heavy. Daniels Año 3 = mayor fluidez. ' +
     'McLaurin como WR1 en sistema air-raid. Rachaad White en rol de receiving back. ' +
     'Riesgo: WAS schedule playoff Brutal (Wk15-17).',
     'FantasyPros / Rotowire'],

    ['CHI', 'HC', 'Matt Eberflus (fired)', 'Ben Johnson',
     'External hire — 2025', 'Pass-heavy ÉLITE — Johnson es el OC más innovador de la última década (DET)',
     'QB, WR, TE', '🔼 UPGRADE MÁXIMO',
     'Año 2 — 2026',
     'Ben Johnson fue el arquitecto de la ofensiva DET (Goff + ARSB + Gibbs). Llegó a CHI con Waldron como OC. ' +
     'Caleb Williams Año 3 en sistema Johnson = breakout candidato. ' +
     'Keenan Allen + Rome Odunze + Cole Kmet = armas sólidas. OL pobre es el único riesgo.',
     'ESPN / NFL.com'],

    ['ATL', 'HC', 'Arthur Smith (fired)', 'Raheem Morris',
     'External hire — 2025', 'Equilibrado — Morris ex-DC que valoriza el running game con Bijan',
     'RB, WR, QB', '↔️ NEUTRAL',
     'Año 2 — 2026',
     'Morris viene de LAR defensiva. Bijan Robinson RB1 es el ancla independiente del HC. ' +
     'Tua Tagovailoa nuevo QB — adaptación al sistema de Robinson (OC) es la variable. ' +
     'Drake London + Darnell Mooney como WRs.',
     'Fantasy Footballers'],

    ['MIA', 'HC', 'Mike McDaniel (fired)', 'Unknown/TBD',
     'Vacante o nuevo hire', 'Incierto — McDaniel era pass-heavy innovador (Tua + Hill)',
     'QB, WR, RB', '⚠️ RIESGO ALTO',
     'Año 1 — nueva era',
     'McDaniel se fue con Tua a ATL (indirectamente). Hill es FA. ' +
     "De'Von Achane y Jaylen Waddle sobreviven cualquier sistema por talento. " +
     'Nuevo HC en Año 1 = máxima incertidumbre de sistema. Evitar stacks MIA.',
     'Rotowire'],

    ['CAR', 'HC', 'Frank Reich (fired)', 'Dave Canales',
     'External hire — 2025', 'Equilibrado — Canales ex-OC SEA/TB, bueno con QBs jóvenes',
     'QB, WR', '🔼 UPGRADE vs anterior',
     'Año 2 — 2026',
     'Canales tiene historial de desarrollar QBs jóvenes. Bryce Young Año 4 o nuevo QB. ' +
     'Tetairoa McMillan WR1 (clase 2025, Año 2) como receptor primario. ' +
     'CAR OL peor del NFL = riesgo estructural independiente del HC.',
     'ESPN'],

    ['DEN', 'HC', 'Sean Payton (continuity)', 'Sean Payton',
     'Continuity', 'Equilibrado — Payton históricamente pass-heavy pero adapta al roster',
     'QB, WR', '↔️ NEUTRAL',
     'Año 3 — establecido',
     'Payton sigue con Bo Nix Año 3. Payton tiene historial de construir offenses de élite. ' +
     'Javonte Williams y RJ Harvey en backfield. ' +
     'Courtland Sutton + Jaylen Waddle (llegó de MIA) como WRs.',
     'NFL.com'],

    ['JAX', 'HC', 'Doug Pederson (fired)', 'Liam Coen',
     'External hire — 2025', 'Pass-heavy — Coen ex-OC TB bajo Bowles, sistema de aire',
     'QB, WR, TE', '↔️ NEUTRAL',
     'Año 2 — 2026',
     'Coen con Kellen Moore OC (Año 2). Lawrence + Brian Thomas Jr. + Travis Hunter. ' +
     'Evan Engram como TE seguro. Doble cambio de staff es el riesgo de Año 1 → Año 2.',
     'Fantasy Footballers / ESPN'],

    ['CLE', 'HC', 'Kevin Stefanski (fired)', 'Tommy Rees',
     'HC/OC combinado (Rees asume ambos roles)', 'Incierto — Rees es muy joven para HC y OC simultáneo',
     'QB, WR, TE', '🔽 DOWNGRADE',
     'Año 2 — riesgo alto',
     'CLE apostó por Rees como head coach y coordinador ofensivo. QB situation caótica. ' +
     'Shedeur Sanders posible QB titular. Muy poco historial exitoso para confiar. ' +
     'Evitar jugadores CLE en drafts salvo riesgo calculado.',
     'ESPN / Rotowire'],
  ];

  // Write HC data
  sheet.getRange(r, 1, hcData.length, 11).setValues(hcData);
  _applyCoachingFormat(sheet, r, hcData, '#7c2d12', '#fef3c7');
  r += hcData.length + 1;

  // ─────────────────────────────────────────────────────────────────────────
  // DATA — OC CHANGES (actualizado y expandido)
  // ─────────────────────────────────────────────────────────────────────────
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue('── OFFENSIVE COORDINATORS (OC) — Determinan sistema de pase, targets y uso de RBs ──')
    .setBackground('#14532d').setFontColor('#bbf7d0').setFontWeight('bold').setFontStyle('italic');
  r++;

  var ocData = [
    ['NYJ', 'OC', 'Todd Downing (fired)', 'Frank Reich',
     'External hire — feb 2026', 'West Coast / Play-action — sistema de pase estructurado',
     'QB, WR, RB', '🔼 UPGRADE',
     'Año 1 — 2026',
     'Reich ex-HC IND/CAR. Sistema más disciplinado con Aaron Rodgers QB. ' +
     'Garrett Wilson sobrevive cualquier OC. Breece Hall mejora en recepción con Reich. ' +
     'NYJ OL sigue siendo el riesgo principal.',
     'CBS Sports / NFL.com — feb 2026'],

    ['PIT', 'OC', 'Matt Canada / anterior', 'Brian Angelichio',
     'External hire — 2026 offseason', 'Run-balanced con libertad para Rodgers en pase',
     'QB, WR, RB', '↔️ NEUTRAL/UPGRADE',
     'Año 1 — 2026',
     'Angelichio de MIN (staff O\'Connell). Rodgers Año 2 con Angelichio = más libertad. ' +
     'George Pickens como WR1. DK Metcalf (exSEA) como WR2 nuevo. ' +
     'Jaylen Warren lead back RB. Run-first pero con Rodgers passing game abre.',
     'CBS Sports / NFL.com'],

    ['WAS', 'OC', 'Eric Bieniemy (fired)', 'Kliff Kingsbury',
     'External hire 2025 → Año 2 en 2026', 'Air raid — motions rápidos, slants, routes cortas',
     'QB, WR', '🔼 UPGRADE',
     'Año 2 — mayor fluidez',
     'Año 2 de Kingsbury = sistema más fluido y químico establecida con Daniels. ' +
     'McLaurin + Dotson como receptores primarios. White (exTB) en receiving back role. ' +
     'WAS schedule playoffs Brutal = riesgo final de temporada.',
     'FantasyPros / Rotowire'],

    ['CHI', 'OC', 'Luke Getsy (fired)', 'Shane Waldron',
     'External hire 2025 → Año 2 en 2026', 'West Coast / RPO — sistema de timing preciso',
     'QB, WR, TE', '🔼 UPGRADE',
     'Año 2 — mayor apertura',
     'Waldron Año 2 bajo Ben Johnson HC = mayor apertura del playbook. ' +
     'Williams Año 3 con sistema establecido. Keenan Allen + Odunze como dupla. ' +
     'Cole Kmet safety blanket. CHI OL pobre es el riesgo único del sistema.',
     'Fantasy Footballers'],

    ['ARI', 'OC', 'Drew Petzing (resigned)', 'Ryan Grubb',
     'External hire 2025 → Año 2 en 2026', 'Motions + RPO — sistema de UW/SEA moderno',
     'QB, WR, TE', '🔼 UPGRADE',
     'Año 2 — mayor explosividad',
     'Año 2 de Grubb. Kyler Murray se fue a MIN; nuevo QB ARI. ' +
     'Marvin Harrison Jr. es el target monster del sistema. ' +
     'Tyler Allgeier nuevo RB. Grubb maximiza WR1 con rutas inteligentes.',
     'ESPN / Rotowire'],

    ['LV', 'OC', 'Mick Lombardi (fired)', 'Luke Getsy',
     'External hire 2025 → Año 2 en 2026', 'Pase moderno centrado en TE — Bowers como eje',
     'QB, TE, WR', '🔼 UPGRADE',
     'Año 2 — Bowers establece récords',
     'Getsy Año 2 con Mendoza QB Año 2 = química real. ' +
     'Bowers target hog de élite (110+ rec proyectadas). ' +
     'Jakobi Meyers en slot como WR2 confiable.',
     'Rotowire / ESPN'],

    ['TEN', 'OC', 'Tim Kelly (resigned)', 'Nick Holz',
     'External hire 2025 → Año 2 en 2026', 'Air raid adaptado de GB — pases rápidos y slants',
     'QB, WR', '🔼 UPGRADE',
     'Año 2 — Cam Ward explosivo',
     'Holz Año 2 con Ward Año 2 = sinfonía de sistemas. ' +
     'Calvin Ridley puede explotar como WR1 en air raid. ' +
     'Tyjae Spears lead back tras la salida de Pollard.',
     'Fantasy Footballers / ESPN'],

    ['ATL', 'OC', 'Dave Ragone (promoted)', 'Zac Robinson',
     'Promoted 2025 → Año 2 en 2026', 'Equilibrado — Robinson adapta al talento disponible',
     'QB, WR, RB', '↔️ NEUTRAL',
     'Año 2 — Tua variable crítica',
     'Robinson Año 2. Tua nuevo QB (exMIA) = adaptar sistema. ' +
     'Bijan Robinson RB1 independiente del OC. ' +
     'Drake London + Darnell Mooney como dupla WR.',
     'Fantasy Footballers'],

    ['NE', 'OC', "Bill O'Brien (fired)", 'Alex Van Pelt',
     'External hire 2025 → Año 2 en 2026', 'Más moderno que O\'Brien — mejor para Maye',
     'QB, RB, WR', '🔼 UPGRADE',
     'Año 2 — Maye liberado',
     'Van Pelt Año 2 bajo Vrabel HC nuevo = triple cambio estabilizándose. ' +
     'Maye Año 3 con sistema establecido = breakout candidate. ' +
     'Stevenson + Elliott en backfield. JaLynn Polk WR joven.',
     'ESPN / PFR'],

    ['JAX', 'OC', 'Press Taylor (fired)', 'Kellen Moore',
     'External hire 2025 → Año 2 en 2026', 'Spread / RPO — bueno en papel, inconsistente en práctica',
     'QB, WR, TE', '↔️ NEUTRAL',
     'Año 2 — esperando consistencia',
     'Moore Año 2. Históricamente inconsistente en DAL/PHI. ' +
     'Lawrence + Brian Thomas Jr. + Travis Hunter. ' +
     'Engram TE confiable independiente del OC.',
     'Fantasy Footballers'],

    ['CLE', 'OC', 'Ken Dorsey (fired)', 'Tommy Rees',
     'External hire 2025 → Año 2 en 2026', 'Desconocido — Rees muy joven para el rol',
     'QB, WR, TE', '🔽 DOWNGRADE',
     'Año 2 — aún con riesgo',
     'Rees doblando como HC. QB situation caótica limita todo. ' +
     'Njoku, Cooper, Tillman dependen de QB funcional. ' +
     'Evitar CLE a menos de que haya QB sorpresa.',
     'ESPN / Rotowire'],

    ['SF', 'OC', 'Kyle Shanahan (HC + OC)', 'Kyle Shanahan (continuity)',
     'Continuity', 'Sistema Shanahan élite — YBC #1, run-pass balance perfecta',
     'QB, RB, WR', '↔️ NEUTRAL (status quo élite)',
     'Año 12+ — predecible',
     'Shanahan sigue siendo el cerebro ofensivo. CMC + Purdy + Mike Evans = sistema élite. ' +
     'Continuidad total = alta predictibilidad = fácil proyectar.',
     'NFL.com'],

    ['KC', 'OC', 'Eric Bieniemy (gone)', 'Matt Nagy',
     'External hire — 2025 offseason', 'West Coast de KC — Nagy conoce el sistema Reid',
     'QB, WR, RB', '↔️ NEUTRAL',
     'Año 2 — Mahomes domina',
     'Mahomes supera cualquier OC. Kenneth Walker III nuevo RB. ' +
     'Rashee Rice regresó de lesión. ' +
     'KC schedule playoffs Brutal = único riesgo.',
     'ESPN'],

    ['BAL', 'OC', 'Greg Roman (prev)', 'Todd Monken (continuity)',
     'Continuity', 'Expandiendo passing game — más RPO y dropback para Lamar',
     'QB, WR, TE', '↔️ NEUTRAL (sistema maduro)',
     'Año 3+ — muy establecido',
     'Monken continuó expandiendo el passing game de BAL. ' +
     'Lamar + Flowers WR1 + Henry RB + Andrews/Likely TEs. ' +
     'Sistema predecible y de alta ceiling.',
     'NFL.com'],

    ['BUF', 'OC', 'Ken Dorsey → Joe Brady', 'Joe Brady (continuity)',
     'Continuity', 'Moderno — maximiza movilidad de Allen y pase rápido',
     'QB, WR, RB', '🔼 UPGRADE (Brady Año 2+)',
     'Año 2+ — establecido',
     'Brady implementó sistema para Allen. DJ Moore WR1 nuevo (exCHI). ' +
     'Shakir en slot, Kincaid TE safety blanket. ' +
     'BUF ofensiva más completa con DJ Moore.',
     'ESPN / FantasyPros'],

    ['DAL', 'OC', 'Brian Schottenheimer', 'Brian Schottenheimer (continuity)',
     'Continuity', 'Equilibrado — CeeDee Lamb como ancla del passing game',
     'QB, WR, RB', '↔️ NEUTRAL',
     'Continuity',
     'Schottenheimer sigue. CeeDee Lamb WR1 sólido. ' +
     'George Pickens nuevo WR2 (exPIT). Javonte Williams nuevo RB. ' +
     'QB situation = variable principal de riesgo DAL.',
     'NFL.com / ESPN'],
  ];

  sheet.getRange(r, 1, ocData.length, 11).setValues(ocData);
  _applyCoachingFormat(sheet, r, ocData, '#14532d', '#f0fdf4');
  r += ocData.length + 1;

  // ─────────────────────────────────────────────────────────────────────────
  // DATA — DC CHANGES (impacto en game script y DST)
  // ─────────────────────────────────────────────────────────────────────────
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue(
      '── DEFENSIVE COORDINATORS (DC) — Afectan game script (leads → más rush) y valor de DST ──'
    )
    .setBackground('#1e3a5f').setFontColor('#bfdbfe').setFontWeight('bold').setFontStyle('italic');
  r++;

  var dcData = [
    ['PHI', 'DC', 'Sean Desai (fired)', 'Vic Fangio',
     'External hire — 2025', 'Cover-2 / Zone agresivo — una de las mejores mentes defensivas',
     'DST, RB (game script)', '🔼 UPGRADE DÉFENSA',
     'Año 2 — establecido',
     'Fangio con PHI = defensiva que podría ser top-3 NFL. ' +
     'Con Barkley RB1 + OL élite + Fangio defense = PHI juega con leads constantemente. ' +
     'Fantasy: Barkley acumula carries en cuarto trimestre. DST PHI = top pick.',
     'ESPN / PFR'],

    ['BAL', 'DC', 'Mike Macdonald (hired por SEA)', 'Zach Orr',
     'Promoted — 2025', 'Zona agresiva — continuidad de cultura BAL defensiva',
     'DST, RB (game script)', '↔️ NEUTRAL',
     'Año 2 — Orr aprendió de Macdonald',
     'BAL defensiva sigue siendo élite bajo Orr. Lamar + Henry con leads = más rushing. ' +
     'DST BAL = top-5 pick. ' +
     'Henry beneficiado por carries en cuarto con ventajas cómodas.',
     'NFL.com'],

    ['KC', 'DC', 'Steve Spagnuolo (continuity)', 'Steve Spagnuolo',
     'Continuity', 'Cover-1 Man / Blitz selectivo — disruption pura',
     'DST', '↔️ NEUTRAL',
     'Año 8+ — dominante',
     'Spagnuolo sigue siendo top-3 DC del NFL. ' +
     'KC D permite suficientes puntos que Mahomes siempre tiene que pasar en cuartos tardíos. ' +
     'Schedule Brutal (Wk15-17) limita el upside de KC DST en playoffs.',
     'NFL.com'],

    ['DET', 'DC', 'Aaron Glenn (fue a NYJ como HC)', 'Kelvin Sheppard',
     'Internal promotion — 2025', 'Zona / Cover-2 — Sheppard aprendió del sistema Glenn',
     'DST', '⚠️ DOWNGRADE temporario',
     'Año 2 — adaptando',
     'Glenn construyó la defensiva DET de scratch. Sheppard tiene menos crédito probado. ' +
     'DET puede perder algo de su edge defensiva en 2026. ' +
     'Gibbs y ARSB no se ven muy afectados (ofensiva DET es independiente de la D).',
     'ESPN'],

    ['SF', 'DC', 'Steve Wilks / Nick Sorensen', 'Nick Sorensen (promoted)',
     'Internal promotion', 'Shanahan-influenced D — aggressive press-man',
     'DST, RB (game script)', '↔️ NEUTRAL',
     'Año 2 — adaptando',
     'SF defense sigue siendo élite por talento (Nick Bosa). ' +
     'CMC y backfield SF beneficiados cuando SF juega con ventajas. ' +
     'DST SF = top-5 pick cuando Bosa está sano.',
     'NFL.com'],

    ['MIN', 'DC', 'Brian Flores (continuity)', 'Brian Flores',
     'Continuity', 'Man-heavy / Blitz pesado — uno de los mejores DCs del NFL',
     'DST', '↔️ NEUTRAL',
     'Año 3+ — élite',
     'Flores sigue. MIN defense + Kyler Murray nuevo QB con movilidad. ' +
     'Aaron Jones en backfield = safe receiving back. ' +
     'DST MIN = top-10 pick con schedule playoff Elite (Wk15-17).',
     'NFL.com'],

    ['NYG', 'DC', 'Wink Martindale (resigned)', 'Shane Bowen',
     'External hire — 2025', 'Agresivo / Blitz — similar a Martindale en filosofía',
     'DST', '↔️ NEUTRAL',
     'Año 2 — adaptando',
     'NYG defense en reconstrucción. Jaxson Dart QB en sistema de cuestionable OL. ' +
     'Nabers WR1 absorbe contacto y genera YAC independiente de la D. ' +
     'DST NYG = evitar con schedule Brutal en playoffs.',
     'Rotowire'],
  ];

  sheet.getRange(r, 1, dcData.length, 11).setValues(dcData);
  _applyCoachingFormat(sheet, r, dcData, '#1e3a5f', '#eff6ff');
  r += dcData.length;

  // ── NOTA FINAL ──────────────────────────────────────────────────────────────
  r += 2;
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue(
      '⚠️ NOTA: Cambios de HC y DC son estimados con base en CBS Sports / NFL.com / ESPN offseason 2026. ' +
      'Verificar status antes del draft en cbssports.com/nfl/teams. ' +
      'DC impacta game script: défensa élite = más leads = más carries para RB del equipo. ' +
      'DC débil = más passing = más valor a WR/TE del equipo.'
    )
    .setBackground('#fffbeb').setFontColor('#92400e').setFontStyle('italic').setWrap(true);

  // Column widths
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 380);
  sheet.setColumnWidth(11, 180);
  sheet.setFrozenRows(4);

  SpreadsheetApp.getUi().alert(
    '✅ Coaching Hub creado.\n\n' +
    '  HC: ' + hcData.length + ' equipos\n' +
    '  OC: ' + ocData.length + ' equipos\n' +
    '  DC: ' + dcData.length + ' equipos\n\n' +
    'Puedes eliminar el tab "OC Changes" — queda reemplazado por Coaching Hub.\n\n' +
    '⚠️ Verificar cambios de HC/DC antes del draft en cbssports.com/nfl/teams'
  );
}

// ─── HELPER — aplica formato por filas alternadas y colores de impacto ─────────
function _applyCoachingFormat(sheet, startRow, data, headerColor, altColor) {
  var impactColors = {
    '🔼 UPGRADE':           '#c8e6c9',
    '🔼 UPGRADE MÁXIMO':    '#a5d6a7',
    '🔼 UPGRADE DÉFENSA':   '#b2dfdb',
    '🔼 UPGRADE vs anterior': '#c8e6c9',
    '↔️ NEUTRAL':           '#fff9c4',
    '↔️ NEUTRAL (status quo élite)': '#fff9c4',
    '↔️ NEUTRAL (sistema maduro)':   '#fff9c4',
    '⚠️ RIESGO ALTO':       '#ffe0b2',
    '⚠️ DOWNGRADE temporario': '#ffe0b2',
    '🔽 DOWNGRADE':         '#ffcdd2',
  };

  for (var i = 0; i < data.length; i++) {
    var bg     = i % 2 === 0 ? altColor : '#ffffff';
    var impact = data[i][7];
    var impBg  = impactColors[impact] || bg;

    sheet.getRange(startRow + i, 1, 1, 11).setBackground(bg);
    sheet.getRange(startRow + i, 8).setBackground(impBg).setFontWeight('bold');

    // Role cell color
    var role  = data[i][1];
    var roleBg = role === 'HC' ? '#7c2d12' :
                 role === 'OC' ? '#14532d' : '#1e3a5f';
    sheet.getRange(startRow + i, 2)
      .setBackground(roleBg).setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
}
