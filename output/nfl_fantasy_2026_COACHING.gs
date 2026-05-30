/**
 * NFL Fantasy 2026 — Coaching Hub
 * ================================
 * 32 equipos × HC + OC + DC verificados
 * Fuente: tabla oficial de coaching staff 2026
 *
 * Corre: createCoachingHub_2026()
 */

function createCoachingHub_2026() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Coaching Hub') || ss.insertSheet('Coaching Hub');
  sheet.clear();

  var r = 1;

  // ── HEADER ──────────────────────────────────────────────────────────────────
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue('🏈  NFL FANTASY 2026 — COACHING HUB  |  32 equipos · HC + OC + DC verificados')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  r++;
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue(
      'HC ofensivo = más puntos/juego = QBs/WRs suben  |  ' +
      'HC defensivo = más control = RBs ganan en leads tardíos  |  ' +
      'DC élite = defensiva fuerte = leads = más carreras = RB value  |  ' +
      'Año 1 nuevo OC = alto riesgo distribución targets'
    )
    .setBackground('#1a1f36').setFontColor('#9ca3af').setHorizontalAlignment('center').setFontStyle('italic');
  r += 2;

  // ── LEYENDA DE ESTILOS ───────────────────────────────────────────────────────
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue(
      'ESTILO OC:  RUN-HEAVY = más carries RB  |  PASS-HEAVY = más targets WR/TE  |  ' +
      'BALANCED = equilibrado  |  INNOVADOR = motions/RPO/pre-snap  |  DESCONOCIDO = Año 1 sin historial NFL'
    )
    .setBackground('#374151').setFontColor('#d1d5db').setFontStyle('italic');
  r += 2;

  // ── HEADERS DE TABLA ────────────────────────────────────────────────────────
  var cols = [
    'Team', 'HC', 'OC', 'DC',
    'HC Perfil', 'OC Estilo Ofensivo',
    'Fantasy Impact', 'Riesgo', 'Jugadores Afectados', 'Nota Clave', 'Division'
  ];
  sheet.getRange(r, 1, 1, 11).setValues([cols])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  // ─────────────────────────────────────────────────────────────────────────────
  // 32 EQUIPOS — datos verificados mayo 2026
  // Columnas: Team | HC | OC | DC | HC_Perfil | OC_Estilo | Fantasy_Impact | Riesgo | Jugadores | Nota | Division
  // ─────────────────────────────────────────────────────────────────────────────
  var data = [

    // ── NFC EAST ──────────────────────────────────────────────────────────────
    ['DAL', 'Brian Schottenheimer', 'Klayton Adams', 'Christian Parker',
     'Run-balanced (ex-DAL OC)',
     'Desconocido — Adams sin historial coordinador',
     '⚠️ INCIERTO',
     'Alto — nuevo OC sin track record',
     'CeeDee Lamb, Javonte Williams',
     'Schottenheimer es run-first. Adams como OC nuevo = incertidumbre en targets. Lamb es tan bueno que sobrevive, pero stack con QB incierto.',
     'NFC East'],

    ['NYG', 'John Harbaugh', 'Matt Nagy', 'Dennard Wilson',
     '🔼 OFENSIVO — Harbaugh construyó la ofensiva BAL de élite',
     'PASS-HEAVY — Nagy es discípulo de Reid, conoce el spread',
     '🔼 UPGRADE MÁXIMO',
     'Bajo — Harbaugh + Nagy = combinación probada',
     'Jaxson Dart, Malik Nabers',
     'SORPRESA DEL DRAFT: Harbaugh deja BAL y llega a NYG con Nabers. Dart Año 2 + Harbaugh + Nagy = Stack que puede explotar. Nagy (exCHI HC, exKC OC) conoce la lectura de defensa.',
     'NFC East'],

    ['PHI', 'Nick Sirianni', 'Sean Mannion', 'Vic Fangio',
     'BALANCEADO — Sirianni ha construido offenses sólidas con PHI',
     'PASS-HEAVY — Mannion ex-QB turned OC, pase rápido',
     '🔼 UPGRADE (Fangio defense)',
     'Bajo — sistema establecido',
     'Saquon Barkley, A.J. Brown, Jalen Hurts',
     'Fangio como DC = defensiva top-3 NFL = PHI juega con LEADS = Barkley acumula carries en Q4. Brown + Hurts con Mannion OC nuevo. Ofensiva sólida.',
     'NFC East'],

    ['WAS', 'Dan Quinn', 'David Blough', 'Daronte Jones',
     'BALANCEADO — Quinn ex-DC, abre offense con buenos coordinadores',
     'DESCONOCIDO — Blough ex-QB backup, Año 1 como OC',
     '⚠️ INCIERTO',
     'Alto — OC sin historial como coordinador',
     'Jayden Daniels, Terry McLaurin',
     'Quinn (exDAL DC) sabe construir culturas ganadoras. RIESGO: Blough como OC es la gran incógnita. Ex-QB convirtiendo en OC Año 1 con Daniels = ¿cuánta libertad le da? McLaurin sobrevive cualquier sistema.',
     'NFC East'],

    // ── NFC NORTH ─────────────────────────────────────────────────────────────
    ['CHI', 'Ben Johnson', 'Press Taylor', 'Dennis Allen',
     '🔼 OFENSIVO ÉLITE — Johnson arquitecto de la mejor ofensiva DET',
     'PASS-HEAVY / INNOVADOR — Taylor ex-IND OC + BAL, entiende a QBs jóvenes',
     '🔼 UPGRADE MÁXIMO',
     'Bajo — mejor combinación HC/OC para un QB joven en el draft',
     'Caleb Williams, Rome Odunze, Keenan Allen',
     'MEJOR ECOSISTEMA para Caleb Williams Año 3: Johnson (el mejor diseñador de offense del NFL) + Taylor (desarrolló Luck en IND). OL pobre sigue siendo el único riesgo. Williams puede ser QB de élite este año.',
     'NFC North'],

    ['DET', 'Dan Campbell', 'Drew Petzing', 'Kelvin Sheppard',
     'RUN-HEAVY cultura — Campbell identidad de poder físico',
     'BALANCED — Petzing conoce el sistema DET',
     '↔️ NEUTRAL',
     'Bajo — continuidad y sistema establecido',
     'Jahmyr Gibbs, Amon-Ra St. Brown, Jared Goff',
     'Gibbs es RB1 claro con Petzing OC que entiende su uso dual. ARSB en slot élite. Goff protegido. DET es el equipo más predecible para proyectar en fantasy.',
     'NFC North'],

    ['GB', 'Matt LaFleur', 'Adam Stenavich', 'Jonathan Gannon',
     'BALANCEADO West Coast — LaFleur moldea a QBs jóvenes',
     'BALANCED — Stenavich conoce el sistema GB internamente',
     '↔️ NEUTRAL',
     'Bajo — LaFleur continuity',
     'Jordan Love, Davante Adams, Josh Jacobs',
     'LaFleur sigue. Stenavich como OC interno = menor disrupción. GB schedule playoff Elite. Jacobs RB1 sólido. Adams-Love dupla probada. Equipo muy proyectable.',
     'NFC North'],

    ['MIN', 'Kevin O\'Connell', 'Wes Phillips', 'Brian Flores',
     '🔼 OFENSIVO MODERNO — O\'Connell ex-LAR OC, spread élite',
     'PASS-HEAVY / INNOVADOR — Phillips ex-LAR TE coach turned OC',
     '↔️ NEUTRAL',
     'Bajo — O\'Connell continuity',
     'Justin Jefferson, Kyler Murray, Aaron Jones',
     'O\'Connell sigue con Kyler Murray nuevo QB. Jefferson + Murray = stack a considerar. Jones como receiving back complementa bien. Flores DC = defensiva agresiva = leads = más Jones carries en Q4.',
     'NFC North'],

    // ── NFC SOUTH ─────────────────────────────────────────────────────────────
    ['ATL', 'Kevin Stefanski', 'Tommy Rees', 'Jeff Ulbrich',
     'RUN-BALANCED — Stefanski construyó el juego terrestre en CLE con Chubb',
     'DESCONOCIDO — Rees ex-Notre Dame OC, Año 1 en NFL como coordinador',
     '⚠️ INCIERTO',
     'Medio — Rees sin historial NFL como OC',
     'Bijan Robinson, Drake London, Tua Tagovailoa',
     'Bijan Robinson va a estar BIEN independientemente — Stefanski ama el run game. El riesgo es Tua + Rees OC nuevo en Año 1. London como WR1 sobrevive. Bijan es el jugador más seguro de ATL.',
     'NFC South'],

    ['CAR', 'Dave Canales', 'Brad Idzik', 'Ejiro Evero',
     'QB-DEVELOPMENT — Canales desarrolló a Geno Smith y Baker Mayfield',
     'DESCONOCIDO — Idzik sin historial relevante como OC',
     '⚠️ INCIERTO',
     'Alto — CAR OL peor NFL + OC nuevo + QB joven',
     'Tetairoa McMillan, Bryce Young',
     'McMillan WR1 tiene talento real. El riesgo es el sistema entero: Idzik OC nuevo, OL pobre, QB joven. Canales sabe desarrollar QBs pero necesita protección para que el sistema funcione.',
     'NFC South'],

    ['NO', 'Kellen Moore', 'Doug Nussmeier', 'Brandon Staley',
     '🔼 OFENSIVO — Moore ex-OC PHI/JAX/DAL, sistemas de pase creativos',
     'PASS-HEAVY — Nussmeier ex-OC en college/NFL, sistema aéreo',
     '🔼 UPGRADE',
     'Medio — Moore Año 1 como HC',
     'Chris Olave, Tony Pollard, Travis Etienne',
     'Moore llega de JAX como HC con historial de construir pass-heavy offenses. Olave WR1 puede explotar. Pollard + Etienne backfield de 2 cabezas. Nussmeier OC complementa la visión de Moore.',
     'NFC South'],

    ['TB', 'Todd Bowles', 'Zac Robinson', 'N/A',
     'DEFENSIVO — Bowles HC de origen DC, conservador en offense',
     'PASS-BALANCED — Robinson ex-ATL OC, conoce el pass game moderno',
     '↔️ NEUTRAL',
     'Bajo — Robinson Año 1 en TB',
     'Baker Mayfield, Chris Godwin, Rachaad White',
     'Bowles conservador = White y Godwin con piso sólido pero ceiling limitado. Robinson OC puede abrir la ofensiva. Bucky Irving tiene upside como receiving back.',
     'NFC South'],

    // ── NFC WEST ──────────────────────────────────────────────────────────────
    ['ARI', 'Mike LaFleur', 'Nathaniel Hackett', 'Nick Rallis',
     'BALANCEADO — Mike LaFleur ex-NYJ/SF OC, conoce el West Coast',
     'PASS-HEAVY — Hackett ex-DEN/JAX/CLE OC, offenses de aire',
     '🔼 UPGRADE',
     'Medio — nuevo sistema completo en ARI',
     'Marvin Harrison Jr., Cam Skattebo, Jeremiyah Love',
     'Harrison Jr. WR1 de talento generacional FINALMENTE con un OC competente (Hackett). Skattebo y Love como backfield dual. LaFleur + Hackett = más pase que antes para ARI.',
     'NFC West'],

    ['SF', 'Kyle Shanahan', 'Klay Kubiak', 'Raheem Morris',
     'SISTEMA ÉLITE — Shanahan el mejor scheming HC del NFL',
     'BALANCED / SISTEMA — Kubiak entiende el sistema Shanahan internamente',
     '↔️ NEUTRAL',
     'Bajo — Shanahan continuity',
     'Christian McCaffrey, Brock Purdy, Mike Evans',
     'Shanahan sigue. Kubiak (Gary Kubiak\'s son) como OC interno = menor disrupción al sistema. CMC si sano = RB1 overall. Morris DC continuidad defensiva. SF muy predecible.',
     'NFC West'],

    ['LAR', 'Sean McVay', 'Nate Scheelhaase', 'Chris Shula',
     '🔼 OFENSIVO — McVay revolucionó el football moderno con el spread',
     'PASS-HEAVY — Scheelhaase entiende el sistema McVay',
     '↔️ NEUTRAL',
     'Bajo — McVay continuity',
     'Puka Nacua, Kyren Williams, Cooper Kupp',
     'McVay sigue. Nacua WR1 con target share enorme. Williams RB1 sólido. Kupp como veteran presence en red zone. Schedule Neutral = no ventajas ni desventajas en playoffs.',
     'NFC West'],

    ['SEA', 'Mike Macdonald', 'Brian Fleury', 'Aden Durde',
     'DEFENSIVO — Macdonald ex-BAL DC, defense-first culture',
     'DESCONOCIDO — Fleury Año 1 como coordinador',
     '⚠️ INCIERTO',
     'Alto — Fleury sin historial + Macdonald Año 2',
     'Jaxson Smith-Njigba, DK Metcalf, Sam Darnold',
     'JSN WR1 con target share creciente. Metcalf rojo deep threat. Darnold QB serviceable. El riesgo es Fleury OC nuevo + sistema defensivo de Macdonald = offense puede ser muy conservadora.',
     'NFC West'],

    // ── AFC EAST ──────────────────────────────────────────────────────────────
    ['BUF', 'Joe Brady', 'Pete Carmichael', 'Jim Leonhard',
     'OFENSIVO — Brady ex-BUF OC, conoce el sistema de Allen',
     'PASS-HEAVY ÉLITE — Carmichael OC de NO por 15+ años bajo Payton/Brees',
     '🔼 UPGRADE',
     'Medio — Brady Año 1 como HC',
     'Josh Allen, DJ Moore, James Cook',
     'Carmichael es uno de los OCs más subestimados del NFL — construyó la ofensiva Brees/NO durante 15 años. Con Allen = ceiling altísimo. Brady sabe el sistema Allen. Schedule playoff Elite para BUF.',
     'AFC East'],

    ['MIA', 'Jeff Hafley', 'Bobby Slowick', 'Sean Duggan',
     'DESCONOCIDO — Hafley ex-BC college HC, background defensivo',
     'DESCONOCIDO — Slowick ex-SF/SEA analítico, primer OC de equipo titular',
     '⚠️ RIESGO ALTO',
     'Alto — HC y OC nuevos + Tua se fue + Hill FA',
     "De'Von Achane, Jaylen Waddle",
     'MIA es el equipo más incierto del draft. Hafley y Slowick son dos coordinadores sin historial probado como líderes. Achane tiene talento puro y sobrevive cualquier sistema por velocidad. Waddle también. Evitar QB de MIA.',
     'AFC East'],

    ['NE', 'Mike Vrabel', 'Josh McDaniels', 'Zak Kuhr',
     'BALANCEADO — Vrabel ex-HC TEN, entiende QBs jóvenes',
     'PASS-HEAVY — McDaniels construyó la ofensiva Brady/Belichick en NE',
     '🔼 UPGRADE',
     'Bajo — Vrabel + McDaniels = historial NE probado',
     'Drake Maye, Rhamondre Stevenson, JaLynn Polk',
     'Maye Año 3 con McDaniels (que conoce NE como nadie) + Vrabel = la dupla perfecta para su desarrollo. McDaniels sabe cómo maximizar QBs en el sistema Patriot. Maye breakout candidate real.',
     'AFC East'],

    ['NYJ', 'Aaron Glenn', 'Frank Reich', 'Brian Duker',
     'DEFENSIVO — Glenn ex-DET DC de élite',
     'PASS-BALANCED — Reich ex-HC IND/CAR, sistema disciplinado',
     '↔️ NEUTRAL',
     'Bajo — Glenn + Reich = combinación sólida',
     'Breece Hall, Garrett Wilson',
     'Glenn + Reich es la combinación correcta: Glenn maneja cultura y defensa, Reich maneja la ofensiva. Wilson sobrevive cualquier OC por talento. Hall mejora en recepción con sistema Reich. OL NYJ sigue siendo el mayor riesgo.',
     'AFC East'],

    // ── AFC NORTH ─────────────────────────────────────────────────────────────
    ['BAL', 'Jesse Minter', 'Declan Doyle', 'Anthony Weaver',
     'DESCONOCIDO — Minter ex-DB coach/joven, Año 1 como HC',
     'DESCONOCIDO — Doyle Año 1 como OC, sin historial relevante',
     '⚠️ RIESGO ALTO',
     'Alto — HC y OC completamente nuevos con Lamar Jackson',
     'Lamar Jackson, Zay Flowers, Derrick Henry',
     'MAYOR SORPRESA del offseason: Harbaugh se fue y BAL apostó por Minter (inexperience). Lamar Jackson con HC y OC nuevos en Año 1 = MUCHO riesgo de regresión en proyecciones. FP proyecta conservador. Flowers y Henry tienen pisos altos independientemente.',
     'AFC North'],

    ['CIN', 'Zac Taylor', 'Dan Pitcher', 'Al Golden',
     'BALANCEADO — Taylor entiende el sistema Burrow/Chase',
     'PASS-HEAVY — Pitcher conoce el sistema CIN internamente',
     '↔️ NEUTRAL',
     'Bajo — Taylor continuity con Chase+Burrow',
     "Joe Burrow, Ja'Marr Chase, Tee Higgins",
     'Taylor + Pitcher (interno) = menor disrupción al sistema Chase-Burrow. CIN es el stack de WR más confiable del draft. Burrow+Chase Año continuity.',
     'AFC North'],

    ['CLE', 'Todd Monken', 'Travis Switzer', 'Mike Rutenberg',
     'OFENSIVO — Monken ex-BAL OC, construyó el passing game de Lamar',
     'DESCONOCIDO — Switzer sin historial relevante como coordinador',
     '⚠️ INCIERTO',
     'Alto — QB situation + OC nuevo en CLE',
     'Jerry Jeudy, Harold Fannin Jr.',
     'Monken es buena contratación como HC pero Switzer OC es incógnita. QB situation caótica sigue siendo el mayor riesgo. Jeudy y Fannin Jr. dependen del QB. Evitar CLE salvo riesgo calculado.',
     'AFC North'],

    ['PIT', 'Mike McCarthy', 'Brian Angelichio', 'Patrick Graham',
     'OFENSIVO — McCarthy construyó offenses élite en GB y DAL',
     'BALANCED / RUN — Angelichio ex-MIN staff, moderno',
     '🔼 UPGRADE',
     'Bajo — McCarthy tiene historial enorme',
     'George Pickens, DK Metcalf, Aaron Rodgers',
     'McCarthy + Angelichio = sistema más sofisticado que Tomlin para Rodgers. McCarthy sabe maximizar QBs veteranos (Favre, Romo, Rodgers antes). Pickens WR1 + Metcalf WR2 nuevo = dupla élite. Graham DC sólido.',
     'AFC North'],

    // ── AFC SOUTH ─────────────────────────────────────────────────────────────
    ['HOU', 'DeMeco Ryans', 'Nick Caley', 'Matt Burke',
     'DEFENSIVO — Ryans ex-DC, culture-builder, confía en el run',
     'BALANCED — Caley ex-NE TE/OC coach, conoce los sistemas modernos',
     '↔️ NEUTRAL',
     'Bajo — Ryans continuity',
     'CJ Stroud, Nico Collins, David Montgomery',
     'Ryans + Caley = sistema establecido. Stroud QB sólido con Collins + Tank Dell + Johnson como WRs. Montgomery nuevo RB. HOU tiene buen balance.',
     'AFC South'],

    ['IND', 'Shane Steichen', 'Jim Bob Cooter', 'Lou Anarumo',
     'OFENSIVO — Steichen ex-PHI OC, pass-heavy con QBs móviles',
     'BALANCED — Cooter ex-DET/JAX OC, conoce a QBs jóvenes',
     '↔️ NEUTRAL',
     'Bajo — sistema establecido',
     'Anthony Richardson, Jonathan Taylor, Alec Pierce',
     'Richardson + Steichen + Cooter = sistema diseñado para QBs duales. Taylor RB1 predecible. Pierce WR1 extendido. Anarumo DC sólido (ex-CIN DC). IND es estable.',
     'AFC South'],

    ['JAX', 'Liam Coen', 'Grant Udinski', 'Anthony Campanile',
     'OFENSIVO — Coen ex-TB/LAR OC, pass-friendly',
     'DESCONOCIDO — Udinski Año 1 como coordinador principal',
     '⚠️ INCIERTO',
     'Medio — Coen Año 2 como HC pero OC nuevo',
     'Trevor Lawrence, Brian Thomas Jr., Travis Hunter',
     'Coen Año 2 con Udinski OC nuevo = mayor certeza que Año 1 pero aún incierto. BTJ WR1 con talento real. Hunter dual-threat WR/CB Año 2 de crecimiento. Lawrence en su mejor situación desde Pederson.',
     'AFC South'],

    ['TEN', 'Robert Saleh', 'Brian Daboll', 'Gus Bradley',
     'DEFENSIVO — Saleh ex-NYJ HC, culture de defensa y disciplina',
     '🔼 PASS-HEAVY CREATIVO — Daboll ex-BUF OC/NYG HC, maximiza QBs móviles',
     '🔼 UPGRADE',
     'Medio — Saleh + Daboll Año 1 juntos',
     'Cam Ward, Calvin Ridley, Tyjae Spears',
     'COMBINACIÓN SUBESTIMADA: Daboll fue el arquitecto del sistema que hizo a Josh Allen élite en BUF. Cam Ward = QB muy similar a Allen en movilidad. Ward + Daboll = stack a considerar en R5-R7. Ridley WR1 beneficiado.',
     'AFC South'],

    // ── AFC WEST ──────────────────────────────────────────────────────────────
    ['DEN', 'Sean Payton', 'Davis Webb', 'Vance Joseph',
     '🔼 OFENSIVO — Payton construyó la dynasty NO con Brees',
     'DESCONOCIDO — Webb ex-QB backup turned OC, Año 1 en rol',
     '⚠️ INCIERTO',
     'Medio — Webb OC sin historial + Bo Nix Año 3',
     'Bo Nix, Courtland Sutton, Javonte Williams',
     'Payton sabe offense pero Webb OC nuevo es el riesgo. Similar a Blough en WAS. Bo Nix Año 3 necesita un OC que lo libere. Sutton WR1 sólido. Williams RB1 si Webb entiende el run game.',
     'AFC West'],

    ['KC', 'Andy Reid', 'Eric Bieniemy', 'Steve Spagnuolo',
     '🔼 GENIUS OFENSIVO — Reid mejor HC de todos los tiempos en offense',
     'PASS-HEAVY ÉLITE — Bieniemy de vuelta a KC, conoce a Mahomes perfecto',
     '🔼 UPGRADE (Bieniemy regresa)',
     'Bajo — Reid + Bieniemy + Mahomes = dynasty',
     'Patrick Mahomes, Rashee Rice, Xavier Worthy',
     'BIENIEMY REGRESA A KC — esto es enorme. Bieniemy fue el arquitecto del sistema que dio 4 Super Bowls. Con Mahomes Año whatever = ceiling limitado solo por el schedule Brutal. KC DST élite con Spagnuolo.',
     'AFC West'],

    ['LAC', 'Jim Harbaugh', 'Mike McDaniel', 'Chris O\'Leary',
     'RUN-BALANCED — Harbaugh ama el power run + play-action',
     '🔼 PASS-HEAVY INNOVADOR — McDaniel ex-MIA HC, sistema de aire creativo',
     '🔼 UPGRADE',
     'Bajo — McDaniel + Harbaugh = duo élite',
     'Justin Herbert, Omarion Hampton, Ladd McConkey',
     'McDaniel como OC bajo Harbaugh es una combo poderosa: Harbaugh da la cultura física, McDaniel diseña la offense. Herbert puede explotar con el sistema más creativo de su carrera. McConkey WR1 slot. Hampton RB emergente.',
     'AFC West'],

    ['LV', 'Klint Kubiak', 'Andrew Janocko', 'Rob Leonard',
     'DESCONOCIDO — Klint Kubiak (hijo de Gary), run/play-action en el ADN',
     'DESCONOCIDO — Janocko sin historial como OC principal',
     '⚠️ INCIERTO',
     'Alto — HC y OC completamente nuevos + Mendoza QB',
     'Brock Bowers, Jakobi Meyers',
     'Bowers TE1 sobrevive cualquier sistema — es tan bueno que cualquier OC lo usa como ancla. Kubiak puede traer ADN de run/play-action de su padre Gary (famoso por el sistema de boot). Janocko OC es la gran incógnita.',
     'AFC West'],
  ];

  // ── WRITE DATA ───────────────────────────────────────────────────────────────
  sheet.getRange(r, 1, data.length, 11).setValues(data);

  // Colores por impacto y por division
  var impactColors = {
    '🔼 UPGRADE MÁXIMO':         '#166534',
    '🔼 UPGRADE':                '#15803d',
    '🔼 UPGRADE (Bieniemy regresa)': '#15803d',
    '🔼 UPGRADE (Fangio defense)': '#15803d',
    '↔️ NEUTRAL':                '#a16207',
    '⚠️ INCIERTO':               '#c2410c',
    '⚠️ RIESGO ALTO':            '#b91c1c',
  };
  var impactTextColors = {
    '🔼 UPGRADE MÁXIMO':         '#ffffff',
    '🔼 UPGRADE':                '#ffffff',
    '🔼 UPGRADE (Bieniemy regresa)': '#ffffff',
    '🔼 UPGRADE (Fangio defense)': '#ffffff',
    '↔️ NEUTRAL':                '#ffffff',
    '⚠️ INCIERTO':               '#ffffff',
    '⚠️ RIESGO ALTO':            '#ffffff',
  };

  var divColors = {
    'NFC East':  '#dbeafe', 'NFC North': '#dcfce7',
    'NFC South': '#fef9c3', 'NFC West':  '#fce7f3',
    'AFC East':  '#e0e7ff', 'AFC North': '#fee2e2',
    'AFC South': '#fff7ed', 'AFC West':  '#f5f3ff',
  };

  for (var i = 0; i < data.length; i++) {
    var row    = r + i;
    var impact = data[i][6];
    var div    = data[i][10];
    var divBg  = divColors[div] || '#ffffff';
    var altBg  = (i % 2 === 0) ? divBg : '#ffffff';

    sheet.getRange(row, 1, 1, 11).setBackground(altBg);

    // Impact column
    var impBg   = impactColors[impact]     || '#ffffff';
    var impText = impactTextColors[impact] || '#000000';
    sheet.getRange(row, 7).setBackground(impBg).setFontColor(impText).setFontWeight('bold');

    // Division column subtle color
    sheet.getRange(row, 11).setBackground(divBg).setFontColor('#374151');
  }

  // ── TOP CHANGES SECTION ──────────────────────────────────────────────────────
  var summaryRow = r + data.length + 2;

  sheet.getRange(summaryRow, 1, 1, 11).merge()
    .setValue('🏆  TOP CAMBIOS CON MAYOR IMPACTO FANTASY 2026')
    .setBackground('#0d1117').setFontColor('#f5c518')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  summaryRow++;

  var topChanges = [
    ['🥇', 'CHI — Ben Johnson HC + Press Taylor OC',
     'Caleb Williams Año 3 en el mejor sistema posible. El OC más innovador del NFL como HC. Williams breakout = clave del draft.'],
    ['🥈', 'NYG — John Harbaugh HC + Matt Nagy OC',
     'Harbaugh construyó offenses élite en BAL. Nagy conoce el spread. Jaxson Dart + Nabers = stack sorpresa del draft.'],
    ['🥉', 'NE — Mike Vrabel HC + Josh McDaniels OC',
     'McDaniels conoce el sistema NE como nadie. Maye Año 3 + McDaniels = breakout candidate real. Mayor upgrade QB2 del draft.'],
    ['4', 'TEN — Robert Saleh HC + Brian Daboll OC',
     'Daboll hizo a Josh Allen élite en BUF. Cam Ward tiene perfil MUY similar. Ward + Daboll = el stack más infravalorado del draft.'],
    ['5', 'BUF — Joe Brady HC + Pete Carmichael OC',
     'Carmichael OC de NO por 15 años con Brees/Payton. Uno de los OCs más efectivos de la historia. Con Allen = ceiling enorme.'],
    ['6', 'LAC — Jim Harbaugh HC + Mike McDaniel OC',
     'McDaniel bajó de HC a OC pero trae su sistema innovador. Herbert + McDaniel = potencialmente la mejor QB-OC pairing de su carrera.'],
    ['7', 'KC — Andy Reid + Eric Bieniemy (regresa)',
     'Bieniemy construyó el sistema que ganó 4 Super Bowls. De vuelta a KC con Mahomes. Schedule Brutal es el único limitante.'],
    ['⚠️', 'BAL — Jesse Minter HC + Declan Doyle OC',
     'Harbaugh se fue. Minter y Doyle = Año 1, sin historial probado. Lamar Jackson con HC/OC nuevos = proyecciones FP son conservadoras (344 pts). Riesgo real.'],
    ['⚠️', 'WAS — Dan Quinn HC + David Blough OC',
     'Blough ex-backup QB como OC = altísima incertidumbre en el sistema de Daniels. Quinn es bueno pero el OC importa para proyectar WRs.'],
    ['⚠️', 'MIA — Jeff Hafley HC + Bobby Slowick OC',
     'Dos coordinadores sin historial como titulares. Achane y Waddle sobreviven por talento pero evitar cualquier stack MIA.'],
  ];

  var topHeaders = ['#', 'Equipo / Cambio', 'Por qué importa para fantasy'];
  sheet.getRange(summaryRow, 1, 1, 3).setValues([topHeaders])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  summaryRow++;

  for (var t = 0; t < topChanges.length; t++) {
    var bg = t < 7 ? '#f0fdf4' : '#fef2f2';
    sheet.getRange(summaryRow + t, 1, 1, 3).setValues([topChanges[t]]).setBackground(bg);
    sheet.getRange(summaryRow + t, 3).setWrap(true);
  }

  // Column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 165);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 160);
  sheet.setColumnWidth(10, 400);
  sheet.setColumnWidth(11, 80);
  sheet.setFrozenRows(5);

  SpreadsheetApp.getUi().alert(
    '✅ Coaching Hub creado con datos verificados 2026.\n\n' +
    '  32 equipos × HC + OC + DC\n' +
    '  Top 10 cambios más impactantes\n\n' +
    'HALLAZGOS CLAVE:\n' +
    '  🔼 CHI Ben Johnson + Press Taylor = Caleb Williams breakout\n' +
    '  🔼 NYG John Harbaugh + Matt Nagy = Dart/Nabers stack sorpresa\n' +
    '  🔼 TEN Brian Daboll OC = Cam Ward como Josh Allen 2.0\n' +
    '  🔼 KC Bieniemy regresa = Mahomes sigue siendo élite\n' +
    '  ⚠️ BAL Jesse Minter HC = Lamar con sistema nuevo = riesgo\n\n' +
    'Puedes eliminar el tab "OC Changes" — queda reemplazado.'
  );
}
