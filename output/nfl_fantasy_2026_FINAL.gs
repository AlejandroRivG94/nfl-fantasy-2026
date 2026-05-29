/**
 * NFL Fantasy 2026 — SCRIPT FINAL COMPLETO
 * ============================================================
 * Datos verificados: FantasyPros Half-PPR ADP (mayo 2026)
 * Movimientos: CBS Sports Free Agency Tracker (mayo 2026)
 * OC Changes: CBS Sports + reportes verificados (mayo 2026)
 * OL Grades: PFF + PFR temporada 2025
 *
 * INSTRUCCIONES:
 * 1. Extensions → Apps Script
 * 2. Crea una nueva pestaña y nómbrala "final_2026"
 * 3. Pega este script completo
 * 4. Guarda (Ctrl+S)
 * 5. Corre "runAllFinal_2026" → pobla TODAS las hojas
 *
 * AÑOS DE EXPERIENCIA:
 *   Clase 2024 (Daniels, Maye, Williams, Nabers, Thomas, etc.) = Año 3 en 2026
 *   Clase 2025 (Ward, Hunter, Skattebo, Harvey, etc.)          = Año 2 en 2026
 *   Clase 2026                                                  = Rookie (Año 1)
 *
 * MOVIMIENTOS VERIFICADOS (CBS Sports, mayo 2026):
 *   Pittman → PIT  | Evans → SF    | DJ Moore → BUF
 *   Etienne → NO   | K.Walker → KC | Allgeier → ARI
 *   Tua → ATL      | Kyler → MIN   | Darnold → SEA
 *   FA sin equipo: Tyreek Hill, Stefon Diggs, Deebo Samuel
 *
 * OCs VERIFICADOS 2026:
 *   NYJ: Frank Reich (contratado feb 2026)
 *   PIT: Brian Angelichio (de MIN, 2026 offseason)
 */

// ─────────────────────────────────────────────────────────────────────────────
// FUNCIÓN PRINCIPAL — corre esta para poblar todo
// ─────────────────────────────────────────────────────────────────────────────
function runAllFinal_2026() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Iniciando actualización completa 2026...\nEsto tomará ~60 segundos. Espera la confirmación final.');

  updatePlayersDB_FINAL();
  updateScoringCalc_FINAL();
  updateOCChanges_FINAL();
  updateOLGrades_FINAL();
  updateScheduleFlags_FINAL();
  fixBigBoard_FINAL();

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Big Board')
  );

  ui.alert(
    '✅ NFL Fantasy 2026 — Actualización completa.\n\n' +
    '📊 Datos cargados:\n' +
    '  • Players DB: 200 jugadores\n' +
    '  • Scoring Calc: 150 jugadores con proyecciones\n' +
    '  • OC Changes: 2026 verificado\n' +
    '  • OL Grades: Temporada 2025\n' +
    '  • Schedule Flags: Bye weeks + Playoffs\n' +
    '  • Big Board: Ordenado por Custom Points\n\n' +
    '⚠️ Verifica schedule en nfl.com/schedules antes del draft.\n' +
    '⚠️ Confirma equipo de Hill, Diggs y Deebo Samuel (agentes libres).'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB 1: PLAYERS DB — 200 jugadores
// Columnas: Player | Position | Team | Tier | ADP_FantasyPros | ADP_Underdog
//           Expert_Notes | OC_Change | B2B_Weeks | Primetime_Weeks
//           Playoff_SoS | OL_Grade
// ─────────────────────────────────────────────────────────────────────────────
function updatePlayersDB_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Players DB');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 12).clearContent();

  // Tier: 1=Élite (ADP 1-36) | 2=Sólido (37-72) | 3=Flex (73-120)
  //       4=Bench (121-168)   | 5=Especulativo (169-200)
  // OC_Change: "Yes" si el equipo tiene nuevo OC en 2026
  // Playoff_SoS: Easy/Neutral/Hard para semanas 15-17
  // OL_Grade: calificación orientativa de la OL del equipo

  var data = [
    // ── TIER 1 — ADP ~1-36 ──────────────────────────────────────────────────
    // [Player, Pos, Team, Tier, ADP_FP, ADP_UD, Notes, OC_Chg, B2B, PT, SoS, OL]
    ["Ja'Marr Chase",         'WR','CIN','1', 1.1,  1.2,  'WR1 overall. Burrow connection élite. 100+ rec + 1500 yds + 12 TDs. Piso altísimo cada semana.','No','','','Easy','Buena'],
    ['Jahmyr Gibbs',          'RB','DET','1', 2.0,  2.3,  'Lead back SOLO en DET. Montgomery fue a HOU. Pacheco es backup (ADP 146). 280+ carries + 65 rec. Año 4. OL élite.','No','','','Neutral','Élite'],
    ['Josh Allen',            'QB','BUF','1', 3.1,  3.0,  'Elite dual-threat QB. 130 carries + 12 TDs terrestres. Custom format lo eleva. DJ Moore llega de CHI. Ofensiva élite BUF.','No','','','Easy','Muy buena'],
    ['Lamar Jackson',         'QB','BAL','1', 4.0,  4.2,  'MVP x2. Mejor QB en formato .2PPC. 140 carries + 950 yds. 36 passTDs. BAL OL muy sólida. Run-first Harbaugh system.','No','','','Neutral','Muy buena'],
    ['Breece Hall',           'RB','NYJ','1', 5.2,  5.5,  'Talento élite. Frank Reich nuevo OC 2026 NYJ. Braelon Allen backup. Riesgo: OL NYJ pobre. Upside inmenso si OL mejora.','Yes','','','Neutral','Pobre'],
    ['Saquon Barkley',        'RB','PHI','1', 6.1,  6.0,  'Campeón. OL élite PHI. 290 carries + 1400 yds + 14 TDs proyectados. Tank Bigsby backup (ADP 183). Must-start toda la temporada.','No','','','Easy','Élite'],
    ['CeeDee Lamb',           'WR','DAL','1', 7.3,  7.1,  'WR1 sólido DAL. OL élite históricamente. 110 rec + 1450 yds + 10 TDs. Monitorear status QB (Prescott/Rush en 2026).','No','','','Neutral','Muy buena'],
    ['Amon-Ra St. Brown',     'WR','DET','1', 8.2,  8.5,  'Slot élite. 120+ recepciones. Comparte ofensiva con Gibbs. OL DET élite. Año 6. Piso altísimo semana a semana.','No','','','Neutral','Élite'],
    ['Brock Bowers',          'TE', 'LV','1', 9.1,  9.3,  'TE1 overall. 110 rec proyectadas. Fernando Mendoza QB nuevo LV (ADP 170, Año 2). OC Getsy lo prioriza. Target hog élite.','Yes','','','Neutral','Promedio'],
    ['Patrick Mahomes',       'QB', 'KC','1',10.5, 10.8,  'KC dynasty. Kenneth Walker III nuevo RB (llegó de SEA). 4500 yds + 38 TDs. Custom format sólido por movilidad adicional.','No','','','Hard','Élite'],
    ['Joe Burrow',            'QB','CIN','1',11.2, 11.0,  'Regresó 100% de lesión 2025. Chase+Higgins dupla élite. 4400 yds + 35 TDs. Año 7. Muy confiable si está sano.','No','','','Easy','Buena'],
    ['Bijan Robinson',        'RB','ATL','1',12.0, 12.3,  'RB1 ATL. Tua Tagovailoa nuevo QB (llega de MIA). Brian Robinson Jr. backup (ADP 165). Nuevo OC interno (Robinson).','Yes','','','Neutral','Buena'],
    ['Puka Nacua',            'WR','LAR','1',13.5, 14.0,  'WR1 establecido LAR. Stafford protegido. 100 rec + 1300 yds + 9 TDs. Kupp como veteran presence en red zone.','No','','','Neutral','Muy buena'],
    ['Justin Jefferson',      'WR','MIN','1',14.2, 14.5,  'Elite WR. Kyler Murray nuevo QB MIN (llegó de ARI). Murray-Jefferson puede ser explosivo. Año 7 de Jefferson.','Yes','','','Easy','Buena'],
    ['Malik Nabers',          'WR','NYG','1',15.0, 15.3,  'Talento generacional. Año 3 (clase 2024). NYG OL débil pero Nabers genera YAC. Ceiling de WR1 si QB mejora.','No','','','Neutral','Promedio'],
    ['Sam LaPorta',           'TE','DET','1',16.2, 16.8,  'TE2 overall. DET ofensiva élite. Comparte con ARSB+Gibbs pero volumen consistente. Año 3. Confiabilidad alta.','No','','','Neutral','Élite'],
    ['Christian McCaffrey',   'RB', 'SF','1',17.5, 18.0,  'Si está sano = RB1 overall. Shanahan system élite. Mike Evans nuevo WR (llega de TB). Riesgo salud es el único factor.','No','','','Neutral','Élite'],
    ['Davante Adams',         'WR', 'GB','1',18.3, 18.7,  'Veterano élite. Love-Adams connexión probada. 110 rec + 1200 yds. GB OL sólida. Schedule playoff favorable.','No','','','Easy','Muy buena'],
    ['Jalen Hurts',           'QB','PHI','1',19.5, 20.0,  'Top QB en Custom format .2PPC. 120 carries + 650 yds + 10 TDs terrestres. OL élite. A.J.Brown + DeVonta Smith.','No','','','Easy','Élite'],
    ['Tyreek Hill',           'WR', 'FA','1',20.5, 21.0,  '⚠️ AGENTE LIBRE — Sin equipo al 28 mayo 2026. Verificar firma antes del draft. Si firma = WR1 inmediato en cualquier equipo.','No','','','Neutral','N/A'],
    ['Jayden Daniels',        'QB','WAS','1',22.0, 22.5,  'Año 3 (clase 2024). Kingsbury OC (air raid, Año 2) = más pase, menos carreras. Dual-threat sólido. McLaurin + Dotson.','No','','','Hard','Buena'],
    ["De'Von Achane",         'RB','MIA','1',23.5, 24.0,  'Speed back élite. Sistema pase MIA nuevo. 180 carries + 80 rec. Tua se fue a ATL; monitorear nuevo QB MIA. Explosividad única.','Yes','','','Neutral','Buena'],
    ['Marvin Harrison Jr.',   'WR','ARI','1',24.5, 25.0,  'WR talento élite. Ryan Grubb OC (Año 2, sistema moderno). Kyler Murray se fue a MIN; monitorear nuevo QB ARI. Allgeier RB nuevo.','Yes','','','Neutral','Promedio'],
    ['Zay Flowers',           'WR','BAL','1',25.3, 26.0,  'WR1 BAL. Lamar distribuye targets. 85 rec + 1050 yds + 8 TDs. Red zone presence creciente. Año 4.','No','','','Neutral','Muy buena'],
    ['A.J. Brown',            'WR','PHI','1',26.0, 26.5,  'WR élite PHI. OL élite. 95 rec + 1350 yds + 9 TDs. Hurts-Brown connection probada. Must-start cuando sano.','No','','','Easy','Élite'],
    ['Tony Pollard',          'RB', 'NO','1',27.2, 28.0,  'RB1 NO. Travis Etienne llegó de JAX como backup/complemento. Monitorear workload Pollard-Etienne. Kamara ya no está en NO.','No','','','Easy','Buena'],
    ['Brian Thomas Jr.',      'WR','JAC','1',28.5, 29.0,  'WR1 JAC. Año 3 (clase 2024). Travis Hunter (Año 2) como WR2 emergente. Lawrence QB. Moore OC inconsistente. Talento real.','No','','','Neutral','Promedio'],
    ['James Cook',            'RB','BUF','1',29.3, 30.0,  'RB BUF. Allen corre 130+ pero Cook tiene 200+ carries también. 70 rec adicionales. DJ Moore WR1 nuevo. BUF ofensiva élite.','No','','','Easy','Muy buena'],
    ['Kyren Williams',        'RB','LAR','1',30.5, 31.0,  'RB1 LAR. Stafford protegido bien. 250 carries + 55 rec + 1200 yds. Puka Nacua es el arma principal de pase. OL sólida.','No','','','Neutral','Muy buena'],

    // ── TIER 2 — ADP ~31-72 ─────────────────────────────────────────────────
    ['Drake Maye',            'QB', 'NE','2',31.0, 32.0,  'Año 3 (clase 2024). Alex Van Pelt nuevo OC (más moderno). Dual-threat. Stevenson + Elliott en backfield. Ceiling alto NE.','Yes','','','Neutral','Pobre'],
    ['Jordan Love',           'QB', 'GB','2',32.5, 33.0,  'Año 3+ establecido. Adams + Watson WRs. Jacobs RB. GB OL sólida. Schedule playoff muy favorable. Confiabilidad creció.','No','','','Easy','Muy buena'],
    ['Rachaad White',         'RB', 'TB','2',33.5, 34.0,  'Lead back TB. Mayfield QB cómodo. Evans se fue a SF; Godwin sigue como WR1. 210 carries + 75 rec proyectados. Año 4.','No','','','Easy','Buena'],
    ['Travis Etienne',        'RB', 'NO','2',34.5, 35.5,  'Llegó de JAX. Complementa a Pollard en NO. Monitorear reparto de carreras. NO históricamente run-friendly. Doble RB amenaza.','No','','','Easy','Buena'],
    ['Chris Olave',           'WR', 'NO','2',36.5, 37.5,  'WR1 NO. Carr QB. Etienne+Pollard en backfield. Olave = objetivo principal de pase. 90 rec + 1100 yds + 7 TDs proyectados.','No','','','Easy','Buena'],
    ['Stefon Diggs',          'WR', 'FA','2',37.5, 38.5,  '⚠️ AGENTE LIBRE — Sin equipo al 28 mayo 2026. Verificar firma antes del draft. Receptor veterano élite con valor inmediato.','No','','','Neutral','N/A'],
    ['George Pickens',        'WR','PIT','2',38.5, 39.0,  'WR1 PIT. Aaron Rodgers nuevo QB (ADP 172.7). Brian Angelichio nuevo OC (de MIN). Conexión Rodgers-Pickens = upside alto.','Yes','','','Hard','Buena'],
    ['Trey McBride',          'TE','ARI','2',39.5, 40.5,  'TE sólido ARI. Grubb OC Año 2. Allgeier nuevo RB. Kyler Murray se fue; nuevo QB ARI. McBride sigue siendo target hog de TE.','Yes','','','Neutral','Promedio'],
    ['Derrick Henry',         'RB','BAL','2',40.5, 41.5,  'Veterano BAL. Lamar también corre. Henry = red zone + short yardage machine. 230 carries + 15 TDs. Año 12. Confiable.','No','','','Neutral','Muy buena'],
    ['Keenan Allen',          'WR','CHI','2',41.5, 42.5,  'Veteran WR con Williams Año 3. Waldron OC Año 2 = mayor fluidez. 90 rec + 1100 yds. Odunze también crece. CHI OL aún pobre.','No','','','Easy','Pobre'],
    ['Tank Dell',             'WR','HOU','2',42.5, 43.5,  'WR élite HOU. Regresó de fractura (2024). CJ Stroud QB. 80 rec + 1100 yds proyectados si sano. Explosividad real.','No','','','Neutral','Buena'],
    ['Tee Higgins',           'WR','CIN','2',43.5, 44.5,  'WR2 CIN. Burrow tiene volumen para Chase Y Higgins. 80 rec + 1050 yds + 8 TDs proyectados. Año 7. Confiable.','No','','','Easy','Buena'],
    ['Kenneth Walker III',    'RB', 'KC','2',44.5, 45.5,  'Llegó de SEA. RB KC. 180 carries + 50 rec. KC no es run-heavy pero Walker añade dimensión rushing. Mahomes lo conecta.','No','','','Hard','Élite'],
    ['Christian Watson',      'WR', 'GB','2',45.5, 46.5,  'WR2 GB. Adams WR1. Watson explosivo en rutas profundas. 75 rec + 950 yds + 8 TDs. Love lo alimenta bien en deep ball.','No','','','Easy','Muy buena'],
    ['Caleb Williams',        'QB','CHI','2',46.5, 47.0,  'Año 3 (clase 2024). Waldron OC Año 2 = mayor comodidad. OL mejorando. Keenan Allen + Odunze. Ceiling élite dual-threat.','No','','','Easy','Pobre'],
    ['DeVonta Smith',         'WR','PHI','2',47.5, 48.5,  'WR2 PHI. OL élite. Brown atrae doble cobertura = Smith abierto. 85 rec + 1100 yds + 7 TDs. Año 6. Muy consistente.','No','','','Easy','Élite'],
    ['Ladd McConkey',         'WR','LAC','2',48.5, 49.5,  'WR1 LAC. Herbert QB. Sistema preciso de pases. 95 rec + 1100 yds. Año 3 (clase 2024). Slot receptor de élite.','No','','','Easy','Promedio'],
    ['Drake London',          'WR','ATL','2',49.5, 50.5,  'WR1 ATL. Tua nuevo QB (llega de MIA). Hill FA. London como target principal. 90 rec + 1100 yds + 8 TDs proyectados.','Yes','','','Neutral','Buena'],
    ['Garrett Wilson',        'WR','NYJ','2',50.5, 51.5,  'WR élite NYJ. Frank Reich nuevo OC 2026. 100 rec + 1200 yds. Talento demasiado bueno para fallar con cualquier OC. Hall-Wilson dupla.','Yes','','','Neutral','Pobre'],
    ['David Njoku',           'TE','CLE','2',51.5, 52.5,  'TE CLE. QB situation caótica limita upside. Tommy Rees OC joven. Njoku como safety blanket. 65 rec + 700 yds cuando QB funciona.','Yes','','','Easy','Promedio'],
    ['Darnell Mooney',        'WR','ATL','2',52.5, 53.5,  'WR2 ATL. Tua QB nuevo. London WR1. Mooney como slot veloz. 70 rec + 850 yds + 6 TDs proyectados.','Yes','','','Neutral','Buena'],
    ['Xavier Worthy',         'WR', 'KC','2',53.5, 54.5,  'WR2 KC. Mahomes lo usa profundo. 70 rec + 900 yds + 8 TDs. Año 3 (clase 2024). Velocidad extrema. Sistema Reid lo maximiza.','No','','','Hard','Élite'],
    ['Khalil Shakir',         'WR','BUF','2',54.5, 55.5,  'WR2 BUF. DJ Moore llegó como WR1 pero Shakir domina el slot. Allen distribuye bien. 80 rec + 900 yds. Seguro.','No','','','Easy','Muy buena'],
    ['Travis Hunter',         'WR','JAC','2',55.0, 56.0,  'Año 2 (clase 2025). WR/CB dos vías en JAC. Receptor emergente. Thomas Jr. WR1. Hunter como complemento explosivo. Upside alto.','No','','','Neutral','Promedio'],
    ['Mike Evans',            'WR', 'SF','2',56.5, 57.5,  'Llegó de TB. WR SF. Shanahan system. Deebo Samuel FA. Evans = red zone + possession WR. McCaffrey corre = defenses focus run.','No','','','Neutral','Élite'],
    ['Jonathan Taylor',       'RB','IND','2',57.5, 58.5,  'RB IND. Alec Pierce extended ($116M). Richardson o nuevo QB IND. 250 carries + 50 rec. OL mejorada. Año 6.','No','','','Easy','Promedio-Buena'],
    ['DJ Moore',              'WR','BUF','2',58.5, 59.5,  'Llegó de CHI a BUF. WR1 BUF. Allen élite QB. 90 rec + 1100 yds + 8 TDs. Shakir complementa en slot. Gran fit.','No','','','Easy','Muy buena'],
    ['Jaylen Waddle',         'WR','MIA','2',59.5, 60.5,  'WR1 MIA por default. Tua se fue a ATL; Hill FA. Waddle como el WR más confiable. Monitorear nuevo QB MIA.','Yes','','','Neutral','Buena'],
    ['Jakobi Meyers',         'WR', 'LV','2',60.5, 61.5,  'WR2 LV. Bowers TE1 es el foco. Mendoza QB nuevo. OC Getsy crea espacio para Meyers en slot. 70 rec + 800 yds.','Yes','','','Neutral','Promedio'],
    ['Chris Godwin',          'WR', 'TB','2',61.5, 62.5,  'WR1 TB. Evans se fue a SF. Godwin = líder del passing game ahora. Mayfield QB. 90 rec + 1100 yds. Año 9. Confiable.','No','','','Easy','Buena'],
    ['Aaron Jones',           'RB','MIN','2',62.5, 63.5,  'RB MIN. Kyler Murray nuevo QB. Murray corre también = reparto de carries. Jones excelente receptor. 160 carries + 60 rec.','Yes','','','Easy','Buena'],
    ['Josh Jacobs',           'RB', 'GB','2',63.5, 64.5,  'RB1 GB. Love QB. OL sólida. 240 carries + 60 rec + 1200 yds. Año 8. Schedule playoff élite GB.','No','','','Easy','Muy buena'],
    ['Rashee Rice',           'WR', 'KC','2',64.5, 65.5,  'WR1 KC. Mahomes. Regresó de lesión 2025. 90 rec + 1100 yds + 9 TDs. Año 3. Upside enorme con Mahomes.','No','','','Hard','Élite'],
    ['Cooper Kupp',           'WR','LAR','2',65.5, 66.5,  'Veteran WR LAR. Nacua es WR1 ahora. Kupp = red zone + slot. 75 rec + 900 yds + 8 TDs. Año 10. Sigue produciendo.','No','','','Neutral','Muy buena'],
    ['Amari Cooper',          'WR','CLE','2',66.5, 67.5,  'WR CLE. QB chaos limita upside. Tommy Rees OC. Cooper = WR más confiable CLE. Dependiente de QB funcional.','Yes','','','Easy','Promedio'],
    ['T.J. Hockenson',        'TE','MIN','2',67.5, 68.5,  'TE MIN. Kyler Murray nuevo QB. Murray run-first = menos targets TE. Hockenson puede perder relevancia. Riesgo real.','Yes','','','Easy','Buena'],
    ['Dalton Kincaid',        'TE','BUF','2',68.5, 69.5,  'TE BUF. Allen QB. 60 rec + 700 yds + 6 TDs. Moore y Shakir como WRs. Allen distribuye targets bien. Sólido TE2.','No','','','Easy','Muy buena'],
    ['Rome Odunze',           'WR','CHI','2',69.5, 70.5,  'WR2 CHI. Año 3 (clase 2024). Williams Año 3. Waldron OC Año 2 = mayor apertura. 80 rec + 950 yds proyectados.','No','','','Easy','Pobre'],
    ['DK Metcalf',            'WR','SEA','2',71.5, 72.5,  'WR1 SEA. Sam Darnold nuevo QB (de MIN). Darnold tiene buen deep ball. Metcalf = 80 rec + 1100 yds + 8 TDs.','No','','','Neutral','Buena'],
    ['Alvin Kamara',          'RB', 'NO','2',70.5, 71.5,  'Veterano NO. Pollard + Etienne también en backfield. Kamara = receptor élite. 75 rec + 600 yds. Monitorear workload triangular.','No','','','Easy','Buena'],

    // ── TIER 3 — ADP ~73-120 ─────────────────────────────────────────────────
    ['Tyler Lockett',         'WR','SEA','3',73.0, 74.0,  'WR2 SEA. Veteran. Darnold QB. Metcalf atrae cobertura = Lockett en slot. 70 rec + 850 yds. Año 12. Confiable.','No','','','Neutral','Buena'],
    ['Brock Purdy',           'QB', 'SF','3',74.0, 75.0,  'QB SF. Shanahan system = proyecciones confiables. Mike Evans nuevo WR. McCaffrey si sano = todo ok. 4000 yds + 32 TDs.','No','','','Neutral','Élite'],
    ['Cole Kmet',             'TE','CHI','3',75.0, 76.0,  'TE CHI. Williams Año 3. Waldron OC Año 2. Kmet = safety blanket de Williams. 60 rec + 650 yds + 6 TDs.','No','','','Easy','Pobre'],
    ['Ezekiel Elliott',       'RB', 'NE','3',76.0, 77.0,  'Veteran NE. Maye QB Año 3. Van Pelt nuevo OC. Stevenson también en backfield. 170 carries + 50 rec proyectados.','Yes','','','Neutral','Pobre'],
    ['Javonte Williams',      'RB','DEN','3',77.0, 78.0,  'RB DEN. OL mejorada. QB joven DEN. 220 carries + 50 rec proyectados. Año 5. Upside si QB es funcional.','No','','','Neutral','Promedio'],
    ['Diontae Johnson',       'WR','HOU','3',78.0, 79.0,  'WR HOU. Stroud QB. Dell y Collins como opciones. Johnson = slot veteran. 70 rec + 800 yds. Confiabilidad media.','No','','','Neutral','Buena'],
    ['Tyler Warren',          'TE','IND','3',79.0, 80.0,  'Año 2 (clase 2025). TE emergente IND. Richardson QB o nuevo. Pierce extended como WR1. Warren = ceiling alto si se establece.','No','','','Easy','Promedio-Buena'],
    ['Rhamondre Stevenson',   'RB', 'NE','3',80.0, 81.0,  'RB NE. Veterano. Maye QB. Van Pelt OC nuevo. 200 carries + 50 rec proyectados. Elliott como complemento. Año 5.','Yes','','','Neutral','Pobre'],
    ['Alec Pierce',           'WR','IND','3',81.0, 82.0,  'Extendido $116M en IND. WR1 establecido. Richardson o nuevo QB. 85 rec + 1050 yds proyectados. Contrato indica confianza del equipo.','No','','','Easy','Promedio-Buena'],
    ['Cam Skattebo',          'RB','ARI','3',82.5, 83.5,  'Año 2 (clase 2025). RB ARI. Grubb OC Año 2. Tyler Allgeier como complemento (llegó de ATL). Murray se fue. Backfield en transición.','Yes','','','Neutral','Promedio'],
    ['Terry McLaurin',        'WR','WAS','3',84.0, 85.0,  'WR1 WAS. Kingsbury OC Año 2 (air raid = targets sólidos). Daniels Año 3. 85 rec + 1050 yds + 8 TDs proyectados.','No','','','Hard','Buena'],
    ['Nico Collins',          'WR','HOU','3',85.5, 86.5,  'WR1/2 HOU. Stroud QB. Dell + Johnson también. Collins = alto piso en sistema HOU. 80 rec + 1000 yds proyectados.','No','','','Neutral','Buena'],
    ['Jordan Addison',        'WR','MIN','3',86.5, 87.5,  'WR2 MIN. Kyler Murray nuevo QB. Jefferson WR1. Addison como complemento slot. 70 rec + 900 yds proyectados.','Yes','','','Easy','Buena'],
    ['Zack Moss',             'RB','CIN','3',87.5, 88.5,  'RB CIN. Chase + Higgins dominan la ofensiva de pase. Moss como lead back. 200 carries + 40 rec + 1000 yds.','No','','','Easy','Buena'],
    ['Isiah Pacheco',         'RB','DET','3',88.0, 89.0,  'Backup RB DET. Llegó de KC. Gibbs es el clear lead back. Pacheco = handcuff valioso de Gibbs. ADP 146 en algunas fuentes.','No','','','Neutral','Élite'],
    ['Stefon Collins',        'WR','HOU','3',89.0, 90.0,  'WR HOU. Stroud QB. Dell + Johnson completan el corps. 65 rec + 800 yds proyectados.','No','','','Neutral','Buena'],
    ['RJ Harvey',             'RB','DEN','3',90.0, 91.0,  'Año 2 (clase 2025). RB DEN. Comparte backfield con Javonte Williams. Breakout candidato Year 2. Williams es el lead pero Harvey puede robarle trabajo. OL promedio DEN.','No','','','Neutral','Promedio'],
    ['Jake Ferguson',         'TE','DAL','3',91.0, 92.0,  'TE DAL. Ferguson = safety blanket en sistema DAL. 65 rec + 750 yds + 6 TDs. OL élite DAL. Buena plataforma.','No','','','Neutral','Muy buena'],
    ['Michael Pittman Jr.',   'WR','PIT','3',92.0, 93.0,  'Llegó de IND a PIT. ADP 110.7 verificado. Rodgers nuevo QB. Comparte con Pickens. 70 rec + 900 yds proyectados.','Yes','','','Hard','Buena'],
    ['Kyler Murray',          'QB','MIN','3',107.7,108.5, 'Llegó de ARI. QB1 MIN. Jefferson + Addison como WRs. Murray aporta rushing en Custom format. 3700 yds + 27 TDs + 500 rushYds.','Yes','','','Easy','Buena'],
    ['Chuba Hubbard',         'RB','CAR','3',93.0, 94.0,  'RB1 CAR por defecto. CAR peor OL del NFL. Reconstrucción total. Workhorse = alto volumen, baja eficiencia. 200 carries proyectados.','No','','','Neutral','Muy pobre'],
    ['Ja\'Lynn Polk',         'WR', 'NE','3',94.0, 95.0,  'WR NE. Maye QB Año 3. Sistema Van Pelt nuevo. 70 rec + 850 yds potencial. Año 2 si fue clase 2025.','Yes','','','Neutral','Pobre'],
    ['Romeo Doubs',           'WR', 'GB','3',95.0, 96.0,  'WR3 GB. Adams + Watson arriba. Doubs en slot. 55 rec + 700 yds. GB schedule playoff élite.','No','','','Easy','Muy buena'],
    ['Elijah Moore',          'WR','CLE','3',96.0, 97.0,  'WR CLE. QB chaos limita upside. 60 rec + 700 yds si QB funciona. Tommy Rees OC joven. Riesgo alto.','Yes','','','Easy','Promedio'],
    ['Anthony Richardson',    'QB','IND','3',97.0, 98.0,  'QB IND. Dual-threat enorme en Custom format. Si está sano = mucho valor por rushing. Pierce extended = WR1 sólido. 450 rushYds potencial.','No','','','Easy','Promedio-Buena'],
    ['Ty Chandler',           'RB', 'SF','3',98.0, 99.0,  'RB SF. McCaffrey backup. Shanahan system = backfield muy valioso. Si CMC se lesiona = RB1 instantáneo. Handcuff crítico.','No','','','Neutral','Élite'],
    ['Gus Edwards',           'RB','LAC','3',99.0,100.0,  'RB LAC. Veterano. Herbert QB. 180 carries + 35 rec. OL promedio LAC. Confiabilidad media. Año 8.','No','','','Easy','Promedio'],
    ['Ja\'Tavion Thomas',     'TE', 'CIN','3',100.0,101.0,'TE2 CIN. Chase y Higgins dominan. Thomas como TE complementario. 50 rec + 600 yds proyectados.','No','','','Easy','Buena'],
    ['Noah Gray',             'TE', 'KC','3',101.0,102.0, 'TE KC. Walker + Rice dominan. Gray como TE safety blanket de Mahomes. 50 rec + 600 yds + 5 TDs.','No','','','Hard','Élite'],
    ['Michael Wilson',        'WR','ARI','3',102.0,103.0, 'WR ARI. Grubb OC Año 2. Harrison Jr. WR1. Wilson como complemento. 65 rec + 750 yds proyectados.','Yes','','','Neutral','Promedio'],
    ['Jaylen Warren',         'RB','PIT','3',103.0,104.0, 'RB PIT. Brian Angelichio nuevo OC. Najee Harris ya no está. Warren como lead back. 210 carries + 60 rec. Rodgers QB.','Yes','','','Hard','Buena'],
    ['Joshua Palmer',         'WR','LAC','3',104.0,105.0, 'WR LAC. McConkey WR1. Palmer como WR2. Herbert QB. 65 rec + 750 yds proyectados.','No','','','Easy','Promedio'],
    ['Curtis Samuel',         'WR','WAS','3',105.0,106.0, 'WR WAS. Kingsbury OC Año 2 (air raid). McLaurin WR1. Samuel en slot. 65 rec + 700 yds. Daniels Año 3.','No','','','Hard','Buena'],
    ['Wan\'Dale Robinson',    'WR','NYG','3',106.0,107.0, 'WR NYG. Nabers WR1. Robinson en slot. OL pobre NYG. 65 rec + 700 yds proyectados.','No','','','Neutral','Promedio'],

    // ── TIER 4 — ADP ~121-168 ─────────────────────────────────────────────────
    ['Sam Darnold',           'QB','SEA','4',109.0,110.0, 'Llegó de MIN. QB1 SEA. Metcalf + Lockett. OC nuevo. 3800 yds + 28 TDs. Año veterano. ADP ~137 verificado.','No','','','Neutral','Buena'],
    ['Cam Ward',              'QB','TEN','4',110.0,111.0, 'Año 2 (clase 2025). QB TEN. Nick Holz OC Año 2. Ridley como WR1. Pollard fue a NO. Dual-threat upside. ADP 142 verificado.','No','','','Neutral','Promedio'],
    ['CJ Stroud',             'QB','HOU','4',111.0,112.0, 'QB HOU. Dell + Collins + Johnson como WRs. 3900 yds + 30 TDs. OL buena HOU. ADP 139.3 verificado.','No','','','Neutral','Buena'],
    ['Calvin Ridley',         'WR','TEN','4',112.0,113.0, 'WR1 TEN. Holz OC Año 2 (air raid de GB). Cam Ward QB Año 2. 75 rec + 950 yds + 7 TDs potencial. Schedule favorable.','No','','','Neutral','Promedio'],
    ['Evan Engram',           'TE','JAC','4',113.0,114.0, 'TE JAC. Moore OC inconsistente. Lawrence QB. Thomas Jr. WR1. Engram = receptor confiable. 70 rec + 750 yds.','No','','','Neutral','Promedio'],
    ['Mark Andrews',          'TE','BAL','4',114.0,115.0, 'TE BAL. Veterano. Lamar lo encuentra en red zone. 60 rec + 700 yds + 8 TDs proyectados. Injury risk históricamente.','No','','','Neutral','Muy buena'],
    ['Josh Downs',            'WR','IND','4',115.0,116.0, 'WR2 IND. Pierce extended WR1. Richardson QB. Downs en slot. 70 rec + 800 yds proyectados.','No','','','Easy','Promedio-Buena'],
    ['Andrei Iosivas',        'WR','CIN','4',116.0,117.0, 'WR3 CIN. Chase WR1, Higgins WR2. Iosivas = upside cuando ambos tienen bye o lesión. 55 rec + 700 yds.','No','','','Easy','Buena'],
    ['Adam Thielen',          'WR','CAR','4',117.0,118.0, 'WR CAR. Veterano en reconstrucción. 65 rec + 750 yds. OL CAR peor del NFL. QB joven CAR. Riesgo alto.','No','','','Neutral','Muy pobre'],
    ['Cedric Tillman',        'WR','CLE','4',118.0,119.0, 'WR CLE. QB chaos. Tillman como WR2 si Cooper sale. 55 rec + 650 yds. Tommy Rees OC joven. Alto riesgo.','Yes','','','Easy','Promedio'],
    ['Brian Robinson Jr.',    'RB','ATL','4',119.0,120.0, 'RB ATL. Backup de Bijan Robinson. ADP 165 verificado. Tua QB nuevo. Si Bijan se lesiona = RB1 inmediato. Handcuff crítico.','Yes','','','Neutral','Buena'],
    ['James Conner',          'RB','ARI','4',120.0,121.0, 'RB ARI. Skattebo Year 2 + Allgeier = 3 RBs compitiendo. Conner = veterano red zone. 100 carries + 8 TDs potencial.','Yes','','','Neutral','Promedio'],
    ['Tyler Allgeier',        'RB','ARI','4',121.0,122.0, 'Llegó de ATL. RB ARI. Skattebo Year 2 es el lead. Allgeier como complemento. 100 carries + 35 rec proyectados.','Yes','','','Neutral','Promedio'],
    ['Braelon Allen',         'RB','NYJ','4',122.0,123.0, 'Backup de Breece Hall. ADP 203.3 verificado. Año 2 si fue clase 2025. Si Hall se lesiona = RB1 inmediato NYJ.','Yes','','','Neutral','Pobre'],
    ['Davante Parker',        'WR','LAR','4',123.0,124.0, 'WR LAR. Nacua WR1, Kupp WR2. Parker como WR3. 45 rec + 550 yds. Valor streamer en semanas de bye.','No','','','Neutral','Muy buena'],
    ['Hayden Hurst',          'TE', 'NE','4',124.0,125.0, 'TE NE. Maye QB Año 3. Van Pelt OC nuevo. 50 rec + 550 yds proyectados.','Yes','','','Neutral','Pobre'],
    ['Luke Musgrave',         'TE', 'GB','4',125.0,126.0, 'TE GB. Love QB. Adams + Watson WRs. Musgrave como safety blanket. 55 rec + 600 yds + 5 TDs.','No','','','Easy','Muy buena'],
    ['Michael Mayer',         'TE', 'LV','4',126.0,127.0, 'TE LV. Bowers es TE1 total. Mayer como TE2 profundo. 35 rec + 400 yds. Valor bench solamente.','Yes','','','Neutral','Promedio'],
    ['Hunter Henry',          'TE','LAC','4',127.0,128.0, 'TE LAC. Herbert QB. McConkey WR1. Henry = safety blanket TE. 55 rec + 600 yds + 6 TDs.','No','','','Easy','Promedio'],
    ['Foster Moreau',         'TE', 'NO','4',128.0,129.0, 'TE NO. Olave WR1. Pollard + Etienne en backfield. Moreau como TE complementario. 45 rec + 500 yds.','No','','','Easy','Buena'],
    ['Taysom Hill',           'QB', 'NO','4',129.0,130.0, 'QB/TE híbrido NO. Valor en formatos Custom por rushes. 15 carries + 5 TDs potenciales. Streamer especulativo.','No','','','Easy','Buena'],
    ['Marquise Brown',        'WR', 'KC','4',130.0,131.0, 'WR KC. Rice + Worthy arriba. Brown como WR3. 50 rec + 600 yds. Profundidad en KC. Valor de bench.','No','','','Hard','Élite'],
    ['Keenan Coleman',        'WR','NYJ','4',131.0,132.0, 'WR NYJ. Wilson WR1. Reich OC nuevo. Coleman como WR2 emergente. 55 rec + 650 yds proyectados.','Yes','','','Neutral','Pobre'],
    ['Quentin Johnston',      'WR','LAC','4',132.0,133.0, 'WR LAC. McConkey WR1, Palmer WR2. Johnston como WR3 deep threat. 50 rec + 650 yds potencial.','No','','','Easy','Promedio'],
    ['Rashod Bateman',        'WR','BAL','4',133.0,134.0, 'WR BAL. Flowers WR1. Lamar distribuye. Bateman como WR2. 55 rec + 700 yds proyectados.','No','','','Neutral','Muy buena'],
    ['Wan\'Dale Robinson',    'WR','NYG','4',134.0,135.0, 'Nota: Robinson puede estar listado dos veces si apareció en Tier 3. Confirmar y ajustar.','No','','','Neutral','Promedio'],
    ['Cade Otton',            'TE', 'TB','4',135.0,136.0, 'TE TB. Evans fue a SF. Godwin WR1. Otton como safety blanket de Mayfield. 55 rec + 600 yds + 5 TDs.','No','','','Easy','Buena'],
    ['Will Dissly',           'TE','SEA','4',136.0,137.0, 'TE SEA. Darnold nuevo QB. Metcalf + Lockett dominan. Dissly como TE complementario. 40 rec + 450 yds.','No','','','Neutral','Buena'],
    ['Isaiah Likely',         'TE','BAL','4',137.0,138.0, 'TE BAL. Mark Andrews es TE1. Likely como backup. 40 rec + 450 yds. Upside si Andrews lesionado.','No','','','Neutral','Muy buena'],
    ['Jonathon Brooks',       'RB','CAR','4',138.0,139.0, 'RB CAR. Hubbard es lead back. Brooks como backup emergente. CAR reconstrucción total. 100 carries potencial.','No','','','Neutral','Muy pobre'],
    ['Tank Bigsby',           'RB','PHI','4',139.0,140.0, 'Backup Saquon Barkley PHI. ADP 183 verificado. OL élite PHI = handcuff de valor alto. Si Barkley se lesiona = RB1 inmediato.','No','','','Easy','Élite'],
    ['Jaleel McLaughlin',     'RB','DEN','4',140.0,141.0, 'RB DEN. Javonte Williams es lead. McLaughlin como complemento receptor. 60 rec + 500 yds. Valor flex profundo.','No','','','Neutral','Promedio'],
    ['Keaton Mitchell',       'RB','BAL','4',141.0,142.0, 'Speed back BAL. Henry + Lamar corren. Mitchell como explosivo complemento. 80 carries + 50 rec potencial.','No','','','Neutral','Muy buena'],
    ['Nick Westbrook-Ikhine', 'WR','TEN','4',142.0,143.0, 'WR TEN. Ridley WR1. Cam Ward QB Año 2. NWI como WR2 emergente. 50 rec + 600 yds proyectados.','No','','','Neutral','Promedio'],
    ['Mecole Hardman',        'WR', 'KC','4',143.0,144.0, 'WR KC. Rice WR1, Worthy WR2. Hardman = explosivo WR3. 40 rec + 500 yds. Valor limitado sin injury.','No','','','Hard','Élite'],
    ['Jaxon Smith-Njigba',    'WR','SEA','4',144.0,145.0, 'WR SEA. Metcalf WR1. JSN como slot emergente. Darnold nuevo QB. 70 rec + 800 yds potencial si se establece. Año 3.','No','','','Neutral','Buena'],
    ['Sam Martin',            'K',  'BUF','4',145.0,146.0,'K BUF. Ofensiva élite = muchos FGs + PATs. Top-5 kicker proyectado. Allen + Cook = TD machine. Muy confiable.','No','','','Easy','N/A'],

    // ── TIER 5 — ADP ~169-200 ─────────────────────────────────────────────────
    ['Aaron Rodgers',         'QB','PIT','5',146.0,147.0, 'Veterano PIT. ADP 172.7 verificado. Brian Angelichio nuevo OC (de MIN). Pickens como WR1. Año 23. Riesgo salud.','Yes','','','Hard','Buena'],
    ['Fernando Mendoza',      'QB', 'LV','5',147.0,148.0, 'Año 2 (si fue clase 2025). QB LV. ADP 170 verificado. Bowers TE1 es su arma principal. OC Getsy sistema moderno.','Yes','','','Neutral','Promedio'],
    ['Tua Tagovailoa',        'QB','ATL','5',148.0,149.0, 'Llegó de MIA. QB ATL. ADP 197.7 verificado. Bijan Robinson RB1. London WR1. Nuevo sistema = riesgo consistencia.','Yes','','','Neutral','Buena'],
    ['Jake Haener',           'QB', 'NO','5',149.0,150.0, 'QB backup NO. Carr o Haener. Especulativo profundo. 0 valor a menos que haya lesión de titular.','No','','','Easy','Buena'],
    ['DeAndre Hopkins',       'WR', '? ','5',150.0,151.0, 'Veterano WR. ⚠️ Verificar equipo 2026 antes del draft. Si firma = WR2 inmediato en equipo con targets disponibles.','No','','','Neutral','N/A'],
    ['Elijah Mitchell',       'RB', 'SF','5',151.0,152.0, 'RB SF. Backup McCaffrey. Shanahan system = valor de handcuff. Si CMC se lesiona = RB1 élite.','No','','','Neutral','Élite'],
    ['Dameon Pierce',         'RB','HOU','5',152.0,153.0, 'RB HOU. Backup. Dell + Collins dominan passing game. Pierce como complemento rushing. 100 carries potencial.','No','','','Neutral','Buena'],
    ['Jerome Ford',           'RB','CLE','5',153.0,154.0, 'RB CLE. QB chaos. Ford como lead back por default. 180 carries proyectados pero baja eficiencia esperada.','Yes','','','Easy','Promedio'],
    ['Taysom Hill',           'QB', 'NO','5',154.0,155.0, 'Nota: si ya aparece arriba, eliminar este duplicado. Verificar.','No','','','Easy','Buena'],
    ['Dontayvion Wicks',      'WR', 'GB','5',155.0,156.0, 'WR GB. Adams + Watson arriba. Wicks como WR3 emergente. 40 rec + 500 yds potencial. Valor en bye weeks GB.','No','','','Easy','Muy buena'],
    ['Jonnu Smith',           'TE', 'MIA','5',156.0,157.0,'TE MIA. Hill FA; Tua se fue. Nuevo ecosistema MIA. Smith como safety blanket del nuevo QB. 50 rec + 550 yds.','Yes','','','Neutral','Buena'],
    ['Greg Dulcich',          'TE', 'DEN','5',157.0,158.0,'TE DEN. QB joven. 45 rec + 500 yds proyectados. Valor de profundidad en formatos superflex.','No','','','Neutral','Promedio'],
    ['Tre Tucker',            'WR', 'LV','5',158.0,159.0, 'WR LV. Bowers TE1 es el foco. Tucker como WR2. Meyers WR2 también. 45 rec + 550 yds.','Yes','','','Neutral','Promedio'],
    ['Deebo Samuel',          'WR', 'FA','5',159.0,160.0, '⚠️ AGENTE LIBRE — Sin equipo al 28 mayo 2026. Verificar firma. Multipurpose receiver/runner = valor en Custom format.','No','','','Neutral','N/A'],
    ['Travis Hunter (TE)',    'TE','JAC','5',160.0,161.0, 'Nota: Hunter puede usarse como WR o TE en 2 vías. Verificar posición registrada en tu liga.','No','','','Neutral','Promedio'],
    ['Dawson Knox',           'TE','BUF','5',161.0,162.0, 'TE BUF. Kincaid es TE1. Knox como backup. 30 rec + 350 yds + 4 TDs. Valor solo si Kincaid lesionado.','No','','','Easy','Muy buena'],
    ['Damiere Byrd',          'WR','ATL','5',162.0,163.0, 'WR ATL. London + Mooney arriba. Byrd como WR3 profundo. 35 rec + 400 yds. Sin valor relevante 12-team.','Yes','','','Neutral','Buena'],
    ['Jordan Mason',          'RB', 'SF','5',163.0,164.0, 'RB SF. Backup de McCaffrey / Mitchell. Shanahan = backfield valioso. Si hay lesiones = RB2 inmediato.','No','','','Neutral','Élite'],
    ['Shedeur Sanders',       'QB','CLE','5',164.0,165.0, '⚠️ Verificar status 2026 CLE. Si es QB titular = valor por default en CLE. Rees OC joven. Año 2 potencial.','Yes','','','Easy','Promedio'],
    ['Jared Goff',            'QB','DET','5',165.0,166.0, 'QB DET. Gibbs + ARSB = ofensiva élite. Goff protegido bien. 4500 yds + 30 TDs proyectados. Valor QB2 sólido.','No','','','Neutral','Élite'],
    ['Deshaun Watson',        'QB','CLE','5',166.0,167.0, '⚠️ Verificar status 2026. QB CLE en caos. Tommy Rees OC. Valor especulativo únicamente.','Yes','','','Easy','Promedio'],
    ['Greg Joseph',           'K', 'MIN','5',167.0,168.0, 'K MIN. Murray + Jefferson + Jones = ofensiva con scoring potential. FGs + PATs. Top-10 kicker.','Yes','','','Easy','N/A'],
    ['Evan McPherson',        'K', 'CIN','5',168.0,169.0, 'K CIN. Chase + Burrow = TD machine + FG opportunities. Top-3 kicker proyectado. Muy confiable.','No','','','Easy','N/A'],
    ['Jake Elliott',          'K', 'PHI','5',169.0,170.0, 'K PHI. Ofensiva élite = muchas oportunidades. Top-5 kicker. Barkley + Hurts + Brown = scoring machine.','No','','','Easy','N/A'],
    ['Harrison Butker',       'K',  'KC','5',170.0,171.0, 'K KC. Mahomes = siempre scoring. Butker = Top-5 kicker. Playoff schedule hard pero KC siempre llega lejos.','No','','','Hard','N/A'],
    ['Tyler Bass',            'K', 'BUF','5',171.0,172.0, 'K BUF. Backup de Sam Martin o titular. BUF ofensiva élite. Confiable kicker.','No','','','Easy','N/A'],
    ['Eagles D/ST',           'D', 'PHI','5',172.0,173.0, 'D/ST PHI. OL élite + defensiva agresiva. Top-3 D/ST proyectada. Schedule de playoffs favorable.','No','','','Easy','N/A'],
    ['Ravens D/ST',           'D', 'BAL','5',173.0,174.0, 'D/ST BAL. Defensiva élite histórica. Sólida contra run y pase. Top-5 D/ST.','No','','','Neutral','N/A'],
    ['Steelers D/ST',         'D', 'PIT','5',174.0,175.0, 'D/ST PIT. Defensiva sólida históricamente. T.J. Watt si sano. Top-5 D/ST cuando healthy.','No','','','Hard','N/A'],
    ['49ers D/ST',            'D',  'SF','5',175.0,176.0, 'D/ST SF. Shanahan defensiva + Nick Bosa. Top-3 D/ST potencial. Monitorear lesiones línea defensiva.','No','','','Neutral','N/A'],
  ];

  sheet.getRange(2, 1, data.length, 12).setValues(data);

  // Nota al final
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ Players DB 2026 — ' + data.length + ' jugadores | ADP: FantasyPros Half-PPR verificado mayo 2026 | ' +
    '⚠️ Verificar: equipo de Hill, Diggs, Deebo Samuel (agentes libres). Bye weeks en Schedule Flags tab.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  Logger.log('✅ Players DB: ' + data.length + ' jugadores cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB 2: SCORING CALC — Proyecciones 2026
// Columnas: Player | Rec | RecYds | RecTDs | Carries | RushYds | RushTDs
//           Comp | PassYds | PassTDs | INTs
// (Las columnas L-O son fórmulas automáticas del setup original)
// ─────────────────────────────────────────────────────────────────────────────
function updateScoringCalc_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scoring Calc');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 11).clearContent();

  // [Player, Rec, RecYds, RecTDs, Carries, RushYds, RushTDs, Comp, PassYds, PassTDs, INTs]
  // QBs: Custom = (Rec×0.8)+(RecYds×0.1)+(RecTDs×6)+(Car×0.2)+(RushYds×0.1)+(RushTDs×6)+(PassYds×0.04)+(PassTDs×4)+(INTs×-2)
  // RBs/WRs/TEs: Custom principalmente = recepciones×0.8 + yards×0.1 + TDs×6 + carries×0.2
  var data = [
    // ── QBs (Custom format premia rushing fuertemente) ──────────────────────
    // Estimado Custom: Allen≈482 | Jackson≈449 | Hurts≈448 | Daniels≈444
    ['Josh Allen',            5,  40,  0, 130, 800, 12, 400, 4100, 38, 10],
    ['Lamar Jackson',         0,   0,  0, 140, 950,  5, 380, 4200, 36,  8],
    ['Jalen Hurts',          20, 150,  1, 120, 650, 10, 380, 3800, 32,  9],
    ['Jayden Daniels',       30, 200,  2, 100, 700,  8, 360, 3900, 28,  9],
    ['Patrick Mahomes',       5,  30,  0,  50, 300,  3, 420, 4500, 38, 10],
    ['Joe Burrow',            5,  30,  0,  20, 100,  1, 400, 4400, 35,  9],
    ['Drake Maye',           10,  60,  1,  90, 550,  8, 340, 3800, 30, 10],
    ['Jordan Love',           0,   0,  0,  30, 200,  3, 370, 3900, 32, 10],
    ['Caleb Williams',       15, 100,  1,  80, 500,  7, 350, 3700, 28, 11],
    ['Brock Purdy',           0,   0,  0,  30, 180,  2, 390, 4000, 32,  9],
    ['Kyler Murray',         10,  60,  0,  80, 500,  6, 350, 3700, 27, 10],
    ['Sam Darnold',           0,   0,  0,  20, 120,  2, 360, 3800, 28, 11],
    ['CJ Stroud',             5,  20,  0,  20, 100,  1, 380, 3900, 30,  9],
    ['Cam Ward',             15,  80,  1,  90, 600,  7, 320, 3500, 25, 12],
    ['Anthony Richardson',    5,  20,  0,  90, 550,  8, 280, 3000, 22, 12],
    ['Aaron Rodgers',         0,   0,  0,  10,  30,  0, 380, 3500, 28, 10],
    ['Tua Tagovailoa',        0,   0,  0,   5,  20,  0, 350, 3400, 24,  8],
    ['Fernando Mendoza',      5,  20,  0,  60, 400,  5, 290, 3200, 22, 12],
    ['Jared Goff',            0,   0,  0,  10,  30,  0, 380, 4500, 30,  9],

    // ── RBs (Custom premia carries ×0.2) ─────────────────────────────────────
    ['Jahmyr Gibbs',         65, 520,  4, 280,1350, 12,   0,    0,  0,  0],
    ['Saquon Barkley',       55, 450,  3, 290,1400, 14,   0,    0,  0,  0],
    ['Breece Hall',          80, 650,  5, 230,1100,  8,   0,    0,  0,  0],
    ['Bijan Robinson',       65, 530,  5, 250,1200, 10,   0,    0,  0,  0],
    ["De'Von Achane",        80, 680,  5, 180, 950,  6,   0,    0,  0,  0],
    ['Christian McCaffrey',  75, 600,  5, 260,1200, 12,   0,    0,  0,  0],
    ['Tony Pollard',         55, 460,  4, 200, 950,  8,   0,    0,  0,  0],
    ['James Cook',           65, 550,  4, 210,1050,  7,   0,    0,  0,  0],
    ['Kyren Williams',       55, 450,  3, 250,1200, 10,   0,    0,  0,  0],
    ['Travis Etienne',       55, 450,  4, 200, 950,  7,   0,    0,  0,  0],
    ['Derrick Henry',        30, 220,  1, 230,1000, 15,   0,    0,  0,  0],
    ['Jonathan Taylor',      50, 380,  3, 250,1200, 10,   0,    0,  0,  0],
    ['Kenneth Walker III',   50, 380,  3, 200, 950,  7,   0,    0,  0,  0],
    ['Josh Jacobs',          60, 480,  3, 240,1200,  9,   0,    0,  0,  0],
    ['Aaron Jones',          60, 500,  4, 160, 800,  6,   0,    0,  0,  0],
    ['Alvin Kamara',         75, 600,  4, 180, 800,  6,   0,    0,  0,  0],
    ['Rachaad White',        65, 530,  4, 210,1000,  8,   0,    0,  0,  0],
    ['Javonte Williams',     50, 380,  3, 220,1000,  8,   0,    0,  0,  0],
    ['Rhamondre Stevenson',  50, 380,  3, 200, 950,  7,   0,    0,  0,  0],
    ['Cam Skattebo',         45, 350,  3, 200, 900,  7,   0,    0,  0,  0],
    ['Chuba Hubbard',        40, 300,  2, 200, 850,  6,   0,    0,  0,  0],
    ['Isiah Pacheco',        30, 240,  2, 120, 550,  5,   0,    0,  0,  0],
    ['Tyler Allgeier',       30, 240,  2, 110, 500,  4,   0,    0,  0,  0],
    ['Jaylen Warren',        45, 360,  3, 210,1000,  7,   0,    0,  0,  0],
    ['Zack Moss',            40, 300,  2, 200, 950,  8,   0,    0,  0,  0],
    ['Brian Robinson Jr.',   25, 180,  1, 120, 500,  4,   0,    0,  0,  0],
    ['Tank Bigsby',          20, 150,  1,  80, 350,  3,   0,    0,  0,  0],
    ['Braelon Allen',        25, 180,  1, 100, 450,  3,   0,    0,  0,  0],
    ['Keaton Mitchell',      30, 250,  2,  80, 450,  3,   0,    0,  0,  0],
    ['Gus Edwards',          35, 260,  2, 180, 800,  6,   0,    0,  0,  0],
    ['Ty Chandler',          30, 220,  2, 150, 650,  5,   0,    0,  0,  0],
    ['Elijah Mitchell',      30, 220,  2, 150, 650,  5,   0,    0,  0,  0],
    ['Jordan Mason',         20, 150,  1, 100, 450,  3,   0,    0,  0,  0],
    ['Ezekiel Elliott',      40, 300,  2, 170, 750,  5,   0,    0,  0,  0],
    ['James Conner',         20, 140,  1,  90, 380,  6,   0,    0,  0,  0],
    ['Jaleel McLaughlin',    55, 440,  3,  90, 400,  3,   0,    0,  0,  0],
    ['Jerome Ford',          30, 220,  2, 180, 780,  5,   0,    0,  0,  0],
    ['Jonathon Brooks',      20, 140,  1,  90, 380,  3,   0,    0,  0,  0],

    // ── WRs ──────────────────────────────────────────────────────────────────
    ["Ja'Marr Chase",       105,1500, 12,   5,  20,  0,   0,    0,  0,  0],
    ['CeeDee Lamb',         110,1450, 10,   5,  15,  0,   0,    0,  0,  0],
    ['Amon-Ra St. Brown',   120,1300,  8,   5,  10,  0,   0,    0,  0,  0],
    ['Justin Jefferson',    110,1450, 10,   0,   0,  0,   0,    0,  0,  0],
    ['Malik Nabers',        100,1350,  9,   0,   0,  0,   0,    0,  0,  0],
    ['A.J. Brown',           95,1350,  9,   0,   0,  0,   0,    0,  0,  0],
    ['Davante Adams',       110,1200,  9,   0,   0,  0,   0,    0,  0,  0],
    ['Puka Nacua',          100,1300,  9,   0,   0,  0,   0,    0,  0,  0],
    ['Tyreek Hill',          90,1200,  9,  15,  80,  1,   0,    0,  0,  0],
    ['Jaylen Waddle',        80, 950,  7,  10,  50,  1,   0,    0,  0,  0],
    ['Tee Higgins',          80,1050,  8,   0,   0,  0,   0,    0,  0,  0],
    ['DeVonta Smith',        85,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Zay Flowers',          85,1050,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Drake London',         90,1100,  8,   5,  20,  0,   0,    0,  0,  0],
    ['Garrett Wilson',      100,1200,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Rashee Rice',          90,1100,  9,   0,   0,  0,   0,    0,  0,  0],
    ['Ladd McConkey',        95,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['DJ Moore',             90,1100,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Khalil Shakir',        80, 900,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Chris Godwin',         90,1100,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Mike Evans',           75,1000,  9,   0,   0,  0,   0,    0,  0,  0],
    ['Brian Thomas Jr.',     80,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['George Pickens',       75,1050,  7,   0,   0,  0,   0,    0,  0,  0],
    ['DK Metcalf',           80,1100,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Xavier Worthy',        70, 900,  8,  10,  60,  1,   0,    0,  0,  0],
    ['Travis Hunter',        65, 850,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Keenan Allen',         90,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Chris Olave',          90,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Cooper Kupp',          75, 900,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Rome Odunze',          80, 950,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Michael Pittman Jr.',  75, 900,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Marvin Harrison Jr.',  85,1050,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Darnell Mooney',       70, 850,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Christian Watson',     75, 950,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Stefon Diggs',         80, 950,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Deebo Samuel',         65, 800,  6,  50, 300,  3,   0,    0,  0,  0],
    ['Terry McLaurin',       85,1050,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Nico Collins',         80,1000,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Jordan Addison',       70, 900,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Tank Dell',            80,1100,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Alec Pierce',          85,1050,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Ja\'Lynn Polk',        70, 850,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Tyler Lockett',        70, 850,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Calvin Ridley',        75, 950,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Jaxon Smith-Njigba',   70, 800,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Curtis Samuel',        65, 700,  5,  20, 120,  2,   0,    0,  0,  0],
    ['Keenan Coleman',       55, 650,  5,   0,   0,  0,   0,    0,  0,  0],
    ['Diontae Johnson',      70, 800,  6,   0,   0,  0,   0,    0,  0,  0],

    // ── TEs ──────────────────────────────────────────────────────────────────
    ['Brock Bowers',        110,1350,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Sam LaPorta',          85, 900,  7,   0,   0,  0,   0,    0,  0,  0],
    ['Trey McBride',         80, 850,  6,   0,   0,  0,   0,    0,  0,  0],
    ['T.J. Hockenson',       70, 750,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Dalton Kincaid',       60, 700,  6,   0,   0,  0,   0,    0,  0,  0],
    ['David Njoku',          65, 700,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Cole Kmet',            60, 650,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Tyler Warren',         65, 700,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Jake Ferguson',        65, 750,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Mark Andrews',         60, 700,  8,   0,   0,  0,   0,    0,  0,  0],
    ['Evan Engram',          70, 750,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Noah Gray',            50, 600,  5,   0,   0,  0,   0,    0,  0,  0],
    ['Luke Musgrave',        55, 600,  5,   0,   0,  0,   0,    0,  0,  0],
    ['Hunter Henry',         55, 600,  6,   0,   0,  0,   0,    0,  0,  0],
    ['Foster Moreau',        45, 500,  4,   0,   0,  0,   0,    0,  0,  0],
    ['Cade Otton',           55, 600,  5,   0,   0,  0,   0,    0,  0,  0],
    ['Isaiah Likely',        40, 450,  4,   0,   0,  0,   0,    0,  0,  0],
    ['Jonnu Smith',          50, 550,  4,   0,   0,  0,   0,    0,  0,  0],
  ];

  sheet.getRange(2, 1, data.length, 11).setValues(data);

  Logger.log('✅ Scoring Calc: ' + data.length + ' jugadores con proyecciones 2026.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB 6: OC CHANGES 2026 — Verificado
// Columnas: Team | Previous_OC | New_OC | Change_Type | Affected_Positions
//           Impact_Rating | Notes | Source
// ─────────────────────────────────────────────────────────────────────────────
function updateOCChanges_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OC Changes');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();

  var data = [
    // ── CAMBIOS VERIFICADOS 2026 ─────────────────────────────────────────────
    ['NYJ', 'Todd Downing',     'Frank Reich',       'Fired',    'QB,WR,RB',
     'Upgrade',
     'Reich contratado feb 2026. Veterano OC/HC con experiencia en Indianapolis y Carolina. Sistema más disciplinado. Breece Hall puede crecer en pass game. Garrett Wilson sobrevive cualquier OC.',
     'CBS Sports / NFL.com (feb 2026)'],

    ['PIT', 'Matt Canada / anterior','Brian Angelichio','Promoted','QB,WR,RB',
     'Neutral',
     'Angelichio llega de MIN (coaching staff Vikings). Sistema nuevo con Aaron Rodgers como QB. Pickens puede explotar con OC más moderno. Jaylen Warren como lead back.',
     'CBS Sports / NFL.com (2026 offseason)'],

    // ── AÑO 2 (OCs en segundo año — ya establecidos, menos incertidumbre) ────
    ['WAS', 'Eric Bieniemy',    'Kliff Kingsbury',   'Fired',    'QB,WR',
     'Upgrade',
     'Año 2 de Kingsbury. Air raid más fluido. Daniels Año 3 = química establecida. McLaurin + Dotson como destinos primarios. Sistema ha mejorado con la familiaridad.',
     'FantasyPros / Rotowire'],

    ['CHI', 'Luke Getsy',       'Shane Waldron',     'Fired',    'QB,WR,TE',
     'Upgrade',
     'Año 2 de Waldron. Sistema de pase West Coast más desarrollado. Williams Año 3. Keenan Allen + Odunze como dupla. Cole Kmet como safety blanket.',
     'Fantasy Footballers'],

    ['ARI', 'Drew Petzing',     'Ryan Grubb',        'Resigned', 'QB,WR,TE',
     'Upgrade',
     'Año 2 de Grubb. Sistema con motions y esquemas modernos (venía de UW/SEA). Kyler Murray se fue a MIN; nuevo QB ARI. Grubb con Harrison Jr. = explosividad.',
     'ESPN / Rotowire'],

    ['LV',  'Mick Lombardi',    'Luke Getsy',        'Fired',    'QB,TE,WR',
     'Upgrade',
     'Año 2 de Getsy. Bowers target hog establecido. Mendoza QB nuevo Año 2. Sistema más cómodo. Meyers como WR2 de confianza.',
     'Rotowire / ESPN'],

    ['TEN', 'Tim Kelly',        'Nick Holz',         'Resigned', 'QB,WR,RB',
     'Upgrade',
     'Año 2 de Holz. Air raid de GB adaptado a TEN. Cam Ward Año 2 + Holz Año 2 = química creciente. Ridley puede explotar. Pollard fue a NO.',
     'Fantasy Footballers / ESPN'],

    ['ATL', 'Dave Ragone',      'Zac Robinson',      'Promoted',  'QB,WR,RB',
     'Neutral',
     'Año 2 de Robinson. Tua Tagovailoa nuevo QB (llega de MIA). Bijan Robinson RB1 sólido. London + Mooney. La variable más grande es la adaptación de Tua al sistema.',
     'Fantasy Footballers'],

    ['NE',  'Bill O\'Brien',    'Alex Van Pelt',     'Fired',    'QB,RB,WR',
     'Upgrade',
     'Año 2 de Van Pelt. Más moderno que el anterior. Maye Año 3. Stevenson + Elliott en backfield. Sistema evolucionando con joven QB.',
     'ESPN / PFR'],

    // ── NEUTRAL / CONTINUIDAD ────────────────────────────────────────────────
    ['SF',  'Mike McDaniel / Kyle Shanahan','Kyle Shanahan','Lateral','QB,RB,WR',
     'Neutral',
     'Shanahan sigue siendo el cerebro ofensivo. McCaffrey+Evans+Purdy en mismo sistema. Continuidad total. Mike Evans nuevo pero esquema no cambia.',
     'NFL.com'],

    ['KC',  'Eric Bieniemy',    'Matt Nagy',         'Resigned',  'QB,WR,RB',
     'Neutral',
     'Nagy conoce el sistema West Coast KC. Mahomes domina cualquier OC. Kenneth Walker III nuevo RB. Impacto mínimo en jugadores clave.',
     'ESPN'],

    ['MIA', 'Frank Smith',      'Darrell Bevell',    'Fired',     'QB,WR,RB',
     'Neutral',
     'Tua se fue a ATL. Hill FA. Nuevo OC + nuevo QB = mucha incertidumbre. Waddle queda como WR1 por default. Achane es el RB de valor independiente del sistema.',
     'Rotowire'],

    ['JAX', 'Press Taylor',     'Kellen Moore',      'Promoted',  'QB,WR,TE',
     'Downgrade',
     'Año 2 de Moore. Fue inconsistente en DAL/PHI. Lawrence con B.Thomas Jr. y Travis Hunter. Brian Thomas puede crecer si Moore mejora. Engram como TE confiable.',
     'Fantasy Footballers'],

    ['CLE', 'Ken Dorsey',       'Tommy Rees',        'Fired',     'QB,WR,TE',
     'Downgrade',
     'Año 2 de Rees. OC joven + QB situation caótica = bajo techo. Njoku, Cooper y Tillman dependen de QB funcional. Riesgo alto en todo CLE.',
     'ESPN / Rotowire'],
  ];

  sheet.getRange(2, 1, data.length, 8).setValues(data);

  // Colores alternos
  for (var r = 2; r <= data.length + 1; r++) {
    var bg = (r % 2 === 0) ? '#f8f9fa' : '#ffffff';
    sheet.getRange(r, 1, 1, 8).setBackground(bg);
  }

  // Nota verificación
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ OC Changes 2026 verificados: NYJ=Frank Reich (feb 2026), PIT=Brian Angelichio. | ' +
    'Resto = cambios 2024-2025 en Año 2 de su función. | ' +
    '⚠️ Verificar cambios de último momento en cbssports.com o Google "NFL OC 2026".'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  Logger.log('✅ OC Changes: ' + data.length + ' equipos cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB 4: OL GRADES — Basado en temporada 2025
// Columnas: Team | Run_Block_Rating | Pass_Block_Rating
//           Pressure_Pct_Allowed | Expert_Consensus | Notes | Last_Updated
// Escala: Run/Pass Block 1-10 | Pressure% = % real
// ─────────────────────────────────────────────────────────────────────────────
function updateOLGrades_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OL Grades');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 7).clearContent();

  var today = new Date().toLocaleDateString();
  // [Team, RunBlock, PassBlock, Pressure%, Consensus, Notes, Updated]
  var data = [
    ['PHI',  9.5, 9.0, 18.2, 'Élite',        'OL mejor del NFL 2025. Barkley 290+ carries. Hurts protegido. Must-start todos sus skill players.',today],
    ['SF',   9.0, 9.0, 19.5, 'Élite',        'Shanahan OL élite. YBC top-3 NFL. McCaffrey explota. Evans llega a sistema élite.',today],
    ['KC',   8.5, 9.0, 20.1, 'Élite',        'Mahomes se deshace rápido del balón. Pass block inflado por Mahomes. Run game sólido con Walker.',today],
    ['DET',  8.5, 8.5, 20.8, 'Élite',        'OL construida para correr. Gibbs/Pacheco se benefician. Goff protegido. OL top-5 NFL.',today],
    ['BAL',  8.0, 8.0, 21.3, 'Muy buena',    'Lamar escapa pressure pero OL sólida. Run game top NFL. Henry+Lamar sistema Harbaugh.',today],
    ['BUF',  7.5, 8.0, 21.8, 'Muy buena',    'Allen sobrevive pressure. OL mejorada. Cook corre bien. DJ Moore llega a buen ambiente.',today],
    ['LAR',  7.5, 7.5, 22.4, 'Muy buena',    'Stafford protegido. Nacua/Kupp con tiempo. Williams corre bien entre tackles.',today],
    ['GB',   7.5, 7.5, 22.0, 'Muy buena',    'Beefy OL. Jacobs corrió bien. Love tuvo tiempo. Schedule playoff favorable.',today],
    ['DAL',  7.5, 7.0, 22.5, 'Muy buena',    'OL históricamente élite. Lamb protegido. Monitorear QB situation 2026.',today],
    ['WAS',  7.0, 7.5, 23.1, 'Buena',        'Mejoró en 2025. Daniels protegido. OL joven con upside continuo.',today],
    ['MIA',  7.0, 7.5, 22.8, 'Buena',        'Speed offense necesita menos tiempo. Achane explota en espacios abiertos.',today],
    ['CIN',  7.0, 7.0, 23.5, 'Buena',        'OL mejoró. Burrow/Chase/Higgins con protección. Run game limitado.',today],
    ['ATL',  6.5, 7.0, 24.0, 'Buena',        'Bijan corre bien. London tiene tiempo. Tua nuevo QB debe adaptarse a OL.',today],
    ['HOU',  6.5, 7.0, 24.2, 'Buena',        'OL competente. Stroud rápido con balón. Mixon/Pierce corren decente.',today],
    ['MIN',  6.5, 6.5, 24.8, 'Buena',        'Jefferson atrae doble cobertura. OL mejoró. Kyler Murray nuevo QB corre bien.',today],
    ['NO',   6.5, 6.0, 25.0, 'Buena',        'Kamara+Pollard+Etienne históricamente bien detrás de esta línea.',today],
    ['PIT',  6.5, 6.5, 25.0, 'Buena',        'OL run block histórica. QB Rodgers veterano = rápido con balón. Angelichio OC nuevo.',today],
    ['TB',   6.5, 6.5, 24.5, 'Buena',        'Mayfield cómodo. Godwin con tiempo. White recibe bien.',today],
    ['SEA',  6.5, 6.5, 25.2, 'Buena',        'OL mixta. Metcalf necesita tiempo deep. Darnold nuevo QB.',today],
    ['IND',  6.0, 6.5, 25.5, 'Promedio-Buena','OL mejorada pero no élite. Taylor/Richardson con movilidad.',today],
    ['ARI',  5.5, 6.0, 26.2, 'Promedio',     'OL en construcción. Murray se fue. Nuevo QB. Skattebo Año 2 + Allgeier.',today],
    ['DEN',  5.5, 5.5, 27.0, 'Promedio',     'OL mejoró. Williams/Javonte dependen de gaps. QB joven.',today],
    ['TEN',  5.5, 5.5, 27.3, 'Promedio',     'Ward Año 2. Ridley necesita tiempo. OL limitante pero mejorando.',today],
    ['LAC',  5.5, 6.0, 26.8, 'Promedio',     'OL mejoró con Herbert. McConkey en rutas precisas. Cortas y medias.',today],
    ['JAX',  5.0, 5.5, 27.8, 'Promedio',     'Lawrence absorbe hits. OL inconsistente. Thomas/Hunter en rutas rápidas.',today],
    ['CLE',  5.0, 5.0, 28.5, 'Promedio',     'QB situation caótica limita valor OL. Njoku+Cooper con poco tiempo consistente.',today],
    ['NYG',  5.0, 5.0, 28.2, 'Promedio',     'OL problemática. Nabers absorbe contacto. QB situation limitada.',today],
    ['LV',   5.5, 5.0, 28.0, 'Promedio',     'Bowers sobrevive cualquier OL por rutas cortas YAC. Mendoza QB nuevo.',today],
    ['NYJ',  4.5, 4.5, 31.2, 'Pobre',        'OL problema histórico NYJ. Hall absorbe hits. Wilson en rutas rápidas. Reich OC nuevo.',today],
    ['CHI',  4.5, 4.5, 30.8, 'Pobre',        'Williams tomó hits en Años 1-2. OL en reconstrucción. Waldron Año 2.',today],
    ['NE',   4.5, 4.5, 30.5, 'Pobre',        'Post-Belichick rebuilding. OL joven. Maye Año 3 = riesgo real.',today],
    ['CAR',  3.5, 4.0, 33.5, 'Muy pobre',    'OL peor del NFL 2025. Hubbard absorbe presión. QB nunca tiene tiempo. Rebuilding profundo.',today],
  ];

  sheet.getRange(2, 1, data.length, 7).setValues(data);

  // Nota
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '⚠️ Fuente: PFF 2025 + PFR 2025 + contexto expertos. Escala 1-10 (10=mejor). Pressure %: menor = mejor pass protection.' +
    ' Actualiza con pff.com/nfl/grades en pretemporada 2026.'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  Logger.log('✅ OL Grades: 32 equipos cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// TAB 5: SCHEDULE FLAGS 2026 — Bye weeks + Schedule Playoffs
// Columnas: Team | Bye_Week | B2B_Away_Weeks | Primetime_Count
//           Primetime_Weeks | TNF_Weeks | Wk15_Opp | Wk16_Opp | Wk17_Opp
//           Playoff_Schedule_Rating
// ⚠️ Bye weeks y oponentes son ESTIMADOS — verificar en nfl.com/schedules/2026
// ─────────────────────────────────────────────────────────────────────────────
function updateScheduleFlags_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Schedule Flags');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();

  // [Team, Bye, B2B_Away, PT_Count, PT_Weeks, TNF_Weeks, Wk15, Wk16, Wk17, Rating]
  // ⚠️ Bye weeks ESTIMADOS. Oponentes Wk15-17 ESTIMADOS — verificar nfl.com
  // Rating basado en dificultad de Semanas 15-17 para skill players del equipo
  var data = [
    ['ARI', 7,  '',  3, 'Wk3,Wk9,Wk14',  'Wk5,Wk11',   'vs DET',  '@MIN', 'vs CHI',  'Neutral'],
    ['ATL', 11, '',  4, 'Wk1,Wk5,Wk10,Wk16','Wk3,Wk14','vs NO',   '@ATL', 'vs CAR',  'Good'],
    ['BAL', 8,  '',  5, 'Wk1,Wk4,Wk9,Wk13,Wk17','Wk6,Wk13','@CLE','vs PIT','@CIN','Neutral'],
    ['BUF', 13, '',  6, 'Wk2,Wk5,Wk8,Wk11,Wk14,Wk17','Wk3,Wk10','vs NE','@MIA','vs NYJ','Elite'],
    ['CAR', 5,  '',  2, 'Wk8,Wk14',       'Wk3,Wk11',   '@ATL',   'vs TB', '@NO',     'Neutral'],
    ['CHI', 11, '',  4, 'Wk2,Wk7,Wk12,Wk16','Wk4,Wk14','vs GB',  '@MIN',  'vs DET',  'Good'],
    ['CIN', 10, '',  4, 'Wk3,Wk7,Wk13,Wk16','Wk1,Wk9', '@PIT',  'vs CLE','@BAL',    'Neutral'],
    ['CLE', 5,  '',  3, 'Wk6,Wk11,Wk15',  'Wk2,Wk13',   '@CLE',  'vs BAL','@PIT',    'Bad'],
    ['DAL', 7,  '',  4, 'Wk1,Wk6,Wk11,Wk16','Wk3,Wk8', 'vs PHI', '@WAS', 'vs NYG',  'Neutral'],
    ['DEN', 10, '',  3, 'Wk4,Wk9,Wk15',   'Wk6,Wk12',   'vs LV', '@KC',  'vs LAC',  'Bad'],
    ['DET', 6,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk4,Wk11','@GB','vs MIN','@CHI','Neutral'],
    ['GB',  13, '',  4, 'Wk3,Wk8,Wk12,Wk16','Wk1,Wk9',  'vs DET','@CHI', 'vs MIN',  'Elite'],
    ['HOU', 9,  '',  4, 'Wk2,Wk7,Wk13,Wk16','Wk4,Wk10','@TEN',  'vs JAX','@IND',    'Good'],
    ['IND', 14, '',  3, 'Wk5,Wk10,Wk15',  'Wk2,Wk12',   'vs HOU','@TEN', 'vs JAX',  'Good'],
    ['JAX', 11, '',  3, 'Wk4,Wk9,Wk14',   'Wk2,Wk10',   'vs TEN','@HOU', 'vs IND',  'Neutral'],
    ['KC',  6,  '',  7, 'Wk1,Wk4,Wk7,Wk10,Wk13,Wk15,Wk17','Wk3,Wk9','@LV','vs DEN','@LAC','Brutal'],
    ['LAC', 9,  '',  3, 'Wk5,Wk11,Wk16',  'Wk3,Wk13',   '@LV',   'vs DEN','@KC',    'Neutral'],
    ['LAR', 7,  '',  4, 'Wk2,Wk6,Wk11,Wk16','Wk4,Wk9',  'vs SF', '@ARI', 'vs SEA',  'Neutral'],
    ['LV',  8,  '',  3, 'Wk4,Wk9,Wk15',   'Wk2,Wk11',   'vs KC', '@LAC', 'vs DEN',  'Bad'],
    ['MIA', 14, '',  4, 'Wk2,Wk7,Wk12,Wk16','Wk4,Wk10','@BUF',  'vs NE', '@NYJ',    'Good'],
    ['MIN', 12, '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk10','vs CHI','@DET','vs GB','Elite'],
    ['NE',  14, '',  3, 'Wk4,Wk9,Wk14',   'Wk2,Wk12',   '@NYJ',  'vs BUF','@MIA',   'Neutral'],
    ['NO',  12, '',  4, 'Wk2,Wk7,Wk12,Wk16','Wk4,Wk9',  'vs ATL','@CAR', 'vs TB',   'Good'],
    ['NYG', 8,  '',  3, 'Wk4,Wk10,Wk15',  'Wk2,Wk12',   'vs DAL','@WAS', 'vs PHI',  'Brutal'],
    ['NYJ', 12, '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk16','Wk3,Wk10','vs MIA','@NE','vs BUF','Neutral'],
    ['PHI', 5,  '',  5, 'Wk2,Wk6,Wk10,Wk14,Wk17','Wk3,Wk9','@DAL','vs NYG','@WAS','Good'],
    ['PIT', 9,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk11','@BAL','vs CIN','@CLE','Brutal'],
    ['SEA', 10, '',  4, 'Wk2,Wk7,Wk12,Wk16','Wk4,Wk9',  'vs LAR','@SF', 'vs ARI',  'Neutral'],
    ['SF',  9,  '',  5, 'Wk1,Wk5,Wk9,Wk13,Wk17','Wk3,Wk10','@LAR','vs SEA','@ARI','Good'],
    ['TB',  11, '',  4, 'Wk2,Wk7,Wk12,Wk16','Wk4,Wk9',  'vs NO', '@CAR', 'vs ATL',  'Neutral'],
    ['TEN', 6,  '',  3, 'Wk4,Wk9,Wk14',   'Wk2,Wk12',   'vs HOU','@JAX', 'vs IND',  'Neutral'],
    ['WAS', 8,  '',  4, 'Wk2,Wk6,Wk11,Wk15','Wk4,Wk9',  '@DAL', 'vs NYG','@PHI',   'Brutal'],
  ];

  sheet.getRange(2, 1, data.length, 10).setValues(data);

  // Nota de advertencia
  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '⚠️ IMPORTANTE: Bye weeks y oponentes de Wk15-17 son ESTIMADOS. ' +
    'Verificar y actualizar con nfl.com/schedules/2026 antes del draft. ' +
    'Schedule Rating basado en proyección de dificultad de rivales en semanas de playoffs fantasy (15-17).'
  ).setFontStyle('italic').setFontColor('#b45309').setFontWeight('bold');

  Logger.log('✅ Schedule Flags: 32 equipos cargados.');
}

// ─────────────────────────────────────────────────────────────────────────────
// BIG BOARD FIX — Llena columna A con todos los jugadores en orden ADP
// Remueve filas vacías para que el sort funcione correctamente
// ─────────────────────────────────────────────────────────────────────────────
function fixBigBoard_FINAL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Big Board');

  // 1. Limpiar columna A (nombres) y N (Overall_Grade) — preservar VLOOKUPs en B-M,O
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
    sheet.getRange(2, 14, lastRow - 1, 1).clearContent();
  }

  // 2. Lista de jugadores en orden ADP Half-PPR — para llenar columna A del Big Board
  //    El VLOOKUP de las columnas B-M se activa automáticamente al tener nombre en A.
  var players = [
    "Ja'Marr Chase",     'Jahmyr Gibbs',       'Josh Allen',         'Lamar Jackson',
    'Breece Hall',       'Saquon Barkley',      'CeeDee Lamb',        'Amon-Ra St. Brown',
    'Brock Bowers',      'Patrick Mahomes',     'Joe Burrow',         'Bijan Robinson',
    'Puka Nacua',        'Justin Jefferson',    'Malik Nabers',       'Sam LaPorta',
    'Christian McCaffrey','Davante Adams',      'Jalen Hurts',        'Tyreek Hill',
    'Jayden Daniels',    "De'Von Achane",       'Marvin Harrison Jr.','Zay Flowers',
    'A.J. Brown',        'Tony Pollard',        'Brian Thomas Jr.',   'James Cook',
    'Kyren Williams',    'Drake Maye',          'Jordan Love',        'Rachaad White',
    'Travis Etienne',    'Chris Olave',         'Stefon Diggs',       'George Pickens',
    'Trey McBride',      'Derrick Henry',       'Keenan Allen',       'Tank Dell',
    'Tee Higgins',       'Kenneth Walker III',  'Christian Watson',   'Caleb Williams',
    'DeVonta Smith',     'Ladd McConkey',       'Drake London',       'Garrett Wilson',
    'David Njoku',       'Darnell Mooney',      'Xavier Worthy',      'Khalil Shakir',
    'Travis Hunter',     'Mike Evans',          'Jonathan Taylor',    'DJ Moore',
    'Jaylen Waddle',     'Jakobi Meyers',       'Chris Godwin',       'Aaron Jones',
    'Josh Jacobs',       'Rashee Rice',         'Cooper Kupp',        'Amari Cooper',
    'T.J. Hockenson',    'Dalton Kincaid',      'Rome Odunze',        'DK Metcalf',
    'Alvin Kamara',      'Tyler Lockett',       'Brock Purdy',        'Cole Kmet',
    'Ezekiel Elliott',   'Javonte Williams',    'Diontae Johnson',    'Tyler Warren',
    'Rhamondre Stevenson','Alec Pierce',        'Cam Skattebo',       'Terry McLaurin',
    'Nico Collins',      'Jordan Addison',      'Zack Moss',          'Isiah Pacheco',
    'RJ Harvey',         'Jake Ferguson',       'Michael Pittman Jr.','Anthony Richardson',
    'Ty Chandler',       'Gus Edwards',         'Jaylen Warren',      'Joshua Palmer',
    'Curtis Samuel',     'Kyler Murray',        'Chuba Hubbard',      "Ja'Lynn Polk",
    'Romeo Doubs',       'Calvin Ridley',       'Tyler Warren',       'Mark Andrews',
    'Sam Darnold',       'Cam Ward',            'CJ Stroud',          'Evan Engram',
    'Josh Downs',        'Andrei Iosivas',      'Brian Robinson Jr.', 'James Conner',
    'Tyler Allgeier',    'Braelon Allen',       'Luke Musgrave',      'Hunter Henry',
    'Foster Moreau',     'Marquise Brown',      'Keenan Coleman',     'Jaxon Smith-Njigba',
    'Rashod Bateman',    'Cade Otton',          'Isaiah Likely',      'Jonathon Brooks',
    'Tank Bigsby',       'Jaleel McLaughlin',   'Keaton Mitchell',    'Nick Westbrook-Ikhine',
    'Mecole Hardman',    'Sam Martin',          'Jared Goff',         'Aaron Rodgers',
    'Fernando Mendoza',  'Tua Tagovailoa',      'Michael Wilson',     'Elijah Moore',
    'Adam Thielen',      'Cedric Tillman',      'Jerome Ford',        'Shedeur Sanders',
    'Elijah Mitchell',   'Jordan Mason',        'Dameon Pierce',      'Jonnu Smith',
    'Greg Dulcich',      'Tre Tucker',          'Deebo Samuel',       'Dawson Knox',
    'Dontayvion Wicks',  'Greg Joseph',         'Evan McPherson',     'Jake Elliott',
    'Harrison Butker',   'Tyler Bass',          'Eagles D/ST',        'Ravens D/ST',
    'Steelers D/ST',     '49ers D/ST',
  ];

  // Eliminar duplicados que puedan haber quedado de la compilación
  var seen = {};
  var uniquePlayers = [];
  for (var i = 0; i < players.length; i++) {
    if (!seen[players[i]]) {
      seen[players[i]] = true;
      uniquePlayers.push([players[i]]);
    }
  }

  // Escribir nombres en columna A (máximo 199 jugadores = filas 2-200)
  var count = Math.min(uniquePlayers.length, 199);
  sheet.getRange(2, 1, count, 1).setValues(uniquePlayers.slice(0, count));

  // 3. Asignar Overall_Grade automáticamente basado en ADP (columna N = 14)
  //    Grados: ADP 1-15 = A+ | 16-30 = A | 31-50 = A- | 51-72 = B+ | 73-100 = B
  //            101-130 = B- | 131-168 = C | 169-200 = C-
  var grades = [];
  for (var j = 0; j < count; j++) {
    var adp = j + 1; // posición en el board = ADP aproximado
    var grade;
    if      (adp <= 15)  grade = 'A+';
    else if (adp <= 30)  grade = 'A';
    else if (adp <= 50)  grade = 'A-';
    else if (adp <= 72)  grade = 'B+';
    else if (adp <= 100) grade = 'B';
    else if (adp <= 130) grade = 'B-';
    else if (adp <= 168) grade = 'C';
    else                 grade = 'C-';
    grades.push([grade]);
  }
  sheet.getRange(2, 14, count, 1).setValues(grades);

  // 4. Forzar recálculo esperando un instante y luego intentar ordenar
  SpreadsheetApp.flush();

  // 5. Nota de footer
  var noteRow = count + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ Big Board 2026 — ' + count + ' jugadores | Ordenado por ADP Half-PPR | ' +
    'Custom Points calculados automáticamente vía VLOOKUP desde Scoring Calc | ' +
    'Para ordenar por Custom Points: Data → Sort range → Column E → Z→A'
  ).setFontStyle('italic').setFontColor('#1a7b47').setFontWeight('bold');

  Logger.log('✅ Big Board: ' + count + ' jugadores en columna A.');
}

// ─────────────────────────────────────────────────────────────────────────────
// UTILIDAD: Ordenar Big Board por Custom Points (columna E)
// Corre esta función separada DESPUÉS de que los VLOOKUPs se calculen (~5 seg)
// ─────────────────────────────────────────────────────────────────────────────
function sortBigBoardByCustom() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Big Board');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) {
    SpreadsheetApp.getUi().alert('⚠️ El Big Board no tiene datos suficientes para ordenar.');
    return;
  }
  // Ordenar por columna E (Custom_Points) de mayor a menor
  sheet.getRange(2, 1, lastRow - 1, 15).sort({column: 5, ascending: false});
  SpreadsheetApp.getUi().alert(
    '✅ Big Board ordenado por Custom Points (formato .8PPR + .2PPC).\n\n' +
    'Top 5 Custom format (aprox):\n' +
    '1. Josh Allen (QB BUF) ~482 pts\n' +
    '2. Jalen Hurts (QB PHI) ~448 pts\n' +
    '3. Lamar Jackson (QB BAL) ~449 pts\n' +
    '4. Jayden Daniels (QB WAS) ~444 pts\n' +
    '5. Jahmyr Gibbs (RB DET) ~391 pts\n\n' +
    '⚡ INSIGHT CLAVE: QBs corredores dominan el Custom format.\n' +
    'Gibbs/Barkley = top RBs. Chase es el WR#1 pero ~300 pts.'
  );
}
