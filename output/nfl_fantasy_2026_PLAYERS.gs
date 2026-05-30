/**
 * NFL Fantasy 2026 — PLAYERS DB CORREGIDO
 * Fuente: FantasyPros Half-PPR ADP verificado directo (mayo 2026)
 * Corre esta función: updatePlayersDB_2026_VERIFIED()
 *
 * CORRECCIONES MAYORES vs script anterior:
 *  Javonte Williams → DAL (no DEN)
 *  Cam Skattebo     → NYG (no ARI)
 *  George Pickens   → DAL (no PIT)
 *  DK Metcalf       → PIT (no SEA)
 *  Jaylen Waddle    → DEN (no MIA)
 *  Davante Adams    → LAR (no GB)
 *  Jordan Mason     → MIN (no SF)
 *  Rachaad White    → WAS (no TB)
 *  Jakobi Meyers    → JAC (no LV)
 *  Isaiah Likely    → NYG (no BAL)
 *  Jaxon Smith-Njigba → SEA ADP 5.3 (no ADP 144!)
 *  Ashton Jeanty    → LV ADP 13 (nuevo - no estaba)
 *  Omarion Hampton  → LAC ADP 17.3 (nuevo - no estaba)
 *  Chase Brown      → CIN ADP 19.3 (nuevo)
 *  Colston Loveland → CHI ADP 36 (nuevo)
 */

function updatePlayersDB_2026_VERIFIED() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Players DB');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 12).clearContent();

  // Columnas: Player | Position | Team | Tier | ADP_FP | ADP_UD
  //           Expert_Notes | OC_Change | B2B_Weeks | Primetime_Weeks
  //           Playoff_SoS | OL_Grade
  // Tier: 1=ADP 1-30 | 2=ADP 31-72 | 3=ADP 73-120 | 4=ADP 121-168 | 5=169+
  // Fuente ADP: fantasypros.com Half-PPR mayo 2026

  var data = [
    // ── TIER 1 — ADP 1-30 ──────────────────────────────────────────────────
    ['Bijan Robinson',        'RB','ATL','1',  1.0,  1.1, 'RB1 overall. ATL OL sólida. Tua Tagovailoa nuevo QB (de MIA). Kyle Pitts TE. London+Mooney WRs. 260 carries + 70 rec + 12 TDs.','Yes','','','Neutral','Buena'],
    ['Jahmyr Gibbs',          'RB','DET','1',  2.0,  2.2, 'Lead back solo DET. Pacheco backup (ADP 146). 280 carries + 65 rec. Año 4. OL élite DET. ARSB comparte ofensiva.','No','','','Neutral','Élite'],
    ["Ja'Marr Chase",         'WR','CIN','1',  3.0,  3.1, 'WR1 overall. Burrow connection élite. 100+ rec + 1500 yds + 12 TDs. Piso altísimo semana a semana. OL CIN mejoró.','No','','','Easy','Buena'],
    ['Puka Nacua',            'WR','LAR','1',  4.0,  4.2, 'WR1 LAR. Adams llegó de GB. Stafford QB. 100 rec + 1300 yds + 9 TDs. OL LAR muy buena. Must-start semanal.','No','','','Neutral','Muy buena'],
    ['Jaxon Smith-Njigba',    'WR','SEA','1',  5.3,  5.5, 'WR1 SEA. Año 3 (clase 2024). Explosión total. DK Metcalf se fue a PIT. Sam Darnold QB. 105 rec + 1300 yds proyectados.','No','','','Neutral','Buena'],
    ['Christian McCaffrey',   'RB', 'SF','1',  6.0,  6.2, 'Si está sano = RB1 overall. Shanahan system élite. Mike Evans WR. Kittle TE. Riesgo salud es el único factor. Guerendo handcuff.','No','','','Neutral','Élite'],
    ['Jonathan Taylor',       'RB','IND','1',  7.0,  7.3, 'RB IND. Alec Pierce WR extendido. Daniel Jones QB. 250 carries + 55 rec. OL mejorada. Año 6. Muy confiable.','No','','','Easy','Promedio-Buena'],
    ['Amon-Ra St. Brown',     'WR','DET','1',  7.7,  8.0, 'Slot élite. 120+ recepciones. OL DET élite. Gibbs comparte ofensiva. Goff QB. Año 6. Piso altísimo.','No','','','Neutral','Élite'],
    ['Justin Jefferson',      'WR','MIN','1', 10.0, 10.3, 'WR élite. Kyler Murray QB nuevo (de ARI). Murray-Jefferson = explosividad. Jordan Love se fue? Monitorear contexto MIN.','Yes','','','Easy','Buena'],
    ['James Cook',            'RB','BUF','1', 10.7, 11.0, 'RB BUF. Allen corre también pero Cook 200+ carries. 70 rec. DJ Moore WR1. BUF ofensiva élite. OL muy buena.','No','','','Easy','Muy buena'],
    ['CeeDee Lamb',           'WR','DAL','1', 11.3, 11.5, 'WR1 DAL. Dak Prescott QB regresó. Pickens + Williams también en DAL. 110 rec + 1450 yds + 10 TDs. OL élite.','No','','','Neutral','Muy buena'],
    ['Ashton Jeanty',         'RB', 'LV','1', 13.0, 13.3, 'Año 2 (clase 2025). RB LV. Breakout esperado. Bowers TE1 domina la ofensiva. Mendoza QB. OC Getsy segundo año.','Yes','','','Neutral','Promedio'],
    ["De'Von Achane",         'RB','MIA','1', 13.0, 13.4, 'Speed back élite. MIA nuevo QB (Malik Willis). Tua se fue a ATL. 185 carries + 80 rec. Explosividad única.','Yes','','','Neutral','Buena'],
    ['Saquon Barkley',        'RB','PHI','1', 14.7, 15.0, 'OL élite PHI. Fangio DC = defensiva top-3 NFL = leads constantes = Barkley carries en Q4. 290 carries + 1400 yds + 14 TDs. Must-start toda la temporada.','No','','','Easy','Élite'],
    ['Trey McBride',          'TE','ARI','1', 16.7, 17.0, 'TE1 overall. ARI OL promedio pero McBride target hog. Grubb OC. Harrison Jr. WR1. 90 rec + 1050 yds + 7 TDs.','Yes','','','Neutral','Promedio'],
    ['Brock Bowers',          'TE', 'LV','1', 17.3, 17.5, 'TE2 overall. 110 rec proyectadas. Mendoza QB Año 2. OC Getsy prioriza Bowers. Target hog élite. Jeanty nuevo RB.','Yes','','','Neutral','Promedio'],
    ['Omarion Hampton',       'RB','LAC','1', 17.3, 17.6, 'Año 2 (clase 2025). RB LAC. Herbert QB. McConkey WR1. Gadsden TE. 200 carries + 60 rec proyectados. Upside enorme.','No','','','Easy','Promedio'],
    ['Kenneth Walker III',    'RB', 'KC','1', 18.7, 19.0, 'RB KC. Mahomes QB. Rice WR1, Worthy WR2. 190 carries + 55 rec. Sistema Reid. Hard schedule en playoffs.','No','','','Hard','Élite'],
    ['Chase Brown',           'RB','CIN','1', 19.3, 19.6, 'RB CIN. Burrow QB. Chase+Higgins WRs. 220 carries + 55 rec. OL mejorada CIN. Año 3 si fue clase 2024.','No','','','Easy','Buena'],
    ['Drake London',          'WR','ATL','1', 19.7, 20.0, 'WR1 ATL. Tua nuevo QB. Bijan RB1. Kyle Pitts TE1. London = objetivo principal de pase. 90 rec + 1100 yds + 8 TDs.','Yes','','','Neutral','Buena'],
    ['Derrick Henry',         'RB','BAL','1', 20.3, 20.6, 'Veterano BAL. Lamar también corre. Henry = red zone machine. 230 carries + 15 TDs. Año 12. Confiable.','No','','','Neutral','Muy buena'],
    ['Josh Allen',            'QB','BUF','1', 20.7, 21.0, 'Elite dual-threat QB. Brady HC + Carmichael OC (15 años con Brees/NO) = sistema sólido. Schedule playoff Elite. DJ Moore WR1. ADP 21 = mayor valor Custom del draft.','No','','','Easy','Muy buena'],
    ['Jeremiyah Love',        'RB','ARI','1', 22.7, 23.0, 'Año 2 (clase 2025). RB ARI. Grubb OC. Allgeier también en backfield. Harrison Jr. WR1. Breakout potencial Year 2.','Yes','','','Neutral','Promedio'],
    ['Nico Collins',          'WR','HOU','1', 23.7, 24.0, 'WR1 HOU. CJ Stroud QB. Dell + Higgins como WRs adicionales. 90 rec + 1150 yds + 9 TDs proyectados.','No','','','Neutral','Buena'],
    ['George Pickens',        'WR','DAL','1', 24.0, 24.3, 'WR DAL. Llegó a DAL. Dak Prescott QB. CeeDee Lamb WR1. Pickens WR2 = upside enorme. Javonte Williams también en DAL.','No','','','Neutral','Muy buena'],
    ['Malik Nabers',          'WR','NYG','1', 25.7, 26.0, '🔼 John Harbaugh HC + Matt Nagy OC = mayor upgrade NYG en décadas. Nabers WR1 en sistema de élite. Dart Año 2 + Harbaugh = breakout candidate y stack a considerar.','No','','','Neutral','Promedio'],
    ['A.J. Brown',            'WR','PHI','1', 28.7, 29.0, 'WR élite PHI. OL élite. Hurts QB. Smith WR2. Brown atrae doble cobertura. 95 rec + 1350 yds + 9 TDs.','No','','','Easy','Élite'],
    ['Rashee Rice',           'WR', 'KC','1', 28.7, 29.1, 'WR1 KC. Mahomes. Regresó de lesión. 90 rec + 1100 yds + 9 TDs. Worthy como WR2 deep threat.','No','','','Hard','Élite'],
    ['Chris Olave',           'WR', 'NO','1', 29.3, 29.6, 'WR1 NO. Tyler Shough QB. Etienne + Kamara en backfield. Olave = objetivo principal. 90 rec + 1100 yds + 7 TDs.','No','','','Easy','Buena'],

    // ── TIER 2 — ADP 31-72 ─────────────────────────────────────────────────
    ['Josh Jacobs',           'RB', 'GB','2', 31.7, 32.0, 'RB1 GB. Jordan Love QB. Christian Watson WR. Tucker Kraft TE. 240 carries + 60 rec. OL muy buena GB. Schedule élite.','No','','','Easy','Muy buena'],
    ['Kyren Williams',        'RB','LAR','2', 32.0, 32.3, 'RB1 LAR. Stafford QB. Nacua + Adams WRs. 250 carries + 55 rec + 1200 yds. OL muy buena LAR.','No','','','Neutral','Muy buena'],
    ['Breece Hall',           'RB','NYJ','2', 33.0, 33.4, 'Talento élite. Frank Reich nuevo OC 2026. Wilson WR1. OL NYJ pobre. 230 carries + 80 rec si OL mejora.','Yes','','','Neutral','Pobre'],
    ['Travis Etienne',        'RB', 'NO','2', 35.3, 35.6, 'RB NO. Llegó de JAX. Kamara también en backfield. Tyler Shough QB. Monitorear reparto. NO run-friendly históricamente.','No','','','Easy','Buena'],
    ['Javonte Williams',      'RB','DAL','2', 35.7, 36.0, 'RB DAL. Llegó a DAL. Pickens también en DAL. Dak Prescott QB. 200 carries + 45 rec. OL élite DAL.','No','','','Neutral','Muy buena'],
    ['Colston Loveland',      'TE','CHI','2', 36.0, 36.3, 'Año 2. TE CHI. Ben Johnson HC + Press Taylor OC = mejor sistema posible. Caleb Williams QB Año 3. Odunze + Burden WRs. Loveland safety blanket. Breakout TE potencial.','No','','','Easy','Pobre'],
    ['Tee Higgins',           'WR','CIN','2', 37.3, 37.6, 'WR2 CIN. Burrow tiene volumen para Chase Y Higgins. 80 rec + 1050 yds + 8 TDs. Año 7. Muy confiable.','No','','','Easy','Buena'],
    ['DeVonta Smith',         'WR','PHI','2', 39.0, 39.3, 'WR2 PHI. OL élite. Brown atrae cobertura = Smith abierto. 85 rec + 1100 yds + 7 TDs. Año 6.','No','','','Easy','Élite'],
    ['Emeka Egbuka',          'WR', 'TB','2', 39.3, 39.6, 'Año 2 (clase 2025). WR1 TB. Baker Mayfield QB. Godwin veterano como complemento. Breakout Year 2 esperado.','No','','','Easy','Buena'],
    ['Tetairoa McMillan',     'WR','CAR','2', 39.3, 39.7, 'Año 2 (clase 2025). WR1 CAR. Bryce Young QB. Reconstrucción total pero McMillan = talento élite. OL pobre = risk.','No','','','Neutral','Muy pobre'],
    ['Zay Flowers',           'WR','BAL','2', 41.7, 42.0, 'WR1 BAL. ⚠️ Minter HC + Doyle OC Año 1 = incertidumbre en sistema con Lamar. Flowers sobrevive por talento. Ceiling limitado hasta que el nuevo sistema se establezca.','No','','','Neutral','Muy buena'],
    ['Cam Skattebo',          'RB','NYG','2', 42.0, 42.3, 'Año 2. RB NYG. Harbaugh HC + Nagy OC = sistema mucho mejor. Skattebo workhorse con upside real. Nabers WR1. OL pobre NYG sigue siendo el riesgo.','No','','','Neutral','Promedio'],
    ['Lamar Jackson',         'QB','BAL','2', 43.0, 43.4, '⚠️ Jesse Minter HC Año 1 + Declan Doyle OC Año 1 = cambio total de sistema en BAL. FP proyecta conservador (344 pts). Si el sistema funciona = élite. Riesgo real en ADP 43.','No','','','Neutral','Muy buena'],
    ['Garrett Wilson',        'WR','NYJ','2', 43.3, 43.6, 'WR élite NYJ. Frank Reich nuevo OC 2026. 100 rec + 1200 yds. Talento demasiado bueno para fallar con cualquier OC.','Yes','','','Neutral','Pobre'],
    ['DJ Moore',              'WR','BUF','2', 46.0, 46.3, 'WR1 BUF. Allen QB élite. 90 rec + 1100 yds + 8 TDs. Shakir complementa en slot. Gran fit con Allen.','No','','','Easy','Muy buena'],
    ['Ladd McConkey',         'WR','LAC','2', 46.7, 47.0, 'WR1 LAC. Herbert QB. 95 rec + 1100 yds. Año 3 (clase 2024). Hampton RB nuevo. Sistema preciso.','No','','','Easy','Promedio'],
    ['Bucky Irving',          'RB', 'TB','2', 48.3, 48.6, 'RB1 TB. Baker Mayfield QB. Egbuka WR1. 210 carries + 60 rec proyectados. Año 3 si fue clase 2024. Sólido.','No','','','Easy','Buena'],
    ['Joe Burrow',            'QB','CIN','2', 48.7, 49.0, 'Regresó 100% de lesión. Chase+Higgins dupla élite. 4400 yds + 35 TDs. Año 7. ADP 49 = valor enorme.','No','','','Easy','Buena'],
    ['Quinshon Judkins',      'RB','CLE','2', 49.0, 49.3, 'Año 2 (clase 2025). RB CLE. QB situation caótica. Pero Judkins = workhorse garantizado. 210 carries. Fannin Jr. TE.','Yes','','','Easy','Promedio'],
    ['David Montgomery',      'RB','HOU','2', 50.0, 50.3, 'RB HOU. Llegó de DET. CJ Stroud QB. Collins + Dell WRs. 200 carries + 45 rec. OL buena HOU.','No','','','Neutral','Buena'],
    ['Luther Burden III',     'WR','CHI','2', 51.0, 51.3, 'Año 2 (clase 2025). WR CHI. Williams Año 3. Waldron OC Año 2. Slot receiver explosivo. 80 rec + 1000 yds potencial.','No','','','Easy','Pobre'],
    ['Tyler Warren',          'TE','IND','2', 51.0, 51.3, 'Año 2 (clase 2025). TE IND. Daniel Jones QB. Pierce WR1. 65 rec + 700 yds + 6 TDs. Breakout candidato.','No','','','Easy','Promedio-Buena'],
    ['TreVeyon Henderson',    'RB', 'NE','2', 52.3, 52.6, 'Año 2. RB NE. Drake Maye QB Año 3. Vrabel HC + McDaniels OC (conoce NE como nadie). 180 carries + 55 rec proyectados. Upside alto con sistema mejorado.','Yes','','','Neutral','Pobre'],
    ['Davante Adams',         'WR','LAR','2', 52.7, 53.0, 'Veterano WR LAR. Llegó de GB. Stafford QB. Nacua WR1. Adams como WR2 = sigue produciendo. 85 rec + 1000 yds.','No','','','Neutral','Muy buena'],
    ['Mike Evans',            'WR', 'SF','2', 53.3, 53.6, 'WR SF. Shanahan system. Kittle TE. Purdy QB. McCaffrey corre = espacio para Evans. 80 rec + 1000 yds + 9 TDs.','No','','','Neutral','Élite'],
    ['Jadarian Price',        'RB','SEA','2', 55.0, 55.3, 'Año 2 (clase 2025). RB SEA. Sam Darnold QB. JSN WR1. 180 carries + 50 rec. Metcalf se fue a PIT.','No','','','Neutral','Buena'],
    ['Jameson Williams',      'WR','DET','2', 56.0, 56.3, 'WR2 DET. Goff QB. ARSB WR1. Williams deep threat explosivo. 70 rec + 950 yds + 8 TDs. OL DET élite.','No','','','Neutral','Élite'],
    ['Jaylen Waddle',         'WR','DEN','2', 56.0, 56.3, 'WR DEN. Llegó de MIA. Bo Nix QB. Sutton también en DEN. 80 rec + 950 yds proyectados. Sistema diferente a MIA.','No','','','Neutral','Promedio'],
    ['Terry McLaurin',        'WR','WAS','2', 56.3, 56.6, 'WR1 WAS. Daniels Año 3. Kingsbury OC Año 2 (air raid). 85 rec + 1050 yds + 8 TDs. Confiable como WR2.','No','','','Hard','Buena'],
    ['Drake Maye',            'QB', 'NE','2', 57.0, 57.3, '🔼 Año 3. Vrabel HC + Josh McDaniels OC (construyó el sistema Brady/Belichick en NE). Maye breakout candidate real. Henderson RB. FP proyecta 4394 pass yds (real 2025).','Yes','','','Neutral','Pobre'],
    ['Jayden Daniels',        'QB','WAS','2', 57.3, 57.6, '⚠️ Dan Quinn HC + David Blough OC (ex-QB backup, Año 1 sin historial como OC) = incertidumbre en sistema. Daniels tiene talento dual-threat pero el OC es el riesgo.','No','','','Hard','Buena'],
    ['D\'Andre Swift',        'RB','CHI','2', 58.0, 58.3, 'RB CHI. Williams QB Año 3. Waldron OC. Burden + Odunze WRs. 190 carries + 60 rec. OL CHI pobre pero Swift veloz.','No','','','Easy','Pobre'],
    ['Rome Odunze',           'WR','CHI','2', 63.0, 63.3, 'WR CHI. Año 3 (clase 2024). Williams QB Año 3. Waldron OC Año 2. Luther Burden WR2. 80 rec + 950 yds.','No','','','Easy','Pobre'],
    ['Bhayshul Tuten',        'RB','JAC','2', 64.3, 64.6, 'RB JAC. Lawrence QB. Thomas Jr. WR1. Hunter WR2. Moore OC. 200 carries + 50 rec proyectados.','No','','','Neutral','Promedio'],
    ['Jalen Hurts',           'QB','PHI','2', 66.3, 66.6, 'Top QB Custom format. 120 carries + 650 yds + 10 TDs terrestres. OL élite. Brown + Smith WRs. Warren TE.','No','','','Easy','Élite'],
    ['Caleb Williams',        'QB','CHI','2', 66.7, 67.0, '🔼 Ben Johnson HC ÉLITE + Press Taylor OC = MEJOR sistema posible para Williams Año 3. Swift RB. Burden + Odunze WRs. OL pobre sigue siendo el único riesgo real.','No','','','Easy','Pobre'],
    ['Harold Fannin Jr.',     'TE','CLE','2', 68.3, 68.6, 'Año 2 (clase 2025). TE CLE. QB caos pero Fannin = target hog. Judkins RB. 70 rec + 750 yds potencial breakout.','Yes','','','Easy','Promedio'],
    ['Tucker Kraft',          'TE', 'GB','2', 71.0, 71.3, 'TE GB. Love QB. Jacobs RB. Adams + Watson WRs. Kraft como safety blanket. 60 rec + 650 yds + 6 TDs.','No','','','Easy','Muy buena'],
    ['Christian Watson',      'WR', 'GB','2', 71.0, 71.3, 'WR2 GB. Love QB. Adams WR1. 75 rec + 950 yds + 8 TDs. Deep threat élite. Schedule playoff GB élite.','No','','','Easy','Muy buena'],
    ['Marvin Harrison Jr.',   'WR','ARI','2', 71.0, 71.3, 'WR1 ARI. Grubb OC Año 2. McBride TE. Love/Brissett QB. 85 rec + 1050 yds + 7 TDs. Sistema moderno Grubb.','Yes','','','Neutral','Promedio'],
    ['Brian Thomas Jr.',      'WR','JAC','2', 71.7, 72.0, 'WR1 JAC. Año 3 (clase 2024). Lawrence QB. Hunter WR2. Meyers WR3. 80 rec + 1100 yds + 7 TDs.','No','','','Neutral','Promedio'],
    ['Chuba Hubbard',         'RB','CAR','2', 72.0, 72.3, 'RB1 CAR workhorse. Bryce Young QB. McMillan WR1. OL peor del NFL. 210 carries alto volumen, baja eficiencia.','No','','','Neutral','Muy pobre'],
    ['Carnell Tate',          'WR','TEN','2', 72.7, 73.0, '🔼 Año 2. WR TEN. Cam Ward QB + Brian Daboll OC (construyó Josh Allen en BUF). Breakout candidato real. TEN el stack más infravalorado del draft.','No','','','Neutral','Promedio'],

    // ── TIER 3 — ADP 73-120 ─────────────────────────────────────────────────
    ['Jordyn Tyson',          'WR', 'NO','3', 76.7, 77.0, 'Año 2 (clase 2025). WR NO. Shough QB. Olave WR1. Etienne+Kamara en backfield. 70 rec + 850 yds potencial.','No','','','Easy','Buena'],
    ['Alec Pierce',           'WR','IND','3', 78.3, 78.6, 'WR1 IND extendido ($116M). Daniel Jones QB. Warren TE. Taylor RB. 85 rec + 1050 yds proyectados.','No','','','Easy','Promedio-Buena'],
    ['Sam LaPorta',           'TE','DET','3', 79.3, 79.6, 'TE DET. Goff QB. ARSB WR1. Gibbs RB élite. Volumen consistente. 80 rec + 850 yds + 6 TDs. Año 3.','No','','','Neutral','Élite'],
    ['Kyle Pitts',            'TE','ATL','3', 79.7, 80.0, 'TE ATL. Tua nuevo QB. Bijan RB1. London WR1. Si Pitts regresa forma = TE1 en ofensiva élite ATL. Riesgo salud.','Yes','','','Neutral','Buena'],
    ['Trevor Lawrence',       'QB','JAC','3', 81.0, 81.3, 'QB JAC. Thomas Jr. WR1. Hunter WR2. Tuten RB. Meyers WR3. Moore OC inconsistente. 3800 yds + 27 TDs.','No','','','Neutral','Promedio'],
    ['Justin Herbert',        'QB','LAC','3', 82.3, 82.6, 'QB LAC. McConkey WR1. Hampton RB. Johnston WR. Gadsden TE. 4200 yds + 30 TDs. Preciso y confiable.','No','','','Easy','Promedio'],
    ['Jaylen Warren',         'RB','PIT','3', 82.7, 83.0, 'RB PIT. Angelichio nuevo OC. DK Metcalf llegó de SEA. Rodgers QB. 210 carries + 60 rec. Rico Dowdle compite.','Yes','','','Hard','Buena'],
    ['Jaxson Dart',           'QB','NYG','3', 84.7, 85.0, '🔼 Año 2 NYG. John Harbaugh HC + Matt Nagy OC = upgrade masivo. Dart + Nabers = STACK SORPRESA del draft. Harbaugh construyó offenses élite en BAL durante 16 años.','No','','','Neutral','Promedio'],
    ['DK Metcalf',            'WR','PIT','3', 84.7, 85.0, 'WR PIT. Llegó de SEA. Rodgers QB. Angelichio nuevo OC. Warren RB. 80 rec + 1100 yds + 8 TDs. Fit interesante.','Yes','','','Hard','Buena'],
    ['Courtland Sutton',      'WR','DEN','3', 86.0, 86.3, 'WR DEN. Bo Nix QB. Waddle llegó de MIA. Harvey RB. 80 rec + 1000 yds + 7 TDs proyectados.','No','','','Neutral','Promedio'],
    ['Tony Pollard',          'RB','TEN','3', 86.0, 86.3, 'RB TEN. Saleh HC + Daboll OC nuevo. Cam Ward QB dual-threat. Tyjae Spears también en backfield. 200 carries + 55 rec proyectados.','No','','','Neutral','Promedio'],
    ['Rico Dowdle',           'RB','PIT','3', 86.7, 87.0, 'RB PIT. Warren compite. Metcalf llegó. Rodgers QB. Angelichio OC. 150 carries + 40 rec si gana el starting job.','Yes','','','Hard','Buena'],
    ['Patrick Mahomes',       'QB', 'KC','3', 88.7, 89.0, '⭐ Eric Bieniemy REGRESA como OC = construyó los 4 Super Bowls con Mahomes. Rice WR1 + Worthy WR2. ADP 89 = MAYOR VOR POR PICK del draft en formato Custom.','No','','','Hard','Élite'],
    ['Kyle Monangai',         'RB','CHI','3', 89.3, 89.6, 'RB CHI. Swift también. Williams QB Año 3. Burden + Odunze WRs. 130 carries + 45 rec como complemento.','No','','','Easy','Pobre'],
    ['Michael Wilson',        'WR','ARI','3', 90.0, 90.3, 'WR ARI. Grubb OC. Harrison Jr. WR1. McBride TE. 70 rec + 850 yds proyectados.','Yes','','','Neutral','Promedio'],
    ['Makai Lemon',           'WR','PHI','3', 90.3, 90.6, 'Año 2 (clase 2025). WR PHI. Hurts QB. Brown + Smith WRs. Lemon como WR3 en ofensiva élite. Upside en bye semanas.','No','','','Easy','Élite'],
    ['Chris Godwin',          'WR', 'TB','3', 91.0, 91.3, 'WR1 TB. Mayfield QB. Egbuka WR2. Godwin slot veteran. 85 rec + 1050 yds + 7 TDs. Año 9. Confiable.','No','','','Easy','Buena'],
    ['Matthew Stafford',      'QB','LAR','3', 91.7, 92.0, 'QB LAR. Nacua + Adams WRs. Williams RB. OL muy buena. 4200 yds + 28 TDs. Bajo ADP = valor en superflex.','No','','','Neutral','Muy buena'],
    ['George Kittle',         'TE', 'SF','3', 95.3, 95.6, 'TE SF. Shanahan system. Purdy QB. McCaffrey RB. Evans WR. Kittle = safety blanket élite. 75 rec + 850 yds + 7 TDs.','No','','','Neutral','Élite'],
    ['Brock Purdy',           'QB', 'SF','3', 96.7, 97.0, 'QB SF. Shanahan system élite. Evans + Kittle. McCaffrey si sano. 4000 yds + 32 TDs. Muy confiable si CMC sano.','No','','','Neutral','Élite'],
    ['Bo Nix',                'QB','DEN','3', 98.3, 98.6, 'QB DEN. Segundo año. Waddle + Sutton WRs. Harvey RB. 3700 yds + 27 TDs. Upside si OL DEN mejora.','No','','','Neutral','Promedio'],
    ['Blake Corum',           'RB','LAR','3', 99.7,100.0, 'RB2 LAR. Williams es lead. Corum como complemento. 120 carries + 35 rec. Shanahan-like system LAR.','No','','','Neutral','Muy buena'],
    ['RJ Harvey',             'RB','DEN','3', 78.3, 78.6, 'Año 2 (clase 2025). RB DEN. Bo Nix QB. Waddle + Sutton WRs. Comparte con J.K. Dobbins. Breakout candidato Year 2.','No','','','Neutral','Promedio'],
    ['J.K. Dobbins',          'RB','DEN','3',101.7,102.0, 'RB DEN. Harvey + Dobbins = backfield compartido. Bo Nix QB. 140 carries potencial. Veterano experimentado.','No','','','Neutral','Promedio'],
    ['Jayden Reed',           'WR', 'GB','3',101.7,102.0, 'WR GB. Love QB. Adams WR1, Watson WR2. Reed en slot. 70 rec + 850 yds proyectados. Schedule élite GB.','No','','','Easy','Muy buena'],
    ['Jared Goff',            'QB','DET','3',102.0,102.3, 'QB DET. ARSB + Williams WRs. Gibbs RB. OL élite DET. 4500 yds + 30 TDs. LaPorta TE. ADP bajo = QB2 de valor.','No','','','Neutral','Élite'],
    ['Jordan Addison',        'WR','MIN','3',103.7,104.0, 'WR MIN. Kyler Murray QB. Jefferson WR1. 75 rec + 950 yds proyectados. Murray run-first pero Addison se beneficia.','Yes','','','Easy','Buena'],
    ['Jakobi Meyers',         'WR','JAC','3',105.0,105.3, 'WR JAC. Llegó de LV a JAC. Lawrence QB. Thomas Jr. WR1. Hunter WR2. Meyers en slot. 65 rec + 750 yds.','No','','','Neutral','Promedio'],
    ['Kyler Murray',          'QB','MIN','3',107.3,107.6, 'QB MIN. Llegó de ARI. Jefferson WR1. Jones + Mason RBs. Addison WR2. Dual-threat = Custom format. 3700 yds + 27 TDs + 500 rushYds.','Yes','','','Easy','Buena'],
    ['Wan\'Dale Robinson',    'WR','TEN','3',107.7,108.0, 'WR TEN. Llegó de NYG a TEN. Cam Ward QB Año 2. Tate WR. Pollard RB. 70 rec + 800 yds en slot.','No','','','Neutral','Promedio'],

    // ── TIER 4 — ADP 109-168 ─────────────────────────────────────────────────
    ['Ricky Pearsall',        'WR', 'SF','4',109.7,110.0, 'WR SF. Año 2 (clase 2025). Evans WR1. Kittle TE. Shanahan system = muchos targets. Breakout candidato.','No','','','Neutral','Élite'],
    ['Dalton Kincaid',        'TE','BUF','4',109.7,110.0, 'TE BUF. Allen QB. Moore + Shakir WRs. 60 rec + 700 yds + 6 TDs. Allen distribuye bien. Sólido TE2.','No','','','Easy','Muy buena'],
    ['Michael Pittman Jr.',   'WR','PIT','4',110.0,110.3, 'WR PIT. Llegó de IND. DK Metcalf es WR1 ahora. Rodgers QB. Angelichio OC. 70 rec + 900 yds en WR2 role.','Yes','','','Hard','Buena'],
    ['Jake Ferguson',         'TE','DAL','4',110.7,111.0, 'TE DAL. Dak Prescott QB. Lamb + Pickens WRs. 65 rec + 750 yds + 6 TDs. OL élite DAL. Confiable.','No','','','Neutral','Muy buena'],
    ['Quentin Johnston',      'WR','LAC','4',112.0,112.3, 'WR LAC. Herbert QB. McConkey WR1. Johnston como WR2 deep. 65 rec + 800 yds potencial.','No','','','Easy','Promedio'],
    ['Tyler Shough',          'QB', 'NO','4',112.3,112.6, 'QB NO. Nuevo titular. Olave WR1. Etienne + Kamara en backfield. Tyson WR emergente. 3200 yds + 22 TDs est.','No','','','Easy','Buena'],
    ['Isaiah Likely',         'TE','NYG','4',113.3,113.6, 'TE NYG. Llegó de BAL a NYG. Dart QB. Nabers WR1. Skattebo RB. 55 rec + 600 yds + 5 TDs.','No','','','Neutral','Promedio'],
    ['Josh Downs',            'WR','IND','4',113.0,113.3, 'WR IND. Pierce WR1. Daniel Jones QB. Taylor RB. Warren TE. Downs en slot. 70 rec + 800 yds.','No','','','Easy','Promedio-Buena'],
    ['Jordan Love',           'QB', 'GB','4',115.3,115.6, 'QB GB. Jacobs + Watson + Adams. OL muy buena. 3900 yds + 32 TDs. Schedule playoff élite GB.','No','','','Easy','Muy buena'],
    ['Baker Mayfield',        'QB', 'TB','4',116.3,116.6, 'QB TB. Egbuka + Godwin WRs. Irving RB. 3800 yds + 28 TDs. Sistema cómodo para Mayfield.','No','','','Easy','Buena'],
    ['Dallas Goedert',        'TE','PHI','4',117.0,117.3, 'TE PHI. Hurts QB. Brown + Smith WRs. OL élite. Goedert safety blanket. 65 rec + 700 yds + 7 TDs.','No','','','Easy','Élite'],
    ['Jordan Mason',          'RB','MIN','4',117.3,117.6, 'RB MIN. Llegó de SF a MIN. Kyler Murray QB. Jefferson WR1. Aaron Jones también. 140 carries compartido.','Yes','','','Easy','Buena'],
    ['Mark Andrews',          'TE','BAL','4',119.0,119.3, 'TE BAL. Lamar QB. Flowers WR1. 60 rec + 700 yds + 8 TDs. Injury risk históricamente. Valor alto si sano.','No','','','Neutral','Muy buena'],
    ['Xavier Worthy',         'WR', 'KC','4',120.0,120.3, 'WR KC. Mahomes. Rice WR1. 70 rec + 900 yds + 8 TDs. Deep threat élite. Sistema Reid lo maximiza.','No','','','Hard','Élite'],
    ['Rachaad White',         'RB','WAS','4',120.0,120.3, 'RB WAS. Llegó de TB a WAS. Daniels QB. McLaurin WR1. Croskey-Merritt compite. 170 carries + 55 rec.','No','','','Hard','Buena'],
    ['Matthew Golden',        'WR', 'GB','4',121.3,121.6, 'Año 2 (clase 2025). WR GB. Love QB. Adams WR1. Watson WR2. Golden como WR3 emergente. Upside.','No','','','Easy','Muy buena'],
    ['Jacory Croskey-Merritt', 'RB','WAS','4',122.0,122.3,'Año 2 (clase 2025). RB WAS. Daniels QB. McLaurin WR1. White compite. Backup que puede robar carries.','No','','','Hard','Buena'],
    ['Aaron Jones',           'RB','MIN','4',122.3,122.6, 'RB MIN. Murray QB. Jefferson WR1. Mason también. Jones receptor. 130 carries + 55 rec. Veterano.','Yes','','','Easy','Buena'],
    ['Jonathon Brooks',       'RB','CAR','4',124.3,124.6, 'RB CAR. Hubbard es lead. Brooks como backup emergente. Bryce Young QB. OL peor. 100 carries potencial.','No','','','Neutral','Muy pobre'],
    ['Chris Rodriguez Jr.',   'RB','JAC','4',124.7,125.0, 'RB JAC. Tuten es lead. Rodriguez como complemento. Lawrence QB. Thomas Jr. WR1. 100 carries potencial.','No','','','Neutral','Promedio'],
    ['Tyrone Tracy Jr.',      'RB','NYG','4',130.3,130.6, 'RB NYG. Skattebo ADP 42 es lead. Tracy como backup. Dart QB. Nabers WR1. OL pobre. 100 carries potencial.','No','','','Neutral','Promedio'],
    ['Kenyon Sadiq',          'TE','NYJ','4',130.0,130.3, 'TE NYJ. Frank Reich nuevo OC. Hall RB. Wilson WR1. 55 rec + 600 yds potencial en Reich system.','Yes','','','Neutral','Pobre'],
    ['Oronde Gadsden II',     'TE','LAC','4',131.0,131.3, 'TE LAC. Herbert QB. McConkey WR1. Hampton RB. 60 rec + 650 yds potencial. Año 3.','No','','','Easy','Promedio'],
    ['Malik Willis',          'QB','MIA','4',132.3,132.6, 'QB MIA. Tua se fue a ATL. Willis nuevo titular. Achane RB élite. Waddle se fue a DEN. Nuevo ecosistema.','Yes','','','Neutral','Buena'],
    ['Khalil Shakir',         'WR','BUF','4',133.7,134.0, 'WR BUF. Moore WR1. Shakir slot. Allen distribuyente. 75 rec + 850 yds. Confiabilidad alta.','No','','','Easy','Muy buena'],
    ['Sam Darnold',           'QB','SEA','4',137.3,137.6, 'QB SEA. JSN WR1 (ADP 5.3). Metcalf se fue a PIT. Price RB. 3800 yds + 28 TDs. Nuevo sistema.','No','','','Neutral','Buena'],
    ['Jayden Higgins',        'WR','HOU','4',138.0,138.3, 'Año 2 (clase 2025). WR HOU. Stroud QB. Collins WR1. Dell WR2. Higgins como WR3 emergente.','No','','','Neutral','Buena'],
    ['CJ Stroud',             'QB','HOU','4',138.7,139.0, 'QB HOU. Collins WR1. Dell WR2. Montgomery RB. Higgins emergente. 3900 yds + 30 TDs. Confiable.','No','','','Neutral','Buena'],
    ['Cam Ward',              'QB','TEN','4',141.0,141.3, '🔼 Año 2 TEN. Saleh HC + Brian Daboll OC (construyó Josh Allen en BUF). Ward = perfil idéntico a Allen. STACK MÁS INFRAVALORADO del draft. ADP 141 = valor masivo en Custom.','No','','','Neutral','Promedio'],
    ['Juwan Johnson',         'TE', 'NO','4',140.3,140.6, 'TE NO. Shough QB. Olave WR1. Etienne + Kamara RBs. Johnson safety blanket. 50 rec + 550 yds.','No','','','Easy','Buena'],
    ['Jalen Coker',           'WR','CAR','4',141.3,141.6, 'WR CAR. Bryce Young QB. McMillan WR1. Coker como WR2. OL pobre pero WRs necesitan producir.','No','','','Neutral','Muy pobre'],
    ['Chig Okonkwo',          'TE','WAS','4',145.3,145.6, 'TE WAS. Daniels QB. McLaurin WR1. Kingsbury OC air raid. 55 rec + 600 yds potencial.','No','','','Hard','Buena'],
    ['Isiah Pacheco',         'RB','DET','4',146.0,146.3, 'RB DET. Gibbs lead back. Pacheco backup valioso. OL élite DET. Handcuff crítico de Gibbs. 100 carries si Gibbs lesiona.','No','','','Neutral','Élite'],
    ['Zach Charbonnet',       'RB','SEA','4',147.3,147.6, 'RB SEA. Darnold QB. JSN WR1. Price RB también. 130 carries + 40 rec. Backup valioso.','No','','','Neutral','Buena'],
    ['Hunter Henry',          'TE', 'NE','4',147.3,147.6, 'TE NE. Maye QB. Henderson RB. 55 rec + 600 yds + 5 TDs. Nuevo sistema Van Pelt.','Yes','','','Neutral','Pobre'],
    ['Daniel Jones',          'QB','IND','4',147.0,147.3, 'QB IND. Taylor RB élite. Pierce WR extendido. Warren TE. Jones como gestor. 3200 yds + 22 TDs.','No','','','Easy','Promedio-Buena'],
    ['Jonah Coleman',         'RB','DEN','4',148.7,149.0, 'RB DEN. Harvey + Dobbins + Coleman = 3 RBs. Bo Nix QB. Waddle + Sutton WRs. Rotación profunda.','No','','','Neutral','Promedio'],
    ['Tyler Allgeier',        'RB','ARI','4',149.7,150.0, 'RB ARI. Jeremiyah Love es lead. Allgeier como complemento. Grubb OC. Harrison Jr. WR1.','Yes','','','Neutral','Promedio'],
    ['T.J. Hockenson',        'TE','MIN','4',150.3,150.6, 'TE MIN. Kyler Murray QB. Jefferson WR1. Murray run-first = menos TE targets. Riesgo en nuevo sistema.','Yes','','','Easy','Buena'],
    ['Bryce Young',           'QB','CAR','4',151.7,152.0, 'QB CAR. McMillan WR1. Hubbard RB. OL peor NFL. Reconstrucción total. Solo valor especulativo.','No','','','Neutral','Muy pobre'],
    ['Tyjae Spears',          'RB','TEN','4',159.7,160.0, 'RB TEN. Pollard es lead. Spears como complemento receptor. 60 rec + 450 yds potencial.','No','','','Neutral','Promedio'],
    ['Alvin Kamara',          'RB', 'NO','4',161.7,162.0, 'Veterano NO. Etienne + Kamara = doble RB. Kamara = receptor élite veterano. 65 rec + 500 yds potencial.','No','','','Easy','Buena'],
    ['Keaton Mitchell',       'RB','LAC','4',162.3,162.6, 'Speed back LAC. Hampton es lead. Mitchell complemento explosivo. 60 carries + 45 rec potencial.','No','','','Easy','Promedio'],
    ['Jalen Nailor',          'WR', 'LV','4',164.3,164.6, 'WR LV. Bowers TE1. Jeanty RB. Mendoza QB. 60 rec + 700 yds potencial.','Yes','','','Neutral','Promedio'],
    ['Tyreek Hill',           'WR', 'FA','4',165.3,165.6, '⚠️ AGENTE LIBRE — Sin equipo confirmado mayo 2026. Si firma = WR1 inmediato. Verificar antes del draft.','No','','','Neutral','N/A'],
    ['Brian Robinson Jr.',    'RB','ATL','4',168.7,169.0, 'RB ATL. Bijan Robinson lead. Robinson Jr. backup (ADP 169). Tua QB. Si Bijan lesionado = RB1 inmediato.','Yes','','','Neutral','Buena'],

    // ── TIER 5 — ADP 169+ ─────────────────────────────────────────────────────
    ['Jacoby Brissett',       'QB','ARI','5',169.0,169.3, 'QB ARI. Harrison Jr. WR1. McBride TE. Grubb OC. Backup confiable. 3000 yds + 20 TDs est.','Yes','','','Neutral','Promedio'],
    ['Brandon Aiyuk',         'WR', 'SF','5',169.0,169.3, 'WR SF. Evans WR1. Shanahan system. Pearsall también. 65 rec + 800 yds potencial.','No','','','Neutral','Élite'],
    ['Fernando Mendoza',      'QB', 'LV','5',169.3,169.6, 'Año 2 QB LV. Bowers TE1. Jeanty RB. OC Getsy. 3200 yds + 22 TDs est. Nailor WR.','Yes','','','Neutral','Promedio'],
    ['Dylan Sampson',         'RB','CLE','5',170.7,171.0, 'Año 2 RB CLE. Judkins es lead. Sampson como complemento. QB caos CLE. 100 carries potencial.','Yes','','','Easy','Promedio'],
    ['Nicholas Singleton',    'RB','TEN','5',171.0,171.3, 'Año 2 RB TEN. Pollard lead. Spears + Singleton en rotación. 80 carries potencial.','No','','','Neutral','Promedio'],
    ['Aaron Rodgers',         'QB','PIT','5',172.3,172.6, 'Veterano PIT. Angelichio nuevo OC. Metcalf WR1. Pittman WR2. Warren RB. 3500 yds + 28 TDs.','Yes','','','Hard','Buena'],
    ['Travis Hunter',         'WR','JAC','5',174.5,174.8, 'Año 2 (clase 2025). WR/CB doble función JAC. Thomas Jr. WR1. Upside alto si se establece. 60 rec potencial.','No','','','Neutral','Promedio'],
    ['Mike Washington Jr.',   'RB', 'LV','5',174.0,174.3, 'RB LV. Jeanty lead back. Washington Jr. como backup. 80 carries potencial.','Yes','','','Neutral','Promedio'],
    ['Geno Smith',            'QB','NYJ','5',175.7,176.0, 'QB NYJ. Frank Reich OC. Hall RB. Wilson WR1. Sadiq TE. 3200 yds + 22 TDs est.','Yes','','','Neutral','Pobre'],
    ['Deebo Samuel',          'WR', 'FA','5',177.7,178.0, '⚠️ AGENTE LIBRE — Sin equipo mayo 2026. Valor en Custom format por rushing (50+ carries + 300 yds).','No','','','Neutral','N/A'],
    ['Antonio Williams',      'WR','WAS','5',177.0,177.3, 'WR WAS. Daniels QB. McLaurin WR1. Kingsbury air raid. Williams como WR2/3. 55 rec + 650 yds.','No','','','Hard','Buena'],
    ['Jerry Jeudy',           'WR','CLE','5',178.7,179.0, 'WR CLE. QB caos. Jeudy como WR receptor. Judkins RB. 55 rec + 650 yds si QB funciona.','Yes','','','Easy','Promedio'],
    ['Tank Bigsby',           'RB','PHI','5',180.3,180.6, 'Backup Barkley. OL élite PHI. Handcuff crítico. Si Barkley lesionado = RB1 inmediato en ofensiva élite.','No','','','Easy','Élite'],
    ['Calvin Ridley',         'WR','TEN','5',181.3,181.6, 'WR TEN. Ward QB Año 2. Holz OC. Tate WR. Pollard RB. 70 rec + 850 yds potencial.','No','','','Neutral','Promedio'],
    ['Tre Tucker',            'WR', 'LV','5',182.0,182.3, 'WR LV. Bowers TE1. Jeanty RB. 55 rec + 650 yds potencial.','Yes','','','Neutral','Promedio'],
    ['Tua Tagovailoa',        'QB','ATL','5',197.7,198.0, 'QB ATL. Llegó de MIA. Bijan RB1. London WR1. Pitts TE. Robinson Jr. backup. 3400 yds + 24 TDs.','Yes','','','Neutral','Buena'],

    // ── K y D/ST ─────────────────────────────────────────────────────────────
    ['Eagles D/ST',           'D', 'PHI','5',172.0,173.0, 'D/ST PHI. OL élite + defensiva agresiva. Top-3 D/ST. Schedule playoff favorable.','No','','','Easy','N/A'],
    ['Ravens D/ST',           'D', 'BAL','5',173.0,174.0, 'D/ST BAL. Defensiva élite histórica. Wk15-17 neutral. Top-5 D/ST.','No','','','Neutral','N/A'],
    ['Steelers D/ST',         'D', 'PIT','5',174.0,175.0, 'D/ST PIT. T.J. Watt si sano. Históricamente top D/ST. Schedule Wk15-17 difícil.','No','','','Hard','N/A'],
    ['Jake Elliott',          'K', 'PHI','5',169.0,170.0, 'K PHI. Ofensiva élite = muchas oportunidades. Top-5 kicker. Barkley + Hurts = scoring machine.','No','','','Easy','N/A'],
    ['Evan McPherson',        'K', 'CIN','5',168.0,169.0, 'K CIN. Chase + Burrow = TD machine + FG. Top-3 kicker proyectado.','No','','','Easy','N/A'],
    ['Harrison Butker',       'K',  'KC','5',170.0,171.0, 'K KC. Mahomes = siempre scoring. Top-5 kicker. Hard schedule playoffs pero KC siempre llega lejos.','No','','','Hard','N/A'],
  ];

  sheet.getRange(2, 1, data.length, 12).setValues(data);

  var noteRow = data.length + 3;
  sheet.getRange(noteRow, 1).setValue(
    '✅ Players DB 2026 VERIFICADO — ' + data.length + ' jugadores | ' +
    'ADP: FantasyPros Half-PPR directo mayo 2026 | ' +
    '⚠️ Tyreek Hill + Deebo Samuel = agentes libres. Verificar firma.'
  ).setFontStyle('italic').setFontColor('#1a7b47').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '✅ Players DB actualizado: ' + data.length + ' jugadores.\n\n' +
    'Correcciones principales aplicadas:\n' +
    '• Javonte Williams → DAL\n' +
    '• Cam Skattebo → NYG\n' +
    '• George Pickens → DAL\n' +
    '• DK Metcalf → PIT\n' +
    '• Jaylen Waddle → DEN\n' +
    '• Davante Adams → LAR\n' +
    '• Jordan Mason → MIN\n' +
    '• Rachaad White → WAS\n' +
    '• Jaxon Smith-Njigba → SEA ADP 5.3\n' +
    '• Ashton Jeanty, Omarion Hampton, Chase Brown añadidos'
  );
}
