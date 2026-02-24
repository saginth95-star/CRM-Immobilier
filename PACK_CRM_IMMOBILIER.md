# CRM Immobilier V2 — Google Sheets + Apps Script (mono-agent, duplicable)

Version pensée **terrain** : moins de saisie, plus de menus déroulants, plus d’automatisation, interface lisible.

---

## 1) Nouvelle structure d’onglets (orientée usage)

## INTERFACE (utilisateur)
1. `INTERFACE_DASHBOARD`
2. `INTERFACE_BIENS`
3. `INTERFACE_CLIENTS`
4. `INTERFACE_MATCHING` (fiche dynamique par client)
5. `INTERFACE_SUIVI` (historique actions client-bien)
6. `INTERFACE_COMMISSIONS`
7. `PARAMETRES`

## DATA (protégée)
8. `DATA_LISTES`
9. `DATA_BIENS`
10. `DATA_CLIENTS`
11. `DATA_SUIVI_CLIENT_BIEN` ✅ (historique indépendant par client/bien)
12. `DATA_COMMISSIONS`
13. `DATA_STATS`

> Logique clé V2 : le **bien est global**, mais le **statut (visité/refusé/favori/à relancer)** est stocké par couple `Client + Bien` dans `DATA_SUIVI_CLIENT_BIEN`.

---

## 2) Script Apps Script complet (Code.gs unique)

> Copiez-collez ce code dans un seul fichier `Code.gs`.

```javascript

/**
 * CRM Immobilier V4 - Google Sheets
 * UX premium, terrain, mono-agent, duplicable
 * Un seul fichier Code.gs
 */

const APP = {
  menu: '📊 CRM Immobilier V4',
  sh: {
    dashboard: 'INTERFACE_DASHBOARD',
    biens: 'INTERFACE_BIENS',
    clients: 'INTERFACE_CLIENTS',
    matching: 'INTERFACE_MATCHING',
    suivi: 'INTERFACE_SUIVI',
    commissions: 'INTERFACE_COMMISSIONS',
    params: 'PARAMETRES',
    listes: 'DATA_LISTES',
    dataBiens: 'DATA_BIENS',
    dataClients: 'DATA_CLIENTS',
    dataSuivi: 'DATA_SUIVI_CLIENT_BIEN',
    dataCom: 'DATA_COMMISSIONS',
    dataStats: 'DATA_STATS'
  },
  ui: {
    header: '#1f2937',
    headerText: '#ffffff',
    zebra1: '#ffffff',
    zebra2: '#f8fafc',
    success: '#dcfce7',
    warning: '#ffedd5',
    danger: '#fee2e2',
    info: '#dbeafe',
    neutral: '#e5e7eb',
    border: '#e5e7eb',
    textMuted: '#4b5563'
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP.menu)
    .addItem('🚀 Initialiser CRM V4', 'initialiserCRMV4')
    .addSeparator()
    .addItem('🔄 Rafraîchir tout', 'refreshAllV4_')
    .addItem('🎯 Recalculer matching client', 'recalculerMatchingClient')
    .addItem('✅ Valider ligne matching sélectionnée', 'validerLigneMatchingSelectionnee')
    .addItem('📋 Trier suivi', 'trierSuivi_')
    .addItem('📊 Mettre à jour dashboard', 'majDashboardV2')
    .addItem('🎨 Réappliquer design premium', 'appliquerDesignV4_')
    .addItem('🔒 Protéger onglets DATA', 'protegerData_')
    .addSeparator()
    .addItem('➕ Ajouter un client', 'ajouterClientRapide')
    .addItem('🏠 Ajouter un bien', 'ajouterBienRapide')
    .addToUi();
}

function initialiserCRMV4() {
  const ss = SpreadsheetApp.getActive();
  creerOngletsSiAbsents_(ss);
  definirEntetes_();
  remplirListes_();
  remplirParametres_();
  setupValidationsV2_();
  protegerData_();
  refreshAllV4_();
  SpreadsheetApp.getUi().alert('CRM V4 initialisé ✅');
}

// Compatibilité ancien menu
function initialiserCRMV2() { initialiserCRMV4(); }
function refreshAllV2() { refreshAllV4_(); }
function appliquerDesignV2() { appliquerDesignV4_(); }

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!sh) return;
  const name = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (name === APP.sh.biens) {
    autoDateCreation_(sh, row, 1, 2);
    majScoreBienLigne_(row);
    syncInterfaceVersDataBiens_();
    appliquerStylesBiens_();
  }

  if (name === APP.sh.clients) {
    autoDateCreation_(sh, row, 1, 2);
    majScoreClientLigne_(row);
    syncInterfaceVersDataClients_();
    appliquerStylesClients_();
  }

  if (name === APP.sh.matching) {
    if ((row === 2 || row === 3) && col === 2) {
      genererFicheMatching_();
      return;
    }
    if (row >= 6 && col === 7) {
      enregistrerStatutDepuisMatching_(row);
      majScoreClientDepuisSuivi_(row);
      majDashboardV2();
      genererFicheMatching_(); // cache immédiatement Refusés si filtre actif
    }
  }

  if (name === APP.sh.suivi) {
    syncInterfaceSuiviVersData_();
    trierSuivi_();
    majScoreActiviteClients_();
    appliquerStylesSuivi_();
    majDashboardV2();
  }

  if (name === APP.sh.commissions) {
    recalculCommissionsV2_();
    majDashboardV2();
  }
}

function refreshAllV4_() {
  majScoresBiens_();
  majScoreActiviteClients_();
  syncInterfaceVersDataBiens_();
  syncInterfaceVersDataClients_();
  genererFicheMatching_();
  syncDataSuiviVersInterface_();
  trierSuivi_();
  majDashboardV2();
  appliquerDesignV4_();
  SpreadsheetApp.getActive().toast('Rafraîchissement complet terminé ✅', 'CRM V4', 3);
}

function ajouterClientRapide() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const rep = ui.prompt('Nouveau client', 'Nom du client :', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const nom = String(rep.getResponseText() || '').trim();
  if (!nom) return;

  sh.appendRow([nom, new Date(), '', '', '', '', '', '', 'Nouveau', '', 0, 'Moyenne']);
  syncInterfaceVersDataClients_();
  appliquerStylesClients_();
}

function ajouterBienRapide() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const rep = ui.prompt('Nouveau bien', 'Adresse du bien :', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const adresse = String(rep.getResponseText() || '').trim();
  if (!adresse) return;

  sh.appendRow([adresse, new Date(), '', '', '', '', '', '', '', 'Prospection', '', '', '', '', '']);
  syncInterfaceVersDataBiens_();
  appliquerStylesBiens_();
}

function recalculerMatchingClient() {
  genererFicheMatching_();
  SpreadsheetApp.getActive().toast('Matching client recalculé ✅', 'CRM V4', 3);
}

function validerLigneMatchingSelectionnee() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== APP.sh.matching) {
    ss.toast('Placez-vous sur INTERFACE_MATCHING.', 'CRM V4', 4);
    return;
  }
  const row = sh.getActiveRange().getRow();
  if (row < 6) {
    ss.toast('Sélectionnez une ligne de matching (6+).', 'CRM V4', 4);
    return;
  }
  if (!String(sh.getRange('B2').getValue() || '')) {
    ss.toast('Sélectionnez un client en B2.', 'CRM V4', 4);
    return;
  }

  enregistrerStatutDepuisMatching_(row);
  majScoreClientDepuisSuivi_(row);
  majDashboardV2();
  genererFicheMatching_();
  ss.toast('Ligne validée ✅', 'CRM V4', 3);
}

function genererFicheMatching_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.matching);
  const clients = getRows_(APP.sh.clients, 12);
  const biens = getRows_(APP.sh.biens, 15);
  const suivi = getRows_(APP.sh.dataSuivi, 8);

  sh.getRange('A1:H1').merge().setValue('FICHE MATCHING CLIENT').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('Client sélectionné');
  sh.getRange('A3').setValue('Filtre affichage');
  sh.getRange('D2').setValue('Dernière MAJ');
  sh.getRange('E2').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('A4:H4').setValues([['Bien', 'Ville', 'Type', 'Prix', 'Score', 'Niveau', 'Statut client-bien', 'Dernière action']]);

  const listClientNames = clients.map(r => r[0]).filter(Boolean);
  if (listClientNames.length) {
    const ruleClient = SpreadsheetApp.newDataValidation().requireValueInList(listClientNames, true).setAllowInvalid(false).build();
    sh.getRange('B2').setDataValidation(ruleClient);
  }

  const filtres = ['Tout', 'Non traités', 'Favoris seulement', 'À relancer seulement', 'Masquer Refusés'];
  const ruleFiltre = SpreadsheetApp.newDataValidation().requireValueInList(filtres, true).setAllowInvalid(false).build();
  sh.getRange('B3').setDataValidation(ruleFiltre);
  if (!sh.getRange('B3').getValue()) sh.getRange('B3').setValue('Tout');

  const clientNom = String(sh.getRange('B2').getValue() || '');
  const filtre = String(sh.getRange('B3').getValue() || 'Tout');

  sh.getRange('A6:H700').clearContent().clearDataValidations().setBackground(null);
  if (!clientNom) {
    designMatchingV4_();
    return;
  }

  const c = clients.find(r => String(r[0]) === clientNom);
  if (!c) return;

  const budget = toNum_(c[4]);
  const villeRecherche = String(c[5] || '').toLowerCase();
  const typeRecherche = String(c[6] || '').toLowerCase();
  const surfaceMin = toNum_(c[7]);

  const rows = [];
  for (let i = 0; i < biens.length; i++) {
    const b = biens[i];
    const statutBien = String(b[9] || '');
    if (['Vendu', 'Retiré', 'Expiré'].includes(statutBien)) continue;

    const score = calculScoreMatching_(
      { ville: String(b[2] || ''), type: String(b[3] || ''), prix: toNum_(b[4]), surface: toNum_(b[5]) },
      { villeRecherche, typeRecherche, budget, surfaceMin }
    );

    const niveau = score >= 80 ? 'Très bon' : score >= 60 ? 'Moyen' : 'Faible';
    const found = suivi.find(s => keySuivi_(String(s[0]), String(s[1])) === keySuivi_(clientNom, String(b[0])));
    const statut = found ? String(found[4] || 'À proposer') : 'À proposer';
    const dateAct = found ? found[6] : '';

    if (!filtreMatchingOK_(filtre, statut)) continue;
    rows.push([b[0], b[2], b[3], b[4], score, niveau, statut, dateAct]);
  }

  rows.sort((a, b) => b[4] - a[4]);
  if (!rows.length) {
    designMatchingV4_();
    return;
  }

  sh.getRange(6, 1, rows.length, 8).setValues(rows);
  sh.getRange(6, 4, rows.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(6, 5, rows.length, 1).setNumberFormat('0');
  sh.getRange(6, 8, rows.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  const ruleStatut = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(APP.sh.listes).getRange('H2:H10'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(6, 7, rows.length, 1).setDataValidation(ruleStatut);

  applyMatchingStatusColors_(sh, rows.length);
  designMatchingV4_();
}

function calculScoreMatching_(bien, crit) {
  let ville = 0, budget = 0, type = 0, surface = 10;
  const vBien = String(bien.ville || '').toLowerCase();
  if (crit.villeRecherche && vBien === crit.villeRecherche) ville = 35;

  const p = toNum_(bien.prix);
  const b = toNum_(crit.budget);
  if (b > 0) {
    if (p <= b) budget = 35;
    else {
      const dep = (p - b) / b;
      budget = dep < 0.15 ? 22 : dep < 0.30 ? 10 : 0;
    }
  }

  const tBien = String(bien.type || '').toLowerCase();
  if (crit.typeRecherche && tBien === crit.typeRecherche) type = 20;

  const sMin = toNum_(crit.surfaceMin);
  if (sMin > 0) {
    const sBien = toNum_(bien.surface);
    surface = sBien >= sMin ? 10 : Math.max(0, 10 - Math.round((sMin - sBien) / 5));
  }

  return Math.max(0, Math.min(100, Math.round(ville + budget + type + surface)));
}

function filtreMatchingOK_(filtre, statut) {
  const st = String(statut || 'À proposer');
  if (filtre === 'Masquer Refusés') return st !== 'Refusé';
  if (filtre === 'Favoris seulement') return st === 'Favori';
  if (filtre === 'À relancer seulement') return st === 'À relancer';
  if (filtre === 'Non traités') return st === 'À proposer';
  return true;
}

function applyMatchingStatusColors_(sh, nRows) {
  const lv = sh.getRange(6, 6, nRows, 1).getValues();
  const st = sh.getRange(6, 7, nRows, 1).getValues();
  sh.getRange(6, 6, nRows, 1).setBackgrounds(lv.map(r => r[0] === 'Très bon' ? [APP.ui.success] : (r[0] === 'Moyen' ? [APP.ui.warning] : [APP.ui.danger])));
  sh.getRange(6, 7, nRows, 1).setBackgrounds(st.map(r => {
    const x = String(r[0] || '');
    if (x === 'Favori') return [APP.ui.success];
    if (x === 'À relancer') return [APP.ui.warning];
    if (x === 'Refusé') return [APP.ui.danger];
    if (x === 'Visité') return [APP.ui.info];
    return ['#ffffff'];
  }));
}

function enregistrerStatutDepuisMatching_(row) {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getSheetByName(APP.sh.matching);
  const clientNom = String(m.getRange('B2').getValue() || '');
  if (!clientNom) return;

  const bienAdresse = String(m.getRange(row, 1).getValue() || '');
  const score = toNum_(m.getRange(row, 5).getValue());
  const statut = String(m.getRange(row, 7).getValue() || 'À proposer');
  const now = new Date();

  const d = ss.getSheetByName(APP.sh.dataSuivi);
  const vals = getRows_(APP.sh.dataSuivi, 8);
  let idx = -1;
  for (let i = 0; i < vals.length; i++) {
    if (keySuivi_(String(vals[i][0]), String(vals[i][1])) === keySuivi_(clientNom, bienAdresse)) {
      idx = i + 2;
      break;
    }
  }

  if (idx === -1) d.appendRow([clientNom, bienAdresse, now, '', statut, score, now, Session.getActiveUser().getEmail() || '']);
  else d.getRange(idx, 5, 1, 3).setValues([[statut, score, now]]);

  syncDataSuiviVersInterface_();
}

function majScoreClientDepuisSuivi_(matchingRow) {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getSheetByName(APP.sh.matching);
  const clientNom = String(m.getRange('B2').getValue() || '');
  if (!clientNom) return;

  const clientsSh = ss.getSheetByName(APP.sh.clients);
  const clients = getRows_(APP.sh.clients, 12);
  const idx = clients.findIndex(r => String(r[0]) === clientNom);
  if (idx === -1) return;

  clientsSh.getRange(idx + 2, 10).setValue(new Date());
  majScoreClientLigne_(idx + 2);
  syncInterfaceVersDataClients_();
}

function keySuivi_(clientNom, bienAdresse) {
  return `${clientNom}|||${bienAdresse}`;
}

function autoDateCreation_(sheet, row, colReference, colDateCreation) {
  const ref = sheet.getRange(row, colReference).getValue();
  const d = sheet.getRange(row, colDateCreation).getValue();
  if (ref && !d) sheet.getRange(row, colDateCreation).setValue(new Date());
}

function majScoresBiens_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const last = sh.getLastRow();
  if (last < 2) return;
  for (let r = 2; r <= last; r++) majScoreBienLigne_(r);
  appliquerStylesBiens_();
}

function majScoreBienLigne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const p = getParams_();
  const dpe = String(sh.getRange(row, 7).getValue() || '').toUpperCase();
  const etat = String(sh.getRange(row, 8).getValue() || '');
  const dateFinMandat = sh.getRange(row, 9).getValue();

  const scoreDPE = ({ A: 7, B: 6, C: 5, D: 4, E: 3, F: 2, G: 1 }[dpe]) || 0;
  const scoreEtat = ({ 'Neuf': 4, 'Bon': 3, 'À rafraîchir': 2, 'Travaux': 1 }[etat]) || 0;
  const scoreGlobal = Math.round((scoreDPE * 0.55 + scoreEtat * 0.45) * 10) / 10;

  let alerte = '';
  if (dateFinMandat instanceof Date) {
    const jours = Math.floor((dateFinMandat - new Date()) / (1000 * 3600 * 24));
    if (jours < 0) alerte = 'Mandat expiré ❌';
    else if (jours <= p.alMandat) alerte = `Expire dans ${jours}j ⚠️`;
    else alerte = 'OK';
  }

  sh.getRange(row, 11, 1, 4).setValues([[scoreDPE, scoreEtat, scoreGlobal, alerte]]);
}

function majScoreActiviteClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const last = sh.getLastRow();
  if (last < 2) return;
  for (let r = 2; r <= last; r++) majScoreClientLigne_(r);
  appliquerStylesClients_();
}

function majScoreClientLigne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const suivi = getRows_(APP.sh.dataSuivi, 8);

  const nom = String(sh.getRange(row, 1).getValue() || '');
  if (!nom) return;

  const dernierContact = sh.getRange(row, 10).getValue();
  const actions = suivi.filter(r => String(r[0]) === nom);

  let visites = 0, refus = 0, favoris = 0;
  actions.forEach(a => {
    const st = String(a[4] || '');
    if (st === 'Visité') visites++;
    if (st === 'Refusé') refus++;
    if (st === 'Favori') favoris++;
  });

  let inact = 30;
  if (dernierContact instanceof Date) inact = Math.max(0, Math.floor((new Date() - dernierContact) / (1000 * 3600 * 24)));

  const score = Math.max(0, Math.min(100, Math.round(visites * 12 + favoris * 10 - refus * 8 + (30 - Math.min(30, inact)))));
  let priorite = 'Moyenne';
  if (score < 40 || inact > 14) priorite = 'Haute';
  else if (score >= 70 && inact <= 14) priorite = 'Basse';

  sh.getRange(row, 11, 1, 2).setValues([[score, priorite]]);
}

function majDashboardV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.dashboard);

  const biens = getRows_(APP.sh.biens, 15);
  const clients = getRows_(APP.sh.clients, 12);
  const suivi = getRows_(APP.sh.dataSuivi, 8);
  const coms = getRows_(APP.sh.commissions, 8);
  const p = getParams_();

  const clientsActifs = clients.filter(c => ['Actif', 'Tiède'].includes(String(c[8]))).length;
  const biensActifs = biens.filter(b => !['Vendu', 'Retiré', 'Expiré'].includes(String(b[9]))).length;

  const now = new Date();
  const mois = now.getMonth();
  const annee = now.getFullYear();
  const actionsMois = suivi.filter(r => r[6] instanceof Date && r[6].getMonth() === mois && r[6].getFullYear() === annee);

  const visitesMois = actionsMois.filter(a => String(a[4]) === 'Visité').length;
  const refusMois = actionsMois.filter(a => String(a[4]) === 'Refusé').length;
  const favoris = actionsMois.filter(a => String(a[4]) === 'Favori').length;
  const actionsRelance = suivi.filter(a => String(a[4]) === 'À relancer').length;
  const tauxRefus = visitesMois ? refusMois / visitesMois : 0;

  const caMensuel = coms
    .filter(c => c[0] instanceof Date && c[0].getMonth() === mois && c[0].getFullYear() === annee)
    .reduce((acc, c) => acc + toNum_(c[7]), 0);

  const pctObjectif = p.objectifMensuel ? caMensuel / p.objectifMensuel : 0;
  const objectifAtteint = caMensuel >= p.objectifMensuel;

  sh.getRange('A1:J80').clearContent().clearFormat();
  sh.getRange('A1:J1').merge().setValue('DASHBOARD TERRAIN — CRM IMMOBILIER V4');

  // Cards KPI
  drawKpiCard_(sh, 'A3:B6', 'Clients actifs', clientsActifs, APP.ui.info);
  drawKpiCard_(sh, 'C3:D6', 'Biens actifs', biensActifs, APP.ui.info);
  drawKpiCard_(sh, 'E3:F6', 'Visites (mois)', visitesMois, APP.ui.info);
  drawKpiCard_(sh, 'G3:H6', 'Refus (mois)', refusMois, refusMois > 0 ? APP.ui.warning : APP.ui.info);

  drawKpiCard_(sh, 'A7:B10', 'Favoris (mois)', favoris, APP.ui.success);
  drawKpiCard_(sh, 'C7:D10', 'Actions à relancer', actionsRelance, actionsRelance > 0 ? APP.ui.warning : APP.ui.success);
  drawKpiCard_(sh, 'E7:F10', 'CA mensuel net', formatEuro_(caMensuel), APP.ui.info);
  drawKpiCard_(sh, 'G7:H10', 'Objectif mensuel', formatEuro_(p.objectifMensuel), APP.ui.neutral);

  drawKpiCard_(sh, 'A11:C14', '% objectif atteint', Utilities.formatString('%.1f%%', pctObjectif * 100), objectifAtteint ? APP.ui.success : APP.ui.warning);
  drawKpiCard_(sh, 'D11:F14', 'Taux refus', Utilities.formatString('%.1f%%', tauxRefus * 100), tauxRefus > 0.5 ? APP.ui.danger : (tauxRefus > 0.3 ? APP.ui.warning : APP.ui.success));
  drawKpiCard_(sh, 'G11:H14', 'État objectif', objectifAtteint ? 'Atteint ✅' : 'En cours ⏳', objectifAtteint ? APP.ui.success : APP.ui.warning);

  sh.getRange('I3:J3').setValues([['Top clients actifs', 'Nb']]);
  const mapClients = mapCountByStatus_(suivi, 0, ['Visité', 'Favori', 'À relancer']);
  writeMapSorted_(sh, 4, 9, mapClients, 6);

  sh.getRange('I12:J12').setValues([['Top biens visités', 'Nb']]);
  const mapBiens = mapCountByStatus_(suivi, 1, ['Visité']);
  writeMapSorted_(sh, 13, 9, mapBiens, 6);

  buildChartsDashboardV2_(sh);
  designDashboardV4_();
}

function drawKpiCard_(sh, a1, title, value, bg) {
  const r = sh.getRange(a1);
  r.merge().setBackground(bg).setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);
  const row = r.getRow();
  const col = r.getColumn();
  sh.getRange(row, col).setValue(`${title}
${value}`).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold').setFontSize(11);
}

function mapCountByStatus_(rows, keyIdx, statuses) {
  const out = {};
  rows.forEach(r => {
    if (!statuses.includes(String(r[4]))) return;
    const k = String(r[keyIdx] || 'N/A');
    out[k] = (out[k] || 0) + 1;
  });
  return out;
}

function writeMapSorted_(sh, startRow, startCol, obj, maxRows) {
  const arr = Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, maxRows);
  if (!arr.length) return;
  sh.getRange(startRow, startCol, arr.length, 2).setValues(arr);
}

function buildChartsDashboardV2_(sh) {
  sh.getCharts().forEach(c => sh.removeChart(c));

  const c1 = sh.newChart()
    .asColumnChart()
    .addRange(sh.getRange('I3:J9'))
    .setOption('title', 'Top clients actifs')
    .setPosition(20, 1, 0, 0)
    .build();
  sh.insertChart(c1);

  const c2 = sh.newChart()
    .asBarChart()
    .addRange(sh.getRange('I12:J18'))
    .setOption('title', 'Top biens visités')
    .setPosition(20, 6, 0, 0)
    .build();
  sh.insertChart(c2);
}

function recalculCommissionsV2_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.commissions);
  const p = getParams_();
  const last = sh.getLastRow();
  if (last < 2) return;

  const vals = sh.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < vals.length; i++) {
    const prix = toNum_(vals[i][2]);
    if (!prix) continue;
    const taux = prix <= p.p1max ? p.p1taux : (prix <= p.p2max ? p.p2taux : p.p3taux);
    const brute = prix * taux;
    const ht = brute / (1 + p.tva);
    const tva = brute - ht;
    vals[i][3] = taux;
    vals[i][4] = round2_(brute);
    vals[i][5] = round2_(ht);
    vals[i][6] = round2_(tva);
    vals[i][7] = round2_(ht);
  }
  sh.getRange(2, 1, vals.length, 8).setValues(vals);
  sh.getRange(2, 4, vals.length, 1).setNumberFormat('0,00%');
  sh.getRange(2, 5, vals.length, 4).setNumberFormat('#,##0.00 €');
}

function creerOngletsSiAbsents_(ss) {
  Object.values(APP.sh).forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
}

function definirEntetes_() {
  const ss = SpreadsheetApp.getActive();

  setHeaders_(ss.getSheetByName(APP.sh.biens), ['Bien', 'Date création', 'Ville', 'Type', 'Prix affiché', 'Surface m²', 'DPE', 'État bien', 'Date fin mandat', 'Statut bien', 'Score DPE', 'Score État', 'Score global', 'Alerte mandat', 'Note interne']);
  setHeaders_(ss.getSheetByName(APP.sh.clients), ['Client', 'Date création', 'Téléphone', 'Email', 'Budget max', 'Ville recherchée', 'Type recherché', 'Surface min', 'Statut client', 'Dernier contact', 'Score activité', 'Priorité relance']);
  setHeaders_(ss.getSheetByName(APP.sh.matching), ['Zone', 'Valeur']);
  setHeaders_(ss.getSheetByName(APP.sh.suivi), ['Client', 'Bien', 'Date création', 'Canal', 'Statut client-bien', 'Score match', 'Dernière action', 'Auteur']);
  setHeaders_(ss.getSheetByName(APP.sh.commissions), ['Date vente', 'Bien', 'Prix vente', 'Commission %', 'Commission brute', 'Commission HT', 'Commission TVA', 'Commission nette']);
  setHeaders_(ss.getSheetByName(APP.sh.listes), ['Villes', 'Types bien', 'Statuts bien', 'DPE', 'État bien', 'Statuts client', 'Résultats visite', 'Statut client-bien']);
  setHeaders_(ss.getSheetByName(APP.sh.params), ['Paramètre', 'Valeur']);

  setHeaders_(ss.getSheetByName(APP.sh.dataBiens), ['Bien', 'Ville', 'Type', 'Prix affiché', 'Surface m²', 'DPE', 'État bien', 'Date fin mandat', 'Statut bien', 'Score global']);
  setHeaders_(ss.getSheetByName(APP.sh.dataClients), ['Client', 'Budget max', 'Ville recherchée', 'Type recherché', 'Surface min', 'Statut client', 'Dernier contact', 'Score activité', 'Priorité relance']);
  setHeaders_(ss.getSheetByName(APP.sh.dataSuivi), ['Client', 'Bien', 'Date création', 'Canal', 'Statut client-bien', 'Score match', 'Dernière action', 'Auteur']);
  setHeaders_(ss.getSheetByName(APP.sh.dataCom), ['Date vente', 'Bien', 'Prix vente', 'Commission %', 'Commission brute', 'Commission HT', 'Commission TVA', 'Commission nette']);
  setHeaders_(ss.getSheetByName(APP.sh.dataStats), ['Clé', 'Valeur']);
}

function remplirListes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.listes);
  sh.getRange('A2:A8').setValues([['Paris'], ['Lyon'], ['Marseille'], ['Toulouse'], ['Nantes'], ['Nice'], ['Bordeaux']]);
  sh.getRange('B2:B7').setValues([['Appartement'], ['Maison'], ['Terrain'], ['Local commercial'], ['Immeuble'], ['Studio']]);
  sh.getRange('C2:C7').setValues([['Prospection'], ['Mandat simple'], ['Mandat exclusif'], ['Sous offre'], ['Vendu'], ['Retiré']]);
  sh.getRange('D2:D8').setValues([['A'], ['B'], ['C'], ['D'], ['E'], ['F'], ['G']]);
  sh.getRange('E2:E5').setValues([['Neuf'], ['Bon'], ['À rafraîchir'], ['Travaux']]);
  sh.getRange('F2:F6').setValues([['Nouveau'], ['Actif'], ['Tiède'], ['Froid'], ['Signé']]);
  sh.getRange('G2:G4').setValues([['Visité'], ['Refusé'], ['À relancer']]);
  sh.getRange('H2:H6').setValues([['À proposer'], ['Visité'], ['Refusé'], ['Favori'], ['À relancer']]);
}

function remplirParametres_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  sh.getRange('A2:B12').setValues([
    ['Nom agent', ''],
    ['Nom agence', ''],
    ['Objectif mensuel (€)', 15000],
    ['Palier 1 max (€)', 60000],
    ['Palier 1 taux', 0.10],
    ['Palier 2 max (€)', 130000],
    ['Palier 2 taux', 0.15],
    ['Palier 3 taux', 0.20],
    ['TVA', 0.20],
    ['Seuil inactivité (jours)', 21],
    ['Alerte mandat (jours)', 30]
  ]);
  sh.getRange('B6').setNumberFormat('0,00%');
  sh.getRange('B8:B10').setNumberFormat('0,00%');
}

function setupValidationsV2_() {
  const ss = SpreadsheetApp.getActive();
  const list = ss.getSheetByName(APP.sh.listes);

  const vVille = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('A2:A200'), true).setAllowInvalid(false).build();
  const vType = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('B2:B200'), true).setAllowInvalid(false).build();
  const vStatutBien = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('C2:C200'), true).setAllowInvalid(false).build();
  const vDpe = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('D2:D200'), true).setAllowInvalid(false).build();
  const vEtat = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('E2:E200'), true).setAllowInvalid(false).build();
  const vStatutClient = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('F2:F200'), true).setAllowInvalid(false).build();
  const vSuivi = SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('H2:H200'), true).setAllowInvalid(false).build();

  const b = ss.getSheetByName(APP.sh.biens);
  b.getRange('C2:C5000').setDataValidation(vVille);
  b.getRange('D2:D5000').setDataValidation(vType);
  b.getRange('G2:G5000').setDataValidation(vDpe);
  b.getRange('H2:H5000').setDataValidation(vEtat);
  b.getRange('J2:J5000').setDataValidation(vStatutBien);

  const c = ss.getSheetByName(APP.sh.clients);
  c.getRange('F2:F5000').setDataValidation(vVille);
  c.getRange('G2:G5000').setDataValidation(vType);
  c.getRange('I2:I5000').setDataValidation(vStatutClient);

  const s = ss.getSheetByName(APP.sh.suivi);
  s.getRange('E2:E10000').setDataValidation(vSuivi);
}

function syncInterfaceVersDataBiens_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.biens, 15);
  const dst = ss.getSheetByName(APP.sh.dataBiens);
  dst.getRange('A2:J12000').clearContent();
  if (!src.length) return;
  const out = src.filter(r => String(r[0]).trim() !== '').map(r => [r[0], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[12]]);
  if (out.length) dst.getRange(2, 1, out.length, 10).setValues(out);
}

function syncInterfaceVersDataClients_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.clients, 12);
  const dst = ss.getSheetByName(APP.sh.dataClients);
  dst.getRange('A2:I12000').clearContent();
  if (!src.length) return;
  const out = src.filter(r => String(r[0]).trim() !== '').map(r => [r[0], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11] || 'Moyenne']);
  if (out.length) dst.getRange(2, 1, out.length, 9).setValues(out);
}

function syncDataSuiviVersInterface_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.dataSuivi, 8);
  const dst = ss.getSheetByName(APP.sh.suivi);
  dst.getRange('A2:H12000').clearContent();
  if (src.length) dst.getRange(2, 1, src.length, 8).setValues(src);
  trierSuivi_();
  appliquerStylesSuivi_();
}

function syncInterfaceSuiviVersData_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.suivi, 8);
  const dst = ss.getSheetByName(APP.sh.dataSuivi);
  dst.getRange('A2:H12000').clearContent();
  if (src.length) dst.getRange(2, 1, src.length, 8).setValues(src);
}

function appliquerDesignV4_() {
  designBiensV4_();
  designClientsV4_();
  designMatchingV4_();
  designSuiviV4_();
  designDashboardV4_();
  designCommissionsV4_();
  designParamsV4_();
}

function applyBaseSheetStyle_(sh, headerRow, zebraStart, widthMap) {
  if (!sh) return;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), headerRow + 1);

  sh.getRange(headerRow, 1, 1, lastCol)
    .setBackground(APP.ui.header)
    .setFontColor(APP.ui.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.setRowHeight(headerRow, 30);
  if (lastRow > zebraStart) applyZebra_(sh, zebraStart, lastRow, 1, lastCol);
  sh.getRange(1, 1, lastRow, lastCol).setBorder(false, false, false, false, false, false);
  sh.getRange(headerRow, 1, Math.min(lastRow - headerRow + 1, 400), lastCol)
    .setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);

  if (widthMap) Object.keys(widthMap).forEach(c => sh.setColumnWidth(Number(c), widthMap[c]));
}

function designBiensV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:260,2:125,3:140,4:130,5:130,6:120,7:85,8:110,9:125,10:120,11:85,12:85,13:95,14:150,15:220});
  sh.setFrozenRows(1);
  sh.getRange('C2:N5000').setHorizontalAlignment('center');
  sh.getRange('A2:A5000').setHorizontalAlignment('left');
  sh.getRange('I2:I5000').setNumberFormat('dd/mm/yyyy');
  sh.setRowHeights(2, Math.max(sh.getLastRow()-1,1), 27);
  ensureFilter_(sh, 1);
  appliquerStylesBiens_();
}

function designClientsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:230,2:120,3:130,4:220,5:120,6:135,7:130,8:100,9:115,10:125,11:105,12:115});
  sh.setFrozenRows(1);
  sh.getRange('F2:L5000').setHorizontalAlignment('center');
  sh.getRange('A2:A5000').setHorizontalAlignment('left');
  sh.getRange('J2:J5000').setNumberFormat('dd/mm/yyyy');
  sh.setRowHeights(2, Math.max(sh.getLastRow()-1,1), 27);
  ensureFilter_(sh, 1);
  appliquerStylesClients_();
}

function designMatchingV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.matching);
  if (!sh) return;
  const lastRow = Math.max(sh.getLastRow(), 6);
  sh.setFrozenRows(4);

  sh.getRange('A1:H1').setBackground(APP.ui.header).setFontColor(APP.ui.headerText).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('A4:H4').setBackground(APP.ui.header).setFontColor(APP.ui.headerText).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32); sh.setRowHeight(4, 30);

  if (lastRow >= 6) {
    applyZebra_(sh, 6, lastRow, 1, 5);
    applyZebra_(sh, 6, lastRow, 8, 1);
  }

  [1,2,3,4,5,6,7,8].forEach((c,i)=>sh.setColumnWidth(c,[270,130,120,120,90,110,145,145][i]));
  sh.getRange('B2:H700').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('A6:A700').setHorizontalAlignment('left');
  sh.getRange('D6:D700').setNumberFormat('#,##0.00 €');
  sh.getRange('H6:H700').setNumberFormat('dd/mm/yyyy hh:mm');
  ensureFilter_(sh, 4);
}

function designSuiviV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:220,2:250,3:140,4:95,5:140,6:95,7:145,8:210});
  sh.setFrozenRows(1);
  sh.getRange('C2:G10000').setHorizontalAlignment('center');
  sh.getRange('A2:B10000').setHorizontalAlignment('left');
  sh.setRowHeights(2, Math.max(sh.getLastRow()-1,1), 27);
  ensureFilter_(sh, 1);
  appliquerStylesSuivi_();
}

function designDashboardV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.dashboard);
  if (!sh) return;
  sh.setColumnWidths(1, 10, 130);
  sh.getRange('A1:J1').setBackground('#0f172a').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeights(1, 16, 30);
  sh.getRange('I3:J3').setBackground(APP.ui.header).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('I12:J12').setBackground(APP.ui.header).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('I4:J18').setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);
}

function designCommissionsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.commissions);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:120,2:250,3:120,4:95,5:130,6:130,7:130,8:130});
  sh.setFrozenRows(1);
  sh.getRange('A2:A5000').setNumberFormat('dd/mm/yyyy');
  sh.getRange('C2:H5000').setHorizontalAlignment('center');
  ensureFilter_(sh, 1);
}

function designParamsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:260,2:180});
}

function trierSuivi_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh || sh.getLastRow() < 3) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).sort([{ column: 7, ascending: false }]);
  ensureFilter_(sh, 1);
}

function appliquerStylesSuivi_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 5, n, 2).getValues();

  const bgs = vals.map(v => {
    const st = String(v[0] || '');
    if (st === 'Favori') return [APP.ui.success];
    if (st === 'À relancer') return [APP.ui.warning];
    if (st === 'Refusé') return [APP.ui.danger];
    if (st === 'Visité') return [APP.ui.info];
    return [APP.ui.neutral];
  });
  sh.getRange(2, 5, n, 1).setBackgrounds(bgs);
  sh.getRange(2, 3, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, 7, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, 6, n, 1).setNumberFormat('0');
}

function appliquerStylesBiens_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;

  // 1) Reset propre : zébrage de base + texte normal pour éviter les artefacts
  applyZebra_(sh, 2, sh.getLastRow(), 1, 15);
  sh.getRange(2, 1, n, 15).setFontColor('#111827');

  // 2) Lire les données puis appliquer styles spécifiques
  const vals = sh.getRange(2, 1, n, 15).getValues();

  const alertBg = [];
  const scoreBg = [];
  for (let i = 0; i < vals.length; i++) {
    const statut = String(vals[i][9] || '');
    const alert = String(vals[i][13] || '').toLowerCase();
    const score = toNum_(vals[i][12]);

    if (alert.includes('expiré') || alert.includes('❌')) alertBg.push([APP.ui.danger]);
    else if (alert.includes('expire') || alert.includes('⚠️')) alertBg.push([APP.ui.warning]);
    else if (alert === 'ok') alertBg.push([APP.ui.success]);
    else alertBg.push([i % 2 === 0 ? APP.ui.zebra1 : APP.ui.zebra2]);

    scoreBg.push([score >= 5.5 ? '#ecfccb' : (i % 2 === 0 ? APP.ui.zebra1 : APP.ui.zebra2)]);

    if (statut === 'Vendu' || statut === 'Retiré') {
      sh.getRange(i + 2, 1, 1, 15).setBackground(APP.ui.neutral).setFontColor(APP.ui.textMuted);
    }
  }

  // 3) Recoloration ciblée après la couche de base
  sh.getRange(2, 14, n, 1).setBackgrounds(alertBg);
  sh.getRange(2, 13, n, 1).setBackgrounds(scoreBg);
  sh.getRange(2, 9, n, 1).setNumberFormat('dd/mm/yyyy');
}

function appliquerStylesClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 11, n, 1).getValues();
  const bg = vals.map(v => {
    const score = toNum_(v[0]);
    if (score >= 70) return [APP.ui.success];
    if (score >= 40) return [APP.ui.warning];
    return [APP.ui.danger];
  });
  sh.getRange(2, 11, n, 1).setBackgrounds(bg);
  sh.getRange(2, 10, n, 1).setNumberFormat('dd/mm/yyyy');
}

function ensureFilter_(sh, headerRow) {
  if (!sh) return;
  const h = headerRow || 1;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), h);

  const existing = sh.getFilter();
  if (existing) existing.remove();

  sh.getRange(h, 1, lastRow - h + 1, lastCol).createFilter();
}

function applyZebra_(sh, startRow, endRow, startCol, numCols) {
  if (endRow < startRow) return;
  const bgs = [];
  for (let r = startRow; r <= endRow; r++) bgs.push(Array(numCols).fill(r % 2 === 0 ? APP.ui.zebra1 : APP.ui.zebra2));
  sh.getRange(startRow, startCol, bgs.length, numCols).setBackgrounds(bgs);
}

function protegerData_() {
  const ss = SpreadsheetApp.getActive();
  [APP.sh.listes, APP.sh.dataBiens, APP.sh.dataClients, APP.sh.dataSuivi, APP.sh.dataCom, APP.sh.dataStats].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    if (sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) return;
    const p = sh.protect().setDescription('Protection DATA auto V4');
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

function getRows_(sheetName, cols) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}

function setHeaders_(sh, headers) {
  const maxCols = Math.max(sh.getLastColumn(), headers.length);
  sh.getRange(1, 1, 1, maxCols).clearContent();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getParams_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  const vals = sh.getRange('A2:B12').getValues();
  const m = {};
  vals.forEach(r => m[String(r[0])] = r[1]);
  return {
    objectifMensuel: toNum_(m['Objectif mensuel (€)']) || 15000,
    p1max: toNum_(m['Palier 1 max (€)']) || 60000,
    p1taux: toNum_(m['Palier 1 taux']) || 0.10,
    p2max: toNum_(m['Palier 2 max (€)']) || 130000,
    p2taux: toNum_(m['Palier 2 taux']) || 0.15,
    p3taux: toNum_(m['Palier 3 taux']) || 0.20,
    tva: toNum_(m['TVA']) || 0.20,
    inact: toNum_(m['Seuil inactivité (jours)']) || 21,
    alMandat: toNum_(m['Alerte mandat (jours)']) || 30
  };
}

function formatEuro_(n) {
  const v = Number(n || 0);
  return v.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
}

function toNum_(x) {
  if (typeof x === 'number') return x;
  if (typeof x === 'string') {
    const n = Number(String(x).replace(/\s/g, '').replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

```

---

## 3) Menus déroulants + validations + automatisations (résumé)

### Menus déroulants actifs
- `INTERFACE_BIENS`
  - Ville, Type, DPE, État bien, Statut bien
- `INTERFACE_CLIENTS`
  - Ville recherchée, Type recherché, Statut client
- `INTERFACE_MATCHING`
  - Client sélectionné (`B2`)
  - Statut client-bien (`G6:G...`) : À proposer / Visité / Refusé / Favori / À relancer
- `INTERFACE_SUIVI`
  - Statut client-bien

### Automatisations clés
- Date de création auto (biens/clients)
- Date dernier contact auto lors d’une action matching
- Score DPE, score état, score global bien
- Score activité client basé sur actions réelles
- Matching dynamique par client (trié + couleur)
- Suivi client-bien indépendant (DATA dédiée)
- Dashboard terrain auto

---

## 4) Guide d’installation simple (pas à pas)

1. Créer un Google Sheet vide.
2. Ouvrir **Extensions > Apps Script**.
3. Coller le script complet dans `Code.gs`.
4. Enregistrer puis recharger le Sheet.
5. Dans le menu **📊 CRM Immobilier V2**, cliquer **🚀 Initialiser CRM V2**.
6. Aller dans `PARAMETRES` pour définir objectif, paliers, TVA.
7. Mettre à jour les listes déroulantes dans `DATA_LISTES` (villes, types, statuts).
8. Saisir vos biens dans `INTERFACE_BIENS`.
9. Saisir vos clients dans `INTERFACE_CLIENTS`.
10. Ouvrir `INTERFACE_MATCHING`, choisir un client en `B2`.
11. Le système affiche automatiquement les biens compatibles triés.
12. Mettre à jour le statut par ligne (Visité/Refusé/Favori/À relancer).
13. Vérifier `INTERFACE_SUIVI` (historique complet client-bien).
14. Consulter `INTERFACE_DASHBOARD` pour le pilotage quotidien.

---

## 5) Pourquoi cette V2 est plus "pro, simple et rapide"

- Interface lisible (colonnes compréhensibles, moins d’IDs techniques en façade)
- Matching exploitable en rendez-vous en 2 clics (choix client + statut)
- Historique indépendant par client/bien (vrai comportement CRM métier)
- Menus déroulants partout (moins d’erreurs, saisie rapide)
- Dashboard orienté action terrain (visites, refus, favoris, tops, objectif)
- Toujours duplicable en mono-agent (copie du fichier + adaptation `PARAMETRES`)
