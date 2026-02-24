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
 * CRM Immobilier V2 - Google Sheets
 * Orienté utilisateur : rapide, simple, duplicable
 * Un seul fichier Code.gs
 */

const APP = {
  menu: '📊 CRM Immobilier V2',
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
  colors: {
    header: '#111827',
    headerText: '#ffffff',
    zebra1: '#ffffff',
    zebra2: '#f8fafc',
    green: '#dcfce7',
    orange: '#ffedd5',
    red: '#fee2e2',
    kpiDark: '#0f172a'
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP.menu)
    .addItem('🚀 Initialiser CRM V2', 'initialiserCRMV2')
    .addSeparator()
    .addItem('➕ Ajouter un client', 'ajouterClientRapide')
    .addItem('🏠 Ajouter un bien', 'ajouterBienRapide')
    .addSeparator()
    .addItem('🎯 Recalculer matching client', 'recalculerMatchingClient')
    .addItem('📊 Mettre à jour dashboard', 'majDashboardV2')
    .addItem('🎨 Appliquer design pro', 'appliquerDesignV2')
    .addToUi();
}

function initialiserCRMV2() {
  const ss = SpreadsheetApp.getActive();
  creerOngletsSiAbsents_(ss);
  definirEntetes_();
  remplirListes_();
  remplirParametres_();
  setupValidationsV2_();
  appliquerDesignV2();
  protegerData_();
  majScoresBiens_();
  majScoreActiviteClients_();
  genererFicheMatching_();
  majDashboardV2();
  SpreadsheetApp.getUi().alert('CRM V2 initialisé ✅');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (name === APP.sh.biens) {
    if (col >= 1 && col <= 12) {
      autoDateCreation_(sh, row, 1, 2);
      majScoreBienLigne_(row);
      syncInterfaceVersDataBiens_();
    }
  }

  if (name === APP.sh.clients) {
    if (col >= 1 && col <= 11) {
      autoDateCreation_(sh, row, 1, 2);
      majScoreClientLigne_(row);
      syncInterfaceVersDataClients_();
    }
  }

  if (name === APP.sh.matching) {
    // B2 = client sélectionné (menu)
    if (row === 2 && col === 2) {
      genererFicheMatching_();
      return;
    }

    // Colonne statut client-bien (G)
    if (row >= 6 && col === 7) {
      enregistrerStatutDepuisMatching_(row);
      majScoreClientDepuisSuivi_(row);
      majDashboardV2();
    }
  }

  if (name === APP.sh.suivi) {
    // si l'utilisateur modifie directement le suivi
    syncInterfaceSuiviVersData_();
    majScoreActiviteClients_();
    majDashboardV2();
  }

  if (name === APP.sh.commissions) {
    recalculCommissionsV2_();
    majDashboardV2();
  }
}

// =========================
// ACTIONS RAPIDES MENU
// =========================

function ajouterClientRapide() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const rep = ui.prompt('Nouveau client', 'Nom du client :', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const nom = String(rep.getResponseText() || '').trim();
  if (!nom) return;

  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 11).setValues([[
    nom, new Date(), '', '', '', '', '', '', 'Nouveau', '', 0
  ]]);
  syncInterfaceVersDataClients_();
  appliquerDesignV2();
}

function ajouterBienRapide() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const rep = ui.prompt('Nouveau bien', 'Adresse du bien :', ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const adresse = String(rep.getResponseText() || '').trim();
  if (!adresse) return;

  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 12).setValues([[
    adresse, new Date(), '', '', '', '', '', '', '', 'Prospection', '', ''
  ]]);
  syncInterfaceVersDataBiens_();
  appliquerDesignV2();
}

function recalculerMatchingClient() {
  genererFicheMatching_();
  SpreadsheetApp.getUi().alert('Matching client recalculé ✅');
}

// =========================
// MATCHING V2 (PRIORITÉ)
// =========================

function genererFicheMatching_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.matching);
  const clients = getRows_(APP.sh.clients, 11);
  const biens = getRows_(APP.sh.biens, 12);
  const suivi = getRows_(APP.sh.dataSuivi, 8);

  // Structure haute lisible
  sh.getRange('A1:H1').merge().setValue('FICHE MATCHING CLIENT').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('Client sélectionné');
  sh.getRange('A4:H4').setValues([['Bien', 'Ville', 'Type', 'Prix', 'Score', 'Niveau', 'Statut client-bien', 'Dernière action']]);

  // Validation liste clients sur B2
  const listClientNames = clients.map(r => r[0]).filter(Boolean);
  if (listClientNames.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(listClientNames, true).setAllowInvalid(false).build();
    sh.getRange('B2').setDataValidation(rule);
  }

  const clientNom = String(sh.getRange('B2').getValue() || '');
  if (!clientNom) {
    sh.getRange('A6:H500').clearContent().clearDataValidations().setBackground(null);
    return;
  }

  const c = clients.find(r => String(r[0]) === clientNom);
  if (!c) return;

  const budget = toNum_(c[5]);
  const villeRecherche = String(c[6] || '').toLowerCase();
  const typeRecherche = String(c[7] || '').toLowerCase();
  const surfaceMin = toNum_(c[8]);

  const rows = [];
  for (let i = 0; i < biens.length; i++) {
    const b = biens[i];
    const statutBien = String(b[9] || '');
    if (['Vendu', 'Retiré', 'Expiré'].includes(statutBien)) continue;

    const score = calculScoreMatching_(
      {
        ville: String(b[2] || ''),
        type: String(b[3] || ''),
        prix: toNum_(b[4]),
        surface: toNum_(b[5])
      },
      {
        villeRecherche,
        typeRecherche,
        budget,
        surfaceMin
      }
    );

    const niveau = score >= 80 ? 'Très bon' : score >= 60 ? 'Moyen' : 'Faible';
    const statutParDefaut = 'À proposer';

    const key = keySuivi_(clientNom, String(b[0]));
    const found = suivi.find(s => keySuivi_(String(s[0]), String(s[1])) === key);
    const statut = found ? String(found[4] || statutParDefaut) : statutParDefaut;
    const dateAct = found ? found[6] : '';

    rows.push([
      b[0], // Bien (adresse)
      b[2], // Ville
      b[3], // Type
      b[4], // Prix
      score,
      niveau,
      statut,
      dateAct
    ]);
  }

  rows.sort((a, b) => b[4] - a[4]);

  sh.getRange('A6:H500').clearContent().clearDataValidations().setBackground(null);
  if (!rows.length) return;
  sh.getRange(6, 1, rows.length, 8).setValues(rows);
  sh.getRange(6, 4, rows.length, 1).setNumberFormat('#,##0.00 €');

  // Dropdown statut client-bien
  const ruleStatut = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(APP.sh.listes).getRange('H2:H10'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(6, 7, rows.length, 1).setDataValidation(ruleStatut);

  // Couleur par niveau de match
  const levelRng = sh.getRange(6, 6, rows.length, 1);
  const lv = levelRng.getValues();
  const bgs = lv.map(r => {
    if (r[0] === 'Très bon') return [APP.colors.green];
    if (r[0] === 'Moyen') return [APP.colors.orange];
    return [APP.colors.red];
  });
  levelRng.setBackgrounds(bgs);

  // Couleur score
  sh.getRange(6, 5, rows.length, 1).setNumberFormat('0');

  appliquerDesignV2();
}

function calculScoreMatching_(bien, crit) {
  // Pondération métier
  // Ville 35%, Budget 35%, Type 20%, Surface 10%
  let ville = 0;
  let budget = 0;
  let type = 0;
  let surface = 10; // bonus neutre si non utilisé

  const vBien = String(bien.ville || '').toLowerCase();
  if (crit.villeRecherche && vBien === crit.villeRecherche) ville = 35;

  const p = toNum_(bien.prix);
  const b = toNum_(crit.budget);
  if (b > 0) {
    if (p <= b) budget = 35;
    else {
      const depassement = (p - b) / b;
      budget = depassement < 0.15 ? 22 : depassement < 0.30 ? 10 : 0;
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

  if (idx === -1) {
    d.appendRow([clientNom, bienAdresse, now, '', statut, score, now, Session.getActiveUser().getEmail() || '']);
  } else {
    d.getRange(idx, 5).setValue(statut);
    d.getRange(idx, 6).setValue(score);
    d.getRange(idx, 7).setValue(now);
  }

  // miroir interface suivi
  syncDataSuiviVersInterface_();
}

function majScoreClientDepuisSuivi_(matchingRow) {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getSheetByName(APP.sh.matching);
  const clientNom = String(m.getRange('B2').getValue() || '');
  if (!clientNom) return;

  const statut = String(m.getRange(matchingRow, 7).getValue() || '');
  const clientsSh = ss.getSheetByName(APP.sh.clients);
  const clients = getRows_(APP.sh.clients, 11);

  const idx = clients.findIndex(r => String(r[0]) === clientNom);
  if (idx === -1) return;

  // dernier contact automatique sur action
  clientsSh.getRange(idx + 2, 10).setValue(new Date());

  // score activité recalculé global
  majScoreClientLigne_(idx + 2);

  if (statut === 'Visité' || statut === 'Refusé' || statut === 'Favori' || statut === 'À relancer') {
    syncInterfaceVersDataClients_();
  }
}

function keySuivi_(clientNom, bienAdresse) {
  return `${clientNom}|||${bienAdresse}`;
}

// =========================
// SCORES & AUTOMATISATIONS
// =========================

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
}

function majScoreBienLigne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const dpe = String(sh.getRange(row, 11).getValue() || '').toUpperCase();
  const etat = String(sh.getRange(row, 12).getValue() || '');

  const scoreDPE = ({ A: 7, B: 6, C: 5, D: 4, E: 3, F: 2, G: 1 }[dpe]) || 0;
  const scoreEtat = ({ 'Neuf': 4, 'Bon': 3, 'À rafraîchir': 2, 'Travaux': 1 }[etat]) || 0;
  const scoreGlobal = Math.round((scoreDPE * 0.55 + scoreEtat * 0.45) * 10) / 10;

  sh.getRange(row, 13).setValue(scoreDPE);
  sh.getRange(row, 14).setValue(scoreEtat);
  sh.getRange(row, 15).setValue(scoreGlobal);
}

function majScoreActiviteClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const last = sh.getLastRow();
  if (last < 2) return;
  for (let r = 2; r <= last; r++) majScoreClientLigne_(r);
}

function majScoreClientLigne_(row) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.clients);
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
  if (dernierContact instanceof Date) {
    inact = Math.max(0, Math.floor((new Date() - dernierContact) / (1000 * 3600 * 24)));
  }

  const score = Math.max(0, Math.min(100, Math.round(visites * 12 + favoris * 10 - refus * 8 + (30 - Math.min(30, inact)))));
  sh.getRange(row, 11).setValue(score);
}

// =========================
// DASHBOARD V2 (TERRAIN)
// =========================

function majDashboardV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.dashboard);

  const biens = getRows_(APP.sh.biens, 15);
  const clients = getRows_(APP.sh.clients, 11);
  const suivi = getRows_(APP.sh.dataSuivi, 8);
  const coms = getRows_(APP.sh.commissions, 8);
  const p = getParams_();

  const clientsActifs = clients.filter(c => ['Actif', 'Tiède'].includes(String(c[9]))).length;
  const biensActifs = biens.filter(b => !['Vendu', 'Retiré', 'Expiré'].includes(String(b[9]))).length;

  const mois = new Date().getMonth();
  const annee = new Date().getFullYear();
  const actionsMois = suivi.filter(r => r[6] instanceof Date && r[6].getMonth() === mois && r[6].getFullYear() === annee);

  const visitesMois = actionsMois.filter(a => String(a[4]) === 'Visité').length;
  const refusMois = actionsMois.filter(a => String(a[4]) === 'Refusé').length;
  const favoris = actionsMois.filter(a => String(a[4]) === 'Favori').length;
  const tauxRefus = visitesMois ? refusMois / visitesMois : 0;

  const caMensuel = coms
    .filter(c => c[0] instanceof Date && c[0].getMonth() === mois && c[0].getFullYear() === annee)
    .reduce((a, c) => a + toNum_(c[7]), 0);

  sh.clear();
  sh.getRange('A1:J1').merge().setValue('DASHBOARD TERRAIN - CRM IMMOBILIER V2').setFontWeight('bold').setFontSize(16);

  const kpi = [
    ['Clients actifs', clientsActifs],
    ['Biens actifs', biensActifs],
    ['Visites (mois)', visitesMois],
    ['Refus (mois)', refusMois],
    ['Taux refus', tauxRefus],
    ['Favoris (mois)', favoris],
    ['CA mensuel net', caMensuel],
    ['Objectif', p.objectifMensuel],
    ['Objectif atteint', caMensuel >= p.objectifMensuel ? 'Oui ✅' : 'Non ⏳']
  ];

  sh.getRange(3, 1, kpi.length, 2).setValues(kpi);
  sh.getRange('B7').setNumberFormat('0,00%');
  sh.getRange('B9:B10').setNumberFormat('#,##0.00 €');

  // Top clients actifs
  sh.getRange('D3').setValue('Top clients actifs');
  const mapClients = mapCountByStatus_(suivi, 0, ['Visité', 'Favori', 'À relancer']);
  writeMapSorted_(sh, 4, 4, mapClients, 5);

  // Top biens visités
  sh.getRange('G3').setValue('Top biens visités');
  const mapBiens = mapCountByStatus_(suivi, 1, ['Visité']);
  writeMapSorted_(sh, 4, 7, mapBiens, 5);

  buildChartsDashboardV2_(sh);
  styleDashboardV2_(sh, caMensuel >= p.objectifMensuel, tauxRefus);
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
    .addRange(sh.getRange('A3:B8'))
    .setOption('title', 'KPI opérationnels')
    .setPosition(13, 1, 0, 0)
    .build();
  sh.insertChart(c1);

  const c2 = sh.newChart()
    .asBarChart()
    .addRange(sh.getRange('D4:E8'))
    .setOption('title', 'Top clients actifs')
    .setPosition(13, 6, 0, 0)
    .build();
  sh.insertChart(c2);

  const c3 = sh.newChart()
    .asBarChart()
    .addRange(sh.getRange('G4:H8'))
    .setOption('title', 'Top biens visités')
    .setPosition(26, 6, 0, 0)
    .build();
  sh.insertChart(c3);
}

// =========================
// COMMISSIONS
// =========================

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

// =========================
// STRUCTURE + VALIDATIONS
// =========================

function creerOngletsSiAbsents_(ss) {
  Object.values(APP.sh).forEach(n => {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });
}

function definirEntetes_() {
  const ss = SpreadsheetApp.getActive();

  setHeaders_(ss.getSheetByName(APP.sh.biens), [
    'Bien', 'Date création', 'Ville', 'Type', 'Prix affiché', 'Surface m²', 'DPE', 'État bien', 'Date fin mandat', 'Statut bien', 'Score DPE', 'Score État', 'Score global', 'Alerte mandat', 'Note interne'
  ]);

  setHeaders_(ss.getSheetByName(APP.sh.clients), [
    'Client', 'Date création', 'Téléphone', 'Email', 'Budget max', 'Ville recherchée', 'Type recherché', 'Surface min', 'Statut client', 'Dernier contact', 'Score activité'
  ]);

  setHeaders_(ss.getSheetByName(APP.sh.matching), ['Zone', 'Valeur']);

  setHeaders_(ss.getSheetByName(APP.sh.suivi), [
    'Client', 'Bien', 'Date création', 'Canal', 'Statut client-bien', 'Score match', 'Dernière action', 'Auteur'
  ]);

  setHeaders_(ss.getSheetByName(APP.sh.commissions), [
    'Date vente', 'Bien', 'Prix vente', 'Commission %', 'Commission brute', 'Commission HT', 'Commission TVA', 'Commission nette'
  ]);

  setHeaders_(ss.getSheetByName(APP.sh.listes), [
    'Villes', 'Types bien', 'Statuts bien', 'DPE', 'État bien', 'Statuts client', 'Résultats visite', 'Statut client-bien'
  ]);

  setHeaders_(ss.getSheetByName(APP.sh.params), ['Paramètre', 'Valeur']);

  setHeaders_(ss.getSheetByName(APP.sh.dataBiens), ['Bien', 'Ville', 'Type', 'Prix affiché', 'Surface m²', 'DPE', 'État bien', 'Date fin mandat', 'Statut bien', 'Score global']);
  setHeaders_(ss.getSheetByName(APP.sh.dataClients), ['Client', 'Budget max', 'Ville recherchée', 'Type recherché', 'Surface min', 'Statut client', 'Dernier contact', 'Score activité']);
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
  sh.getRange('B6:B10').setNumberFormat('0,00%');
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
  s.getRange('E2:E8000').setDataValidation(vSuivi);
}

// =========================
// SYNCHRO INTERFACE <-> DATA
// =========================

function syncInterfaceVersDataBiens_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.biens, 15);
  const dst = ss.getSheetByName(APP.sh.dataBiens);
  dst.getRange('A2:J10000').clearContent();
  if (!src.length) return;

  const out = src
    .filter(r => String(r[0]).trim() !== '')
    .map(r => [r[0], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[12]]);

  if (out.length) dst.getRange(2, 1, out.length, 10).setValues(out);
}

function syncInterfaceVersDataClients_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.clients, 11);
  const dst = ss.getSheetByName(APP.sh.dataClients);
  dst.getRange('A2:H10000').clearContent();
  if (!src.length) return;

  const out = src
    .filter(r => String(r[0]).trim() !== '')
    .map(r => [r[0], r[4], r[5], r[6], r[7], r[8], r[9], r[10]]);

  if (out.length) dst.getRange(2, 1, out.length, 8).setValues(out);
}

function syncDataSuiviVersInterface_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.dataSuivi, 8);
  const dst = ss.getSheetByName(APP.sh.suivi);
  dst.getRange('A2:H10000').clearContent();
  if (src.length) dst.getRange(2, 1, src.length, 8).setValues(src);
}

function syncInterfaceSuiviVersData_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.suivi, 8);
  const dst = ss.getSheetByName(APP.sh.dataSuivi);
  dst.getRange('A2:H10000').clearContent();
  if (src.length) dst.getRange(2, 1, src.length, 8).setValues(src);
}

// =========================
// DESIGN
// =========================

function appliquerDesignV2() {
  const ss = SpreadsheetApp.getActive();
  const interfaces = [APP.sh.dashboard, APP.sh.biens, APP.sh.clients, APP.sh.matching, APP.sh.suivi, APP.sh.commissions, APP.sh.params];

  interfaces.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;

    const lastCol = Math.max(sh.getLastColumn(), 8);
    const lastRow = Math.max(sh.getLastRow(), 2);

    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, lastCol)
      .setBackground(APP.colors.header)
      .setFontColor(APP.colors.headerText)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    const body = sh.getRange(2, 1, lastRow - 1, lastCol);
    const bgs = [];
    for (let r = 0; r < lastRow - 1; r++) {
      const line = [];
      for (let c = 0; c < lastCol; c++) line.push(r % 2 === 0 ? APP.colors.zebra1 : APP.colors.zebra2);
      bgs.push(line);
    }
    body.setBackgrounds(bgs)
      .setVerticalAlignment('middle');

    // centrer les colonnes importantes
    if (n === APP.sh.biens) sh.getRange('C2:J5000').setHorizontalAlignment('center');
    if (n === APP.sh.clients) sh.getRange('F2:I5000').setHorizontalAlignment('center');
    if (n === APP.sh.matching) sh.getRange('B2:H500').setHorizontalAlignment('center');

    sh.autoResizeColumns(1, Math.min(lastCol, 12));
    sh.setColumnWidth(1, Math.max(180, sh.getColumnWidth(1)));

    // bordures propres sur zone utile
    sh.getRange(1, 1, Math.min(lastRow, 200), Math.min(lastCol, 12)).setBorder(true, true, true, true, true, true, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
  });

  styleDashboardV2_(ss.getSheetByName(APP.sh.dashboard), false, 0);
}

function styleDashboardV2_(sh, objectifAtteint, tauxRefus) {
  if (!sh) return;
  sh.getRange('A1:J1').setBackground(APP.colors.kpiDark).setFontColor('#fff').setHorizontalAlignment('center');
  sh.getRange('A3:B11').setBorder(true, true, true, true, true, true);

  if (objectifAtteint) sh.getRange('B11').setBackground(APP.colors.green);
  else sh.getRange('B11').setBackground(APP.colors.orange);

  if (tauxRefus > 0.5) sh.getRange('B7').setBackground(APP.colors.red);
  else if (tauxRefus > 0.3) sh.getRange('B7').setBackground(APP.colors.orange);
  else sh.getRange('B7').setBackground(APP.colors.green);
}

function protegerData_() {
  const ss = SpreadsheetApp.getActive();
  [APP.sh.listes, APP.sh.dataBiens, APP.sh.dataClients, APP.sh.dataSuivi, APP.sh.dataCom, APP.sh.dataStats].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    if (sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) return;
    const p = sh.protect().setDescription('Protection DATA auto V2');
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

// =========================
// UTILS
// =========================

function getRows_(sheetName, cols) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}

function setHeaders_(sh, headers) {
  sh.clear();
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
