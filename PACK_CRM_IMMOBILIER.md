# CRM Immobilier Google Sheets + Apps Script (mono-agent duplicable)

## 1) Structure des onglets

### A. Onglets à créer (ordre recommandé)

1. **INTERFACE_DASHBOARD**
2. **INTERFACE_BIENS**
3. **INTERFACE_CLIENTS**
4. **INTERFACE_VISITES**
5. **INTERFACE_MATCHING**
6. **INTERFACE_COMMISSIONS**
7. **PARAMETRES**
8. **DATA_LISTES** (protégé)
9. **DATA_BIENS** (protégé)
10. **DATA_CLIENTS** (protégé)
11. **DATA_VISITES** (protégé)
12. **DATA_MATCHING** (protégé)
13. **DATA_COMMISSIONS** (protégé)
14. **DATA_STATS** (protégé)

> Principe : les onglets `INTERFACE_*` sont éditables au quotidien, les onglets `DATA_*` servent de base structurée, calcul et historique.

---

### B. Colonnes recommandées

#### INTERFACE_BIENS / DATA_BIENS
- A `ID_BIEN`
- B `Date_entrée`
- C `Adresse`
- D `Ville`
- E `Code_postal`
- F `Type_bien`
- G `Surface_m2`
- H `Prix_affiché`
- I `Statut` (Prospection / Mandat simple / Mandat exclusif / Compromis / Vendu / Expiré)
- J `Date_expiration_mandat`
- K `DPE` (A→G)
- L `Etat_bien` (Neuf / Bon / À rafraîchir / Travaux)
- M `Score_DPE` (auto)
- N `Score_Etat` (auto)
- O `Score_Global_Bien` (auto)
- P `Alerte`

#### INTERFACE_CLIENTS / DATA_CLIENTS
- A `ID_CLIENT`
- B `Date_création`
- C `Nom_prenom`
- D `Téléphone`
- E `Email`
- F `Budget_max`
- G `Ville_recherche`
- H `Type_recherche`
- I `Statut_client` (Nouveau / Actif / Tiède / Froid / Signé / Perdu)
- J `Dernier_contact`
- K `Nb_visites` (auto)
- L `Nb_refus` (auto)
- M `Score_activité` (auto)
- N `Alerte`

#### INTERFACE_VISITES / DATA_VISITES
- A `ID_VISITE`
- B `Date_visite`
- C `ID_CLIENT`
- D `ID_BIEN`
- E `Résultat` (Intéressé / Refus / À relancer)
- F `Commentaire`

#### INTERFACE_MATCHING / DATA_MATCHING
- A `ID_MATCH`
- B `ID_CLIENT`
- C `ID_BIEN`
- D `Score_budget` (0-100)
- E `Score_ville` (0-100)
- F `Score_type` (0-100)
- G `Score_compatibilité` (pondéré auto)
- H `Priorité`

#### INTERFACE_COMMISSIONS / DATA_COMMISSIONS
- A `Date_vente`
- B `ID_BIEN`
- C `Prix_vente`
- D `Commission_%`
- E `Commission_brute`
- F `Commission_HT`
- G `Commission_TVA`
- H `Commission_Nette`
- I `Palier_appliqué`

#### PARAMETRES
- B2 `Nom agent`
- B3 `Nom agence`
- B4 `Objectif mensuel (€)`
- B6 `Palier 1 max (€)` = 60000
- B7 `Palier 1 taux` = 10%
- B8 `Palier 2 max (€)` = 130000
- B9 `Palier 2 taux` = 15%
- B10 `Palier 3 taux` = 20%
- B11 `TVA` = 20%
- B13 `Seuil inactivité client (jours)` = 21
- B14 `Seuil alerte mandat (jours avant expiration)` = 30

#### DATA_LISTES
- A2:A : statuts biens
- B2:B : types biens
- C2:C : DPE
- D2:D : états bien
- E2:E : statuts clients
- F2:F : résultats visite

---

## 2) Apps Script complet

> Créez **un seul fichier** `Code.gs` et collez tout le script suivant.

```javascript
/**
 * CRM Immobilier - Google Sheets
 * Mono-agent, duplicable, performant
 * Tout-en-un dans un seul fichier
 */

const CFG = {
  menuName: '📊 CRM Immobilier',
  sheets: {
    dashboard: 'INTERFACE_DASHBOARD',
    biens: 'INTERFACE_BIENS',
    clients: 'INTERFACE_CLIENTS',
    visites: 'INTERFACE_VISITES',
    matching: 'INTERFACE_MATCHING',
    commissions: 'INTERFACE_COMMISSIONS',
    params: 'PARAMETRES',
    dataListes: 'DATA_LISTES',
    dataBiens: 'DATA_BIENS',
    dataClients: 'DATA_CLIENTS',
    dataVisites: 'DATA_VISITES',
    dataMatching: 'DATA_MATCHING',
    dataCommissions: 'DATA_COMMISSIONS',
    dataStats: 'DATA_STATS'
  },
  col: {
    // Biens
    bien: {
      id: 1, dateEntree: 2, statut: 9, expiration: 10, dpe: 11, etat: 12,
      scoreDpe: 13, scoreEtat: 14, scoreGlobal: 15, alerte: 16
    },
    // Clients
    client: {
      id: 1, dateCreation: 2, statut: 9, dernierContact: 10,
      nbVisites: 11, nbRefus: 12, scoreActivite: 13, alerte: 14
    },
    // Visites
    visite: {
      id: 1, dateVisite: 2, idClient: 3, idBien: 4, resultat: 5
    },
    // Matching
    match: {
      scoreBudget: 4, scoreVille: 5, scoreType: 6, scoreCompat: 7, priorite: 8
    },
    // Commissions
    commission: {
      prixVente: 3, taux: 4, brute: 5, ht: 6, tva: 7, nette: 8, palier: 9
    }
  },
  ui: {
    headerBg: '#1f2937',
    headerFont: '#ffffff',
    zebra1: '#ffffff',
    zebra2: '#f8fafc',
    ok: '#dcfce7',
    warn: '#fef3c7',
    danger: '#fee2e2'
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CFG.menuName)
    .addItem('🚀 Initialiser CRM', 'initialiserCRM')
    .addSeparator()
    .addItem('🔁 Rafraîchir scores & alertes', 'refreshAll')
    .addItem('🎯 Recalcul matching', 'recalculMatching')
    .addItem('💶 Recalcul commissions', 'recalculCommissions')
    .addItem('📊 Mettre à jour dashboard', 'majDashboard')
    .addSeparator()
    .addItem('🎨 Appliquer design pro', 'appliquerDesignPro')
    .addItem('🔒 Protéger onglets DATA', 'protegerDataSheets')
    .addToUi();
}

function initialiserCRM() {
  const ss = SpreadsheetApp.getActive();
  createMissingSheets_(ss);
  seedParametres_();
  seedListes_();
  setupHeaders_();
  setupValidations_();
  appliquerDesignPro();
  protegerDataSheets();
  refreshAll();
  SpreadsheetApp.getUi().alert('CRM initialisé ✅');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  const row = e.range.getRow();
  if (row < 2) return;

  if (name === CFG.sheets.biens) {
    majLigneBien_(row);
  }

  if (name === CFG.sheets.clients) {
    majLigneClient_(row);
  }

  if (name === CFG.sheets.visites) {
    syncCompteursDepuisVisites_();
  }

  if (name === CFG.sheets.commissions) {
    recalculCommissions();
  }

  if ([CFG.sheets.biens, CFG.sheets.clients, CFG.sheets.visites, CFG.sheets.commissions].includes(name)) {
    majDashboard();
  }
}

function refreshAll() {
  majScoresBiens_();
  syncCompteursDepuisVisites_();
  majScoresClients_();
  recalculMatching();
  recalculCommissions();
  majDashboard();
}

function recalculMatching() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheets.matching);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < 2) return;

  const vals = sh.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < vals.length; i++) {
    const sb = toNum_(vals[i][3]);
    const sv = toNum_(vals[i][4]);
    const st = toNum_(vals[i][5]);
    const compat = Math.round(sb * 0.5 + sv * 0.3 + st * 0.2);
    vals[i][6] = compat;
    vals[i][7] = compat >= 80 ? 'Haute' : compat >= 60 ? 'Moyenne' : 'Basse';
  }

  sh.getRange(2, 1, vals.length, 8).setValues(vals);
  colorPrioriteMatching_(sh, vals.length);
}

function recalculCommissions() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheets.commissions);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < 2) return;

  const p = getParams_();
  const data = sh.getRange(2, 1, last - 1, 9).getValues();

  for (let i = 0; i < data.length; i++) {
    const prix = toNum_(data[i][CFG.col.commission.prixVente - 1]);
    if (!prix) continue;

    const taux = prix <= p.palier1Max ? p.palier1Taux : (prix <= p.palier2Max ? p.palier2Taux : p.palier3Taux);
    const brute = prix * taux;
    const ht = brute / (1 + p.tva);
    const tvaMontant = brute - ht;
    const nette = ht;

    data[i][CFG.col.commission.taux - 1] = taux;
    data[i][CFG.col.commission.brute - 1] = round2_(brute);
    data[i][CFG.col.commission.ht - 1] = round2_(ht);
    data[i][CFG.col.commission.tva - 1] = round2_(tvaMontant);
    data[i][CFG.col.commission.nette - 1] = round2_(nette);
    data[i][CFG.col.commission.palier - 1] = taux === p.palier1Taux ? 'Palier 1' : (taux === p.palier2Taux ? 'Palier 2' : 'Palier 3');
  }

  sh.getRange(2, 1, data.length, 9).setValues(data);
}

function majDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheets.dashboard);
  if (!sh) return;

  const biens = getRows_(CFG.sheets.biens, 16);
  const clients = getRows_(CFG.sheets.clients, 14);
  const visites = getRows_(CFG.sheets.visites, 6);
  const commissions = getRows_(CFG.sheets.commissions, 9);
  const p = getParams_();

  const nbBiens = biens.length;
  const nbClients = clients.length;
  const nbVisites = visites.length;
  const caMensuel = commissions
    .filter(r => isCurrentMonth_(r[0]))
    .reduce((acc, r) => acc + toNum_(r[7]), 0);

  const refus = visites.filter(v => String(v[4]).toLowerCase() === 'refus').length;
  const tauxRefus = nbVisites ? refus / nbVisites : 0;
  const objectifAtteint = caMensuel >= p.objectifMensuel;

  sh.getRange('A1:F1').merge();
  sh.getRange('A1').setValue('DASHBOARD CRM IMMOBILIER').setFontWeight('bold').setFontSize(16);

  sh.getRange('A3').setValue('KPI');
  sh.getRange('A4').setValue('Biens actifs');
  sh.getRange('A5').setValue('Clients');
  sh.getRange('A6').setValue('Visites');
  sh.getRange('A7').setValue('CA mensuel net');
  sh.getRange('A8').setValue('Taux refus');
  sh.getRange('A9').setValue('Objectif');

  sh.getRange('B4').setValue(nbBiens);
  sh.getRange('B5').setValue(nbClients);
  sh.getRange('B6').setValue(nbVisites);
  sh.getRange('B7').setValue(caMensuel);
  sh.getRange('B8').setValue(tauxRefus).setNumberFormat('0,00%');
  sh.getRange('B9').setValue(objectifAtteint ? 'Atteint ✅' : 'En cours ⏳');

  sh.getRange('A11').setValue('Objectif mensuel (€)');
  sh.getRange('B11').setValue(p.objectifMensuel);
  sh.getRange('A12').setValue('Réalisation (€)');
  sh.getRange('B12').setValue(caMensuel);
  sh.getRange('A13').setValue('% Objectif');
  sh.getRange('B13').setValue(p.objectifMensuel ? caMensuel / p.objectifMensuel : 0).setNumberFormat('0,00%');

  sh.getRange('D3').setValue('Biens par statut');
  const statMap = mapCount_(biens, 8);
  writeMap_(sh, 'D4', statMap);

  sh.getRange('D11').setValue('Visites par client');
  const visClientMap = mapCount_(visites, 2);
  writeMap_(sh, 'D12', visClientMap);

  sh.getRange('A15').setValue('Graphiques auto');
  buildDashboardCharts_(sh);
  styleDashboard_(sh, objectifAtteint);
}

function appliquerDesignPro() {
  const ss = SpreadsheetApp.getActive();
  const interfaces = [
    CFG.sheets.dashboard,
    CFG.sheets.biens,
    CFG.sheets.clients,
    CFG.sheets.visites,
    CFG.sheets.matching,
    CFG.sheets.commissions,
    CFG.sheets.params
  ];

  interfaces.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastCol = Math.max(sh.getLastColumn(), 16);
    sh.setFrozenRows(1);

    sh.getRange(1, 1, 1, lastCol)
      .setBackground(CFG.ui.headerBg)
      .setFontColor(CFG.ui.headerFont)
      .setFontWeight('bold');

    const lastRow = Math.max(sh.getLastRow(), 2);
    if (lastRow >= 2) {
      const rg = sh.getRange(2, 1, lastRow - 1, lastCol);
      const bgs = [];
      for (let r = 0; r < lastRow - 1; r++) {
        const line = [];
        for (let c = 0; c < lastCol; c++) line.push(r % 2 === 0 ? CFG.ui.zebra1 : CFG.ui.zebra2);
        bgs.push(line);
      }
      rg.setBackgrounds(bgs);
    }

    sh.autoResizeColumns(1, Math.min(lastCol, 12));
  });
}

function protegerDataSheets() {
  const ss = SpreadsheetApp.getActive();
  const dataSheets = [
    CFG.sheets.dataListes,
    CFG.sheets.dataBiens,
    CFG.sheets.dataClients,
    CFG.sheets.dataVisites,
    CFG.sheets.dataMatching,
    CFG.sheets.dataCommissions,
    CFG.sheets.dataStats
  ];

  dataSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (!protections.length) {
      const p = sh.protect().setDescription('Protection auto DATA');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    }
  });
}

// ===== Helpers métier =====

function majScoresBiens_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.biens);
  if (!sh || sh.getLastRow() < 2) return;
  for (let r = 2; r <= sh.getLastRow(); r++) majLigneBien_(r);
}

function majLigneBien_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.biens);
  const p = getParams_();
  const vals = sh.getRange(row, 1, 1, 16).getValues()[0];

  const dpe = String(vals[CFG.col.bien.dpe - 1] || '').toUpperCase();
  const etat = String(vals[CFG.col.bien.etat - 1] || '');
  const exp = vals[CFG.col.bien.expiration - 1];

  const dpeScore = ({ A: 7, B: 6, C: 5, D: 4, E: 3, F: 2, G: 1 }[dpe]) || 0;
  const etatScore = ({ 'Neuf': 4, 'Bon': 3, 'À rafraîchir': 2, 'Travaux': 1 }[etat]) || 0;
  const global = round2_(dpeScore * 0.6 + etatScore * 0.4);

  let alerte = '';
  if (exp instanceof Date) {
    const jours = Math.floor((exp - new Date()) / (1000 * 3600 * 24));
    if (jours < 0) alerte = 'Mandat expiré ❌';
    else if (jours <= p.seuilMandat) alerte = `Mandat expire dans ${jours}j ⚠️`;
  }

  sh.getRange(row, CFG.col.bien.scoreDpe).setValue(dpeScore);
  sh.getRange(row, CFG.col.bien.scoreEtat).setValue(etatScore);
  sh.getRange(row, CFG.col.bien.scoreGlobal).setValue(global);
  sh.getRange(row, CFG.col.bien.alerte).setValue(alerte);

  const alertCell = sh.getRange(row, CFG.col.bien.alerte);
  if (alerte.includes('❌')) alertCell.setBackground(CFG.ui.danger);
  else if (alerte.includes('⚠️')) alertCell.setBackground(CFG.ui.warn);
  else alertCell.setBackground(CFG.ui.ok);
}

function syncCompteursDepuisVisites_() {
  const ss = SpreadsheetApp.getActive();
  const shClients = ss.getSheetByName(CFG.sheets.clients);
  const shVisites = ss.getSheetByName(CFG.sheets.visites);
  if (!shClients || !shVisites || shClients.getLastRow() < 2) return;

  const clients = shClients.getRange(2, 1, shClients.getLastRow() - 1, 14).getValues();
  const visites = shVisites.getLastRow() >= 2 ? shVisites.getRange(2, 1, shVisites.getLastRow() - 1, 6).getValues() : [];

  const map = {};
  clients.forEach(c => { map[c[0]] = { v: 0, r: 0, last: null }; });

  visites.forEach(v => {
    const idClient = v[2];
    if (!map[idClient]) return;
    map[idClient].v += 1;
    if (String(v[4]).toLowerCase() === 'refus') map[idClient].r += 1;
    const d = v[1];
    if (d instanceof Date) {
      if (!map[idClient].last || d > map[idClient].last) map[idClient].last = d;
    }
  });

  for (let i = 0; i < clients.length; i++) {
    const id = clients[i][0];
    clients[i][CFG.col.client.nbVisites - 1] = map[id]?.v || 0;
    clients[i][CFG.col.client.nbRefus - 1] = map[id]?.r || 0;
    if (map[id]?.last) clients[i][CFG.col.client.dernierContact - 1] = map[id].last;
  }

  shClients.getRange(2, 1, clients.length, 14).setValues(clients);
}

function majScoresClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.clients);
  if (!sh || sh.getLastRow() < 2) return;
  for (let r = 2; r <= sh.getLastRow(); r++) majLigneClient_(r);
}

function majLigneClient_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.clients);
  const p = getParams_();
  const vals = sh.getRange(row, 1, 1, 14).getValues()[0];

  const nbVisites = toNum_(vals[CFG.col.client.nbVisites - 1]);
  const nbRefus = toNum_(vals[CFG.col.client.nbRefus - 1]);
  const last = vals[CFG.col.client.dernierContact - 1];

  let joursInact = 999;
  if (last instanceof Date) {
    joursInact = Math.floor((new Date() - last) / (1000 * 3600 * 24));
  }

  const score = Math.max(0, Math.min(100, Math.round(nbVisites * 15 - nbRefus * 10 + (30 - Math.min(30, joursInact)))));
  const alerte = joursInact > p.seuilInactivite ? `Inactif depuis ${joursInact}j ⚠️` : 'OK';

  sh.getRange(row, CFG.col.client.scoreActivite).setValue(score);
  sh.getRange(row, CFG.col.client.alerte).setValue(alerte);
  sh.getRange(row, CFG.col.client.alerte).setBackground(alerte === 'OK' ? CFG.ui.ok : CFG.ui.warn);
}

function setupHeaders_() {
  const ss = SpreadsheetApp.getActive();

  setHeader_(ss.getSheetByName(CFG.sheets.biens), [
    'ID_BIEN','Date_entrée','Adresse','Ville','Code_postal','Type_bien','Surface_m2','Prix_affiché','Statut','Date_expiration_mandat','DPE','Etat_bien','Score_DPE','Score_Etat','Score_Global_Bien','Alerte'
  ]);

  setHeader_(ss.getSheetByName(CFG.sheets.clients), [
    'ID_CLIENT','Date_création','Nom_prenom','Téléphone','Email','Budget_max','Ville_recherche','Type_recherche','Statut_client','Dernier_contact','Nb_visites','Nb_refus','Score_activité','Alerte'
  ]);

  setHeader_(ss.getSheetByName(CFG.sheets.visites), [
    'ID_VISITE','Date_visite','ID_CLIENT','ID_BIEN','Résultat','Commentaire'
  ]);

  setHeader_(ss.getSheetByName(CFG.sheets.matching), [
    'ID_MATCH','ID_CLIENT','ID_BIEN','Score_budget','Score_ville','Score_type','Score_compatibilité','Priorité'
  ]);

  setHeader_(ss.getSheetByName(CFG.sheets.commissions), [
    'Date_vente','ID_BIEN','Prix_vente','Commission_%','Commission_brute','Commission_HT','Commission_TVA','Commission_Nette','Palier_appliqué'
  ]);
}

function setupValidations_() {
  const ss = SpreadsheetApp.getActive();
  const list = ss.getSheetByName(CFG.sheets.dataListes);
  if (!list) return;

  const rgStatutBiens = list.getRange('A2:A20');
  const rgTypeBiens = list.getRange('B2:B20');
  const rgDpe = list.getRange('C2:C20');
  const rgEtat = list.getRange('D2:D20');
  const rgStatutClients = list.getRange('E2:E20');
  const rgResultat = list.getRange('F2:F20');

  const vb = SpreadsheetApp.newDataValidation().requireValueInRange(rgStatutBiens, true).setAllowInvalid(false).build();
  const vt = SpreadsheetApp.newDataValidation().requireValueInRange(rgTypeBiens, true).setAllowInvalid(false).build();
  const vd = SpreadsheetApp.newDataValidation().requireValueInRange(rgDpe, true).setAllowInvalid(false).build();
  const ve = SpreadsheetApp.newDataValidation().requireValueInRange(rgEtat, true).setAllowInvalid(false).build();
  const vc = SpreadsheetApp.newDataValidation().requireValueInRange(rgStatutClients, true).setAllowInvalid(false).build();
  const vr = SpreadsheetApp.newDataValidation().requireValueInRange(rgResultat, true).setAllowInvalid(false).build();

  const shBiens = ss.getSheetByName(CFG.sheets.biens);
  const shClients = ss.getSheetByName(CFG.sheets.clients);
  const shVisites = ss.getSheetByName(CFG.sheets.visites);

  shBiens.getRange('F2:F2000').setDataValidation(vt);
  shBiens.getRange('I2:I2000').setDataValidation(vb);
  shBiens.getRange('K2:K2000').setDataValidation(vd);
  shBiens.getRange('L2:L2000').setDataValidation(ve);

  shClients.getRange('I2:I2000').setDataValidation(vc);

  shVisites.getRange('E2:E5000').setDataValidation(vr);
}

function seedParametres_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.params);
  if (!sh) return;

  const labels = [
    ['Nom agent', ''],
    ['Nom agence', ''],
    ['Objectif mensuel (€)', 15000],
    ['', ''],
    ['Palier 1 max (€)', 60000],
    ['Palier 1 taux', 0.10],
    ['Palier 2 max (€)', 130000],
    ['Palier 2 taux', 0.15],
    ['Palier 3 taux', 0.20],
    ['TVA', 0.20],
    ['', ''],
    ['Seuil inactivité client (jours)', 21],
    ['Seuil alerte mandat (jours)', 30]
  ];

  sh.getRange(1, 1).setValue('Paramètre');
  sh.getRange(1, 2).setValue('Valeur');
  sh.getRange(2, 1, labels.length, 2).setValues(labels);
  sh.getRange('B7:B11').setNumberFormat('0,00%');
}

function seedListes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.dataListes);
  if (!sh) return;

  sh.getRange('A1:F1').setValues([['Statut_bien','Type_bien','DPE','Etat_bien','Statut_client','Résultat_visite']]);
  sh.getRange('A2:A7').setValues([['Prospection'],['Mandat simple'],['Mandat exclusif'],['Compromis'],['Vendu'],['Expiré']]);
  sh.getRange('B2:B6').setValues([['Appartement'],['Maison'],['Terrain'],['Local commercial'],['Immeuble']]);
  sh.getRange('C2:C8').setValues([['A'],['B'],['C'],['D'],['E'],['F'],['G']]);
  sh.getRange('D2:D5').setValues([['Neuf'],['Bon'],['À rafraîchir'],['Travaux']]);
  sh.getRange('E2:E7').setValues([['Nouveau'],['Actif'],['Tiède'],['Froid'],['Signé'],['Perdu']]);
  sh.getRange('F2:F4').setValues([['Intéressé'],['Refus'],['À relancer']]);
}

// ===== Utilitaires =====

function createMissingSheets_(ss) {
  Object.values(CFG.sheets).forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
}

function getParams_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.params);
  const v = sh.getRange('B2:B14').getValues().flat();
  return {
    agent: v[0] || '',
    agence: v[1] || '',
    objectifMensuel: toNum_(v[2]),
    palier1Max: toNum_(v[4]) || 60000,
    palier1Taux: toNum_(v[5]) || 0.10,
    palier2Max: toNum_(v[6]) || 130000,
    palier2Taux: toNum_(v[7]) || 0.15,
    palier3Taux: toNum_(v[8]) || 0.20,
    tva: toNum_(v[9]) || 0.20,
    seuilInactivite: toNum_(v[11]) || 21,
    seuilMandat: toNum_(v[12]) || 30
  };
}

function getRows_(sheetName, cols) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}

function setHeader_(sh, headers) {
  if (!sh) return;
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function mapCount_(arr, idx) {
  const map = {};
  arr.forEach(r => {
    const k = r[idx] || 'N/A';
    map[k] = (map[k] || 0) + 1;
  });
  return map;
}

function writeMap_(sh, startA1, map) {
  const rg = sh.getRange(startA1);
  const r0 = rg.getRow();
  const c0 = rg.getColumn();
  const entries = Object.entries(map);
  if (!entries.length) return;
  sh.getRange(r0, c0, entries.length, 2).setValues(entries);
}

function buildDashboardCharts_(sh) {
  const charts = sh.getCharts();
  charts.forEach(ch => sh.removeChart(ch));

  const chart1 = sh.newChart()
    .asColumnChart()
    .addRange(sh.getRange('D4:E10'))
    .setPosition(16, 1, 0, 0)
    .setOption('title', 'Biens par statut')
    .build();
  sh.insertChart(chart1);

  const chart2 = sh.newChart()
    .asBarChart()
    .addRange(sh.getRange('D12:E30'))
    .setPosition(16, 8, 0, 0)
    .setOption('title', 'Visites par client')
    .build();
  sh.insertChart(chart2);
}

function styleDashboard_(sh, objectifAtteint) {
  sh.getRange('A1:F1').setBackground(CFG.ui.headerBg).setFontColor('#fff').setHorizontalAlignment('center');
  sh.getRange('A3:B13').setBorder(true, true, true, true, true, true);
  sh.getRange('B7:B12').setNumberFormat('#,##0.00 €');
  sh.getRange('B9').setBackground(objectifAtteint ? CFG.ui.ok : CFG.ui.warn);
  sh.getRange('B8').setBackground('#fef9c3');
}

function colorPrioriteMatching_(sh, nRows) {
  const rg = sh.getRange(2, CFG.col.match.priorite, nRows, 1);
  const vals = rg.getValues();
  const bgs = vals.map(v => {
    const p = String(v[0]);
    if (p === 'Haute') return [CFG.ui.ok];
    if (p === 'Moyenne') return [CFG.ui.warn];
    return [CFG.ui.danger];
  });
  rg.setBackgrounds(bgs);
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

function isCurrentMonth_(d) {
  if (!(d instanceof Date)) return false;
  const now = new Date();
  return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
}
```

---

## 3) Formules / validations

### A. Formules (optionnelles, légères, FR `;`)

> Le script calcule déjà la majorité des éléments. Si vous voulez une redondance visuelle dans les onglets interface :

- `INTERFACE_BIENS!M2` (Score DPE):
```excel
=SI(K2="";"";SI(K2="A";7;SI(K2="B";6;SI(K2="C";5;SI(K2="D";4;SI(K2="E";3;SI(K2="F";2;1)))))))
```

- `INTERFACE_BIENS!N2` (Score État):
```excel
=SI(L2="";"";SI(L2="Neuf";4;SI(L2="Bon";3;SI(L2="À rafraîchir";2;SI(L2="Travaux";1;0)))))
```

- `INTERFACE_BIENS!O2` (Score Global):
```excel
=SI(OU(M2="";N2="");"";ARRONDI(M2*0,6+N2*0,4;2))
```

- `INTERFACE_CLIENTS!M2` (Score activité simplifié):
```excel
=MAX(0;MIN(100;K2*15-L2*10+(30-MIN(30;AUJOURDHUI()-J2))))
```

Copiez vers le bas seulement sur une plage utile (ex: 2000 lignes), pas sur des colonnes entières.

---

### B. Validations de données (sans conflit)

Configurer via **Données > Validation des données** (si vous ne lancez pas l’initialisation script) :

- `INTERFACE_BIENS!F2:F2000` → Liste depuis `DATA_LISTES!B2:B20`
- `INTERFACE_BIENS!I2:I2000` → Liste depuis `DATA_LISTES!A2:A20`
- `INTERFACE_BIENS!K2:K2000` → Liste depuis `DATA_LISTES!C2:C20`
- `INTERFACE_BIENS!L2:L2000` → Liste depuis `DATA_LISTES!D2:D20`
- `INTERFACE_CLIENTS!I2:I2000` → Liste depuis `DATA_LISTES!E2:E20`
- `INTERFACE_VISITES!E2:E5000` → Liste depuis `DATA_LISTES!F2:F20`

Options recommandées :
- **Refuser l’entrée** si invalide (zéro erreur de validation).
- Afficher le menu déroulant dans la cellule.

---

## 4) Guide installation pas à pas

1. **Créer un Google Sheet vierge** (nom recommandé: `CRM Immobilier - [VotreNom]`).
2. **Créer les onglets** listés en bloc 1 (ou laissez le script les créer).
3. Ouvrir **Extensions > Apps Script**.
4. Supprimer le contenu par défaut de `Code.gs`.
5. **Copier-coller tout le script** du bloc 2.
6. Enregistrer (`Ctrl+S`) avec le nom de projet: `CRM Immobilier`.
7. Revenir au Sheet puis recharger la page.
8. Ouvrir le menu **📊 CRM Immobilier**.
9. Cliquer **🚀 Initialiser CRM** (permissions Google à accepter la 1ère fois).
10. Aller dans `PARAMETRES` et personnaliser :
    - Nom agent / agence
    - Objectif mensuel
    - Paliers de commission
    - TVA
    - Seuil inactivité / mandat
11. Commencer la saisie dans `INTERFACE_BIENS`, `INTERFACE_CLIENTS`, `INTERFACE_VISITES`.
12. Utiliser les boutons menu pour maintenance :
    - `🔁 Rafraîchir scores & alertes`
    - `🎯 Recalcul matching`
    - `💶 Recalcul commissions`
    - `📊 Mettre à jour dashboard`

### Conseils “revendable / duplicable”
- Dupliquez le fichier via **Fichier > Créer une copie** par agent.
- Gardez `DATA_*` protégés pour éviter la casse.
- Modifiez uniquement `PARAMETRES` et `DATA_LISTES` pour adapter un nouveau client.
- Ajoutez un onglet `README_CLIENT` avec vos conditions d’utilisation / support.
