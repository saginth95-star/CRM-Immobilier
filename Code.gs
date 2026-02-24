// V4 FIX SAFE SYNC SUIVI
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
    dataStats: 'DATA_STATS',
    criteresPoids: 'CRITERES_POIDS',
    paramScoring: 'PARAM_SCORING'
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
  },
  col: {
    biens: {
      id: 1, bien: 2, dateCreation: 3, ville: 4, type: 5, prix: 6, surface: 7, pieces: 8, chambres: 9,
      exterieur: 10, parking: 11, ascenseur: 12, etage: 13, dernierEtage: 14, charges: 15, taxe: 16,
      etatCopro: 17, scoreCopro: 18, dpe: 19, etatBien: 20, dateFinMandat: 21, statut: 22,
      scoreDpe: 23, scoreEtat: 24, scoreGlobal: 25, alerte: 26, note: 27
    },
    clients: {
      id: 1, client: 2, dateCreation: 3, tel: 4, email: 5, budget: 6, ville: 7, type: 8, surfaceMin: 9,
      statut: 10, dernierContact: 11, score: 12, priorite: 13,
      piecesMin: 14, chambresMin: 15, exterieur: 16, parking: 17, ascenseur: 18, etageMax: 19,
      dernierEtage: 20, chargesMax: 21, taxeMax: 22, dpeMin: 23, etatBienMin: 24, etatCoproMin: 25
    },
    suivi: {
      idSuivi: 1, idClient: 2, client: 3, idBien: 4, bien: 5, dateCreation: 6, canal: 7, statut: 8,
      score: 9, derniereAction: 10, auteur: 11, pointsPositifs: 12, pointsNegatifs: 13
    },
    commissions: {
      date: 1, reference: 2, proprietaire: 3, acquereur: 4, prixNetVendeur: 5, commissionBrute: 6,
      tva: 7, commissionHt: 8, entree: 9, sortie: 10, taux: 11, urssaf: 12, impot: 13,
      commissionNette: 14, ratioEntree: 15, ratioSortie: 16, totalCumule: 17
    },
    matching: {
      bien: 1, ville: 2, type: 3, prix: 4, score: 5, niveau: 6, statut: 7, derniereAction: 8, idClient: 9, idBien: 10,
      scoreVille: 11, scoreType: 12, scoreBudget: 13, scoreSurface: 14, scorePieces: 15, scoreChambres: 16, scoreOptions: 17, scoreComplements: 18
    }
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(APP.menu)
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
    .addItem('⚡ Actions rapides — ✅ Marquer Visité', 'actionRapideVisite_')
    .addItem('⚡ Actions rapides — ❤️ Marquer Favori', 'actionRapideFavori_')
    .addItem('⚡ Actions rapides — 🔁 Marquer À relancer', 'actionRapideRelance_')
    .addItem('⚡ Actions rapides — ❌ Marquer Refusé', 'actionRapideRefuse_')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚡ Dates rapides')
      .addItem('📅 Date du jour (cellule active)', 'dateRapideAujourdhui_')
      .addItem('📅+7 Date +7 jours (cellule active)', 'dateRapidePlus7_')
      .addItem('📅+30 Date +30 jours (cellule active)', 'dateRapidePlus30_')
      .addItem('🕒 Date+heure maintenant (cellule active)', 'dateRapideMaintenant_'))
    .addSeparator()
    .addItem('➕ Ajouter un client', 'ajouterClientRapide')
    .addItem('🏠 Ajouter un bien', 'ajouterBienRapide')
    .addToUi();
}


function getUiSafe_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (_) {
    return null;
  }
}

function safeUiAlert_(message, title) {
  const msg = String(message || '');
  const ttl = String(title || 'CRM V4');
  const ui = getUiSafe_();
  if (ui) {
    ui.alert(ttl, msg, ui.ButtonSet.OK);
    return;
  }
  try {
    SpreadsheetApp.getActive().toast(msg, ttl, 5);
  } catch (_) {}
}

function safeUiPrompt_(title, prompt) {
  const ui = getUiSafe_();
  if (!ui) return null;
  return ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
}

function initialiserCRMV4() {
  const ss = SpreadsheetApp.getActive();
  const hasBiens = !!ss.getSheetByName(APP.sh.biens);
  const hasClients = !!ss.getSheetByName(APP.sh.clients);
  const hasMatching = !!ss.getSheetByName(APP.sh.matching);
  const hasSuivi = !!ss.getSheetByName(APP.sh.suivi);
  const hasDashboard = !!ss.getSheetByName(APP.sh.dashboard);
  const freshInstall = !(hasBiens && hasClients && hasMatching && hasSuivi && hasDashboard);

  creerOngletsSiAbsents_(ss);

  if (freshInstall) {
    definirEntetes_();
    remplirListes_();
    remplirParametres_();
    remplirParamScoring_();
  } else {
    assurerListesVillesOfficielles_();
  }

  appliquerCorrectifsV4_();
  refreshAllV4_();
  safeUiAlert_(freshInstall ? 'CRM V4 initialisé ✅' : 'Correctifs CRM V4 appliqués ✅', 'CRM V4');
}

function initialiserCRMV2() { initialiserCRMV4(); }
function refreshAllV2() { refreshAllV4_(); }
function appliquerDesignV2() { appliquerDesignV4_(); }

function onEdit(e) {
  if (!e || !e.range) return;
  const rg = e.range;
  if (rg.getNumRows() > 1 || rg.getNumColumns() > 1) return;

  const sh = rg.getSheet();
  if (!sh) return;
  const row = rg.getRow();
  const col = rg.getColumn();
  const name = sh.getName();
  if (name === APP.sh.suivi && row === 1 && col === 16) {
    filtrerSuiviParClient_();
    trierSuivi_();
    appliquerStylesSuivi_();
    return;
  }
  if (row < 2) return;

  if (name === APP.sh.biens) {
    autoIdBien_(row);
    autoDateCreation_(sh, row, APP.col.biens.bien, APP.col.biens.dateCreation);
    majScoreBienLigne_(row);
    syncInterfaceVersDataBiens_();
    appliquerStylesBiens_();
    return;
  }

  if (name === APP.sh.clients) {
    if (col === APP.col.clients.ville || col === APP.col.clients.type) {
      appliquerMultiSelectClient_(e);
    }
    if (col === APP.col.clients.tel) {
      formaterTelephoneClient_(sh, row);
    }
    autoIdClient_(row);
    autoDateCreation_(sh, row, APP.col.clients.client, APP.col.clients.dateCreation);
    majScoreClientLigne_(row);
    syncInterfaceVersDataClients_();
    appliquerStylesClients_();
    return;
  }

  if (name === APP.sh.matching) {
    if ((row === 2 || row === 3) && col === 2) {
      genererFicheMatching_();
      return;
    }
    if (row >= 6 && col === APP.col.matching.statut) {
      enregistrerStatutDepuisMatching_(row);
      majScoreClientDepuisSuivi_(row);
      majCriteresPoids_();
  majDashboardV2();
      genererFicheMatching_();
      return;
    }
  }

  if (name === APP.sh.suivi) {
    if (col === APP.col.suivi.pointsPositifs || col === APP.col.suivi.pointsNegatifs) {
      appliquerMultiSelectSuivi_(e);
    }
    syncInterfaceSuiviVersData_();
    trierSuivi_();
    majScoreActiviteClients_();
    appliquerStylesSuivi_();
    majCriteresPoids_();
    majDashboardV2();
    return;
  }

  if (name === APP.sh.commissions) {
    recalculCommissionsV2_();
    majDashboardV2();
  }
}

function onSelectionChange(e) {
  return;
}

function appliquerMultiSelectClient_(e) {
  if (!e || !e.range) return;
  const rg = e.range;
  const row = rg.getRow();
  const col = rg.getColumn();
  if (row < 2) return;
  if (col !== APP.col.clients.ville && col !== APP.col.clients.type) return;

  const selected = String(e.value || '').trim();
  if (!selected) return;

  const oldRaw = String(e.oldValue || '').trim();
  if (!oldRaw) {
    rg.setValue(selected);
    return;
  }

  const vals = splitMulti_(oldRaw);
  const exists = vals.includes(selected);
  const next = exists ? vals.filter(v => v !== selected) : vals.concat([selected]);
  rg.setValue(next.join(' | '));
}


function appliquerMultiSelectSuivi_(e) {
  if (!e || !e.range) return;
  const rg = e.range;
  const row = rg.getRow();
  const col = rg.getColumn();
  if (row < 2) return;
  if (col !== APP.col.suivi.pointsPositifs && col !== APP.col.suivi.pointsNegatifs) return;

  const selected = String(e.value || '').trim();
  if (!selected) return;

  const oldRaw = String(e.oldValue || '').trim();
  if (!oldRaw) {
    rg.setValue(selected);
    return;
  }

  const vals = splitMulti_(oldRaw);
  const exists = vals.includes(selected);
  const next = exists ? vals.filter(v => v !== selected) : vals.concat([selected]);
  rg.setValue(next.join(' | '));
}

function formaterTelephoneClient_(sh, row) {
  if (!sh || row < 2) return;
  const rg = sh.getRange(row, APP.col.clients.tel);
  const raw = String(rg.getValue() || '').trim();
  if (!raw) {
    rg.setNumberFormat('@');
    return;
  }

  let v = raw.replace(/[\.\-\s]/g, '');
  v = v.replace(/^\+33/, '0').replace(/^0033/, '0');
  v = v.replace(/\D/g, '');
  if (v.length === 9) v = '0' + v;

  if (v.length === 10) {
    v = v.replace(/(\d{2})(?=\d)/g, '$1 ').trim();
  }

  rg.setNumberFormat('@');
  rg.setValue(v);
}

function assurerListesVillesOfficielles_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.listes);
  if (!sh) return;
  const villes = ['Colombes','Bois-Colombes','La Garenne-Colombes','Asnières-sur-Seine','Gennevilliers','Clichy','Courbevoie','Argenteuil','Paris','Levallois-Perret','Boulogne Billancourt'];
  sh.getRange(2, 1, villes.length, 1).setValues(villes.map(v => [v]));
}

function appliquerCorrectifsV4_() {
  assurerListesVillesOfficielles_();
  remplirParamScoring_();
  setupValidationsV2_();
  appliquerDesignV4_();
  appliquerStylesBiens_();
  appliquerStylesClients_();
  appliquerStylesSuivi_();

  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  if (sh && sh.getLastRow() >= 2) {
    for (let r = 2; r <= sh.getLastRow(); r++) formaterTelephoneClient_(sh, r);
  }

  protegerData_();
}

function refreshAllV4_() {
  majScoresBiens_();
  majScoreActiviteClients_();
  syncInterfaceVersDataBiens_();
  syncInterfaceVersDataClients_();
  genererFicheMatching_();
  syncDataSuiviVersInterface_();
  trierSuivi_();
  majCriteresPoids_();
  majDashboardV2();
  appliquerDesignV4_();
  SpreadsheetApp.getActive().toast('Rafraîchissement complet terminé ✅', 'CRM V4', 3);
}

function ajouterClientRapide() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const rep = safeUiPrompt_('Nouveau client', 'Nom du client :');
  if (!rep) { safeUiAlert_("Action disponible uniquement depuis l'interface Google Sheets.", 'CRM V4'); return; }
  const ui = getUiSafe_();
  if (!ui || rep.getSelectedButton() !== ui.Button.OK) return;
  const nom = String(rep.getResponseText() || '').trim();
  if (!nom) return;

  const id = nextId_('CL', APP.sh.clients, APP.col.clients.id);
  sh.appendRow([id, nom, new Date(), '', '', '', '', '', '', 'Nouveau', '', 0, 'Moyenne', '', '', '', '', '', '', '', '', '', '', '', '']);
  syncInterfaceVersDataClients_();
  appliquerStylesClients_();
}

function ajouterBienRapide() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const rep = safeUiPrompt_('Nouveau bien', 'Adresse du bien :');
  if (!rep) { safeUiAlert_("Action disponible uniquement depuis l'interface Google Sheets.", 'CRM V4'); return; }
  const ui = getUiSafe_();
  if (!ui || rep.getSelectedButton() !== ui.Button.OK) return;
  const adresse = String(rep.getResponseText() || '').trim();
  if (!adresse) return;

  const id = nextId_('BI', APP.sh.biens, APP.col.biens.id);
  sh.appendRow([id, adresse, new Date(), '', '', '', '', '', '', 'Non', 'Non', 'Non', '', 'Non', '', '', '', '', '', '', '', 'Prospection', '', '', '', '', '']);
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
  const rg = sh.getActiveRange();
  if (!rg || rg.getRow() < 6) {
    ss.toast('Sélectionnez une ligne de matching (6+).', 'CRM V4', 4);
    return;
  }
  if (!String(sh.getRange('B2').getValue() || '')) {
    ss.toast('Sélectionnez un client en B2.', 'CRM V4', 4);
    return;
  }
  enregistrerStatutDepuisMatching_(rg.getRow());
  majScoreClientDepuisSuivi_(rg.getRow());
  majDashboardV2();
  genererFicheMatching_();
  ss.toast('Ligne validée ✅', 'CRM V4', 3);
}

function actionRapideVisite_() { appliquerActionRapideStatut_('Visité'); }
function actionRapideFavori_() { appliquerActionRapideStatut_('Favori'); }
function actionRapideRelance_() { appliquerActionRapideStatut_('À relancer'); }
function actionRapideRefuse_() { appliquerActionRapideStatut_('Refusé'); }

function appliquerActionRapideStatut_(statut) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== APP.sh.matching) {
    ss.toast('Placez-vous sur INTERFACE_MATCHING.', 'CRM V4', 4);
    return;
  }
  const rg = sh.getActiveRange();
  if (!rg || rg.getRow() < 6) {
    ss.toast('Sélectionnez une ligne de matching (6+).', 'CRM V4', 4);
    return;
  }
  if (!String(sh.getRange('B2').getValue() || '').trim()) {
    ss.toast('Sélectionnez un client en B2.', 'CRM V4', 4);
    return;
  }
  sh.getRange(rg.getRow(), APP.col.matching.statut).setValue(statut);
  enregistrerStatutDepuisMatching_(rg.getRow());
  majScoreClientDepuisSuivi_(rg.getRow());
  genererFicheMatching_();
  majDashboardV2();
  ss.toast(`Statut mis à jour : ${statut} ✅`, 'CRM V4', 3);
}

function dateRapideAujourdhui_() { ecrireDateRapide_(new Date(), 'dd/mm/yyyy'); }
function dateRapidePlus7_() { const d = new Date(); d.setDate(d.getDate() + 7); ecrireDateRapide_(d, 'dd/mm/yyyy'); }
function dateRapidePlus30_() { const d = new Date(); d.setDate(d.getDate() + 30); ecrireDateRapide_(d, 'dd/mm/yyyy'); }
function dateRapideMaintenant_() { ecrireDateRapide_(new Date(), 'dd/mm/yyyy hh:mm'); }

function ecrireDateRapide_(d, fmt) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const rg = sh ? sh.getActiveRange() : null;
  if (!sh || !rg) { ss.toast('Sélectionnez une cellule.', 'CRM V4', 4); return; }
  const dataSheets = [APP.sh.listes, APP.sh.dataBiens, APP.sh.dataClients, APP.sh.dataSuivi, APP.sh.dataCom, APP.sh.dataStats];
  if (dataSheets.includes(sh.getName())) { ss.toast('Action bloquée sur onglet DATA protégé.', 'CRM V4', 4); return; }
  try {
    rg.setValue(d);
    rg.setNumberFormat(fmt);
  } catch (e) {
    ss.toast('Impossible d\'écrire la date ici.', 'CRM V4', 4);
  }
}

function genererFicheMatching_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.matching);
  if (!sh) return;
  const clients = getRows_(APP.sh.clients, 25);
  const biens = getRows_(APP.sh.biens, 27);
  const suivi = getRows_(APP.sh.dataSuivi, 13);

  const selClient = String(sh.getRange('B2').getValue() || '');
  const selFiltre = String(sh.getRange('B3').getValue() || 'Tout');

  try { sh.getRange('A1:R1').breakApart(); } catch (_) {}
  sh.getRange('A1:R1').merge().setValue('FICHE MATCHING CLIENT').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('Client sélectionné');
  sh.getRange('A3').setValue('Filtre affichage');
  sh.getRange('D2').setValue('Dernière MAJ');
  sh.getRange('E2').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('A4:R4').setValues([['Bien', 'Ville', 'Type', 'Prix', 'Score', 'Niveau', 'Statut client-bien', 'Dernière action', 'ID Client', 'ID Bien', 'Score Ville', 'Score Type', 'Score Budget', 'Score Surface', 'Score Pièces', 'Score Chambres', 'Score Options', 'Score Compléments']]);

  const clientLabels = clients.filter(r => r[APP.col.clients.id-1] && r[APP.col.clients.client-1]).map(r => `${r[0]} — ${r[1]}`);
  if (clientLabels.length) {
    sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(clientLabels, true).setAllowInvalid(false).build());
  } else {
    sh.getRange('B2').clearDataValidations();
  }
  const filtres = ['Tout', 'Non traités', 'Favoris seulement', 'À relancer seulement', 'Masquer Refusés'];
  sh.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(filtres, true).setAllowInvalid(false).build());

  const nextClient = (clientLabels.includes(selClient) ? selClient : (clientLabels[0] || ''));
  const nextFiltre = (filtres.includes(selFiltre) ? selFiltre : 'Tout');
  if (String(sh.getRange('B2').getValue() || '') !== nextClient) sh.getRange('B2').setValue(nextClient);
  if (String(sh.getRange('B3').getValue() || '') !== nextFiltre) sh.getRange('B3').setValue(nextFiltre);

  const clientLabel = String(sh.getRange('B2').getValue() || '');
  const filtre = String(sh.getRange('B3').getValue() || 'Tout');
  const idClient = parseIdFromLabel_(clientLabel);

  sh.getRange('A6:R700').clearContent();
  sh.getRange('A6:R700').clearDataValidations();
  if (!idClient) { designMatchingV4_(); return; }

  const c = clients.find(r => String(r[APP.col.clients.id-1]) === idClient);
  if (!c) return;

  const budget = toNum_(c[APP.col.clients.budget-1]);
  const villeRecherche = splitMulti_(c[APP.col.clients.ville-1]);
  const typeRecherche = splitMulti_(c[APP.col.clients.type-1]);
  const surfaceMin = toNum_(c[APP.col.clients.surfaceMin-1]);
  const piecesMin = toNum_(c[APP.col.clients.piecesMin-1]);
  const chambresMin = toNum_(c[APP.col.clients.chambresMin-1]);
  const exterieur = String(c[APP.col.clients.exterieur-1] || '');
  const parking = String(c[APP.col.clients.parking-1] || '');
  const ascenseur = String(c[APP.col.clients.ascenseur-1] || '');
  const etageMax = toNum_(c[APP.col.clients.etageMax-1]);
  const dernierEtage = String(c[APP.col.clients.dernierEtage-1] || '');
  const chargesMax = toNum_(c[APP.col.clients.chargesMax-1]);
  const taxeMax = toNum_(c[APP.col.clients.taxeMax-1]);
  const dpeMin = String(c[APP.col.clients.dpeMin-1] || '');
  const etatBienMin = String(c[APP.col.clients.etatBienMin-1] || '');
  const etatCoproMin = String(c[APP.col.clients.etatCoproMin-1] || '');

  const rows = [];
  for (let i = 0; i < biens.length; i++) {
    const b = biens[i];
    const statutBien = String(b[APP.col.biens.statut-1] || '');
    if (['Vendu', 'Retiré', 'Expiré'].includes(statutBien)) continue;

    const idBien = String(b[APP.col.biens.id-1] || '');
    const detail = calculScoreMatching_(
      {
        ville: String(b[APP.col.biens.ville-1] || ''),
        type: String(b[APP.col.biens.type-1] || ''),
        prix: toNum_(b[APP.col.biens.prix-1]),
        surface: toNum_(b[APP.col.biens.surface-1]),
        pieces: toNum_(b[APP.col.biens.pieces-1]),
        chambres: toNum_(b[APP.col.biens.chambres-1]),
        exterieur: String(b[APP.col.biens.exterieur-1] || ''),
        parking: String(b[APP.col.biens.parking-1] || ''),
        ascenseur: String(b[APP.col.biens.ascenseur-1] || ''),
        etage: toNum_(b[APP.col.biens.etage-1]),
        dernierEtage: String(b[APP.col.biens.dernierEtage-1] || ''),
        charges: toNum_(b[APP.col.biens.charges-1]),
        taxe: toNum_(b[APP.col.biens.taxe-1]),
        dpe: String(b[APP.col.biens.dpe-1] || ''),
        etatBien: String(b[APP.col.biens.etatBien-1] || ''),
        etatCopro: String(b[APP.col.biens.etatCopro-1] || '')
      },
      { villeRecherche, typeRecherche, budget, surfaceMin, piecesMin, chambresMin, exterieur, parking, ascenseur, etageMax, dernierEtage, chargesMax, taxeMax, dpeMin, etatBienMin, etatCoproMin, idClient }
    );

    const score = detail.total;
    const niveau = score >= 80 ? 'Très bon' : score >= 60 ? 'Moyen' : 'Faible';
    const found = suivi.find(s => String(s[APP.col.suivi.idClient-1]) === idClient && String(s[APP.col.suivi.idBien-1]) === idBien);
    const statut = found ? String(found[APP.col.suivi.statut-1] || 'À proposer') : 'À proposer';
    const dateAct = found ? found[APP.col.suivi.derniereAction-1] : '';

    if (!filtreMatchingOK_(filtre, statut)) continue;
    rows.push([
      `${idBien} — ${String(b[APP.col.biens.bien-1] || '')}`,
      b[APP.col.biens.ville-1],
      b[APP.col.biens.type-1],
      b[APP.col.biens.prix-1],
      score,
      niveau,
      statut,
      dateAct,
      idClient,
      idBien,
      detail.ville,
      detail.type,
      detail.budget,
      detail.surface,
      detail.pieces,
      detail.chambres,
      detail.options,
      detail.complements
    ]);
  }

  rows.sort((a, b) => toNum_(b[4]) - toNum_(a[4]));
  if (!rows.length) { designMatchingV4_(); return; }

  sh.getRange(6, 1, rows.length, 18).setValues(rows);
  sh.getRange(6, APP.col.matching.prix, rows.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(6, APP.col.matching.score, rows.length, 1).setNumberFormat('0');
  sh.getRange(6, APP.col.matching.derniereAction, rows.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  const listes = ss.getSheetByName(APP.sh.listes);
  if (listes) {
    sh.getRange(6, APP.col.matching.statut, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('H2:H200'), true).setAllowInvalid(false).build()
    );
  }

  applyMatchingStatusColors_(sh, rows.length);
  designMatchingV4_();
}


function hasValue_(v) {
  return !(v === '' || v === null || v === undefined);
}

function parseFloor_(v) {
  const t = normTxt_(v);
  if (!t) return null;
  if (t === 'rdc') return 0;
  const n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}

function getParamScoring_() {
  const defaults = [
    { key: 'villeRecherchee', points: 20, mode: 'LIST' },
    { key: 'typeRecherche', points: 15, mode: 'LIST' },
    { key: 'budgetMax', points: 20, mode: 'MAX' },
    { key: 'surfaceMin', points: 10, mode: 'MIN' },
    { key: 'piecesMin', points: 10, mode: 'MIN_ABSOLU' },
    { key: 'chambresMin', points: 10, mode: 'MIN_ABSOLU' },
    { key: 'exterieurSouhaite', points: 3, mode: 'SPECIAL_BOOL_WISH' },
    { key: 'parkingSouhaite', points: 4, mode: 'SPECIAL_BOOL_WISH' },
    { key: 'ascenseurSouhaite', points: 3, mode: 'SPECIAL_BOOL_WISH' },
    { key: 'dernierEtage', points: 6, mode: 'SPECIAL_DERNIER_ETAGE' },
    { key: 'etageMinimum', points: 5, mode: 'MIN_FLOOR' },
    { key: 'chargesMax', points: 1, mode: 'MAX' },
    { key: 'taxeFonciereMax', points: 1, mode: 'MAX' },
    { key: 'dpeMin', points: 2, mode: 'MIN_QUALITY_DPE' },
    { key: 'etatBienMin', points: 1, mode: 'MIN_QUALITY_ETAT_BIEN' },
    { key: 'etatCoproMin', points: 1, mode: 'MIN_QUALITY_COPRO' }
  ];

  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.paramScoring);
  if (!sh || sh.getLastRow() < 2) return defaults;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  const rows = vals
    .map(r => ({ key: String(r[0] || '').trim(), points: toNum_(r[1]), mode: String(r[2] || '').trim().toUpperCase(), active: normTxt_(r[3]) !== 'non' && normTxt_(r[3]) !== 'false' }))
    .filter(r => r.key && r.active);
  if (!rows.length) return defaults;
  return rows.map(r => ({ key: r.key, points: r.points || 0, mode: (r.mode || 'LIST'), absolute: String(r.mode || '').toUpperCase() === 'ABSOLU' }));
}

function computeConfigScore_(bien, crit) {
  const cfg = getParamScoring_();
  const mapping = {
    budgetMax: { b: 'prix', c: 'budget', bucket: 'budget' },
    villeRecherchee: { b: 'ville', c: 'villeRecherche', bucket: 'ville' },
    typeRecherche: { b: 'type', c: 'typeRecherche', bucket: 'type' },
    surfaceMin: { b: 'surface', c: 'surfaceMin', bucket: 'surface' },
    piecesMin: { b: 'pieces', c: 'piecesMin', bucket: 'pieces' },
    chambresMin: { b: 'chambres', c: 'chambresMin', bucket: 'chambres' },
    exterieurSouhaite: { b: 'exterieur', c: 'exterieur', bucket: 'options' },
    parkingSouhaite: { b: 'parking', c: 'parking', bucket: 'options' },
    ascenseurSouhaite: { b: 'ascenseur', c: 'ascenseur', bucket: 'options' },
    etageMinimum: { b: 'etage', c: 'etageMax', bucket: 'complements' },
    etageMax: { b: 'etage', c: 'etageMax', bucket: 'complements' },
    dernierEtage: { b: 'dernierEtage', c: 'dernierEtage', bucket: 'options' },
    chargesMax: { b: 'charges', c: 'chargesMax', bucket: 'complements' },
    taxeFonciereMax: { b: 'taxe', c: 'taxeMax', bucket: 'complements' },
    dpeMin: { b: 'dpe', c: 'dpeMin', bucket: 'complements' },
    etatBienMin: { b: 'etatBien', c: 'etatBienMin', bucket: 'complements' },
    etatCoproMin: { b: 'etatCopro', c: 'etatCoproMin', bucket: 'complements' }
  };
  const out = { ville: 0, type: 0, budget: 0, surface: 0, pieces: 0, chambres: 0, options: 0, complements: 0, blocked: false };

  cfg.forEach(rule => {
    const map = mapping[rule.key];
    if (!map) return;
    const wanted = crit[map.c];
    const bienVal = bien[map.b];
    const mode = String(rule.mode || '').toUpperCase();
    const isAbsolute = !!rule.absolute || mode === 'ABSOLU';
    const points = Math.max(0, toNum_(rule.points));
    let pts = 0;
    let pass = true;

    if (!hasValue_(wanted) || String(wanted).trim() === '') {
      out[map.bucket] += points;
      return;
    }

    if (mode === 'LIST') {
      const arr = Array.isArray(wanted) ? wanted : splitMulti_(wanted);
      pts = scoreListMatch_(bienVal, arr, points);
      pass = pts > 0;
    } else if (mode === 'MAX') {
      pts = scoreMaxCriterion_(bienVal, wanted, points);
      pass = toNum_(bienVal) <= toNum_(wanted);
    } else if (mode === 'MIN' || mode === 'MIN_ABSOLU') {
      pts = scoreMinCriterion_(bienVal, wanted, points);
      pass = toNum_(bienVal) >= toNum_(wanted);
      if (mode === 'MIN_ABSOLU' && !pass) out.blocked = true;
    } else if (mode === 'SPECIAL_BOOL_WISH' || mode === 'SPECIAL_DERNIER_ETAGE') {
      const wish = normTxt_(wanted);
      pts = scoreOuiNon_(bienVal, wanted, points);
      if (!wish || wish === 'indifferent' || wish === 'indifférent') {
        pass = true;
        pts = Math.round(points * 0.2);
      } else if (wish === 'obligatoire') {
        pass = normBoolOuiNon_(bienVal) === 'oui';
      } else if (wish === 'oui') pass = normBoolOuiNon_(bienVal) === 'oui';
      else if (wish === 'non') pass = normBoolOuiNon_(bienVal) === 'non';
      else pass = true;
      if (wish === 'obligatoire' && !pass) out.blocked = true;
    } else if (mode === 'MAX_FLOOR' || mode === 'MIN_FLOOR') {
      const minFloor = parseFloor_(wanted);
      const bFloor = parseFloor_(bienVal);
      if (minFloor === null) pts = points;
      else {
        pass = bFloor !== null && bFloor >= minFloor;
        pts = pass ? points : 0;
      }
    } else if (mode === 'MIN_QUALITY_DPE') {
      pts = scoreThreshold_(bienVal, wanted, ['A', 'B', 'C', 'D', 'E', 'F', 'G'], points);
      pass = pts > 0;
    } else if (mode === 'MIN_QUALITY_ETAT_BIEN' || mode === 'MIN_QUALITY_COPRO') {
      pts = scoreThreshold_(bienVal, wanted, ['Gros travaux', 'Travaux', 'À rafraîchir', 'Bon état', 'Très bon état', 'Neuf'], points);
      pass = pts > 0;
    } else {
      if (Array.isArray(wanted)) pass = wanted.map(normTxt_).includes(normTxt_(bienVal));
      else pass = normTxt_(bienVal) === normTxt_(wanted);
      pts = pass ? points : 0;
    }

    if (isAbsolute && !pass) out.blocked = true;
    out[map.bucket] += toNum_(pts);
  });

  Object.keys(out).forEach(k => {
    if (k !== 'blocked') out[k] = Math.max(0, toNum_(out[k]));
  });

  return out;
}

function calculScoreMatching_(bien, crit) {
  const out = computeConfigScore_(bien, crit);
  if (out.blocked) {
    out.total = 0;
    return out;
  }
  const adjust = scoreFeedbackAdjustement_(bien, crit.idClient);
  out.total = Math.max(0, Math.min(100, Math.round(out.ville + out.type + out.budget + out.surface + out.pieces + out.chambres + out.options + out.complements + adjust)));
  return out;
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
  if (!nRows) return;
  const lv = sh.getRange(6, APP.col.matching.niveau, nRows, 1).getValues();
  const st = sh.getRange(6, APP.col.matching.statut, nRows, 1).getValues();
  sh.getRange(6, APP.col.matching.niveau, nRows, 1).setBackgrounds(lv.map(r => r[0] === 'Très bon' ? [APP.ui.success] : (r[0] === 'Moyen' ? [APP.ui.warning] : [APP.ui.danger])));
  sh.getRange(6, APP.col.matching.statut, nRows, 1).setBackgrounds(st.map(r => {
    const x = String(r[0] || '');
    if (x === 'Favori') return [APP.ui.success];
    if (x === 'À relancer') return [APP.ui.warning];
    if (x === 'Refusé') return [APP.ui.danger];
    if (x === 'Visité') return [APP.ui.info];
    if (x === 'Offre formulée') return ['#e9d5ff'];
    return ['#ffffff'];
  }));
}

function enregistrerStatutDepuisMatching_(row) {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getSheetByName(APP.sh.matching);
  if (!m) return;
  const idClient = String(m.getRange(row, APP.col.matching.idClient).getValue() || parseIdFromLabel_(String(m.getRange('B2').getValue() || '')));
  const idBien = String(m.getRange(row, APP.col.matching.idBien).getValue() || parseIdFromLabel_(String(m.getRange(row, APP.col.matching.bien).getValue() || '')));
  if (!idClient || !idBien) return;

  const clientsSh = ss.getSheetByName(APP.sh.clients);
  const biensSh = ss.getSheetByName(APP.sh.biens);
  const clientRow = findRowByValue_(clientsSh, APP.col.clients.id, idClient);
  const bienRow = findRowByValue_(biensSh, APP.col.biens.id, idBien);
  const clientNom = clientRow ? String(clientsSh.getRange(clientRow, APP.col.clients.client).getValue() || '') : '';
  const bienNom = bienRow ? String(biensSh.getRange(bienRow, APP.col.biens.bien).getValue() || '') : '';

  const score = toNum_(m.getRange(row, APP.col.matching.score).getValue());
  const statut = String(m.getRange(row, APP.col.matching.statut).getValue() || 'À proposer');
  const now = new Date();

  const d = ss.getSheetByName(APP.sh.dataSuivi);
  if (!d) return;
  const vals = getRows_(APP.sh.dataSuivi, 13);
  let idx = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][APP.col.suivi.idClient-1]) === idClient && String(vals[i][APP.col.suivi.idBien-1]) === idBien) {
      idx = i + 2;
      break;
    }
  }

  if (idx === -1) {
    d.appendRow([`SU${Utilities.formatString('%06d', nextNumberForPrefix_('SU', APP.sh.dataSuivi, APP.col.suivi.idSuivi))}`, idClient, `${idClient} — ${clientNom}`, idBien, `${idBien} — ${bienNom}`, now, '', statut, score, now, Session.getActiveUser().getEmail() || '', '', '']);
  } else {
    d.getRange(idx, APP.col.suivi.statut, 1, 3).setValues([[statut, score, now]]);
  }

  syncDataSuiviVersInterface_();
}

function majScoreClientDepuisSuivi_(matchingRow) {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getSheetByName(APP.sh.matching);
  const idClient = String(m.getRange(matchingRow, APP.col.matching.idClient).getValue() || parseIdFromLabel_(String(m.getRange('B2').getValue() || '')));
  if (!idClient) return;

  const clientsSh = ss.getSheetByName(APP.sh.clients);
  const row = findRowByValue_(clientsSh, APP.col.clients.id, idClient);
  if (!row) return;

  clientsSh.getRange(row, APP.col.clients.dernierContact).setValue(new Date());
  majScoreClientLigne_(row);
  syncInterfaceVersDataClients_();
}

function majScoresBiens_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const last = sh.getLastRow();
  if (last < 2) return;
  for (let r = 2; r <= last; r++) {
    autoIdBien_(r);
    majScoreBienLigne_(r);
  }
  appliquerStylesBiens_();
}

function majScoreActiviteClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const last = sh.getLastRow();
  if (last < 2) return;
  for (let r = 2; r <= last; r++) {
    autoIdClient_(r);
    majScoreClientLigne_(r);
  }
  appliquerStylesClients_();
}

function majScoreClientLigne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const suivi = getRows_(APP.sh.dataSuivi, 13);

  const idClient = String(sh.getRange(row, APP.col.clients.id).getValue() || '');
  const nom = String(sh.getRange(row, APP.col.clients.client).getValue() || '');
  if (!idClient && !nom) return;

  const dernierContact = sh.getRange(row, APP.col.clients.dernierContact).getValue();
  const actions = suivi.filter(r => String(r[APP.col.suivi.idClient-1]) === idClient);

  let visites = 0, refus = 0, favoris = 0;
  actions.forEach(a => {
    const st = String(a[APP.col.suivi.statut-1] || '');
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

  sh.getRange(row, APP.col.clients.score, 1, 2).setValues([[score, priorite]]);
}

function majDashboardV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.dashboard);
  if (!sh) return;

  const biens = getRows_(APP.sh.biens, 27);
  const clients = getRows_(APP.sh.clients, 25);
  const suivi = getRows_(APP.sh.dataSuivi, 13);
  const coms = getRows_(APP.sh.commissions, 16);
  const p = getParams_();

  const clientsActifs = clients.filter(c => ['Actif', 'Tiède'].includes(String(c[APP.col.clients.statut-1]))).length;
  const biensActifs = biens.filter(b => !['Vendu', 'Retiré', 'Expiré'].includes(String(b[APP.col.biens.statut-1]))).length;

  const now = new Date();
  const mois = now.getMonth();
  const annee = now.getFullYear();
  const actionsMois = suivi.filter(r => r[APP.col.suivi.derniereAction-1] instanceof Date && r[APP.col.suivi.derniereAction-1].getMonth() === mois && r[APP.col.suivi.derniereAction-1].getFullYear() === annee);

  const visitesMois = actionsMois.filter(a => String(a[APP.col.suivi.statut-1]) === 'Visité').length;
  const refusMois = actionsMois.filter(a => String(a[APP.col.suivi.statut-1]) === 'Refusé').length;
  const favoris = actionsMois.filter(a => String(a[APP.col.suivi.statut-1]) === 'Favori').length;
  const actionsRelance = suivi.filter(a => String(a[APP.col.suivi.statut-1]) === 'À relancer').length;
  const offresMois = actionsMois.filter(a => String(a[APP.col.suivi.statut-1]) === 'Offre formulée').length;
  const tauxRefus = visitesMois ? refusMois / visitesMois : 0;
  const tauxVisiteOffre = visitesMois ? offresMois / visitesMois : 0;

  const caMensuel = coms
    .filter(c => c[0] instanceof Date && c[0].getMonth() === mois && c[0].getFullYear() === annee)
    .reduce((acc, c) => acc + toNum_(c[APP.col.commissions.commissionNette-1]), 0);

  const pctObjectif = p.objectifMensuel ? caMensuel / p.objectifMensuel : 0;
  const objectifAtteint = caMensuel >= p.objectifMensuel;

  sh.getRange('A1:J80').clearContent();
  sh.getRange('A1:J1').merge().setValue('DASHBOARD TERRAIN — CRM IMMOBILIER V4');

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
  drawKpiCard_(sh, 'A15:C18', 'Offres formulées (mois)', offresMois, '#e9d5ff');
  drawKpiCard_(sh, 'D15:F18', 'Taux visite → offre', Utilities.formatString('%.1f%%', tauxVisiteOffre * 100), tauxVisiteOffre >= 0.3 ? APP.ui.success : APP.ui.warning);

  sh.getRange('I3:J3').setValues([['Top clients actifs', 'Nb']]);
  const mapClients = mapCountByStatus_(suivi, APP.col.suivi.client-1, ['Visité', 'Favori', 'À relancer']);
  writeMapSorted_(sh, 4, 9, mapClients, 6);

  sh.getRange('I12:J12').setValues([['Top biens visités', 'Nb']]);
  const mapBiens = mapCountByStatus_(suivi, APP.col.suivi.bien-1, ['Visité']);
  writeMapSorted_(sh, 13, 9, mapBiens, 6);

  buildChartsDashboardV2_(sh);
  designDashboardV4_();
}

function drawKpiCard_(sh, a1, title, value, bg) {
  const r = sh.getRange(a1);
  r.merge().setBackground(bg).setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r.getRow(), r.getColumn()).setValue(`${title}\n${value}`).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold').setFontSize(11);
}

function mapCountByStatus_(rows, keyIdx, statuses) {
  const out = {};
  rows.forEach(r => {
    if (!statuses.includes(String(r[APP.col.suivi.statut-1]))) return;
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

  sh.insertChart(sh.newChart().asColumnChart().addRange(sh.getRange('I3:J9')).setOption('title', 'Top clients actifs').setPosition(20, 1, 0, 0).build());
  sh.insertChart(sh.newChart().asBarChart().addRange(sh.getRange('I12:J18')).setOption('title', 'Top biens visités').setPosition(20, 6, 0, 0).build());
}

function recalculCommissionsV2_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.commissions);
  if (!sh) return;
  const p = getParams_();
  const last = sh.getLastRow();
  if (last < 2) return;

  const vals = sh.getRange(2, 1, last - 1, 16).getValues();
  for (let i = 0; i < vals.length; i++) {
    const brute = toNum_(vals[i][APP.col.commissions.commissionBrute - 1]);
    if (!brute) {
      vals[i][APP.col.commissions.tva - 1] = '';
      vals[i][APP.col.commissions.commissionHt - 1] = '';
      vals[i][APP.col.commissions.entree - 1] = '';
      vals[i][APP.col.commissions.sortie - 1] = '';
      vals[i][APP.col.commissions.urssaf - 1] = '';
      vals[i][APP.col.commissions.impot - 1] = '';
      vals[i][APP.col.commissions.commissionNette - 1] = '';
      continue;
    }

    const taux = safePercent_(vals[i][APP.col.commissions.taux - 1], 0);
    const ratioEntree = safePercent_(vals[i][APP.col.commissions.ratioEntree - 1], 1);
    const ratioSortie = safePercent_(vals[i][APP.col.commissions.ratioSortie - 1], 1);

    const ht = brute / (1 + p.tva);
    const tva = brute - ht;
    const entree = (ht / 2) * ratioEntree;
    const sortie = (ht / 2) * ratioSortie;
    const base = taux > 0 ? (entree + sortie) * taux : 0;
    const urssaf = base * p.urssaf;
    const impot = base * p.impot;
    const nette = base - urssaf - impot;

    vals[i][APP.col.commissions.ratioEntree - 1] = ratioEntree;
    vals[i][APP.col.commissions.ratioSortie - 1] = ratioSortie;
    vals[i][APP.col.commissions.tva - 1] = round2_(tva);
    vals[i][APP.col.commissions.commissionHt - 1] = round2_(ht);
    vals[i][APP.col.commissions.entree - 1] = round2_(entree);
    vals[i][APP.col.commissions.sortie - 1] = round2_(sortie);
    vals[i][APP.col.commissions.urssaf - 1] = round2_(urssaf);
    vals[i][APP.col.commissions.impot - 1] = round2_(impot);
    vals[i][APP.col.commissions.commissionNette - 1] = round2_(nette);
  }
  if (vals.length) sh.getRange(2, 1, vals.length, 16).setValues(vals);
  if (vals.length && sh.getLastColumn() >= APP.col.commissions.totalCumule) {
    let cum = 0;
    const cumul = vals.map(r => {
      const net = toNum_(r[APP.col.commissions.commissionNette - 1]);
      cum += net;
      return [round2_(cum)];
    });
    sh.getRange(2, APP.col.commissions.totalCumule, cumul.length, 1).setValues(cumul).setNumberFormat('#,##0.00 €');
  }
  sh.getRange(2, 1, vals.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, APP.col.commissions.taux, vals.length, 1).setNumberFormat('0%');
  sh.getRange(2, APP.col.commissions.ratioEntree, vals.length, 2).setNumberFormat('0%');
  sh.getRange(2, APP.col.commissions.prixNetVendeur, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionBrute, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.tva, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionHt, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.entree, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.sortie, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.urssaf, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.impot, vals.length, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionNette, vals.length, 1).setNumberFormat('#,##0.00 €');
}

function creerOngletsSiAbsents_(ss) {
  Object.values(APP.sh).forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
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
  sh.getRange(headerRow, 1, Math.min(lastRow - headerRow + 1, 400), lastCol)
    .setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);

  if (widthMap) Object.keys(widthMap).forEach(c => sh.setColumnWidth(Number(c), widthMap[c]));
}

function designDashboardV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.dashboard);
  if (!sh) return;
  sh.setColumnWidths(1, 18, 120);
  sh.getRange('A1:J1').setBackground('#0f172a').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeights(1, 16, 30);
  sh.getRange('I3:J3').setBackground(APP.ui.header).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('I12:J12').setBackground(APP.ui.header).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('I4:J18').setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);
}

function designCommissionsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.commissions);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:110,2:150,3:170,4:170,5:130,6:130,7:120,8:130,9:120,10:120,11:90,12:110,13:100,14:130,15:110,16:110,17:130});
  sh.setFrozenRows(1);
  sh.getRange('A2:A5000').setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, APP.col.commissions.prixNetVendeur, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionBrute, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.tva, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionHt, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.entree, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.sortie, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.urssaf, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.impot, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.commissionNette, 5000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.commissions.taux, 5000, 1).setNumberFormat('0%');
  sh.getRange(2, APP.col.commissions.ratioEntree, 5000, 2).setNumberFormat('0%');
  if (sh.getLastColumn() >= APP.col.commissions.totalCumule) sh.getRange(2, APP.col.commissions.totalCumule, 5000, 1).setNumberFormat('#,##0.00 €');
  ensureFilter_(sh, 1);
}

function designParamsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:260,2:180});
}


function filtrerSuiviParClient_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.suivi);
  if (!sh) return;

  const data = getRows_(APP.sh.dataSuivi, 13);
  const clients = getRows_(APP.sh.clients, 25);

  sh.getRange('O1').setValue('Client sélectionné').setFontWeight('bold');
  sh.getRange('P2').setValue('Vue filtrée (A:M) — source complète conservée dans DATA_SUIVI_CLIENT_BIEN').setFontColor(APP.ui.textMuted || '#4b5563');

  const labels = ['Tous'].concat(clients.filter(r => r[APP.col.clients.id-1] && r[APP.col.clients.client-1]).map(r => `${r[APP.col.clients.id-1]} — ${r[APP.col.clients.client-1]}`));
  sh.getRange('P1').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(labels, true).setAllowInvalid(false).build());

  let sel = String(sh.getRange('P1').getValue() || 'Tous');
  if (!sel || labels.indexOf(sel) === -1) {
    sel = 'Tous';
    sh.getRange('P1').setValue(sel);
  }

  const id = parseIdFromLabel_(sel);
  const filtered = (sel === 'Tous' || !id) ? data : data.filter(r => String(r[APP.col.suivi.idClient-1]) === id);
  sh.getRange('A2:M12000').clearContent();
  if (filtered.length) sh.getRange(2, 1, filtered.length, 13).setValues(filtered);
}

function trierSuivi_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh || sh.getLastRow() < 3) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, 13).sort([{ column: APP.col.suivi.derniereAction, ascending: false }]);
  ensureFilter_(sh, 1);
}

function appliquerStylesSuivi_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, APP.col.suivi.statut, n, 2).getValues();

  const bgs = vals.map(v => {
    const st = String(v[0] || '');
    if (st === 'Favori') return [APP.ui.success];
    if (st === 'À relancer') return [APP.ui.warning];
    if (st === 'Refusé') return [APP.ui.danger];
    if (st === 'Visité') return [APP.ui.info];
    if (st === 'Offre formulée') return ['#e9d5ff'];
    return [APP.ui.neutral];
  });
  sh.getRange(2, APP.col.suivi.statut, n, 1).setBackgrounds(bgs);
  sh.getRange(2, APP.col.suivi.dateCreation, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, APP.col.suivi.derniereAction, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, APP.col.suivi.score, n, 1).setNumberFormat('0');
}

function appliquerStylesBiens_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;

  applyZebra_(sh, 2, sh.getLastRow(), 1, 27);
  sh.getRange(2, 1, n, 27).setFontColor('#111827');

  const vals = sh.getRange(2, 1, n, 27).getValues();
  const alertBg = [];
  const scoreBg = [];

  for (let i = 0; i < vals.length; i++) {
    const statut = String(vals[i][APP.col.biens.statut-1] || '');
    const alert = String(vals[i][APP.col.biens.alerte-1] || '').toLowerCase();
    const score = toNum_(vals[i][APP.col.biens.scoreGlobal-1]);

    if (alert.includes('expiré') || alert.includes('❌')) alertBg.push([APP.ui.danger]);
    else if (alert.includes('expire') || alert.includes('⚠️')) alertBg.push([APP.ui.warning]);
    else if (alert === 'ok') alertBg.push([APP.ui.success]);
    else alertBg.push([i % 2 === 0 ? APP.ui.zebra1 : APP.ui.zebra2]);

    scoreBg.push([score >= 7.5 ? '#ecfccb' : (i % 2 === 0 ? APP.ui.zebra1 : APP.ui.zebra2)]);

    if (statut === 'Vendu' || statut === 'Retiré') {
      sh.getRange(i + 2, 1, 1, 27).setBackground(APP.ui.neutral).setFontColor(APP.ui.textMuted);
    }
  }

  sh.getRange(2, APP.col.biens.alerte, n, 1).setBackgrounds(alertBg);
  sh.getRange(2, APP.col.biens.scoreGlobal, n, 1).setBackgrounds(scoreBg);
  sh.getRange(2, APP.col.biens.prix, n, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.biens.surface, n, 1).setNumberFormat('0.00');
  sh.getRange(2, APP.col.biens.pieces, n, 1).setNumberFormat('0');
  sh.getRange(2, APP.col.biens.chambres, n, 1).setNumberFormat('0');
  sh.getRange(2, APP.col.biens.etage, n, 1).setNumberFormat('0');
  sh.getRange(2, APP.col.biens.charges, n, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.biens.taxe, n, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.biens.dateCreation, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, APP.col.biens.dateFinMandat, n, 1).setNumberFormat('dd/mm/yyyy');
}

function appliquerStylesClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, APP.col.clients.score, n, 1).getValues();
  const bg = vals.map(v => {
    const score = toNum_(v[0]);
    if (score >= 70) return [APP.ui.success];
    if (score >= 40) return [APP.ui.warning];
    return [APP.ui.danger];
  });
  sh.getRange(2, APP.col.clients.score, n, 1).setBackgrounds(bg);
  sh.getRange(2, APP.col.clients.dateCreation, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, APP.col.clients.dernierContact, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, APP.col.clients.tel, n, 1).setNumberFormat('@');
  sh.getRange(2, APP.col.clients.budget, n, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.clients.surfaceMin, n, 1).setNumberFormat('0.00');
}

function ensureFilter_(sh, headerRow) {
  if (!sh) return;
  const h = headerRow || 1;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), h);
  if (lastRow < h || lastCol < 1) return;

  try {
    const existing = sh.getFilter();
    if (existing) existing.remove();
  } catch (_) {}

  try {
    sh.getRange(h, 1, lastRow - h + 1, lastCol).createFilter();
  } catch (_) {}
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
    try {
      const sh = ss.getSheetByName(n);
      if (!sh) return;
      const existing = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (existing && existing.length) return;
      const p = sh.protect().setDescription('Protection DATA auto V4');
      try { p.removeEditors(p.getEditors()); } catch (_) {}
      try { if (p.canDomainEdit()) p.setDomainEdit(false); } catch (_) {}
    } catch (_) {}
  });
}

function getRows_(sheetName, cols) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}

function setHeaders_(sh, headers) {
  if (!sh) return;
  const maxCols = Math.max(sh.getLastColumn(), headers.length);
  sh.getRange(1, 1, 1, maxCols).clearContent();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getParams_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  if (!sh) {
    return {
      objectifMensuel: 15000,
      p1max: 60000,
      p1taux: 0.10,
      p2max: 130000,
      p2taux: 0.15,
      p3taux: 0.20,
      tva: 0.20,
      inact: 21,
      alMandat: 30,
      urssaf: 0.214,
      impot: 0.022
    };
  }
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
    alMandat: toNum_(m['Alerte mandat (jours)']) || 30,
    urssaf: toNum_(m['URSSAF (%)']) || 0.214,
    impot: toNum_(m['Impôt (%)']) || 0.022
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


function splitMulti_(v) {
  return String(v || '').split('|').map(x => x.trim()).filter(Boolean);
}

function scoreListMatch_(bienVal, selectedList, points) {
  if (!selectedList || !selectedList.length) return points;
  return selectedList.map(normTxt_).includes(normTxt_(bienVal)) ? points : 0;
}

function scoreBudget_(prix, budget, points) {
  const p = toNum_(prix), b = toNum_(budget);
  if (b <= 0) return points;
  if (p <= b) return points;
  const dep = (p - b) / b;
  if (dep <= 0.05) return Math.round(points * 0.85);
  if (dep <= 0.10) return Math.round(points * 0.60);
  if (dep <= 0.20) return Math.round(points * 0.25);
  return 0;
}

function scoreMinCriterion_(bienVal, minVal, points) {
  const min = toNum_(minVal), val = toNum_(bienVal);
  if (min <= 0) return points;
  if (val >= min) return points;
  const ratio = val > 0 ? val / min : 0;
  return Math.max(0, Math.round(points * ratio));
}

function scoreMaxCriterion_(bienVal, maxVal, points) {
  const max = toNum_(maxVal), val = toNum_(bienVal);
  if (max <= 0) return points;
  if (val <= max) return points;
  const dep = (val - max) / max;
  if (dep <= 0.10) return Math.round(points * 0.5);
  if (dep <= 0.20) return Math.round(points * 0.2);
  return 0;
}

function scoreOuiNon_(bienVal, wantedVal, points) {
  const wish = normTxt_(wantedVal);
  const bien = normBoolOuiNon_(bienVal);
  if (!wish || wish === 'indifferent' || wish === 'indifférent') return points;
  if (wish === 'obligatoire') return bien === 'oui' ? points : 0;
  if (wish === 'oui') return bien === 'oui' ? points : 0;
  if (wish === 'non') return bien === 'non' ? points : 0;
  return points;
}

function scoreThreshold_(bienVal, minVal, orderedList, points) {
  const isEtat = orderedList && orderedList.length && normTxt_(orderedList[0]).includes('travaux');
  const min = isEtat ? normEtat_(minVal) : String(minVal || '').trim();
  if (!min) return points;
  const list = isEtat ? orderedList.map(normEtat_) : orderedList;
  const val = isEtat ? normEtat_(bienVal) : String(bienVal || '').trim();
  const idxMin = list.indexOf(min);
  const idxVal = list.indexOf(val);
  if (idxMin === -1 || idxVal === -1) return 0;
  return isEtat ? (idxVal >= idxMin ? points : 0) : (idxVal <= idxMin ? points : 0);
}


function normEtat_(v) {
  const t = normTxt_(v)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return '';
  if (t.includes('gros travaux')) return 'gros travaux';
  if (t.includes('travaux') || t.includes('a renover') || t.includes('a renove')) return 'travaux';
  if (t.includes('rafraich') || t.includes('a rafraich')) return 'a rafraichir';
  if (t.includes('tres bon')) return 'tres bon etat';
  if (t.includes('bon etat')) return 'bon etat';
  if (t.includes('neuf')) return 'neuf';
  return t;
}

function normBoolOuiNon_(v) {
  const t = normTxt_(v);
  if (['oui','yes','true','1'].includes(t)) return 'oui';
  if (['non','no','false','0'].includes(t)) return 'non';
  return '';
}

function safePercent_(v, defVal) {
  const raw = String(v == null ? '' : v).trim();
  if (!raw) return defVal;

  if (raw.includes('%')) {
    const n = Number(raw.replace(/\s/g, '').replace('%', '').replace(',', '.'));
    if (isNaN(n)) return defVal;
    return n / 100;
  }

  const n = Number(raw.replace(/\s/g, '').replace(',', '.'));
  if (isNaN(n)) return defVal;
  if (n === 0) return 0;
  return n > 1 ? n / 100 : n;
}

function normTxt_(v) { return String(v || '').trim().toLowerCase(); }

function parseIdFromLabel_(label) {
  const s = String(label || '');
  const i = s.indexOf('—');
  return (i > -1 ? s.slice(0, i) : s).trim();
}

function autoDateCreation_(sheet, row, colReference, colDateCreation) {
  const d = sheet.getRange(row, colDateCreation).getValue();
  if (d) return;
  const lastCol = Math.max(sheet.getLastColumn(), colReference);
  const rowVals = sheet.getRange(row, colReference, 1, lastCol - colReference + 1).getValues()[0];
  const hasData = rowVals.some((v, idx) => (colReference + idx) !== colDateCreation && String(v || '').trim() !== '');
  if (hasData) sheet.getRange(row, colDateCreation).setValue(new Date());
}

function nextId_(prefix, sheetName, col) {
  const n = nextNumberForPrefix_(prefix, sheetName, col);
  return `${prefix}${Utilities.formatString('%04d', n)}`;
}

function nextNumberForPrefix_(prefix, sheetName, col) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 1;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().flat();
  let max = 0;
  vals.forEach(v => {
    const s = String(v || '');
    if (s.startsWith(prefix)) {
      const n = Number(s.slice(prefix.length));
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max + 1;
}

function autoIdClient_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  const id = String(sh.getRange(row, APP.col.clients.id).getValue() || '').trim();
  const nom = String(sh.getRange(row, APP.col.clients.client).getValue() || '').trim();
  if (!nom || id) return;
  sh.getRange(row, APP.col.clients.id).setValue(nextId_('CL', APP.sh.clients, APP.col.clients.id));
}

function autoIdBien_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  const id = String(sh.getRange(row, APP.col.biens.id).getValue() || '').trim();
  const nom = String(sh.getRange(row, APP.col.biens.bien).getValue() || '').trim();
  if (!nom || id) return;
  sh.getRange(row, APP.col.biens.id).setValue(nextId_('BI', APP.sh.biens, APP.col.biens.id));
}

function findRowByValue_(sh, col, value) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().flat();
  const idx = vals.findIndex(v => String(v) === String(value));
  return idx === -1 ? 0 : idx + 2;
}

function definirEntetes_() {
  const ss = SpreadsheetApp.getActive();
  setHeaders_(ss.getSheetByName(APP.sh.biens), [
    'ID Bien','Bien','Date création','Ville','Type','Prix affiché','Surface m²','Pièces','Chambres',
    'Extérieur','Parking','Ascenseur','Étage','Dernier étage','Charges mensuelles','Taxe foncière',
    'État copropriété','Copropriété_Score','DPE','État bien','Date fin mandat','Statut bien',
    'Score DPE','Score État','Score global','Alerte mandat','Note interne'
  ]);
  setHeaders_(ss.getSheetByName(APP.sh.clients), [
    'ID Client','Client','Date création','Téléphone','Email','Budget max','Ville recherchée','Type recherché','Surface min',
    'Statut client','Dernier contact','Score activité','Priorité relance','Pièces min','Chambres min','Extérieur souhaité','Parking souhaité','Ascenseur souhaité','Étage minimum','Dernier étage souhaité','Charges max','Taxe foncière max','DPE min','État bien min','État copropriété min'
  ]);
  setHeaders_(ss.getSheetByName(APP.sh.matching), ['Bien','Ville','Type','Prix','Score','Niveau','Statut client-bien','Dernière action','ID Client','ID Bien','Score Ville','Score Type','Score Budget','Score Surface','Score Pièces','Score Chambres','Score Options','Score Compléments']);
  setHeaders_(ss.getSheetByName(APP.sh.suivi), ['ID Suivi','ID Client','Client','ID Bien','Bien','Date création','Canal','Statut','Score','Dernière action','Auteur','Points positifs','Points négatifs']);
  setHeaders_(ss.getSheetByName(APP.sh.commissions), ['Date','Référence','Propriétaire','Acquéreur','Prix net vendeur','Commission brute','TVA','Commission HT','Entrée','Sortie','Taux','URSSAF','Impôt','Commission nette','Ratio Entrée','Ratio Sortie','Total cumulé']);
  setHeaders_(ss.getSheetByName(APP.sh.params), ['Paramètre','Valeur']);
  setHeaders_(ss.getSheetByName(APP.sh.listes), ['Villes','Types bien','Statuts bien','DPE','État bien','Statuts client','Résultats visite','Statut client-bien','Canaux','Priorités relance','État copropriété','Oui/Non','Ratios commissions','Points positifs','Points négatifs','Avis global','Décision','Souhait client']);
  setHeaders_(ss.getSheetByName(APP.sh.dataBiens), [
    'ID Bien','Bien','Date création','Ville','Type','Prix affiché','Surface m²','Pièces','Chambres',
    'Extérieur','Parking','Ascenseur','Étage','Dernier étage','Charges mensuelles','Taxe foncière',
    'État copropriété','Copropriété_Score','DPE','État bien','Date fin mandat','Statut bien',
    'Score DPE','Score État','Score global','Alerte mandat','Note interne'
  ]);
  setHeaders_(ss.getSheetByName(APP.sh.dataClients), ['ID Client','Client','Date création','Téléphone','Email','Budget max','Ville recherchée','Type recherché','Surface min','Statut client','Dernier contact','Score activité','Priorité relance','Pièces min','Chambres min','Extérieur souhaité','Parking souhaité','Ascenseur souhaité','Étage minimum','Dernier étage souhaité','Charges max','Taxe foncière max','DPE min','État bien min','État copropriété min']);
  setHeaders_(ss.getSheetByName(APP.sh.dataSuivi), ['ID Suivi','ID Client','Client','ID Bien','Bien','Date création','Canal','Statut','Score','Dernière action','Auteur','Points positifs','Points négatifs']);
  setHeaders_(ss.getSheetByName(APP.sh.dataCom), ['Date','Référence','Propriétaire','Acquéreur','Prix net vendeur','Commission brute','TVA','Commission HT','Entrée','Sortie','Taux','URSSAF','Impôt','Commission nette','Ratio Entrée','Ratio Sortie','Total cumulé']);
  setHeaders_(ss.getSheetByName(APP.sh.dataStats), ['Date','KPI','Valeur']);
  setHeaders_(ss.getSheetByName(APP.sh.criteresPoids), ['Critère','Poids de base','Nb positifs','Nb négatifs','Total feedback','Taux positif','Taux négatif','Coefficient lissage',"Score d'impact",'Poids ajusté final']);
  setHeaders_(ss.getSheetByName(APP.sh.paramScoring), ['Critère','Poids','Mode','Actif','Commentaire']);
}

function remplirListes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.listes);
  if (!sh) return;
  sh.getRange('A2:R500').clearContent();
  const villes = ['Colombes','Bois-Colombes','La Garenne-Colombes','Asnières-sur-Seine','Gennevilliers','Clichy','Courbevoie','Argenteuil','Paris','Levallois-Perret','Boulogne Billancourt'];
  const types = ['Studio','T1','T2','T3','T4','T5+','Maison','Loft','Duplex'];
  const statutsBiens = ['Prospection','Disponible','Offre acceptée','Compromis','Vendu','Retiré','Expiré'];
  const dpe = ['A','B','C','D','E','F','G'];
  const etats = ['Gros travaux','Travaux','À rafraîchir','Bon état','Très bon état','Neuf'];
  const statutsClients = ['Nouveau','Actif','Tiède','Inactif','Clos'];
  const resultats = ['Visité','Favori','À relancer','Refusé','À proposer','Offre formulée'];
  const statutsCB = ['À proposer','Visité','Favori','À relancer','Refusé','Offre formulée'];
  const canaux = ['Téléphone','WhatsApp','Email','Portail','Agence','Recommandation'];
  const priorites = ['Haute','Moyenne','Basse'];
  const ouiNon = ['Oui','Non'];
  const souhaitClient = ['Obligatoire','Indifférent'];
  const ratios = [0.25,0.5,0.75,1];
  const pointsPos = ['Prix','Emplacement','Quartier calme','Transports','Commerces','Luminosité','Exposition','Agencement','Surface','Extérieur','État général','Pas de travaux','Charme','Parking','Cave','Charges','DPE','Potentiel'];
  const pointsNeg = ['Prix','Travaux','Rafraîchissement','Agencement','Surface','Luminosité','Vis-à-vis','Bruit','Quartier','Étage','Sans ascenseur','RDC','Pas d’extérieur','Pas de parking','Charges','DPE','Copropriété','Humidité','Pas de coup de cœur'];
  const avis = ['Très positif','Positif','Mitigé','Négatif','Refus'];
  const decisions = ['Abandon','À réfléchir','2e visite','Offre possible','Offre faite'];

  const write = (col, arr) => sh.getRange(2, col, arr.length, 1).setValues(arr.map(x => [x]));
  write(1, villes); write(2, types); write(3, statutsBiens); write(4, dpe); write(5, etats); write(6, statutsClients);
  write(7, resultats); write(8, statutsCB); write(9, canaux); write(10, priorites); write(11, etats); write(12, ouiNon);
  write(13, ratios); write(14, pointsPos); write(15, pointsNeg); write(16, avis); write(17, decisions); write(18, souhaitClient);
  sh.getRange(2, 13, ratios.length, 1).setNumberFormat('0%');
}

function remplirParamScoring_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.paramScoring);
  if (!sh) return;
  if (sh.getLastRow() < 2) {
    sh.getRange(1, 1, 1, 5).setValues([['Critère','Poids','Mode','Actif','Commentaire']]);
  }
  const existing = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().map(x=>String(x||'').trim()).filter(Boolean) : [];
  if (existing.length) return;
  const rows = [
    ['budgetMax', 16, 'MAX', 'Oui', 'Budget client'],
    ['villeRecherchee', 14, 'LIST', 'Oui', 'Ville recherchée'],
    ['typeRecherche', 12, 'LIST', 'Oui', 'Type recherché'],
    ['surfaceMin', 12, 'MIN', 'Oui', 'Surface minimum'],
    ['piecesMin', 10, 'MIN_ABSOLU', 'Oui', 'Pièces minimales obligatoires'],
    ['chambresMin', 10, 'MIN_ABSOLU', 'Oui', 'Chambres minimales obligatoires'],
    ['exterieurSouhaite', 3, 'SPECIAL_BOOL_WISH', 'Oui', 'Obligatoire => bloquant, Indifférent => faible poids'],
    ['parkingSouhaite', 3, 'SPECIAL_BOOL_WISH', 'Oui', 'Obligatoire => bloquant, Indifférent => faible poids'],
    ['ascenseurSouhaite', 3, 'SPECIAL_BOOL_WISH', 'Oui', 'Obligatoire => bloquant, Indifférent => faible poids'],
    ['etageMinimum', 5, 'MIN_FLOOR', 'Oui', 'RDC=0, étages numériques'],
    ['dernierEtage', 6, 'SPECIAL_DERNIER_ETAGE', 'Oui', 'Obligatoire => bloquant'],
    ['chargesMax', 2, 'MAX', 'Oui', 'Très faible'],
    ['taxeFonciereMax', 2, 'MAX', 'Oui', 'Très faible'],
    ['dpeMin', 4, 'MIN_QUALITY_DPE', 'Oui', 'Moyen'],
    ['etatBienMin', 4, 'MIN_QUALITY_ETAT_BIEN', 'Oui', 'Moyen'],
    ['etatCoproMin', 4, 'MIN_QUALITY_COPRO', 'Oui', 'Moyen']
  ];
  sh.getRange(2,1,rows.length,5).setValues(rows);
  sh.getRange(2,2,rows.length,1).setNumberFormat('0');
  sh.getRange(2,4,rows.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Oui','Non'], true).setAllowInvalid(false).build());
}

function remplirParametres_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.params);
  if (!sh) return;
  const rows = [
    ['Objectif mensuel (€)', 15000],
    ['Palier 1 max (€)', 60000],
    ['Palier 1 taux', 0.10],
    ['Palier 2 max (€)', 130000],
    ['Palier 2 taux', 0.15],
    ['Palier 3 taux', 0.20],
    ['TVA', 0.20],
    ['URSSAF (%)', 0.214],
    ['Impôt (%)', 0.022],
    ['Seuil inactivité (jours)', 21],
    ['Alerte mandat (jours)', 30]
  ];
  sh.getRange('A2:B50').clearContent();
  sh.getRange(2,1,rows.length,2).setValues(rows);
}

function setupValidationsV2_() {
  const ss = SpreadsheetApp.getActive();
  const listes = ss.getSheetByName(APP.sh.listes);
  if (!listes) return;
  const biens = ss.getSheetByName(APP.sh.biens);
  const clients = ss.getSheetByName(APP.sh.clients);
  const suivi = ss.getSheetByName(APP.sh.suivi);
  const matching = ss.getSheetByName(APP.sh.matching);

  const dv = (rg, lst) => rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(lst, true).setAllowInvalid(false).build());

  if (biens) {
    dv(biens.getRange(2, APP.col.biens.ville, 12000, 1), listes.getRange('A2:A500'));
    dv(biens.getRange(2, APP.col.biens.type, 12000, 1), listes.getRange('B2:B500'));
    dv(biens.getRange(2, APP.col.biens.statut, 12000, 1), listes.getRange('C2:C500'));
    dv(biens.getRange(2, APP.col.biens.dpe, 12000, 1), listes.getRange('D2:D500'));
    dv(biens.getRange(2, APP.col.biens.etatBien, 12000, 1), listes.getRange('E2:E500'));
    dv(biens.getRange(2, APP.col.biens.etatCopro, 12000, 1), listes.getRange('K2:K500'));
    dv(biens.getRange(2, APP.col.biens.exterieur, 12000, 1), listes.getRange('L2:L500'));
    dv(biens.getRange(2, APP.col.biens.parking, 12000, 1), listes.getRange('L2:L500'));
    dv(biens.getRange(2, APP.col.biens.ascenseur, 12000, 1), listes.getRange('L2:L500'));
    dv(biens.getRange(2, APP.col.biens.dernierEtage, 12000, 1), listes.getRange('L2:L500'));
    biens.getRange(2, APP.col.biens.dateCreation, 12000, 1).setNumberFormat('dd/mm/yyyy');
    biens.getRange(2, APP.col.biens.dateFinMandat, 12000, 1).setNumberFormat('dd/mm/yyyy');
  }

  if (clients) {
    clients.getRange(2, APP.col.clients.ville, 12000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('A2:A500'), true).setAllowInvalid(true).build()
    );
    clients.getRange(2, APP.col.clients.type, 12000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('B2:B500'), true).setAllowInvalid(true).build()
    );
    dv(clients.getRange(2, APP.col.clients.statut, 12000, 1), listes.getRange('F2:F500'));
    dv(clients.getRange(2, APP.col.clients.priorite, 12000, 1), listes.getRange('J2:J500'));
    dv(clients.getRange(2, APP.col.clients.exterieur, 12000, 1), listes.getRange('R2:R500'));
    dv(clients.getRange(2, APP.col.clients.parking, 12000, 1), listes.getRange('R2:R500'));
    dv(clients.getRange(2, APP.col.clients.ascenseur, 12000, 1), listes.getRange('R2:R500'));
    dv(clients.getRange(2, APP.col.clients.dernierEtage, 12000, 1), listes.getRange('R2:R500'));
    clients.getRange(2, APP.col.clients.etageMax, 12000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['RDC','1','2','3','4','5','6'], true).setAllowInvalid(true).build());
    dv(clients.getRange(2, APP.col.clients.dpeMin, 12000, 1), listes.getRange('D2:D500'));
    dv(clients.getRange(2, APP.col.clients.etatBienMin, 12000, 1), listes.getRange('E2:E500'));
    dv(clients.getRange(2, APP.col.clients.etatCoproMin, 12000, 1), listes.getRange('K2:K500'));
    clients.getRange(2, APP.col.clients.dateCreation, 12000, 1).setNumberFormat('dd/mm/yyyy');
    clients.getRange(2, APP.col.clients.dernierContact, 12000, 1).setNumberFormat('dd/mm/yyyy');
  }

  if (suivi) {
    dv(suivi.getRange(2, APP.col.suivi.canal, 12000, 1), listes.getRange('I2:I500'));
    dv(suivi.getRange(2, APP.col.suivi.statut, 12000, 1), listes.getRange('H2:H500'));
    suivi.getRange(2, APP.col.suivi.pointsPositifs, 12000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('N2:N500'), true).setAllowInvalid(true).build());
    suivi.getRange(2, APP.col.suivi.pointsNegatifs, 12000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('O2:O500'), true).setAllowInvalid(true).build());
    suivi.getRange(2, APP.col.suivi.dateCreation, 12000, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    suivi.getRange(2, APP.col.suivi.derniereAction, 12000, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  }

  if (matching) {
    dv(matching.getRange(6, APP.col.matching.statut, 12000, 1), listes.getRange('H2:H500'));
  }

  const commissions = ss.getSheetByName(APP.sh.commissions);
  if (commissions) {
    commissions.getRange(2, APP.col.commissions.taux, 12000, 1).setNumberFormat('0%');
    commissions.getRange(2, APP.col.commissions.ratioEntree, 12000, 2).setNumberFormat('0%');
    commissions.getRange(2, APP.col.commissions.ratioEntree, 12000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('M2:M10'), true).setAllowInvalid(false).build());
    commissions.getRange(2, APP.col.commissions.ratioSortie, 12000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange('M2:M10'), true).setAllowInvalid(false).build());
  }

  ensureDateValidations_();
}

function ensureDateValidations_() {
  const ss = SpreadsheetApp.getActive();
  const dvr = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  const biens = ss.getSheetByName(APP.sh.biens);
  const clients = ss.getSheetByName(APP.sh.clients);
  const suivi = ss.getSheetByName(APP.sh.suivi);
  const com = ss.getSheetByName(APP.sh.commissions);
  if (biens) { biens.getRange(2, APP.col.biens.dateCreation, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy'); biens.getRange(2, APP.col.biens.dateFinMandat, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy'); }
  if (clients) { clients.getRange(2, APP.col.clients.dateCreation, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy'); clients.getRange(2, APP.col.clients.dernierContact, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy'); }
  if (suivi) { suivi.getRange(2, APP.col.suivi.dateCreation, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy hh:mm'); suivi.getRange(2, APP.col.suivi.derniereAction, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy hh:mm'); }
  if (com) { com.getRange(2, 1, 12000, 1).setDataValidation(dvr).setNumberFormat('dd/mm/yyyy'); }
}

function majScoreBienLigne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  if (!sh) return;
  const p = getParams_();

  const dpe = String(sh.getRange(row, APP.col.biens.dpe).getValue() || '').toUpperCase();
  const etat = String(sh.getRange(row, APP.col.biens.etatBien).getValue() || '');
  const etatCopro = String(sh.getRange(row, APP.col.biens.etatCopro).getValue() || '');
  const dateFinMandat = sh.getRange(row, APP.col.biens.dateFinMandat).getValue();

  const mapDPE = { A: 7, B: 6, C: 5, D: 4, E: 3, F: 2, G: 1 };
  const mapEtat = { 'gros travaux': 1, 'travaux': 2, 'a rafraichir': 3, 'bon etat': 4, 'tres bon etat': 5, 'neuf': 6 };

  const scoreDPE = mapDPE[dpe] || 0;
  const scoreEtat = mapEtat[normEtat_(etat)] || 0;
  const scoreCopro = mapEtat[normEtat_(etatCopro)] || 0;

  const ext = normTxt_(sh.getRange(row, APP.col.biens.exterieur).getValue()) === 'oui' ? 1 : 0;
  const park = normTxt_(sh.getRange(row, APP.col.biens.parking).getValue()) === 'oui' ? 1 : 0;
  const asc = normTxt_(sh.getRange(row, APP.col.biens.ascenseur).getValue()) === 'oui' ? 1 : 0;
  const dernier = normTxt_(sh.getRange(row, APP.col.biens.dernierEtage).getValue()) === 'oui' ? 1 : 0;
  const etage = toNum_(sh.getRange(row, APP.col.biens.etage).getValue());
  const charges = toNum_(sh.getRange(row, APP.col.biens.charges).getValue());
  const taxe = toNum_(sh.getRange(row, APP.col.biens.taxe).getValue());

  const dpeNorm = scoreDPE / 7;
  const etatNorm = scoreEtat / 6;
  const coproNorm = scoreCopro / 6;

  let bonus = 0;
  bonus += ext ? 0.20 : 0;
  bonus += park ? 0.20 : 0;
  bonus += asc ? 0.15 : 0;
  bonus += dernier ? 0.10 : 0;
  bonus += (etage > 0 && etage <= 3) ? 0.10 : 0;
  bonus += (charges > 0 && charges <= 120) ? 0.15 : (charges > 0 && charges <= 250 ? 0.08 : 0);
  bonus += (taxe > 0 && taxe <= 900) ? 0.10 : (taxe > 0 && taxe <= 1700 ? 0.05 : 0);
  if (etage >= 4 && !asc) bonus -= 0.10;
  bonus = Math.max(0, Math.min(1, bonus));

  const globalNorm = Math.max(0, Math.min(1,
    dpeNorm * 0.30 +
    etatNorm * 0.35 +
    coproNorm * 0.20 +
    bonus * 0.15
  ));
  const scoreGlobal = round2_(globalNorm * 10);

  let alerte = '';
  if (dateFinMandat instanceof Date) {
    const jours = Math.floor((dateFinMandat - new Date()) / (1000 * 3600 * 24));
    if (jours < 0) alerte = 'Mandat expiré ❌';
    else if (jours <= p.alMandat) alerte = `Expire dans ${jours}j ⚠️`;
    else alerte = 'OK';
  }

  sh.getRange(row, APP.col.biens.scoreCopro).setValue(scoreCopro);
  sh.getRange(row, APP.col.biens.scoreDpe, 1, 4).setValues([[scoreDPE, scoreEtat, scoreGlobal, alerte]]);
}

function syncInterfaceVersDataBiens_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.biens, 27);
  const dst = ss.getSheetByName(APP.sh.dataBiens);
  dst.getRange('A2:AA12000').clearContent();
  if (!src.length) return;
  const out = src.filter(r => String(r[APP.col.biens.id-1]).trim() !== '');
  if (out.length) dst.getRange(2, 1, out.length, out[0].length).setValues(out);
}

function syncInterfaceVersDataClients_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.clients, 25);
  const dst = ss.getSheetByName(APP.sh.dataClients);
  dst.getRange('A2:Y12000').clearContent();
  if (!src.length) return;
  const out = src.filter(r => String(r[APP.col.clients.id-1]).trim() !== '');
  if (out.length) dst.getRange(2, 1, out.length, out[0].length).setValues(out);
}

function syncDataSuiviVersInterface_() {
  const ss = SpreadsheetApp.getActive();
  const src = getRows_(APP.sh.dataSuivi, 13);
  const dst = ss.getSheetByName(APP.sh.suivi);
  dst.getRange('A2:M12000').clearContent();
  if (src.length) dst.getRange(2, 1, src.length, 13).setValues(src);
  filtrerSuiviParClient_();
  trierSuivi_();
  appliquerStylesSuivi_();
}

function mergeSuiviRowsById_(dataRows, interfaceRows) {
  const byId = {};
  (dataRows || []).forEach(r => {
    const id = String(r[APP.col.suivi.idSuivi - 1] || '').trim();
    if (!id) return;
    byId[id] = r.slice(0, 13);
  });

  (interfaceRows || []).forEach(r => {
    const id = String(r[APP.col.suivi.idSuivi - 1] || '').trim();
    if (!id) return;
    byId[id] = r.slice(0, 13);
  });

  const out = Object.keys(byId).map(k => byId[k]);
  out.sort((a, b) => {
    const da = a[APP.col.suivi.derniereAction - 1] instanceof Date ? a[APP.col.suivi.derniereAction - 1].getTime() : 0;
    const db = b[APP.col.suivi.derniereAction - 1] instanceof Date ? b[APP.col.suivi.derniereAction - 1].getTime() : 0;
    return db - da;
  });
  return out;
}

function syncInterfaceSuiviVersData_() {
  const ss = SpreadsheetApp.getActive();
  const srcSh = ss.getSheetByName(APP.sh.suivi);
  const dst = ss.getSheetByName(APP.sh.dataSuivi);
  if (!srcSh || !dst) return;

  const interfaceRows = getRows_(APP.sh.suivi, 13)
    .filter(r => String(r[APP.col.suivi.idSuivi - 1] || '').trim() !== '');
  const dataRows = getRows_(APP.sh.dataSuivi, 13)
    .filter(r => String(r[APP.col.suivi.idSuivi - 1] || '').trim() !== '');

  const merged = mergeSuiviRowsById_(dataRows, interfaceRows);
  dst.getRange('A2:M12000').clearContent();
  if (merged.length) dst.getRange(2, 1, merged.length, 13).setValues(merged);
}


function designBiensV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.biens);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:110,2:260,3:120,4:140,5:120,6:130,7:100,8:80,9:90,10:95,11:95,12:95,13:80,14:100,15:120,16:120,17:140,18:120,19:80,20:120,21:130,22:120,23:95,24:95,25:100,26:130,27:220});
  sh.setFrozenRows(1);
  sh.getRange('A2:A12000').setHorizontalAlignment('center');
  sh.getRange(2, APP.col.biens.prix, 12000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.biens.surface, 12000, 1).setNumberFormat('0.00');
  sh.getRange(2, APP.col.biens.charges, 12000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.biens.taxe, 12000, 1).setNumberFormat('#,##0.00 €');
  ensureFilter_(sh, 1);
}

function designClientsV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.clients);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:110,2:240,3:120,4:130,5:240,6:130,7:160,8:160,9:110,10:120,11:130,12:100,13:120,14:95,15:110,16:130,17:130,18:130,19:95,20:150,21:110,22:130,23:90,24:110,25:140});
  sh.setFrozenRows(1);
  sh.getRange(2, APP.col.clients.budget, 12000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.clients.surfaceMin, 12000, 1).setNumberFormat('0.00');
  sh.getRange(2, APP.col.clients.piecesMin, 12000, 1).setNumberFormat('0');
  sh.getRange(2, APP.col.clients.chambresMin, 12000, 1).setNumberFormat('0');
  sh.getRange(2, APP.col.clients.etageMax, 12000, 1).setNumberFormat('@');
  sh.getRange(2, APP.col.clients.chargesMax, 12000, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(2, APP.col.clients.taxeMax, 12000, 1).setNumberFormat('#,##0.00 €');
  ensureFilter_(sh, 1);
}

function designMatchingV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.matching);
  if (!sh) return;
  sh.setColumnWidths(1, 18, 120);
  sh.setColumnWidth(1, 290);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(7, 150);
  sh.setColumnWidth(8, 140);
  sh.setFrozenRows(5);
  sh.getRange('A4:R4').setBackground(APP.ui.header).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:A3').setFontWeight('bold');
  sh.getRange('A6:R700').setBorder(true, true, true, true, true, true, APP.ui.border, SpreadsheetApp.BorderStyle.SOLID);
}

function designSuiviV4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.suivi);
  if (!sh) return;
  applyBaseSheetStyle_(sh, 1, 2, {1:120,2:110,3:220,4:110,5:260,6:130,7:120,8:130,9:90,10:130,11:220,12:220,13:220});
  sh.setFrozenRows(1);
  ensureFilter_(sh, 1);
}


function majCriteresPoids_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.sh.criteresPoids);
  if (!sh) return;
  const suivis = getRows_(APP.sh.dataSuivi, 13);
  const criteres = ['Prix','Emplacement','Quartier calme','Transports','Commerces','Luminosité','Exposition','Agencement','Surface','Extérieur','État général','Pas de travaux','Charme','Parking','Cave','Charges','DPE','Potentiel','Travaux','Rafraîchissement','Vis-à-vis','Bruit','Quartier','Étage','Sans ascenseur','RDC','Pas d’extérieur','Pas de parking','Copropriété','Humidité','Pas de coup de cœur'];
  const posCount = {}, negCount = {};
  suivis.forEach(r => {
    splitMulti_(r[APP.col.suivi.pointsPositifs-1]).forEach(x => posCount[x]=(posCount[x]||0)+1);
    splitMulti_(r[APP.col.suivi.pointsNegatifs-1]).forEach(x => negCount[x]=(negCount[x]||0)+1);
  });
  const rows = criteres.map(c => {
    const p = posCount[c] || 0;
    const n = negCount[c] || 0;
    const t = p + n;
    const tp = t ? p / t : 0;
    const tn = t ? n / t : 0;
    const coef = Math.min(1, t / 5);
    const impact = coef * ((tp * 0.6) - (tn * 0.6));
    const final = Math.max(0.5, Math.min(2, 1 + impact));
    return [c,1,p,n,t,tp,tn,coef,impact,final];
  });
  sh.getRange('A2:J200').clearContent();
  if (rows.length) sh.getRange(2,1,rows.length,10).setValues(rows);
  sh.getRange('F2:H200').setNumberFormat('0.00%');
  sh.getRange('I2:J200').setNumberFormat('0.00');
}

function getFeedbackWeights_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.criteresPoids);
  if (!sh || sh.getLastRow() < 2) return {};
  const vals = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
  const out = {};
  vals.forEach(r => out[String(r[0] || '').trim()] = toNum_(r[9]) || 1);
  return out;
}

function scoreFeedbackAdjustement_(bien, idClient) {
  const global = getFeedbackWeights_();
  if (!Object.keys(global).length) return 0;

  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sh.criteresPoids);
  const totalFeedback = sh && sh.getLastRow() >= 2 ? toNum_(sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues().flat().reduce((a, b) => a + toNum_(b), 0)) : 0;
  if (totalFeedback < 3) return 0;

  const local = getClientFeedbackWeights_(idClient);
  const useLocal = local && local.__count >= 3;
  const w = useLocal ? local : global;

  let adj = 0;
  const parkW = (w['Parking'] || 1);
  const extW = (w['Extérieur'] || 1);
  const dpeW = (w['DPE'] || 1);
  const travW = (w['Travaux'] || 1);
  const coproW = (w['Copropriété'] || 1);
  const sansAscW = (w['Sans ascenseur'] || 1);

  if (normBoolOuiNon_(bien.parking) === 'oui') adj += (parkW - 1) * 2;
  if (normBoolOuiNon_(bien.exterieur) === 'oui') adj += (extW - 1) * 2;
  if (String(bien.dpe || '').toUpperCase() && String(bien.dpe || '').toUpperCase() <= 'D') adj += (dpeW - 1) * 2;

  if (normEtat_(bien.etatBien).includes('travaux')) adj -= (travW - 1) * 2;
  if (normEtat_(bien.etatCopro).includes('travaux')) adj -= (coproW - 1) * 2;
  if (toNum_(bien.etage) >= 3 && normBoolOuiNon_(bien.ascenseur) !== 'oui') adj -= (sansAscW - 1) * 2;

  if (useLocal) adj *= 1.15;
  return Math.max(-6, Math.min(6, adj));
}

function getClientFeedbackWeights_(idClient) {
  if (!idClient) return null;
  const suivis = getRows_(APP.sh.dataSuivi, 13).filter(r => String(r[APP.col.suivi.idClient-1]) === String(idClient));
  const pos = {}, neg = {};
  let count = 0;
  suivis.forEach(r => {
    const p = splitMulti_(r[APP.col.suivi.pointsPositifs-1]);
    const n = splitMulti_(r[APP.col.suivi.pointsNegatifs-1]);
    if (p.length || n.length) count++;
    p.forEach(x => pos[x] = (pos[x] || 0) + 1);
    n.forEach(x => neg[x] = (neg[x] || 0) + 1);
  });
  const crits = Array.from(new Set(Object.keys(pos).concat(Object.keys(neg))));
  const out = { __count: count };
  crits.forEach(c => {
    const p = pos[c] || 0;
    const n = neg[c] || 0;
    const t = p + n;
    if (!t) return;
    const coef = Math.min(1, t / 5);
    const impact = coef * (((p / t) * 0.6) - ((n / t) * 0.6));
    out[c] = Math.max(0.5, Math.min(2, 1 + impact));
  });
  return out;
}

/*
CHECKLIST FONCTIONS PRINCIPALES (présentes dans ce fichier)
- onOpen
- initialiserCRMV4
- onEdit
- refreshAllV4_
- creerOngletsSiAbsents_
- definirEntetes_
- remplirListes_
- remplirParametres_
- setupValidationsV2_
- appliquerDesignV4_
- designBiensV4_
- designClientsV4_
- designMatchingV4_
- designSuiviV4_
- designDashboardV4_
- designCommissionsV4_
- designParamsV4_
- majScoreBienLigne_
- majScoreClientLigne_
- genererFicheMatching_
- enregistrerStatutDepuisMatching_
- majDashboardV2
- recalculCommissionsV2_
- syncInterfaceVersDataBiens_
- syncInterfaceVersDataClients_
- syncDataSuiviVersInterface_
- syncInterfaceSuiviVersData_
- protegerData_

AUTO-CONTROLE
- creerOngletsSiAbsents_ defined ? OUI
- majScoreBienLigne_ defined ? OUI
- all sync functions defined ? OUI
- all design* functions defined ? OUI
- no undefined function calls ? OUI (hors helpers locaux inline)
*/


/*
CHANGELOG V4.x
- Migration douce clients A:Y avec nouveaux critères de matching (N:Y) et sync DATA_CLIENTS étendue à 25 colonnes.
- Multi-sélection fiable (toggle) sur villes/types clients avec séparateur " | " et prise en charge dans le matching.
- Matching enrichi : score pondéré détaillé (A:R) avec colonnes de détail par critère et total sur 100.
- onSelectionChange désactivé; validations date ajoutées pour BIENS/CLIENTS/SUIVI/COMMISSIONS + formats conservés.
- Formatage téléphone client renforcé en texte (normalisation 06 XX XX XX XX).
*/
