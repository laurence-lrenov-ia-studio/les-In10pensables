/**
 * IN10-HUB — Back-end Apps Script pour la réception des candidatures
 * ===================================================================
 * Reçoit une candidature depuis le formulaire HTML (Interface 01)
 * - Crée un dossier dédié dans Drive avec les 3 documents (s'ils sont fournis)
 * - Ajoute une ligne récapitulative dans le Google Sheet "Candidatures"
 * - Envoie un accusé de réception au candidat depuis Gmail
 * - Envoie une notification à Audrey et Julie
 */

// ====================================================================
// CONFIGURATION — à compléter par Laurence avant déploiement
// ====================================================================
const CONFIG = {
  // ID du dossier Drive parent où seront créés les sous-dossiers candidats
  // → Va dans Drive, ouvre le dossier "Candidatures In10pensables", copie l'ID de l'URL
  // → https://drive.google.com/drive/folders/[CET_ID]
  DRIVE_FOLDER_ID: '19PsXDuvpV-jcV5Y_TPoh9FFCj0JP9xEo',

  // ID du Google Sheet de suivi des candidatures
  // → Va dans le Sheet, copie l'ID de l'URL
  // → https://docs.google.com/spreadsheets/d/[CET_ID]/edit
  SHEET_ID: '1OAfCY52yshh1O_1nt7yOPn1phjGxlcFZFrMxDbJhPp4',

  // Nom de l'onglet dans le Sheet (par défaut "Candidatures")
  SHEET_NAME: 'Candidatures',

  // Email d'Audrey et Julie (pour la notification + signature de l'accusé)
  // L'accusé de réception sera envoyé DEPUIS l'adresse Gmail propriétaire du script
  // C'est donc cette adresse qui doit être ou bien la boîte commune, ou bien une boîte
  // qui a accès à envoyer "au nom de" contact@in10pensables.fr
  FOUNDERS_EMAILS: ['contact@in10pensables.fr'],

  // Adresse d'envoi affichée dans l'accusé de réception
  REPLY_TO: 'contact@in10pensables.fr',

  // Nom affiché dans l'expéditeur
  SENDER_NAME: 'Audrey & Julie — Les In10pensables'
};

// ====================================================================
// POINT D'ENTRÉE — appelé par le POST du formulaire
// ====================================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // 1. Création du dossier candidat dans Drive
    const folderInfo = createCandidateFolder(payload);

    // 2. Upload des documents fournis (si présents)
    const uploadedDocs = uploadDocuments(payload.documents, folderInfo.folder);

    // 3. Sauvegarde du récap JSON complet dans le dossier
    saveRecapJson(payload, folderInfo.folder);

    // 3bis. Génération du PDF récap complet (pour Audrey & Julie) dans le dossier
    const recapPdf = generateRecapPdf(payload, false);
    folderInfo.folder.createFile(recapPdf);

    // 4. Ligne dans le Sheet de suivi
    appendToSheet(payload, folderInfo, uploadedDocs);

    // 5. Accusé de réception au candidat (avec PDF soft en pièce jointe)
    sendAcknowledgment(payload);

    // 6. Notification aux fondatrices
    notifyFounders(payload, folderInfo, uploadedDocs);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', folder: folderInfo.url }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Erreur doPost: ' + err.message + ' — Stack: ' + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ====================================================================
// 1. CRÉATION DU DOSSIER CANDIDAT
// ====================================================================
function createCandidateFolder(payload) {
  const parent = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const date = new Date();
  const dateStr = Utilities.formatDate(date, 'Europe/Paris', 'yyyy-MM-dd');
  const nomComplet = `${payload.profil.nom.toUpperCase()}_${capitalize(payload.profil.prenom)}`
    .replace(/[^a-zA-Z0-9_-]/g, '');
  const folderName = `${dateStr}_${nomComplet}`;

  const folder = parent.createFolder(folderName);
  return {
    folder: folder,
    name: folderName,
    url: folder.getUrl()
  };
}

// ====================================================================
// 2. UPLOAD DES DOCUMENTS (kbis / urssaf / rcpro)
// ====================================================================
function uploadDocuments(documents, folder) {
  const uploaded = { kbis: null, urssaf: null, rcpro: null };

  if (!documents) return uploaded;

  const docLabels = {
    kbis: 'Kbis',
    urssaf: 'Urssaf',
    rcpro: 'RCpro'
  };

  Object.keys(documents).forEach(key => {
    const doc = documents[key];
    if (doc && doc.data) {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(doc.data),
          doc.type,
          `${docLabels[key]}_${doc.name}`
        );
        const file = folder.createFile(blob);
        uploaded[key] = {
          name: file.getName(),
          url: file.getUrl()
        };
      } catch (err) {
        Logger.log(`Erreur upload ${key}: ${err.message}`);
      }
    }
  });

  return uploaded;
}

// ====================================================================
// 3. SAUVEGARDE DU RÉCAP COMPLET (JSON) DANS LE DOSSIER
// ====================================================================
function saveRecapJson(payload, folder) {
  const recap = {
    timestamp: payload.timestamp,
    profil: payload.profil,
    complements: payload.complements,
    missions: payload.missions
  };
  const json = JSON.stringify(recap, null, 2);
  folder.createFile('recap-candidature.json', json, MimeType.PLAIN_TEXT);
}

// ====================================================================
// 4. AJOUT D'UNE LIGNE DANS LE SHEET DE SUIVI
// ====================================================================
function appendToSheet(payload, folderInfo, uploadedDocs) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // En-têtes
    sheet.appendRow([
      'Date', 'Statut',
      'Prénom', 'Nom', 'Email', 'Téléphone',
      'Société', 'Statut juridique', 'SIRET',
      'Zone d\'intervention', 'Années d\'expérience',
      'Présentation',
      'Missions positionnées', 'Catégories couvertes',
      'Site / plaquette', 'Tarif', 'LinkedIn', 'Instagram', 'Facebook',
      'Kbis', 'Urssaf', 'RC pro',
      'Lien dossier Drive'
    ]);
    sheet.getRange(1, 1, 1, 23)
      .setBackground('#b87a5c')
      .setFontColor('#f5f1e8')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const totalMissions = payload.missions.reduce((acc, cat) => acc + cat.missions.length, 0);
  const categoriesCount = payload.missions.length;

  sheet.appendRow([
    new Date(payload.timestamp),
    'À traiter',
    payload.profil.prenom,
    payload.profil.nom,
    payload.profil.email,
    payload.profil.telephone,
    payload.profil.societe,
    payload.profil.statutJuridique,
    payload.profil.siret,
    payload.profil.zone,
    payload.profil.experience,
    payload.profil.presentation,
    totalMissions,
    categoriesCount,
    payload.complements.site,
    payload.complements.tarif,
    payload.complements.linkedin,
    payload.complements.instagram,
    payload.complements.facebook,
    uploadedDocs.kbis ? uploadedDocs.kbis.url : '',
    uploadedDocs.urssaf ? uploadedDocs.urssaf.url : '',
    uploadedDocs.rcpro ? uploadedDocs.rcpro.url : '',
    folderInfo.url
  ]);
}

// ====================================================================
// 5. ACCUSÉ DE RÉCEPTION AU CANDIDAT
// ====================================================================
function sendAcknowledgment(payload) {
  const subject = 'Votre candidature aux In10pensables · Bien reçue';

  const htmlBody = `
    <div style="font-family: 'Helvetica Neue', Arial, sans-serif; color: #3d2314; max-width: 580px; margin: 0 auto; padding: 30px 20px;">
      <p style="font-size: 11px; letter-spacing: 0.3em; text-transform: uppercase; color: #b87a5c; margin: 0 0 14px 0;">Les In10pensables</p>

      <h1 style="font-family: Georgia, serif; font-size: 26px; font-weight: 500; color: #b87a5c; margin: 0 0 20px 0; line-height: 1.2;">
        Bonjour ${escapeHtml(payload.profil.prenom)},
      </h1>

      <p style="font-size: 14px; line-height: 1.7; color: #5a3a2a; margin: 0 0 18px 0;">
        Nous avons bien reçu votre candidature au collectif Les In10pensables. Merci de la confiance que vous nous témoignez.
      </p>

      <p style="font-size: 14px; line-height: 1.7; color: #5a3a2a; margin: 0 0 18px 0;">
        Nous allons étudier votre dossier dans les prochains jours et aurons le plaisir de vous contacter pour organiser un entretien.
      </p>

      <p style="font-size: 14px; line-height: 1.7; color: #5a3a2a; margin: 0 0 18px 0;">
        Nous espérons pouvoir rapidement vous compter parmi nos partenaires freelances.
      </p>

      <div style="margin: 30px 0; padding: 20px 24px; background: #f5f1e8; border-left: 3px solid #b87a5c; border-radius: 4px;">
        <p style="font-size: 12px; color: #a88974; margin: 0 0 8px 0; letter-spacing: 0.1em; text-transform: uppercase;">Récapitulatif</p>
        <p style="font-size: 13px; color: #3d2314; margin: 0; line-height: 1.6;">
          <strong>${escapeHtml(payload.profil.prenom)} ${escapeHtml(payload.profil.nom)}</strong><br>
          ${escapeHtml(payload.profil.societe || '—')}<br>
          ${payload.missions.reduce((acc, cat) => acc + cat.missions.length, 0)} compétence(s) renseignée(s)
        </p>
      </div>

      <p style="font-size: 13px; line-height: 1.7; color: #a88974; margin: 30px 0 0 0; font-style: italic;">
        Bien à vous,<br>
        Audrey Hubert &amp; Julie Cottereau<br>
        Co-fondatrices · Les In10pensables
      </p>

      <hr style="border: none; border-top: 1px solid #e8d8c4; margin: 30px 0 20px 0;">
      <p style="font-size: 11px; color: #a88974; margin: 0; text-align: center;">
        Cet email a été envoyé en réponse à votre candidature soumise le ${new Date(payload.timestamp).toLocaleDateString('fr-FR')}.
      </p>
    </div>
  `;

  // PDF soft (version candidate, sans les missions non souhaitées)
  const softPdf = generateRecapPdf(payload, true);

  GmailApp.sendEmail(
    payload.profil.email,
    subject,
    'Votre candidature a bien été reçue. Vous recevrez un retour prochainement.',
    {
      htmlBody: htmlBody,
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO,
      attachments: [softPdf]
    }
  );
}

// ====================================================================
// 6. NOTIFICATION AUX FONDATRICES
// ====================================================================
function notifyFounders(payload, folderInfo, uploadedDocs) {
  const totalMissions = payload.missions.reduce((acc, cat) => acc + cat.missions.length, 0);
  const subject = `Nouvelle candidature : ${payload.profil.prenom} ${payload.profil.nom}`;

  const docsList = [];
  if (uploadedDocs.kbis) docsList.push('Kbis');
  if (uploadedDocs.urssaf) docsList.push('Urssaf');
  if (uploadedDocs.rcpro) docsList.push('RC pro');
  const docsStr = docsList.length > 0 ? docsList.join(', ') : 'Aucun document fourni (à demander en entretien)';

  const htmlBody = `
    <div style="font-family: 'Helvetica Neue', Arial, sans-serif; color: #3d2314; max-width: 600px;">
      <h2 style="color: #b87a5c; font-family: Georgia, serif;">Nouvelle candidature reçue</h2>

      <table style="border-collapse: collapse; width: 100%; font-size: 13px;">
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold; width: 180px;">Candidat(e)</td><td style="padding: 8px 12px;">${escapeHtml(payload.profil.prenom)} ${escapeHtml(payload.profil.nom)}</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Société</td><td style="padding: 8px 12px;">${escapeHtml(payload.profil.societe || '—')} (${escapeHtml(payload.profil.statutJuridique || '—')})</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Email</td><td style="padding: 8px 12px;"><a href="mailto:${escapeHtml(payload.profil.email)}">${escapeHtml(payload.profil.email)}</a></td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Téléphone</td><td style="padding: 8px 12px;">${escapeHtml(payload.profil.telephone || '—')}</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Zone</td><td style="padding: 8px 12px;">${escapeHtml(payload.profil.zone || '—')}</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Expérience</td><td style="padding: 8px 12px;">${escapeHtml(payload.profil.experience || '—')}</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Missions positionnées</td><td style="padding: 8px 12px;"><strong>${totalMissions}</strong> sur ${payload.missions.length} catégorie(s)</td></tr>
        <tr><td style="padding: 8px 12px; background: #f5f1e8; font-weight: bold;">Documents fournis</td><td style="padding: 8px 12px;">${docsStr}</td></tr>
      </table>

      <p style="margin: 24px 0;">
        <a href="${folderInfo.url}" style="background: #b87a5c; color: #f5f1e8; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold; letter-spacing: 0.05em; text-transform: uppercase; font-size: 12px;">
          Ouvrir le dossier Drive
        </a>
      </p>

      <p style="font-size: 12px; color: #a88974;">
        La ligne de suivi est aussi disponible dans le Google Sheet "Candidatures".
      </p>
    </div>
  `;

  CONFIG.FOUNDERS_EMAILS.forEach(email => {
    GmailApp.sendEmail(
      email,
      subject,
      `Nouvelle candidature de ${payload.profil.prenom} ${payload.profil.nom}. Dossier Drive : ${folderInfo.url}`,
      {
        htmlBody: htmlBody,
        name: 'In10-Hub'
      }
    );
  });
}


// ====================================================================
// GÉNÉRATION DU PDF RÉCAP (charte In10pensables)
// soft = true  -> version candidate (sans les missions "non souhaitées")
// soft = false -> version complète (Audrey & Julie)
// ====================================================================
function generateRecapPdf(payload, soft) {
  const p = payload.profil;
  const c = payload.complements || {};
  const dateStr = Utilities.formatDate(new Date(payload.timestamp), 'Europe/Paris', 'dd MMMM yyyy');

  // Décompte par positionnement
  var nbDo = 0, nbDont = 0, nbLearn = 0;
  (payload.missions || []).forEach(function(cat) {
    cat.missions.forEach(function(m) {
      if (m.positionnement === 'Compétence maîtrisée') nbDo++;
      else if (m.positionnement === 'Compétence maîtrisée mais non souhaitée') nbDont++;
      else if (m.positionnement === 'Souhaite se former') nbLearn++;
    });
  });
  var total = nbDo + nbDont + nbLearn;

  // Construction des blocs de catégories
  var catsHtml = '';
  (payload.missions || []).forEach(function(cat) {
    var rows = '';
    var count = 0;
    cat.missions.forEach(function(m) {
      // En version soft, on masque les "non souhaitées"
      if (soft && m.positionnement === 'Compétence maîtrisée mais non souhaitée') return;
      var tagClass = 'pos-do', tagLabel = 'Maîtrisée';
      if (m.positionnement === 'Compétence maîtrisée mais non souhaitée') { tagClass = 'pos-dont'; tagLabel = 'Non souhaitée'; }
      else if (m.positionnement === 'Souhaite se former') { tagClass = 'pos-learn'; tagLabel = 'À se former'; }
      rows += '<div class="skill-item"><span class="pos-tag ' + tagClass + '">' + tagLabel + '</span><span class="skill-label">' + escapeHtml(m.mission) + '</span></div>';
      count++;
    });
    if (count > 0) {
      catsHtml += '<div class="cat-block"><div class="cat-name">' + escapeHtml(cat.categorie) +
        ' <span class="cat-count">· ' + count + ' positionnée(s)</span></div>' + rows + '</div>';
    }
  });

  // Bloc statistiques (en version soft, pas de "non souhaitées")
  var statsHtml = '<div class="skill-stat do"><div class="skill-stat-num">' + nbDo + '</div><div class="skill-stat-label">Compétences maîtrisées</div></div>';
  if (!soft) {
    statsHtml += '<div class="skill-stat dont"><div class="skill-stat-num">' + nbDont + '</div><div class="skill-stat-label">Maîtrisées mais non souhaitées</div></div>';
  }
  statsHtml += '<div class="skill-stat learn"><div class="skill-stat-num">' + nbLearn + '</div><div class="skill-stat-label">Souhaite se former</div></div>';

  // Compléments
  var compRows = '';
  if (c.site)      compRows += '<div class="comp-row"><span class="comp-label">Site / plaquette</span><span class="comp-value">' + escapeHtml(c.site) + '</span></div>';
  if (c.tarif)     compRows += '<div class="comp-row"><span class="comp-label">Doc tarifaire</span><span class="comp-value">' + escapeHtml(c.tarif) + '</span></div>';
  if (c.linkedin)  compRows += '<div class="comp-row"><span class="comp-label">LinkedIn</span><span class="comp-value">' + escapeHtml(c.linkedin) + '</span></div>';
  if (c.instagram) compRows += '<div class="comp-row"><span class="comp-label">Instagram</span><span class="comp-value">' + escapeHtml(c.instagram) + '</span></div>';
  if (c.facebook)  compRows += '<div class="comp-row"><span class="comp-label">Facebook</span><span class="comp-value">' + escapeHtml(c.facebook) + '</span></div>';
  var compSection = compRows ? '<div class="section"><div class="section-title">Compléments</div><div class="comp-grid">' + compRows + '</div></div>' : '';

  // Documents (version complète uniquement)
  var docsSection = '';
  if (!soft && payload.documents) {
    var chips = '';
    if (payload.documents.kbis)   chips += '<span class="doc-chip">✓ Extrait Kbis</span>';
    if (payload.documents.urssaf) chips += '<span class="doc-chip">✓ Attestation Urssaf</span>';
    if (payload.documents.rcpro)  chips += '<span class="doc-chip">✓ RC Pro</span>';
    if (chips) docsSection = '<div class="section"><div class="section-title">Pièces jointes</div><div class="docs-list">' + chips + '</div></div>';
  }

  var titreCompetences = soft
    ? 'Vos compétences renseignées'
    : 'Cartographie des compétences · ' + total + ' renseignée(s)';

  var presentation = p.presentation
    ? '<div class="presentation">« ' + escapeHtml(p.presentation) + ' »</div>'
    : '';

  var html =
  '<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><style>' +
  '* { margin:0; padding:0; box-sizing:border-box; }' +
  'body { font-family:Arial, Helvetica, sans-serif; color:#3d2314; }' +
  '.header { background:#f5f1e8; border-bottom:3px solid #b87a5c; padding:30px 40px 24px; position:relative; }' +
  '.header-eyebrow { font-size:10px; letter-spacing:0.25em; text-transform:uppercase; color:#b87a5c; margin-bottom:8px; font-weight:bold; }' +
  '.header-name { font-size:28px; color:#3d2314; }' +
  '.header-societe { font-size:14px; color:#a88974; margin-top:5px; }' +
  '.header-meta { margin-top:14px; font-size:12px; color:#5a4636; }' +
  '.header-meta span { margin-right:22px; }' +
  '.header-date { position:absolute; top:30px; right:40px; text-align:right; font-size:10px; color:#a88974; }' +
  '.section { padding:22px 40px 0; }' +
  '.section-title { font-size:15px; color:#b87a5c; margin-bottom:12px; padding-bottom:6px; border-bottom:1px solid #e8d8c4; }' +
  '.info-grid { width:100%; }' +
  '.info-row { font-size:12px; margin-bottom:7px; }' +
  '.info-label { color:#a88974; display:inline-block; width:150px; }' +
  '.info-value { color:#3d2314; font-weight:bold; }' +
  '.presentation { font-size:12px; line-height:1.6; color:#5a4636; background:#faf7f0; padding:13px 16px; border-left:3px solid #d4a888; font-style:italic; margin-top:12px; }' +
  '.skills-summary { width:100%; margin-bottom:16px; }' +
  '.skill-stat { display:inline-block; width:31%; text-align:center; padding:12px 4px; border-radius:6px; vertical-align:top; margin-right:1.5%; }' +
  '.skill-stat.do { background:#eef5f0; border:1px solid #6b9b7a; }' +
  '.skill-stat.dont { background:#fbf6e6; border:1px solid #d9ad3e; }' +
  '.skill-stat.learn { background:#eef2f7; border:1px solid #6a8caf; }' +
  '.skill-stat-num { font-size:24px; font-weight:bold; }' +
  '.skill-stat.do .skill-stat-num { color:#5c8a6b; }' +
  '.skill-stat.dont .skill-stat-num { color:#c79a2f; }' +
  '.skill-stat.learn .skill-stat-num { color:#5b7c9c; }' +
  '.skill-stat-label { font-size:10px; color:#5a4636; margin-top:3px; }' +
  '.cat-block { margin-bottom:14px; }' +
  '.cat-name { font-size:12px; font-weight:bold; color:#3d2314; margin-bottom:6px; }' +
  '.cat-count { font-size:10px; color:#a88974; font-weight:normal; }' +
  '.skill-item { font-size:11px; padding:3px 0 3px 14px; color:#5a4636; }' +
  '.pos-tag { font-size:9px; font-weight:bold; padding:2px 7px; border-radius:9px; color:#fff; margin-right:8px; }' +
  '.pos-do { background:#6b9b7a; } .pos-dont { background:#d9ad3e; } .pos-learn { background:#6a8caf; }' +
  '.comp-grid { font-size:12px; } .comp-row { margin-bottom:6px; }' +
  '.comp-label { color:#a88974; display:inline-block; width:110px; } .comp-value { color:#b87a5c; }' +
  '.docs-list span { font-size:10px; padding:5px 11px; background:#fdf5ee; border:1px solid #d4a888; border-radius:5px; color:#b87a5c; margin-right:8px; }' +
  '.footer { margin-top:26px; padding:14px 40px 0; border-top:1px solid #e8d8c4; font-size:9px; color:#a88974; text-align:center; line-height:1.5; }' +
  '</style></head><body>' +
  '<div class="header">' +
    '<div class="header-date">Candidature reçue le<br><b>' + dateStr + '</b></div>' +
    '<div class="header-eyebrow">Profil candidat·e · Les In10pensables</div>' +
    '<div class="header-name">' + escapeHtml(p.prenom) + ' ' + escapeHtml(p.nom) + '</div>' +
    '<div class="header-societe">' + escapeHtml(p.societe || '—') + (p.statutJuridique ? ' · ' + escapeHtml(p.statutJuridique) : '') + '</div>' +
    '<div class="header-meta"><span>✉ ' + escapeHtml(p.email) + '</span>' +
      (p.telephone ? '<span>☎ ' + escapeHtml(p.telephone) + '</span>' : '') +
      (p.zone ? '<span>⌖ ' + escapeHtml(p.zone) + '</span>' : '') + '</div>' +
  '</div>' +
  '<div class="section"><div class="section-title">Profil</div><div class="info-grid">' +
    (p.siret ? '<div class="info-row"><span class="info-label">SIRET</span><span class="info-value">' + escapeHtml(p.siret) + '</span></div>' : '') +
    (p.experience ? '<div class="info-row"><span class="info-label">Expérience</span><span class="info-value">' + escapeHtml(p.experience) + '</span></div>' : '') +
    (p.zone ? '<div class="info-row"><span class="info-label">Zone d\'intervention</span><span class="info-value">' + escapeHtml(p.zone) + '</span></div>' : '') +
    (p.adresse ? '<div class="info-row"><span class="info-label">Adresse</span><span class="info-value">' + escapeHtml(p.adresse) + '</span></div>' : '') +
  '</div>' + presentation + '</div>' +
  '<div class="section"><div class="section-title">' + titreCompetences + '</div>' +
    '<div class="skills-summary">' + statsHtml + '</div>' + catsHtml + '</div>' +
  compSection + docsSection +
  '<div class="footer"><b>Les In10pensables</b> · Réseau de professionnel·le·s de la gestion administrative<br>' +
    (soft ? 'Récapitulatif de votre candidature' : 'Document généré automatiquement · Confidentiel · Usage interne du réseau') + '</div>' +
  '</body></html>';

  // Conversion HTML -> PDF via le service Drive intégré
  var blob = Utilities.newBlob(html, 'text/html', 'recap.html');
  var pdf = blob.getAs('application/pdf');
  var nomFichier = soft
    ? 'Recap_' + p.nom.toUpperCase() + '_' + capitalize(p.prenom) + '.pdf'
    : 'Profil_' + p.nom.toUpperCase() + '_' + capitalize(p.prenom) + '.pdf';
  pdf.setName(nomFichier.replace(/[^a-zA-Z0-9_.-]/g, ''));
  return pdf;
}

// ====================================================================
// UTILS
// ====================================================================
function capitalize(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ====================================================================
// TEST RAPIDE (à exécuter une fois pour vérifier les permissions)
// ====================================================================
function testSetup() {
  // Vérifier accès Drive
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  Logger.log('Dossier Drive OK : ' + folder.getName());

  // Vérifier accès Sheet
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  Logger.log('Sheet OK : ' + ss.getName());

  // Tester l'envoi d'email à soi-même
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    '[Test] In10-Hub Apps Script configuré',
    'Si vous recevez ce mail, le script Apps Script est correctement configuré et a accès à Drive, Sheets et Gmail.'
  );
  Logger.log('Email test envoyé à ' + Session.getActiveUser().getEmail());
}
