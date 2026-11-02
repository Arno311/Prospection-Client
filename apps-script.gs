/**
 * Pont entre l'app de prospection et la feuille « Prospection ».
 *
 * MISE EN PLACE (5 minutes, une seule fois)
 * 1. Ouvre la feuille Google « Prospection » → Extensions → Apps Script.
 * 2. Efface le contenu, colle tout ce fichier, change CLE ci-dessous (au moins 12 caractères au hasard).
 * 3. Enregistre, puis Déployer → Nouveau déploiement → type « Application Web ».
 *      Exécuter en tant que : Moi
 *      Qui a accès : Tout le monde
 *    Autorise l'accès à ta feuille quand Google le demande, puis copie l'URL qui finit par /exec.
 * 4. Dans l'app : ⚙ Réglages → Synchronisation → colle l'URL et la même clé → « Tester la connexion ».
 *
 * Mise à jour du code (obligatoire pour iPhone/Safari) : Déployer → Gérer les déploiements → crayon → Version : Nouvelle version → Déployer. L'URL ne change pas.
 * Si tu modifies ce code plus tard : Déployer → Gérer les déploiements → crayon → Nouvelle version.
 *
 * Ce que fait le script : écrit l'état en colonne F, et la date du jour en colonne I
 * (« Dernier contact », créée si absente) pour que les relances (J+7…) soient exactes sur tous tes appareils.
 */
const CLE   = "CHANGE-MOI-par-une-longue-phrase";
const ONGLET = "Feuille 1";
const ETATS = ["À contacter", "Mail envoyé", "Relancé", "RDV", "Pas de réponse", "Ne pas contacter"];
const COL_NOM = 1, COL_ETAT = 6, COL_DATE = 9;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const S = o => sortie(o, p.callback);
  const lock = LockService.getScriptLock();
  try {
    if (p.cle !== CLE) return S({ ok: false, error: "clé invalide" });
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET);
    if (!sh) return S({ ok: false, error: "onglet « " + ONGLET + " » introuvable" });
    if (p.ping) return S({ ok: true, onglet: ONGLET, lignes: sh.getLastRow() - 1 });

    if (ETATS.indexOf(p.etat) < 0) return S({ ok: false, error: "état inconnu" });
    const ligne = parseInt(p.ligne, 10);
    if (!(ligne >= 2) || ligne > sh.getLastRow()) return S({ ok: false, error: "ligne invalide" });

    lock.waitLock(10000);
    // Garde-fou : si des lignes ont été triées ou insérées depuis la dernière lecture, on ne touche à rien.
    const nom = String(sh.getRange(ligne, COL_NOM).getValue()).trim();
    if (nom !== String(p.nom || "").trim())
      return S({ ok: false, error: "la ligne a bougé dans la feuille : actualise l'app (↻)" });

    sh.getRange(ligne, COL_ETAT).setValue(p.etat);
    if (!sh.getRange(1, COL_DATE).getValue()) sh.getRange(1, COL_DATE).setValue("Dernier contact");
    const d = sh.getRange(ligne, COL_DATE);
    d.setNumberFormat("@");
    d.setValue(p.etat === "À contacter" ? "" :
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    return S({ ok: true });
  } catch (err) {
    return S({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// Si l'app envoie un nom de fonction (callback), on répond en JavaScript : c'est ce qui marche sur iPhone/Safari.
function sortie(o, callback) {
  if (callback && /^[A-Za-z_][A-Za-z0-9_]*$/.test(callback))
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(o) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
