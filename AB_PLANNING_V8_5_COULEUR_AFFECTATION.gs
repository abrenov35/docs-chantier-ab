// AB PLANNING - APPS SCRIPT CRUD COMPLET
// V8.5 - couleur optionnelle par affectation (colonne N)
// Sheet ID: 1ub__l8g3LSiikfKST-43Q0PvpujqENf2VQm6XzK0vd4

const SHEET_ID = "1ub__l8g3LSiikfKST-43Q0PvpujqENf2VQm6XzK0vd4";


const CALENDAR_ID = "bgconstruction35@gmail.com";
// ============ MAIN HANDLER ============
function doGet(e) {

  const action =
    e.parameter.action ||
    "getAll";

  const callback =
    String(
      e.parameter.callback || ""
    ).trim();

  try {

    let result;

    switch (action) {

      case "getAll":
        result = {
          ouvriers: getOuvriers(),
          chantiers: getChantiers(),
          affectations: getAffectations()
        };
        break;

      case "getOuvriers":
        result = getOuvriers();
        break;

      case "getChantiers":
        result = getChantiers();
        break;

      case "getAffectations":
        result = getAffectations();
        break;

      case "createOuvrier":
        result = createOuvrier(
          e.parameter.nom,
          e.parameter.type,
          e.parameter.metier
        );
        break;

      case "createChantier":
        result = createChantier(
          e.parameter.nom,
          e.parameter.dateDebut,
          e.parameter.dateFin,
          e.parameter.description,
          e.parameter.couleur,
          e.parameter.dateSignature,
          e.parameter.typeChantier
        );
        break;

      case "createAffectation":
        result = createAffectation(
          e.parameter.ouvrierID,
          e.parameter.chantierId,
          e.parameter.dateDebut,
          e.parameter.dateFin,
          e.parameter.tache,
          e.parameter.nomAffectation,
          e.parameter.typeAffectation,
          e.parameter.couleur
        );
        break;

      case "updateOuvrier":
        result = updateOuvrier(
          e.parameter.id,
          e.parameter.nom,
          e.parameter.type,
          e.parameter.metier,
          e.parameter.statut,
          e.parameter.ordre,
          e.parameter.separateurApres
        );
        break;

      case "updateChantier":
        result = updateChantier(
          e.parameter.id,
          e.parameter.nom,
          e.parameter.dateDebut,
          e.parameter.dateFin,
          e.parameter.description,
          e.parameter.statut,
          e.parameter.couleur,
          e.parameter.dateSignature,
          e.parameter.typeChantier
        );
        break;

      case "updateAffectation":
        result = updateAffectation(
          e.parameter.id,
          e.parameter.dateDebut,
          e.parameter.dateFin,
          e.parameter.tache,
          e.parameter.statut,
          e.parameter.nomAffectation,
          e.parameter.chantierId,
          e.parameter.couleur
        );
        break;

      case "deleteChantier":
        result =
          deleteChantier(
            e.parameter.id
          );
        break;

      case "deleteAffectation":
        result =
          deleteAffectation(
            e.parameter.id
          );
        break;

      default:
        result = {
          error:
            "Action inconnue: " +
            action
        };
    }

    return sortieApiV55(
      result,
      callback
    );

  } catch (err) {

    return sortieApiV55(
      {
        error:
          err.toString(),
        message:
          err.message
      },
      callback
    );
  }
}


// ======================================================
// V5.5 - SORTIE API JSON + JSONP
//
// Sans callback : JSON classique.
// Avec callback : JavaScript JSONP.
//
// Le frontend GitHub Pages utilise JSONP pour éviter
// le blocage CORS entre abrenov35.github.io et
// script.google.com.
// ======================================================

function sortieApiV55(
  result,
  callback
) {

  const json =
    JSON.stringify(result);

  if (!callback) {

    return ContentService
      .createTextOutput(json)
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }

  // Sécurité : autoriser uniquement un nom de callback JS.
  if (
    !/^[A-Za-z_$][0-9A-Za-z_$]*$/
      .test(callback)
  ) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          error:
            "Callback JSONP invalide"
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }

  return ContentService
    .createTextOutput(
      callback +
      "(" +
      json +
      ");"
    )
    .setMimeType(
      ContentService
        .MimeType
        .JAVASCRIPT
    );
}


// ============ GET FUNCTIONS ============

function getOuvriers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("ouvriers");
  const data = sheet.getDataRange().getValues();
  const ouvriers = [];
  
  // Sauter la première ligne (headers)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "") break; // Fin des données
    
    ouvriers.push({
      id: data[i][0],
      nom: data[i][1],
      type: data[i][2],
      metier: data[i][3],
      statut: data[i][4],

      // F = ordre partagé PC / iPhone
      ordre:
        data[i][5] === "" ||
        data[i][5] === null ||
        typeof data[i][5] === "undefined"
          ? ""
          : Number(data[i][5]),

      // G = trait gris après cet ouvrier
      separateurApres:
        data[i][6] === true ||
        String(data[i][6] || "")
          .trim()
          .toUpperCase() === "TRUE"
    });
  }
  
  return ouvriers;
}

function formatDateSignatureApi_(value) {
  if (!value) return "";

  const texte = String(value).trim();
  const iso = texte.match(/^(\d{4})-(\d{2})/);
  if (iso) return iso[1] + "-" + iso[2];

  const date = new Date(value);
  if (isNaN(date.getTime())) return "";

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );
}

function getChantiers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("chantiers");
  const data = sheet.getDataRange().getValues();
  const chantiers = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "") break;
    
    chantiers.push({
      id: data[i][0],
      nom: data[i][1],
      dateDebut: data[i][2],
      dateFin: data[i][3],
      description: data[i][4],
      statut: data[i][5],
      // Colonne G : vide = couleur automatique
      couleur: String(data[i][6] || "").trim(),
      // Colonne H : mois et année de signature (AAAA-MM)
      dateSignature: formatDateSignatureApi_(data[i][7]),
      // Colonne I : Rénovation / Sinistre / Autre
      typeChantier: String(data[i][8] || "Rénovation").trim()
    });
  }
  
  return chantiers;
}

function getAffectations() {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();
  const affectations = [];

  for (let i = 1; i < data.length; i++) {

    // Ligne vide : on l'ignore,
    // mais on CONTINUE à lire les lignes suivantes.
    if (data[i][0] === "") {
      continue;
    }

    affectations.push({
      id: data[i][0],
      ouvrierID: data[i][1],
      chantierId: data[i][2],
      dateDebut: data[i][3],
      dateFin: data[i][4],
      tache: data[i][5],
      statut: data[i][6],
      googleEventId: data[i][7] || "",
      source: data[i][8] || "",
     dateModification: data[i][9] || "",
derniereSync: data[i][10] || "",
typeAffectation: data[i][11] || "CHANTIER",
nomExterne: data[i][12] || "",
couleur: String(data[i][13] || "").trim()
    });
  }

  return affectations;
}
// ============ CREATE FUNCTIONS ============

function createOuvrier(nom, type, metier) {
  if (!nom || !type || !metier) {
    return { error: "nom, type et metier sont obligatoires" };
  }
  
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("ouvriers");
  const data = sheet.getDataRange().getValues();
  
  // Trouver le prochain ID
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "") break;
    maxId = Math.max(maxId, parseInt(data[i][0]));
  }
  
  const newId = maxId + 1;
  const newRow = [newId, nom, type, metier, "Actif"];
  
  sheet.appendRow(newRow);
  
  return { success: true, id: newId, message: "Ouvrier créé" };
}

function createChantier(nom, dateDebut, dateFin, description, couleur, dateSignature, typeChantier) {
  if (!nom) {
    return { error: "Le nom est obligatoire" };
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("chantiers");
  const data = sheet.getDataRange().getValues();

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "") break;
    maxId = Math.max(maxId, parseInt(data[i][0]));
  }

  const newId = maxId + 1;
  const couleurFinale = String(couleur || "").trim();
  const dateSignatureFinale = String(dateSignature || "").trim().substring(0, 7);
  const typesAutorises = ["Rénovation", "Sinistre", "Autre"];
  const typeChantierFinal = typesAutorises.includes(String(typeChantier || "").trim())
    ? String(typeChantier).trim()
    : "Rénovation";

  if (couleurFinale && !/^#[0-9A-Fa-f]{6}$/.test(couleurFinale)) {
    return { error: "Couleur invalide. Format attendu : #RRGGBB" };
  }

  const newRow = [
    newId,
    nom,
    dateDebut,
    dateFin,
    description || "",
    "Actif",
    couleurFinale,
    dateSignatureFinale,
    typeChantierFinal
  ];

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  return {
    success: true,
    id: newId,
    couleur: couleurFinale,
    dateSignature: dateSignatureFinale,
    typeChantier: typeChantierFinal,
    message: "Chantier créé"
  };
}



function createAffectation(
  ouvrierID,
  chantierId,
  dateDebut,
  dateFin,
  tache,
  nomAffectation,
  typeAffectation,
  couleur
) {

  const type =
    normalizeAB(
      typeAffectation || "CHANTIER"
    );

  const estLibre =
    type === "HORS_GANTT" ||
    type === "DIVERS" ||
    type === "ABSENCE" ||
    type === "FORMATION";

  const nomLibre =
    String(
      nomAffectation || ""
    ).trim();

  if (
    !ouvrierID ||
    !dateDebut ||
    !dateFin
  ) {
    return {
      error:
        "Ouvrier et jours obligatoires"
    };
  }

  if (
    !estLibre &&
    !chantierId
  ) {
    return {
      error:
        "Choisissez un chantier"
    };
  }

  if (
    estLibre &&
    !nomLibre
  ) {
    return {
      error:
        "Donnez un nom à l'affectation"
    };
  }

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const newId =
    getNextAffectationId();

  const maintenant =
    new Date();

  const couleurFinale =
    String(couleur || "").trim();

  if (
    couleurFinale &&
    !/^#[0-9A-Fa-f]{6}$/.test(couleurFinale)
  ) {
    return {
      error:
        "Couleur d'affectation invalide"
    };
  }

  sheet.appendRow([
    newId,
    ouvrierID,
    estLibre ? "" : chantierId,
    dateDebut,
    dateFin,
    tache || "",
    "Actif",
    "",
    "GANTT",
    maintenant,
    maintenant,
    estLibre ? "HORS_GANTT" : "CHANTIER",
    estLibre ? nomLibre : "",
    couleurFinale
  ]);

  SpreadsheetApp.flush();

  mettreEnFileGoogleV75({
    action: "CREATE",
    affectationId:
      Number(newId)
  });

  Logger.log(
    "✅ V7.5 GANTT créé immédiatement | AffectationID=" +
    newId +
    " | Google en attente"
  );

  return {
    success: true,
    id: newId,
    message:
      estLibre
        ? "Autre affectation créée"
        : "Affectation chantier créée",
    googlePending: true,
    couleur: couleurFinale
  };
}

function creerEvenementsGoogleImmediatsV73(
  affectation
) {

  const calendar =
    CalendarApp.getCalendarById(
      CALENDAR_ID
    );

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " +
      CALENDAR_ID
    );
  }

  const ouvriers =
    getOuvriers();

  const chantiers =
    getChantiers();

  const ouvrier =
    ouvriers.find(function(o) {
      return (
        Number(o.id) ===
        Number(
          affectation.ouvrierID
        )
      );
    });

  if (!ouvrier) {
    throw new Error(
      "Ouvrier introuvable : " +
      affectation.ouvrierID
    );
  }

  const estLibre =
    normalizeAB(
      affectation.typeAffectation || ""
    ) === "HORS_GANTT";

  let titreGoogle = "";
  let chantier = null;

  if (estLibre) {

    titreGoogle =
      String(
        affectation.nomAffectation || ""
      ).trim();

    if (!titreGoogle) {
      throw new Error(
        "Nom d'affectation libre manquant"
      );
    }

  } else {

    chantier =
      chantiers.find(function(c) {
        return (
          Number(c.id) ===
          Number(
            affectation.chantierId
          )
        );
      });

    if (!chantier) {
      throw new Error(
        "Chantier introuvable : " +
        affectation.chantierId
      );
    }

    titreGoogle =
      chantier.nom;
  }

  const jours =
    abChaqueJour(
      affectation.dateDebut,
      affectation.dateFin
    );

  const cacheLignes = {};
  const idsCrees = [];

  jours.forEach(function(jour) {

    const dateObj =
      new Date(
        jour +
        "T12:00:00"
      );

    const jourSemaine =
      dateObj.getDay();

    if (
      jourSemaine === 0 ||
      jourSemaine === 6
    ) {
      return;
    }

    const dimanche =
      abDimancheSemaine(
        dateObj
      );

    const cleSemaine =
      Utilities.formatDate(
        dimanche,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    if (
      !cacheLignes[
        cleSemaine
      ]
    ) {

      cacheLignes[
        cleSemaine
      ] =
        abLignesOuvriersSemaine(
          calendar,
          dimanche,
          ouvriers
        );
    }

    const lignes =
      cacheLignes[
        cleSemaine
      ];

    const ligne =
      lignes.find(function(l) {
        return (
          Number(
            l.ouvrierID
          ) ===
          Number(
            affectation.ouvrierID
          )
        );
      });

    if (!ligne) {

      throw new Error(
        "Repère Google introuvable pour " +
        ouvrier.nom +
        " le " +
        jour
      );
    }

    const debutEvent =
      abDateAvecHeure(
        jour,
        ligne.heure,
        ligne.minute
      );

    const dureeMinutes =
      normalizeAB(
        ouvrier.nom
      ) === "MORVAN"
        ? 105
        : 60;

    const finEvent =
      new Date(
        debutEvent.getTime() +
        dureeMinutes *
        60 *
        1000
      );

    const description =
      "AB PLANNING\n" +
      "OuvrierID: " +
      affectation.ouvrierID +
      "\n" +
      (
        estLibre
          ? "Type: HORS_GANTT\nNom: " +
            titreGoogle
          : "ChantierID: " +
            affectation.chantierId
      ) +
      "\nAffectationID: " +
      affectation.affectationId +
      "\nTache: " +
      (
        affectation.tache ||
        ""
      );

    const eventCree =
      Calendar.Events.insert(
        {
          summary:
            titreGoogle,

          description:
            description,

          start: {
            dateTime:
              debutEvent.toISOString(),
            timeZone:
              "Europe/Paris"
          },

          end: {
            dateTime:
              finEvent.toISOString(),
            timeZone:
              "Europe/Paris"
          }
        },
        CALENDAR_ID
      );

    const eventId =
      String(
        eventCree.id || ""
      ).trim();

    if (!eventId) {
      throw new Error(
        "Google a créé un événement sans event.id"
      );
    }

    idsCrees.push(
      eventId
    );

    Logger.log(
      "⚡ GOOGLE IMMÉDIAT V7.3 + | " +
      ouvrier.nom +
      " | " +
      titreGoogle +
      " | " +
      jour +
      " | " +
      eventId
    );
  });

  return idsCrees;
}


function creerEvenementsGoogleImmediatsV57(
  affectation
) {

  const calendar =
    CalendarApp.getCalendarById(
      CALENDAR_ID
    );

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " +
      CALENDAR_ID
    );
  }

  const ouvriers =
    getOuvriers();

  const chantiers =
    getChantiers();

  const ouvrier =
    ouvriers.find(function(o) {
      return (
        Number(o.id) ===
        Number(
          affectation.ouvrierID
        )
      );
    });

  const chantier =
    chantiers.find(function(c) {
      return (
        Number(c.id) ===
        Number(
          affectation.chantierId
        )
      );
    });

  if (!ouvrier) {
    throw new Error(
      "Ouvrier introuvable : " +
      affectation.ouvrierID
    );
  }

  if (!chantier) {
    throw new Error(
      "Chantier introuvable : " +
      affectation.chantierId
    );
  }

  const jours =
    abChaqueJour(
      affectation.dateDebut,
      affectation.dateFin
    );

  const cacheLignes = {};
  const idsCrees = [];

  jours.forEach(function(jour) {

    // Le Gantt n'affiche/travaille que du lundi au vendredi.
    const dateObj =
      new Date(
        jour +
        "T12:00:00"
      );

    const jourSemaine =
      dateObj.getDay();

    if (
      jourSemaine === 0 ||
      jourSemaine === 6
    ) {
      return;
    }

    const dimanche =
      abDimancheSemaine(
        dateObj
      );

    const cleSemaine =
      Utilities.formatDate(
        dimanche,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    if (
      !cacheLignes[
        cleSemaine
      ]
    ) {

      cacheLignes[
        cleSemaine
      ] =
        abLignesOuvriersSemaine(
          calendar,
          dimanche,
          ouvriers
        );
    }

    const lignes =
      cacheLignes[
        cleSemaine
      ];

    const ligne =
      lignes.find(function(l) {
        return (
          Number(
            l.ouvrierID
          ) ===
          Number(
            affectation.ouvrierID
          )
        );
      });

    if (!ligne) {

      throw new Error(
        "Repère Google introuvable pour " +
        ouvrier.nom +
        " le " +
        jour
      );
    }

    const debutEvent =
      abDateAvecHeure(
        jour,
        ligne.heure,
        ligne.minute
      );

    const dureeMinutes =
      normalizeAB(
        ouvrier.nom
      ) === "MORVAN"
        ? 105
        : 60;

    const finEvent =
      new Date(
        debutEvent.getTime() +
        dureeMinutes *
        60 *
        1000
      );

    const eventCree =
      Calendar.Events.insert(
        {
          summary:
            chantier.nom,

          description:
            "AB PLANNING\n" +
            "OuvrierID: " +
            affectation.ouvrierID +
            "\n" +
            "ChantierID: " +
            affectation.chantierId +
            "\n" +
            "AffectationID: " +
            affectation.affectationId +
            "\n" +
            "Tache: " +
            (
              affectation.tache ||
              "ND"
            ),

          start: {
            dateTime:
              debutEvent.toISOString(),
            timeZone:
              "Europe/Paris"
          },

          end: {
            dateTime:
              finEvent.toISOString(),
            timeZone:
              "Europe/Paris"
          }
        },
        CALENDAR_ID
      );

    const eventId =
      String(
        eventCree.id || ""
      ).trim();

    if (!eventId) {
      throw new Error(
        "Google a créé un événement sans event.id"
      );
    }

    idsCrees.push(
      eventId
    );

    Logger.log(
      "⚡ GOOGLE IMMÉDIAT V5.7 + | " +
      ouvrier.nom +
      " | " +
      chantier.nom +
      " | " +
      jour +
      " | " +
      eventId
    );
  });

  return idsCrees;
}

// ============ UPDATE FUNCTIONS ============

function updateOuvrier(
  id,
  nom,
  type,
  metier,
  statut,
  ordre,
  separateurApres
) {

  if (!id) {
    return {
      error: "ID obligatoire"
    };
  }

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("ouvriers");

  const data =
    sheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < data.length; i++) {

    if (
      parseInt(data[i][0], 10) !==
      parseInt(id, 10)
    ) {
      continue;
    }

    if (nom) {
      sheet
        .getRange(i + 1, 2)
        .setValue(nom);
    }

    if (type) {
      sheet
        .getRange(i + 1, 3)
        .setValue(type);
    }

    if (metier) {
      sheet
        .getRange(i + 1, 4)
        .setValue(metier);
    }

    if (statut) {
      sheet
        .getRange(i + 1, 5)
        .setValue(statut);
    }

    // ==================================================
    // V8.2 - ORDRE PARTAGÉ ENTRE TOUS LES APPAREILS
    // F = ordre
    //
    // Si le frontend envoie une valeur vide, on efface F.
    // Cela permet notamment de retirer proprement un ouvrier
    // de l'ordre actif lors de son archivage.
    // ==================================================

    if (
      ordre !== undefined &&
      ordre !== null
    ) {

      const ordreTexte =
        String(ordre).trim();

      if (ordreTexte === "") {

        sheet
          .getRange(i + 1, 6)
          .clearContent();

      } else {

        const ordreNombre =
          Number(ordreTexte);

        if (
          !isNaN(ordreNombre) &&
          ordreNombre > 0
        ) {

          sheet
            .getRange(i + 1, 6)
            .setValue(ordreNombre);
        }
      }
    }

    // ==================================================
    // V8.2 - SÉPARATEUR PARTAGÉ
    // G = separateurApres
    // ==================================================

    if (
      separateurApres !== undefined &&
      separateurApres !== null &&
      String(separateurApres).trim() !== ""
    ) {

      const actif =
        separateurApres === true ||
        String(separateurApres)
          .trim()
          .toUpperCase() === "TRUE";

      sheet
        .getRange(i + 1, 7)
        .setValue(actif);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message:
        "Ouvrier modifié - ordre partagé enregistré"
    };
  }

  return {
    error:
      "Ouvrier non trouvé"
  };
}

function updateChantier(id, nom, dateDebut, dateFin, description, statut, couleur, dateSignature, typeChantier) {
  if (!id) return { error: "ID obligatoire" };

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("chantiers");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(id)) {
      if (nom) sheet.getRange(i + 1, 2).setValue(nom);
      if (dateDebut) sheet.getRange(i + 1, 3).setValue(dateDebut);
      if (dateFin) sheet.getRange(i + 1, 4).setValue(dateFin);

      if (description !== undefined && description !== null) {
        sheet.getRange(i + 1, 5).setValue(description);
      }

      if (statut) sheet.getRange(i + 1, 6).setValue(statut);

      // Colonne G : vide = retour en couleur automatique
      if (couleur !== undefined && couleur !== null) {
        const couleurFinale = String(couleur).trim();

        if (couleurFinale === "") {
          sheet.getRange(i + 1, 7).clearContent();
        } else {
          if (!/^#[0-9A-Fa-f]{6}$/.test(couleurFinale)) {
            return { error: "Couleur invalide. Format attendu : #RRGGBB" };
          }
          sheet.getRange(i + 1, 7).setValue(couleurFinale);
        }
      }

      // Colonne H : date de signature optionnelle au format AAAA-MM
      if (dateSignature !== undefined && dateSignature !== null) {
        const dateSignatureFinale = String(dateSignature).trim().substring(0, 7);
        if (dateSignatureFinale) {
          sheet.getRange(i + 1, 8).setValue(dateSignatureFinale);
        } else {
          sheet.getRange(i + 1, 8).clearContent();
        }
      }

      // Colonne I : type de chantier
      if (typeChantier !== undefined && typeChantier !== null) {
        const typeFinal = String(typeChantier).trim();
        const typesAutorises = ["Rénovation", "Sinistre", "Autre"];
        sheet.getRange(i + 1, 9).setValue(
          typesAutorises.includes(typeFinal) ? typeFinal : "Autre"
        );
      }

      SpreadsheetApp.flush();
      return { success: true, message: "Chantier modifié" };
    }
  }

  return { error: "Chantier non trouvé" };
}


// ======================================================
// V6.3 - TOMBSTONES ANTI-COURSE
//
// Problème corrigé :
// 1) suppression dans le Gantt ;
// 2) Google n'a pas encore propagé la suppression ;
// 3) trigger Calendar relit encore l'ancien event actif ;
// 4) il recrée l'affectation dans le Gantt.
//
// Solution : dès qu'un event est supprimé/remplacé depuis
// le Gantt, son event.id est "gelé" pendant 10 minutes.
// Les synchronisations Google -> Gantt ignorent ce retour
// transitoire. Le cycle complet ne peut pas le ressusciter.
// ======================================================

const AB_TOMBSTONE_TTL_MS_V63 = 12 * 60 * 1000;
const AB_TOMBSTONE_PROP_V63 = "AB_GOOGLE_TOMBSTONES_V63";

function lireTombstonesV63() {

  const props =
    PropertiesService.getScriptProperties();

  const brut =
    props.getProperty(
      AB_TOMBSTONE_PROP_V63
    );

  let map = {};

  if (brut) {
    try {
      map = JSON.parse(brut) || {};
    } catch (e) {
      map = {};
    }
  }

  const maintenant =
    new Date().getTime();

  let modifie = false;

  Object.keys(map)
    .forEach(function(id) {

      if (
        !map[id] ||
        Number(map[id]) <= maintenant
      ) {
        delete map[id];
        modifie = true;
      }
    });

  if (modifie) {
    props.setProperty(
      AB_TOMBSTONE_PROP_V63,
      JSON.stringify(map)
    );
  }

  return map;
}

function ajouterTombstonesV63(ids) {

  const liste =
    (ids || [])
      .map(function(id) {
        return normaliserGoogleIdV6(id);
      })
      .filter(Boolean);

  if (liste.length === 0) {
    return;
  }

  const props =
    PropertiesService.getScriptProperties();

  const map =
    lireTombstonesV63();

  const expiration =
    new Date().getTime() +
    AB_TOMBSTONE_TTL_MS_V63;

  liste.forEach(function(id) {
    map[id] = expiration;
  });

  props.setProperty(
    AB_TOMBSTONE_PROP_V63,
    JSON.stringify(map)
  );

  Logger.log(
    "🪦 TOMBSTONE V6.3 + | IDs=" +
    liste.join(",")
  );
}

function estTombstoneV63(id) {

  const idN =
    normaliserGoogleIdV6(id);

  if (!idN) {
    return false;
  }

  const map =
    lireTombstonesV63();

  return Boolean(map[idN]);
}

function retirerTombstoneV63(id) {

  const idN =
    normaliserGoogleIdV6(id);

  if (!idN) {
    return;
  }

  const props =
    PropertiesService.getScriptProperties();

  const map =
    lireTombstonesV63();

  if (map[idN]) {
    delete map[idN];

    props.setProperty(
      AB_TOMBSTONE_PROP_V63,
      JSON.stringify(map)
    );
  }
}


function updateAffectation(
  id,
  dateDebut,
  dateFin,
  tache,
  statut,
  nomAffectation,
  nouveauChantierId,
  couleur
) {

  if (!id) {
    return { error: "ID obligatoire" };
  }

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const data =
    sheet.getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      Number(data[i][0]) !==
      Number(id)
    ) {
      continue;
    }

    if (
      couleur !== undefined &&
      couleur !== null
    ) {

      const couleurFinale =
        String(couleur).trim();

      if (
        couleurFinale &&
        !/^#[0-9A-Fa-f]{6}$/.test(couleurFinale)
      ) {
        return {
          error:
            "Couleur d'affectation invalide"
        };
      }

      if (couleurFinale) {
        sheet
          .getRange(i + 1, 14)
          .setValue(couleurFinale);
      } else {
        sheet
          .getRange(i + 1, 14)
          .clearContent();
      }
    }

    const chantierIdActuel =
      data[i][2] === "" ||
      data[i][2] === null
        ? ""
        : Number(data[i][2]);

    const ancienneDateDebut =
      abIsoDate(data[i][3]);

    const ancienneDateFin =
      abIsoDate(data[i][4]);

    const ancienneTache =
      data[i][5] === null ||
      data[i][5] === undefined
        ? ""
        : String(data[i][5]);

    const typeAffectation =
      normalizeAB(
        data[i][11] ||
        (
          data[i][2] === "" ||
          data[i][2] === null
            ? "HORS_GANTT"
            : "CHANTIER"
        )
      );

    const anciensIds =
      abListeIdsSafe(
        data[i][7]
      );

    const nouvelleDateDebut =
      abIsoDate(
        dateDebut ||
        ancienneDateDebut
      );

    const nouvelleDateFin =
      abIsoDate(
        dateFin ||
        ancienneDateFin
      );

    const nouvelleTache =
      tache === undefined ||
      tache === null
        ? ancienneTache
        : String(tache);

    const datesChangees =
      nouvelleDateDebut !==
      ancienneDateDebut ||
      nouvelleDateFin !==
      ancienneDateFin;

    if (
      typeAffectation ===
      "HORS_GANTT"
    ) {

      const ancienNom =
        String(
          data[i][12] || ""
        ).trim();

      const nouveauNom =
        String(
          nomAffectation !== undefined &&
          nomAffectation !== null &&
          String(
            nomAffectation
          ).trim() !== ""
            ? nomAffectation
            : ancienNom
        ).trim();

      if (!nouveauNom) {
        return {
          error:
            "Le nom de l'affectation est obligatoire"
        };
      }

      sheet
        .getRange(i + 1, 4)
        .setValue(
          nouvelleDateDebut
        );

      sheet
        .getRange(i + 1, 5)
        .setValue(
          nouvelleDateFin
        );

      sheet
        .getRange(i + 1, 6)
        .setValue(
          nouvelleTache
        );

      if (statut) {

        sheet
          .getRange(i + 1, 7)
          .setValue(
            statut
          );
      }

      sheet
        .getRange(i + 1, 13)
        .setValue(
          nouveauNom
        );

      sheet
        .getRange(i + 1, 11)
        .setValue(
          new Date()
        );

      SpreadsheetApp.flush();

      if (datesChangees) {

        ajouterTombstonesV63(
          anciensIds
        );

        mettreEnFileGoogleV75({
          action: "REPLACE",
          affectationId:
            Number(id),
          oldEventIds:
            anciensIds
        });

      } else {

        mettreEnFileGoogleV75({
          action: "PATCH",
          affectationId:
            Number(id)
        });
      }

      return {
        success: true,
        message:
          "Affectation enregistrée dans Gantt",
        googlePending: true
      };
    }

    const chantierId =
      nouveauChantierId
        ? Number(
            nouveauChantierId
          )
        : Number(
            chantierIdActuel
          );

    if (
      !chantierId ||
      isNaN(chantierId)
    ) {
      return {
        error:
          "Chantier invalide"
      };
    }

    const chantierChange =
      Number(chantierId) !==
      Number(chantierIdActuel);

    sheet
      .getRange(i + 1, 3)
      .setValue(
        chantierId
      );

    sheet
      .getRange(i + 1, 4)
      .setValue(
        nouvelleDateDebut
      );

    sheet
      .getRange(i + 1, 5)
      .setValue(
        nouvelleDateFin
      );

    sheet
      .getRange(i + 1, 6)
      .setValue(
        nouvelleTache
      );

    if (statut) {

      sheet
        .getRange(i + 1, 7)
        .setValue(
          statut
        );
    }

    sheet
      .getRange(i + 1, 9)
      .setValue("GANTT");

    sheet
      .getRange(i + 1, 11)
      .setValue(
        new Date()
      );

    SpreadsheetApp.flush();

    if (
      datesChangees ||
      chantierChange
    ) {

      ajouterTombstonesV63(
        anciensIds
      );

      mettreEnFileGoogleV75({
        action: "REPLACE",
        affectationId:
          Number(id),
        oldEventIds:
          anciensIds
      });

    } else {

      mettreEnFileGoogleV75({
        action: "PATCH",
        affectationId:
          Number(id)
      });
    }

    return {
      success: true,
      message:
        "Affectation enregistrée dans Gantt",
      googlePending: true
    };
  }

  return {
    error:
      "Affectation non trouvée"
  };
}

function construireDescriptionGoogleV74(
  affectation
) {

  let lignes = [
    "AB PLANNING",
    "OuvrierID: " +
      affectation.ouvrierID
  ];

  if (
    normalizeAB(
      affectation.typeAffectation || ""
    ) === "HORS_GANTT"
  ) {

    lignes.push(
      "Type: HORS_GANTT"
    );

    lignes.push(
      "Nom: " +
      String(
        affectation.nomAffectation || ""
      )
    );

  } else {

    lignes.push(
      "ChantierID: " +
      affectation.chantierId
    );
  }

  lignes.push(
    "AffectationID: " +
    affectation.affectationId
  );

  lignes.push(
    "Tache: " +
    String(
      affectation.tache || ""
    )
  );

  return lignes.join("\n");
}



// ============ DELETE FUNCTIONS ============


// ======================================================
// V8.3 - SUPPRESSION CHANTIER PROTÉGÉE
//
// Règle : suppression autorisée uniquement si AUCUNE
// affectation ne référence ce chantier.
// ======================================================

function deleteChantier(id) {

  if (!id) {
    return {
      error: "ID obligatoire"
    };
  }

  const spreadsheet =
    SpreadsheetApp.openById(SHEET_ID);

  const affectationsSheet =
    spreadsheet.getSheetByName("affectations");

  const affectations =
    affectationsSheet
      .getDataRange()
      .getValues();

  let nbAffectations = 0;

  for (let i = 1; i < affectations.length; i++) {

    if (affectations[i][0] === "") {
      continue;
    }

    if (
      Number(affectations[i][2]) ===
      Number(id)
    ) {
      nbAffectations++;
    }
  }

  if (nbAffectations > 0) {
    return {
      success: false,
      code: "HAS_AFFECTATIONS",
      count: nbAffectations,
      error:
        "Suppression impossible : ce chantier possède encore " +
        nbAffectations +
        " affectation(s)."
    };
  }

  const chantiersSheet =
    spreadsheet.getSheetByName("chantiers");

  const chantiers =
    chantiersSheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < chantiers.length; i++) {

    if (
      Number(chantiers[i][0]) !==
      Number(id)
    ) {
      continue;
    }

    const nom =
      String(chantiers[i][1] || "");

    chantiersSheet.deleteRow(i + 1);

    SpreadsheetApp.flush();

    Logger.log(
      "🗑 CHANTIER SUPPRIMÉ | ID=" +
      id +
      " | " +
      nom
    );

    return {
      success: true,
      id: Number(id),
      nom: nom,
      message: "Chantier supprimé"
    };
  }

  return {
    error: "Chantier non trouvé"
  };
}


function deleteAffectation(id) {

  if (!id) {
    return {
      error:
        "ID obligatoire"
    };
  }

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const data =
    sheet.getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      Number(data[i][0]) !==
      Number(id)
    ) {
      continue;
    }

    const googleIds =
      abListeIdsSafe(
        data[i][7]
      );

    // Bloquer immédiatement toute réapparition venant de Google
    // pendant que la suppression attend dans la file.
    ajouterTombstonesV63(
      googleIds
    );

    sheet.deleteRow(
      i + 1
    );

    SpreadsheetApp.flush();

    if (
      googleIds.length > 0
    ) {

      mettreEnFileGoogleV75({
        action: "DELETE",
        affectationId:
          Number(id),
        eventIds:
          googleIds
      });
    }

    Logger.log(
      "🗑 V7.5 GANTT supprimé immédiatement | AffectationID=" +
      id +
      " | Google en attente=" +
      googleIds.length
    );

    return {
      success: true,
      message:
        "Affectation supprimée du Gantt",
      googlePending:
        googleIds.length > 0
    };
  }

  return {
    error:
      "Affectation non trouvée"
  };
}

function normalizeAB(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}


/**
 * Test de lecture du planning Google.
 * NE MODIFIE RIEN.
 */
function testLectureGooglePlanning() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " + CALENDAR_ID
    );
  }

  const start = new Date("2026-08-01T00:00:00");
  const end   = new Date("2026-10-31T23:59:59");

  const events = calendar.getEvents(start, end);

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();

  Logger.log("========================================");
  Logger.log("AB PLANNING - TEST GOOGLE CALENDAR");
  Logger.log("Agenda : " + calendar.getName());
  Logger.log("Événements trouvés : " + events.length);
  Logger.log("========================================");

  events.forEach(function(event) {

    const titre = event.getTitle();

    const debut = event.getStartTime();
    const fin = event.getEndTime();

    const titreNormalise = normalizeAB(titre);

    // Recherche chantier
    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) === titreNormalise;
    });

    Logger.log("----------------------------------------");
    Logger.log("Titre Google : " + titre);
    Logger.log("Début : " + debut);
    Logger.log("Fin : " + fin);

    if (chantier) {
      Logger.log(
        "✅ CHANTIER RECONNU : " +
        chantier.nom +
        " | ID = " +
        chantier.id
      );
    } else {
      Logger.log(
        "⚪ Hors Gantt / chantier non reconnu"
      );
    }
  });

  Logger.log("========================================");
  Logger.log("TEST TERMINÉ - AUCUNE DONNÉE MODIFIÉE");

}function testAgendaAccessible() {
  const agendas = CalendarApp.getAllCalendars();

  Logger.log("Nombre d'agendas accessibles : " + agendas.length);

  agendas.forEach(function(cal) {
    Logger.log(
      "Nom = " + cal.getName() +
      " | ID = " + cal.getId() +
      " | Propriétaire = " + cal.isOwnedByMe()
    );
  });
}
function testReconnaissanceOuvriersEtChantiers() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const ouvriers = getOuvriers();
  const chantiers = getChantiers();

  const start = new Date("2026-08-30T00:00:00");
  const end = new Date("2026-09-07T23:59:59");

  const events = calendar.getEvents(start, end);

  // 1. Trouver les événements du dimanche qui correspondent aux ouvriers
  const dimanche = new Date("2026-08-30T00:00:00");
  const lundi = new Date("2026-08-31T00:00:00");

  const eventsDimanche = calendar.getEvents(dimanche, lundi);

  const lignesOuvriers = [];

  eventsDimanche.forEach(function(event) {
    const titre = normalizeAB(event.getTitle());

    let ouvrier = ouvriers.find(function(o) {
      return normalizeAB(o.nom) === titre;
    });

    // Alias connus
    if (!ouvrier && titre === "ALEX") {
      ouvrier = ouvriers.find(o => normalizeAB(o.nom) === "ALEXANDRE");
    }

    if (!ouvrier && titre === "MOHAMED") {
      ouvrier = ouvriers.find(o => normalizeAB(o.nom) === "MOMO");
    }

    if (!ouvrier && titre === "UMAR") {
      ouvrier = ouvriers.find(o =>
        normalizeAB(o.nom) === "EQUIPE UMAR"
      );
    }

    if (ouvrier) {
      lignesOuvriers.push({
        ouvrierID: ouvrier.id,
        nom: ouvrier.nom,
        heure: event.getStartTime().getHours(),
        minute: event.getStartTime().getMinutes()
      });
    }
  });

  Logger.log("===== LIGNES OUVRIERS DÉTECTÉES =====");

  lignesOuvriers.forEach(function(ligne) {
    Logger.log(
      ligne.nom +
      " | ID=" + ligne.ouvrierID +
      " | " +
      String(ligne.heure).padStart(2, "0") +
      ":" +
      String(ligne.minute).padStart(2, "0")
    );
  });

  Logger.log("===== AFFECTATIONS RECONNUES =====");

  events.forEach(function(event) {

    const date = event.getStartTime();

    // Ignorer dimanche
    if (date.getDay() === 0) return;

    const titre = normalizeAB(event.getTitle());

    // Chercher chantier exact
    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) === titre;
    });

    if (!chantier) return;

    const h = date.getHours();
    const m = date.getMinutes();

    // Trouver la ligne correspondant exactement au créneau
    const ligne = lignesOuvriers.find(function(l) {
      return l.heure === h && l.minute === m;
    });

    if (!ligne) {
      Logger.log(
        "⚠️ Chantier reconnu mais ouvrier non trouvé : " +
        event.getTitle() +
        " | " +
        date
      );
      return;
    }

    Logger.log(
      "✅ " +
      Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy") +
      " | " +
      ligne.nom +
      " (ID " + ligne.ouvrierID + ")" +
      " | " +
      chantier.nom +
      " (chantier " + chantier.id + ")"
    );
  });

  Logger.log("===== FIN TEST =====");
}
function testConsolidationAffectations() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const ouvriers = getOuvriers();
  const chantiers = getChantiers();

  const start = new Date("2026-08-30T00:00:00");
  const end = new Date("2026-09-08T23:59:59");

  // Dimanche de référence
  const dimanche = new Date("2026-08-30T00:00:00");
  const lundi = new Date("2026-08-31T00:00:00");

  const eventsDimanche = calendar.getEvents(dimanche, lundi);

  const lignesOuvriers = [];

  eventsDimanche.forEach(function(event) {
    const titre = normalizeAB(event.getTitle());

    let ouvrier = ouvriers.find(function(o) {
      return normalizeAB(o.nom) === titre;
    });

    if (!ouvrier && titre === "ALEX") {
      ouvrier = ouvriers.find(o => normalizeAB(o.nom) === "ALEXANDRE");
    }

    if (!ouvrier && titre === "MOHAMED") {
      ouvrier = ouvriers.find(o => normalizeAB(o.nom) === "MOMO");
    }

    if (!ouvrier && titre === "UMAR") {
      ouvrier = ouvriers.find(o => normalizeAB(o.nom) === "EQUIPE UMAR");
    }

    if (ouvrier) {
      lignesOuvriers.push({
        ouvrierID: Number(ouvrier.id),
        nom: ouvrier.nom,
        heure: event.getStartTime().getHours(),
        minute: event.getStartTime().getMinutes()
      });
    }
  });

  const events = calendar.getEvents(start, end);

  const jours = [];

  events.forEach(function(event) {
    const date = event.getStartTime();

    if (date.getDay() === 0) return;

    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) === normalizeAB(event.getTitle());
    });

    if (!chantier) return;

    const ligne = lignesOuvriers.find(function(l) {
      return l.heure === date.getHours() &&
             l.minute === date.getMinutes();
    });

    if (!ligne) return;

    const dateISO = Utilities.formatDate(
      date,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    jours.push({
      ouvrierID: Number(ligne.ouvrierID),
      ouvrierNom: ligne.nom,
      chantierId: Number(chantier.id),
      chantierNom: chantier.nom,
      date: dateISO
    });
  });

  // Trier
  jours.sort(function(a, b) {
    if (a.ouvrierID !== b.ouvrierID) {
      return a.ouvrierID - b.ouvrierID;
    }

    if (a.chantierId !== b.chantierId) {
      return a.chantierId - b.chantierId;
    }

    return a.date.localeCompare(b.date);
  });

  // Regrouper les jours consécutifs
  const groupes = [];

  jours.forEach(function(item) {
    const dernier = groupes[groupes.length - 1];

    if (
      dernier &&
      dernier.ouvrierID === item.ouvrierID &&
      dernier.chantierId === item.chantierId
    ) {
      const dateDerniere = new Date(dernier.dateFin + "T12:00:00");
      const dateCourante = new Date(item.date + "T12:00:00");

      const diff =
        Math.round((dateCourante - dateDerniere) / 86400000);

      if (diff === 1) {
        dernier.dateFin = item.date;
        return;
      }
    }

    groupes.push({
      ouvrierID: item.ouvrierID,
      ouvrierNom: item.ouvrierNom,
      chantierId: item.chantierId,
      chantierNom: item.chantierNom,
      dateDebut: item.date,
      dateFin: item.date
    });
  });

  Logger.log("===== AFFECTATIONS CONSOLIDÉES =====");

  groupes.forEach(function(g) {
    Logger.log(
      g.ouvrierNom +
      " | " +
      g.chantierNom +
      " | " +
      g.dateDebut +
      " -> " +
      g.dateFin
    );
  });

  Logger.log("===== FIN TEST =====");
}
// ======================================================
// DRY RUN GOOGLE -> GANTT
// Compare Google Calendar avec la feuille affectations
// AUCUNE ÉCRITURE
// ======================================================

function testComparaisonGoogleGantt() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();

  // Semaine test
  const dimanche = new Date("2026-08-30T00:00:00");
  const lundi = new Date("2026-08-31T00:00:00");
  const finTest = new Date("2026-09-09T00:00:00");

  // --------------------------------------------------
  // 1. CONSTRUIRE LES LIGNES OUVRIERS DEPUIS DIMANCHE
  // --------------------------------------------------

  const eventsDimanche = calendar.getEvents(dimanche, lundi);
  const lignesOuvriers = [];

  eventsDimanche.forEach(function(event) {

    const titre = normalizeAB(event.getTitle());

    let ouvrier = ouvriers.find(function(o) {
      return normalizeAB(o.nom) === titre;
    });

    // Alias Google -> Sheet
    if (!ouvrier && titre === "ALEX") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "ALEXANDRE";
      });
    }

    if (!ouvrier && titre === "MOHAMED") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "MOMO";
      });
    }

    if (!ouvrier && titre === "UMAR") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "EQUIPE UMAR";
      });
    }

    if (ouvrier) {
      lignesOuvriers.push({
        ouvrierID: Number(ouvrier.id),
        nom: ouvrier.nom,
        heure: event.getStartTime().getHours(),
        minute: event.getStartTime().getMinutes()
      });
    }
  });


  // --------------------------------------------------
  // 2. LIRE LES ÉVÉNEMENTS GOOGLE RECONNUS
  // --------------------------------------------------

  const events = calendar.getEvents(lundi, finTest);
  const jours = [];

  events.forEach(function(event) {

    const date = event.getStartTime();

    // Ignorer les dimanches
    if (date.getDay() === 0) return;

    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) === normalizeAB(event.getTitle());
    });

    // Petit chantier absent du Gantt = ignoré
    if (!chantier) return;

    const ligne = lignesOuvriers.find(function(l) {
      return (
        l.heure === date.getHours() &&
        l.minute === date.getMinutes()
      );
    });

    if (!ligne) return;

    const dateISO = Utilities.formatDate(
      date,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    jours.push({
      ouvrierID: ligne.ouvrierID,
      ouvrierNom: ligne.nom,
      chantierId: Number(chantier.id),
      chantierNom: chantier.nom,
      date: dateISO
    });
  });


  // --------------------------------------------------
  // 3. TRIER
  // --------------------------------------------------

  jours.sort(function(a, b) {

    if (a.ouvrierID !== b.ouvrierID) {
      return a.ouvrierID - b.ouvrierID;
    }

    if (a.chantierId !== b.chantierId) {
      return a.chantierId - b.chantierId;
    }

    return a.date.localeCompare(b.date);
  });


  // --------------------------------------------------
  // 4. CONSOLIDER LES JOURS CONSÉCUTIFS
  // --------------------------------------------------

  const groupes = [];

  jours.forEach(function(item) {

    const dernier = groupes[groupes.length - 1];

    if (
      dernier &&
      dernier.ouvrierID === item.ouvrierID &&
      dernier.chantierId === item.chantierId
    ) {

      const d1 = new Date(dernier.dateFin + "T12:00:00");
      const d2 = new Date(item.date + "T12:00:00");

      const diff = Math.round(
        (d2.getTime() - d1.getTime()) / 86400000
      );

      if (diff === 1) {
        dernier.dateFin = item.date;
        return;
      }
    }

    groupes.push({
      ouvrierID: item.ouvrierID,
      ouvrierNom: item.ouvrierNom,
      chantierId: item.chantierId,
      chantierNom: item.chantierNom,
      dateDebut: item.date,
      dateFin: item.date
    });
  });


  // --------------------------------------------------
  // 5. NORMALISER UNE DATE DU SHEET
  // --------------------------------------------------

  function dateSheetISO(value) {

    if (!value) return "";

    // Si Apps Script reçoit déjà une vraie Date
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(
        value,
        "Europe/Paris",
        "yyyy-MM-dd"
      );
    }

    const texte = String(value).trim();

    // YYYY-MM-DD ou date ISO
    const iso = texte.match(/^(\d{4})-(\d{2})-(\d{2})/);

    if (iso) {
      return iso[1] + "-" + iso[2] + "-" + iso[3];
    }

    // DD/MM/YYYY
    const fr = texte.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);

    if (fr) {
      return fr[3] + "-" + fr[2] + "-" + fr[1];
    }

    return "";
  }


  // --------------------------------------------------
  // 6. COMPARER GOOGLE AVEC AFFECTATIONS
  // --------------------------------------------------

  Logger.log("===== COMPARAISON GOOGLE -> GANTT =====");

  let nbOK = 0;
  let nbModifier = 0;
  let nbCreer = 0;

  groupes.forEach(function(g) {

    // Toutes les affectations du même ouvrier
    // sur le même chantier
    const candidates = affectations.filter(function(a) {
      return (
        Number(a.ouvrierID) === Number(g.ouvrierID) &&
        Number(a.chantierId) === Number(g.chantierId) &&
        normalizeAB(a.statut) !== "ARCHIVE"
      );
    });


    // ----------------------------------------------
    // CAS 1 : période strictement identique
    // ----------------------------------------------

    const identique = candidates.find(function(a) {

      return (
        dateSheetISO(a.dateDebut) === g.dateDebut &&
        dateSheetISO(a.dateFin) === g.dateFin
      );
    });

    if (identique) {

      Logger.log(
        "✅ OK | " +
        g.ouvrierNom +
        " | " +
        g.chantierNom +
        " | " +
        g.dateDebut +
        " -> " +
        g.dateFin +
        " | tâche conservée = " +
        (identique.tache || "ND") +
        " | affectation ID=" +
        identique.id
      );

      nbOK++;
      return;
    }


    // ----------------------------------------------
    // CAS 2 : même ouvrier + même chantier,
    // mais dates différentes
    // ----------------------------------------------

    if (candidates.length > 0) {

      const existante = candidates[0];

      Logger.log(
        "🟠 À MODIFIER | " +
        g.ouvrierNom +
        " | " +
        g.chantierNom +
        " | Google=" +
        g.dateDebut +
        " -> " +
        g.dateFin +
        " | Gantt=" +
        dateSheetISO(existante.dateDebut) +
        " -> " +
        dateSheetISO(existante.dateFin) +
        " | tâche à CONSERVER = " +
        (existante.tache || "ND") +
        " | affectation ID=" +
        existante.id
      );

      nbModifier++;
      return;
    }


    // ----------------------------------------------
    // CAS 3 : aucune affectation correspondante
    // ----------------------------------------------

    Logger.log(
      "🆕 À CRÉER | " +
      g.ouvrierNom +
      " | " +
      g.chantierNom +
      " | " +
      g.dateDebut +
      " -> " +
      g.dateFin +
      " | tâche = ND"
    );

    nbCreer++;
  });


  // --------------------------------------------------
  // 7. RÉSUMÉ
  // --------------------------------------------------

  Logger.log("======================================");
  Logger.log("OK : " + nbOK);
  Logger.log("À MODIFIER : " + nbModifier);
  Logger.log("À CRÉER : " + nbCreer);
  Logger.log("======================================");
  Logger.log("TEST TERMINÉ - AUCUNE DONNÉE MODIFIÉE");
}
// ======================================================
// DRY RUN INITIALISATION GOOGLE <-> GANTT
// AUCUNE ECRITURE
// ======================================================

function testInitialisationBidirectionnelle() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();

  const dimanche = new Date("2026-08-30T00:00:00");
  const debut = new Date("2026-08-31T00:00:00");
  const fin = new Date("2026-09-09T00:00:00");

  // ===== 1. LIGNES OUVRIERS =====

  const eventsDimanche = calendar.getEvents(
    dimanche,
    debut
  );

  const lignes = [];

  eventsDimanche.forEach(function(event) {

    const titre = normalizeAB(event.getTitle());

    let ouvrier = ouvriers.find(function(o) {
      return normalizeAB(o.nom) === titre;
    });

    if (!ouvrier && titre === "ALEX") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "ALEXANDRE";
      });
    }

    if (!ouvrier && titre === "MOHAMED") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "MOMO";
      });
    }

    if (!ouvrier && titre === "UMAR") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "EQUIPE UMAR";
      });
    }

    if (ouvrier) {
      lignes.push({
        ouvrierID: Number(ouvrier.id),
        nom: ouvrier.nom,
        heure: event.getStartTime().getHours(),
        minute: event.getStartTime().getMinutes()
      });
    }
  });


  // ===== 2. JOURS PRESENTS DANS GOOGLE =====

  const googleJours = [];

  calendar.getEvents(debut, fin).forEach(function(event) {

    const date = event.getStartTime();

    if (date.getDay() === 0) return;

    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) ===
             normalizeAB(event.getTitle());
    });

    if (!chantier) return;

    const ligne = lignes.find(function(l) {
      return l.heure === date.getHours() &&
             l.minute === date.getMinutes();
    });

    if (!ligne) return;

    googleJours.push({
      ouvrierID: Number(ligne.ouvrierID),
      ouvrierNom: ligne.nom,
      chantierId: Number(chantier.id),
      chantierNom: chantier.nom,
      date: Utilities.formatDate(
        date,
        "Europe/Paris",
        "yyyy-MM-dd"
      )
    });
  });


  // ===== OUTIL DATE =====

  function iso(value) {

    if (!value) return "";

    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(
        value,
        "Europe/Paris",
        "yyyy-MM-dd"
      );
    }

    const texte = String(value).trim();

    const m = texte.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );

    if (m) {
      return m[1] + "-" + m[2] + "-" + m[3];
    }

    return "";
  }


  function chaqueJour(dateDebut, dateFin) {

    const resultat = [];

    let d = new Date(dateDebut + "T12:00:00");
    const f = new Date(dateFin + "T12:00:00");

    while (d <= f) {

      // Pas dimanche
      if (d.getDay() !== 0) {
        resultat.push(
          Utilities.formatDate(
            d,
            "Europe/Paris",
            "yyyy-MM-dd"
          )
        );
      }

      d.setDate(d.getDate() + 1);
    }

    return resultat;
  }


  // ===== 3. GANTT -> GOOGLE =====

  Logger.log(
    "===== À AJOUTER DANS GOOGLE DEPUIS GANTT ====="
  );

  let ajoutGoogle = 0;

  affectations.forEach(function(a) {

    const d1 = iso(a.dateDebut);
    const d2 = iso(a.dateFin);

    if (!d1 || !d2) return;

    // Initialisation seulement à partir du 01/08/2026
    if (d2 < "2026-08-01") return;

    const ouvrier = ouvriers.find(function(o) {
      return Number(o.id) === Number(a.ouvrierID);
    });

    const chantier = chantiers.find(function(c) {
      return Number(c.id) === Number(a.chantierId);
    });

    if (!ouvrier || !chantier) return;

    chaqueJour(d1, d2).forEach(function(jour) {

      // Seulement dans notre fenêtre de test
      if (jour < "2026-08-31" ||
          jour > "2026-09-08") return;

      const existeGoogle = googleJours.some(function(g) {
        return (
          g.ouvrierID === Number(a.ouvrierID) &&
          g.chantierId === Number(a.chantierId) &&
          g.date === jour
        );
      });

      if (!existeGoogle) {

        Logger.log(
          "➡ GOOGLE + | " +
          jour +
          " | " +
          ouvrier.nom +
          " | " +
          chantier.nom +
          " | tâche Gantt=" +
          (a.tache || "ND") +
          " | affectation ID=" +
          a.id
        );

        ajoutGoogle++;
      }
    });
  });


  // ===== 4. GOOGLE -> GANTT =====

  Logger.log(
    "===== À AJOUTER DANS GANTT DEPUIS GOOGLE ====="
  );

  let ajoutGantt = 0;

  googleJours.forEach(function(g) {

    const existeGantt = affectations.some(function(a) {

      if (
        Number(a.ouvrierID) !== g.ouvrierID ||
        Number(a.chantierId) !== g.chantierId
      ) {
        return false;
      }

      const d1 = iso(a.dateDebut);
      const d2 = iso(a.dateFin);

      if (!d1 || !d2) return false;

      return g.date >= d1 && g.date <= d2;
    });

    if (!existeGantt) {

      Logger.log(
        "⬅ GANTT + | " +
        g.date +
        " | " +
        g.ouvrierNom +
        " | " +
        g.chantierNom +
        " | tâche=ND"
      );

      ajoutGantt++;
    }
  });


  // ===== 5. RÉSUMÉ =====

  Logger.log("=====================================");
  Logger.log(
    "Jours à compléter dans Google : " +
    ajoutGoogle
  );
  Logger.log(
    "Jours à intégrer dans Gantt : " +
    ajoutGantt
  );
  Logger.log("=====================================");
  Logger.log(
    "DRY RUN TERMINÉ - AUCUNE ÉCRITURE"
  );
}
// ======================================================
// INITIALISATION REELLE GOOGLE <-> GANTT
// - complète Google depuis le Gantt
// - complète le Gantt depuis Google
// - tâche nouvelle = ND
// - aucune suppression
// - aucune modification des tâches existantes
// - plusieurs chantiers par ouvrier/jour autorisés
// ======================================================

function initialiserSynchronisationGoogleGantt() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Agenda Google introuvable : " + CALENDAR_ID);
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();

  const dimanche = new Date("2026-08-30T00:00:00");
  const debut = new Date("2026-08-31T00:00:00");
  const fin = new Date("2026-09-09T00:00:00");

  // ==================================================
  // OUTILS
  // ==================================================

  function iso(value) {

    if (!value) return "";

    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(
        value,
        "Europe/Paris",
        "yyyy-MM-dd"
      );
    }

    const texte = String(value).trim();

    let m = texte.match(/^(\d{4})-(\d{2})-(\d{2})/);

    if (m) {
      return m[1] + "-" + m[2] + "-" + m[3];
    }

    m = texte.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);

    if (m) {
      return m[3] + "-" + m[2] + "-" + m[1];
    }

    return "";
  }


  function chaqueJour(dateDebut, dateFin) {

    const jours = [];

    let d = new Date(dateDebut + "T12:00:00");
    const f = new Date(dateFin + "T12:00:00");

    while (d <= f) {

      // Dimanche exclu
      if (d.getDay() !== 0) {
        jours.push(
          Utilities.formatDate(
            d,
            "Europe/Paris",
            "yyyy-MM-dd"
          )
        );
      }

      d.setDate(d.getDate() + 1);
    }

    return jours;
  }


  function dateLocale(isoDate, heure, minute) {

    const morceaux = isoDate.split("-");

    return new Date(
      Number(morceaux[0]),
      Number(morceaux[1]) - 1,
      Number(morceaux[2]),
      heure,
      minute,
      0
    );
  }


  // ==================================================
  // 1. DETECTER LES LIGNES OUVRIERS
  // ==================================================

  const lignes = [];

  calendar.getEvents(dimanche, debut).forEach(function(event) {

    const titre = normalizeAB(event.getTitle());

    let ouvrier = ouvriers.find(function(o) {
      return normalizeAB(o.nom) === titre;
    });

    if (!ouvrier && titre === "ALEX") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "ALEXANDRE";
      });
    }

    if (!ouvrier && titre === "MOHAMED") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "MOMO";
      });
    }

    if (!ouvrier && titre === "UMAR") {
      ouvrier = ouvriers.find(function(o) {
        return normalizeAB(o.nom) === "EQUIPE UMAR";
      });
    }

    if (ouvrier) {

      lignes.push({
        ouvrierID: Number(ouvrier.id),
        nom: ouvrier.nom,
        heure: event.getStartTime().getHours(),
        minute: event.getStartTime().getMinutes()
      });
    }
  });


  Logger.log("Lignes ouvriers détectées : " + lignes.length);


  // ==================================================
  // 2. LIRE GOOGLE AVANT MODIFICATION
  // ==================================================

  const googleJours = [];

  calendar.getEvents(debut, fin).forEach(function(event) {

    const date = event.getStartTime();

    if (date.getDay() === 0) return;

    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) ===
             normalizeAB(event.getTitle());
    });

    if (!chantier) return;

    const ligne = lignes.find(function(l) {
      return (
        l.heure === date.getHours() &&
        l.minute === date.getMinutes()
      );
    });

    if (!ligne) return;

    googleJours.push({
      ouvrierID: ligne.ouvrierID,
      ouvrierNom: ligne.nom,
      chantierId: Number(chantier.id),
      chantierNom: chantier.nom,
      date: Utilities.formatDate(
        date,
        "Europe/Paris",
        "yyyy-MM-dd"
      ),
      eventId: event.getId()
    });
  });


  // ==================================================
  // 3. GANTT -> GOOGLE
  // ==================================================

  Logger.log("===== GANTT -> GOOGLE =====");

  let nbGoogle = 0;

  affectations.forEach(function(a) {

    const d1 = iso(a.dateDebut);
    const d2 = iso(a.dateFin);

    if (!d1 || !d2) return;

    if (d2 < "2026-08-01") return;

    const ouvrier = ouvriers.find(function(o) {
      return Number(o.id) === Number(a.ouvrierID);
    });

    const chantier = chantiers.find(function(c) {
      return Number(c.id) === Number(a.chantierId);
    });

    const ligne = lignes.find(function(l) {
      return Number(l.ouvrierID) === Number(a.ouvrierID);
    });

    if (!ouvrier || !chantier || !ligne) return;

    chaqueJour(d1, d2).forEach(function(jour) {

      // Fenêtre d'initialisation actuelle
      if (jour < "2026-08-31" ||
          jour > "2026-09-08") {
        return;
      }

      const existe = googleJours.some(function(g) {
        return (
          Number(g.ouvrierID) === Number(a.ouvrierID) &&
          Number(g.chantierId) === Number(a.chantierId) &&
          g.date === jour
        );
      });

      if (existe) return;


      // ----------------------------------------------
      // Créer l'événement à l'heure exacte de l'ouvrier
      // ----------------------------------------------

      const start = dateLocale(
        jour,
        ligne.heure,
        ligne.minute
      );

      const end = new Date(
        start.getTime() + (60 * 60 * 1000)
      );

      const event = calendar.createEvent(
        chantier.nom,
        start,
        end,
        {
          description:
            "AB PLANNING\n" +
            "OuvrierID: " + a.ouvrierID + "\n" +
            "ChantierID: " + a.chantierId + "\n" +
            "AffectationID: " + a.id + "\n" +
            "Tache: " + (a.tache || "ND")
        }
      );

      googleJours.push({
        ouvrierID: Number(a.ouvrierID),
        ouvrierNom: ouvrier.nom,
        chantierId: Number(a.chantierId),
        chantierNom: chantier.nom,
        date: jour,
        eventId: event.getId()
      });

      nbGoogle++;

      Logger.log(
        "✅ GOOGLE CRÉÉ | " +
        jour +
        " | " +
        ouvrier.nom +
        " | " +
        chantier.nom
      );
    });
  });


  // ==================================================
  // 4. GOOGLE -> GANTT
  // ==================================================

  Logger.log("===== GOOGLE -> GANTT =====");

  // Important :
  // on reprend googleJours, qui contient maintenant
  // également les événements ajoutés à l'étape précédente.

  const manquants = [];

  googleJours.forEach(function(g) {

    const existe = affectations.some(function(a) {

      if (
        Number(a.ouvrierID) !== Number(g.ouvrierID) ||
        Number(a.chantierId) !== Number(g.chantierId)
      ) {
        return false;
      }

      const d1 = iso(a.dateDebut);
      const d2 = iso(a.dateFin);

      if (!d1 || !d2) return false;

      return g.date >= d1 && g.date <= d2;
    });

    if (!existe) {
      manquants.push(g);
    }
  });


  // ==================================================
  // 5. CONSOLIDER LES JOURS MANQUANTS
  // ==================================================

  manquants.sort(function(a, b) {

    if (a.ouvrierID !== b.ouvrierID) {
      return a.ouvrierID - b.ouvrierID;
    }

    if (a.chantierId !== b.chantierId) {
      return a.chantierId - b.chantierId;
    }

    return a.date.localeCompare(b.date);
  });


  const groupes = [];

  manquants.forEach(function(item) {

    const dernier = groupes[groupes.length - 1];

    if (
      dernier &&
      dernier.ouvrierID === item.ouvrierID &&
      dernier.chantierId === item.chantierId
    ) {

      const d1 = new Date(
        dernier.dateFin + "T12:00:00"
      );

      const d2 = new Date(
        item.date + "T12:00:00"
      );

      const diff = Math.round(
        (d2.getTime() - d1.getTime()) / 86400000
      );

      if (diff === 1) {
        dernier.dateFin = item.date;
        dernier.eventIds.push(item.eventId);
        return;
      }
    }

    groupes.push({
      ouvrierID: item.ouvrierID,
      ouvrierNom: item.ouvrierNom,
      chantierId: item.chantierId,
      chantierNom: item.chantierNom,
      dateDebut: item.date,
      dateFin: item.date,
      eventIds: [item.eventId]
    });
  });


  // ==================================================
// 6. ÉCRIRE OU ÉTENDRE LES AFFECTATIONS
// ==================================================

const sheet = SpreadsheetApp
  .openById(SHEET_ID)
  .getSheetByName("affectations");

const data = sheet.getDataRange().getValues();

let maxId = 0;

for (let i = 1; i < data.length; i++) {
  const id = parseInt(data[i][0], 10);

  if (!isNaN(id)) {
    maxId = Math.max(maxId, id);
  }
}

let nbGanttCrees = 0;
let nbGanttEtendus = 0;

groupes.forEach(function(g) {

  // Chercher une affectation existante même ouvrier + chantier
  const candidates = affectations.filter(function(a) {
    return (
      Number(a.ouvrierID) === Number(g.ouvrierID) &&
      Number(a.chantierId) === Number(g.chantierId) &&
      normalizeAB(a.statut) !== "ARCHIVE"
    );
  });

  let affectationAEtendre = null;
  let nouvelleDateDebut = g.dateDebut;
  let nouvelleDateFin = g.dateFin;

  candidates.forEach(function(a) {

    if (affectationAEtendre) return;

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    const avant = new Date(g.dateFin + "T12:00:00");
    avant.setDate(avant.getDate() + 1);

    const apres = new Date(d2 + "T12:00:00");
    apres.setDate(apres.getDate() + 1);

    const jourApresNouveauBloc = Utilities.formatDate(
      avant,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    const jourApresExistant = Utilities.formatDate(
      apres,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    // Le nouveau bloc est juste avant l'existant
    if (jourApresNouveauBloc === d1) {
      affectationAEtendre = a;
      nouvelleDateDebut = g.dateDebut;
      nouvelleDateFin = d2;
      return;
    }

    // Le nouveau bloc est juste après l'existant
    if (jourApresExistant === g.dateDebut) {
      affectationAEtendre = a;
      nouvelleDateDebut = d1;
      nouvelleDateFin = g.dateFin;
      return;
    }
  });

  const maintenant = new Date();

  // ==================================================
  // CAS A : on étend une affectation existante
  // ==================================================

  if (affectationAEtendre) {

    const rowIndex = data.findIndex(function(row, index) {
      return (
        index > 0 &&
        Number(row[0]) === Number(affectationAEtendre.id)
      );
    });

    if (rowIndex !== -1) {

      const ligneSheet = rowIndex + 1;

      // D = dateDebut
      sheet
        .getRange(ligneSheet, 4)
        .setValue(nouvelleDateDebut);

      // E = dateFin
      sheet
        .getRange(ligneSheet, 5)
        .setValue(nouvelleDateFin);

      // H = googleEventId
      const anciensIds =
        String(affectationAEtendre.googleEventId || "")
          .split("|")
          .filter(Boolean);

      const nouveauxIds =
        g.eventIds
          .filter(Boolean);

      const idsFusionnes = Array.from(
        new Set(
          anciensIds.concat(nouveauxIds)
        )
      ).join("|");

      sheet
        .getRange(ligneSheet, 8)
        .setValue(idsFusionnes);

      // I = source
      if (!affectationAEtendre.source) {
        sheet
          .getRange(ligneSheet, 9)
          .setValue("SYNC");
      }

      // J et K
      sheet
        .getRange(ligneSheet, 10)
        .setValue(maintenant);

      sheet
        .getRange(ligneSheet, 11)
        .setValue(maintenant);

      // Mise à jour mémoire
      affectationAEtendre.dateDebut =
        nouvelleDateDebut;

      affectationAEtendre.dateFin =
        nouvelleDateFin;

      affectationAEtendre.googleEventId =
        idsFusionnes;

      nbGanttEtendus++;

      Logger.log(
        "🔄 GANTT ÉTENDU | " +
        g.ouvrierNom +
        " | " +
        g.chantierNom +
        " | " +
        nouvelleDateDebut +
        " -> " +
        nouvelleDateFin +
        " | tâche conservée=" +
        (affectationAEtendre.tache || "ND") +
        " | ID=" +
        affectationAEtendre.id
      );
    }

    return;
  }

  // ==================================================
  // CAS B : nouvelle affectation réellement distincte
  // ==================================================

  maxId++;

  const googleIds = g.eventIds
    .filter(Boolean)
    .join("|");

  sheet.appendRow([
    maxId,
    g.ouvrierID,
    g.chantierId,
    g.dateDebut,
    g.dateFin,
    "ND",
    "Actif",
    googleIds,
    "GOOGLE",
    maintenant,
    maintenant,
    "CHANTIER",
    "",
    ""
  ]);

  affectations.push({
    id: maxId,
    ouvrierID: g.ouvrierID,
    chantierId: g.chantierId,
    dateDebut: g.dateDebut,
    dateFin: g.dateFin,
    tache: "ND",
    statut: "Actif",
    googleEventId: googleIds,
    source: "GOOGLE",
    dateModification: maintenant,
    derniereSync: maintenant
  });

  nbGanttCrees++;

  Logger.log(
    "✅ GANTT + | " +
    g.ouvrierNom +
    " | " +
    g.chantierNom +
    " | " +
    g.dateDebut +
    " -> " +
    g.dateFin +
    " | ND"
  );
});

Logger.log("================================");

Logger.log(
  "Google créés : " +
  nbGoogleCrees
);

Logger.log(
  "Gantt créés : " +
  nbGanttCrees
);

Logger.log(
  "Gantt étendus : " +
  nbGanttEtendus
);

Logger.log("================================");

  // ==================================================
  // RÉSUMÉ
  // ==================================================

  Logger.log("====================================");
  Logger.log(
    "Événements ajoutés dans Google : " +
    nbGoogle
  );

  Logger.log(
    "Affectations ajoutées dans Gantt : " +
    nbGantt
  );

  Logger.log("====================================");
  Logger.log(
    "INITIALISATION TERMINÉE"
  );
}
// ======================================================
// SYNCHRONISATION PERMANENTE AB PLANNING <-> GOOGLE
// VERSION 1 : ADDITIVE ET SÉCURISÉE
// ======================================================

const AB_SYNC_START = "2026-08-01";
const AB_SYNC_FUTURE_DAYS = 120;


// ------------------------------------------------------
// DATE -> YYYY-MM-DD
// ------------------------------------------------------

function abIsoDate(value) {

  if (!value) return "";

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      "Europe/Paris",
      "yyyy-MM-dd"
    );
  }

  const texte = String(value).trim();

  let m = texte.match(/^(\d{4})-(\d{2})-(\d{2})/);

  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  m = texte.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);

  if (m) {
    return m[3] + "-" + m[2] + "-" + m[1];
  }

  return "";
}


// ------------------------------------------------------
// DIMANCHE PRÉCÉDANT UNE DATE
// ------------------------------------------------------

function abDimancheSemaine(date) {

  const d = new Date(date);

  d.setHours(0, 0, 0, 0);

  d.setDate(
    d.getDate() - d.getDay()
  );

  return d;
}


// ------------------------------------------------------
// ALIAS OUVRIERS GOOGLE -> SHEET
// ------------------------------------------------------

function abTrouverOuvrierParTitre(titre, ouvriers) {

  const normalise = normalizeAB(titre);

  let recherche = normalise;

  const alias = {
    "ALEX": "ALEXANDRE",
    "MOHAMED": "MOMO",
    "UMAR": "EQUIPE UMAR"
  };

  if (alias[normalise]) {
    recherche = alias[normalise];
  }

  return ouvriers.find(function(o) {
    return normalizeAB(o.nom) === recherche;
  }) || null;
}


// ------------------------------------------------------
// LIRE LES LIGNES OUVRIERS D'UNE SEMAINE
// ------------------------------------------------------

function abLignesOuvriersSemaine(
  calendar,
  dimanche,
  ouvriers
) {

  const lendemain = new Date(dimanche);

  lendemain.setDate(
    lendemain.getDate() + 1
  );

  const events = calendar.getEvents(
    dimanche,
    lendemain
  );

  const lignes = [];

  events.forEach(function(event) {

    const ouvrier = abTrouverOuvrierParTitre(
      event.getTitle(),
      ouvriers
    );

    if (!ouvrier) return;

    lignes.push({
      ouvrierID: Number(ouvrier.id),
      nom: ouvrier.nom,
      heure: event.getStartTime().getHours(),
      minute: event.getStartTime().getMinutes()
    });
  });

  return lignes;
}


// ------------------------------------------------------
// CRÉER UNE DATE À L'HEURE DE LA LIGNE OUVRIER
// ------------------------------------------------------

function abDateAvecHeure(
  dateISO,
  heure,
  minute
) {

  const p = dateISO.split("-");

  return new Date(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2]),
    heure,
    minute,
    0,
    0
  );
}


// ------------------------------------------------------
// TOUS LES JOURS D'UNE PÉRIODE
// DIMANCHE EXCLU
// ------------------------------------------------------

function abChaqueJour(
  dateDebut,
  dateFin
) {

  const resultat = [];

  let d = new Date(
    dateDebut + "T12:00:00"
  );

  const fin = new Date(
    dateFin + "T12:00:00"
  );

  while (d <= fin) {

    // dimanche = colonne repère ouvriers
    if (d.getDay() !== 0) {

      resultat.push(
        Utilities.formatDate(
          d,
          "Europe/Paris",
          "yyyy-MM-dd"
        )
      );
    }

    d.setDate(
      d.getDate() + 1
    );
  }

  return resultat;
}

// ======================================================
// SYNCHRONISATION ADDITIVE OPTIMISÉE
// - CHANTIER : affectation normale
// - HORS_GANTT : événement Google absent de la liste chantiers
// ======================================================

function syncGoogleGanttAdditiveOptimise(indexGoogle) {

  const debutTimer = new Date();

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Agenda introuvable : " + CALENDAR_ID);
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  let affectations = getAffectations();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const sheetData = sheet.getDataRange().getValues();

  const rowById = {};

  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] !== "") {
      rowById[Number(sheetData[i][0])] = i + 1;
    }
  }

  // ==================================================
  // FENÊTRE DE SYNCHRONISATION
  // ==================================================

  const aujourdHui = new Date();
  aujourdHui.setHours(0, 0, 0, 0);

  const debutLecture = new Date("2026-08-01T00:00:00");

  const finLecture = new Date(aujourdHui);
  finLecture.setDate(
    finLecture.getDate() + AB_SYNC_FUTURE_DAYS
  );

  const debutISO = Utilities.formatDate(
    debutLecture,
    "Europe/Paris",
    "yyyy-MM-dd"
  );

  const finISO = Utilities.formatDate(
    finLecture,
    "Europe/Paris",
    "yyyy-MM-dd"
  );

  // ==================================================
  // CACHE LIGNES OUVRIERS
  // ==================================================

  const cacheLignes = {};

  function lignesPourDate(date) {

    const dimanche = abDimancheSemaine(date);

    const cle = Utilities.formatDate(
      dimanche,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    if (!cacheLignes[cle]) {
      cacheLignes[cle] = abLignesOuvriersSemaine(
        calendar,
        dimanche,
        ouvriers
      );
    }

    return cacheLignes[cle];
  }

  // ==================================================
  // ÉVÉNEMENTS GOOGLE RECONNUS
  // ==================================================

  const googleJours = [];

  indexGoogle.actifs.forEach(function(event) {

    if (!event.start || !event.start.dateTime) {
      return;
    }

    const dateISO = String(event.start.dateTime).substring(0, 10);

    if (!dateISO) return;

    if (
      dateISO < debutISO ||
      dateISO > finISO
    ) {
      return;
    }

    const nomGoogle = String(event.summary || "").trim();

    if (!nomGoogle) return;

    const chantier = chantiers.find(function(c) {
      return (
        normalizeAB(c.nom) ===
        normalizeAB(nomGoogle)
      );
    });

    const horsGantt = !chantier;

    const dateEvent = new Date(event.start.dateTime);

    const heure = Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "HH"
      )
    );

    const minute = Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "mm"
      )
    );

    const lignes = lignesPourDate(
      new Date(dateISO + "T12:00:00")
    );

    // ==================================================
    // RATTACHEMENT STRICT À LA PLAGE RÉELLE DE L'OUVRIER
    //
    // Le repère Google de l'ouvrier définit le début de
    // sa plage. La plage est strictement limitée à 60 min.
    //
    // Kevin 00:00 -> 01:00 :
    // - Mvondo 00:00 -> 00:30    => Kevin
    // - Firminski 00:30 -> 01:00 => Kevin
    // - événement à 01:00        => PAS Kevin
    //
    // On ne prolonge jamais la plage jusqu'au repère
    // de l'ouvrier suivant.
    // ==================================================

    const minuteEvent =
      (heure * 60) + minute;

    let ligne = null;

    for (let i = 0; i < lignes.length; i++) {

      const debutLigne =
        (Number(lignes[i].heure) * 60) +
        Number(lignes[i].minute);

      const nomOuvrier =
        normalizeAB(
          lignes[i].nom || ""
        );

      // Durée réelle de la plage Google par ouvrier.
      // Tous les ouvriers = 60 min.
      // Morvan = 105 min (1 h 45).
      const dureePlage =
        nomOuvrier === "MORVAN"
          ? 105
          : 60;

      const finLigne =
        debutLigne +
        dureePlage;

      if (
        minuteEvent >= debutLigne &&
        minuteEvent < finLigne
      ) {
        ligne = lignes[i];
        break;
      }
    }

    if (!ligne) return;

    googleJours.push({
      ouvrierID: Number(ligne.ouvrierID),
      ouvrierNom: ligne.nom,

      chantierId: chantier ? Number(chantier.id) : null,
      chantierNom: chantier ? chantier.nom : nomGoogle,

      typeAffectation: horsGantt ? "HORS_GANTT" : "CHANTIER",
      nomExterne: horsGantt ? nomGoogle : "",

      date: dateISO,

      eventId:
        event.iCalUID ||
        event.id
    });
  });

  Logger.log(
    "⚡ Jours Google reconnus : " +
    googleJours.length
  );

  // ==================================================
  // 1. GANTT -> GOOGLE
  // ==================================================

  let nbGoogleCrees = 0;

  affectations.forEach(function(a) {

    if (
      normalizeAB(a.statut) ===
      "ARCHIVE"
    ) {
      return;
    }

    // Les HORS_GANTT naissent uniquement depuis Google.
    // Leur suppression Gantt -> Google est gérée par deleteAffectation().
    if (
      normalizeAB(a.typeAffectation) ===
      "HORS_GANTT"
    ) {
      return;
    }

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    if (
      d2 < debutISO ||
      d1 > finISO
    ) {
      return;
    }

    const ouvrier = ouvriers.find(function(o) {
      return (
        Number(o.id) ===
        Number(a.ouvrierID)
      );
    });

    const chantier = chantiers.find(function(c) {
      return (
        Number(c.id) ===
        Number(a.chantierId)
      );
    });

    if (!ouvrier || !chantier) return;

    abChaqueJour(d1, d2).forEach(function(jour) {

      if (
        jour < debutISO ||
        jour > finISO
      ) {
        return;
      }

      const existeGoogle = googleJours.some(function(g) {
        return (
          g.typeAffectation === "CHANTIER" &&
          g.ouvrierID === Number(a.ouvrierID) &&
          g.chantierId === Number(a.chantierId) &&
          g.date === jour
        );
      });

      if (existeGoogle) return;

      const lignes = lignesPourDate(
        new Date(jour + "T12:00:00")
      );

      const ligne = lignes.find(function(l) {
        return (
          Number(l.ouvrierID) ===
          Number(a.ouvrierID)
        );
      });

      if (!ligne) {

        Logger.log(
          "⚠ Ligne ouvrier absente | " +
          jour +
          " | " +
          ouvrier.nom
        );

        return;
      }

      const debutEvent = abDateAvecHeure(
        jour,
        ligne.heure,
        ligne.minute
      );

      const finEvent = new Date(
        debutEvent.getTime() +
        60 * 60 * 1000
      );

      const event = calendar.createEvent(
        chantier.nom,
        debutEvent,
        finEvent,
        {
          description:
            "AB PLANNING\n" +
            "OuvrierID: " + a.ouvrierID + "\n" +
            "ChantierID: " + a.chantierId + "\n" +
            "AffectationID: " + a.id + "\n" +
            "Tache: " + (a.tache || "ND")
        }
      );

      const eventId = event.getId();

      googleJours.push({
        ouvrierID: Number(a.ouvrierID),
        ouvrierNom: ouvrier.nom,
        chantierId: Number(a.chantierId),
        chantierNom: chantier.nom,
        typeAffectation: "CHANTIER",
        nomExterne: "",
        date: jour,
        eventId: eventId
      });

      const ligneSheet =
        rowById[Number(a.id)];

      if (ligneSheet) {

        const anciensIds = String(
          a.googleEventId || ""
        )
          .split("|")
          .filter(Boolean);

        if (!anciensIds.includes(eventId)) {
          anciensIds.push(eventId);
        }

        const idsFusionnes =
          Array.from(
            new Set(anciensIds)
          ).join("|");

        sheet
          .getRange(ligneSheet, 8)
          .setValue(idsFusionnes);

        if (!a.source) {
          sheet
            .getRange(ligneSheet, 9)
            .setValue("GANTT");
        }

        sheet
          .getRange(ligneSheet, 11)
          .setValue(new Date());

        if (!a.typeAffectation) {
          sheet
            .getRange(ligneSheet, 12)
            .setValue("CHANTIER");
        }

        a.googleEventId = idsFusionnes;
      }

      nbGoogleCrees++;

      Logger.log(
        "✅ GOOGLE + | " +
        jour +
        " | " +
        ouvrier.nom +
        " | " +
        chantier.nom
      );
    });
  });

  // ==================================================
  // 2. GOOGLE -> GANTT
  // ==================================================

  const manquants = [];

  googleJours.forEach(function(g) {

    const existeGantt =
      affectations.some(function(a) {

        if (
          Number(a.ouvrierID) !==
          Number(g.ouvrierID)
        ) {
          return false;
        }

        const d1 = abIsoDate(a.dateDebut);
        const d2 = abIsoDate(a.dateFin);

        if (!d1 || !d2) return false;

        if (
          g.date < d1 ||
          g.date > d2
        ) {
          return false;
        }

        if (
          g.typeAffectation ===
          "CHANTIER"
        ) {
          return (
            normalizeAB(a.typeAffectation || "CHANTIER") !== "HORS_GANTT" &&
            Number(a.chantierId) ===
            Number(g.chantierId)
          );
        }

        return (
          normalizeAB(a.typeAffectation) ===
            "HORS_GANTT" &&
          normalizeAB(a.nomExterne || "") ===
            normalizeAB(g.nomExterne || "")
        );
      });

    if (!existeGantt) {
      manquants.push(g);
    }
  });

  // ==================================================
  // 3. TRIER ET CONSOLIDER
  // ==================================================

  manquants.sort(function(a, b) {

    if (
      a.ouvrierID !==
      b.ouvrierID
    ) {
      return (
        a.ouvrierID -
        b.ouvrierID
      );
    }

    const cleA =
      a.typeAffectation +
      "|" +
      (
        a.typeAffectation === "CHANTIER"
          ? String(a.chantierId)
          : normalizeAB(a.nomExterne || "")
      );

    const cleB =
      b.typeAffectation +
      "|" +
      (
        b.typeAffectation === "CHANTIER"
          ? String(b.chantierId)
          : normalizeAB(b.nomExterne || "")
      );

    if (cleA !== cleB) {
      return cleA.localeCompare(cleB);
    }

    return a.date.localeCompare(b.date);
  });

  const groupes = [];

  manquants.forEach(function(item) {

    const dernier =
      groupes[groupes.length - 1];

    const memeAffectation =
      dernier &&
      dernier.ouvrierID === item.ouvrierID &&
      dernier.typeAffectation === item.typeAffectation &&
      (
        item.typeAffectation === "CHANTIER"
          ? Number(dernier.chantierId) === Number(item.chantierId)
          : normalizeAB(dernier.nomExterne || "") === normalizeAB(item.nomExterne || "")
      );

    if (memeAffectation) {

      const dateDerniere =
        new Date(
          dernier.dateFin +
          "T12:00:00"
        );

      const dateCourante =
        new Date(
          item.date +
          "T12:00:00"
        );

      const diff =
        Math.round(
          (
            dateCourante -
            dateDerniere
          ) /
          86400000
        );

      if (diff === 1) {

        dernier.dateFin =
          item.date;

        dernier.eventIds.push(
          item.eventId
        );

        return;
      }
    }

    groupes.push({
      ouvrierID: item.ouvrierID,
      ouvrierNom: item.ouvrierNom,

      chantierId: item.chantierId,
      chantierNom: item.chantierNom,

      typeAffectation: item.typeAffectation,
      nomExterne: item.nomExterne || "",

      dateDebut: item.date,
      dateFin: item.date,

      eventIds: [item.eventId]
    });
  });

  // ==================================================
  // 4. ÉTENDRE OU CRÉER
  // ==================================================

  let nbGanttCrees = 0;
  let nbGanttEtendus = 0;

  groupes.forEach(function(g) {

    const candidates =
      affectations.filter(function(a) {

        if (
          Number(a.ouvrierID) !==
          Number(g.ouvrierID)
        ) {
          return false;
        }

        if (
          normalizeAB(a.statut) ===
          "ARCHIVE"
        ) {
          return false;
        }

        if (
          g.typeAffectation ===
          "CHANTIER"
        ) {
          return (
            normalizeAB(a.typeAffectation || "CHANTIER") !== "HORS_GANTT" &&
            Number(a.chantierId) ===
            Number(g.chantierId)
          );
        }

        return (
          normalizeAB(a.typeAffectation) ===
            "HORS_GANTT" &&
          normalizeAB(a.nomExterne || "") ===
            normalizeAB(g.nomExterne || "")
        );
      });

    let aEtendre = null;

    let nouvelleDateDebut =
      g.dateDebut;

    let nouvelleDateFin =
      g.dateFin;

    candidates.forEach(function(a) {

      if (aEtendre) return;

      const d1 =
        abIsoDate(a.dateDebut);

      const d2 =
        abIsoDate(a.dateFin);

      if (!d1 || !d2) return;

      const apresExistant =
        new Date(
          d2 + "T12:00:00"
        );

      apresExistant.setDate(
        apresExistant.getDate() + 1
      );

      const lendemainExistant =
        Utilities.formatDate(
          apresExistant,
          "Europe/Paris",
          "yyyy-MM-dd"
        );

      if (
        lendemainExistant ===
        g.dateDebut
      ) {
        aEtendre = a;
        nouvelleDateDebut = d1;
        nouvelleDateFin = g.dateFin;
        return;
      }

      const apresNouveau =
        new Date(
          g.dateFin +
          "T12:00:00"
        );

      apresNouveau.setDate(
        apresNouveau.getDate() + 1
      );

      const lendemainNouveau =
        Utilities.formatDate(
          apresNouveau,
          "Europe/Paris",
          "yyyy-MM-dd"
        );

      if (
        lendemainNouveau === d1
      ) {
        aEtendre = a;
        nouvelleDateDebut = g.dateDebut;
        nouvelleDateFin = d2;
      }
    });

    const maintenant =
      new Date();

    // ----------------------------------------------
    // ÉTENDRE
    // ----------------------------------------------

    if (aEtendre) {

      const ligneSheet =
        rowById[
          Number(aEtendre.id)
        ];

      if (ligneSheet) {

        sheet
          .getRange(ligneSheet, 4)
          .setValue(nouvelleDateDebut);

        sheet
          .getRange(ligneSheet, 5)
          .setValue(nouvelleDateFin);

        const anciensIds =
          String(
            aEtendre.googleEventId ||
            ""
          )
            .split("|")
            .filter(Boolean);

        const idsFusionnes =
          Array.from(
            new Set(
              anciensIds.concat(
                g.eventIds.filter(Boolean)
              )
            )
          ).join("|");

        sheet
          .getRange(ligneSheet, 8)
          .setValue(idsFusionnes);

        if (!aEtendre.source) {
          sheet
            .getRange(ligneSheet, 9)
            .setValue("SYNC");
        }

        sheet
          .getRange(ligneSheet, 10)
          .setValue(maintenant);

        sheet
          .getRange(ligneSheet, 11)
          .setValue(maintenant);

        sheet
          .getRange(ligneSheet, 12)
          .setValue(g.typeAffectation);

        sheet
          .getRange(ligneSheet, 13)
          .setValue(g.nomExterne || "");

        aEtendre.dateDebut = nouvelleDateDebut;
        aEtendre.dateFin = nouvelleDateFin;
        aEtendre.googleEventId = idsFusionnes;
        aEtendre.typeAffectation = g.typeAffectation;
        aEtendre.nomExterne = g.nomExterne || "";

        nbGanttEtendus++;

        Logger.log(
          "🔄 GANTT ÉTENDU | " +
          g.ouvrierNom +
          " | " +
          (
            g.typeAffectation === "HORS_GANTT"
              ? g.nomExterne
              : g.chantierNom
          ) +
          " | " +
          nouvelleDateDebut +
          " -> " +
          nouvelleDateFin
        );
      }

      return;
    }

    // ----------------------------------------------
    // CRÉER
    // ----------------------------------------------

    const newId =
      getNextAffectationId();

    const googleIds =
      Array.from(
        new Set(
          g.eventIds.filter(Boolean)
        )
      ).join("|");

    sheet.appendRow([
      newId,
      g.ouvrierID,
      g.typeAffectation === "CHANTIER"
        ? g.chantierId
        : "",
      g.dateDebut,
      g.dateFin,
      g.typeAffectation === "HORS_GANTT" ? "" : "ND",
      "Actif",
      googleIds,
      "GOOGLE",
      maintenant,
      maintenant,
      g.typeAffectation,
      g.nomExterne || "",
      ""
    ]);

    affectations.push({
      id: newId,

      ouvrierID:
        g.ouvrierID,

      chantierId:
        g.typeAffectation === "CHANTIER"
          ? g.chantierId
          : "",

      dateDebut:
        g.dateDebut,

      dateFin:
        g.dateFin,

      tache:
        g.typeAffectation === "HORS_GANTT" ? "" : "ND",

      statut:
        "Actif",

      googleEventId:
        googleIds,

      source:
        "GOOGLE",

      dateModification:
        maintenant,

      derniereSync:
        maintenant,

      typeAffectation:
        g.typeAffectation,

      nomExterne:
        g.nomExterne || ""
    });

    nbGanttCrees++;

    Logger.log(
      "✅ GANTT + | ID=" +
      newId +
      " | " +
      g.ouvrierNom +
      " | " +
      (
        g.typeAffectation === "HORS_GANTT"
          ? g.nomExterne
          : g.chantierNom
      ) +
      " | " +
      g.dateDebut +
      " -> " +
      g.dateFin
    );
  });

  Logger.log(
    "⚡ ADDITIF OPTIMISÉ | Google créés=" +
    nbGoogleCrees +
    " | Gantt créés=" +
    nbGanttCrees +
    " | Gantt étendus=" +
    nbGanttEtendus +
    " | durée=" +
    ((new Date() - debutTimer) / 1000) +
    "s"
  );
}


function rattacherTousLesGoogleEventIds() {


  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Agenda introuvable : " + CALENDAR_ID);
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();
  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const debut = new Date("2026-08-01T00:00:00");

  const fin = new Date();
  fin.setDate(fin.getDate() + AB_SYNC_FUTURE_DAYS);

  // Cache des lignes ouvriers par semaine
  const cacheLignes = {};

  function lignesPourDate(date) {

    const dimanche = abDimancheSemaine(date);

    const cle = Utilities.formatDate(
      dimanche,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    if (!cacheLignes[cle]) {
      cacheLignes[cle] = abLignesOuvriersSemaine(
        calendar,
        dimanche,
        ouvriers
      );
    }

    return cacheLignes[cle];
  }

  // Tous les événements Google reconnus
  const googleJours = [];

  calendar.getEvents(debut, fin).forEach(function(event) {

    const date = event.getStartTime();

    // dimanche = repères ouvriers
    if (date.getDay() === 0) return;

    const chantier = chantiers.find(function(c) {
      return normalizeAB(c.nom) === normalizeAB(event.getTitle());
    });

    if (!chantier) return;

    const lignes = lignesPourDate(date);

    const ligne = lignes.find(function(l) {
      return (
        l.heure === date.getHours() &&
        l.minute === date.getMinutes()
      );
    });

    if (!ligne) return;

    googleJours.push({
      ouvrierID: Number(ligne.ouvrierID),
      chantierId: Number(chantier.id),

      date: Utilities.formatDate(
        date,
        "Europe/Paris",
        "yyyy-MM-dd"
      ),

      eventId: event.getId()
    });
  });

  let nbLignes = 0;
  let nbIds = 0;

  affectations.forEach(function(a) {

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;
    if (d2 < AB_SYNC_START) return;

    const ids = googleJours
      .filter(function(g) {

        return (
          g.ouvrierID === Number(a.ouvrierID) &&
          g.chantierId === Number(a.chantierId) &&
          g.date >= d1 &&
          g.date <= d2
        );
      })
      .map(function(g) {
        return g.eventId;
      });

    const idsUniques = Array.from(
      new Set(ids.filter(Boolean))
    );

    if (idsUniques.length === 0) return;

    // retrouver la ligne par son ID affectation
    const data = sheet.getDataRange().getValues();

    const rowIndex = data.findIndex(function(row, index) {
      return (
        index > 0 &&
        Number(row[0]) === Number(a.id)
      );
    });

    if (rowIndex === -1) return;

    const ligneSheet = rowIndex + 1;

    // H
    sheet
      .getRange(ligneSheet, 8)
      .setValue(idsUniques.join("|"));

    // K = dernière synchro
    sheet
      .getRange(ligneSheet, 11)
      .setValue(new Date());

    nbLignes++;
    nbIds += idsUniques.length;

    Logger.log(
      "🔗 ID GOOGLE | affectation " +
      a.id +
      " | " +
      idsUniques.length +
      " événement(s)"
    );
  });

  Logger.log("================================");
  Logger.log("Affectations rattachées : " + nbLignes);
  Logger.log("IDs Google enregistrés : " + nbIds);
  Logger.log("================================");
}
function testDeplacementsGoogleVersGantt() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Agenda introuvable : " + CALENDAR_ID);
  }

  const affectations = getAffectations();

  Logger.log("===== TEST DÉPLACEMENTS GOOGLE -> GANTT =====");

  let nbDeplacements = 0;
  let nbIntrouvables = 0;

  affectations.forEach(function(a) {

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    const ids = String(a.googleEventId || "")
      .split("|")
      .filter(Boolean);

    if (ids.length === 0) return;

    ids.forEach(function(eventId) {

      let event = null;

      try {
        event = calendar.getEventById(eventId);
      } catch (err) {
        event = null;
      }

      if (!event) {
        nbIntrouvables++;

        Logger.log(
          "⚠ EVENT INTROUVABLE | affectation ID=" +
          a.id +
          " | eventId=" +
          eventId
        );

        return;
      }

      const nouvelleDate = Utilities.formatDate(
        event.getStartTime(),
        "Europe/Paris",
        "yyyy-MM-dd"
      );

      // si la nouvelle date n'est plus comprise
      // dans la période actuelle du Gantt
      if (
        nouvelleDate < d1 ||
        nouvelleDate > d2
      ) {

        nbDeplacements++;

        Logger.log(
          "🔄 DÉPLACEMENT DÉTECTÉ | affectation ID=" +
          a.id +
          " | Gantt=" +
          d1 +
          " -> " +
          d2 +
          " | Google=" +
          nouvelleDate +
          " | titre=" +
          event.getTitle()
        );
      }
    });
  });

  Logger.log("====================================");
  Logger.log("Déplacements détectés : " + nbDeplacements);
  Logger.log("Événements introuvables : " + nbIntrouvables);
  Logger.log("====================================");
  Logger.log("TEST TERMINÉ - AUCUNE MODIFICATION");
}
function appliquerDeplacementsGoogleVersGantt() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Agenda introuvable : " + CALENDAR_ID);
  }

  const affectations = getAffectations();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  Logger.log("===== APPLICATION DÉPLACEMENTS GOOGLE -> GANTT =====");

  let nbModifies = 0;
  let nbIgnores = 0;

  affectations.forEach(function(a) {

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    const ids = String(a.googleEventId || "")
      .split("|")
      .filter(Boolean);

    if (ids.length === 0) return;

    // SÉCURITÉ :
    // pour l'instant uniquement les affectations
    // d'une seule journée avec un seul événement Google
    if (d1 !== d2 || ids.length !== 1) {
      return;
    }

    let event = null;

    try {
      event = calendar.getEventById(ids[0]);
    } catch (err) {
      event = null;
    }

    // Une suppression sera traitée séparément.
    if (!event) {
      nbIgnores++;
      return;
    }

    const nouvelleDate = Utilities.formatDate(
      event.getStartTime(),
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    if (nouvelleDate === d1) {
      return;
    }

    const rowIndex = data.findIndex(function(row, index) {
      return (
        index > 0 &&
        Number(row[0]) === Number(a.id)
      );
    });

    if (rowIndex === -1) {
      nbIgnores++;
      return;
    }

    const ligneSheet = rowIndex + 1;
    const maintenant = new Date();

    // D = dateDebut
    sheet
      .getRange(ligneSheet, 4)
      .setValue(nouvelleDate);

    // E = dateFin
    sheet
      .getRange(ligneSheet, 5)
      .setValue(nouvelleDate);

    // J = dateModification
    sheet
      .getRange(ligneSheet, 10)
      .setValue(maintenant);

    // K = derniereSync
    sheet
      .getRange(ligneSheet, 11)
      .setValue(maintenant);

    nbModifies++;

    Logger.log(
      "✅ GANTT DÉPLACÉ | ID=" +
      a.id +
      " | " +
      d1 +
      " -> " +
      nouvelleDate +
      " | " +
      event.getTitle()
    );
  });

  Logger.log("====================================");
  Logger.log("Affectations déplacées : " + nbModifies);
  Logger.log("Cas ignorés : " + nbIgnores);
  Logger.log("====================================");
}

function testSuppressionsGoogleVersGantt() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const affectations = getAffectations();

  Logger.log("===== TEST SUPPRESSIONS GOOGLE -> GANTT =====");

  let suppressionsDetectees = 0;
  let casNonFiables = 0;

  affectations.forEach(function(a) {

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    const ids = String(a.googleEventId || "")
      .split("|")
      .filter(Boolean);

    if (ids.length === 0) return;

    const jours = abChaqueJour(d1, d2);

    // On ne touche qu'aux affectations où chaque jour
    // possède bien un ID Google identifiable.
    if (ids.length !== jours.length) {

      Logger.log(
        "⚠ NON FIABLE | ID=" + a.id +
        " | jours=" + jours.length +
        " | IDs=" + ids.length
      );

      casNonFiables++;
      return;
    }

    const joursSupprimes = [];

    ids.forEach(function(eventId, index) {

      let event = null;

      try {
        event = calendar.getEventById(eventId);
      } catch (err) {
        event = null;
      }

      if (!event) {
        joursSupprimes.push(jours[index]);
      }
    });

    if (joursSupprimes.length === 0) return;

    suppressionsDetectees++;

    Logger.log(
      "🗑 SUPPRESSION DÉTECTÉE | affectation ID=" +
      a.id +
      " | Gantt=" +
      d1 +
      " -> " +
      d2 +
      " | jours supprimés Google=" +
      joursSupprimes.join(", ")
    );

    if (joursSupprimes.length === jours.length) {

      Logger.log(
        "➡ ACTION FUTURE : SUPPRIMER TOUTE L'AFFECTATION GANTT ID=" +
        a.id
      );
    }
  });

  Logger.log("====================================");
  Logger.log(
    "Affectations avec suppression détectée : " +
    suppressionsDetectees
  );
  Logger.log(
    "Cas non fiables ignorés : " +
    casNonFiables
  );
  Logger.log("====================================");
  Logger.log("TEST TERMINÉ - AUCUNE MODIFICATION");
}
function verifierGoogleIdsAffectation161() {

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  const ids = [
    "0kqpmbilspomov3erp9aef47n8@google.com",
    "5e44ca9kh1if19jqjmro6u9h4n@google.com",
    "0m84pd8hk6rsp7vcn7a6k7b7r0@google.com",
    "lq4rfvspjaepl5u6eql137stqs@google.com"
  ];

  Logger.log("===== VÉRIFICATION ID 161 =====");

  ids.forEach(function(id) {

    let event = null;

    try {
      event = calendar.getEventById(id);
    } catch (e) {
      event = null;
    }

    if (!event) {
      Logger.log("❌ SUPPRIMÉ / INTROUVABLE | " + id);
      return;
    }

    Logger.log(
      "✅ EXISTE | " +
      id +
      " | " +
      Utilities.formatDate(
        event.getStartTime(),
        "Europe/Paris",
        "yyyy-MM-dd HH:mm"
      ) +
      " | " +
      event.getTitle()
    );
  });

  Logger.log("===== FIN =====");
}
function verifierStatutGoogleIdsAffectation161() {

  const ids = [
    "0kqpmbilspomov3erp9aef47n8",
    "5e44ca9kh1if19jqjmro6u9h4n",
    "0m84pd8hk6rsp7vcn7a6k7b7r0",
    "lq4rfvspjaepl5u6eql137stqs"
  ];

  Logger.log("===== STATUT API GOOGLE - AFFECTATION 161 =====");

  ids.forEach(function(id) {

    try {

      const event = Calendar.Events.get(
        CALENDAR_ID,
        id
      );

      Logger.log(
        "ID=" + id +
        " | status=" + event.status +
        " | titre=" + (event.summary || "-")
      );

    } catch (err) {

      Logger.log(
        "ERREUR / INTROUVABLE | " +
        id +
        " | " +
        err.message
      );
    }
  });

  Logger.log("===== FIN =====");
}
function appliquerSuppressionsGoogleCompletes() {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  Logger.log("===== SUPPRESSIONS GOOGLE -> GANTT =====");

  let nbSupprimees = 0;
  let nbIgnorees = 0;

  // On part du bas pour pouvoir supprimer des lignes
  // sans décaler les lignes suivantes.
  for (let i = data.length - 1; i >= 1; i--) {

    const idAffectation = data[i][0];

    // H = googleEventId
    const ids = String(data[i][7] || "")
      .split("|")
      .map(function(id) {
        return id.trim();
      })
      .filter(Boolean);

    // Pas d'ID Google = on ne touche pas.
    if (ids.length === 0) {
      continue;
    }

    let nbCancelled = 0;
    let nbActifs = 0;
    let nbInconnus = 0;

    ids.forEach(function(idComplet) {

      // Calendar API avancée attend l'ID sans @google.com
      const eventId = idComplet.replace(
        /@google\.com$/i,
        ""
      );

      try {

        const event = Calendar.Events.get(
          CALENDAR_ID,
          eventId
        );

        if (
          event &&
          event.status === "cancelled"
        ) {
          nbCancelled++;
        } else {
          nbActifs++;
        }

      } catch (err) {

        // Une erreur n'est PAS considérée automatiquement
        // comme une suppression.
        nbInconnus++;

        Logger.log(
          "⚠ ID NON VÉRIFIÉ | affectation=" +
          idAffectation +
          " | " +
          eventId +
          " | " +
          err.message
        );
      }
    });

    // ==================================================
    // TOUS LES IDs SONT CONFIRMÉS CANCELLED
    // ==================================================

    if (
      nbCancelled === ids.length &&
      nbActifs === 0 &&
      nbInconnus === 0
    ) {

      Logger.log(
        "🗑 GANTT SUPPRIMÉ | affectation ID=" +
        idAffectation +
        " | tous les événements Google sont cancelled"
      );

      sheet.deleteRow(i + 1);

      nbSupprimees++;

      continue;
    }

    // ==================================================
    // MÉLANGE ACTIF / CANCELLED
    // ==================================================

    if (nbCancelled > 0) {

      Logger.log(
        "⚠ SUPPRESSION PARTIELLE NON TRAITÉE | ID=" +
        idAffectation +
        " | cancelled=" +
        nbCancelled +
        " | actifs=" +
        nbActifs +
        " | inconnus=" +
        nbInconnus
      );

      nbIgnorees++;
    }
  }

  Logger.log("====================================");
  Logger.log(
    "Affectations supprimées : " +
    nbSupprimees
  );
  Logger.log(
    "Suppressions partielles ignorées : " +
    nbIgnorees
  );
  Logger.log("====================================");
}
function appliquerSuppressionsPartiellesGoogleVersGantt() {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0], 10);
    if (!isNaN(id)) {
      maxId = Math.max(maxId, id);
    }
  }

  Logger.log("===== SUPPRESSIONS PARTIELLES GOOGLE -> GANTT =====");

  let nbReduites = 0;
  let nbScindees = 0;

  // du bas vers le haut pour éviter les décalages
  for (let i = data.length - 1; i >= 1; i--) {

    const idAffectation = data[i][0];
    const ouvrierID = data[i][1];
    const chantierId = data[i][2];

    const dateDebut = abIsoDate(data[i][3]);
    const dateFin = abIsoDate(data[i][4]);

    const tache = data[i][5] || "ND";
    const statut = data[i][6] || "Actif";

    const ids = String(data[i][7] || "")
      .split("|")
      .map(id => id.trim())
      .filter(Boolean);

    if (!dateDebut || !dateFin || ids.length === 0) {
      continue;
    }

    const joursActifs = [];

    ids.forEach(function(idComplet) {

      const eventId = idComplet.replace(
        /@google\.com$/i,
        ""
      );

      try {

        const event = Calendar.Events.get(
          CALENDAR_ID,
          eventId
        );

        if (
          event &&
          event.status !== "cancelled" &&
          event.start
        ) {

          const dateEvent =
            event.start.date ||
            String(event.start.dateTime || "").substring(0, 10);

          if (dateEvent) {
            joursActifs.push({
              date: dateEvent,
              id: idComplet
            });
          }
        }

      } catch (err) {
        Logger.log(
          "⚠ ID non vérifié | affectation=" +
          idAffectation +
          " | " +
          eventId
        );
      }
    });

    // Aucun actif :
    // traité par appliquerSuppressionsGoogleCompletes()
    if (joursActifs.length === 0) {
      continue;
    }

    // Dates uniques et triées
    const actifsParDate = {};

    joursActifs.forEach(function(j) {
      actifsParDate[j.date] = j.id;
    });

    const datesActives =
      Object.keys(actifsParDate).sort();

    const tousLesJours =
      abChaqueJour(dateDebut, dateFin);

    // Aucun jour supprimé
    if (datesActives.length === tousLesJours.length) {
      continue;
    }

    // Construire les blocs consécutifs encore actifs
    const blocs = [];

    datesActives.forEach(function(date) {

      const dernier =
        blocs[blocs.length - 1];

      if (dernier) {

        const d1 =
          new Date(
            dernier.dateFin +
            "T12:00:00"
          );

        const d2 =
          new Date(
            date +
            "T12:00:00"
          );

        const diff =
          Math.round(
            (d2 - d1) / 86400000
          );

        if (diff === 1) {

          dernier.dateFin = date;
          dernier.ids.push(
            actifsParDate[date]
          );

          return;
        }
      }

      blocs.push({
        dateDebut: date,
        dateFin: date,
        ids: [actifsParDate[date]]
      });
    });

    const maintenant = new Date();

    // ==================================================
    // UN SEUL BLOC RESTANT :
    // on réduit simplement l'affectation
    // ==================================================

    if (blocs.length === 1) {

      const bloc = blocs[0];

      sheet
        .getRange(i + 1, 4)
        .setValue(bloc.dateDebut);

      sheet
        .getRange(i + 1, 5)
        .setValue(bloc.dateFin);

      sheet
        .getRange(i + 1, 8)
        .setValue(
          bloc.ids.join("|")
        );

      sheet
        .getRange(i + 1, 10)
        .setValue(maintenant);

      sheet
        .getRange(i + 1, 11)
        .setValue(maintenant);

      nbReduites++;

      Logger.log(
        "✂ GANTT RÉDUIT | ID=" +
        idAffectation +
        " | " +
        bloc.dateDebut +
        " -> " +
        bloc.dateFin
      );

      continue;
    }

    // ==================================================
    // PLUSIEURS BLOCS :
    // on conserve le premier sur la ligne existante
    // et on crée les autres
    // ==================================================

    const premier = blocs[0];

    sheet
      .getRange(i + 1, 4)
      .setValue(premier.dateDebut);

    sheet
      .getRange(i + 1, 5)
      .setValue(premier.dateFin);

    sheet
      .getRange(i + 1, 8)
      .setValue(
        premier.ids.join("|")
      );

    sheet
      .getRange(i + 1, 10)
      .setValue(maintenant);

    sheet
      .getRange(i + 1, 11)
      .setValue(maintenant);

    for (let b = 1; b < blocs.length; b++) {

      maxId++;

      const bloc = blocs[b];

      sheet.appendRow([
        maxId,
        ouvrierID,
        chantierId,
        bloc.dateDebut,
        bloc.dateFin,
        tache,
        statut,
        bloc.ids.join("|"),
        "SYNC",
        maintenant,
        maintenant,
        "CHANTIER",
        "",
        data[i][13] || ""
      ]);
    }

    nbScindees++;

    Logger.log(
      "✂✂ GANTT SCINDÉ | ancien ID=" +
      idAffectation +
      " | blocs=" +
      blocs.map(function(b) {
        return (
          b.dateDebut +
          "->" +
          b.dateFin
        );
      }).join(" + ")
    );
  }

  Logger.log("====================================");
  Logger.log("Affectations réduites : " + nbReduites);
  Logger.log("Affectations scindées : " + nbScindees);
  Logger.log("====================================");
}


function chargerIndexGoogleOptimise() {
  const timeMin = new Date("2026-08-01T00:00:00").toISOString();

  const fin = new Date();
  fin.setDate(fin.getDate() + AB_SYNC_FUTURE_DAYS);
  const timeMax = fin.toISOString();

  let pageToken = null;
  const parId = {};
  const actifs = [];
  const cancelled = [];

  do {
    const options = {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      showDeleted: true,
      maxResults: 2500
    };

    if (pageToken) {
      options.pageToken = pageToken;
    }

    const response = Calendar.Events.list(
      CALENDAR_ID,
      options
    );

    (response.items || []).forEach(function(event) {
      // IMPORTANT :
      // event.id identifie UNE occurrence.
      // iCalUID est partagé par toutes les occurrences d'une série récurrente
      // et ne doit donc jamais servir d'identité unique.
      if (event.id) {
        parId[String(event.id)] = event;
      }

      if (event.status === "cancelled") {
        cancelled.push(event);
      } else {
        actifs.push(event);
      }
    });

    pageToken = response.nextPageToken || null;
  } while (pageToken);

  Logger.log(
    "⚡ INDEX GOOGLE SAFE | actifs=" +
    actifs.length +
    " | supprimés=" +
    cancelled.length
  );

  return {
    parId: parId,
    actifs: actifs,
    cancelled: cancelled
  };
}



function testIndexGoogleOptimise() {

  const debut = new Date();

  const index =
    chargerIndexGoogleOptimise();

  const duree =
    (new Date().getTime() -
     debut.getTime()) / 1000;

  Logger.log(
    "Événements indexés : " +
    Object.keys(index.parId).length
  );

  Logger.log(
    "Actifs : " +
    index.actifs.length
  );

  Logger.log(
    "Cancelled : " +
    index.cancelled.length
  );

  Logger.log(
    "Durée : " +
    duree +
    " secondes"
  );
}
// ======================================================
// SUPPRESSIONS COMPLÈTES - VERSION OPTIMISÉE
// ======================================================

function appliquerSuppressionsGoogleCompletesOptimise(indexGoogle) {

  const debutTimer = new Date();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();

  let nbSupprimees = 0;
  let nbPartielles = 0;
  let nbRepares = 0;

  // ==================================================
  // 1. INDEX RAPIDE DES OUVRIERS
  // ==================================================

  const ouvrierParNom = {};

  ouvriers.forEach(function(o) {
    ouvrierParNom[normalizeAB(o.nom)] = Number(o.id);
  });

  // ==================================================
  // 2. RECONSTRUIRE LES LIGNES OUVRIERS
  //    DIRECTEMENT DEPUIS indexGoogle
  //
  //    clé semaine = dimanche YYYY-MM-DD
  //    puis heure -> ouvrierID
  // ==================================================

  const lignesParSemaine = {};

  indexGoogle.actifs.forEach(function(event) {

    if (
      !event.start ||
      !event.start.dateTime
    ) {
      return;
    }

    const dateEvent =
      new Date(event.start.dateTime);

    // Dimanche uniquement = lignes ouvriers
    if (dateEvent.getDay() !== 0) {
      return;
    }

    const ouvrierID =
      ouvrierParNom[
        normalizeAB(event.summary || "")
      ];

    if (!ouvrierID) {
      return;
    }

    const dimancheISO =
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    const heureCle =
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "HH:mm"
      );

    if (!lignesParSemaine[dimancheISO]) {
      lignesParSemaine[dimancheISO] = {};
    }

    lignesParSemaine[dimancheISO][heureCle] =
      ouvrierID;
  });

  // ==================================================
  // 3. INDEX DES CHANTIERS PAR NOM
  // ==================================================

  const chantierParNom = {};

  chantiers.forEach(function(c) {
    chantierParNom[
      normalizeAB(c.nom)
    ] = Number(c.id);
  });

  // ==================================================
  // 4. INDEX DES ÉVÉNEMENTS GOOGLE ACTIFS
  //
  //    clé :
  //    ouvrierID|chantierId|date
  // ==================================================

  const actifsParCle = {};

  indexGoogle.actifs.forEach(function(event) {

    if (
      !event.start ||
      !event.start.dateTime
    ) {
      return;
    }

    const chantierId =
      chantierParNom[
        normalizeAB(event.summary || "")
      ];

    if (!chantierId) {
      return;
    }

    const dateEvent =
      new Date(event.start.dateTime);

    // Le dimanche sert uniquement de repère ouvrier
    if (dateEvent.getDay() === 0) {
      return;
    }

    const dateISO =
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    const dimanche =
      abDimancheSemaine(dateEvent);

    const dimancheISO =
      Utilities.formatDate(
        dimanche,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    const heureCle =
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "HH:mm"
      );

    const lignesSemaine =
      lignesParSemaine[dimancheISO];

    if (!lignesSemaine) {
      return;
    }

    const ouvrierID =
      lignesSemaine[heureCle];

    if (!ouvrierID) {
      return;
    }

    const cle =
      Number(ouvrierID) +
      "|" +
      Number(chantierId) +
      "|" +
      dateISO;

    actifsParCle[cle] = event;
  });

  // ==================================================
  // 5. VÉRIFIER LES AFFECTATIONS
  // ==================================================

  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    const idAffectation =
      data[i][0];

    const ids =
      String(data[i][7] || "")
        .split("|")
        .map(function(id) {
          return id.trim();
        })
        .filter(Boolean);

    if (ids.length === 0) {
      continue;
    }

    let nbCancelled = 0;
    let nbActifs = 0;
    let nbInconnus = 0;

    ids.forEach(function(idComplet) {

      const idSansGoogle =
        idComplet.replace(
          /@google\.com$/i,
          ""
        );

      const event =
        indexGoogle.parId[idComplet] ||
        indexGoogle.parId[idSansGoogle];

      if (!event) {
        nbInconnus++;
        return;
      }

      if (
        event.status === "cancelled"
      ) {
        nbCancelled++;
      } else {
        nbActifs++;
      }
    });

    // ================================================
    // Les IDs ne sont pas tous supprimés
    // ================================================

    if (
      nbCancelled !== ids.length ||
      nbActifs !== 0 ||
      nbInconnus !== 0
    ) {

      if (nbCancelled > 0) {
        nbPartielles++;
      }

      continue;
    }

    // ================================================
    // Tous les anciens IDs sont CANCELLED
    // On cherche un événement actif correspondant.
    // ================================================

    const ouvrierID =
      Number(data[i][1]);

    const chantierId =
      Number(data[i][2]);

    const dateDebut =
      abIsoDate(data[i][3]);

    const dateFin =
      abIsoDate(data[i][4]);

    let evenementActifCorrespondant =
      null;

    if (
      dateDebut &&
      dateFin
    ) {

      const jours =
        abChaqueJour(
          dateDebut,
          dateFin
        );

      for (
        let j = 0;
        j < jours.length;
        j++
      ) {

        const cle =
          ouvrierID +
          "|" +
          chantierId +
          "|" +
          jours[j];

        if (actifsParCle[cle]) {

          evenementActifCorrespondant =
            actifsParCle[cle];

          break;
        }
      }
    }

    // ================================================
    // CAS A
    // L'événement existe toujours avec un nouvel ID
    // => réparer H
    // ================================================

    if (evenementActifCorrespondant) {

      const nouvelId =
        evenementActifCorrespondant.iCalUID ||
        evenementActifCorrespondant.id;

      sheet
        .getRange(i + 1, 8)
        .setValue(nouvelId);

      sheet
        .getRange(i + 1, 11)
        .setValue(new Date());

      nbRepares++;

      Logger.log(
        "🔧 ID GOOGLE RÉPARÉ | ID=" +
        idAffectation +
        " | " +
        nouvelId
      );

      continue;
    }

    // ================================================
    // CAS B
    // Plus aucun événement actif
    // => suppression réelle du Gantt
    // ================================================

    sheet.deleteRow(i + 1);

    nbSupprimees++;

    Logger.log(
      "🗑 GANTT SUPPRIMÉ | ID=" +
      idAffectation
    );
  }

  Logger.log(
    "⚡ Suppressions complètes | supprimées=" +
    nbSupprimees +
    " | réparées=" +
    nbRepares +
    " | partielles=" +
    nbPartielles +
    " | durée=" +
    (
      (new Date() - debutTimer) /
      1000
    ) +
    "s"
  );
}

// ======================================================
// SUPPRESSIONS PARTIELLES - VERSION OPTIMISÉE
// ======================================================

function appliquerSuppressionsPartiellesGoogleVersGanttOptimise(indexGoogle) {

  const debutTimer = new Date();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  let maxId = 0;

  for (let i = 1; i < data.length; i++) {

    const id = parseInt(data[i][0], 10);

    if (!isNaN(id)) {
      maxId = Math.max(maxId, id);
    }
  }

  let nbReduites = 0;
  let nbScindees = 0;

  for (let i = data.length - 1; i >= 1; i--) {

    const idAffectation = data[i][0];
    const ouvrierID = data[i][1];
    const chantierId = data[i][2];

    const dateDebut =
      abIsoDate(data[i][3]);

    const dateFin =
      abIsoDate(data[i][4]);

    const tache =
      data[i][5] || "ND";

    const statut =
      data[i][6] || "Actif";

    const ids = String(data[i][7] || "")
      .split("|")
      .map(function(id) {
        return id.trim();
      })
      .filter(Boolean);

    if (
      !dateDebut ||
      !dateFin ||
      ids.length === 0
    ) {
      continue;
    }

    const joursActifs = [];
    let nbCancelled = 0;

    ids.forEach(function(idComplet) {

      const idSansGoogle =
        idComplet.replace(
          /@google\.com$/i,
          ""
        );

      const event =
        indexGoogle.parId[idComplet] ||
        indexGoogle.parId[idSansGoogle];

      if (!event) {
        return;
      }

      if (event.status === "cancelled") {
        nbCancelled++;
        return;
      }

      if (!event.start) return;

      const dateEvent =
        event.start.date ||
        String(
          event.start.dateTime || ""
        ).substring(0, 10);

      if (dateEvent) {

        joursActifs.push({
          date: dateEvent,
          id: idComplet
        });
      }
    });

    // Pas de suppression partielle
    if (nbCancelled === 0) {
      continue;
    }

    // Tout supprimé :
    // déjà traité par la fonction précédente
    if (joursActifs.length === 0) {
      continue;
    }

    const actifsParDate = {};

    joursActifs.forEach(function(j) {
      actifsParDate[j.date] = j.id;
    });

    const datesActives =
      Object.keys(actifsParDate)
        .sort();

    const tousLesJours =
      abChaqueJour(
        dateDebut,
        dateFin
      );

    // Rien à modifier
    if (
      datesActives.length ===
      tousLesJours.length
    ) {
      continue;
    }

    // Construire les blocs consécutifs
    const blocs = [];

    datesActives.forEach(function(date) {

      const dernier =
        blocs[blocs.length - 1];

      if (dernier) {

        const d1 =
          new Date(
            dernier.dateFin +
            "T12:00:00"
          );

        const d2 =
          new Date(
            date +
            "T12:00:00"
          );

        const diff =
          Math.round(
            (d2 - d1) / 86400000
          );

        if (diff === 1) {

          dernier.dateFin = date;

          dernier.ids.push(
            actifsParDate[date]
          );

          return;
        }
      }

      blocs.push({
        dateDebut: date,
        dateFin: date,
        ids: [
          actifsParDate[date]
        ]
      });
    });

    const maintenant =
      new Date();

    // ================================================
    // RÉDUCTION SIMPLE
    // ================================================

    if (blocs.length === 1) {

      const bloc = blocs[0];

      sheet
        .getRange(i + 1, 4)
        .setValue(
          bloc.dateDebut
        );

      sheet
        .getRange(i + 1, 5)
        .setValue(
          bloc.dateFin
        );

      sheet
        .getRange(i + 1, 8)
        .setValue(
          bloc.ids.join("|")
        );

      sheet
        .getRange(i + 1, 10)
        .setValue(maintenant);

      sheet
        .getRange(i + 1, 11)
        .setValue(maintenant);

      nbReduites++;

      Logger.log(
        "✂ GANTT RÉDUIT | ID=" +
        idAffectation +
        " | " +
        bloc.dateDebut +
        " -> " +
        bloc.dateFin
      );

      continue;
    }

    // ================================================
    // SCISSION
    // ================================================

    const premier =
      blocs[0];

    sheet
      .getRange(i + 1, 4)
      .setValue(
        premier.dateDebut
      );

    sheet
      .getRange(i + 1, 5)
      .setValue(
        premier.dateFin
      );

    sheet
      .getRange(i + 1, 8)
      .setValue(
        premier.ids.join("|")
      );

    sheet
      .getRange(i + 1, 10)
      .setValue(maintenant);

    sheet
      .getRange(i + 1, 11)
      .setValue(maintenant);

    for (
      let b = 1;
      b < blocs.length;
      b++
    ) {

      const newId = getNextAffectationId();

      const bloc =
        blocs[b];

      sheet.appendRow([
        newId,
        ouvrierID,
        chantierId,
        bloc.dateDebut,
        bloc.dateFin,
        tache,
        statut,
        bloc.ids.join("|"),
        "SYNC",
        maintenant,
        maintenant,
        data[i][11] || "CHANTIER",
        data[i][12] || "",
        data[i][13] || ""
      ]);
    }

    nbScindees++;

    Logger.log(
      "✂✂ GANTT SCINDÉ | ID=" +
      idAffectation +
      " | " +
      blocs
        .map(function(b) {
          return (
            b.dateDebut +
            "->" +
            b.dateFin
          );
        })
        .join(" + ")
    );
  }

  Logger.log(
    "⚡ Suppressions partielles | réduites=" +
    nbReduites +
    " | scindées=" +
    nbScindees +
    " | durée=" +
    ((new Date() - debutTimer) / 1000) +
    "s"
  );
}


// ======================================================
// DÉPLACEMENTS 1 JOUR - VERSION OPTIMISÉE
// ======================================================

function appliquerDeplacementsGoogleVersGanttOptimise(indexGoogle) {

  const debutTimer = new Date();

  const affectations =
    getAffectations();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data =
    sheet
      .getDataRange()
      .getValues();

  let nbModifies = 0;

  affectations.forEach(function(a) {

    const d1 =
      abIsoDate(a.dateDebut);

    const d2 =
      abIsoDate(a.dateFin);

    if (!d1 || !d2) return;

    const ids =
      String(a.googleEventId || "")
        .split("|")
        .filter(Boolean);

    // Pour l'instant :
    // déplacements simples d'une journée
    if (
      d1 !== d2 ||
      ids.length !== 1
    ) {
      return;
    }

    const idComplet =
      ids[0];

    const idSansGoogle =
      idComplet.replace(
        /@google\.com$/i,
        ""
      );

    const event =
      indexGoogle.parId[idComplet] ||
      indexGoogle.parId[idSansGoogle];

    if (!event) return;

    // Suppression traitée ailleurs
    if (
      event.status ===
      "cancelled"
    ) {
      return;
    }

    if (!event.start) return;

    const nouvelleDate =
      event.start.date ||
      String(
        event.start.dateTime || ""
      ).substring(0, 10);

    if (
      !nouvelleDate ||
      nouvelleDate === d1
    ) {
      return;
    }

    const rowIndex =
      data.findIndex(
        function(row, index) {

          return (
            index > 0 &&
            Number(row[0]) ===
            Number(a.id)
          );
        }
      );

    if (rowIndex === -1) {
      return;
    }

    const ligneSheet =
      rowIndex + 1;

    const maintenant =
      new Date();

    sheet
      .getRange(
        ligneSheet,
        4
      )
      .setValue(
        nouvelleDate
      );

    sheet
      .getRange(
        ligneSheet,
        5
      )
      .setValue(
        nouvelleDate
      );

    sheet
      .getRange(
        ligneSheet,
        10
      )
      .setValue(
        maintenant
      );

    sheet
      .getRange(
        ligneSheet,
        11
      )
      .setValue(
        maintenant
      );

    nbModifies++;

    Logger.log(
      "🔄 GANTT DÉPLACÉ | ID=" +
      a.id +
      " | " +
      d1 +
      " -> " +
      nouvelleDate
    );
  });

  Logger.log(
    "⚡ Déplacements | modifiés=" +
    nbModifies +
    " | durée=" +
    ((new Date() - debutTimer) / 1000) +
    "s"
  );
}

// ======================================================
// RÉPARATION AFFICHAGE HORS_GANTT
// - typeAffectation = HORS_GANTT si chantierId vide
// - nomExterne repris depuis Google si possible
// - tache "ND" supprimée pour ne rien afficher sous le cube
// ======================================================

function reparerAffichageHorsGantt() {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  const indexGoogle = chargerIndexGoogleOptimise();

  let nbRepares = 0;

  for (let i = 1; i < data.length; i++) {

    if (data[i][0] === "") continue;

    const chantierId = data[i][2];
    const tache = String(data[i][5] || "").trim();
    const ids = String(data[i][7] || "")
      .split("|")
      .map(function(id) { return id.trim(); })
      .filter(Boolean);

    const chantierVide =
      chantierId === "" ||
      chantierId === null ||
      typeof chantierId === "undefined";

    if (!chantierVide) continue;

    let nomExterne = String(data[i][12] || "").trim();

    if (!nomExterne) {

      for (let j = 0; j < ids.length; j++) {

        const idComplet = ids[j];
        const idSansGoogle = idComplet.replace(/@google\.com$/i, "");

        const event =
          indexGoogle.parId[idComplet] ||
          indexGoogle.parId[idSansGoogle];

        if (event && event.status !== "cancelled") {
          nomExterne = String(event.summary || "").trim();
          if (nomExterne) break;
        }
      }
    }

    // L = typeAffectation
    sheet.getRange(i + 1, 12).setValue("HORS_GANTT");

    // M = nomExterne
    if (nomExterne) {
      sheet.getRange(i + 1, 13).setValue(nomExterne);
    }

    // F = tâche : supprimer ND pour les événements Google externes
    if (normalizeAB(tache) === "ND") {
      sheet.getRange(i + 1, 6).setValue("");
    }

    // K = dernière synchro
    sheet.getRange(i + 1, 11).setValue(new Date());

    nbRepares++;
  }

  Logger.log(
    "🩶 HORS_GANTT RÉPARÉS : " + nbRepares
  );
}


// ======================================================
// SYNCHRONISATION SAFE V2
// ======================================================

function abCleCibleSafe(typeAffectation, chantierId, nomExterne) {
  return normalizeAB(typeAffectation) === "HORS_GANTT"
    ? "H:" + normalizeAB(nomExterne || "")
    : "C:" + Number(chantierId);
}

function abCleJourSafe(ouvrierID, typeAffectation, chantierId, nomExterne, dateISO) {
  return (
    Number(ouvrierID) +
    "|" +
    abCleCibleSafe(typeAffectation, chantierId, nomExterne) +
    "|" +
    dateISO
  );
}

function abListeIdsSafe(value) {
  return String(value || "")
    .split("|")
    .map(function(id) { return id.trim(); })
    .filter(Boolean);
}


function nettoyerDoublonsAffectationsSafe() {

  const debutTimer = new Date();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log("🧹 DÉDOUBLONNAGE V4 | rien à traiter");
    return 0;
  }

  const header = data[0];
  const lastCol = header.length;

  const lignesGardees = [];
  const eventIdDejaVu = {};

  let nbSupprimees = 0;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    if (row[0] === "") {
      continue;
    }

    const ids = abListeIdsSafe(row[7]).map(function(id) {
      return String(id)
        .trim()
        .replace(/@google\.com$/i, "");
    }).filter(Boolean);

    // Sans ID Google : aucune déduplication automatique.
    // On ne veut jamais fusionner deux affectations distinctes
    // uniquement parce qu'elles ont le même titre ou la même date.
    if (ids.length === 0) {
      lignesGardees.push(row);
      continue;
    }

    const ouvrierID = Number(row[1]);
    const dateDebut = abIsoDate(row[3]);
    const dateFin = abIsoDate(row[4]);

    const typeAffectation =
      normalizeAB(
        row[11] ||
        (
          row[2] === "" ||
          row[2] === null
            ? "HORS_GANTT"
            : "CHANTIER"
        )
      );

    const cible =
      typeAffectation === "HORS_GANTT"
        ? "H:" + normalizeAB(row[12] || "")
        : "C:" + Number(row[2]);

    // V4.5 :
    // un vrai doublon doit avoir le même event.id ET la même
    // identité visuelle/fonctionnelle dans le Gantt.
    //
    // Cela évite de supprimer une occurrence correctement réparée
    // simplement parce que le même event.id existait auparavant
    // sur une mauvaise ligne historique.
    const identiteLigne =
      ouvrierID +
      "|" +
      dateDebut +
      "|" +
      dateFin +
      "|" +
      typeAffectation +
      "|" +
      cible;

    const tousDejaVus = ids.every(function(id) {
      return eventIdDejaVu[
        identiteLigne + "|" + id
      ] === true;
    });

    if (tousDejaVus) {
      nbSupprimees++;
      continue;
    }

    ids.forEach(function(id) {
      eventIdDejaVu[
        identiteLigne + "|" + id
      ] = true;
    });

    lignesGardees.push(row);
  }

  const ancienneNbLignes =
    Math.max(0, data.length - 1);

  if (ancienneNbLignes > 0) {
    sheet
      .getRange(
        2,
        1,
        ancienneNbLignes,
        lastCol
      )
      .clearContent();
  }

  if (lignesGardees.length > 0) {
    sheet
      .getRange(
        2,
        1,
        lignesGardees.length,
        lastCol
      )
      .setValues(
        lignesGardees.map(function(row) {
          const copie = row.slice(0, lastCol);

          while (copie.length < lastCol) {
            copie.push("");
          }

          return copie;
        })
      );
  }

  SpreadsheetApp.flush();

  Logger.log(
    "🧹 DÉDOUBLONNAGE V4 PAR EVENT.ID | supprimées=" +
    nbSupprimees +
    " | lignes restantes=" +
    lignesGardees.length +
    " | durée=" +
    ((new Date() - debutTimer) / 1000) +
    "s"
  );

  return nbSupprimees;
}


function construireGoogleJoursSafe(indexGoogle, calendar, ouvriers, chantiers) {

  const cacheLignes = {};
  const resultat = [];
  const idsVus = {};

  function lignesPourDate(date) {

    const dimanche = abDimancheSemaine(date);

    const cle = Utilities.formatDate(
      dimanche,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

    if (!cacheLignes[cle]) {
      cacheLignes[cle] = abLignesOuvriersSemaine(
        calendar,
        dimanche,
        ouvriers
      );
    }

    return cacheLignes[cle];
  }

  indexGoogle.actifs.forEach(function(event) {

    if (
      !event.id ||
      !event.start ||
      !event.start.dateTime
    ) {
      return;
    }

    // event.id = identité unique de l'occurrence Google.
    // iCalUID n'est volontairement pas utilisé ici.
    const eventId = String(event.id).trim();

    if (!eventId || idsVus[eventId]) {
      return;
    }

    idsVus[eventId] = true;

    const dateISO =
      String(event.start.dateTime).substring(0, 10);

    if (!dateISO) return;

    const dateEvent = new Date(event.start.dateTime);

    // Dimanche = repère des lignes ouvriers.
    if (dateEvent.getDay() === 0) {
      return;
    }

    const heure = Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "HH"
      )
    );

    const minute = Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "mm"
      )
    );

    const lignes = lignesPourDate(
      new Date(dateISO + "T12:00:00")
    );

    // ==================================================
    // V4.6 - RATTACHEMENT À LA PLAGE RÉELLE DE L'OUVRIER
    //
    // IMPORTANT : c'est CE bloc qui est réellement utilisé
    // par syncGoogleGantt() via construireGoogleJoursSafe().
    //
    // Tous les ouvriers : 60 minutes
    // Morvan : 105 minutes (1 h 45)
    //
    // Exemple Kevin 00:00 -> 01:00 :
    // Mvondo    00:00 -> 00:30 => Kevin
    // Firminski 00:30 -> 01:00 => Kevin
    // ==================================================

    const minuteEvent =
      (heure * 60) + minute;

    let ligne = null;

    for (let i = 0; i < lignes.length; i++) {

      const debutLigne =
        (Number(lignes[i].heure) * 60) +
        Number(lignes[i].minute);

      const nomOuvrier =
        normalizeAB(lignes[i].nom || "");

      const dureePlage =
        nomOuvrier === "MORVAN"
          ? 105
          : 60;

      const finLigne =
        debutLigne + dureePlage;

      if (
        minuteEvent >= debutLigne &&
        minuteEvent < finLigne
      ) {
        ligne = lignes[i];
        break;
      }
    }

    if (!ligne) return;

    const nomGoogle =
      String(event.summary || "").trim();

    if (!nomGoogle) return;

    const chantier = chantiers.find(function(c) {
      return (
        normalizeAB(c.nom) ===
        normalizeAB(nomGoogle)
      );
    });

    const typeAffectation =
      chantier ? "CHANTIER" : "HORS_GANTT";

    resultat.push({
      ouvrierID: Number(ligne.ouvrierID),
      ouvrierNom: ligne.nom,
      chantierId: chantier ? Number(chantier.id) : "",
      chantierNom: chantier ? chantier.nom : "",
      typeAffectation: typeAffectation,
      nomExterne: chantier ? "" : nomGoogle,
      date: dateISO,
      eventId: eventId,
      eventIds: [eventId]
    });
  });

  Logger.log(
    "⚡ OCCURRENCES GOOGLE V4 : " +
    resultat.length
  );

  return resultat;
}

function syncGoogleGanttSafe(indexGoogle) {

  const debutTimer = new Date();

  const calendar =
    CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " + CALENDAR_ID
    );
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const maintenant = new Date();

  const googleJoursBruts =
    construireGoogleJoursSafe(
      indexGoogle,
      calendar,
      ouvriers,
      chantiers
    );

  const googleJours =
    googleJoursBruts.filter(function(g) {

      if (
        estTombstoneV63(
          g.eventId
        )
      ) {

        Logger.log(
          "🪦 FULL SYNC V6.3 IGNORE | event.id=" +
          g.eventId
        );

        return false;
      }

      return true;
    });

  Logger.log(
    "⚡ OCCURRENCES GOOGLE SAFE V6.3 : " +
    googleJours.length
  );

  // ==================================================
  // V5.6 - INDEX DES AFFECTATIONS EN MÉMOIRE
  //
  // Ancienne version :
  // pour chaque événement Google, balayage de toutes
  // les affectations avec .some(), puis parsing des dates.
  //
  // Nouvelle version :
  // construction UNE SEULE FOIS de :
  // - index event.id
  // - index ouvrier + chantier + jour
  // ==================================================

  const indexEventId = {};
  const indexChantierJour = {};

  function normaliserEventId(id) {
    return String(id || "")
      .trim()
      .replace(/@google\.com$/i, "");
  }

  function typeAffectationDe(a) {
    return normalizeAB(
      a.typeAffectation ||
      (
        a.chantierId === "" ||
        a.chantierId === null
          ? "HORS_GANTT"
          : "CHANTIER"
      )
    );
  }

  function cleChantierJour(
    ouvrierID,
    chantierId,
    date
  ) {
    return (
      Number(ouvrierID) +
      "|" +
      Number(chantierId) +
      "|" +
      date
    );
  }

  affectations.forEach(function(a) {

    if (
      normalizeAB(a.statut) ===
      "ARCHIVE"
    ) {
      return;
    }

    const typeA =
      typeAffectationDe(a);

    const d1 =
      abIsoDate(a.dateDebut);

    const d2 =
      abIsoDate(a.dateFin);

    const descriptor = {
      ouvrierID:
        Number(a.ouvrierID),

      chantierId:
        a.chantierId === "" ||
        a.chantierId === null
          ? ""
          : Number(a.chantierId),

      typeAffectation:
        typeA,

      nomExterne:
        normalizeAB(
          a.nomExterne || ""
        ),

      dateDebut:
        d1,

      dateFin:
        d2
    };

    abListeIdsSafe(
      a.googleEventId
    ).forEach(function(id) {

      const idN =
        normaliserEventId(id);

      if (!idN) {
        return;
      }

      if (!indexEventId[idN]) {
        indexEventId[idN] = [];
      }

      indexEventId[idN].push(
        descriptor
      );
    });

    if (
      typeA !== "HORS_GANTT" &&
      d1 &&
      d2 &&
      a.chantierId !== "" &&
      a.chantierId !== null
    ) {

      abChaqueJour(
        d1,
        d2
      ).forEach(function(date) {

        indexChantierJour[
          cleChantierJour(
            a.ouvrierID,
            a.chantierId,
            date
          )
        ] = true;
      });
    }
  });

  // ==================================================
  // DÉTECTION DES ÉVÉNEMENTS MANQUANTS
  // ==================================================

  const manquants = [];

  googleJours.forEach(function(g) {

    const eventId =
      normaliserEventId(
        g.eventId
      );

    if (!eventId) {
      return;
    }

    const typeG =
      normalizeAB(
        g.typeAffectation ||
        "CHANTIER"
      );

    const candidats =
      indexEventId[eventId] ||
      [];

    const eventCorrectementRattache =
      candidats.some(function(a) {

        if (
          Number(a.ouvrierID) !==
          Number(g.ouvrierID)
        ) {
          return false;
        }

        if (
          !a.dateDebut ||
          !a.dateFin ||
          g.date < a.dateDebut ||
          g.date > a.dateFin
        ) {
          return false;
        }

        if (
          a.typeAffectation !==
          typeG
        ) {
          return false;
        }

        if (
          typeG === "CHANTIER"
        ) {
          return (
            Number(a.chantierId) ===
            Number(g.chantierId)
          );
        }

        return (
          a.nomExterne ===
          normalizeAB(
            g.nomExterne || ""
          )
        );
      });

    if (eventCorrectementRattache) {
      return;
    }

    // Un chantier déjà couvert dans le Gantt ne doit
    // jamais être recréé uniquement parce que son ID
    // Google historique diffère.
    if (
      typeG === "CHANTIER" &&
      indexChantierJour[
        cleChantierJour(
          g.ouvrierID,
          g.chantierId,
          g.date
        )
      ]
    ) {
      return;
    }

    manquants.push(g);
  });

  Logger.log(
    "⚡ ÉVÉNEMENTS MANQUANTS V5.6 : " +
    manquants.length
  );

  if (manquants.length === 0) {

    Logger.log(
      "✅ SAFE V5.6 | nouvelles affectations=0" +
      " | durée=" +
      (
        (new Date() - debutTimer) /
        1000
      ) +
      "s"
    );

    return;
  }

  // ==================================================
  // TRI + CONSOLIDATION
  // ==================================================

  manquants.sort(function(a, b) {

    if (
      Number(a.ouvrierID) !==
      Number(b.ouvrierID)
    ) {
      return (
        Number(a.ouvrierID) -
        Number(b.ouvrierID)
      );
    }

    if (
      a.typeAffectation !==
      b.typeAffectation
    ) {
      return String(
        a.typeAffectation
      ).localeCompare(
        String(
          b.typeAffectation
        )
      );
    }

    if (
      a.typeAffectation ===
      "CHANTIER" &&
      Number(a.chantierId) !==
      Number(b.chantierId)
    ) {
      return (
        Number(a.chantierId) -
        Number(b.chantierId)
      );
    }

    return String(
      a.date
    ).localeCompare(
      String(
        b.date
      )
    );
  });

  const groupes = [];

  manquants.forEach(function(item) {

    if (
      item.typeAffectation ===
      "HORS_GANTT"
    ) {

      groupes.push({
        ouvrierID:
          item.ouvrierID,

        ouvrierNom:
          item.ouvrierNom,

        chantierId:
          "",

        chantierNom:
          "",

        typeAffectation:
          "HORS_GANTT",

        nomExterne:
          item.nomExterne || "",

        dateDebut:
          item.date,

        dateFin:
          item.date,

        eventIds: [
          String(
            item.eventId
          )
        ]
      });

      return;
    }

    const dernier =
      groupes[
        groupes.length - 1
      ];

    const memeChantier =
      dernier &&
      dernier.typeAffectation ===
        "CHANTIER" &&
      Number(
        dernier.ouvrierID
      ) ===
        Number(
          item.ouvrierID
        ) &&
      Number(
        dernier.chantierId
      ) ===
        Number(
          item.chantierId
        );

    if (memeChantier) {

      const d1 =
        new Date(
          dernier.dateFin +
          "T12:00:00"
        );

      const d2 =
        new Date(
          item.date +
          "T12:00:00"
        );

      const diff =
        Math.round(
          (
            d2.getTime() -
            d1.getTime()
          ) /
          86400000
        );

      if (diff === 1) {

        dernier.dateFin =
          item.date;

        dernier.eventIds =
          Array.from(
            new Set(
              dernier.eventIds.concat(
                item.eventIds || [
                  item.eventId
                ]
              )
            )
          );

        return;
      }
    }

    groupes.push({
      ouvrierID:
        item.ouvrierID,

      ouvrierNom:
        item.ouvrierNom,

      chantierId:
        item.chantierId,

      chantierNom:
        item.chantierNom,

      typeAffectation:
        "CHANTIER",

      nomExterne:
        "",

      dateDebut:
        item.date,

      dateFin:
        item.date,

      eventIds:
        Array.from(
          new Set(
            item.eventIds || [
              item.eventId
            ]
          )
        )
    });
  });

  // ==================================================
  // ÉCRITURE BATCH
  //
  // Plus aucun appendRow() dans une boucle.
  // Un seul setValues() pour toutes les nouvelles lignes.
  // ==================================================

  const ids =
    reserverIdsAffectationsV56(
      groupes.length
    );

  const rows = [];

  groupes.forEach(function(g, index) {

    const googleIds =
      Array.from(
        new Set(
          (g.eventIds || [])
            .map(function(id) {
              return String(
                id || ""
              ).trim();
            })
            .filter(Boolean)
        )
      );

    if (
      googleIds.length === 0
    ) {
      return;
    }

    const newId =
      ids[index];

    const googleIdsTexte =
      googleIds.join("|");

    rows.push([
      newId,
      g.ouvrierID,
      g.typeAffectation ===
        "CHANTIER"
          ? g.chantierId
          : "",
      g.dateDebut,
      g.dateFin,
      g.typeAffectation ===
        "HORS_GANTT"
          ? ""
          : "ND",
      "Actif",
      googleIdsTexte,
      "GOOGLE",
      maintenant,
      maintenant,
      g.typeAffectation,
      g.nomExterne || "",
      ""
    ]);

    Logger.log(
      "✅ GANTT V5.6 + | ID=" +
      newId +
      " | " +
      g.ouvrierNom +
      " | " +
      (
        g.typeAffectation ===
        "HORS_GANTT"
          ? g.nomExterne
          : g.chantierNom
      ) +
      " | " +
      g.dateDebut +
      " -> " +
      g.dateFin
    );
  });

  if (rows.length > 0) {

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        14
      )
      .setValues(rows);

    SpreadsheetApp.flush();
  }

  Logger.log(
    "✅ SAFE V5.6 | nouvelles affectations=" +
    rows.length +
    " | durée=" +
    (
      (new Date() - debutTimer) /
      1000
    ) +
    "s"
  );
}


// ======================================================
// V5.6 - RÉSERVATION D'IDS EN LOT
//
// Un seul accès PropertiesService + un seul verrou,
// même si 20 nouvelles affectations sont créées.
// ======================================================

function reserverIdsAffectationsV56(
  quantite
) {

  quantite =
    Number(quantite) || 0;

  if (quantite <= 0) {
    return [];
  }

  const lock =
    LockService.getUserLock();

  if (!lock.tryLock(3000)) {
    throw new Error(
      "Réservation groupée d'IDs temporairement occupée."
    );
  }

  try {

    const props =
      PropertiesService
        .getScriptProperties();

    let nextId =
      Number(
        props.getProperty(
          "NEXT_AFFECTATION_ID"
        )
      );

    if (
      !nextId ||
      isNaN(nextId)
    ) {

      const sheet =
        SpreadsheetApp
          .openById(SHEET_ID)
          .getSheetByName(
            "affectations"
          );

      const data =
        sheet
          .getRange(
            2,
            1,
            Math.max(
              sheet.getLastRow() - 1,
              1
            ),
            1
          )
          .getValues();

      let maxId = 0;

      data.forEach(function(row) {

        const id =
          Number(row[0]);

        if (!isNaN(id)) {
          maxId =
            Math.max(
              maxId,
              id
            );
        }
      });

      nextId =
        maxId + 1;
    }

    const ids = [];

    for (
      let i = 0;
      i < quantite;
      i++
    ) {
      ids.push(
        nextId + i
      );
    }

    props.setProperty(
      "NEXT_AFFECTATION_ID",
      String(
        nextId +
        quantite
      )
    );

    return ids;

  } finally {

    lock.releaseLock();
  }
}

function appliquerSuppressionsGoogleVersGanttSafeV51(indexGoogle) {

  const DATE_DEBUT_SYNC = "2026-08-03";

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data =
    sheet.getDataRange().getValues();

  let nbLignesSupprimees = 0;
  let nbLignesReduites = 0;
  let nbLignesScindees = 0;
  let nbOccurrencesCancelled = 0;

  function getEventParId(id) {

    const brut =
      String(id || "").trim();

    if (!brut) return null;

    const sansGoogle =
      brut.replace(
        /@google\.com$/i,
        ""
      );

    return (
      indexGoogle.parId[brut] ||
      indexGoogle.parId[sansGoogle] ||
      null
    );
  }

  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    if (data[i][0] === "") {
      continue;
    }

    const idAffectation =
      Number(data[i][0]);

    const dateDebut =
      abIsoDate(data[i][3]);

    const dateFin =
      abIsoDate(data[i][4]);

    if (
      !dateDebut ||
      !dateFin ||
      dateFin < DATE_DEBUT_SYNC
    ) {
      continue;
    }

    const ids =
      abListeIdsSafe(
        data[i][7]
      );

    // ==================================================
    // PRIORITÉ V5.5 : CRÉATION GANTT PROTÉGÉE
    //
    // Une ligne sans googleEventId n'a jamais encore été
    // synchronisée vers Google. Une case vide dans Google
    // NE PEUT donc PAS être interprétée comme une suppression.
    //
    // Exemple : STANUS créé dans le Gantt
    // -> aucun googleEventId
    // -> on protège la ligne
    // -> l'étape Gantt -> Google la créera ensuite.
    // ==================================================
    if (ids.length === 0) {

      Logger.log(
        "🛡 GANTT PROTÉGÉ V5.5 | AffectationID=" +
        idAffectation +
        " | aucun googleEventId"
      );

      continue;
    }

    const actifsParDate = {};
    let nbCancelledLigne = 0;
    let nbInconnus = 0;

    ids.forEach(function(id) {

      const event =
        getEventParId(id);

      if (!event) {
        nbInconnus++;
        return;
      }

      if (
        event.status === "cancelled"
      ) {
        nbCancelledLigne++;
        nbOccurrencesCancelled++;
        return;
      }

      if (
        !event.start
      ) {
        return;
      }

      const dateEvent =
        event.start.date ||
        String(
          event.start.dateTime || ""
        ).substring(0, 10);

      if (!dateEvent) {
        return;
      }

      if (!actifsParDate[dateEvent]) {
        actifsParDate[dateEvent] = [];
      }

      actifsParDate[dateEvent].push(
        String(event.id || id)
      );
    });

    // Aucun status=cancelled explicite :
    // aucune suppression.
    if (nbCancelledLigne === 0) {
      continue;
    }

    // S'il existe un ID historique inconnu, on ne fait PAS
    // de suppression totale destructive sans certitude.
    if (
      nbInconnus > 0 &&
      Object.keys(actifsParDate).length === 0
    ) {

      Logger.log(
        "⚠ SUPPRESSION V5.1 IGNORÉE (ID inconnu) | AffectationID=" +
        idAffectation
      );

      continue;
    }

    const datesActives =
      Object.keys(actifsParDate)
        .filter(function(date) {
          return (
            date >= DATE_DEBUT_SYNC
          );
        })
        .sort();

    // ==================================================
    // CAS 1 : tous les IDs connus de la ligne sont cancelled
    // ==================================================

    if (
      datesActives.length === 0 &&
      nbInconnus === 0
    ) {

      sheet.deleteRow(i + 1);

      nbLignesSupprimees++;

      Logger.log(
        "🗑 GOOGLE -> GANTT V5.1 | supprimé AffectationID=" +
        idAffectation
      );

      continue;
    }

    // ==================================================
    // CAS 2 : suppression partielle
    // On reconstruit seulement les jours encore actifs.
    // ==================================================

    if (datesActives.length === 0) {
      continue;
    }

    const blocs = [];

    datesActives.forEach(function(date) {

      const dernier =
        blocs[blocs.length - 1];

      if (dernier) {

        const d1 =
          new Date(
            dernier.dateFin +
            "T12:00:00"
          );

        const d2 =
          new Date(
            date +
            "T12:00:00"
          );

        const diff =
          Math.round(
            (
              d2.getTime() -
              d1.getTime()
            ) /
            86400000
          );

        if (diff === 1) {

          dernier.dateFin = date;

          dernier.ids =
            dernier.ids.concat(
              actifsParDate[date]
            );

          return;
        }
      }

      blocs.push({
        dateDebut: date,
        dateFin: date,
        ids:
          actifsParDate[date].slice()
      });
    });

    const maintenant =
      new Date();

    const premier =
      blocs[0];

    // Réécrire la ligne d'origine avec le premier bloc actif.
    sheet
      .getRange(i + 1, 4)
      .setValue(
        premier.dateDebut
      );

    sheet
      .getRange(i + 1, 5)
      .setValue(
        premier.dateFin
      );

    sheet
      .getRange(i + 1, 8)
      .setValue(
        Array.from(
          new Set(premier.ids)
        ).join("|")
      );

    sheet
      .getRange(i + 1, 10)
      .setValue(maintenant);

    sheet
      .getRange(i + 1, 11)
      .setValue(maintenant);

    if (blocs.length === 1) {

      nbLignesReduites++;

      Logger.log(
        "✂ GOOGLE -> GANTT V5.1 | réduit AffectationID=" +
        idAffectation +
        " | " +
        premier.dateDebut +
        " -> " +
        premier.dateFin
      );

      continue;
    }

    // Plusieurs blocs actifs séparés :
    // scinder la ligne.
    for (
      let b = 1;
      b < blocs.length;
      b++
    ) {

      const bloc =
        blocs[b];

      const newId =
        getNextAffectationId();

      sheet.appendRow([
        newId,
        data[i][1],                  // ouvrierID
        data[i][2],                  // chantierId
        bloc.dateDebut,
        bloc.dateFin,
        data[i][5] || "ND",          // tâche
        data[i][6] || "Actif",       // statut
        Array.from(
          new Set(bloc.ids)
        ).join("|"),
        data[i][8] || "SYNC",        // source
        maintenant,
        maintenant,
        data[i][11] || "CHANTIER",   // type
        data[i][12] || "",            // nomExterne
        data[i][13] || ""              // couleur
      ]);
    }

    nbLignesScindees++;

    Logger.log(
      "✂✂ GOOGLE -> GANTT V5.1 | scindé AffectationID=" +
      idAffectation +
      " | blocs=" +
      blocs.length
    );
  }

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ SUPPRESSIONS GOOGLE -> GANTT V5.1" +
    " | occurrences cancelled=" +
    nbOccurrencesCancelled +
    " | lignes supprimées=" +
    nbLignesSupprimees +
    " | réduites=" +
    nbLignesReduites +
    " | scindées=" +
    nbLignesScindees
  );

  return {
    occurrencesCancelled:
      nbOccurrencesCancelled,

    supprimees:
      nbLignesSupprimees,

    reduites:
      nbLignesReduites,

    scindees:
      nbLignesScindees
  };
}


// ======================================================
// V5.9 - GOOGLE -> GANTT : SUPPRESSION PAR ABSENCE ACTIVE
//
// Cas corrigé : un événement a été supprimé dans Google,
// mais le googleEventId historique stocké dans le Gantt
// ne correspond plus exactement à l'occurrence cancelled.
// L'ancienne logique ne supprimait alors pas la ligne,
// puis Gantt -> Google recréait l'événement.
//
// Règle sûre :
// - on ne traite QUE les lignes ayant déjà un googleEventId ;
// - une ligne récente (< 2 min) est protégée contre les races ;
// - pour chaque jour, si aucun événement Google ACTIF ne
//   correspond à ouvrier + chantier + date, ce jour est retiré ;
// - HORS_GANTT utilise ouvrier + nomExterne + date.
// ======================================================

function appliquerSuppressionsParAbsenceGoogleV59(indexGoogle) {

  const DATE_DEBUT_SYNC = "2026-08-03";
  const DELAI_PROTECTION_MS = 5 * 60 * 1000;

  const calendar =
    CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " + CALENDAR_ID
    );
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();

  const googleJours =
    construireGoogleJoursSafe(
      indexGoogle,
      calendar,
      ouvriers,
      chantiers
    );

  const actifsChantier = {};
  const actifsHorsGantt = {};

  googleJours.forEach(function(g) {

    if (
      estTombstoneV63(
        g.eventId
      )
    ) {
      return;
    }

    if (g.date < DATE_DEBUT_SYNC) {
      return;
    }

    if (g.typeAffectation === "CHANTIER") {

      actifsChantier[
        Number(g.ouvrierID) +
        "|" +
        Number(g.chantierId) +
        "|" +
        g.date
      ] = true;

    } else {

      actifsHorsGantt[
        Number(g.ouvrierID) +
        "|" +
        normalizeAB(g.nomExterne || "") +
        "|" +
        g.date
      ] = true;
    }
  });

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const data =
    sheet.getDataRange().getValues();

  let supprimees = 0;
  let reduites = 0;
  let protegees = 0;

  for (let i = data.length - 1; i >= 1; i--) {

    const row = data[i];

    if (row[0] === "") {
      continue;
    }

    const ids =
      abListeIdsSafe(row[7]);

    // Sans ID Google = création Gantt en attente.
    if (ids.length === 0) {
      continue;
    }

    const d1 = abIsoDate(row[3]);
    const d2 = abIsoDate(row[4]);

    if (
      !d1 ||
      !d2 ||
      d2 < DATE_DEBUT_SYNC
    ) {
      continue;
    }

    // Protection contre une création Gantt qui vient juste
    // d'être poussée dans Google mais n'est pas encore visible
    // dans l'index Calendar au même instant.
    const derniereModif =
      row[10] instanceof Date
        ? row[10]
        : (
            row[9] instanceof Date
              ? row[9]
              : null
          );

    if (
      derniereModif &&
      (
        new Date().getTime() -
        derniereModif.getTime()
      ) < DELAI_PROTECTION_MS
    ) {
      protegees++;
      continue;
    }

    const type =
      normalizeAB(
        row[11] ||
        (
          row[2] === "" ||
          row[2] === null
            ? "HORS_GANTT"
            : "CHANTIER"
        )
      );

    const joursActifs = [];

    abChaqueJour(d1, d2)
      .forEach(function(jour) {

        if (jour < DATE_DEBUT_SYNC) {
          joursActifs.push(jour);
          return;
        }

        let existe = false;

        if (type === "HORS_GANTT") {

          existe =
            actifsHorsGantt[
              Number(row[1]) +
              "|" +
              normalizeAB(row[12] || "") +
              "|" +
              jour
            ] === true;

        } else {

          existe =
            actifsChantier[
              Number(row[1]) +
              "|" +
              Number(row[2]) +
              "|" +
              jour
            ] === true;
        }

        if (existe) {
          joursActifs.push(jour);
        }
      });

    // Tous les jours ont disparu de Google.
    if (joursActifs.length === 0) {

      Logger.log(
        "🗑 ABSENCE GOOGLE V5.9 | AffectationID=" +
        row[0] +
        " | supprimée du Gantt"
      );

      sheet.deleteRow(i + 1);
      supprimees++;
      continue;
    }

    // Aucun changement.
    const tousJours =
      abChaqueJour(d1, d2);

    if (
      joursActifs.length === tousJours.length
    ) {
      continue;
    }

    // Pour rester sûr et simple :
    // on conserve le premier bloc contigu de jours encore actifs
    // et on crée les autres blocs séparés si nécessaire.
    const blocs = [];

    joursActifs.sort().forEach(function(jour) {

      const dernier =
        blocs[blocs.length - 1];

      if (dernier) {

        const a =
          new Date(
            dernier.fin +
            "T12:00:00"
          );

        const b =
          new Date(
            jour +
            "T12:00:00"
          );

        const diff =
          Math.round(
            (
              b.getTime() -
              a.getTime()
            ) /
            86400000
          );

        if (diff === 1) {
          dernier.fin = jour;
          return;
        }
      }

      blocs.push({
        debut: jour,
        fin: jour
      });
    });

    const maintenant = new Date();

    const premier = blocs[0];

    sheet
      .getRange(i + 1, 4)
      .setValue(premier.debut);

    sheet
      .getRange(i + 1, 5)
      .setValue(premier.fin);

    sheet
      .getRange(i + 1, 10)
      .setValue(maintenant);

    sheet
      .getRange(i + 1, 11)
      .setValue(maintenant);

    // Si plusieurs blocs actifs restent, créer les suites.
    for (let b = 1; b < blocs.length; b++) {

      const bloc = blocs[b];
      const newId = getNextAffectationId();

      sheet.appendRow([
        newId,
        row[1],
        row[2],
        bloc.debut,
        bloc.fin,
        row[5] || "ND",
        row[6] || "Actif",
        row[7],
        row[8] || "SYNC",
        maintenant,
        maintenant,
        row[11] || "CHANTIER",
        row[12] || "",
        row[13] || ""
      ]);
    }

    reduites++;

    Logger.log(
      "✂ ABSENCE GOOGLE V5.9 | AffectationID=" +
      row[0] +
      " | " +
      d1 +
      " -> " +
      d2 +
      " devient " +
      premier.debut +
      " -> " +
      premier.fin
    );
  }

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ ABSENCE GOOGLE V5.9" +
    " | supprimées=" +
    supprimees +
    " | réduites=" +
    reduites +
    " | protégées récentes=" +
    protegees
  );

  return {
    supprimees: supprimees,
    reduites: reduites,
    protegees: protegees
  };
}

function syncGanttVersGoogleSafeV5(indexGoogle) {

  const DATE_DEBUT_SYNC = "2026-08-03";

  const calendar =
    CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " + CALENDAR_ID
    );
  }

  const ouvriers = getOuvriers();
  const chantiers = getChantiers();
  const affectations = getAffectations();

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("affectations");

  const data = sheet.getDataRange().getValues();

  const rowById = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== "") {
      rowById[Number(data[i][0])] = i + 1;
    }
  }

  // Lecture des événements Google avec EXACTEMENT la même
  // logique de rattachement ouvrier que Google -> Gantt.
  const googleJours =
    construireGoogleJoursSafe(
      indexGoogle,
      calendar,
      ouvriers,
      chantiers
    );

  const cacheLignes = {};

  function lignesPourDate(dateISO) {

    const date =
      new Date(dateISO + "T12:00:00");

    const dimanche =
      abDimancheSemaine(date);

    const cle =
      Utilities.formatDate(
        dimanche,
        "Europe/Paris",
        "yyyy-MM-dd"
      );

    if (!cacheLignes[cle]) {
      cacheLignes[cle] =
        abLignesOuvriersSemaine(
          calendar,
          dimanche,
          ouvriers
        );
    }

    return cacheLignes[cle];
  }

  let nbGoogleCrees = 0;
  let nbIdsRattaches = 0;
  let nbIdsCreationsEnregistres = 0;

  // V5.6 : index Google chantier par ouvrier + chantier + date.
  const indexGoogleChantier = {};

  googleJours.forEach(function(g) {

    if (
      g.typeAffectation !== "CHANTIER"
    ) {
      return;
    }

    const cle =
      Number(g.ouvrierID) +
      "|" +
      Number(g.chantierId) +
      "|" +
      g.date;

    if (!indexGoogleChantier[cle]) {
      indexGoogleChantier[cle] = [];
    }

    indexGoogleChantier[cle].push(g);
  });

  const chantiersParId = {};
  chantiers.forEach(function(c) {
    chantiersParId[Number(c.id)] = c;
  });

  const ouvriersParId = {};
  ouvriers.forEach(function(o) {
    ouvriersParId[Number(o.id)] = o;
  });

  affectations.forEach(function(a) {

    if (
      normalizeAB(a.statut) === "ARCHIVE"
    ) {
      return;
    }

    // Les blocs gris/HORS_GANTT sont exclusivement Google -> Gantt.
    if (
      normalizeAB(a.typeAffectation) ===
      "HORS_GANTT"
    ) {
      return;
    }


    // ==================================================
    // V6.2 - ANTI-RÉSURRECTION GOOGLE
    //
    // Si la ligne possède déjà un googleEventId, elle a déjà
    // été synchronisée au moins une fois avec Google.
    //
    // Si elle est ensuite absente de Google, on NE LA RECRÉE PAS
    // depuis le scan Gantt -> Google : cette absence peut être une
    // suppression volontaire faite dans Google.
    //
    // Seules les créations Gantt réellement neuves (googleEventId vide)
    // sont autorisées à créer un événement Google depuis ce scan.
    //
    // Les modifications Gantt explicites passent déjà par
    // updateAffectation() qui synchronise directement la seule ligne.
    // ==================================================

    const idsDejaConnus =
      abListeIdsSafe(
        a.googleEventId
      );

    if (idsDejaConnus.length > 0) {

      Logger.log(
        "🛡 ANTI-RÉSURRECTION V6.3 | AffectationID=" +
        a.id +
        " | googleEventId déjà connu"
      );

      return;
    }

    const chantier =
      chantiersParId[
        Number(a.chantierId)
      ];

    const ouvrier =
      ouvriersParId[
        Number(a.ouvrierID)
      ];

    if (!chantier || !ouvrier) {
      return;
    }

    const d1 = abIsoDate(a.dateDebut);
    const d2 = abIsoDate(a.dateFin);

    if (!d1 || !d2) {
      return;
    }

    abChaqueJour(d1, d2)
      .forEach(function(jour) {

        if (jour < DATE_DEBUT_SYNC) {
          return;
        }

        // Tous les événements Google correspondant réellement
        // à cette case Gantt.
        const cleGoogle =
          Number(a.ouvrierID) +
          "|" +
          Number(a.chantierId) +
          "|" +
          jour;

        const correspondants =
          indexGoogleChantier[
            cleGoogle
          ] || [];

        // ==================================================
        // CAS 1 : l'événement existe déjà dans Google
        // -> NE PAS recréer
        // -> rattacher son event.id canonique à la ligne Gantt
        // ==================================================

        if (correspondants.length > 0) {

          const ligneSheet =
            rowById[Number(a.id)];

          if (!ligneSheet) {
            return;
          }

          const anciensIds =
            abListeIdsSafe(
              a.googleEventId
            );

          const nouveauxIds =
            correspondants
              .map(function(g) {
                return String(
                  g.eventId || ""
                ).trim();
              })
              .filter(Boolean);

          const idsFusionnes =
            Array.from(
              new Set(
                anciensIds.concat(
                  nouveauxIds
                )
              )
            );

          if (
            idsFusionnes.join("|") !==
            anciensIds.join("|")
          ) {
            sheet
              .getRange(
                ligneSheet,
                8
              )
              .setValue(
                idsFusionnes.join("|")
              );

            sheet
              .getRange(
                ligneSheet,
                11
              )
              .setValue(new Date());

            a.googleEventId =
              idsFusionnes.join("|");

            nbIdsRattaches++;
          }

          return;
        }

        // ==================================================
        // CAS 2 : affectation Gantt absente de Google
        // -> créer UNE occurrence Google
        // ==================================================

        const lignes =
          lignesPourDate(jour);

        const ligne =
          lignes.find(function(l) {
            return (
              Number(l.ouvrierID) ===
              Number(a.ouvrierID)
            );
          });

        if (!ligne) {

          Logger.log(
            "⚠ GANTT -> GOOGLE impossible : " +
            ouvrier.nom +
            " sans repère le " +
            jour
          );

          return;
        }

        const debutEvent =
          abDateAvecHeure(
            jour,
            ligne.heure,
            ligne.minute
          );

        // La durée du bloc créé depuis le Gantt reste
        // la plage complète de l'ouvrier.
        const dureeMinutes =
          normalizeAB(ouvrier.nom) ===
          "MORVAN"
            ? 105
            : 60;

        const finEvent =
          new Date(
            debutEvent.getTime() +
            dureeMinutes *
            60 *
            1000
          );

        // IMPORTANT :
        // Advanced Calendar API => retourne le vrai event.id.
        const eventCree =
          Calendar.Events.insert(
            {
              summary:
                chantier.nom,

              description:
                "AB PLANNING\n" +
                "OuvrierID: " +
                a.ouvrierID +
                "\n" +
                "ChantierID: " +
                a.chantierId +
                "\n" +
                "AffectationID: " +
                a.id +
                "\n" +
                "Tache: " +
                (a.tache || "ND"),

              start: {
                dateTime:
                  debutEvent.toISOString(),
                timeZone:
                  "Europe/Paris"
              },

              end: {
                dateTime:
                  finEvent.toISOString(),
                timeZone:
                  "Europe/Paris"
              }
            },
            CALENDAR_ID
          );

        const eventId =
          String(
            eventCree.id || ""
          ).trim();

        if (!eventId) {
          throw new Error(
            "Google a créé l'événement sans event.id : " +
            ouvrier.nom +
            " / " +
            chantier.nom +
            " / " +
            jour
          );
        }

        // Ajouter immédiatement à l'index mémoire pour empêcher
        // une deuxième création pendant la même exécution.
        googleJours.push({
          ouvrierID:
            Number(a.ouvrierID),

          ouvrierNom:
            ouvrier.nom,

          chantierId:
            Number(a.chantierId),

          chantierNom:
            chantier.nom,

          typeAffectation:
            "CHANTIER",

          nomExterne:
            "",

          date:
            jour,

          eventId:
            eventId,

          eventIds: [
            eventId
          ]
        });

        if (
          !indexGoogleChantier[
            cleGoogle
          ]
        ) {
          indexGoogleChantier[
            cleGoogle
          ] = [];
        }

        indexGoogleChantier[
          cleGoogle
        ].push({
          ouvrierID:
            Number(a.ouvrierID),
          chantierId:
            Number(a.chantierId),
          date:
            jour,
          eventId:
            eventId,
          typeAffectation:
            "CHANTIER"
        });

        const ligneSheet =
          rowById[Number(a.id)];

        if (ligneSheet) {

          const anciensIds =
            abListeIdsSafe(
              a.googleEventId
            );

          const idsFusionnes =
            Array.from(
              new Set(
                anciensIds.concat([
                  eventId
                ])
              )
            );

          sheet
            .getRange(
              ligneSheet,
              8
            )
            .setValue(
              idsFusionnes.join("|")
            );

          sheet
            .getRange(
              ligneSheet,
              9
            )
            .setValue(
              a.source ||
              "GANTT"
            );

          sheet
            .getRange(
              ligneSheet,
              11
            )
            .setValue(
              new Date()
            );

          sheet
            .getRange(
              ligneSheet,
              12
            )
            .setValue(
              "CHANTIER"
            );

          a.googleEventId =
            idsFusionnes.join("|");

          nbIdsCreationsEnregistres++;
        }

        nbGoogleCrees++;

        Logger.log(
          "✅ GOOGLE V5 + | " +
          ouvrier.nom +
          " | " +
          chantier.nom +
          " | " +
          jour +
          " | event.id=" +
          eventId
        );
      });
  });

  SpreadsheetApp.flush();

  Logger.log(
    "➡ GANTT -> GOOGLE V6.3 ANTI-RÉSURRECTION | créés=" +
    nbGoogleCrees +
    " | IDs existants rattachés=" +
    nbIdsRattaches +
    " | IDs créations enregistrés=" +
    nbIdsCreationsEnregistres
  );

  return {
    crees: nbGoogleCrees,
    idsRattaches: nbIdsRattaches,
    idsCreationsEnregistres:
      nbIdsCreationsEnregistres
  };
}



// ======================================================
// V5.2 - INSTALLATION DU DÉCLENCHEUR 1 MINUTE
// Supprime uniquement les anciens déclencheurs de
// syncGoogleGantt puis installe un seul déclencheur.
// ======================================================

function installerSyncAutomatique5Minutes() {

  const triggers = ScriptApp.getProjectTriggers();
  let supprimes = 0;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "syncGoogleGantt") {
      ScriptApp.deleteTrigger(trigger);
      supprimes++;
    }
  });

  ScriptApp
    .newTrigger("syncGoogleGantt")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log(
    "✅ SYNC AUTO V5.8 | toutes les 5 minutes | anciens déclencheurs supprimés=" +
    supprimes
  );
}

function syncGoogleGantt() {
  Logger.log("⛔ DÉSACTIVÉ V8 : Google -> Gantt interdit. Gantt est la source unique.");
  return;
}


function diagnosticKevinStanus() {

  const affectations = getAffectations();

  Logger.log("===== DIAGNOSTIC KEVIN / STANUS =====");

  affectations.forEach(function(a) {

    if (
      Number(a.ouvrierID) === 1 &&
      Number(a.chantierId) === 10
    ) {

      Logger.log(
        "ID=" + a.id +
        " | ouvrierID=" + a.ouvrierID +
        " | chantierId=" + a.chantierId +
        " | début brut=" + a.dateDebut +
        " | fin brute=" + a.dateFin +
        " | début ISO=" + abIsoDate(a.dateDebut) +
        " | fin ISO=" + abIsoDate(a.dateFin) +
        " | statut=" + a.statut +
        " | googleEventId=" + a.googleEventId
      );
    }
  });

  Logger.log(
    "Nombre total affectations lues : " +
    affectations.length
  );

  Logger.log("===== FIN =====");
}

// ======================================================
// V5.8 - GOOGLE -> GANTT ÉVÉNEMENTIEL / RAPIDE
//
// Google Apps Script permet un déclencheur Calendar
// "onEventUpdated" qui se déclenche dès qu'un événement
// est créé, modifié ou supprimé.
//
// Cette fonction ne fait QUE Google -> Gantt :
// - suppressions
// - déplacements
// - nouveaux événements
//
// Le Gantt -> Google est déjà immédiat dans createAffectation()
// depuis la V5.7.
//
// Le cycle complet de 5 minutes reste uniquement comme
// réconciliation de sécurité.
// ======================================================


// ======================================================
// V6 - GOOGLE -> GANTT ULTRA-RAPIDE
//
// Le déclencheur Calendar ne relit plus 400+ événements.
// Il ne demande à Google que les événements MODIFIÉS
// depuis le dernier passage (avec chevauchement de sécurité).
//
// Le cycle complet syncGoogleGantt() toutes les 5 minutes
// reste le filet de sécurité.
// ======================================================

function onGoogleCalendarChangeV6(e) {
  Logger.log("⛔ DÉSACTIVÉ V8 : événement Google ignoré. Aucun retour vers Gantt.");
  return;
}


// ======================================================
// V6 - CHARGER UNIQUEMENT LES ÉVÉNEMENTS MODIFIÉS
//
// On conserve un watermark dans ScriptProperties.
// Chevauchement de 2 minutes pour ne rien rater.
// Les doublons de traitement sont sans danger car event.id
// reste l'identité unique.
// ======================================================

function chargerChangementsGoogleV6() {

  const props =
    PropertiesService.getScriptProperties();

  const maintenant =
    new Date();

  const precedent =
    props.getProperty(
      "AB_CALENDAR_LAST_UPDATE_V6"
    );

  let depuis;

  if (precedent) {

    depuis =
      new Date(precedent);

  } else {

    // Premier passage : regarder seulement les 10 dernières minutes.
    depuis =
      new Date(
        maintenant.getTime() -
        20 * 60 * 1000
      );
  }

  // Chevauchement volontaire de 2 minutes.
  depuis =
    new Date(
      depuis.getTime() -
      2 * 60 * 1000
    );

  let pageToken = null;
  const parId = {};

  do {

    const options = {
      updatedMin:
        depuis.toISOString(),

      showDeleted:
        true,

      singleEvents:
        true,

      maxResults:
        2500
    };

    if (pageToken) {
      options.pageToken =
        pageToken;
    }

    const response =
      Calendar.Events.list(
        CALENDAR_ID,
        options
      );

    (response.items || [])
      .forEach(function(event) {

        if (!event.id) {
          return;
        }

        parId[
          String(event.id)
        ] = event;
      });

    pageToken =
      response.nextPageToken ||
      null;

  } while (pageToken);

  // Positionner le watermark légèrement avant la fin
  // du traitement pour protéger les modifications concurrentes.
  props.setProperty(
    "AB_CALENDAR_LAST_UPDATE_V6",
    new Date(
      maintenant.getTime() -
      5000
    ).toISOString()
  );

  return Object.keys(parId)
    .map(function(id) {
      return parId[id];
    });
}


// ======================================================
// V6 - TRAITEMENT CIBLÉ DES ÉVÉNEMENTS MODIFIÉS
// ======================================================

function traiterChangementsGoogleV6(
  changements
) {

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const ouvriers =
    getOuvriers();

  const chantiers =
    getChantiers();

  const calendar =
    CalendarApp.getCalendarById(
      CALENDAR_ID
    );

  if (!calendar) {
    throw new Error(
      "Agenda introuvable : " +
      CALENDAR_ID
    );
  }

  // Index event.id -> AffectationID, construit une seule fois.
  const data =
    sheet.getDataRange().getValues();

  const affectationParEventId = {};

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const affectationId =
      Number(data[i][0]);

    if (!affectationId) {
      continue;
    }

    abListeIdsSafe(
      data[i][7]
    ).forEach(function(id) {

      const idN =
        normaliserGoogleIdV6(id);

      if (idN) {
        affectationParEventId[
          idN
        ] = affectationId;
      }
    });
  }

  let suppressions = 0;
  let modifications = 0;
  let creations = 0;
  let ignores = 0;

  changements.forEach(function(event) {

    const eventId =
      normaliserGoogleIdV6(
        event.id
      );

    if (!eventId) {
      return;
    }

    const affectationId =
      affectationParEventId[
        eventId
      ] || null;

    // ==================================================
    // V6.3 - ANTI-COURSE :
    // un event récemment supprimé/remplacé depuis le Gantt
    // peut encore remonter quelques secondes comme "actif".
    // On l'ignore tant que son tombstone est vivant.
    // ==================================================

    if (
      event.status !== "cancelled" &&
      estTombstoneV63(eventId)
    ) {

      ignores++;

      Logger.log(
        "🪦 GOOGLE -> GANTT V6.4 IGNORÉ (ancien ID supprimé du Gantt) | event.id=" +
        eventId
      );

      return;
    }

    if (
      event.status !== "cancelled" &&
      !affectationId
    ) {
      Logger.log(
        "🆕 GOOGLE -> GANTT V6.4 NOUVEL EVENT | event.id=" +
        eventId +
        " | summary=" +
        String(event.summary || "")
      );
    }

    if (
      event.status === "cancelled" &&
      estTombstoneV63(eventId)
    ) {
      retirerTombstoneV63(eventId);
    }

    // ==================================================
    // SUPPRESSION GOOGLE
    // ==================================================

    if (
      event.status ===
      "cancelled"
    ) {

      if (!affectationId) {

        ignores++;
        return;
      }

      retirerOccurrenceGoogleV6(
        sheet,
        affectationId,
        eventId,
        dateOccurrenceGoogleV6(
          event
        ),
        calendar,
        ouvriers,
        chantiers
      );

      suppressions++;

      Logger.log(
        "🗑 GOOGLE -> GANTT V6 | event.id=" +
        eventId +
        " | AffectationID=" +
        affectationId
      );

      return;
    }

    // Les repères du dimanche ne sont pas des affectations.
    const jour =
      convertirEventGoogleEnJourV6(
        event,
        calendar,
        ouvriers,
        chantiers
      );

    if (!jour) {
      ignores++;
      return;
    }

    // ==================================================
    // MODIFICATION / DÉPLACEMENT D'UN EVENT EXISTANT
    //
    // On retire d'abord son ancienne occurrence du Gantt,
    // puis on l'insère au nouvel endroit.
    // ==================================================

    if (affectationId) {

      retirerOccurrenceGoogleV6(
        sheet,
        affectationId,
        eventId,
        null,
        calendar,
        ouvriers,
        chantiers
      );

      ajouterOccurrenceGoogleV6(
        sheet,
        jour
      );

      modifications++;

      Logger.log(
        "🔄 GOOGLE -> GANTT V6 | " +
        jour.ouvrierNom +
        " | " +
        (
          jour.typeAffectation ===
          "CHANTIER"
            ? jour.chantierNom
            : jour.nomExterne
        ) +
        " | " +
        jour.date
      );

      return;
    }

    // ==================================================
    // NOUVEL ÉVÉNEMENT GOOGLE
    // ==================================================

    ajouterOccurrenceGoogleV6(
      sheet,
      jour
    );

    creations++;

    Logger.log(
      "✅ GOOGLE -> GANTT V6 + | " +
      jour.ouvrierNom +
      " | " +
      (
        jour.typeAffectation ===
        "CHANTIER"
          ? jour.chantierNom
          : jour.nomExterne
      ) +
      " | " +
      jour.date
    );
  });

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ V6 RAPIDE | créations=" +
    creations +
    " | modifications=" +
    modifications +
    " | suppressions=" +
    suppressions +
    " | ignorés=" +
    ignores
  );
}


// ======================================================
// V6 - CONVERTIR UN SEUL EVENT GOOGLE EN JOUR GANTT
// ======================================================

function convertirEventGoogleEnJourV6(
  event,
  calendar,
  ouvriers,
  chantiers
) {

  if (
    !event ||
    !event.id ||
    !event.start ||
    !event.start.dateTime
  ) {
    return null;
  }

  const dateEvent =
    new Date(
      event.start.dateTime
    );

  // Dimanche = repère des lignes ouvriers.
  if (
    dateEvent.getDay() === 0
  ) {
    return null;
  }

  const dateISO =
    Utilities.formatDate(
      dateEvent,
      "Europe/Paris",
      "yyyy-MM-dd"
    );

  if (
    dateISO < "2026-08-03"
  ) {
    return null;
  }

  const lignes =
    abLignesOuvriersSemaine(
      calendar,
      abDimancheSemaine(
        dateEvent
      ),
      ouvriers
    );

  const minuteEvent =
    Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "HH"
      )
    ) *
    60 +
    Number(
      Utilities.formatDate(
        dateEvent,
        "Europe/Paris",
        "mm"
      )
    );

  let ligne = null;

  for (
    let i = 0;
    i < lignes.length;
    i++
  ) {

    const debut =
      Number(lignes[i].heure) *
      60 +
      Number(lignes[i].minute);

    const duree =
      normalizeAB(
        lignes[i].nom || ""
      ) === "MORVAN"
        ? 105
        : 60;

    if (
      minuteEvent >= debut &&
      minuteEvent <
        debut + duree
    ) {

      ligne =
        lignes[i];

      break;
    }
  }

  if (!ligne) {
    return null;
  }

  const nomGoogle =
    String(
      event.summary || ""
    ).trim();

  if (!nomGoogle) {
    return null;
  }

  const chantier =
    chantiers.find(function(c) {

      return (
        normalizeAB(c.nom) ===
        normalizeAB(nomGoogle)
      );
    });

  return {
    ouvrierID:
      Number(ligne.ouvrierID),

    ouvrierNom:
      ligne.nom,

    chantierId:
      chantier
        ? Number(chantier.id)
        : "",

    chantierNom:
      chantier
        ? chantier.nom
        : "",

    typeAffectation:
      chantier
        ? "CHANTIER"
        : "HORS_GANTT",

    nomExterne:
      chantier
        ? ""
        : nomGoogle,

    date:
      dateISO,

    eventId:
      String(event.id),

    eventIds: [
      String(event.id)
    ]
  };
}


// ======================================================
// V6 - AJOUT / RATTACHEMENT D'UNE OCCURRENCE
// ======================================================

function ajouterOccurrenceGoogleV6(
  sheet,
  g
) {

  const data =
    sheet.getDataRange().getValues();

  const eventId =
    normaliserGoogleIdV6(
      g.eventId
    );

  // V6.3 : défense en profondeur contre un retour Google stale.
  if (
    estTombstoneV63(eventId)
  ) {

    Logger.log(
      "🪦 AJOUT GANTT BLOQUÉ V6.3 | event.id=" +
      eventId
    );

    return;
  }

  // 1. L'ID existe déjà : ne rien créer.
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const ids =
      abListeIdsSafe(
        data[i][7]
      )
        .map(
          normaliserGoogleIdV6
        );

    if (
      ids.indexOf(eventId) !== -1
    ) {
      return;
    }
  }

  // 2. Chercher une ligne sémantiquement identique
  //    couvrant déjà le jour.
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];

    if (
      Number(row[1]) !==
      Number(g.ouvrierID)
    ) {
      continue;
    }

    const d1 =
      abIsoDate(row[3]);

    const d2 =
      abIsoDate(row[4]);

    if (
      !d1 ||
      !d2 ||
      g.date < d1 ||
      g.date > d2
    ) {
      continue;
    }

    const type =
      normalizeAB(
        row[11] ||
        (
          row[2] === "" ||
          row[2] === null
            ? "HORS_GANTT"
            : "CHANTIER"
        )
      );

    if (
      type !==
      g.typeAffectation
    ) {
      continue;
    }

    const memeCible =
      g.typeAffectation ===
      "CHANTIER"
        ? (
            Number(row[2]) ===
            Number(g.chantierId)
          )
        : (
            normalizeAB(
              row[12] || ""
            ) ===
            normalizeAB(
              g.nomExterne || ""
            )
          );

    if (!memeCible) {
      continue;
    }

    const ids =
      abListeIdsSafe(
        row[7]
      );

    ids.push(
      g.eventId
    );

    sheet
      .getRange(
        i + 1,
        8
      )
      .setValue(
        Array.from(
          new Set(ids)
        ).join("|")
      );

    sheet
      .getRange(
        i + 1,
        11
      )
      .setValue(
        new Date()
      );

    return;
  }

  // 3. Nouvelle ligne.
  const newId =
    getNextAffectationId();

  const maintenant =
    new Date();

  sheet.appendRow([
    newId,
    g.ouvrierID,
    g.typeAffectation ===
      "CHANTIER"
        ? g.chantierId
        : "",
    g.date,
    g.date,
    g.typeAffectation ===
      "CHANTIER"
        ? "ND"
        : "",
    "Actif",
    g.eventId,
    "GOOGLE",
    maintenant,
    maintenant,
    g.typeAffectation,
    g.nomExterne || "",
    ""
  ]);
}


// ======================================================
// V6 - RETIRER UNE OCCURRENCE event.id D'UNE LIGNE
// ======================================================

function retirerOccurrenceGoogleV6(
  sheet,
  affectationId,
  eventId,
  dateOccurrence,
  calendar,
  ouvriers,
  chantiers
) {

  const data =
    sheet.getDataRange().getValues();

  let rowIndex = -1;
  let row = null;

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      Number(data[i][0]) ===
      Number(affectationId)
    ) {

      rowIndex =
        i + 1;

      row =
        data[i];

      break;
    }
  }

  if (
    rowIndex === -1 ||
    !row
  ) {
    return;
  }

  const idN =
    normaliserGoogleIdV6(
      eventId
    );

  const idsRestants =
    abListeIdsSafe(
      row[7]
    ).filter(function(id) {

      return (
        normaliserGoogleIdV6(id) !==
        idN
      );
    });

  // Ligne mono-occurrence : suppression directe.
  if (
    idsRestants.length === 0
  ) {

    sheet.deleteRow(
      rowIndex
    );

    return;
  }

  // Pour les lignes multi-jours, reconstruire uniquement
  // depuis les événements encore actifs.
  const joursRestants = [];

  idsRestants.forEach(function(id) {

    try {

      const event =
        Calendar.Events.get(
          CALENDAR_ID,
          normaliserGoogleIdV6(id)
        );

      if (
        !event ||
        event.status ===
          "cancelled"
      ) {
        return;
      }

      const g =
        convertirEventGoogleEnJourV6(
          event,
          calendar,
          ouvriers,
          chantiers
        );

      if (g) {
        joursRestants.push(g);
      }

    } catch (err) {

      Logger.log(
        "⚠ V6 event restant introuvable : " +
        id
      );
    }
  });

  if (
    joursRestants.length === 0
  ) {

    sheet.deleteRow(
      rowIndex
    );

    return;
  }

  joursRestants.sort(function(a, b) {

    return String(a.date)
      .localeCompare(
        String(b.date)
      );
  });

  // Grouper les dates consécutives avec même ouvrier/cible.
  const blocs = [];

  joursRestants.forEach(function(g) {

    const dernier =
      blocs[
        blocs.length - 1
      ];

    let memeIdentite =
      false;

    if (dernier) {

      memeIdentite =
        Number(
          dernier.ouvrierID
        ) ===
          Number(
            g.ouvrierID
          ) &&
        dernier.typeAffectation ===
          g.typeAffectation &&
        (
          g.typeAffectation ===
          "CHANTIER"
            ? Number(
                dernier.chantierId
              ) ===
              Number(
                g.chantierId
              )
            : normalizeAB(
                dernier.nomExterne
              ) ===
              normalizeAB(
                g.nomExterne
              )
        );
    }

    if (
      dernier &&
      memeIdentite
    ) {

      const d1 =
        new Date(
          dernier.dateFin +
          "T12:00:00"
        );

      const d2 =
        new Date(
          g.date +
          "T12:00:00"
        );

      const diff =
        Math.round(
          (
            d2.getTime() -
            d1.getTime()
          ) /
          86400000
        );

      if (diff === 1) {

        dernier.dateFin =
          g.date;

        dernier.ids.push(
          g.eventId
        );

        return;
      }
    }

    blocs.push({
      ouvrierID:
        g.ouvrierID,

      chantierId:
        g.chantierId,

      typeAffectation:
        g.typeAffectation,

      nomExterne:
        g.nomExterne,

      dateDebut:
        g.date,

      dateFin:
        g.date,

      ids: [
        g.eventId
      ]
    });
  });

  const maintenant =
    new Date();

  const premier =
    blocs[0];

  // Réutiliser la ligne existante pour le premier bloc.
  sheet
    .getRange(
      rowIndex,
      2,
      1,
      12
    )
    .setValues([[
      premier.ouvrierID,
      premier.typeAffectation ===
        "CHANTIER"
          ? premier.chantierId
          : "",
      premier.dateDebut,
      premier.dateFin,
      row[5] || "ND",
      "Actif",
      Array.from(
        new Set(
          premier.ids
        )
      ).join("|"),
      row[8] || "GOOGLE",
      row[9] || maintenant,
      maintenant,
      premier.typeAffectation,
      premier.nomExterne || "",
      row[13] || ""
    ]]);

  // Ajouter les blocs supplémentaires.
  for (
    let b = 1;
    b < blocs.length;
    b++
  ) {

    const bloc =
      blocs[b];

    const newId =
      getNextAffectationId();

    sheet.appendRow([
      newId,
      bloc.ouvrierID,
      bloc.typeAffectation ===
        "CHANTIER"
          ? bloc.chantierId
          : "",
      bloc.dateDebut,
      bloc.dateFin,
      row[5] || "ND",
      "Actif",
      Array.from(
        new Set(
          bloc.ids
        )
      ).join("|"),
      row[8] || "GOOGLE",
      maintenant,
      maintenant,
      bloc.typeAffectation,
      bloc.nomExterne || "",
      row[13] || ""
    ]);
  }
}


// ======================================================
// V6 - DATE D'UNE OCCURRENCE SUPPRIMÉE
// ======================================================

function dateOccurrenceGoogleV6(
  event
) {

  const value =
    event &&
    event.originalStartTime
      ? (
          event.originalStartTime.dateTime ||
          event.originalStartTime.date
        )
      : (
          event &&
          event.start
            ? (
                event.start.dateTime ||
                event.start.date
              )
            : ""
        );

  if (!value) {
    return "";
  }

  return String(value)
    .substring(0, 10);
}


// ======================================================
// V6 - NORMALISATION event.id
// ======================================================

function normaliserGoogleIdV6(
  id
) {

  return String(id || "")
    .trim()
    .replace(
      /@google\.com$/i,
      ""
    );
}


// ======================================================
// V6 - INSTALLATION DES DÉCLENCHEURS
//
// 1) Google Calendar -> Gantt : événementiel / rapide.
// 2) Réconciliation complète : toutes les 5 minutes.
// ======================================================

function installerSynchronisationV6() {

  const triggers =
    ScriptApp.getProjectTriggers();

  let supprimes = 0;

  triggers.forEach(function(trigger) {

    const handler =
      trigger.getHandlerFunction();

    if (
      handler ===
        "syncGoogleGantt" ||
      handler ===
        "onGoogleCalendarChangeV58" ||
      handler ===
        "onGoogleCalendarChangeV6"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );

      supprimes++;
    }
  });

  // Initialiser le watermark juste avant l'installation.
  PropertiesService
    .getScriptProperties()
    .setProperty(
      "AB_CALENDAR_LAST_UPDATE_V6",
      new Date(
        new Date().getTime() -
        60 * 1000
      ).toISOString()
    );

  ScriptApp
    .newTrigger(
      "syncGoogleGantt"
    )
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp
    .newTrigger(
      "onGoogleCalendarChangeV6"
    )
    .forUserCalendar(
      CALENDAR_ID
    )
    .onEventUpdated()
    .create();

  Logger.log(
    "✅ V6 installée | Google -> Gantt rapide + réconciliation 5 min | anciens triggers supprimés=" +
    supprimes
  );
}


// ======================================================
// V6 - DIAGNOSTIC DÉCLENCHEURS
// ======================================================

function diagnosticSynchronisationV6() {

  Logger.log(
    "===== DIAGNOSTIC SYNC V6 ====="
  );

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {

      Logger.log(
        "Handler=" +
        trigger.getHandlerFunction() +
        " | EventType=" +
        trigger.getEventType() +
        " | Source=" +
        trigger.getTriggerSource()
      );
    });

  Logger.log(
    "Dernier watermark=" +
    (
      PropertiesService
        .getScriptProperties()
        .getProperty(
          "AB_CALENDAR_LAST_UPDATE_V6"
        ) ||
      "AUCUN"
    )
  );

  Logger.log(
    "===== FIN DIAGNOSTIC V6 ====="
  );
}




function diagnosticTombstonesV63() {

  const map =
    lireTombstonesV63();

  Logger.log(
    "===== TOMBSTONES V6.3 ====="
  );

  Object.keys(map)
    .forEach(function(id) {

      Logger.log(
        id +
        " | expire=" +
        new Date(
          Number(map[id])
        ).toISOString()
      );
    });

  Logger.log(
    "Total=" +
    Object.keys(map).length
  );

  Logger.log(
    "===== FIN TOMBSTONES V6.3 ====="
  );
}


// ======================================================
// V6.4 - TEST CIBLÉ GOOGLE -> GANTT
// Lit uniquement les événements récemment modifiés et
// applique immédiatement les créations/modifications/suppressions.
// ======================================================

function testerGoogleVersGanttV64() {

  Logger.log(
    "===== TEST GOOGLE -> GANTT V6.4 ====="
  );

  const changements =
    chargerChangementsGoogleV6();

  Logger.log(
    "Changements récents trouvés=" +
    changements.length
  );

  changements.forEach(function(event) {

    Logger.log(
      "EVENT | id=" +
      String(event.id || "") +
      " | status=" +
      String(event.status || "") +
      " | summary=" +
      String(event.summary || "") +
      " | updated=" +
      String(event.updated || "")
    );
  });

  traiterChangementsGoogleV6(
    changements
  );

  Logger.log(
    "===== FIN TEST V6.4 ====="
  );
}



// ============================================================================
// AB PLANNING V7 — MOTEUR DE SYNCHRONISATION DELTA
// ============================================================================
//
// ARCHITECTURE :
// 1) GANTT -> GOOGLE : les CRUD create/update/delete déjà présents agissent
//    directement sur l'affectation concernée. Aucun scan.
//
// 2) GOOGLE -> GANTT : déclencheur Calendar onEventUpdated.
//    On ne lit QUE les événements modifiés depuis le dernier watermark.
//
// 3) AUDIT 5 MIN : fenêtre opérationnelle limitée à
//    semaine précédente + semaine courante + 3 semaines suivantes.
//    Il ne parcourt jamais une année entière.
//
// 4) ANTI-COURSE : tombstone uniquement sur l'ancien event.id supprimé.
//    Une nouvelle affectation reçoit un nouvel event.id et reste autorisée.
//
// 5) WATERMARK : validé uniquement APRÈS traitement réussi.
// ============================================================================

const AB_V7_TIMEZONE = "Europe/Paris";
const AB_V7_DATE_MIN = "2026-08-03";
const AB_V7_WATERMARK = "AB_CALENDAR_LAST_SUCCESS_V7";
const AB_V7_LOOKBACK_INITIAL_MIN = 20;
const AB_V7_OVERLAP_MIN = 5;


// ============================================================================
// V7 - FENÊTRE DE TRAVAIL
// Semaine précédente + courante + 3 suivantes = 5 semaines.
// ============================================================================

function fenetrePlanningV7() {

  const maintenant = new Date();

  const jour = maintenant.getDay();
  const decalageLundi =
    jour === 0
      ? -6
      : 1 - jour;

  const lundiCourant =
    new Date(maintenant);

  lundiCourant.setDate(
    lundiCourant.getDate() +
    decalageLundi
  );

  lundiCourant.setHours(
    0, 0, 0, 0
  );

  const debut =
    new Date(lundiCourant);

  debut.setDate(
    debut.getDate() - 7
  );

  const fin =
    new Date(lundiCourant);

  // 4 semaines après le lundi courant = lundi suivant les 3 semaines futures.
  fin.setDate(
    fin.getDate() + 28
  );

  fin.setHours(
    23, 59, 59, 999
  );

  return {
    debut: debut,
    fin: fin,
    debutISO:
      Utilities.formatDate(
        debut,
        AB_V7_TIMEZONE,
        "yyyy-MM-dd"
      ),
    finISO:
      Utilities.formatDate(
        fin,
        AB_V7_TIMEZONE,
        "yyyy-MM-dd"
      )
  };
}


// ============================================================================
// V7 - LOCK GLOBAL SYNCHRO
// Un seul moteur de synchro peut écrire à la fois.
// Les CRUD utilisateur restent courts et ciblés.
// ============================================================================

function verrouSyncV7() {

  return LockService
    .getScriptLock();
}


// ============================================================================
// V7 - TRIGGER GOOGLE -> GANTT
// ============================================================================

function onGoogleCalendarChangeV7(e) {
  Logger.log("⛔ DÉSACTIVÉ V8 : événement Google ignoré. Aucun retour vers Gantt.");
  return;
}


// ============================================================================
// V7 - CHARGEMENT DELTA GOOGLE
// Ne récupère que ce qui a été MODIFIÉ récemment.
// ============================================================================

function chargerDeltaGoogleV7() {

  const props =
    PropertiesService
      .getScriptProperties();

  const maintenant =
    new Date();

  const saved =
    props.getProperty(
      AB_V7_WATERMARK
    );

  let depuis;

  if (saved) {

    depuis =
      new Date(saved);

  } else {

    depuis =
      new Date(
        maintenant.getTime() -
        AB_V7_LOOKBACK_INITIAL_MIN *
        60 *
        1000
      );
  }

  // Chevauchement anti-perte.
  depuis =
    new Date(
      depuis.getTime() -
      AB_V7_OVERLAP_MIN *
      60 *
      1000
    );

  let pageToken = null;

  const parId = {};

  do {

    const options = {
      updatedMin:
        depuis.toISOString(),

      showDeleted:
        true,

      singleEvents:
        true,

      maxResults:
        2500
    };

    if (pageToken) {
      options.pageToken =
        pageToken;
    }

    const response =
      Calendar.Events.list(
        CALENDAR_ID,
        options
      );

    (response.items || [])
      .forEach(function(event) {

        if (!event.id) {
          return;
        }

        // Garder les suppressions si l'event.id est connu dans le Gantt.
        // Pour les événements actifs : on limite aux semaines opérationnelles.
        if (
          event.status !== "cancelled" &&
          !eventDansFenetreV7(
            event
          )
        ) {
          return;
        }

        parId[
          normaliserGoogleIdV6(
            event.id
          )
        ] = event;
      });

    pageToken =
      response.nextPageToken ||
      null;

  } while (pageToken);

  return {
    events:
      Object.keys(parId)
        .map(function(id) {
          return parId[id];
        }),

    // 5 secondes de marge : modification concurrente protégée.
    watermark:
      new Date(
        maintenant.getTime() -
        5000
      ).toISOString()
  };
}


// ============================================================================
// V7 - WATERMARK COMMIT
// ============================================================================

function validerWatermarkV7(
  watermark
) {

  if (!watermark) {
    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      AB_V7_WATERMARK,
      watermark
    );
}


// ============================================================================
// V7 - TEST FENÊTRE D'UN EVENT
// ============================================================================

function eventDansFenetreV7(
  event
) {

  const fenetre =
    fenetrePlanningV7();

  const raw =
    event &&
    event.start
      ? (
          event.start.dateTime ||
          event.start.date
        )
      : (
          event &&
          event.originalStartTime
            ? (
                event.originalStartTime.dateTime ||
                event.originalStartTime.date
              )
            : ""
        );

  if (!raw) {
    return false;
  }

  const date =
    String(raw)
      .substring(0, 10);

  return (
    date >= fenetre.debutISO &&
    date <= fenetre.finISO
  );
}


// ============================================================================
// V7 - CACHE CONTEXTE
// Une seule lecture Sheets + une seule lecture repères par semaine.
// ============================================================================

function construireContexteV7() {

  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      "affectations"
    );

  const data =
    sheet.getDataRange()
      .getValues();

  const ouvriers =
    getOuvriers();

  const chantiers =
    getChantiers();

  const calendar =
    CalendarApp
      .getCalendarById(
        CALENDAR_ID
      );

  if (!calendar) {

    throw new Error(
      "Agenda introuvable : " +
      CALENDAR_ID
    );
  }

  const ouvriersParId = {};
  const chantiersParId = {};
  const chantiersParNom = {};
  const rowParAffectationId = {};
  const affectationParEventId = {};

  ouvriers.forEach(function(o) {
    ouvriersParId[
      Number(o.id)
    ] = o;
  });

  chantiers.forEach(function(c) {

    chantiersParId[
      Number(c.id)
    ] = c;

    chantiersParNom[
      normalizeAB(c.nom)
    ] = c;
  });

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const id =
      Number(data[i][0]);

    if (!id) {
      continue;
    }

    rowParAffectationId[id] =
      i + 1;

    abListeIdsSafe(
      data[i][7]
    ).forEach(function(eventId) {

      const idN =
        normaliserGoogleIdV6(
          eventId
        );

      if (!idN) {
        return;
      }

      affectationParEventId[
        idN
      ] = id;
    });
  }

  return {
    ss: ss,
    sheet: sheet,
    data: data,
    ouvriers: ouvriers,
    chantiers: chantiers,
    calendar: calendar,
    ouvriersParId:
      ouvriersParId,
    chantiersParId:
      chantiersParId,
    chantiersParNom:
      chantiersParNom,
    rowParAffectationId:
      rowParAffectationId,
    affectationParEventId:
      affectationParEventId,
    lignesParSemaine: {}
  };
}


// ============================================================================
// V7 - CACHE DES REPÈRES OUVRIERS POUR UNE SEMAINE
// ============================================================================

function lignesSemaineV7(
  contexte,
  dateEvent
) {

  const dimanche =
    abDimancheSemaine(
      dateEvent
    );

  const cle =
    Utilities.formatDate(
      dimanche,
      AB_V7_TIMEZONE,
      "yyyy-MM-dd"
    );

  if (
    !contexte
      .lignesParSemaine[
        cle
      ]
  ) {

    contexte
      .lignesParSemaine[
        cle
      ] =
      abLignesOuvriersSemaine(
        contexte.calendar,
        dimanche,
        contexte.ouvriers
      );
  }

  return contexte
    .lignesParSemaine[
      cle
    ];
}


// ============================================================================
// V7 - CONVERTIR UN EVENT ACTIF EN OCCURRENCE GANTT
// ============================================================================

function eventVersJourV7(
  event,
  contexte
) {

  if (
    !event ||
    !event.id ||
    !event.start ||
    !event.start.dateTime
  ) {
    return null;
  }

  const dateEvent =
    new Date(
      event.start.dateTime
    );

  if (
    dateEvent.getDay() === 0
  ) {
    // Dimanche = repère ouvriers.
    return null;
  }

  const dateISO =
    Utilities.formatDate(
      dateEvent,
      AB_V7_TIMEZONE,
      "yyyy-MM-dd"
    );

  if (
    dateISO < AB_V7_DATE_MIN
  ) {
    return null;
  }

  const lignes =
    lignesSemaineV7(
      contexte,
      dateEvent
    );

  const minuteEvent =
    Number(
      Utilities.formatDate(
        dateEvent,
        AB_V7_TIMEZONE,
        "HH"
      )
    ) *
    60 +
    Number(
      Utilities.formatDate(
        dateEvent,
        AB_V7_TIMEZONE,
        "mm"
      )
    );

  let ligne = null;

  for (
    let i = 0;
    i < lignes.length;
    i++
  ) {

    const debut =
      Number(lignes[i].heure) *
      60 +
      Number(lignes[i].minute);

    const duree =
      normalizeAB(
        lignes[i].nom || ""
      ) === "MORVAN"
        ? 105
        : 60;

    if (
      minuteEvent >= debut &&
      minuteEvent <
        debut + duree
    ) {

      ligne =
        lignes[i];

      break;
    }
  }

  if (!ligne) {
    return null;
  }

  const titre =
    String(
      event.summary || ""
    ).trim();

  if (!titre) {
    return null;
  }

  const chantier =
    contexte
      .chantiersParNom[
        normalizeAB(titre)
      ] ||
    null;

  return {
    ouvrierID:
      Number(
        ligne.ouvrierID
      ),

    ouvrierNom:
      ligne.nom,

    chantierId:
      chantier
        ? Number(
            chantier.id
          )
        : "",

    chantierNom:
      chantier
        ? chantier.nom
        : "",

    typeAffectation:
      chantier
        ? "CHANTIER"
        : "HORS_GANTT",

    nomExterne:
      chantier
        ? ""
        : titre,

    date:
      dateISO,

    eventId:
      normaliserGoogleIdV6(
        event.id
      )
  };
}


// ============================================================================
// V7 - TRAITER LE DELTA GOOGLE EN UNE SEULE PASSE
// ============================================================================

function traiterDeltaGoogleV7(
  events
) {

  const contexte =
    construireContexteV7();

  let creations = 0;
  let modifications = 0;
  let suppressions = 0;
  let ignores = 0;

  events.forEach(function(event) {

    const eventId =
      normaliserGoogleIdV6(
        event.id
      );

    if (!eventId) {
      ignores++;
      return;
    }

    const affectationId =
      contexte
        .affectationParEventId[
          eventId
        ] ||
      null;

    // --------------------------------------------------
    // Ancien ID supprimé depuis le Gantt :
    // ne jamais le réinjecter pendant sa propagation.
    // --------------------------------------------------

    if (
      event.status !== "cancelled" &&
      estTombstoneV63(
        eventId
      )
    ) {

      Logger.log(
        "🪦 V7 retour stale ignoré | " +
        eventId
      );

      ignores++;
      return;
    }

    // --------------------------------------------------
    // SUPPRESSION GOOGLE
    // --------------------------------------------------

    if (
      event.status ===
      "cancelled"
    ) {

      if (
        estTombstoneV63(
          eventId
        )
      ) {
        retirerTombstoneV63(
          eventId
        );
      }

      if (!affectationId) {
        ignores++;
        return;
      }

      retirerOccurrenceGoogleV6(
        contexte.sheet,
        affectationId,
        eventId,
        dateOccurrenceGoogleV6(
          event
        ),
        contexte.calendar,
        contexte.ouvriers,
        contexte.chantiers
      );

      suppressions++;

      Logger.log(
        "🗑 V7 GOOGLE -> GANTT | " +
        eventId
      );

      return;
    }

    // --------------------------------------------------
    // CREATION / MODIFICATION
    // --------------------------------------------------

    const jour =
      eventVersJourV7(
        event,
        contexte
      );

    if (!jour) {
      ignores++;
      return;
    }

    if (affectationId) {

      retirerOccurrenceGoogleV6(
        contexte.sheet,
        affectationId,
        eventId,
        null,
        contexte.calendar,
        contexte.ouvriers,
        contexte.chantiers
      );

      ajouterOccurrenceGoogleV6(
        contexte.sheet,
        jour
      );

      modifications++;

      Logger.log(
        "🔄 V7 GOOGLE -> GANTT | " +
        jour.ouvrierNom +
        " | " +
        (
          jour.chantierNom ||
          jour.nomExterne
        ) +
        " | " +
        jour.date
      );

      return;
    }

    ajouterOccurrenceGoogleV6(
      contexte.sheet,
      jour
    );

    creations++;

    Logger.log(
      "✅ V7 GOOGLE -> GANTT + | " +
      jour.ouvrierNom +
      " | " +
      (
        jour.chantierNom ||
        jour.nomExterne
      ) +
      " | " +
      jour.date
    );
  });

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ V7 DELTA | créations=" +
    creations +
    " | modifications=" +
    modifications +
    " | suppressions=" +
    suppressions +
    " | ignorés=" +
    ignores
  );
}


// ============================================================================
// V7 - AUDIT LÉGER TOUTES LES 5 MINUTES
//
// Source de vérité :
// - ligne Gantt avec googleEventId => Google décide si elle existe encore.
// - ligne Gantt SANS googleEventId => nouvelle saisie Gantt / retry autorisé.
//
// Fenêtre : seulement 5 semaines opérationnelles.
// ============================================================================

function syncAuditV7() {
  Logger.log("⛔ DÉSACTIVÉ V8 : ancien audit Google -> Gantt neutralisé.");
  return;
}


// ============================================================================
// V7 - INDEX GOOGLE ACTIF LIMITÉ À LA FENÊTRE
// ============================================================================

function chargerIndexGoogleFenetreV7(
  fenetre
) {

  let pageToken = null;

  const actifs = [];
  const parId = {};

  do {

    const options = {
      timeMin:
        fenetre.debut.toISOString(),

      timeMax:
        fenetre.fin.toISOString(),

      showDeleted:
        false,

      singleEvents:
        true,

      maxResults:
        2500
    };

    if (pageToken) {
      options.pageToken =
        pageToken;
    }

    const response =
      Calendar.Events.list(
        CALENDAR_ID,
        options
      );

    (response.items || [])
      .forEach(function(event) {

        if (
          !event.id ||
          event.status ===
            "cancelled"
        ) {
          return;
        }

        const id =
          normaliserGoogleIdV6(
            event.id
          );

        actifs.push(event);
        parId[id] = event;
      });

    pageToken =
      response.nextPageToken ||
      null;

  } while (pageToken);

  Logger.log(
    "⚡ V7 index Google fenêtre=" +
    actifs.length
  );

  return {
    actifs: actifs,
    parId: parId
  };
}


// ============================================================================
// V7 - AUDIT DES ABSENCES GOOGLE
//
// Ici il n'y a aucune résurrection :
// une ligne AVEC googleEventId absente de Google est retirée/réduite du Gantt.
// ============================================================================

function auditerAbsencesGoogleV7(
  indexGoogle,
  fenetre
) {

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName(
        "affectations"
      );

  const data =
    sheet.getDataRange()
      .getValues();

  let supprimees = 0;
  let reduites = 0;

  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    const row =
      data[i];

    if (!row[0]) {
      continue;
    }

    const ids =
      abListeIdsSafe(
        row[7]
      );

    // Sans ID = nouvelle saisie Gantt, ne jamais supprimer ici.
    if (ids.length === 0) {
      continue;
    }

    const d1 =
      abIsoDate(
        row[3]
      );

    const d2 =
      abIsoDate(
        row[4]
      );

    if (
      !d1 ||
      !d2 ||
      d2 <
        fenetre.debutISO ||
      d1 >
        fenetre.finISO
    ) {
      continue;
    }

    const actifs =
      ids.filter(function(id) {

        const idN =
          normaliserGoogleIdV6(
            id
          );

        if (
          estTombstoneV63(
            idN
          )
        ) {
          return false;
        }

        return Boolean(
          indexGoogle.parId[
            idN
          ]
        );
      });

    if (
      actifs.length ===
      ids.length
    ) {
      continue;
    }

    if (
      actifs.length === 0
    ) {

      sheet.deleteRow(
        i + 1
      );

      supprimees++;

      continue;
    }

    // Reconstruction précise via les events encore actifs.
    const jours = [];

    actifs.forEach(function(id) {

      const event =
        indexGoogle.parId[
          normaliserGoogleIdV6(
            id
          )
        ];

      if (
        !event ||
        !event.start
      ) {
        return;
      }

      const raw =
        event.start.dateTime ||
        event.start.date;

      if (!raw) {
        return;
      }

      jours.push({
        date:
          String(raw)
            .substring(0, 10),
        id:
          normaliserGoogleIdV6(
            id
          )
      });
    });

    if (
      jours.length === 0
    ) {

      sheet.deleteRow(
        i + 1
      );

      supprimees++;
      continue;
    }

    jours.sort(function(a, b) {
      return a.date.localeCompare(
        b.date
      );
    });

    sheet
      .getRange(
        i + 1,
        4
      )
      .setValue(
        jours[0].date
      );

    sheet
      .getRange(
        i + 1,
        5
      )
      .setValue(
        jours[
          jours.length - 1
        ].date
      );

    sheet
      .getRange(
        i + 1,
        8
      )
      .setValue(
        jours
          .map(function(j) {
            return j.id;
          })
          .join("|")
      );

    sheet
      .getRange(
        i + 1,
        11
      )
      .setValue(
        new Date()
      );

    reduites++;
  }

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ V7 absences Google | supprimées=" +
    supprimees +
    " | réduites=" +
    reduites
  );
}


// ============================================================================
// V7 - AUDIT DES AJOUTS GOOGLE MANQUANTS
// ============================================================================

function auditerAjoutsGoogleV7(
  indexGoogle
) {

  const contexte =
    construireContexteV7();

  let ajouts = 0;

  indexGoogle.actifs
    .forEach(function(event) {

      const id =
        normaliserGoogleIdV6(
          event.id
        );

      if (
        estTombstoneV63(
          id
        )
      ) {
        return;
      }

      if (
        contexte
          .affectationParEventId[
            id
          ]
      ) {
        return;
      }

      const jour =
        eventVersJourV7(
          event,
          contexte
        );

      if (!jour) {
        return;
      }

      ajouterOccurrenceGoogleV6(
        contexte.sheet,
        jour
      );

      ajouts++;
    });

  SpreadsheetApp.flush();

  Logger.log(
    "⬅ V7 ajouts Google manquants=" +
    ajouts
  );
}


// ============================================================================
// V7 - RETRY GANTT -> GOOGLE
// Seulement les lignes neuves dont googleEventId est VIDE.
// ============================================================================

function retryNouvellesAffectationsGanttV7(
  indexGoogle,
  fenetre
) {

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName(
        "affectations"
      );

  const data =
    sheet.getDataRange()
      .getValues();

  let retries = 0;

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];

    const affectationId =
      Number(row[0]);

    if (!affectationId) {
      continue;
    }

    const type =
      normalizeAB(
        row[11] ||
        "CHANTIER"
      );

    if (
      type === "HORS_GANTT"
    ) {
      continue;
    }

    const ids =
      abListeIdsSafe(
        row[7]
      );

    // V7 règle absolue anti-résurrection :
    // s'il y a déjà eu un ID Google, ce scan ne recrée jamais.
    if (
      ids.length > 0
    ) {
      continue;
    }

    const d1 =
      abIsoDate(
        row[3]
      );

    const d2 =
      abIsoDate(
        row[4]
      );

    if (
      !d1 ||
      !d2 ||
      d2 <
        fenetre.debutISO ||
      d1 >
        fenetre.finISO
    ) {
      continue;
    }

    try {

      const nouveaux =
        creerEvenementsGoogleImmediatsV57({
          affectationId:
            affectationId,

          ouvrierID:
            Number(row[1]),

          chantierId:
            Number(row[2]),

          dateDebut:
            d1,

          dateFin:
            d2,

          tache:
            row[5] || "ND"
        });

      if (
        nouveaux.length > 0
      ) {

        sheet
          .getRange(
            i + 1,
            8
          )
          .setValue(
            nouveaux.join("|")
          );

        sheet
          .getRange(
            i + 1,
            11
          )
          .setValue(
            new Date()
          );

        retries++;
      }

    } catch (err) {

      Logger.log(
        "⚠ V7 retry Gantt->Google AffectationID=" +
        affectationId +
        " | " +
        (
          err && err.message
            ? err.message
            : String(err)
        )
      );
    }
  }

  SpreadsheetApp.flush();

  Logger.log(
    "➡ V7 retry Gantt -> Google=" +
    retries
  );
}


// ============================================================================
// V7 - INSTALLATION
//
// A lancer UNE SEULE FOIS après avoir posé le Code.gs V7.
// Supprime les anciens triggers de synchronisation et n'active que V7.
// ============================================================================

function installerSynchronisationV7() {

  const triggers =
    ScriptApp
      .getProjectTriggers();

  const handlersSync = {
    syncGoogleGantt: true,
    onGoogleCalendarChangeV58: true,
    onGoogleCalendarChangeV6: true,
    syncAuditV7: true,
    onGoogleCalendarChangeV7: true
  };

  let supprimes = 0;

  triggers.forEach(function(trigger) {

    const handler =
      trigger.getHandlerFunction();

    if (
      handlersSync[
        handler
      ]
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );

      supprimes++;
    }
  });

  // Watermark initial : 10 minutes avant installation.
  PropertiesService
    .getScriptProperties()
    .setProperty(
      AB_V7_WATERMARK,
      new Date(
        new Date().getTime() -
        10 *
        60 *
        1000
      ).toISOString()
    );

  // Google -> Gantt événementiel.
  ScriptApp
    .newTrigger(
      "onGoogleCalendarChangeV7"
    )
    .forUserCalendar(
      CALENDAR_ID
    )
    .onEventUpdated()
    .create();

  // Audit léger des 5 semaines opérationnelles.
  ScriptApp
    .newTrigger(
      "syncAuditV7"
    )
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log(
    "✅ V7 installée | delta Calendar + audit 5 semaines | anciens triggers supprimés=" +
    supprimes
  );
}


// ============================================================================
// V7 - DIAGNOSTIC
// ============================================================================

function diagnosticSynchronisationV7() {

  Logger.log(
    "===== DIAGNOSTIC V7 ====="
  );

  const fenetre =
    fenetrePlanningV7();

  Logger.log(
    "Fenêtre=" +
    fenetre.debutISO +
    " -> " +
    fenetre.finISO
  );

  Logger.log(
    "Watermark=" +
    (
      PropertiesService
        .getScriptProperties()
        .getProperty(
          AB_V7_WATERMARK
        ) ||
      "AUCUN"
    )
  );

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {

      Logger.log(
        "Trigger=" +
        trigger.getHandlerFunction() +
        " | EventType=" +
        trigger.getEventType() +
        " | Source=" +
        trigger.getTriggerSource()
      );
    });

  Logger.log(
    "===== FIN DIAGNOSTIC V7 ====="
  );
}


// ============================================================================
// V7 - TEST DELTA MANUEL
// ============================================================================

function testerGoogleVersGanttV7() {

  const lot =
    chargerDeltaGoogleV7();

  Logger.log(
    "V7 test changements=" +
    lot.events.length
  );

  traiterDeltaGoogleV7(
    lot.events
  );

  validerWatermarkV7(
    lot.watermark
  );
}




// ============================================================================
// V7.5 — GOOGLE = MIROIR DE SECOURS
// ============================================================================
//
// - Gantt est la source principale.
// - Les actions utilisateur sont enregistrées immédiatement dans le Sheet.
// - Google n'est plus appelé pendant l'action utilisateur.
// - Les opérations Gantt -> Google sont mises en file et traitées toutes les 5 min.
// - Google -> Gantt est DÉSACTIVÉ.
// - Audit de cohérence Gantt -> Google toutes les 15 min.
// ============================================================================

const AB_V75_QUEUE_PREFIX = "AB_V75_QUEUE_";
const AB_V75_QUEUE_MAX_RETRY = 6;


// --------------------------------------------------------------------------
// V7.5 - ENQUEUE SANS VERROU GLOBAL
// Une clé unique par opération : aucun conflit avec la synchro.
// --------------------------------------------------------------------------

function mettreEnFileGoogleV75(operation) {

  const props =
    PropertiesService
      .getScriptProperties();

  const cle =
    AB_V75_QUEUE_PREFIX +
    new Date().getTime() +
    "_" +
    Math.floor(Math.random() * 1000000);

  const payload =
    Object.assign(
      {
        createdAt:
          new Date().toISOString(),
        retry:
          0
      },
      operation || {}
    );

  props.setProperty(
    cle,
    JSON.stringify(payload)
  );

  Logger.log(
    "📥 V7.5 FILE + | " +
    String(payload.action || "") +
    " | AffectationID=" +
    String(payload.affectationId || "")
  );

  return cle;
}


// --------------------------------------------------------------------------
// V7.5 - LISTER LA FILE
// --------------------------------------------------------------------------

function lireFileGoogleV75() {

  const props =
    PropertiesService
      .getScriptProperties()
      .getProperties();

  return Object.keys(props)
    .filter(function(cle) {
      return cle.indexOf(
        AB_V75_QUEUE_PREFIX
      ) === 0;
    })
    .sort()
    .map(function(cle) {

      let payload = null;

      try {
        payload =
          JSON.parse(
            props[cle]
          );
      } catch (err) {
        payload = {
          action: "INVALID",
          raw: props[cle]
        };
      }

      return {
        cle: cle,
        payload: payload
      };
    });
}


// --------------------------------------------------------------------------
// V7.5 - SUPPRIMER UNE OPÉRATION DE LA FILE
// --------------------------------------------------------------------------

function retirerDeFileGoogleV75(cle) {

  PropertiesService
    .getScriptProperties()
    .deleteProperty(cle);
}


// --------------------------------------------------------------------------
// V7.5 - RETRY
// --------------------------------------------------------------------------

function reprogrammerFileGoogleV75(
  item,
  erreur
) {

  const payload =
    item.payload || {};

  payload.retry =
    Number(payload.retry || 0) + 1;

  payload.lastError =
    String(
      erreur &&
      erreur.message
        ? erreur.message
        : erreur || ""
    );

  payload.lastRetryAt =
    new Date().toISOString();

  if (
    payload.retry >
    AB_V75_QUEUE_MAX_RETRY
  ) {

    Logger.log(
      "❌ V7.5 FILE ABANDON | " +
      item.cle +
      " | " +
      payload.lastError
    );

    // On conserve la propriété en suffixe ERROR pour diagnostic,
    // mais elle n'est plus retraitée automatiquement.
    const props =
      PropertiesService
        .getScriptProperties();

    props.setProperty(
      "AB_V75_ERROR_" +
      new Date().getTime(),
      JSON.stringify(payload)
    );

    props.deleteProperty(
      item.cle
    );

    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      item.cle,
      JSON.stringify(payload)
    );
}


// --------------------------------------------------------------------------
// V7.5 - TRAITER LA FILE GANTT -> GOOGLE
// --------------------------------------------------------------------------

function traiterFileGoogleV75() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1500)) {
    Logger.log("⏭ V8.2 FILE : autre traitement en cours.");
    return;
  }

  const debut = new Date();

  try {

    let file = lireFileGoogleV75();

    Logger.log(
      "===== V8.2 FILE GANTT -> GOOGLE | opérations=" +
      file.length +
      " ====="
    );

    if (!file.length) return;

    const priorite = {
      "DELETE": 1,
      "REPLACE": 2,
      "CREATE": 3,
      "PATCH": 4
    };

    file.sort(function(a, b) {

      const pa = priorite[String((a.payload || {}).action || "")] || 99;
      const pb = priorite[String((b.payload || {}).action || "")] || 99;

      if (pa !== pb) return pa - pb;

      return String(a.cle).localeCompare(String(b.cle));
    });

    // V8.2 : 100 opérations max par passage.
    file.slice(0, 100).forEach(function(item) {

      const op = item.payload || {};

      try {

        if (op.action === "CREATE") {
          executerCreateGoogleV75(op);

        } else if (op.action === "PATCH") {
          executerPatchGoogleV75(op);

        } else if (op.action === "REPLACE") {
          executerReplaceGoogleV75(op);

        } else if (op.action === "DELETE") {
          executerDeleteGoogleV75(op);

        } else {
          throw new Error(
            "Action file inconnue : " +
            String(op.action || "")
          );
        }

        retirerDeFileGoogleV75(item.cle);

      } catch (err) {

        Logger.log(
          "⚠ V8.2 FILE retry | " +
          item.cle +
          " | " +
          (err && err.message ? err.message : String(err))
        );

        reprogrammerFileGoogleV75(item, err);
      }
    });

    Logger.log(
      "✅ V8.2 FILE terminé | durée=" +
      ((new Date() - debut) / 1000) +
      "s"
    );

  } finally {
    lock.releaseLock();
  }
}


// --------------------------------------------------------------------------
// V7.5 - RETROUVER LA LIGNE D'UNE AFFECTATION
// --------------------------------------------------------------------------

function trouverLigneAffectationV75(
  affectationId
) {

  const sheet =
    SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName("affectations");

  const data =
    sheet.getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      Number(data[i][0]) ===
      Number(affectationId)
    ) {

      return {
        sheet: sheet,
        row: i + 1,
        values: data[i]
      };
    }
  }

  return null;
}


// --------------------------------------------------------------------------
// V7.5 - CREATE GOOGLE
// --------------------------------------------------------------------------

function executerCreateGoogleV75(op) {

  const ref =
    trouverLigneAffectationV75(
      op.affectationId
    );

  if (!ref) {

    // Affectation supprimée avant que Google n'ait eu le temps de la copier :
    // rien à créer.
    Logger.log(
      "🗑 V7.5 CREATE ignoré, affectation supprimée | ID=" +
      op.affectationId
    );

    return;
  }

  const idsExistants =
    abListeIdsSafe(
      ref.values[7]
    );

  if (
    idsExistants.length > 0
  ) {

    Logger.log(
      "✅ V7.5 CREATE déjà synchronisé | ID=" +
      op.affectationId
    );

    return;
  }

  const type =
    normalizeAB(
      ref.values[11] ||
      (
        ref.values[2] === ""
          ? "HORS_GANTT"
          : "CHANTIER"
      )
    );

  const ids =
    type === "HORS_GANTT"
      ? creerEvenementsGoogleImmediatsV73({
          affectationId:
            Number(ref.values[0]),
          ouvrierID:
            Number(ref.values[1]),
          chantierId: "",
          nomAffectation:
            String(ref.values[12] || "").trim(),
          typeAffectation:
            "HORS_GANTT",
          dateDebut:
            abIsoDate(ref.values[3]),
          dateFin:
            abIsoDate(ref.values[4]),
          tache:
            String(ref.values[5] || "")
        })
      : creerEvenementsGoogleImmediatsV57({
          affectationId:
            Number(ref.values[0]),
          ouvrierID:
            Number(ref.values[1]),
          chantierId:
            Number(ref.values[2]),
          dateDebut:
            abIsoDate(ref.values[3]),
          dateFin:
            abIsoDate(ref.values[4]),
          tache:
            String(ref.values[5] || "")
        });

  if (
    ids.length > 0
  ) {

    ref.sheet
      .getRange(
        ref.row,
        8
      )
      .setValue(
        ids.join("|")
      );

    ref.sheet
      .getRange(
        ref.row,
        11
      )
      .setValue(
        new Date()
      );

    SpreadsheetApp.flush();
  }
}


// --------------------------------------------------------------------------
// V7.5 - PATCH GOOGLE
// --------------------------------------------------------------------------

function executerPatchGoogleV75(op) {

  const ref =
    trouverLigneAffectationV75(
      op.affectationId
    );

  if (!ref) {
    return;
  }

  const ids =
    abListeIdsSafe(
      ref.values[7]
    );

  if (!ids.length) {

    // Pas encore créé côté Google : le CREATE / prochain passage s'en chargera.
    mettreEnFileGoogleV75({
      action: "CREATE",
      affectationId:
        Number(op.affectationId)
    });

    return;
  }

  const type =
    normalizeAB(
      ref.values[11] ||
      (
        ref.values[2] === ""
          ? "HORS_GANTT"
          : "CHANTIER"
      )
    );

  const chantiers =
    getChantiers();

  const chantier =
    type === "CHANTIER"
      ? chantiers.find(function(c) {
          return (
            Number(c.id) ===
            Number(ref.values[2])
          );
        })
      : null;

  const summary =
    type === "HORS_GANTT"
      ? String(ref.values[12] || "").trim()
      : (
          chantier
            ? chantier.nom
            : ""
        );

  const description =
    construireDescriptionGoogleV74({
      affectationId:
        Number(ref.values[0]),
      ouvrierID:
        Number(ref.values[1]),
      chantierId:
        ref.values[2],
      typeAffectation:
        type,
      nomAffectation:
        String(ref.values[12] || ""),
      tache:
        String(ref.values[5] || "")
    });

  ids.forEach(function(eventId) {

    const idGoogle =
      normaliserGoogleIdV6(
        eventId
      );

    if (!idGoogle) {
      return;
    }

    Calendar.Events.patch(
      {
        summary: summary,
        description: description
      },
      CALENDAR_ID,
      idGoogle
    );
  });
}


// --------------------------------------------------------------------------
// V7.5 - REPLACE GOOGLE
// --------------------------------------------------------------------------

function executerReplaceGoogleV75(op) {

  const anciensIds =
    Array.isArray(op.oldEventIds)
      ? op.oldEventIds
      : abListeIdsSafe(
          op.oldEventIds || ""
        );

  anciensIds.forEach(function(eventId) {

    const idGoogle =
      normaliserGoogleIdV6(
        eventId
      );

    if (!idGoogle) {
      return;
    }

    try {

      Calendar.Events.remove(
        CALENDAR_ID,
        idGoogle
      );

    } catch (err) {

      Logger.log(
        "⚠ V7.5 ancien event déjà absent | " +
        idGoogle
      );
    }
  });

  const ref =
    trouverLigneAffectationV75(
      op.affectationId
    );

  if (!ref) {
    return;
  }

  ref.sheet
    .getRange(
      ref.row,
      8
    )
    .setValue("");

  SpreadsheetApp.flush();

  executerCreateGoogleV75({
    affectationId:
      op.affectationId
  });
}


// --------------------------------------------------------------------------
// V7.5 - DELETE GOOGLE
// --------------------------------------------------------------------------

function executerDeleteGoogleV75(op) {

  const ids =
    Array.isArray(op.eventIds)
      ? op.eventIds
      : abListeIdsSafe(
          op.eventIds || ""
        );

  ids.forEach(function(eventId) {

    const idGoogle =
      normaliserGoogleIdV6(
        eventId
      );

    if (!idGoogle) {
      return;
    }

    try {

      Calendar.Events.remove(
        CALENDAR_ID,
        idGoogle
      );

    } catch (err) {

      Logger.log(
        "⚠ V7.5 DELETE Google déjà absent | " +
        idGoogle
      );
    }
  });
}


// --------------------------------------------------------------------------
// V7.5 - GOOGLE -> GANTT TOUTES LES 5 MINUTES
// --------------------------------------------------------------------------

function syncGoogleVersGanttV75() {
  Logger.log("⛔ DÉSACTIVÉ V8 : Google -> Gantt neutralisé.");
  return;
}


// --------------------------------------------------------------------------
// V7.5 - AUDIT LÉGER 15 MINUTES
// Pas de push Gantt->Google ici : seulement cohérence Google -> Gantt.
// --------------------------------------------------------------------------

function syncAuditV75() {
  Logger.log("⛔ DÉSACTIVÉ V8 : ancien audit Google -> Gantt neutralisé.");
  return;
}



// ============================================================================
// V8 — AUDIT MIROIR GANTT -> GOOGLE
// ============================================================================
// Gantt/Sheet = source de vérité.
// Google Calendar = miroir de secours.
// Cet audit NE MODIFIE JAMAIS les dates, tâches, chantiers ou ouvriers du Gantt.
// Il contrôle seulement si les occurrences Google correspondant à une
// affectation Gantt existent encore aux bonnes dates.
// En cas d'écart, il place un REPLACE/PATCH dans la file V7.5.
// ============================================================================

function syncAuditGanttVersGoogleV80() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1500)) {
    Logger.log("⏭ V8.2 AUDIT : autre traitement en cours.");
    return;
  }

  try {

    const affectations = getAffectations();
    const file = lireFileGoogleV75();

    const pendingByAffectation = {};

    file.forEach(function(item) {
      const op = item.payload || {};
      const id = Number(op.affectationId);
      if (id) pendingByAffectation[id] = true;
    });

    let controles = 0;
    let createAjoutes = 0;
    let replaceAjoutes = 0;

    affectations.forEach(function(a) {

      if (normalizeAB(a.statut) === "ARCHIVE") return;

      const affectationId = Number(a.id);
      if (!affectationId) return;

      if (pendingByAffectation[affectationId]) return;

      const d1 = abIsoDate(a.dateDebut);
      const d2 = abIsoDate(a.dateFin);

      if (!d1 || !d2) return;

      const joursAttendus = abChaqueJour(d1, d2).filter(function(jour) {
        const d = new Date(jour + "T12:00:00");
        return d.getDay() !== 0 && d.getDay() !== 6;
      });

      if (joursAttendus.length === 0) return;

      const ids = abListeIdsSafe(a.googleEventId);

      // Jamais synchronisée : CREATE.
      if (ids.length === 0) {

        mettreEnFileGoogleV75({
          action: "CREATE",
          affectationId: affectationId
        });

        pendingByAffectation[affectationId] = true;
        createAjoutes++;
        return;
      }

      controles++;

      const datesGoogle = [];
      let anomalie = false;

      ids.forEach(function(id) {

        const idGoogle = normaliserGoogleIdV6(id);

        if (!idGoogle) {
          anomalie = true;
          return;
        }

        try {

          const event = Calendar.Events.get(
            CALENDAR_ID,
            idGoogle
          );

          if (
            !event ||
            event.status === "cancelled" ||
            !event.start
          ) {
            anomalie = true;
            return;
          }

          const dateEvent =
            event.start.date ||
            String(event.start.dateTime || "").substring(0, 10);

          if (!dateEvent) {
            anomalie = true;
            return;
          }

          datesGoogle.push(dateEvent);

        } catch (err) {
          anomalie = true;
        }
      });

      const attendues =
        Array.from(new Set(joursAttendus)).sort();

      const presentes =
        Array.from(new Set(datesGoogle)).sort();

      const memesDates =
        !anomalie &&
        attendues.length === presentes.length &&
        attendues.every(function(date, i) {
          return date === presentes[i];
        });

      // Tout est correct : aucune opération ajoutée.
      if (memesDates) {
        return;
      }

      mettreEnFileGoogleV75({
        action: "REPLACE",
        affectationId: affectationId,
        oldEventIds: ids
      });

      pendingByAffectation[affectationId] = true;
      replaceAjoutes++;
    });

    Logger.log(
      "✅ V8.2 AUDIT GANTT -> GOOGLE | contrôlées=" +
      controles +
      " | CREATE=" +
      createAjoutes +
      " | REPLACE=" +
      replaceAjoutes +
      " | PATCH inutile=0"
    );

  } finally {
    lock.releaseLock();
  }
}


// --------------------------------------------------------------------------
// V7.5 - INSTALLATION
// --------------------------------------------------------------------------

function installerSynchronisationV75() {

  const triggers = ScriptApp.getProjectTriggers();

  // Tous les anciens moteurs susceptibles de lire Google et d'écrire le Gantt
  // sont supprimés. On supprime aussi les anciens triggers sortants pour éviter
  // les doublons avant de recréer exactement les deux triggers V8 voulus.
  const handlers = {
    syncGoogleGantt: true,
    onGoogleCalendarChangeV58: true,
    onGoogleCalendarChangeV6: true,
    onGoogleCalendarChangeV7: true,
    syncAuditV7: true,
    syncAuditV75: true,
    syncGoogleVersGanttV75: true,
    traiterFileGoogleV75: true,
    syncAuditGanttVersGoogleV80: true
  };

  let supprimes = 0;

  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();

    if (handlers[handler]) {
      ScriptApp.deleteTrigger(trigger);
      supprimes++;
    }
  });

  // 1) File des actions utilisateur Gantt -> Google.
  ScriptApp
    .newTrigger("traiterFileGoogleV75")
    .timeBased()
    .everyMinutes(5)
    .create();

  // 2) Contrôle de cohérence sortant.
  // Si Google a été modifié/supprimé manuellement, le Gantt gagne.
  ScriptApp
    .newTrigger("syncAuditGanttVersGoogleV80")
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log(
    "✅ V8 MIROIR installé | Gantt = source unique | " +
    "Gantt->Google file=5 min | audit sortant=15 min | " +
    "Google->Gantt=OFF | anciens triggers supprimés=" +
    supprimes
  );
}


// --------------------------------------------------------------------------
// V7.5 - DIAGNOSTIC
// --------------------------------------------------------------------------

function diagnosticSynchronisationV75() {

  Logger.log(
    "===== DIAGNOSTIC V7.5 ====="
  );

  const file =
    lireFileGoogleV75();

  Logger.log(
    "File Google en attente=" +
    file.length
  );

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {

      Logger.log(
        "Trigger=" +
        trigger.getHandlerFunction() +
        " | EventType=" +
        trigger.getEventType() +
        " | Source=" +
        trigger.getTriggerSource()
      );
    });

  Logger.log(
    "===== FIN DIAGNOSTIC V7.5 ====="
  );
}


function getNextAffectationId() {

  // ==================================================
  // V7.2 - RÉSERVATION ID INDÉPENDANTE DE LA SYNCHRO
  //
  // IMPORTANT :
  // Le moteur V7 utilise ScriptLock pour sérialiser les
  // synchronisations Google <-> Gantt.
  //
  // L'ancienne réservation d'ID utilisait ELLE AUSSI
  // ScriptLock : une synchro de plusieurs secondes pouvait
  // donc bloquer une création utilisateur et provoquer :
  // "Impossible de réserver un ID".
  //
  // V7.2 utilise UserLock pour la très courte section
  // atomique de réservation. Elle ne se bat donc plus avec
  // le verrou global de synchronisation.
  // ==================================================

  const idLock =
    LockService.getUserLock();

  // Une réservation prend normalement quelques millisecondes.
  // On attend jusqu'à 3 s uniquement si deux créations Gantt
  // arrivent exactement en même temps.
  if (!idLock.tryLock(3000)) {

    // Filet de sécurité :
    // on ne bloque jamais l'utilisateur à cause d'une synchro.
    // Date.now() est un entier sûr en JavaScript (~13 chiffres)
    // et très éloigné des IDs séquentiels actuels.
    const fallbackId =
      new Date().getTime();

    Logger.log(
      "⚠ V7.2 ID fallback temporel=" +
      fallbackId
    );

    return fallbackId;
  }

  try {

    const props =
      PropertiesService
        .getScriptProperties();

    let nextId =
      Number(
        props.getProperty(
          "NEXT_AFFECTATION_ID"
        )
      );

    // Initialisation / réparation du compteur uniquement
    // si la propriété n'existe pas.
    if (
      !nextId ||
      isNaN(nextId)
    ) {

      const sheet =
        SpreadsheetApp
          .openById(SHEET_ID)
          .getSheetByName(
            "affectations"
          );

      const lastRow =
        sheet.getLastRow();

      let maxId = 0;

      if (lastRow >= 2) {

        const ids =
          sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              1
            )
            .getValues();

        ids.forEach(function(row) {

          const id =
            Number(row[0]);

          // Ne pas laisser un éventuel fallback timestamp
          // faire exploser le compteur séquentiel historique.
          // Les IDs standards restent dans une plage raisonnable.
          if (
            !isNaN(id) &&
            id > 0 &&
            id < 1000000000
          ) {
            maxId =
              Math.max(
                maxId,
                id
              );
          }
        });
      }

      nextId =
        maxId + 1;
    }

    // Réserver immédiatement l'ID suivant.
    props.setProperty(
      "NEXT_AFFECTATION_ID",
      String(
        nextId + 1
      )
    );

    return nextId;

  } finally {

    idLock.releaseLock();
  }
}


// ======================================================
// V7.2 - DIAGNOSTIC RÉSERVATION IDS
// ======================================================

function diagnosticReservationIdsV72() {

  const props =
    PropertiesService
      .getScriptProperties();

  Logger.log(
    "===== DIAGNOSTIC IDS V7.2 ====="
  );

  Logger.log(
    "NEXT_AFFECTATION_ID=" +
    (
      props.getProperty(
        "NEXT_AFFECTATION_ID"
      ) ||
      "NON INITIALISÉ"
    )
  );

  const idTest =
    getNextAffectationId();

  Logger.log(
    "ID test réservé=" +
    idTest
  );

  Logger.log(
    "NEXT après test=" +
    props.getProperty(
      "NEXT_AFFECTATION_ID"
    )
  );

  Logger.log(
    "===== FIN DIAGNOSTIC IDS V7.2 ====="
  );
}




// ============================================================================
// V8.2 - OUTILS DE REMISE À PLAT DE LA FILE
// ============================================================================

function diagnosticFileGoogleV82() {

  const file = lireFileGoogleV75();

  const stats = {
    CREATE: 0,
    REPLACE: 0,
    DELETE: 0,
    PATCH: 0,
    AUTRE: 0
  };

  file.forEach(function(item) {
    const action = String((item.payload || {}).action || "AUTRE");
    if (stats[action] === undefined) stats.AUTRE++;
    else stats[action]++;
  });

  Logger.log(
    "===== V8.2 FILE DIAGNOSTIC ===== | total=" +
    file.length +
    " | CREATE=" + stats.CREATE +
    " | REPLACE=" + stats.REPLACE +
    " | DELETE=" + stats.DELETE +
    " | PATCH=" + stats.PATCH +
    " | AUTRE=" + stats.AUTRE
  );

  return stats;
}


function nettoyerAncienneFileV82() {

  const props =
    PropertiesService.getScriptProperties();

  const all = props.getProperties();

  let patchSupprimes = 0;
  let erreursCalendarIdSupprimees = 0;

  Object.keys(all).forEach(function(cle) {

    if (cle.indexOf(AB_V75_QUEUE_PREFIX) === 0) {

      try {
        const payload = JSON.parse(all[cle] || "{}");

        // Supprimer seulement les anciens PATCH générés par l'audit V8.0.
        // CREATE / REPLACE / DELETE sont conservés.
        if (String(payload.action || "") === "PATCH") {
          props.deleteProperty(cle);
          patchSupprimes++;
        }
      } catch (_) {}
    }

    if (cle.indexOf("AB_V75_ERROR_") === 0) {

      try {
        const payload = JSON.parse(all[cle] || "{}");

        if (
          String(payload.lastError || "")
            .indexOf("CALENDAR_ID is not defined") !== -1
        ) {
          props.deleteProperty(cle);
          erreursCalendarIdSupprimees++;
        }
      } catch (_) {}
    }
  });

  Logger.log(
    "🧹 V8.2 NETTOYAGE | PATCH supprimés=" +
    patchSupprimes +
    " | anciennes erreurs CALENDAR_ID supprimées=" +
    erreursCalendarIdSupprimees
  );

  // Reconstituer immédiatement uniquement les vraies opérations manquantes.
  syncAuditGanttVersGoogleV80();

  return {
    patchSupprimes: patchSupprimes,
    erreursCalendarIdSupprimees: erreursCalendarIdSupprimees
  };
}
