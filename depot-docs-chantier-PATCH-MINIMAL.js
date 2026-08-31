/* =========================================================
   DOCS CHANTIER — DÉPÔT PDF YAYA / DROPBOX / LES DEUX
   Fichier complet à ajouter au dépôt GitHub docs-chantier-ab.

   À charger APRÈS le script principal de index.html :
   <script src="depot-docs-chantier.js"></script>

   Ce module :
   - neutralise la synchronisation Yaya automatique pendant sauvegarderAuto();
   - remplace le bouton "Partager le PDF" par "Déposer le PDF";
   - réutilise un seul PDF pour Yaya et Dropbox;
   - garde le PDF Drive existant pour "Voir le PDF";
   - utilise l'ID Yaya stable du chantier;
   - utilise la recherche Dropbox d'AB DOCS;
   - date Yaya = date réelle du dépôt;
   - redépôt = même ID Yaya et même nom Dropbox.
========================================================= */

(function () {
  "use strict";

  const AB_DOCS_API_URL =
    "https://script.google.com/macros/s/AKfycbxkGeQXDzn5bcI97mrFq1Lbc4DREz805dDrZxKl6ycTi5skQxQKeLJ6-bw0sNB_E_zL/exec";

  let depotPdfBase64 = "";
  let depotPdfNom = "";
  let depotDropboxId = "";
  let depotDropboxLabel = "";
  let depotRechercheTimer = null;
  let autoriserSyncYaya = false;

  function html(v) {
    return String(v == null ? "" : v)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function dateDepotISO() {
    const d = new Date();
    const offset = d.getTimezoneOffset();
    return new Date(d.getTime() - offset * 60000)
      .toISOString()
      .slice(0, 10);
  }

  function normaliser(v) {
    return String(v || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function nomPdfDeterministe() {
    const dateDoc =
      String(window.currentDocDate || dateDepotISO())
        .slice(0, 10);

    const type =
      typeof window.getDocumentTypeLabel === "function"
        ? window.getDocumentTypeLabel()
        : "Document chantier";

    const chantier =
      String(window.currentChantier || "chantier");

    const propre = function (v) {
      return String(v || "")
        .replace(/[<>:"|?*\/\\]+/g, "-")
        .replace(/\s+/g, " ")
        .trim();
    };

    return (
      propre(dateDoc) +
      " - " +
      propre(type) +
      " - " +
      propre(chantier) +
      ".pdf"
    );
  }

  async function appelerABDocs(action, payload) {
    const data = Object.assign({}, payload || {}, {
      action: action
    });

    const response = await fetch(
      AB_DOCS_API_URL,
      {
        method: "POST",
        headers: {
          "Content-Type":
            "text/plain;charset=utf-8"
        },
        body: JSON.stringify(data),
        redirect: "follow"
      }
    );

    const texte = await response.text();

    let json;
    try {
      json = JSON.parse(texte);
    } catch (e) {
      throw new Error(
        "Réponse AB DOCS invalide."
      );
    }

    if (
      !response.ok ||
      !json ||
      json.ok === false
    ) {
      throw new Error(
        (json && json.error) ||
        "AB DOCS a refusé l'opération."
      );
    }

    return json;
  }

  function creerStylesEtModale() {
    if (
      document.getElementById(
        "docs-chantier-depot-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "docs-chantier-depot-style";

    style.textContent = `
      .depot-back{
        display:none;position:fixed;inset:0;z-index:1600;
        padding:18px;background:rgba(10,20,35,.58);
        align-items:center;justify-content:center
      }
      .depot-back.open{display:flex}
      .depot-card{
        width:min(520px,100%);max-height:90vh;overflow:auto;
        background:#fff;border:1px solid #d6dfeb;border-radius:16px;
        box-shadow:0 24px 70px rgba(10,25,55,.3)
      }
      .depot-head{
        display:flex;align-items:center;justify-content:space-between;
        gap:12px;padding:18px 20px;background:#243f8f;color:#fff;
        font-size:15px;font-weight:850
      }
      .depot-close{
        width:34px;height:32px;border:1px solid rgba(255,255,255,.35);
        border-radius:7px;background:transparent;color:#fff;
        font-size:18px;cursor:pointer
      }
      .depot-body{padding:18px 20px}
      .depot-label{
        display:block;margin:0 0 7px;color:#34465c;
        font-size:11px;font-weight:800
      }
      .depot-options{
        display:grid;grid-template-columns:repeat(3,1fr);
        gap:8px;margin-bottom:18px
      }
      .depot-option{
        min-height:58px;padding:9px;border:1px solid #ccd6e5;
        border-radius:10px;background:#f7f9fc;color:#26384a;
        font-size:11px;font-weight:800;cursor:pointer
      }
      .depot-option.active{
        border-color:#4964ad;background:#eaf0ff;color:#243f8f;
        box-shadow:0 0 0 2px rgba(73,100,173,.1)
      }
      .depot-field{margin-top:14px}
      .depot-field input,.depot-field select{
        width:100%;min-width:0;height:44px;padding:0 11px;
        border:1px solid #c7d2e2;border-radius:9px;
        background:#fff;color:#172b43;font-size:13px
      }
      .depot-results{
        margin-top:5px;max-height:210px;overflow:auto;
        border:1px solid #dce3ef;border-radius:9px;background:#fff
      }
      .depot-results[hidden]{display:none}
      .depot-result{
        display:block;width:100%;padding:10px 11px;border:0;
        border-bottom:1px solid #edf1f5;background:#fff;
        text-align:left;cursor:pointer;font-size:12px
      }
      .depot-result:hover{background:#f1f5ff}
      .depot-result strong{display:block;color:#243f8f}
      .depot-result small{
        display:block;margin-top:2px;color:#718096;
        white-space:normal;overflow-wrap:anywhere
      }
      .depot-info{
        min-height:20px;margin-top:14px;color:#66758a;
        font-size:11px;line-height:1.45
      }
      .depot-actions{
        display:flex;justify-content:flex-end;gap:9px;
        padding:14px 18px;border-top:1px solid #e5eaf1;
        background:#f7f9fc
      }
      .depot-actions button{
        min-width:110px;height:40px;border-radius:8px;
        font-size:12px;font-weight:800;cursor:pointer
      }
      .depot-cancel{
        border:1px solid #cbd5e1;background:#fff;color:#34465c
      }
      .depot-confirm{
        border:1px solid #294796;background:#294796;color:#fff
      }
      .depot-confirm:disabled{opacity:.6;cursor:wait}
      .depot-result-ok{color:#16845b;font-weight:800}
      .depot-result-err{color:#b5453c;font-weight:800}
      @media(max-width:620px){
        .depot-options{grid-template-columns:1fr}
        .depot-card{max-height:94vh}
      }
    `;

    document.head.appendChild(style);

    const back =
      document.createElement("div");

    back.className = "depot-back";
    back.id = "depot-pdf-back";

    back.innerHTML = `
      <div class="depot-card" role="dialog" aria-modal="true" aria-labelledby="depot-pdf-title">
        <div class="depot-head">
          <span id="depot-pdf-title">Déposer le PDF</span>
          <button type="button" class="depot-close" id="depot-pdf-close">✕</button>
        </div>

        <div class="depot-body">
          <span class="depot-label">Destination</span>

          <div class="depot-options">
            <button type="button" class="depot-option active" data-depot="yaya">📄 Yaya</button>
            <button type="button" class="depot-option" data-depot="dropbox">📁 Dropbox</button>
            <button type="button" class="depot-option" data-depot="yaya_dropbox">📄 + 📁 Yaya + Dropbox</button>
          </div>

          <div class="depot-field" id="depot-yaya-bloc">
            <label class="depot-label" for="depot-yaya-chantier">Chantier Yaya</label>
            <select id="depot-yaya-chantier">
              <option value="">Choisir le chantier…</option>
            </select>
          </div>

          <div class="depot-field" id="depot-dropbox-bloc" hidden>
            <label class="depot-label" for="depot-dropbox-recherche">Dossier Dropbox</label>
            <input id="depot-dropbox-recherche" type="search"
              placeholder="Tapez le nom du dossier…" autocomplete="off">
            <div class="depot-results" id="depot-dropbox-resultats" hidden></div>
          </div>

          <div class="depot-info" id="depot-pdf-info"></div>
        </div>

        <div class="depot-actions">
          <button type="button" class="depot-cancel" id="depot-pdf-cancel">Annuler</button>
          <button type="button" class="depot-confirm" id="depot-pdf-confirm">Déposer</button>
        </div>
      </div>
    `;

    document.body.appendChild(back);

    back.addEventListener(
      "click",
      function (e) {
        if (e.target === back) {
          fermerDepot();
        }
      }
    );

    document
      .getElementById("depot-pdf-close")
      .addEventListener("click", fermerDepot);

    document
      .getElementById("depot-pdf-cancel")
      .addEventListener("click", fermerDepot);

    document
      .querySelectorAll(".depot-option")
      .forEach(function (b) {
        b.addEventListener(
          "click",
          function () {
            selectionDestination(
              b.getAttribute("data-depot")
            );
          }
        );
      });

    document
      .getElementById(
        "depot-dropbox-recherche"
      )
      .addEventListener(
        "input",
        rechercherDropbox
      );

    document
      .getElementById(
        "depot-pdf-confirm"
      )
      .addEventListener(
        "click",
        confirmerDepot
      );
  }

  let destinationCourante = "yaya";

  function selectionDestination(dest) {
    destinationCourante = dest;

    document
      .querySelectorAll(".depot-option")
      .forEach(function (b) {
        b.classList.toggle(
          "active",
          b.getAttribute("data-depot") === dest
        );
      });

    const doitYaya =
      dest === "yaya" ||
      dest === "yaya_dropbox";

    const doitDropbox =
      dest === "dropbox" ||
      dest === "yaya_dropbox";

    document.getElementById(
      "depot-yaya-bloc"
    ).hidden = !doitYaya;

    document.getElementById(
      "depot-dropbox-bloc"
    ).hidden = !doitDropbox;

    document.getElementById(
      "depot-pdf-info"
    ).textContent = "";
  }

  function chantierLabel(c) {
    return String(
      c && (
        c.nom ||
        c.chantier ||
        c.client
      ) || ""
    ).trim();
  }

  function remplirChantiersYaya() {
    const select =
      document.getElementById(
        "depot-yaya-chantier"
      );

    select.innerHTML =
      '<option value="">Choisir le chantier…</option>';

    const liste =
      Array.isArray(window.yayaChantiers)
        ? window.yayaChantiers
        : [];

    liste.forEach(function (c) {
      const nom = chantierLabel(c);
      const id = String(c && c.id || "");

      if (!id || !nom) return;

      const option =
        document.createElement("option");

      option.value = id;
      option.textContent = nom;
      select.appendChild(option);
    });

    // Présélection sur le chantier courant.
    const courant =
      normaliser(window.currentChantier || "");

    if (courant) {
      const match =
        liste.find(function (c) {
          return (
            normaliser(
              chantierLabel(c)
            ) === courant
          );
        });

      if (match && match.id) {
        select.value =
          String(match.id);
      }
    }
  }

  function ouvrirDepot() {
    creerStylesEtModale();

    if (!window.currentDocId) {
      alert(
        "Le document doit d'abord être sauvegardé."
      );
      return;
    }

    depotDropboxId = "";
    depotDropboxLabel = "";

    const champ =
      document.getElementById(
        "depot-dropbox-recherche"
      );

    const zone =
      document.getElementById(
        "depot-dropbox-resultats"
      );

    champ.value = "";
    zone.innerHTML = "";
    zone.hidden = true;

    selectionDestination("yaya");
    remplirChantiersYaya();

    document.getElementById(
      "depot-pdf-info"
    ).textContent =
      "Le PDF courant sera utilisé. Aucun doublon de génération.";

    document.getElementById(
      "depot-pdf-back"
    ).classList.add("open");
  }

  function fermerDepot() {
    const back =
      document.getElementById(
        "depot-pdf-back"
      );

    if (back) {
      back.classList.remove("open");
    }
  }

  async function rechercherDropbox() {
    const champ =
      document.getElementById(
        "depot-dropbox-recherche"
      );

    const zone =
      document.getElementById(
        "depot-dropbox-resultats"
      );

    const terme =
      champ.value.trim();

    depotDropboxId = "";
    depotDropboxLabel = "";

    clearTimeout(
      depotRechercheTimer
    );

    if (terme.length < 2) {
      zone.hidden = true;
      return;
    }

    zone.hidden = false;
    zone.innerHTML =
      '<div style="padding:10px;color:#66758a;font-size:12px">Recherche…</div>';

    depotRechercheTimer =
      setTimeout(
        async function () {
          try {
            const r =
              await appelerABDocs(
                "RECHERCHER_DESTINATIONS_DROPBOX",
                {
                  terme: terme
                }
              );

            const resultats =
              Array.isArray(r.destinations)
                ? r.destinations
                : [];

            if (!resultats.length) {
              zone.innerHTML =
                '<div style="padding:10px;color:#66758a;font-size:12px">Aucun dossier trouvé.</div>';
              return;
            }

            zone.innerHTML = "";

            resultats.forEach(
              function (o) {
                const b =
                  document.createElement(
                    "button"
                  );

                b.type = "button";
                b.className =
                  "depot-result";

                b.innerHTML =
                  "<strong>" +
                  html(
                    o.label ||
                    o.nom ||
                    "Dossier"
                  ) +
                  "</strong>" +
                  "<small>" +
                  html(
                    o.chemin || ""
                  ) +
                  "</small>";

                b.addEventListener(
                  "click",
                  function () {
                    depotDropboxId =
                      String(
                        o.value ||
                        o.id ||
                        ""
                      );

                    depotDropboxLabel =
                      String(
                        o.label ||
                        o.nom ||
                        ""
                      );

                    champ.value =
                      depotDropboxLabel;

                    zone.hidden = true;
                  }
                );

                zone.appendChild(b);
              }
            );
          } catch (e) {
            zone.innerHTML =
              '<div style="padding:10px;color:#b5453c;font-size:12px">' +
              html(e.message) +
              "</div>";
          }
        },
        250
      );
  }

  async function fabriquerPdfEtSauverDrive() {
    if (!window.currentDocId) {
      throw new Error(
        "Document non sauvegardé."
      );
    }

    const source =
      document.getElementById(
        "doc-frame"
      );

    if (!source) {
      throw new Error(
        "Document à convertir introuvable."
      );
    }

    const canvas =
      await html2canvas(
        source,
        {
          scale: 1.6,
          useCORS: true,
          backgroundColor: "#ffffff"
        }
      );

    const imgData =
      canvas.toDataURL(
        "image/jpeg",
        0.9
      );

    const jsPDF =
      window.jspdf &&
      window.jspdf.jsPDF;

    if (!jsPDF) {
      throw new Error(
        "Moteur PDF indisponible."
      );
    }

    const pdf =
      new jsPDF(
        "p",
        "mm",
        "a4"
      );

    const pageW =
      pdf.internal.pageSize.getWidth();

    const pageH =
      pdf.internal.pageSize.getHeight();

    const imgW = pageW;
    const imgH =
      canvas.height *
      imgW /
      canvas.width;

    let heightLeft = imgH;
    let position = 0;

    pdf.addImage(
      imgData,
      "JPEG",
      0,
      position,
      imgW,
      imgH
    );

    heightLeft -= pageH;

    while (heightLeft > 0) {
      position =
        heightLeft - imgH;

      pdf.addPage();

      pdf.addImage(
        imgData,
        "JPEG",
        0,
        position,
        imgW,
        imgH
      );

      heightLeft -= pageH;
    }

    depotPdfBase64 =
      pdf.output(
        "datauristring"
      ).split(",")[1];

    depotPdfNom =
      nomPdfDeterministe();

    const r =
      await window.callGAS({
        action: "save_pdf",
        id: window.currentDocId,
        base64: depotPdfBase64
      });

    if (!r || !r.url) {
      throw new Error(
        "Lien PDF Drive manquant."
      );
    }

    window.currentPdfYayaUrl =
      r.url;

    if (
      typeof window.majBoutonPdf ===
      "function"
    ) {
      window.majBoutonPdf();
    }

    return {
      base64: depotPdfBase64,
      nom: depotPdfNom,
      driveUrl: r.url
    };
  }

  // Remplacement global : l'autosave construit le PDF une seule fois
  // et le conserve pour le dépôt.
  window.genererPdfPourYaya =
    fabriquerPdfEtSauverDrive;

  async function assurerPdfCourant() {
    if (
      depotPdfBase64 &&
      depotPdfNom &&
      window.currentPdfYayaUrl
    ) {
      return {
        base64: depotPdfBase64,
        nom: depotPdfNom,
        driveUrl:
          window.currentPdfYayaUrl
      };
    }

    return fabriquerPdfEtSauverDrive();
  }

  async function synchroniserYayaReel(
    chantierId,
    lien
  ) {
    if (!window.currentDocId) {
      throw new Error(
        "Document non sauvegardé."
      );
    }

    if (!chantierId) {
      throw new Error(
        "Choisissez le chantier Yaya."
      );
    }

    const api =
      String(
        window.YAYA_API || ""
      );

    if (!api) {
      throw new Error(
        "API Yaya introuvable."
      );
    }

    const response =
      await fetch(
        api +
        "?_ts=" +
        Date.now(),
        {
          method: "GET",
          cache: "no-store"
        }
      );

    if (!response.ok) {
      throw new Error(
        "Lecture Yaya HTTP " +
        response.status
      );
    }

    const json =
      await response.json();

    const state =
      json && json.data
        ? json.data
        : json;

    const chantiers =
      Array.isArray(
        state &&
        state.chantiers
      )
        ? state.chantiers
        : [];

    const chantier =
      chantiers.find(
        function (c) {
          return (
            String(c.id || "") ===
            String(chantierId)
          );
        }
      );

    if (!chantier) {
      throw new Error(
        "Chantier Yaya introuvable."
      );
    }

    const documents =
      Array.isArray(
        state.documents
      )
        ? state.documents.slice()
        : [];

    const id =
      "DOCS-CHANTIER-" +
      window.currentDocId;

    const documentYaya = {
      id: id,
      chantierId:
        String(chantier.id),
      type:
        typeof window.getDocumentTypeLabel === "function"
          ? window.getDocumentTypeLabel()
          : "Document",
      titre:
        String(
          window.currentTitre ||
          (
            typeof window.getDocumentTypeLabel === "function"
              ? window.getDocumentTypeLabel()
              : "Document chantier"
          )
        ),
      sujet:
        String(
          window.currentOperateur ||
          "Document chantier"
        ),
      date:
        dateDepotISO(),
      lien:
        String(
          lien ||
          window.currentPdfYayaUrl ||
          ""
        ),
      origine:
        "DOCS_CHANTIER",
      publieDocument:
        true
    };

    if (!documentYaya.lien) {
      throw new Error(
        "Lien PDF Yaya manquant."
      );
    }

    const index =
      documents.findIndex(
        function (d) {
          return (
            String(d.id || "") ===
            id
          );
        }
      );

    if (index >= 0) {
      documents[index] =
        Object.assign(
          {},
          documents[index],
          documentYaya
        );
    } else {
      documents.unshift(
        documentYaya
      );
    }

    const save =
      await fetch(
        api,
        {
          method: "POST",
          headers: {
            "Content-Type":
              "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            action: "setDocuments",
            data: documents
          })
        }
      );

    let result = {};
    try {
      result =
        await save.json();
    } catch (e) {}

    if (
      !save.ok ||
      result.ok === false
    ) {
      throw new Error(
        result.error ||
        "Écriture Yaya impossible."
      );
    }

    // Vérification par relecture.
    const verifRep =
      await fetch(
        api +
        "?_ts=" +
        Date.now(),
        {
          method: "GET",
          cache: "no-store"
        }
      );

    if (!verifRep.ok) {
      throw new Error(
        "Vérification Yaya impossible."
      );
    }

    const verifJson =
      await verifRep.json();

    const verifState =
      verifJson && verifJson.data
        ? verifJson.data
        : verifJson;

    const confirme =
      Array.isArray(
        verifState.documents
      ) &&
      verifState.documents.some(
        function (d) {
          return (
            String(d.id || "") ===
            id
          );
        }
      );

    if (!confirme) {
      throw new Error(
        "Yaya a répondu OK mais le document n'est pas retrouvé après relecture."
      );
    }

    return {
      ok: true,
      id: id,
      chantier:
        chantierLabel(chantier)
    };
  }

  /*
   * Neutralise l'appel automatique présent dans sauvegarderAuto().
   * Le vrai envoi ne s'exécute qu'au clic "Déposer".
   */
  window.synchroniserDocumentYaya =
    async function (
      chantierId,
      lien
    ) {
      if (!autoriserSyncYaya) {
        return {
          ok: true,
          skipped: true
        };
      }

      return synchroniserYayaReel(
        chantierId,
        lien
      );
    };

  async function confirmerDepot() {
    const bouton =
      document.getElementById(
        "depot-pdf-confirm"
      );

    const info =
      document.getElementById(
        "depot-pdf-info"
      );

    const doitYaya =
      destinationCourante === "yaya" ||
      destinationCourante ===
        "yaya_dropbox";

    const doitDropbox =
      destinationCourante === "dropbox" ||
      destinationCourante ===
        "yaya_dropbox";

    const chantierId =
      document.getElementById(
        "depot-yaya-chantier"
      ).value;

    if (
      doitYaya &&
      !chantierId
    ) {
      info.innerHTML =
        '<span class="depot-result-err">Choisissez le chantier Yaya.</span>';
      return;
    }

    if (
      doitDropbox &&
      !depotDropboxId
    ) {
      info.innerHTML =
        '<span class="depot-result-err">Choisissez le dossier Dropbox.</span>';
      return;
    }

    bouton.disabled = true;
    bouton.textContent =
      "Dépôt…";

    info.textContent =
      "Préparation du PDF…";

    let pdf;

    try {
      pdf =
        await assurerPdfCourant();
    } catch (e) {
      bouton.disabled = false;
      bouton.textContent =
        "Déposer";

      info.innerHTML =
        '<span class="depot-result-err">' +
        html(e.message) +
        "</span>";

      return;
    }

    const resultats = {
      yaya: null,
      dropbox: null
    };

    let lienDropbox = "";

    if (doitDropbox) {
      info.textContent =
        "Dépôt Dropbox…";

      try {
        resultats.dropbox =
          await appelerABDocs(
            "DEPOSER_PDF_DROPBOX",
            {
              base64:
                pdf.base64,
              nom:
                pdf.nom,
              destinationDropboxId:
                depotDropboxId,
              documentId:
                String(
                  window.currentDocId
                )
            }
          );

        lienDropbox =
          String(
            resultats.dropbox.lien ||
            ""
          );
      } catch (e) {
        resultats.dropbox = {
          ok: false,
          error: e.message
        };
      }
    }

    if (doitYaya) {
      info.textContent =
        "Dépôt Yaya…";

      try {
        autoriserSyncYaya = true;

        resultats.yaya =
          await window
            .synchroniserDocumentYaya(
              chantierId,
              // Pour un dépôt combiné, Yaya pointe vers Dropbox
              // si le lien est disponible ; sinon vers le PDF Drive.
              lienDropbox ||
              pdf.driveUrl
            );
      } catch (e) {
        resultats.yaya = {
          ok: false,
          error: e.message
        };
      } finally {
        autoriserSyncYaya = false;
      }
    }

    bouton.disabled = false;
    bouton.textContent =
      "Déposer";

    const lignes = [];

    if (doitYaya) {
      if (
        resultats.yaya &&
        resultats.yaya.ok
      ) {
        lignes.push(
          '<span class="depot-result-ok">✓ Yaya : déposé</span>'
        );
      } else {
        lignes.push(
          '<span class="depot-result-err">✕ Yaya : ' +
          html(
            resultats.yaya &&
            resultats.yaya.error ||
            "échec"
          ) +
          "</span>"
        );
      }
    }

    if (doitDropbox) {
      if (
        resultats.dropbox &&
        resultats.dropbox.ok
      ) {
        lignes.push(
          '<span class="depot-result-ok">✓ Dropbox : déposé</span>'
        );
      } else {
        lignes.push(
          '<span class="depot-result-err">✕ Dropbox : ' +
          html(
            resultats.dropbox &&
            resultats.dropbox.error ||
            "échec"
          ) +
          "</span>"
        );
      }
    }

    info.innerHTML =
      lignes.join("<br>");

    const toutOk =
      (!doitYaya ||
        (
          resultats.yaya &&
          resultats.yaya.ok
        )) &&
      (!doitDropbox ||
        (
          resultats.dropbox &&
          resultats.dropbox.ok
        ));

    if (toutOk) {
      setTimeout(
        fermerDepot,
        1200
      );
    }
  }

  function adapterBoutonPrincipal() {
    const bouton =
      document.querySelector(
        ".btn-share"
      );

    if (!bouton) return;

    bouton.textContent =
      "📤 Déposer le PDF";

    bouton.removeAttribute(
      "onclick"
    );

    bouton.onclick =
      ouvrirDepot;
  }

  function initialiser() {
    creerStylesEtModale();
    adapterBoutonPrincipal();
  }

  if (
    document.readyState ===
    "loading"
  ) {
    document.addEventListener(
      "DOMContentLoaded",
      initialiser
    );
  } else {
    initialiser();
  }

  // Exposé uniquement pour diagnostic manuel.
  window.ouvrirDepotDocument =
    ouvrirDepot;
})();
