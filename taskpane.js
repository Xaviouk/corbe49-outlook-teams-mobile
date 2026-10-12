/* CORBE49 — Transfert Outlook vers Teams
 * This file reads the selected Outlook message only.
 * Sending to Teams stays disabled until Microsoft Entra authentication,
 * delegated Graph permissions, and attachment handling are configured.
 */
(function () {
  "use strict";

  function setText(id, value) {
    var element = document.getElementById(id);
    if (element) element.textContent = value || "—";
  }

  function showStatus(message) {
    setText("mail-state", message);
  }

  function renderMessage(item) {
    if (!item) {
      showStatus("Aucun message disponible");
      setText("setup-note", "Ouvre un e-mail dans Outlook, puis lance le complément depuis ce message.");
      return;
    }

    setText("mail-subject", item.subject || "(Sans objet)");

    var from = item.from || item.sender;
    var fromText = from && from.emailAddress
      ? [from.displayName, from.emailAddress].filter(Boolean).join(" — ")
      : (from && from.displayName) || "Expéditeur non disponible";
    setText("mail-from", fromText);

    var attachments = Array.isArray(item.attachments) ? item.attachments : [];
    setText("mail-attachments", attachments.length
      ? attachments.length + " pièce(s) jointe(s) détectée(s)"
      : "Aucune pièce jointe détectée dans les informations disponibles");

    if (item.body && typeof item.body.getAsync === "function" && window.Office && Office.AsyncResultStatus) {
      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setText("mail-body", result.value || "(Message vide)");
        } else {
          setText("mail-body", "Outlook n'a pas autorisé la lecture du contenu dans ce contexte.");
        }
      });
    } else {
      setText("mail-body", "L'aperçu du corps du message n'est pas disponible dans ce contexte Outlook.");
    }

    showStatus("Message détecté");
  }

  if (window.Office) {
    Office.onReady(function (info) {
      if (info.host !== Office.HostType.Outlook) {
        showStatus("À ouvrir dans Outlook");
        return;
      }
      renderMessage(Office.context && Office.context.mailbox
        ? Office.context.mailbox.item
        : null);
    });
  } else {
    showStatus("Bibliothèque Outlook indisponible");
  }
})();