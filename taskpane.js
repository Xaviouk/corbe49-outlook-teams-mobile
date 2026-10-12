/* CORBE49 — Transfert Outlook vers Teams
 * Authentification NAA et lecture des conversations : aucun envoi de message ici.
 */
(function () {
  "use strict";

  var CLIENT_ID = "41d79a73-3bbc-47b3-a307-6edb15af1d47";
  var TENANT_ID = "75b354dc-0810-4caa-90a3-ac64ef88c77a";
  var SCOPES = ["User.Read", "Chat.ReadBasic", "Chat.Create", "ChatMessage.Send"];
  var msalApp = null;
  var mailboxItem = null;

  function el(id) { return document.getElementById(id); }
  function setText(id, value) {
    var node = el(id);
    if (node) node.textContent = value || "—";
  }
  function status(message) { setText("setup-note", message); }

  function renderMessage(item) {
    if (!item) {
      setText("mail-state", "Aucun message disponible");
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
        setText("mail-body", result.status === Office.AsyncResultStatus.Succeeded
          ? (result.value || "(Message vide)")
          : "Outlook n'a pas autorisé la lecture du contenu dans ce contexte.");
      });
    } else {
      setText("mail-body", "L'aperçu du corps du message n'est pas disponible dans ce contexte Outlook.");
    }
    setText("mail-state", "Message détecté");
  }

  async function initAuth() {
    if (msalApp) return msalApp;
    if (!window.msal || typeof window.msal.createNestablePublicClientApplication !== "function") {
      throw new Error("La bibliothèque MSAL compatible avec l'authentification imbriquée n'a pas été chargée.");
    }
    msalApp = await window.msal.createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + TENANT_ID
      },
      cache: { cacheLocation: "localStorage" }
    });
    return msalApp;
  }

  async function getToken(interactive) {
    var app = await initAuth();
    var accounts = app.getAllAccounts();
    var request = { scopes: SCOPES };
    if (accounts && accounts.length) request.account = accounts[0];
    try {
      if (request.account) {
        var silent = await app.acquireTokenSilent(request);
        return silent.accessToken;
      }
      if (!interactive) throw new Error("Connecte-toi à Microsoft pour afficher tes conversations Teams.");
      var login = await app.loginPopup(request);
      if (login && login.account) request.account = login.account;
      var token = await app.acquireTokenSilent(request);
      return token.accessToken;
    } catch (err) {
      if (!interactive) throw err;
      if (err && (err.name === "InteractionRequiredAuthError" || err.errorCode === "interaction_required" ||
          err.errorCode === "consent_required" || err.errorCode === "login_required")) {
        var popupToken = await app.acquireTokenPopup(request);
        return popupToken.accessToken;
      }
      throw err;
    }
  }

  async function loadChats(interactive) {
    var button = el("connect-button");
    var select = el("chat-select");
    if (button) { button.disabled = true; button.textContent = "Connexion en cours…"; }
    if (select) select.disabled = true;
    status("Connexion sécurisée à Microsoft et lecture des conversations Teams…");
    try {
      var accessToken = await getToken(interactive);
      var response = await fetch("https://graph.microsoft.com/v1.0/me/chats?$top=50&$select=id,topic,chatType,lastUpdatedDateTime", {
        headers: { Authorization: "Bearer " + accessToken }
      });
      if (!response.ok) {
        var detail = await response.text();
        throw new Error("Microsoft Graph a répondu " + response.status + ". " + detail);
      }
      var data = await response.json();
      var chats = Array.isArray(data.value) ? data.value : [];
      if (select) {
        select.innerHTML = "";
        var first = document.createElement("option");
        first.value = "";
        first.textContent = chats.length ? "Sélectionne une conversation…" : "Aucune conversation retournée";
        select.appendChild(first);
        chats.forEach(function (chat) {
          var option = document.createElement("option");
          option.value = chat.id;
          var title = chat.topic || (chat.chatType === "oneOnOne" ? "Conversation individuelle" :
            chat.chatType === "group" ? "Conversation de groupe" : "Conversation Teams");
          option.textContent = title + " · " + (chat.chatType || "chat");
          select.appendChild(option);
        });
        select.disabled = chats.length === 0;
      }
      if (el("destination-mode")) el("destination-mode").disabled = false;
      if (el("people")) el("people").disabled = false;
      setText("auth-state", "Connecté à Microsoft");
      if (button) { button.disabled = false; button.textContent = "Actualiser les conversations"; }
      status("Accès Teams vérifié dans le complément Outlook. " + chats.length +
        " conversation(s) récupérée(s) (maximum 50). Aucun message n'a été envoyé ; le bouton de transfert reste désactivé.");
    } catch (err) {
      if (button) { button.disabled = false; button.textContent = "Se connecter à Teams"; }
      status("Impossible de charger les conversations Teams. " +
        (err && err.message ? err.message : String(err)) +
        " Aucun message n'a été envoyé.");
    }
  }

  function initUI() {
    var button = el("connect-button");
    if (button) button.addEventListener("click", function () { loadChats(true); });
    var mode = el("destination-mode");
    if (mode) mode.addEventListener("change", function () {
      var existing = mode.value === "existing";
      if (el("chat-select")) el("chat-select").hidden = !existing;
      if (el("people")) {
        el("people").hidden = existing;
        el("people").placeholder = "Saisis les adresses e-mail des participants, séparées par des points-virgules";
      }
    });
    var refresh = el("refresh-button");
    if (refresh) refresh.addEventListener("click", function () { loadChats(true); });
  }

  function start() {
    initUI();
    if (!window.Office) {
      setText("mail-state", "Office.js indisponible");
      status("Cette page doit être ouverte depuis Outlook pour lire l'e-mail sélectionné.");
      return;
    }
    Office.onReady(function (info) {
      if (info.host !== Office.HostType.Outlook) {
        setText("mail-state", "À ouvrir dans Outlook");
        status("Ouvre ce complément depuis un e-mail dans Outlook.");
        return;
      }
      mailboxItem = Office.context && Office.context.mailbox ? Office.context.mailbox.item : null;
      renderMessage(mailboxItem);
      status("Le message est prêt. Connecte-toi pour vérifier l'accès aux conversations Teams.");
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
  else start();
})();