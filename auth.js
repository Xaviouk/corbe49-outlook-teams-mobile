(function () {
  "use strict";
  var CLIENT_ID = "41d79a73-3bbc-47b3-a307-6edb15af1d47";
  var TENANT_ID = "75b354dc-0810-4caa-90a3-ac64ef88c77a";
  var REDIRECT_URI = "https://xaviouk.github.io/corbe49-outlook-teams-mobile/auth.html";
  var GRAPH_SCOPES = ["User.Read", "Chat.ReadBasic", "Chat.Create", "ChatMessage.Send"];
  var status = document.getElementById("status");
  var loginButton = document.getElementById("login");
  var logoutButton = document.getElementById("logout");
  var msalApp;
  function show(message) { status.textContent = message; }

  if (!window.msal || !window.msal.PublicClientApplication) {
    show(window.msalLibraryFailed
      ? "La bibliothèque Microsoft n'a pas pu être téléchargée depuis le CDN. Vérifie le réseau ou le filtrage web."
      : "La bibliothèque Microsoft n'est pas disponible dans la page. Le script externe ne s'est peut-être pas chargé.");
    return;
  }

  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + TENANT_ID,
      redirectUri: REDIRECT_URI,
      postLogoutRedirectUri: REDIRECT_URI,
      navigateToLoginRequestUrl: false
    },
    cache: { cacheLocation: "sessionStorage" }
  });

  function checkGraph(account) {
    show("Connexion réussie. Vérification des autorisations Teams…");
    msalApp.acquireTokenSilent({ account: account, scopes: GRAPH_SCOPES })
      .then(function (tokenResult) {
        return fetch("https://graph.microsoft.com/v1.0/me/chats?$top=10&$select=id,topic,chatType", {
          headers: { "Authorization": "Bearer " + tokenResult.accessToken }
        });
      })
      .then(function (response) {
        if (!response.ok) {
          return response.text().then(function (detail) {
            throw new Error("Microsoft Graph a répondu " + response.status + ". " + detail);
          });
        }
        return response.json();
      })
      .then(function (data) {
        loginButton.hidden = true;
        logoutButton.hidden = false;
        var count = Array.isArray(data.value) ? data.value.length : 0;
        show("Connexion Microsoft réussie.\nCompte : " + (account.username || account.name || "compte connecté") +
          "\n\nAccès Microsoft Graph vérifié : la liste des conversations Teams répond correctement.\n" +
          "Conversations retournées dans ce contrôle : " + count + " (maximum 10).\n\n" +
          "Aucun message n'a été envoyé. L'intégration Outlook et les pièces jointes restent à réaliser.");
      })
      .catch(function (error) {
        loginButton.hidden = false;
        loginButton.disabled = false;
        loginButton.textContent = "Réessayer la connexion et les autorisations";
        logoutButton.hidden = false;
        show("La connexion Microsoft fonctionne, mais l'accès Teams n'a pas pu être vérifié.\n" +
          (error && error.message ? error.message : String(error)) +
          "\n\nSi une fenêtre de consentement n'est pas apparue, réessaie le bouton de connexion. Aucun message n'a été envoyé.");
      });
  }

  var startup = Promise.resolve(typeof msalApp.initialize === "function" ? msalApp.initialize() : undefined)
    .then(function () { return msalApp.handleRedirectPromise(); });

  var watchdog = window.setTimeout(function () {
    show("L'initialisation prend trop de temps. Recharge avec Ctrl + F5. Si cela recommence, envoie-moi cette capture.");
  }, 10000);

  startup.then(function (result) {
    window.clearTimeout(watchdog);
    var account = result && result.account;
    if (!account) {
      var accounts = msalApp.getAllAccounts();
      account = accounts.length ? accounts[0] : null;
    }
    if (account) {
      msalApp.setActiveAccount(account);
      checkGraph(account);
    } else {
      loginButton.disabled = false;
      loginButton.textContent = "Se connecter et autoriser l'accès à Teams";
      show("Prêt. Clique pour connecter Microsoft et demander les autorisations Teams nécessaires.");
    }
  }).catch(function (error) {
    window.clearTimeout(watchdog);
    loginButton.disabled = false;
    loginButton.textContent = "Réessayer la connexion";
    show("La connexion n'a pas abouti.\n" + (error && error.message ? error.message : String(error)) +
      "\n\nURI de redirection attendue :\n" + REDIRECT_URI);
  });

  loginButton.addEventListener("click", function () {
    loginButton.disabled = true;
    show("Redirection vers Microsoft pour confirmer l'accès Teams…");
    msalApp.loginRedirect({ scopes: GRAPH_SCOPES });
  });

  logoutButton.addEventListener("click", function () {
    var account = msalApp.getActiveAccount();
    msalApp.logoutRedirect({ account: account, postLogoutRedirectUri: REDIRECT_URI });
  });
})();
