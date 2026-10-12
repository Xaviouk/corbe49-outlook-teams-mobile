(function () {
  "use strict";
  var CLIENT_ID = "41d79a73-3bbc-47b3-a307-6edb15af1d47";
  var TENANT_ID = "75b354dc-0810-4caa-90a3-ac64ef88c77a";
  var REDIRECT_URI = "https://xaviouk.github.io/corbe49-outlook-teams-mobile/auth.html";
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
   function showAccount(account) {
    loginButton.hidden = true;
    logoutButton.hidden = false;
    show("Connexion Microsoft réussie.\nCompte : " + (account.username || account.name || "compte connecté") + "\n\nCette étape valide uniquement l'authentification. L'envoi vers Teams n'est pas encore activé.");
  }
   var startup = Promise.resolve(typeof msalApp.initialize === "function" ? msalApp.initialize() : undefined)
    .then(function () { return msalApp.handleRedirectPromise(); });
   var watchdog = window.setTimeout(function () {
    show("L'initialisation prend trop de temps. La bibliothèque Microsoft est chargée, mais son démarrage ne répond pas. Recharge avec Ctrl + F5. Si cela recommence, envoie-moi cette capture.");
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
        showAccount(account);
      } else {
        loginButton.disabled = false;
        loginButton.textContent = "Se connecter avec Microsoft";
        show("Prêt. Clique sur « Se connecter avec Microsoft » pour tester la connexion.");
      }
    })
    .catch(function (error) {
      window.clearTimeout(watchdog);
      loginButton.disabled = false;
      loginButton.textContent = "Réessayer la connexion";
      show("La connexion n'a pas abouti.\n" + (error && error.message ? error.message : String(error)) + "\n\nURI de redirection attendue :\n" + REDIRECT_URI);
    });
   loginButton.addEventListener("click", function () {
    loginButton.disabled = true;
    show("Redirection vers la page de connexion Microsoft…");
    msalApp.loginRedirect({ scopes: ["openid", "profile", "email", "User.Read"] });
  });
   logoutButton.addEventListener("click", function () {
    var account = msalApp.getActiveAccount();
    msalApp.logoutRedirect({ account: account, postLogoutRedirectUri: REDIRECT_URI });
  });
})();
