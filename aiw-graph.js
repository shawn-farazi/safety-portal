/* ============================================================
   ADVANCED IRONWORKS - SAFETY PORTAL
   OneDrive upload module (Microsoft Graph + MSAL)
   ------------------------------------------------------------
   Provides window.AIWGraph with:
     isConfigured()      -> true once aiw-config.js is filled in
     isReady()           -> MSAL loaded and configured
     getAccount()        -> signed-in account object or null
     signIn()            -> interactive sign-in (popup)
     signOut()
     uploadPdf(blob, filename, formKey)  -> uploads to OneDrive,
                            returns { ok, webUrl, folder } or throws
     onStatus(cb)        -> subscribe to "signed in / out" changes
   Requires MSAL browser library + aiw-config.js loaded first.
   ============================================================ */
(function () {
  "use strict";

  var CFG = window.AIW_CONFIG || {};
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var msalApp = null;
  var initPromise = null;
  var statusSubs = [];

  function cfgOk() {
    return !!(CFG && typeof CFG.isConfigured === "function" && CFG.isConfigured());
  }

  function notify() {
    var acct = getAccount();
    statusSubs.forEach(function (cb) { try { cb(acct); } catch (e) {} });
  }

  /* ---- MSAL bootstrap ------------------------------------- */
  function ensureMsal() {
    if (initPromise) return initPromise;
    initPromise = new Promise(function (resolve, reject) {
      if (!window.msal || !window.msal.PublicClientApplication) {
        return reject(new Error("Microsoft sign-in library did not load."));
      }
      if (!cfgOk()) {
        return reject(new Error("Safety Portal is not connected to OneDrive yet (config not filled in)."));
      }
      try {
        msalApp = new window.msal.PublicClientApplication({
          auth: {
            clientId: CFG.clientId,
            authority: "https://login.microsoftonline.com/" + CFG.tenantId,
            redirectUri: CFG.redirectUri
          },
          cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
        });
        // MSAL v3 needs initialize(); v2 does not have it. Handle both.
        var done = function () {
          var accts = msalApp.getAllAccounts();
          if (accts.length && !msalApp.getActiveAccount()) {
            msalApp.setActiveAccount(accts[0]);
          }
          resolve(msalApp);
        };
        if (typeof msalApp.initialize === "function") {
          msalApp.initialize().then(done).catch(reject);
        } else {
          done();
        }
      } catch (e) { reject(e); }
    });
    return initPromise;
  }

  function getAccount() {
    if (!msalApp) return null;
    var a = msalApp.getActiveAccount();
    if (a) return a;
    var all = msalApp.getAllAccounts();
    return all.length ? all[0] : null;
  }

  /* ---- token --------------------------------------------- */
  function getToken(interactive) {
    return ensureMsal().then(function (app) {
      var acct = getAccount();
      var req = { scopes: CFG.scopes, account: acct || undefined };
      if (acct) {
        return app.acquireTokenSilent(req)
          .then(function (r) { return r.accessToken; })
          .catch(function () {
            return app.acquireTokenPopup({ scopes: CFG.scopes })
              .then(function (r) { app.setActiveAccount(r.account); notify(); return r.accessToken; });
          });
      }
      if (interactive === false) throw new Error("Not signed in.");
      return app.acquireTokenPopup({ scopes: CFG.scopes })
        .then(function (r) { app.setActiveAccount(r.account); notify(); return r.accessToken; });
    });
  }

  /* ---- folder helpers ------------------------------------ */
  function encPath(path) {
    return path.split("/").map(encodeURIComponent).join("/");
  }

  // Create each folder segment if missing (ignore "already exists").
  function ensureFolders(token, segments) {
    var parent = "";
    var chain = Promise.resolve();
    segments.forEach(function (seg) {
      chain = chain.then(function () {
        var url = parent
          ? GRAPH + "/me/drive/root:/" + encPath(parent) + ":/children"
          : GRAPH + "/me/drive/root/children";
        return fetch(url, {
          method: "POST",
          headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
          body: JSON.stringify({
            name: seg, folder: {},
            "@microsoft.graph.conflictBehavior": "fail"
          })
        }).then(function () {
          parent = parent ? parent + "/" + seg : seg;
        });
      });
    });
    return chain;
  }

  /* ---- upload -------------------------------------------- */
  function uploadPdf(blob, filename, formKey) {
    var folderName = (CFG.folders && CFG.folders[formKey]) || "Safety Records";
    var segments = [CFG.baseFolder, folderName];
    var fullPath = segments.join("/") + "/" + filename;
    return getToken(true).then(function (token) {
      return ensureFolders(token, segments).then(function () {
        var url = GRAPH + "/me/drive/root:/" + encPath(fullPath) + ":/content";
        return fetch(url, {
          method: "PUT",
          headers: { "Authorization": "Bearer " + token, "Content-Type": "application/pdf" },
          body: blob
        });
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            throw new Error("Upload failed (" + res.status + "). " + t.slice(0, 300));
          });
        }
        return res.json();
      }).then(function (item) {
        return { ok: true, webUrl: item.webUrl, folder: segments.join(" / "), name: item.name };
      });
    });
  }

  /* ---- public API ---------------------------------------- */
  window.AIWGraph = {
    isConfigured: cfgOk,
    isReady: function () { return cfgOk() && !!(window.msal && window.msal.PublicClientApplication); },
    getAccount: getAccount,
    init: ensureMsal,
    signIn: function () {
      return ensureMsal().then(function (app) {
        return app.loginPopup({ scopes: CFG.scopes }).then(function (r) {
          app.setActiveAccount(r.account); notify(); return r.account;
        });
      });
    },
    signOut: function () {
      if (!msalApp) return Promise.resolve();
      var acct = getAccount();
      return msalApp.logoutPopup({ account: acct }).then(notify).catch(function () { notify(); });
    },
    uploadPdf: uploadPdf,
    onStatus: function (cb) { statusSubs.push(cb); }
  };

  // Warm up MSAL quietly so the signed-in chip is accurate on load.
  if (cfgOk() && window.msal) {
    ensureMsal().then(notify).catch(function () {});
  }
})();
