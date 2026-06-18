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
      function popup() {
        return app.acquireTokenPopup({ scopes: CFG.scopes })
          .then(function (r) { app.setActiveAccount(r.account); notify(); return r.accessToken; });
      }
      if (acct) {
        // Silent renewal can hang or be blocked by Safari ITP — race it with a
        // timeout, then fall back to an interactive popup.
        var silent = app.acquireTokenSilent({ scopes: CFG.scopes, account: acct });
        var timeout = new Promise(function (_, rej) {
          setTimeout(function () { rej(new Error("silent_timeout")); }, 8000);
        });
        return Promise.race([silent, timeout])
          .then(function (r) { return r.accessToken; })
          .catch(function () {
            if (interactive === false) throw new Error("Sign-in required.");
            return popup();
          });
      }
      if (interactive === false) throw new Error("Not signed in.");
      return popup();
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

  function handleRes(res) {
    if (!res.ok) {
      return res.text().then(function (t) {
        throw new Error("Upload failed (" + res.status + "). " + t.slice(0, 300));
      });
    }
    return res.json();
  }

  /* ---- resolve the shared destination folder ------------- */
  var _target = null;
  function encPathStr(p) { return p.split("/").map(encodeURIComponent).join("/"); }
  function resolveTarget(token) {
    if (_target) return Promise.resolve(_target);
    // Preferred: explicit drive + folder IDs. Locks every save to ONE shared
    // folder (e.g. Shawn's), so anyone signed in (incl. MTR) writes there —
    // not into their own same-named folder. No lookup needed.
    if (CFG.targetDriveId && CFG.targetFolderId) {
      _target = { driveId: CFG.targetDriveId, itemId: CFG.targetFolderId };
      return Promise.resolve(_target);
    }
    var path = CFG.targetFolderPath;
    var segs = path.split("/");
    var leaf = segs[segs.length - 1];
    var parent = segs.slice(0, -1).join("/");
    var H = { "Authorization": "Bearer " + token };
    // 1) the signed-in user's OWN OneDrive (the folder's owner, e.g. Shawn)
    return fetch(GRAPH + "/me/drive/root:/" + encPathStr(path), { headers: H })
      .then(function (res) {
        if (res.ok) {
          return res.json().then(function (it) {
            _target = { driveId: it.parentReference.driveId, itemId: it.id };
            return _target;
          });
        }
        // 2) the destination folder SHARED with the signed-in user (e.g. MTR)
        return fetch(GRAPH + "/me/drive/sharedWithMe", { headers: H })
          .then(function (r) { return r.json(); })
          .then(function (sw) {
            var m = (sw.value || []).filter(function (i) { return i.remoteItem && i.name === leaf; })[0];
            if (m && m.remoteItem) {
              _target = { driveId: m.remoteItem.parentReference.driveId, itemId: m.remoteItem.id };
              return _target;
            }
            // 3) not found -> create the leaf inside its parent in the user's own drive
            var childrenUrl = parent
              ? GRAPH + "/me/drive/root:/" + encPathStr(parent) + ":/children"
              : GRAPH + "/me/drive/root/children";
            return fetch(childrenUrl, {
              method: "POST",
              headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
              body: JSON.stringify({ name: leaf, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
            }).then(handleRes).then(function (it) {
              _target = { driveId: it.parentReference.driveId, itemId: it.id };
              return _target;
            });
          });
      });
  }

  /* ---- upload -------------------------------------------- */
  function uploadPdf(blob, filename, formKey) {
    var folderName = (CFG.folders && CFG.folders[formKey]) || "Safety Records";
    return getToken(true).then(function (token) {
      var auth = { "Authorization": "Bearer " + token };

      // MODE 1: shared destination — everything into one fixed folder,
      // regardless of who is signed in.
      if (CFG.targetFolderId || CFG.targetFolderPath) {
        return resolveTarget(token).then(function (t) {
          var base = GRAPH + "/drives/" + t.driveId + "/items/" + t.itemId;
          // create the per-form subfolder if it isn't there yet (ignore 409)
          return fetch(base + "/children", {
            method: "POST",
            headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
            body: JSON.stringify({
              name: folderName, folder: {},
              "@microsoft.graph.conflictBehavior": "fail"
            })
          }).then(function () {
            var url = base + ":/" + encPath(folderName + "/" + filename) + ":/content";
            return fetch(url, {
              method: "PUT",
              headers: { "Authorization": "Bearer " + token, "Content-Type": "application/pdf" },
              body: blob
            });
          }).then(handleRes).then(function (item) {
            return {
              ok: true, webUrl: item.webUrl, name: item.name,
              folder: (CFG.targetLabel || "OneDrive") + " / " + folderName
            };
          });
        });
      }

      // MODE 2: per-user — save to the signed-in user's own OneDrive.
      var segments = [CFG.baseFolder, folderName];
      var fullPath = segments.join("/") + "/" + filename;
      return ensureFolders(token, segments).then(function () {
        var url = GRAPH + "/me/drive/root:/" + encPath(fullPath) + ":/content";
        return fetch(url, {
          method: "PUT",
          headers: { "Authorization": "Bearer " + token, "Content-Type": "application/pdf" },
          body: blob
        });
      }).then(handleRes).then(function (item) {
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
    ensureToken: function () { return getToken(true); },
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
