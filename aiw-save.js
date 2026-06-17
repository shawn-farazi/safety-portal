/* ============================================================
   ADVANCED IRONWORKS - SAFETY PORTAL
   Shared "deliver a PDF" helper used by every form.
   Tries a one-tap OneDrive upload first; if OneDrive is not
   configured / sign-in is refused / it errors, it falls back
   to the iPad Share Sheet so the form is never a dead end.
   ------------------------------------------------------------
   aiwDeliverPdf({ blob, filename, formKey, shareText,
                   statusEl, onUploaded })
   ============================================================ */
window.aiwDeliverPdf = function (opts) {
  var blob = opts.blob;
  var file = new File([blob], opts.filename, { type: "application/pdf" });
  var status = opts.statusEl || null;
  function setStatus(t) { if (status) status.textContent = t; }

  function shareFallback() {
    try {
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        return navigator.share({ files: [file], title: file.name, text: opts.shareText || file.name })
          .then(function () { setStatus("PDF ready — choose a folder from the Share Sheet."); if (opts.onUploaded) opts.onUploaded(false); })
          .catch(function () { setStatus("Save cancelled."); });
      }
    } catch (e) {}
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
    setStatus("PDF downloaded — move it to the correct folder.");
    if (opts.onUploaded) opts.onUploaded(false);
    return Promise.resolve();
  }

  if (window.AIWGraph && AIWGraph.isConfigured()) {
    setStatus("Saving to OneDrive…");
    return AIWGraph.uploadPdf(blob, opts.filename, opts.formKey)
      .then(function (r) {
        setStatus("✓ Saved to OneDrive  ›  " + r.folder);
        if (opts.onUploaded) opts.onUploaded(true, r);
      })
      .catch(function (e) {
        console.error(e);
        setStatus("OneDrive save unavailable — opening Share Sheet…");
        return shareFallback();
      });
  }
  return shareFallback();
};

/* ============================================================
   aiwRunSave — the orchestrator every form's Save button calls.
   CRITICAL ORDER: it secures the OneDrive token FIRST (while the
   Save tap is still a fresh user gesture, so any sign-in popup is
   allowed), and ONLY THEN builds the PDF. Building first — which
   can take 15-20s on long forms — used to let the gesture expire,
   so iOS blocked the sign-in popup and the save fell back to the
   "browse to a folder" Share Sheet. This prevents that.
   ------------------------------------------------------------
   aiwRunSave({ build () -> Promise<Blob>, filename, formKey,
                shareText, statusEl, onUploaded })
   ============================================================ */
window.aiwRunSave = function (opts) {
  var status = opts.statusEl || null;
  function set(t) { if (status) status.textContent = t; }
  function buildAndDeliver() {
    set("Creating PDF…");
    return Promise.resolve(opts.build()).then(function (blob) {
      return aiwDeliverPdf({
        blob: blob, filename: opts.filename, formKey: opts.formKey,
        shareText: opts.shareText, statusEl: status, onUploaded: opts.onUploaded
      });
    }).catch(function (e) {
      console.error(e);
      set((e && e.message) || "Could not create the PDF.");
    });
  }
  if (window.AIWGraph && AIWGraph.isConfigured()) {
    set("Connecting to OneDrive…");
    return AIWGraph.ensureToken().then(buildAndDeliver, function () {
      set('Sign-in needed — tap "Sign in" at the top, then Save again.');
    });
  }
  return buildAndDeliver();
};
