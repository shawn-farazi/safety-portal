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
