/* ============================================================
   ADVANCED IRONWORKS - SAFETY PORTAL
   Drop-in OneDrive connection / sign-in bar for any form.
   Usage: put <div id="aiwConnBar" class="no-print"></div> near the
   top of the page and include this script after aiw-graph.js.
   Shows connection status and a Sign in / Sign out button so staff
   can connect from any form, not just the dashboard.
   ============================================================ */
(function () {
  function injectStyles() {
    if (document.getElementById("aiwcb-style")) return;
    var s = document.createElement("style");
    s.id = "aiwcb-style";
    s.textContent =
      ".aiwcb-card{max-width:1100px;margin:10px auto;display:flex;align-items:center;gap:12px;background:#fff;border:1px solid #d9dee7;border-radius:14px;padding:10px 14px;box-shadow:0 1px 6px rgba(0,0,0,.06);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif}" +
      ".aiwcb-dot{width:13px;height:13px;border-radius:50%;background:#c2c9d6;flex:0 0 auto}" +
      ".aiwcb-dot.on{background:#0f7b45}.aiwcb-dot.warn{background:#b45309}" +
      ".aiwcb-txt{flex:1;min-width:140px;font-size:14px;color:#243b53;font-weight:700}" +
      ".aiwcb-btn{border:0;border-radius:10px;padding:10px 16px;font-weight:800;font-size:15px;min-height:44px;cursor:pointer;background:#0b5fff;color:#fff}" +
      ".aiwcb-btn.ghost{background:#e9eef7;color:#0b2f6b}" +
      "@media print{#aiwConnBar{display:none!important}}";
    document.head.appendChild(s);
  }

  function init() {
    var host = document.getElementById("aiwConnBar");
    if (!host) return;
    injectStyles();
    host.innerHTML =
      '<div class="aiwcb-card">' +
      '<span class="aiwcb-dot" id="aiwcbDot"></span>' +
      '<span class="aiwcb-txt" id="aiwcbTxt">Checking OneDrive…</span>' +
      '<button class="aiwcb-btn" id="aiwcbBtn" style="display:none">Sign in</button>' +
      '<button class="aiwcb-btn ghost" id="aiwcbOut" style="display:none">Sign out</button>' +
      "</div>";
    var dot = document.getElementById("aiwcbDot");
    var txt = document.getElementById("aiwcbTxt");
    var btn = document.getElementById("aiwcbBtn");
    var out = document.getElementById("aiwcbOut");

    function render() {
      if (!window.AIWGraph || !AIWGraph.isConfigured()) {
        dot.className = "aiwcb-dot warn";
        txt.textContent = "OneDrive not set up — forms fall back to the Share Sheet";
        btn.style.display = "none"; out.style.display = "none";
        return;
      }
      var a = AIWGraph.getAccount();
      if (a) {
        dot.className = "aiwcb-dot on";
        txt.textContent = "OneDrive connected — " + (a.username || a.name);
        btn.style.display = "none"; out.style.display = "inline-block";
      } else {
        dot.className = "aiwcb-dot";
        txt.textContent = "Not signed in to OneDrive — tap Sign in before saving";
        btn.textContent = "Sign in"; btn.style.display = "inline-block"; out.style.display = "none";
      }
    }

    btn.onclick = function () {
      txt.textContent = "Opening Microsoft sign-in…";
      AIWGraph.signIn().then(render).catch(function (e) {
        txt.textContent = "Sign-in didn't complete: " + ((e && e.message) || "try again"); render();
      });
    };
    out.onclick = function () { AIWGraph.signOut().then(render); };

    if (window.AIWGraph) {
      AIWGraph.onStatus(render);
      if (AIWGraph.isReady()) { AIWGraph.init().then(render).catch(render); } else { render(); }
    } else { render(); }
    setTimeout(render, 700);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
