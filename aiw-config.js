/* ============================================================
   ADVANCED IRONWORKS - SAFETY PORTAL CONFIG
   ------------------------------------------------------------
   THIS IS THE ONLY FILE YOU NEED TO EDIT after the Microsoft
   app registration is created. Fill in the three values marked
   "FILL IN". Everything else can stay as-is.
   ============================================================ */
window.AIW_CONFIG = {

  /* 1) FILL IN - Application (client) ID from the Entra app registration.
        Looks like: 11111111-2222-3333-4444-555555555555            */
  clientId: "PASTE_CLIENT_ID_HERE",

  /* 2) FILL IN - Your Microsoft 365 Directory (tenant) ID, also from
        the app registration overview page. Using the tenant ID (instead
        of "organizations") keeps sign-in locked to your company only.  */
  tenantId: "PASTE_TENANT_ID_HERE",

  /* 3) Redirect URI. This is filled in automatically from wherever the
        app is hosted, using the blank.html helper page. The SAME url
        must be added as a "Single-page application" redirect URI in the
        app registration. Leave this as-is unless told otherwise.       */
  redirectUri: (location.origin + location.pathname).replace(/[^/]*$/, "") + "blank.html",

  /* Permissions the app asks for. Files.ReadWrite lets it save PDFs to
     the signed-in user's OneDrive. User.Read shows who is signed in.    */
  scopes: ["User.Read", "Files.ReadWrite"],

  /* Top-level OneDrive folder that holds everything. */
  baseFolder: "Safety Records",

  /* One subfolder per form type (your chosen layout). */
  folders: {
    meeting:     "Daily Safety Meetings",
    hazard:      "Hazard Assessments",
    accident:    "Accident Reports",
    orientation: "New Employee Orientation"
  }
};

/* Helper used by the app to know whether setup is finished. */
window.AIW_CONFIG.isConfigured = function () {
  return this.clientId &&
         this.clientId.indexOf("PASTE_") !== 0 &&
         this.tenantId &&
         this.tenantId.indexOf("PASTE_") !== 0;
};
