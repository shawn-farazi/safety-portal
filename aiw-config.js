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
  clientId: "44feae02-5a01-4ce6-836f-2b5f5a68afc6",

  /* 2) FILL IN - Your Microsoft 365 Directory (tenant) ID, also from
        the app registration overview page. Using the tenant ID (instead
        of "organizations") keeps sign-in locked to your company only.  */
  tenantId: "577b1ab0-c97f-4f69-8858-c2efb831d108",

  /* 3) Redirect URI. This is filled in automatically from wherever the
        app is hosted, using the blank.html helper page. The SAME url
        must be added as a "Single-page application" redirect URI in the
        app registration. Leave this as-is unless told otherwise.       */
  redirectUri: (location.origin + location.pathname).replace(/[^/]*$/, "") + "blank.html",

  /* Permissions the app asks for.
     - Files.ReadWrite      : save to the signed-in user's own OneDrive
     - Files.ReadWrite.All  : save into a folder shared with the signed-in
                              user (lets the iPad/MTR write into Shawn's folder)
     - User.Read            : show who is signed in                          */
  scopes: ["User.Read", "Files.ReadWrite", "Files.ReadWrite.All"],

  /* ----------------------------------------------------------------
     SHARED DESTINATION (everything saves here, no matter who signs in)
     Points at Shawn's OneDrive > "09 SAFETY" folder.
     Leave targetFolderId blank ("") to instead save to each signed-in
     user's OWN OneDrive under baseFolder below.
     ---------------------------------------------------------------- */
  targetDriveId:  "b!oHQoAYGaAUi4wdJ-miC_1QXZm5HBZlNLv0v56E4_mmZxTzHDA8UaS7vC3RMWkKrP",
  targetFolderId: "01QCFSZVSJASRSUDHWZFDIFXAEYHHKU6WD",
  targetLabel:    "09 SAFETY",

  /* Used only when targetFolderId is blank (per-user OneDrive mode). */
  baseFolder: "Safety Records",

  /* One subfolder per form type, created inside the destination. */
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
