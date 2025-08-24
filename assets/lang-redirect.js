// Minimal language redirect based on browser language (not geo).
// Set your GitHub Pages project base path:
const BASE = "/wsus-powershell-reports/"; // <- keep this in sync with your repo name

(function () {
  try {
    // Do not redirect if user is already on FR/ES, or if there's a hash (anchor navigation)
    if (location.pathname.startsWith(BASE + "fr/")) return;
    if (location.pathname.startsWith(BASE + "es/")) return;
    if (location.hash && location.hash.length > 1) return;

    const lang = (navigator.language || navigator.userLanguage || "en").toLowerCase();
    const isFR = lang.startsWith("fr");
    const isES = lang.startsWith("es");

    // Map current EN path to localized equivalent if it exists
    const pathAfterBase = location.pathname.replace(BASE, "");
    const localized = (pref) => {
      // Try to mirror current path inside /fr/ or /es/
      // If current path is "" or "index.html", go to /<pref>/
      if (pathAfterBase === "" || pathAfterBase === "index.html") return BASE + pref + "/";
      return BASE + pref + "/" + pathAfterBase;
    };

    if (isFR) {
      location.replace(localized("fr"));
      return;
    }
    if (isES) {
      location.replace(localized("es"));
      return;
    }
    // default: stay EN
  } catch (e) {
    // no-op
  }
})();
