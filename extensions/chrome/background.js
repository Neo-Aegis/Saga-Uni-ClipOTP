const TRIGGER_URL = "http://localhost:18924/trigger";

chrome.webNavigation.onBeforeNavigate.addListener(
  (details) => {
    if (details.frameId === 0) {
      fetch(TRIGGER_URL, { method: "POST" }).catch(() => {});
    }
  },
  { url: [{ urlMatches: "ssoidp\\.cc\\.saga-u\\.ac\\.jp/idp/profile/SAML2/Redirect" }] }
);
