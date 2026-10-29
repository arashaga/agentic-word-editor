/** @file Apply URL-selected or system-preferred theme before the page is painted. */
/** Initialize data-theme without accessing storage or waiting for the main application. */
(() => {
  const param = new URLSearchParams(window.location.search).get("clawpilotTheme");
  const theme =
    param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  document.documentElement.setAttribute("data-theme", theme);
})();
