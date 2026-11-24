// Loaded before vendor/tex-svg-full.js on every page.
window.MathJax = {
  tex: {
    // Every extension is bundled; turn the common ones on. noundefined/noerrors
    // are off so a typo is reported instead of drawn in red into the email.
    packages: {
      "[+]": ["ams", "newcommand", "boldsymbol", "color", "cancel", "mathtools", "physics", "braket", "cases", "bbox", "unicode"],
      "[-]": ["noundefined", "noerrors", "require", "autoload"],
    },
  },
  svg: { fontCache: "none" }, // each picture self-contained, so it can become a PNG
  startup: { typeset: false },
};
