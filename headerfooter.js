//Applies headers and footers to each page. To be called as script in header to ensure viewport is set before page renders.
//Don't bother checking prerequisites, as it's not a huge problem if this script crashes

"use strict";

const lotoHF = (() => {
  //TODO: Bug in Chrome: defer script load doesn't work with XHTML, so load async and await domLoad promise
  const domLoad = new Promise((res) => { document.addEventListener("DOMContentLoaded", () => { res(); }, { once: true }); });
  // document.addEventListener("DOMContentLoaded", () => { domLoad = Promise.resolve(); }, { once: true });

  //Get nesting
  const pathCpts = location.pathname.split("/");
  const langFolderStep = pathCpts.indexOf(tottext.langFolder);
  let pathPrefix = "";
  if (langFolderStep >= 0) {
    const stepsToTop = pathCpts.length - pathCpts.indexOf(tottext.langFolder) - 1;
    for (let stepId = 0; stepId < stepsToTop; stepId++) {
      pathPrefix += "../";
    }
  }

  //Header: viewport, stylesheet, favicons
  const headFrag = document.createDocumentFragment();
  addElement("meta", headFrag, {
    name: "viewport",
    content: "width=device-width, initial-scale=1.0"
  });
  addElement("link", headFrag, {
    rel: "stylesheet",
    href: pathPrefix + "style.css"
  });
  addElement("link", headFrag, {
    ref: "apple-touch-icon",
    type: "image/png",
    href: pathPrefix + "favicon/noborder152.png"
  });
  ["512", "256", "192", "32", "16"].forEach((size) => {
    addElement("link", headFrag, {
      ref: "icon",
      type: "image/png",
      href: pathPrefix + "favicon/favicon" + size + ".png",
      sizes: size + "x" + size
    });
  });
  document.head.appendChild(headFrag);

  function addElement(elName, parentEl, params = {}) {
    const newEl = document.createElement(elName);
    Object.keys(params).forEach((key) => { newEl[key] = params[key]; });
    parentEl.appendChild(newEl);
    return newEl;
  }

  return {
    domLoad: domLoad,
    addElement: addElement,
    pathPrefix: pathPrefix,
    modDate: new Date(document.lastModified)
  };
})();

(async () => {
  //Footer: modification date, copyright, license
  await lotoHF.domLoad;
  const footer = document.createElement("footer");
  let pEl = lotoHF.addElement("p", footer);
  pEl.classList.add("right");
  pEl.textContent = tottext.lastUpdated + " " + lotoHF.modDate.toLocaleString("en-GB", { dateStyle: "long", timeStyle: "long" });
  pEl = lotoHF.addElement("p", footer);
  pEl.classList.add("right");
  pEl.textContent = tottext.copyright + " © 2017–" + lotoHF.modDate.getFullYear() + " " + tottext.copyrightName + " ";
  const aEl = lotoHF.addElement("a", pEl, { href: lotoHF.pathPrefix + "license.xhtml", target: "help", rel: "license" });
  aEl.textContent = tottext.license;
  document.body.appendChild(footer);
})();
