//Applies headers and footers to each page. To be called as script in header to ensure viewport is set before page renders.

"use strict";

//Get nesting
const langFolder = "en-gb";
const pathCpts = location.pathname.split("/");
const langFolderStep = pathCpts.indexOf(langFolder);
let pathPrefix = "";
if (langFolderStep >= 0) {
  const stepsToTop = pathCpts.length - pathCpts.indexOf(langFolder) - 1;
  for (let stepId = 0; stepId < stepsToTop; stepId++) {
    pathPrefix += "../";
  }
}

//Header: viewport, stylesheet, favicons
document.head.innerHTML += "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />" +
"<link rel=\"stylesheet\" href=\"" + pathPrefix + "style.css\" />" +
"<link rel=\"apple-touch-icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/noborder152.png\" sizes=\"152x152\" />" +
"<link rel=\"icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/favicon512.png\" sizes=\"512x512\" />" +
"<link rel=\"icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/favicon256.png\" sizes=\"256x256\" />" +
"<link rel=\"icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/favicon192.png\" sizes=\"192x192\" />" +
"<link rel=\"icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/favicon32.png\" sizes=\"32x32\" />" +
"<link rel=\"icon\" type=\"image/png\" href=\"" + pathPrefix + "favicon/favicon16.png\" sizes=\"16x16\" />";

//Footer: modification date, copyright, license
addEventListener('DOMContentLoaded', () => {
  const footer = document.createElement("footer");
  const modDate = new Date(document.lastModified);
  footer.innerHTML = "<p class=\"right\">Last updated " +
  modDate.toLocaleString("en-GB", { dateStyle: "long", timeStyle: "long" }) + "</p>" +
  "<p class=\"right\">Copyright (c) 2017-20 TCTemplate contributors, maintained by Tom Dobra: <a href=\"" +
  pathPrefix + "license.xhtml\" target=\"help\" rel=\"license\">view license</a></p>";
  document.body.appendChild(footer);
});
