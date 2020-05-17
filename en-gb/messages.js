//List of language-dependent messages: lang=en-GB

"use strict";

const tcTemplateMsg = {
  stationNameAlert: "Please insert a valid and unique station name",
  taskNameAlert: "Please insert a valid and unique task name",
  confirmDelete: "Are you sure you want to delete? This cannot be undone.",
  numberFieldError: "Must be a number",
  nonNegativeFieldError: "Must be ≥ 0",
  strictPositiveFieldError: "Must be > 0",
  naturalNumberFieldError: "Must be an integer ≥ 1",
  nameSyntax: "Must start with an alphanumeric character and then use only alphanumeric characters, spaces and ,.-_+=",
  notUnique: "Must be unique",
  numKitesRule: "IOF rule: number of kites = 6",
  notSameForAllStations: "Not the same for all stations",
  circle: "Circle",
  diameter: "diameter",
  sideLength: "side length",
  notSameForAllStationsRule: "IOF rule: must be the same for all stations",
  mapSizeRule: "IOF rules: 5 ≤ size ≤ 12; must be the same for all stations",
  mapSizeError: "0 &lt; size ≤ 12",
  mapScaleRule: "IOF rules: normally 1:4000 or 1:5000; must be the same for all stations",
  confirmRemoveTasks: (stationName) => "Some tasks will be removed from station " + stationName + ". Are you sure you want to delete these?",
  taskNameNotUnique: "Must be unique and neither <em>Kites</em> nor <em>VP</em>"
};
