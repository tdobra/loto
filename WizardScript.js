//Put script in body of HTML

//Prevent sloppy programming and throw more errors
"use strict";

//Check browser supports required APIs
//Items to check:
//HTMLCanvasElement.toBlob() (Edge >= 79, Firefox >= 19, Safari >= 11)
//details (Edge >= 79, Firefox >= 49, Safari >= 6)
//async (Edge >= 15, Firefox >= 52, Safari >= 10.1) - no easy way to check and very few browswers to trip up
try {
  if (
    typeof document.createElement("canvas").toBlob !== "function" ||
    typeof document.createElement("details").open !== "boolean"
  ) {
    throw undefined;
  }
  document.getElementById("missingAPIs").hidden = true;   //Hide error message that shows by default
} catch (err) {
  document.getElementById("mainView").hidden = true;
}

document.getElementById("stationProperties").hidden = true;
document.getElementById("savePDF").hidden = true;
document.getElementById("viewLog").hidden = true;

//Keep all functions private and put those with events in HTML tags in a namespace
const tcTemplate = (() => {
  var pdfjsLib;

  //Keep track of whether an input file has been changed in a table to disable autosave
  var paramsSaved = true;

  //Layout default measurements
  const defaultLayout = {
    IDFontSize: 0.7,
    checkWidth: 1.5,
    checkHeight: 1.5,
    checkFontSize: 0.8,
    removeFontSize: 0.3,
    pointHeight: 2.5,
    letterFontSize: 1.8,
    phoneticFontSize: 0.6
  };

  function loadppen(fileInput) {
    //Reads a Purple Pen file
    var fileobj, freader;
    var courseOrderUsed = [];

    function getNodeByID(xmlDoc, tagName, ID) {
      //Returns course-control node with given ID
      var courseControlNodeSet, courseControlNumNodes, itemNum;
      courseControlNodeSet = xmlDoc.getElementsByTagName(tagName);
      courseControlNumNodes = courseControlNodeSet.length;
      for (itemNum = 0; itemNum < courseControlNumNodes; itemNum++) {
        if (courseControlNodeSet[itemNum].getAttribute("id") == ID) {
          return courseControlNodeSet[itemNum];
        }
      }
      return null;
    }

    fileobj = fileInput.files[0];
    if (!fileobj) { return; }	//Nothing selected

    const ppenStatusBox = document.getElementById("ppenStatus");
    freader = new FileReader();
    freader.onload = function () {
      var xmlParser, xmlobj, parsererrorNS, mapFileScale, globalScale, courseNodes, courseNodesId, courseNodesNum, tableRowNode, tableColNode, tableContentNode, layoutRowNode, selectOptionNode, existingRows, existingRowID, existingRow, otherNode, courseControlNode, controlNode, controlsSkipped, numProblems, stationNameRoot, courseScale;
      xmlParser = new DOMParser();
      xmlobj = xmlParser.parseFromString(freader.result, "text/xml");
      //Check XML is well-formed - see Rast on https://stackoverflow.com/questions/11563554/how-do-i-detect-xml-parsing-errors-when-using-javascripts-domparser-in-a-cross - will have a parsererror element somewhere in a namespace that depends on the browser
      try {
        //Reports error on console in Edge, but ok
        parsererrorNS = xmlParser.parseFromString("INVALID", "text/xml").getElementsByTagName("parsererror")[0].namespaceURI;
      } catch (err) {
        //Method of getting namespace is too harsh for browser; try no namespace
        parsererrorNS = "";
      }

      if (xmlobj.getElementsByTagNameNS(parsererrorNS, "parsererror").length > 0) {
        ppenStatusBox.innerHTML = "Could not read Purple Pen file: invalid XML.";
        return;
      }

      try {
        //Reset course table - keep the first row
        existingRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
        for (existingRowID = existingRows.length - 1; existingRowID > 0; existingRowID--) {
          document.getElementById("courseTableBody").removeChild(existingRows[existingRowID]);
        }

        //Reset layout table - keep the first two rows
        existingRows = document.getElementById("layoutTableBody").getElementsByTagName("tr");
        for (existingRowID = existingRows.length - 1; existingRowID > 1; existingRowID--) {
          document.getElementById("layoutTableBody").removeChild(existingRows[existingRowID]);
        }

        //Save map file scale
        otherNode = xmlobj.getElementsByTagName("map")[0];
        if (!otherNode) {
          ppenStatusBox.innerHTML = "Could not read map scale.";
          return;
        }
        mapFileScale = Number(otherNode.getAttribute("scale"));
        if (!(mapFileScale > 0)) {
          ppenStatusBox.innerHTML = "Could not read map scale.";
          return;
        }

        //Global print scale
        otherNode = xmlobj.getElementsByTagName("all-controls")[0];
        if (otherNode) {
          globalScale = Number(otherNode.getAttribute("print-scale"));
          if (!globalScale) {
            globalScale = Number(xmlobj.getElementsByTagName("map")[0].getAttribute("scale"));
            if (!globalScale) {
              globalScale = "";	//Scale could not be found
            }
          }
        }

        courseNodes = xmlobj.getElementsByTagName("course");
        courseNodesNum = courseNodes.length;

        //Special objects use <course> elements with no child element <name>. Remove these.
        //Iterate backwards as we are removing elements on the fly
        for (courseNodesId = courseNodesNum - 1; courseNodesId >= 0; courseNodesId--) {
          if (courseNodes[courseNodesId].getElementsByTagName("name").length == 0) {
            courseNodes[courseNodesId].parentNode.removeChild(courseNodes[courseNodesId]);
          }
        }
        courseNodes = xmlobj.getElementsByTagName("course");
        courseNodesNum = courseNodes.length;

        //Make list of all course order attributes used - to determine print page, so must be completed before rest of file reading
        for (courseNodesId = 0; courseNodesId < courseNodesNum; courseNodesId++) {
          courseOrderUsed.push(courseNodes[courseNodesId].getAttribute("order"));
        }
        courseOrderUsed.sort(function (a, b) { return a - b; });

        //Find courses with name *.1. xpath doesn't appear to be working in Safari, iterate over nodes.
        for (courseNodesId = 0; courseNodesId < courseNodesNum; courseNodesId++) {
          if (courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent.endsWith(".1")) {
            //Check course type = score
            if (courseNodes[courseNodesId].getAttribute("kind") != "score") {
              ppenStatusBox.innerHTML = "Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + " type must be set to score.";
              return;
            }

            //Check zero page margin
            //Portrait vs. landscape is irrelevant, as coordinates are determined by left and bottom attributes
            otherNode = courseNodes[courseNodesId].getElementsByTagName("print-area")[0];
            if (otherNode) {
              if (otherNode.getAttribute("page-margins") > 0) {
                ppenStatusBox.innerHTML = "The page margin must be set to 0 on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
                return;
              }
            } else {
              ppenStatusBox.innerHTML = "The page margin must be set to 0 on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
              return;
            }
            if (otherNode.getAttribute("automatic") == "true") {
              ppenStatusBox.innerHTML = "The print area selection must be set to manual on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
              return;
            }

            //Create new table row
            tableRowNode = document.createElement("tr");

            //Create first column - station name + hidden values
            tableColNode = document.createElement("td");
            stationNameRoot = courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent.slice(0, -2);
            tableContentNode = document.createElement("span");
            tableContentNode.className = "stationName";
            tableContentNode.innerHTML = stationNameRoot;
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);
            //Hidden: Course order
            tableContentNode = document.createElement("span");
            tableContentNode.className = "courseOrder"
            tableContentNode.hidden = true;
            tableContentNode.innerHTML = courseNodes[courseNodesId].getAttribute("order");
            tableColNode.appendChild(tableContentNode);

            //Second column - show station?
            tableColNode = document.createElement("td");
            tableRowNode.appendChild(tableColNode);
            tableContentNode = document.createElement("input");
            tableContentNode.type = "checkbox";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableContentNode.className = "showStation";
            tableContentNode.checked = true;
            tableColNode.appendChild(tableContentNode);

            //Third column - number of kites
            tableColNode = document.createElement("td");
            tableContentNode = document.createElement("input");
            tableContentNode.type = "number";
            tableContentNode.min = 1;
            tableContentNode.max = 6;
            tableContentNode.step = 1;
            tableContentNode.required = true;
            tableContentNode.className = "kites";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Fourth column - zeroes allowed?
            tableColNode = document.createElement("td");
            tableRowNode.appendChild(tableColNode);
            tableContentNode = document.createElement("input");
            tableContentNode.type = "checkbox";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableContentNode.className = "zeroes";
            tableColNode.appendChild(tableContentNode);

            //Fifth column - station heading
            tableColNode = document.createElement("td");
            tableContentNode = document.createElement("input");
            tableContentNode.type = "number";
            tableContentNode.step = "any";
            tableContentNode.required = true;
            tableContentNode.className = "heading";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createTextNode(" " + String.fromCharCode(176));
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Sixth column - map shape
            tableColNode = document.createElement("td");
            tableContentNode = document.createElement("select");
            tableContentNode.required = true;
            tableContentNode.className = "mapShape";
            selectOptionNode = document.createElement("option");
            selectOptionNode.text = "Circle";
            tableContentNode.add(selectOptionNode);
            selectOptionNode = document.createElement("option");
            selectOptionNode.text = "Square";
            tableContentNode.add(selectOptionNode);
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Seventh column - map size
            tableColNode = document.createElement("td");
            tableContentNode = document.createElement("input");
            tableContentNode.type = "number";
            tableContentNode.step = "any";
            tableContentNode.min = 0;
            tableContentNode.max = 12;
            tableContentNode.required = true;
            tableContentNode.className = "mapSize layoutLength";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createTextNode(" cm");
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Eighth column - map scale
            tableColNode = document.createElement("td");
            tableContentNode = document.createTextNode("1:");
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createElement("input");
            tableContentNode.type = "number";
            tableContentNode.step = "any";
            tableContentNode.min = 0;
            tableContentNode.required = true;
            tableContentNode.className = "mapScale";
            //Populate map scale
            courseScale = Number(courseNodes[courseNodesId].getElementsByTagName("options")[0].getAttribute("print-scale"));
            if (!courseScale) {
              courseScale = globalScale;
            }
            tableContentNode.value = courseScale;
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Nineth column - contour interval + hidden values
            tableColNode = document.createElement("td");
            tableContentNode = document.createElement("input");
            tableContentNode.type = "number";
            tableContentNode.step = "any";
            tableContentNode.min = 0;
            tableContentNode.required = true;
            tableContentNode.className = "contourInterval";
            tableContentNode.addEventListener("change", () => { paramsSaved = false; });
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createTextNode(" m");
            tableColNode.appendChild(tableContentNode);
            tableRowNode.appendChild(tableColNode);

            //Hidden: x and y coordinates of centre of circle on map, relative to bottom-left corner
            existingRowID = courseNodesId;
            numProblems = 1;
            //Create hidden spans to save data
            tableContentNode = document.createElement("span");
            tableContentNode.className = "circlex";
            tableContentNode.hidden = true;
            tableContentNode.innerHTML = "{";
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createElement("span");
            tableContentNode.className = "circley";
            tableContentNode.hidden = true;
            tableContentNode.innerHTML = "{";
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createElement("span");
            tableContentNode.className = "printPage";
            tableContentNode.hidden = true;
            tableContentNode.innerHTML = "{";
            tableColNode.appendChild(tableContentNode);
            tableContentNode = document.createElement("span");
            tableContentNode.className = "controlsSkipped";
            tableContentNode.hidden = true; //Don't add brackets to innerHTML
            tableColNode.appendChild(tableContentNode);
            //Loop while another control can be added to the station
            while (existingRowID < courseNodesNum) {
              const printNode = courseNodes[existingRowID].getElementsByTagName("print-area")[0];
              if (!printNode) {
                window.alert("Page setup not complete for Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0] + ".");
                return;
              }
              //Calculate positions of circles relative to bottom left of page in cm
              //Work in mm, changing to cm at the last moment
              //Cannot rely only on bottom and left coordinates because might be for wrong page size if scale has changed recently/multiple scales in use
              //Emulate PPen behaviour: calculates centre of page using left, right, ...; then scale for ratio between map and course scales
              //Get page dimensions in mm; saved in ppen file in hundredths of inch
              let pageHalfWidth, pageHalfHeight;
              if (printNode.getAttribute("page-landscape") === "true") {
                pageHalfWidth = printNode.getAttribute("page-height");
                pageHalfHeight = printNode.getAttribute("page-width");
              } else {
                pageHalfWidth = printNode.getAttribute("page-width");
                pageHalfHeight = printNode.getAttribute("page-height");
              }
              pageHalfWidth = Number(pageHalfWidth) * 0.127;
              pageHalfHeight = Number(pageHalfHeight) * 0.127;
              const pageCentreH = (Number(printNode.getAttribute("right")) + Number(printNode.getAttribute("left"))) / 2;
              const pageCentreV = (Number(printNode.getAttribute("top")) + Number(printNode.getAttribute("bottom"))) / 2;
              if (isNaN(pageCentreH) || isNaN(pageCentreV)) {
                window.alert("Page setup missing attributes for Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0] + ".");
                return;
              }
              courseControlNode = getNodeByID(xmlobj, "course-control", courseNodes[existingRowID].getElementsByTagName("first")[0].getAttribute("course-control"));
              controlNode = getNodeByID(xmlobj, "control", courseControlNode.getAttribute("control"));
              controlsSkipped = 0;    //Keep tabs on position of desired control in control descriptions
              while (controlNode.getAttribute("kind") != "normal") {
                controlsSkipped++;
                courseControlNode = courseControlNode.getElementsByTagName("next")[0];
                if (!courseControlNode) {
                  window.alert("No control added to Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0] + ".");
                  return;
                }
                courseControlNode = getNodeByID(xmlobj, "course-control", courseControlNode.getAttribute("course-control"));
                controlNode = getNodeByID(xmlobj, "control", courseControlNode.getAttribute("control"));
              }

              //Append comma to lists - always have a trailing comma
              //Circle position in cm with origin in bottom left corner; saved in mm relative to map file scale in PPen
              tableColNode.getElementsByClassName("circlex")[0].innerHTML += (0.1 * (pageHalfWidth + (Number(controlNode.getElementsByTagName("location")[0].getAttribute("x")) - pageCentreH) * mapFileScale / courseScale)).toString() + ",";
              tableColNode.getElementsByClassName("circley")[0].innerHTML += (0.1 * (pageHalfHeight + (Number(controlNode.getElementsByTagName("location")[0].getAttribute("y")) - pageCentreV) * mapFileScale / courseScale)).toString() + ",";

              //Read course order attribute, then find its position in list of course order values used. Adding one onto this gives the page number when all courses, except blank, are printed in a single PDF.
              tableColNode.getElementsByClassName("printPage")[0].innerHTML += (courseOrderUsed.indexOf(courseNodes[existingRowID].getAttribute("order")) + 1).toString() + ",";
              tableColNode.getElementsByClassName("controlsSkipped")[0].innerHTML += controlsSkipped + ",";

              //Find next control at station
              otherNode = stationNameRoot + "." + (numProblems + 1);
              for (existingRowID = 0; existingRowID < courseNodesNum; existingRowID++) {
                if (courseNodes[existingRowID].getElementsByTagName("name")[0].textContent == otherNode) {
                  //Found another control
                  numProblems++;
                  //Check new course type, margin and manual print selection
                  if (courseNodes[existingRowID].getAttribute("kind") != "score") {
                    ppenStatusBox.innerHTML = "Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + " type must be set to score.";
                    return;
                  }
                  //Check zero page margin
                  otherNode = courseNodes[existingRowID].getElementsByTagName("print-area")[0];
                  if (otherNode) {
                    if (otherNode.getAttribute("page-margins") > 0) {
                      ppenStatusBox.innerHTML = "The page margin must be set to 0 on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
                      return;
                    }
                  } else {
                    ppenStatusBox.innerHTML = "The page margin must be set to 0 on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
                    return;
                  }
                  if (otherNode.getAttribute("automatic") == "true") {
                    ppenStatusBox.innerHTML = "The print area selection must be set to manual on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.";
                    return;
                  }
                  //Check same scale
                  otherNode = courseNodes[existingRowID].getElementsByTagName("options")[0];
                  if (otherNode) {
                    if (Number(otherNode.getAttribute("print-scale")) != courseScale) {
                      ppenStatusBox.innerHTML = "The print scale is different on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ".";
                      return;
                    }
                  } else if (courseScale != globalScale) {
                    ppenStatusBox.innerHTML = "The print scale is different on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ".";
                    return;
                  }
                  break;
                }
              }
            }
            //Close array brackets
            tableColNode.getElementsByClassName("circlex")[0].innerHTML += "}";
            tableColNode.getElementsByClassName("circley")[0].innerHTML += "}";
            tableColNode.getElementsByClassName("printPage")[0].innerHTML += "}";
            //Save numProblems
            tableContentNode = document.createElement("span");
            tableContentNode.className = "numProblems";
            tableContentNode.hidden = true;
            tableContentNode.innerHTML = numProblems;
            tableColNode.appendChild(tableContentNode);

            //Populate rows in layout table
            //Create new table row
            layoutRowNode = document.createElement("tr");

            //Create first column - station name + hidden values
            tableColNode = document.createElement("td");
            tableColNode.innerHTML = stationNameRoot;
            layoutRowNode.appendChild(tableColNode);

            //Other columns
            for (otherNode in defaultLayout) {
              tableColNode = document.createElement("td");
              tableContentNode = document.createElement("input");
              tableContentNode.type = "number";
              tableContentNode.min = 0;
              tableContentNode.max = 29.7;
              tableContentNode.step = "any";
              tableContentNode.required = true;
              tableContentNode.value = defaultLayout[otherNode];  //Default value
              tableContentNode.className = otherNode + " layoutLength";
              tableContentNode.addEventListener("change", () => { paramsSaved = false; });
              tableColNode.appendChild(tableContentNode);
              layoutRowNode.appendChild(tableColNode);
            }

            //Insert row in correct position in tables for course order
            //existingRows is a list of course order spans, which are 2nd generation descendants of table rows
            existingRows = document.getElementById("courseTableBody").getElementsByClassName("courseOrder");
            for (existingRowID = 0; ; existingRowID++) {
              existingRowID++;
              //One header row in main table
              existingRow = existingRows[existingRowID + 1];
              if (!existingRow) {
                //Run out of rows, so create new row at end
                document.getElementById("courseTableBody").appendChild(tableRowNode);
                document.getElementById("layoutTableBody").appendChild(layoutRowNode);
                break;	//No more rows to consider
              }
              //Is the course order of this existing row greater than the new row? If yes, insert before. If no, iterate.
              if (Number(existingRow.innerHTML) > Number(courseNodes[courseNodesId].getAttribute("order"))) {
                document.getElementById("courseTableBody").insertBefore(tableRowNode, existingRow.parentElement.parentElement);
                //Need to update existingRow to match layout table. There are two header rows in the tbody.
                existingRow = document.getElementById("layoutTableBody").getElementsByTagName("tr")[existingRowID + 2];
                document.getElementById("layoutTableBody").insertBefore(layoutRowNode, existingRow);
                break;	//Current row needs to be inserted before existingRow
              }
            }
          }
        }
      } catch (err) {
        ppenStatusBox.innerHTML = "Error reading Purple Pen file: " + err;
        console.error(err);
        return;
      }

      //Reset update all stations fields
      document.getElementsByClassName("showStation")[0].indeterminate = true;
      document.getElementsByClassName("kites")[0].value = "";
      document.getElementsByClassName("zeroes")[0].indeterminate = true;
      document.getElementsByClassName("mapShape")[0].selectedIndex = 0;
      document.getElementsByClassName("mapSize")[0].value = "";
      document.getElementsByClassName("mapScale")[0].value = "";
      document.getElementsByClassName("contourInterval")[0].value = "";

      for (otherNode in defaultLayout) {
        document.getElementsByClassName(otherNode)[1].value = "";
      }

      //Disable debug mode
      tableContentNode = document.getElementById("debugCircle");
      tableContentNode.checked = false;
      tableContentNode.addEventListener("change", () => { paramsSaved = false; });

      //Prepare view
      tableContentNode = document.getElementById("stationProperties");
      tableContentNode.hidden = false;
      tableContentNode.scrollIntoView();
      ppenStatusBox.innerHTML = "Purple Pen file loaded successfully."

      //Activate compilers
      mapsCompiler.startTeXLive();
    };
    freader.onerror = function () {
      ppenStatusBox.innerHTML = "Could not read Purple Pen file. Try reselecting it, then click Reload.";
    };
    freader.readAsText(fileobj);   //Reads as UTF-8
  }

  function setAllCourses() {
    //Validates then copies value from set all courses into all courses for any fields that have been set
    var control, controlValue, classSet, classSetLength, id, classList;

    classList = ["showStation", "kites", "zeroes", "mapShape", "mapSize", "mapScale", "contourInterval"];
    for (const controlClass of classList) {
      classSet = document.getElementsByClassName(controlClass);
      classSetLength = classSet.length;
      control = classSet[0];

      if (controlClass == "zeroes" || controlClass == "showStation") {
        //Check whether control has been set
        if (control.indeterminate == false) {
          controlValue = control.checked;
          //Copy value
          for (id = 1; id < classSetLength; id++) {
            //id = 0 is the master control
            classSet[id].checked = controlValue;
          }
          //Set to indeterminate
          control.indeterminate = true;
          //Flag parameters as changed
          paramsSaved = false;
        }
      } else if (controlClass == "mapShape") {
        if (control.selectedIndex > 0) {
          controlValue = control.selectedIndex;
          for (id = 1; id < classSetLength; id++) {
            //id = 0 is the master control
            classSet[id].selectedIndex = controlValue - 1;
          }
          control.selectedIndex = 0;
          //Flag parameters as changed
          paramsSaved = false;
        }
      } else {
        //Control is required, so invalid if blank
        if (control.checkValidity() == true) {
          controlValue = control.value;
          for (id = 1; id < classSetLength; id++) {
            //id = 0 is the master control
            classSet[id].value = controlValue;
          }
          control.value = "";
          //Flag parameters as changed
          paramsSaved = false;
        }
      }
    }
  }

  function setAllLayout() {
    //Validates then copies value from set all courses into all courses for any layout fields that have been set
    var controlClass, control, controlValue, classSet, classSetLength, id;

    for (controlClass in defaultLayout) {
      classSet = document.getElementsByClassName(controlClass);
      classSetLength = classSet.length;
      //The first input element is the default button
      control = classSet[1];

      //Control is required, so invalid if blank
      if (control.checkValidity() == true) {
        controlValue = control.value;
        for (id = 2; id < classSetLength; id++) {
          //id = 0 is the master control
          classSet[id].value = controlValue;
        }
        control.value = "";
        //Flag parameters as changed
        paramsSaved = false;
      }
    }
  }

  function loadTeX(fileInput) {
    //Loads existing LaTeX file into memory
    var fileobj, freader, fname;
    const statusBox = document.getElementById("texReadStatus");
    fileobj = fileInput.files[0];
    if (fileobj) {
      fname = fileobj.name;
      freader = new FileReader();
      freader.onload = function () {
        //Populates station tables from previous LaTeX parameters file
        var fileString, startPos, endPos, subString, varArray, fields, rowId, numRows, classRoot;

        try {
          fileString = freader.result;

          //Indicate that parameters data is currently saved - unedited opened file
          paramsSaved = true;

          //Show station
          startPos = fileString.indexOf("\\def\\ShowStationList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("showStation");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (parseInt(varArray[rowId], 10) == 1) {
                fields[rowId + 1].checked = true;
              } else {
                fields[rowId + 1].checked = false;
              }
            }
          }

          //Number of kites
          startPos = fileString.indexOf("\\def\\NumKitesList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("kites");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId + 1].value = parseInt(varArray[rowId], 10);
              }
            }
          }

          //Zeroes
          startPos = fileString.indexOf("\\def\\ZeroOptionList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("zeroes");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (parseInt(varArray[rowId], 10) == 1) {
                fields[rowId + 1].checked = true;
              } else {
                fields[rowId + 1].checked = false;
              }
            }
          }

          //Heading
          startPos = fileString.indexOf("\\def\\MapHeadingList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("heading");
            //No master field for heading
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId].value = Number(varArray[rowId]);
              }
            }
          }

          //Map shape
          startPos = fileString.indexOf("\\def\\SquareMapList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("mapShape");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId + 1].selectedIndex = parseInt(varArray[rowId], 10);
              }
            }
          }

          //Map size
          startPos = fileString.indexOf("\\def\\CircleRadiusList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("mapSize");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId + 1].value = Number(varArray[rowId]) * 2;
              }
            }
          }

          //Map scale
          startPos = fileString.indexOf("\\def\\MapScaleList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("mapScale");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId + 1].value = Number(varArray[rowId]);
              }
            }
          }

          //Contour interval
          startPos = fileString.indexOf("\\def\\ContourIntervalList{{");
          if (startPos >= 0) {
            endPos = fileString.indexOf(",}}", startPos);
            startPos = fileString.indexOf("{{", startPos);
            subString = fileString.slice(startPos + 2, endPos);
            varArray = subString.split(",");
            fields = document.getElementsByClassName("contourInterval");
            numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
            for (rowId = 0; rowId < numRows; rowId++) {
              if (varArray[rowId] !== "") {
                fields[rowId + 1].value = Number(varArray[rowId]);
              }
            }
          }

          //Layout customisations
          //Root of name in TeX file
          const layoutTeXNames = {
            IDFontSize: "StationIDFontSize",
            checkWidth: "SheetCheckBoxWidth",
            checkHeight: "SheetCheckBoxHeight",
            checkFontSize: "CheckNumberHeight",
            removeFontSize: "RemoveTextFontSize",
            pointHeight: "PointingBoxHeight",
            letterFontSize: "PointingLetterFontSize",
            phoneticFontSize: "PointingPhoneticFontSize"
          }
          for (classRoot in layoutTeXNames) {
            startPos = fileString.indexOf("\\def\\" + layoutTeXNames[classRoot] + "List{{");
            if (startPos >= 0) {
              endPos = fileString.indexOf(",}}", startPos);
              startPos = fileString.indexOf("{{", startPos);
              subString = fileString.slice(startPos + 2, endPos);
              varArray = subString.split(",");
              fields = document.getElementsByClassName(classRoot);
              //First two rows are not fields
              numRows = Math.min(fields.length, document.getElementById("layoutTableBody").getElementsByTagName("tr").length - 2);
              for (rowId = 0; rowId < numRows; rowId++) {
                if (varArray[rowId] !== "") {
                  if (varArray[rowId].includes("cm")) {
                    //cm unit and quotes added, needs removing
                    fields[rowId + 2].value = Number(varArray[rowId].slice(1, -3));
                  } else {
                    fields[rowId + 2].value = Number(varArray[rowId]);
                  }
                }
              }
            }
          }

          //Debug circles enabled?
          startPos = fileString.indexOf("\\def\\AdjustMode{");
          if (startPos >= 0) {
            subString = fileString.slice(startPos + 16, startPos + 17);
            if (subString == "1") {
              document.getElementById("debugCircle").checked = true;
            } else {
              document.getElementById("debugCircle").checked = false;
            }
          } else {
            document.getElementById("debugCircle").checked = false;
          }

          statusBox.innerHTML = "Data loaded successfully.";
        } catch (err) {
          statusBox.innerHTML = "Error reading data: " + err;
          console.log(err);
        }
      };
      freader.onerror = function (err) {
        if (err.name == undefined) {
          statusBox.innerHTML = "Could not read file due to an unknown error. This occurs on Safari for files containing the % symbol - try deleting all of them.";
        } else {
          statusBox.innerHTML = "Could not read file: " + err;
        }
        console.log(err);
      };
      freader.readAsText(fileobj);   //Reads as UTF-8
    }
  }

  function generateLaTeX() {
    //Generates LaTeX parameters file
    //Returns string with error message or "ok" if no errors

    var tableRows, layoutRows, numTableRows, tableRowID, contentField, numStations, maxProblems, numProblems, showStationList, numProblemsList, stationName, stationNameList, numKites, kitesList, zeroesList, headingList, shapeList, mapSize, sizeList, briefingWidthList, scaleList, contourList, mapFileList, mapPageList, mapxList, mapyList, CDsFileList, CDsPageList, CDsxList, CDsyList, controlsSkipped, CDsxCoord, CDsyCoord, CDsHeightList, CDsWidthList, CDsScaleList, CDsaFontList, CDsbFontList, fileName, showPointingBoxesList, pointingBoxWidthList, pointingBoxHeightList, pointingLetterFontList, pointingPhoneticFontList, stationIDFontList, checkBoxWidthList, checkBoxHeightList, checkNumberFontList, checkRemoveFontList, fileString, iterNum, CDsxCoordBase, CDsyCoordBase, CDsWidthBase, CDsHeightBase, CDsPDFScale, parametersBlob;

    tableRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
    layoutRows = document.getElementById("layoutTableBody").getElementsByTagName("tr");
    numTableRows = tableRows.length;

    //Define constants
    //Positioning relative to bottom-left corner and size of control descriptions in source PDF
    CDsxCoordBase = 1.25;
    CDsyCoordBase = 26.58;
    CDsHeightBase = 0.77;
    CDsWidthBase = 5.68;
    CDsPDFScale = 7 / 6; //Enlarges CDs from 6mm to 7mm boxes. LaTeX crashes if too many decimal places.

    //Create variables, often strings, to accumulate
    numStations = 0;
    maxProblems = 0;
    showStationList = "\\def\\ShowStationList{{";
    numProblemsList = "\\def\\ProblemsPerStationList{{";
    stationNameList = "\\def\\StationIDList{{";
    kitesList = "\\def\\NumKitesList{{";
    zeroesList = "\\def\\ZeroOptionList{{";
    headingList = "\\def\\MapHeadingList{{";
    shapeList = "\\def\\SquareMapList{{";
    sizeList = "\\def\\CircleRadiusList{{";
    briefingWidthList = "\\def\\BriefingWidthList{{";
    scaleList = "\\def\\MapScaleList{{";
    contourList = "\\def\\ContourIntervalList{{";
    mapFileList = "\\def\\MapFileList{{";
    mapPageList = "\\def\\MapPageList{{";
    mapxList = "\\def\\ControlxCoordinateList{{";
    mapyList = "\\def\\ControlyCoordinateList{{";
    CDsFileList = "\\def\\DescriptionFileList{{";
    CDsPageList = "\\def\\DescriptionPageList{{";
    CDsxList = "\\def\\DescriptionxCoordinateList{{";
    CDsyList = "\\def\\DescriptionyCoordinateList{{";
    CDsHeightList = "\\def\\DescriptionHeightList{{";
    CDsWidthList = "\\def\\DescriptionWidthList{{";
    CDsScaleList = "\\def\\DescriptionPDFScaleList{{";
    CDsaFontList = "\\def\\CDaFontSizeList{{";
    CDsbFontList = "\\def\\CDbFontSizeList{{";
    showPointingBoxesList = "\\def\\ShowPointingBoxesList{{";
    pointingBoxWidthList = "\\def\\PointingBoxWidthList{{";
    pointingBoxHeightList = "\\def\\PointingBoxHeightList{{";
    pointingLetterFontList = "\\def\\PointingLetterFontSizeList{{";
    pointingPhoneticFontList = "\\def\\PointingPhoneticFontSizeList{{";
    stationIDFontList = "\\def\\StationIDFontSizeList{{";
    checkBoxWidthList = "\\def\\SheetCheckBoxWidthList{{";
    checkBoxHeightList = "\\def\\SheetCheckBoxHeightList{{";
    checkNumberFontList = "\\def\\CheckNumberHeightList{{";
    checkRemoveFontList = "\\def\\RemoveTextFontSizeList{{";

    for (tableRowID = 1; tableRowID < numTableRows; tableRowID++) {
      numStations++;
      stationName = tableRows[tableRowID].getElementsByClassName("stationName")[0].innerHTML;	//Store it for useful error messages later
      stationNameList += "\"" + stationName + "\"";

      //Show station
      if (tableRows[tableRowID].getElementsByClassName("showStation")[0].checked == true) {
        showStationList += "1";
      } else {
        showStationList += "0";
      }

      //Number of problems at station
      numProblems = Number(tableRows[tableRowID].getElementsByClassName("numProblems")[0].innerHTML);	//Save for later
      numProblemsList += numProblems;
      if (numProblems > maxProblems) {
        maxProblems = numProblems;
      }

      //Number of kites
      contentField = tableRows[tableRowID].getElementsByClassName("kites")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The number of kites for station " + stationName + " must be an integer between 1 and 6.");
      } else {
        numKites = Number(contentField.value);
        kitesList += numKites;	//Don't add quotes
      }

      if (tableRows[tableRowID].getElementsByClassName("zeroes")[0].checked == true) {
        zeroesList += "1";
        numKites += 1;
      } else {
        zeroesList += "0";
      }
      pointingBoxWidthList += (18 / numKites);

      //Heading
      contentField = tableRows[tableRowID].getElementsByClassName("heading")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The heading of station " + stationName + " must be a number.");
      } else {
        headingList += contentField.value;	//Don't add quotes
      }

      //Map shape
      contentField = tableRows[tableRowID].getElementsByClassName("mapShape")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The map shape for station " + stationName + " must be specified.");
      } else {
        shapeList += contentField.selectedIndex;	//Don't add quotes
      }

      //Map size
      contentField = tableRows[tableRowID].getElementsByClassName("mapSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The map size for station " + stationName + " must be > 0 and <= 12.");
      } else if (contentField.value == 0) {
        contentField.focus();
        throw new Error("The map size for station " + stationName + " must be strictly greater than 0.");
      } else {
        mapSize = Number(contentField.value);
        sizeList += 0.5 * mapSize;	//Don't add quotes
        briefingWidthList += "\"" + (0.7 * contentField.value) + "cm\"";
      }

      //Map scale
      contentField = tableRows[tableRowID].getElementsByClassName("mapScale")[0];
      if (contentField.checkValidity() == false || contentField.value == 0) {
        contentField.focus();
        throw new Error("The map scale for station " + stationName + " must be strictly greater than 0.");
      } else {
        scaleList += contentField.value;	//Don't add quotes
      }

      //Map contour interval
      contentField = tableRows[tableRowID].getElementsByClassName("contourInterval")[0];
      if (contentField.checkValidity() == false || contentField.value == 0) {
        contentField.focus();
        throw new Error("The contour interval for station " + stationName + " must be strictly greater than 0.");
      } else {
        contourList += contentField.value;	//Don't add quotes
      }

      //Map and control description files
      controlsSkipped = tableRows[tableRowID].getElementsByClassName("controlsSkipped")[0].innerHTML.split(",");
      fileName = "\"Maps\"";
      mapFileList += "{";
      CDsFileList += "{";
      CDsxList += "{";
      CDsyList += "{";
      CDsHeightList += "{";
      CDsWidthList += "{";
      CDsScaleList += "{";
      //Add a comma after all elements including the last one. A scalar without comma is misinterpreted by LaTeX.
      for (iterNum = 0; iterNum < numProblems; iterNum++) {
        mapFileList += fileName + ",";
        CDsFileList += "\"CDs\",";
        //Control description size parameters taking 6mm boxes in PDF from PPen and displaying 7mm on map cards
        CDsxList += CDsxCoordBase + ",";	//Must match number just above for Purple Pen
        CDsyCoord = CDsyCoordBase - 0.6 * controlsSkipped[iterNum];
        CDsyList += CDsyCoord + ",";	//Must match number just above for Purple Pen
        CDsHeightList += CDsHeightBase + ",";   //For a 7mm box
        CDsWidthList += CDsWidthBase + ",";   //For a 7mm box
        CDsScaleList += CDsPDFScale.toString() + ",";
      }
      mapFileList += "}";
      CDsFileList += "}";
      CDsxList += "}";
      CDsyList += "}";
      CDsHeightList += "}";
      CDsWidthList += "}";
      CDsScaleList += "}";
      CDsaFontList += "\"0.45cm\"";
      CDsbFontList += "\"0.39cm\"";

      //Coordinate map positions in files
      mapxList += tableRows[tableRowID].getElementsByClassName("circlex")[0].innerHTML;
      mapyList += tableRows[tableRowID].getElementsByClassName("circley")[0].innerHTML;
      mapPageList += tableRows[tableRowID].getElementsByClassName("printPage")[0].innerHTML;
      CDsPageList += tableRows[tableRowID].getElementsByClassName("printPage")[0].innerHTML;

      //Layout parameters. Remember to account for extra header row.

      //Station name font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("IDFontSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The name font size for station " + stationName + " must be between 0 and 29.7.");
      } else {
        stationIDFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box width
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkWidth")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The page order box width for station " + stationName + " must be between 0 and 29.7.");
      } else {
        checkBoxWidthList += contentField.value;
      }

      //Check box height
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkHeight")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The page order box height for station " + stationName + " must be between 0 and 29.7.");
      } else {
        checkBoxHeightList += contentField.value;
      }

      //Check box number font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkFontSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The page number font size for station " + stationName + " must be between 0 and 29.7.");
      } else {
        checkNumberFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box remove text font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("removeFontSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The <em>Remove</em> font size for station " + stationName + " must be between 0 and 29.7.");
      } else {
        checkRemoveFontList += "\"" + contentField.value + "cm\"";
      }

      //Pointing box height
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("pointHeight")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The pointing box height for station " + stationName + " must be between 0 and 29.7.");
      } else {
        pointingBoxHeightList += contentField.value;

        //Show pointing boxes only if they will fit: Map diameter + Pointing box height <= 12.5cm
        if (mapSize + Number(contentField.value) <= 12.5) {
          showPointingBoxesList += "1";
        } else {
          showPointingBoxesList += "0";
        }
      }

      //Pointing box letter font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("letterFontSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The pointing box letter font size for station " + stationName + " must be between 0 and 29.7.");
      } else {
        pointingLetterFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box remove text font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("phoneticFontSize")[0];
      if (contentField.checkValidity() == false) {
        contentField.focus();
        throw new Error("The pointing box phonetic font size for station " + stationName + " must be between 0 and 29.7.");
      } else {
        pointingPhoneticFontList += "\"" + contentField.value + "cm\"";
      }

      //Add a comma after all elements including the last one. A scalar without comma is misinterpreted by LaTeX.
      showStationList += ",";
      numProblemsList += ",";
      stationNameList += ",";
      kitesList += ",";
      zeroesList += ",";
      headingList += ",";
      shapeList += ",";
      sizeList += ",";
      briefingWidthList += ",";
      scaleList += ",";
      contourList += ",";
      mapFileList += ",";
      mapPageList += ",";
      mapxList += ",";
      mapyList += ",";
      CDsFileList += ",";
      CDsPageList += ",";
      CDsxList += ",";
      CDsyList += ",";
      CDsHeightList += ",";
      CDsWidthList += ",";
      CDsScaleList += ",";
      CDsaFontList += ",";
      CDsbFontList += ",";
      showPointingBoxesList += ",";
      pointingBoxWidthList += ",";
      pointingBoxHeightList += ",";
      pointingLetterFontList += ",";
      pointingPhoneticFontList += ",";
      stationIDFontList += ",";
      checkBoxWidthList += ",";
      checkBoxHeightList += ",";
      checkNumberFontList += ",";
      checkRemoveFontList += ",";
    }

    //Show construction circle for lining up maps? Not required when using this wizard, but still a useful feature.
    if (document.getElementById("debugCircle").checked == true) {
      fileString = "\\def\\AdjustMode{1}\n";
    } else {
      fileString = "\\def\\AdjustMode{0}\n";
    }

    //Write to fileString
    //Insert a comma to introduce an extra element to the array where it contains a string ending in cm, otherwise TikZ parses it incorrectly/doesn't recognise an array of length 1.
    fileString += "\\newcommand{\\NumStations}{" + numStations + "}\n";
    fileString += "\\newcommand{\\MaxProblemsPerStation}{" + maxProblems + "}\n";
    fileString += showStationList + "}}\n";
    fileString += numProblemsList + "}}\n";
    fileString += stationNameList + "}}\n";
    fileString += kitesList + "}}\n";
    fileString += zeroesList + "}}\n";
    fileString += headingList + "}}\n";
    fileString += shapeList + "}}\n";
    fileString += sizeList + "}}\n";
    fileString += briefingWidthList + "}}\n";
    fileString += scaleList + "}}\n";
    fileString += contourList + "}}\n";
    fileString += mapFileList + "}}\n";
    fileString += mapPageList + "}}\n";
    fileString += mapxList + "}}\n";
    fileString += mapyList + "}}\n";
    fileString += CDsFileList + "}}\n";
    fileString += CDsPageList + "}}\n";
    fileString += CDsxList + "}}\n";
    fileString += CDsyList + "}}\n";
    fileString += CDsHeightList + "}}\n";
    fileString += CDsWidthList + "}}\n";
    fileString += CDsScaleList + "}}\n";
    fileString += CDsaFontList + "}}\n";
    fileString += CDsbFontList + "}}\n";
    fileString += showPointingBoxesList + "}}\n";
    fileString += pointingBoxWidthList + "}}\n";
    fileString += pointingBoxHeightList + "}}\n";
    fileString += pointingLetterFontList + ",}}\n";
    fileString += pointingPhoneticFontList + "}}\n";
    fileString += stationIDFontList + "}}\n";
    fileString += checkBoxWidthList + "}}\n";
    fileString += checkBoxHeightList + "}}\n";
    fileString += checkNumberFontList + "}}\n";
    fileString += checkRemoveFontList + "}}\n";

    //Create file
    return new Blob([fileString], { type: "text/plain" });
  }

  function downloadFile(fileBlob, fileName) {
    //Downloads a file blob
    var downloadElement, url;
    if (window.navigator.msSaveBlob) {
      //Microsoft
      window.navigator.msSaveBlob(fileBlob, fileName);
    } else {
      //Other browsers
      downloadElement = document.createElement("a");
      url = URL.createObjectURL(fileBlob);
      downloadElement.href = url;
      downloadElement.download = fileName;
      document.body.appendChild(downloadElement);
      downloadElement.click();
      setTimeout(function () {
        document.body.removeChild(downloadElement);
        URL.revokeObjectURL(url);
      }, 0);
    }
  }

  function saveParameters() {
    var rtn;
    rtn = generateLaTeX();
    downloadFile(rtn, "TemplateParameters.tex");

    //Indicate that parameters data is currently saved - unedited opened file
    paramsSaved = true;
  }

  function resetAllLayout() {
    //Inserts the default value into all the layout set all rows fields
    const tableRow = document.getElementById("layoutSetAllRow");
    for (var btnClass in defaultLayout) {
      if (defaultLayout.hasOwnProperty(btnClass)) {
        //Do not act on properties inherited from generic object class
        tableRow.getElementsByClassName(btnClass)[0].value = defaultLayout[btnClass];
      }
    }
  }

  function resetField(btn) {
    //Resets the corresponding field to its default value
    //Get the corresponding input element class of the button
    var btnClass = btn.className;
    //Look up corresponding input element and value
    document.getElementById("layoutSetAllRow").getElementsByClassName(btnClass)[0].value = defaultLayout[btnClass];
  }

  function updateTemplate() {
    const templateName = this.value;
    mapsCompiler.setTemplate(templateName);

    //Show/hide sets of instructions
    const printInstructions = document.getElementById("printInstructions");
    const onlineInstructions = document.getElementById("onlineInstructions");
    if (templateName === "printA5onA4") {
      printInstructions.hidden = false;
      onlineInstructions.hidden = true;
    } else {
      printInstructions.hidden = true;
      onlineInstructions.hidden = false;
    }
  }

  class Compiler {
    constructor(obj) {
      this.statusBox = obj.statusBox;
      this.outDownload = obj.outDownload;
      this.logDownload = obj.logDownload;
    }

    static async generateAll() {
      //Prevent second button press while compiling
      const btn = document.getElementById("compileLaTeXBtn");
      btn.disabled = true;
      try {
        mapsCompiler.resetStatus();

        mapsCompiler.compileUnitsItems = Compiler.totalNumTasks();
        if (mapsCompiler.compileUnitsItems === 0) { throw new Error("No stations selected"); }

        //Check PDF resources are present
        const coursePDFFiles = document.getElementById("coursePDFSelector").files;
        if (coursePDFFiles.length === 0) {
          throw new Error("Course maps PDF is missing. Try selecting it again.");
        }
        const cdPDFFiles = document.getElementById("CDPDFSelector").files;
        if (cdPDFFiles.length === 0) {
          throw new Error("Control descriptions PDF is missing. Try selecting it again.");
        }

        //Make LaTeX parameters file
        mapsCompiler.statusBox.textContent = "Validating station data.";
        const paramFile = generateLaTeX();

        //Read maps and CDs PDFs
        const resourceNames = ["TemplateParameters.tex", "Maps.pdf", "CDs.pdf"];
        const resourceFileArray = [paramFile, coursePDFFiles[0], cdPDFFiles[0]];
        // Blob.arrayBuffer isn't implemented in Safari - switch to new code when available
        // const resourceBuffers = await Promise.all(resourceFileArray.map((file) => file.arrayBuffer()));
        const resourceBuffers = await Promise.all(resourceFileArray.map((file) => new Promise((resolve, reject) => {
          const freader = new FileReader();
          freader.onload = (ev) => { resolve(ev.target.result); };
          freader.onerror = () => { reject(freader.error); };
          freader.readAsArrayBuffer(file);
        })));

        //Start compilers - await promises later
        const compilerPromises = [mapsCompiler.compile(resourceBuffers, resourceNames)];

        //Download a copy of parameters file to save for later
        if (document.getElementById("autoSave").checked && !paramsSaved) {
          downloadFile(paramFile, "TemplateParameters.tex");
          //Indicate that parameters data is currently saved
          paramsSaved = true;
        }

        await Promise.allSettled(compilerPromises);
      } catch (err) {
        mapsCompiler.statusBox.textContent = err;
        console.error(err);
      }
      btn.disabled = false;
    }

    static setDownload(newURL, newName, element) {
      //Revoke old URL
      const oldURL = element.href;
      if (oldURL) {
        URL.revokeObjectURL(oldURL);
      }
      element.href = newURL;
      element.download = newName;
      element.hidden = false;
    }

    static totalNumTasks() {
      //Get total number of tasks to be outputted
      const tableRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
      let numTasks = 0;
      for (const row of tableRows) {
        const numProblemsSet = row.getElementsByClassName("numProblems");
        //First row is defaults so ignore
        if (numProblemsSet.length > 0 && row.getElementsByClassName("showStation")[0].checked) {
          numTasks += Number(numProblemsSet[0].textContent);
        }
      }
      return numTasks;
    }

    async compile(resourceBuffers, resourceNames) {
      try {
        //Might be sensible to preconvert in future
        this.statusBox.textContent = "Reading PDFs.";
        const resourceStrings = resourceBuffers.map((buffer) => TeXLive.arrayBufferToString(buffer));

        //One compile unit = time to insert one map + CD
        this.compileTotalUnits = this.compileUnitsPre + this.compileUnitsItems + this.compileUnitsPost;

        this.statusBox.textContent = "Compiling PDF (0%).";
        const pdfURL = await this.texlive.compile(this.texsrc, resourceStrings, resourceNames);
        this.statusBox.textContent = "Compiling PDF (100%).";
        await this.postCompile(pdfURL);
      } catch (err) {
        this.statusBox.textContent = err;
        console.error(err);
        const logLength = this.logContent.length;
        if (logLength > 0) {
          //Display log
          let logStr = "";
          for (let rowId = 0; rowId < logLength; rowId++) {
            logStr += this.logContent[rowId] + "\n";
          }
          Compiler.setDownload(URL.createObjectURL(new Blob([logStr], { type: "text/plain" })), "TCTemplate.log", this.logDownload);
        }
        //Do not rethrow as any action taken does not depend on success of output
      }
    }

    downloadOutput(newURL) {
      Compiler.setDownload(newURL, this.downloadFileName, this.outDownload);
      this.outDownload.click();
      this.statusBox.textContent = "Map cards produced successfully.";
    }

    logEvent(msg) {
      //Called when a status message is output by TeXLive
      //Provides additionally monitoring and error handling beyond texlivelight
      console.log(msg);
      this.logContent.push(msg);
      if (msg.includes("no output PDF file produced!")) {
        //TeXLive encountered an error -> handle it
        let msg;
        if (this.logContent[this.logContent.length - 3] === "!pdfTeX error: /latex (file ./Maps.pdf): PDF inclusion: required page does not ") {
          msg = "PDF of maps does not contain enough pages. Please follow the instructions in step 8 carefully and try again.";
        } else if (this.logContent[this.logContent.length - 3] === "!pdfTeX error: /latex (file ./CDs.pdf): PDF inclusion: required page does not e") {
          msg = "PDF of control descriptions does not contain enough pages. Please follow the instructions in step 9 carefully and try again.";
        }
        throw new Error(msg);
      } else if (msg.includes("<Maps.pdf,")) {
        this.statusBox.textContent = "Compiling PDF (" + (this.compileUnitsDone / this.compileTotalUnits * 100).toFixed() + "%).";
        this.compileUnitsDone++;
      }
    }

    resetStatus() {
      this.outDownload.hidden = true;
      this.logDownload.hidden = true;
      //Reset log
      this.logContent = [];
      //One compile unit = time to insert one map + CD; populate with start-up time
      this.compileUnitsDone = this.compileUnitsPre;
    }

    async startTeXLive() {
      try {
        await loadScript("texlive");
        TeXLive.workerFolder = "texlive.js/";
        this.texlive = new TeXLive({
          onlog: (msg) => this.logEvent(msg)
        });
        //Fetch template ready for use
        TeXLive.getFile(this.texsrc);
      } catch (err) {
        document.getElementById("compileLaTeXBtn").disabled = true;
        this.statusBox.textContent = "Loading error: " + err;
        console.error(err);
      }
    }
  }

  const loadScript = (() => {
    //Download and store each required script
    const sources = {
      texlive: {
        label: "TeXLive",
        src: "texlive.js/pdftexlight.js",
        remotePath: ""
      },
      pdfjs: {
        label: "PDF.js",
        src: "pdf.min.js",
        remotePath: "https://cdn.jsdelivr.net/npm/pdfjs-dist@2.4.456/build/",
        onready: async (remotePath) => {
          pdfjsLib = window["pdfjs-dist/build/pdf"];
          pdfjsLib.GlobalWorkerOptions.workerSrc = await TeXLive.getFile(remotePath + "pdf.worker.min.js");
        }
      },
      jszip: {
        label: "JSZip",
        src: "jszip.min.js",
        remotePath: "https://cdn.jsdelivr.net/npm/jszip@3.4.0/dist/"
      }
    }

    function add2DOM(source, useRemote) {
      let remotePath;
      if (useRemote) {
        remotePath = source.remotePath;
      } else {
        remotePath = "";
      }
      return new Promise((resolve, reject) => {
        const scriptEl = document.createElement("script");
        scriptEl.addEventListener("load", async () => {
          if (typeof source.onready === "function") { await source.onready(remotePath); }
          resolve();
        }, { once: true });
        scriptEl.addEventListener("error", reject, { once: true });
        scriptEl.src = remotePath + source.src;
        document.head.appendChild(scriptEl);
      });
    }

    return async (scriptName) => {
      const source = sources[scriptName];
      if (typeof source === "undefined") { throw new ReferenceError("Requested load of unrecognised script: " + scriptName); }
      if (source.promise === undefined) {
        //Writing in terms of promises is cleaner here than try await catch
        //Try to load using CDN
        source.promise = add2DOM(source, true).catch((err) => {
          if (source.remotePath !== "") {
            //Look for local copy
            return add2DOM(source, false).catch((err) => {
              let msg;
              if (location.hostname.includes("tdobra.github.io")) {
                throw undefined;
              } else {
                throw new Error("Need to download scripts for " + source.label + " to root folder");
              }
            });
          } else {
            throw undefined;
          }
        }).catch((err) => {
          if (err === undefined) {
            throw new Error("Failed to download script: " + source.label);
          } else {
            throw err;
          }
        });
      }
      return source.promise;
    };
  })();

  const mapsCompiler = new Compiler({
    statusBox: document.getElementById("compileStatus"),
    outDownload: document.getElementById("savePDF"),
    logDownload: document.getElementById("viewLog")
  });

  //Do not use arrow functions when binding methods as properties
  mapsCompiler.makePNGs = async function (pdfURL) {
    this.statusBox.textContent = "Preparing to split into images.";

    //Create objects
    const canvasFull = document.createElement("canvas");
    const canvasCropped = document.createElement("canvas");
    const ctxFull = canvasFull.getContext("2d");
    const ctxCropped = canvasCropped.getContext("2d");
    let pageNum = 1;
    let taskCount = 0;
    const tableRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
    //Row 0 of table is set all stations
    const numStations = tableRows.length - 1;
    const imPromises = [];

    await loadScript("pdfjs");
    const pdfDoc = await pdfjsLib.getDocument(pdfURL).promise;
    const numPages = pdfDoc.numPages;

    await loadScript("jszip");
    const zip = new JSZip();

    this.statusBox.textContent = "Splitting into images (0%).";
    const numImages = this.compileUnitsItems + numStations;
    for (let stationId = 1; stationId <= numStations; stationId++) {
      const numTasks = Number(tableRows[stationId].getElementsByClassName("numProblems")[0].textContent);
      if (tableRows[stationId].getElementsByClassName("showStation")[0].checked) {
        const stationName = tableRows[stationId].getElementsByClassName("stationName")[0].textContent;

        //Image resolution
        const imDPI = 200;
        const pdfScale = imDPI / 72; //PDF renders at 72 DPI by default

        for (let taskId = 0; taskId <= numTasks; taskId++) {
          //Read page into canvas
          const page = await pdfDoc.getPage(pageNum);
          const viewport = page.getViewport({ scale: pdfScale });
          canvasFull.height = viewport.height;
          canvasFull.width = viewport.width;
          const renderContext = {
            canvasContext: ctxFull,
            viewport: viewport
          };
          await page.render(renderContext).promise;

          //Work out blank margins around items on page
          //Use canvasFull rather than viewport dimensions to ensure integers
          const im = ctxFull.getImageData(0, 0, canvasFull.width, canvasFull.height).data;

          function pixelHasContent(row, column) {
            //ImageData is a single array going RGBA, left-to-right then top-to-bottom
            //White or alpha = 0
            const pixelStart = (row * canvasFull.width + column) * 4;
            return ((im[pixelStart] < 255 || im[pixelStart + 1] < 255 || im[pixelStart + 2] < 255) && im[pixelStart + 3] > 0);
          }

          //Find first nonblank row - white or alpha = 100%
          let topRow;
          outerLoop:
          for (topRow = 0; topRow < canvasFull.height; topRow++) {
            for (let colId = 0; colId < canvasFull.width; colId++) {
              if (pixelHasContent(topRow, colId)) { break outerLoop; }
            }
          }
          if (topRow === canvasFull.height) { throw new Error("Image blank"); }
          //Find bottom row
          let bottomRow;
          outerLoop:
          for (bottomRow = canvasFull.height - 1; bottomRow > topRow; bottomRow--) {
            for (let colId = 0; colId < canvasFull.width; colId++) {
              if (pixelHasContent(bottomRow, colId)) { break outerLoop; }
            }
          }
          //Find left column
          let leftCol;
          outerLoop:
          for (leftCol = 0; leftCol < canvasFull.width; leftCol++) {
            for (let rowId = topRow; rowId <= bottomRow; rowId++) {
              if (pixelHasContent(rowId, leftCol)) { break outerLoop; }
            }
          }
          //Find right column
          let rightCol;
          outerLoop:
          for (rightCol = canvasFull.width - 1; rightCol > leftCol; rightCol--) {
            for (let rowId = topRow; rowId <= bottomRow; rowId++) {
              if (pixelHasContent(rowId, rightCol)) { break outerLoop; }
            }
          }

          //Put cropped image onto cropped canvas
          let croppedHeight = bottomRow - topRow + 1;
          const croppedWidth = rightCol - leftCol + 1;

          //Crop map to max height 300px
          if (croppedHeight > 300) {
            const cropAmount = Math.ceil((croppedHeight - 300) / 2);
            bottomRow -= cropAmount;
            topRow += cropAmount;
            croppedHeight -= 2 * cropAmount;
          }

          //Expand to width 1000px, keeping centred
          const targetWidth = 1000;
          const xOffset = croppedWidth < targetWidth ? Math.floor((targetWidth - croppedWidth) / 2) : 0;

          canvasCropped.height = croppedHeight;
          canvasCropped.width = croppedWidth > targetWidth ? croppedWidth : targetWidth;
          //Initialise to white background
          ctxCropped.fillStyle = "white";
          ctxCropped.fillRect(0, 0, canvasCropped.width, canvasCropped.height);
          ctxCropped.drawImage(canvasFull, leftCol, topRow, croppedWidth, croppedHeight, xOffset, 0, croppedWidth, croppedHeight);

          //Save to zip
          const blob = await new Promise((resolve) => { canvasCropped.toBlob(resolve); });
          imPromises.push(zip.file("map-" + stationName + "." + taskId + "z.png", blob));

          pageNum++;
          taskCount++;
          this.statusBox.textContent = "Splitting into images (" + (taskCount / numImages * 100).toFixed() + "%).";
        }
      } else {
        pageNum += numTasks;
      }
    }
    await Promise.all(imPromises);
    const zipBlob = await zip.generateAsync({ type: "blob" });
    this.downloadOutput(URL.createObjectURL(zipBlob));
  }

  mapsCompiler.setTemplate = function (name) {
    //Compile time units are determined empirically for each template to give an approximate progress of compiler
    switch (name) {
      case "printA5onA4":
        this.texsrc = "TCTemplate.tex";
        this.postCompile = this.downloadOutput;
        this.downloadFileName = "TCMapCards.pdf";
        this.compileUnitsPre = 7.5;
        this.compileUnitsPost = 3;
        break;
      case "onlineTempO":
        this.texsrc = "onlinetempo.tex";
        this.postCompile = this.makePNGs;
        this.downloadFileName = "TCMapCards.zip";
        this.compileUnitsPre = 10;
        this.compileUnitsPost = 0.2;
        //Load prequisities for post-processing
        loadScript("pdfjs");
        loadScript("jszip");
        break;
      default:
        throw new ReferenceError("Template " + name + " not recognised");
    }
    if (this.texlive !== undefined) {
      //Fetch template ready for use - async, so won't block code
      TeXLive.getFile(this.texsrc);
    }
  };

  document.getElementById("selectTemplate").addEventListener("change", updateTemplate);
  updateTemplate.call(document.getElementById("selectTemplate"));
  document.getElementById("compileLaTeXBtn").addEventListener("click", Compiler.generateAll, { passive: true });

  //Make required functions globally visible
  return {
    loadppen: loadppen,
    loadTeX: loadTeX,
    saveParameters: saveParameters,
    setAllCourses: setAllCourses,
    resetAllLayout: resetAllLayout,
    setAllLayout: setAllLayout,
    resetField: resetField
  };
})();
