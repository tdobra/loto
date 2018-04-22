//Put script in body of HTML

//Global variables
var fileNameRoot;

//Prevent sloppy programming and throw more errors
"use strict";

//Check browser for XML support
if (window.FileReader && window.DOMParser && window.Blob && window.URL) {
	document.getElementById("missingAPIs").hidden = true;   //Hide error message that shows by default
} else {
	document.getElementById("mainView").hidden = true;
}
document.getElementById("stationProperties").hidden = true;
document.getElementById("compileMaps").hidden = true;

function loadppen(fileInput) {
	//Reads a Purple Pen file
	var fileobj, freader, testPattern;
	var courseOrderUsed = [];
	
	fileobj = fileInput.files[0];
	if (!fileobj) { return; }	//Nothing selected
	fileNameRoot = fileobj.name.slice(0, fileobj.name.lastIndexOf("."));
	
	//Check for spaces - throw error if any found, because LaTeX won't like it
	testPattern = /\s/;
	if (testPattern.test(fileNameRoot)) {
		window.alert("File name cannot contain spaces.");
		return;
	}
	
	freader = new FileReader();
	freader.onload = function () {
		var xmlParser, xmlobj, parsererrorNS, globalScale, courseNodes, courseNodesId, courseNodesNum, tableRowNode, tableColNode, tableContentNode, selectOptionNode, existingRows, existingRowID, otherNode, leftcoord, bottomcoord, courseControlNode, controlNode, controlsSkipped, numProblems, stationNameRoot;
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
			window.alert("Could not read Purple Pen file: its XML is invalid.");
			return;
		}
			
		//Reset course table
		existingRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
		for (existingRowsID = existingRows.length - 1; existingRowsID > 0; existingRowsID--) {
			document.getElementById("courseTableBody").removeChild(existingRows[existingRowsID]);
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
		//Make list of all course order attributes used - to determine print page, so must be completed before rest of file reading
		for (courseNodesId = 0; courseNodesId < courseNodesNum; courseNodesId++) {	
			courseOrderUsed.push(courseNodes[courseNodesId].getAttribute("order"));
		}
		courseOrderUsed.sort(function(a, b){return a - b});
	
		//Find courses with name *.1. xpath doesn't appear to be working in Safari, iterate over nodes.
		for (courseNodesId = 0; courseNodesId < courseNodesNum; courseNodesId++) {
			if (courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent.endsWith(".1")) {
				//Check course type = score
				if (courseNodes[courseNodesId].getAttribute("kind") != "score") {
					window.alert("Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + " type must be set to score.");
					return;
				}
				
				//Check zero page margin
				//Portrait vs. landscape is irrelevant, as coordinates are determined by left and bottom attributes
				otherNode = courseNodes[courseNodesId].getElementsByTagName("print-area")[0];
				if (otherNode) {
					if (otherNode.getAttribute("page-margins") > 0) {
						window.alert("The page margin must be set to 0 on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
						return;
					}
				} else {
					window.alert("The page margin must be set to 0 on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
					return;
				}
				if (otherNode.getAttribute("automatic") == "true") {
					window.alert("The print area selection must be set to manual on Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
					return;
				}
					
				//Create new table row
				tableRowNode = document.createElement("tr");
						
				//Create first column - show + hidden values
				tableColNode = document.createElement("td");
				tableRowNode.appendChild(tableColNode);
				tableContentNode = document.createElement("input");
				tableContentNode.type = "checkbox";
				tableContentNode.checked = true;
				tableContentNode.className = "showStation";
				tableColNode.appendChild(tableContentNode);
				//Hidden: Course order
				tableContentNode = document.createElement("span");
				tableContentNode.className = "courseOrder"
				tableContentNode.hidden = true;
				tableContentNode.innerHTML = courseNodes[courseNodesId].getAttribute("order");
				tableColNode.appendChild(tableContentNode);
											
				//Create second column - station name
				tableColNode = document.createElement("td");
				stationNameRoot = courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent.slice(0,-2);
				tableContentNode = document.createElement("span");
				tableContentNode.className = "stationName";
				tableContentNode.innerHTML = stationNameRoot;
				tableColNode.appendChild(tableContentNode);
				tableRowNode.appendChild(tableColNode);
						
				//Third column - number of kites
				tableColNode = document.createElement("td");
				tableContentNode = document.createElement("input");
				tableContentNode.type = "number";
				tableContentNode.min = 1;
				tableContentNode.max = 6;
				tableContentNode.step = 1;
				tableContentNode.required = true;
				tableContentNode.className = "kites";
				tableColNode.appendChild(tableContentNode);
				tableRowNode.appendChild(tableColNode);
						
				//Fourth column - zeroes allowed?
				tableColNode = document.createElement("td");
				tableRowNode.appendChild(tableColNode);
				tableContentNode = document.createElement("input");
				tableContentNode.type = "checkbox";
				tableContentNode.className = "zeroes";
				tableColNode.appendChild(tableContentNode);
						
				//Fifth column - station heading
				tableColNode = document.createElement("td");
				tableContentNode = document.createElement("input");
				tableContentNode.type = "number";
				tableContentNode.step = "any";
				tableContentNode.required = true;
				tableContentNode.className = "heading";
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
				tableContentNode.className = "mapSize";
				tableColNode.appendChild(tableContentNode);
				tableContentNode = document.createTextNode(" cm");
				tableColNode.appendChild(tableContentNode);
				tableRowNode.appendChild(tableColNode);
						
				//Eigth column - map scale
				tableColNode = document.createElement("td");
				tableContentNode = document.createTextNode("1:");
				tableColNode.appendChild(tableContentNode);
				tableContentNode = document.createElement("input");
				tableContentNode.type = "number";
				tableContentNode.step = "any";
				tableContentNode.min = 0;
				tableContentNode.required = true;
				tableContentNode.className = "mapScale";
				tableColNode.appendChild(tableContentNode);
				tableRowNode.appendChild(tableColNode);
				//Populate map scale
				courseScale = Number(courseNodes[courseNodesId].getElementsByTagName("options")[0].getAttribute("print-scale"));
				if (!courseScale) {
					courseScale = globalScale;
				}
				tableContentNode.value = courseScale;
						
				//Nineth column - contour interval + hidden values
				tableColNode = document.createElement("td");
				tableContentNode = document.createElement("input");
				tableContentNode.type = "number";
				tableContentNode.step = "any";
				tableContentNode.min = 0;
				tableContentNode.required = true;
				tableContentNode.className = "contourInterval";
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
					bottomcoord = courseNodes[existingRowID].getElementsByTagName("print-area")[0];
					if (!bottomcoord) {
						window.alert("Page setup not complete for Purple Pen course " + courseNodes[courseNodesId].getElementsByTagName("name")[0] + ".");
						return;
					}
					leftcoord = Number(bottomcoord.getAttribute("left"));
					bottomcoord = Number(bottomcoord.getAttribute("bottom"));
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
					//Circle position in cm with origin in bottom left corner
					if (numProblems > 1) {
						//Insert comma in list
						tableColNode.getElementsByClassName("circlex")[0].innerHTML += ",";
						tableColNode.getElementsByClassName("circley")[0].innerHTML += ",";
						tableColNode.getElementsByClassName("printPage")[0].innerHTML += ",";
						tableColNode.getElementsByClassName("controlsSkipped")[0].innerHTML += ",";
					}
					tableColNode.getElementsByClassName("circlex")[0].innerHTML += 0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("x")) - leftcoord);
					tableColNode.getElementsByClassName("circley")[0].innerHTML += 0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("y")) - bottomcoord);

					//Read course order attribute, then find its position in list of course order values used. Adding one onto this gives the page number when all courses, except blank, are printed in a single PDF.
					tableColNode.getElementsByClassName("printPage")[0].innerHTML += (courseOrderUsed.indexOf(courseNodes[existingRowID].getAttribute("order")) + 1).toString();
					tableColNode.getElementsByClassName("controlsSkipped")[0].innerHTML += controlsSkipped;
					//Find next control at station
					otherNode = stationNameRoot + "." + (numProblems + 1);
					for (existingRowID = 0; existingRowID < courseNodesNum; existingRowID++) {
						if (courseNodes[existingRowID].getElementsByTagName("name")[0].textContent == otherNode) {
							//Found another control
							numProblems++;
							//Check new course type, margin and manual print selection
							if (courseNodes[existingRowID].getAttribute("kind") != "score") {
								window.alert("Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + " type must be set to score.");
								return;
							}
							//Check zero page margin
							otherNode = courseNodes[existingRowID].getElementsByTagName("print-area")[0];
							if (otherNode) {
								if (otherNode.getAttribute("page-margins") > 0) {
									window.alert("The page margin must be set to 0 on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
									return;
								}
							} else {
								window.alert("The page margin must be set to 0 on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
								return;
							}
							if (otherNode.getAttribute("automatic") == "true") {
								window.alert("The print area selection must be set to manual on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ". Then, recreate the course PDF.");
								return;
							}
							//Check same scale
							otherNode = courseNodes[existingRowID].getElementsByTagName("options")[0];
							if (otherNode) {
								if (Number(otherNode.getAttribute("print-scale")) != courseScale) {
									window.alert("The print scale is different on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ".");
									return;
								}
							} else if (courseScale != globalScale) {
								window.alert("The print scale is different on Purple Pen course " + courseNodes[existingRowID].getElementsByTagName("name")[0].textContent + ".");
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
						
				//Insert row in correct position for course order
				existingRows = document.getElementById("courseTableBody").getElementsByClassName("courseOrder");
				existingRowID = 0;
				for (;;) {
					existingRowID++;
					existingRow = existingRows[existingRowID];
					if (!existingRow) {
						document.getElementById("courseTableBody").appendChild(tableRowNode);
						break;	//No more rows to consider
					}
					if (Number(existingRow.innerHTML) > Number(courseNodes[courseNodesId].getAttribute("order"))) {
						document.getElementById("courseTableBody").insertBefore(tableRowNode, existingRow.parentElement);
						break;	//Current row needs to be inserted before existingRow
					}
				}						
			}
		}
		
		//Reset update all stations fields
		document.getElementsByClassName("kites")[0].value = "";
		document.getElementsByClassName("zeroes")[0].indeterminate = true;
		document.getElementsByClassName("mapShape")[0].selectedIndex = 0;
		document.getElementsByClassName("mapSize")[0].value = "";
		document.getElementsByClassName("mapScale")[0].value = "";
		document.getElementsByClassName("contourInterval")[0].value = "";
		
		//Prepare view
		document.getElementById("stationProperties").hidden = false;
		document.getElementById("stationProperties").scrollIntoView();
	};
	freader.onerror = function () { window.alert("Could not read Purple Pen file. Try reselecting it, then click Reload."); };
	freader.readAsText(fileobj);   //Reads as UTF-8
}
			
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
		
function setAllCourses() {
	//Validates then copies value from set all courses into all courses for any fields that have been set
	var control, controlClass, controlValue, classSet, classSetLength, id, classList;
	
	classList = ["kites", "zeroes", "mapShape", "mapSize", "mapScale", "contourInterval"];
	for (const controlClass of classList) {
		classSet = document.getElementsByClassName(controlClass);
		classSetLength = classSet.length;
		control = classSet[0];
		
		if (controlClass == "zeroes") {
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
			}
		} else if (controlClass == "mapShape") {
			if (control.selectedIndex > 0) {
				controlValue = control.selectedIndex;
				for (id = 1; id < classSetLength; id++) {
					//id = 0 is the master control
					classSet[id].selectedIndex = controlValue - 1;
				}
				control.selectedIndex = 0;
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
			}
		}
	}
}

function loadTeX(fileInput) {
	//Loads existing LaTeX file into memory
	var fileobj, freader, fname;
	//Check browser for file reading support
	if (!window.FileReader) {
		window.alert("Web browser does not support FileReader API. Need to update to newer browser to read files.");
		return -1;
	} else {
		fileobj = fileInput.files[0];
		if (fileobj) {
			fname = fileobj.name;
			freader = new FileReader();
			freader.onload = function () {
				readTeX(freader.result);
			};
			freader.onerror = function () { window.alert("Could not read file"); };
			freader.readAsText(fileobj);   //Reads as UTF-8
		}
	}
}

function readTeX(fileString) {
	//Populates station table from previous LaTeX parameters file
	var startPos, subString, varArray, fields, rowId, numRows;
	
	//Number of kites
	startPos = fileString.indexOf("\\def\\NumKitesList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("kites");
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId + 1].value = parseInt(varArray[rowId], 10);
		}
	}
	
	//Zeroes
	startPos = fileString.indexOf("\\def\\ZeroOptionList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
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
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("heading");
		//No master field for heading
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId].value = Number(varArray[rowId]);
		}
	}
	
	//Map shape
	startPos = fileString.indexOf("\\def\\SquareMapList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("mapShape");
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId + 1].selectedIndex = parseInt(varArray[rowId], 10);
		}
	}
	
	//Map size
	startPos = fileString.indexOf("\\def\\CircleRadiusList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("mapSize");
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId + 1].value = Number(varArray[rowId]) * 2;
		}
	}
	
	//Map scale
	startPos = fileString.indexOf("\\def\\MapScaleList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("mapScale");
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId + 1].value = Number(varArray[rowId]);
		}
	}
	
	//Contour interval
	startPos = fileString.indexOf("\\def\\ContourIntervalList{{");
	if (startPos >= 0) {
		endPos = fileString.indexOf("}}", startPos);
		startPos = fileString.indexOf("{{", startPos);
		subString = fileString.slice(startPos + 2, endPos);
		varArray = subString.split(",");
		fields = document.getElementsByClassName("contourInterval");
		numRows = Math.min(fields.length, document.getElementById("courseTableBody").getElementsByTagName("tr").length - 1);
		for (rowId = 0; rowId < numRows; rowId++) {
			fields[rowId + 1].value = Number(varArray[rowId]);
		}
	}
}

function generateLaTeX() {
	//Generates and downloads LaTeX parameters file
	var tableRows, numTableRows, tableRowID, contentField, numStations, maxProblems, numProblems, numProblemsList, stationName, stationNameList, numKites, kitesList, zeroesList, headingList, shapeList, sizeList, briefingWidthList, scaleList, contourList, mapFileList, mapPageList, mapxList, mapyList, CDsFileList, CDsPageList, CDsxList, CDsyList, controlsSkipped, CDsyCoord, CDsHeightList, CDsWidthList, CDsaFontList, CDsbFontList, fileName, showPointingBoxesList, pointingBoxWidthList, pointingBoxHeightList, pointingLetterFontList, pointingPhoneticFontList, stationIDFontList, checkBoxWidthList, checkBoxHeightList, checkNumberFontList, checkRemoveFontList, fileString, fileBlob, downloadElement, url, iterNum, CDsxCoordBase, CDsyCoordBase, CDsWidthBase, CDsHeightBase;
	
	tableRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
	numTableRows = tableRows.length;

    //Define constants
    //Positioning relative to bottom-left corner and size of control descriptions in source PDF
	CDsxCoordBase = 1.25;
	CDsyCoordBase = 26.28;
	CDsHeightBase = 0.77;
	CDsWidthBase = 5.68;
	
	//Create variables, often strings, to accumulate
	numStations = 0;
	maxProblems = 0;
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
		if (tableRows[tableRowID].getElementsByClassName("showStation")[0].checked == true) {
			numStations++;
			if (numStations > 1) {
				//Insert commas in lists
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
			
			stationName = tableRows[tableRowID].getElementsByClassName("stationName")[0].innerHTML;	//Store it for useful error messages later
			stationNameList += "\"" + stationName + "\"";
			
			//Number of problems at station
			numProblems = Number(tableRows[tableRowID].getElementsByClassName("numProblems")[0].innerHTML);	//Save for later
			numProblemsList += numProblems;
			if (numProblems > maxProblems) {
				maxProblems = numProblems;
			}
			
			//Number of kites
			contentField = tableRows[tableRowID].getElementsByClassName("kites")[0];
			if (contentField.checkValidity() == false) {
				window.alert("The number of kites for station " + stationName + " must be an integer between 1 and 6.");
				contentField.focus();
				return;
			}
			numKites = Number(contentField.value);
			kitesList += numKites;	//Don't add quotes
			
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
				window.alert("The heading of station " + stationName + " must be a number.");
				contentField.focus();
				return;
			}
			headingList += contentField.value;	//Don't add quotes
			
			//Map shape
			contentField = tableRows[tableRowID].getElementsByClassName("mapShape")[0];
			if (contentField.checkValidity() == false) {
				window.alert("The map shape for station " + stationName + " must be specified.");
				contentField.focus();
				return;
			}
			shapeList += contentField.selectedIndex;	//Don't add quotes
			
			//Map size
			contentField = tableRows[tableRowID].getElementsByClassName("mapSize")[0];
			if (contentField.checkValidity() == false) {
				window.alert("The map size for station " + stationName + " must be > 0 and <= 12.");
				contentField.focus();
				return;
			}
			if (contentField.value == 0) {
				window.alert("The map size for station " + stationName + " must be strictly greater than 0.");
				contentField.focus();
				return;
			}
			sizeList += 0.5 * contentField.value;	//Don't add quotes
			briefingWidthList += "\"" + (0.7 * contentField.value) + "cm\"";
			//Show pointing boxes only if diameter < 10cm
			if (contentField.value <= 10) {
				showPointingBoxesList += "1";
			} else {
				showPointingBoxesList += "0";
			}

			//Map scale
			contentField = tableRows[tableRowID].getElementsByClassName("mapScale")[0];
			if (contentField.checkValidity() == false || contentField.value == 0) {
				window.alert("The map scale for station " + stationName + " must be strictly greater than 0.");
				contentField.focus();
				return;
			}
			scaleList += contentField.value;	//Don't add quotes
			
			//Map contour interval
			contentField = tableRows[tableRowID].getElementsByClassName("contourInterval")[0];
			if (contentField.checkValidity() == false || contentField.value == 0) {
				window.alert("The contour interval for station " + stationName + " must be strictly greater than 0.");
				contentField.focus();
				return;
			}
			contourList += contentField.value;	//Don't add quotes
			
			//Map and control description files
			fileName = "\"" + fileNameRoot + "\"";
			mapFileList += "{" + fileName;
			CDsFileList += "{\"CDs\"";
			CDsxList += "{" + CDsxCoordBase;
			controlsSkipped = tableRows[tableRowID].getElementsByClassName("controlsSkipped")[0].innerHTML.split(",");
			CDsyCoord = CDsyCoordBase - 0.7 * controlsSkipped[0];
			CDsyList += "{" + CDsyCoord;
			CDsHeightList += "{" + CDsHeightBase;   //For a 7mm box
			CDsWidthList += "{" + CDsWidthBase;   //For a 7mm box
			for (iterNum = 1; iterNum < numProblems; iterNum++) {
				mapFileList += "," + fileName;
				CDsFileList += ",\"CDs\"";
				CDsxList += "," + CDsxCoordBase;	//Must match number just above for Purple Pen
				CDsyCoord = CDsyCoordBase - 0.7 * controlsSkipped[iterNum];
				CDsyList += "," + CDsyCoord;	//Must match number just above for Purple Pen
				CDsHeightList += "," + CDsHeightBase;
				CDsWidthList += "," + CDsWidthBase;   //For a 7mm box
            }
			mapFileList += "}";
			CDsFileList += "}";
			CDsxList += "}";
			CDsyList += "}";
			CDsHeightList += "}";
			CDsWidthList += "}";
			CDsaFontList += "\"0.45cm\"";
			CDsbFontList += "\"0.39cm\"";
									
			//Coordinate map positions in files
			mapxList += tableRows[tableRowID].getElementsByClassName("circlex")[0].innerHTML;
			mapyList += tableRows[tableRowID].getElementsByClassName("circley")[0].innerHTML;
			mapPageList += tableRows[tableRowID].getElementsByClassName("printPage")[0].innerHTML;
			CDsPageList += tableRows[tableRowID].getElementsByClassName("printPage")[0].innerHTML;

			//Layout constant parameters - for A5
			pointingBoxHeightList += "2.5";
			pointingLetterFontList += "\"1.8cm\"";
			pointingPhoneticFontList += "\"0.6cm\"";
			stationIDFontList += "\"0.7cm\"";
			checkBoxWidthList += "1.5";	
			checkBoxHeightList += "1.5";	
			checkNumberFontList += "\"0.8cm\"";
			checkRemoveFontList += "\"0.3cm\"";
		}
	}
	
	//Hide construction circle for lining up maps - not required when using this wizard
	fileString = "\\def\\AdjustMode{0}\n";
	
    //Write to fileString
    //Insert a comma to introduce an extra element to the array where it contains a string ending in cm, otherwise TikZ parses it incorrectly/doesn't recognise an array of length 1.
	fileString += "\\newcommand{\\NumStations}{" + numStations + "}\n";
	fileString += "\\newcommand{\\MaxProblemsPerStation}{" + maxProblems + "}\n";
	fileString += numProblemsList + "}}\n";
	fileString += stationNameList + "}}\n";
	fileString += kitesList + "}}\n";
	fileString += zeroesList + "}}\n";
	fileString += headingList + "}}\n";
	fileString += shapeList + "}}\n";
	fileString += sizeList + "}}\n";
	fileString += briefingWidthList + ",}}\n";
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
	fileString += CDsaFontList + ",}}\n";
	fileString += CDsbFontList + ",}}\n";
	fileString += showPointingBoxesList + "}}\n";
	fileString += pointingBoxWidthList + "}}\n";
	fileString += pointingBoxHeightList + "}}\n";
	fileString += pointingLetterFontList + ",}}\n";
	fileString += pointingPhoneticFontList + ",}}\n";
	fileString += stationIDFontList + ",}}\n";
	fileString += checkBoxWidthList + "}}\n";
	fileString += checkBoxHeightList + "}}\n";
	fileString += checkNumberFontList + ",}}\n";
	fileString += checkRemoveFontList + ",}}\n";
	
	//Create and download file
	fileBlob = new Blob([fileString], { type: "text/plain" });
	if (window.navigator.msSaveBlob) {
		//Microsoft
		window.navigator.msSaveBlob(fileBlob, "TemplateParameters.tex");
	} else {
		//Other browsers
		downloadElement = document.createElement("a");
		url = URL.createObjectURL(fileBlob);
		downloadElement.href = url;
		downloadElement.download = "TemplateParameters.tex";
		document.body.appendChild(downloadElement);
		downloadElement.click();
		setTimeout(function () {
			document.body.removeChild(downloadElement);
			URL.revokeObjectURL(url);
		}, 0);
	}
	
	//Show next stage of instructions
	document.getElementById("compileMaps").hidden = false;
	document.getElementById("compileMaps").scrollIntoView();
}
