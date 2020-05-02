//Put script in body of HTML

//Prevent sloppy programming and throw more errors
"use strict";

//Check browser supports required APIs
if (FileReader && DOMParser && Blob && URL && fetch) {
  document.getElementById("missingAPIs").hidden = true;   //Hide error message that shows by default
} else {
  document.getElementById("mainView").hidden = true;
}

//Old MS Edge/IE doesn't always work - <details> tag not implemented
if ("open" in document.createElement("details")) {
  document.getElementById("MSEdgeWarning").hidden = true;
}

//Keep all functions private and put those with events in HTML tags in a namespace
const tcTemplate = (() => {
  //Keep track of whether an input file has been changed in a table to disable autosave
  var paramsSaved = true;

  //Object structure: Reorderable list->Station list-array of>Station->(Field->Specialised field->)Named field

  //Classes are not hoisted, so declare them first
  class IterableList {
    constructor(obj) {
      this.selectorSpan = obj.selectorSpan;
      this.addBtn = obj.addBtn;
      this.deleteBtn = obj.deleteBtn;
      this.upBtn = obj.upBtn;
      this.downBtn = obj.downBtn;
      if (obj.itemInFocus === undefined) { obj.itemInFocus = 0; };
      this.defaultInFocus = obj.defaultInFocus;
      this.itemInFocus = obj.itemInFocus;
      this.items = [];
      //Create HTML select element, but don't insert yet
      this.selector = document.createElement("select");
      this.selector.addEventListener("change", () => { this.refresh(); });
    }

    get activeItem() {
      if (this.defaultInFocus) {
        return this.default;
      } else {
        return this.items[this.itemInFocus];
      }
    }

    createBtnListeners() {
      //Do not call until activeItem is defined
      //Always use arrow functions to ensure this points to the IterableList rather than the calling DOM element
      this.addBtn.addEventListener("click", () => { this.add(); });
      this.deleteBtn.addEventListener("click", () => { this.activeItem.deleteThis(true); });
      this.upBtn.addEventListener("click", () => { this.activeItem.move(-1); });
      this.downBtn.addEventListener("click", () => { this.activeItem.move(1); });
    }

    add() {
      //Create option in item selector
      const newNode = this.selector.appendChild(document.createElement("option"));

      //New item initially takes default values
      const newItem = this.newItem(newNode);

      //Name the new item
      newItem.itemName.value = (this.items.length + 1).toString();
      newNode.innerHTML = newItem.itemName.value;

      //Check validity of data in new item including all fields
      newItem.checkValidity(true);

      //Add to the end
      this.items.push(newItem);

      //Try to change to new itme. The add station button is disabled when on defaults.
      newNode.selected = true;
      this.refresh(true);
    }

    applyAll(action, fieldList) {
      //Applies specified action function to all fields with reset buttons
      for (const field of fieldList) {
        if (this.activeItem[field].resetBtn !== undefined) {
          this.activeItem[field][action]();
        }
      }
    }

    showHideMoveBtns() {
      //Show the delete button only if more than one option
      if (this.selector.length > 1) {
        this.deleteBtn.disabled = false;
      } else {
        this.deleteBtn.disabled = true;
      }

      //Shows or hides the move up and move down buttons
      switch (this.selector.selectedIndex) {
      case 0:
        //First station, so can't move it up
        this.upBtn.disabled = true;
        this.downBtn.disabled = false;
        break;
      case (this.selector.length - 1):
        //Last station, so can't move it down
        this.upBtn.disabled = false;
        this.downBtn.disabled = true;
        break;
      default:
        //Can move in both directions
        this.upBtn.disabled = false;
        this.downBtn.disabled = false;
      }
    }
  }

  class StationList extends IterableList {
    constructor() {
      const obj = {
        selectorSpan: document.getElementById("stationSelect"),
        addBtn: document.getElementById("addStation"),
        deleteBtn: document.getElementById("deleteStation"),
        upBtn: document.getElementById("moveUpStation"),
        downBtn: document.getElementById("moveDownStation"),
        defaultInFocus: true,
        itemInFocus: 0
      };
      super(obj);
      //Insert selector of size 5
      this.selectorSpan.appendChild(this.selector);
      this.selector.size = 5;
      //Radio button options
      this.defaultRadio = document.getElementById("defaultRadio");
      this.stationRadio = document.getElementById("stationRadio");
      //Set up current/dynamic defaults
      this.default = new Station({
        parentObj: this,
        optionElement: undefined,
        copyStation: undefined
      });
      this.default.checkValidity(true);
    }

    newItem(newNode) {
      return new Station({
        parentObj: this,
        optionElement: newNode,
        copyStation: this.default
      });
    }

    refresh(storeValues = false) {
      //storeValues is a boolean stating whether to commit values in form fields to variables in memory.

      //Other DOM elements
      const setAllResetCSS = document.getElementById("showSetAllCSS");

      if (storeValues) {
        //Only permit change of station if the name is valid and unique
        if (this.activeItem.itemName.valid === false) {
          //Abort
          alert("Please insert a valid and unique station name");
          //Change radio buttons and selectors back to original values
          if (this.defaultInFocus) {
            this.defaultRadio.checked = true;
          } else {
            this.stationRadio.checked = true;
            this.selector.selectedIndex = stationList.itemInFocus;
          }
          return;
        }
      }

      //Remove task etc. selectors for station to be hidden
      if (this.defaultInFocus === false) {
        this.activeItem.taskList.selector.remove();
      }

      //Show/hide or enable/disable HTML elements according to new selected station
      this.itemInFocus = this.selector.selectedIndex;
      if (this.defaultRadio.checked) {
        //Setting defaults for all stations
        this.defaultInFocus = true;

        //Show/hide any buttons as required
        setAllResetCSS.innerHTML = ".hideIfDefault{ display: none; }";

        //Some fields need disabling
        this.selector.disabled = true;
        this.addBtn.disabled = true;
        this.deleteBtn.disabled = true;
        this.upBtn.disabled = true;
        this.downBtn.disabled = true;
        this.default.itemName.inputElement.disabled = true;
        this.default.heading.inputElement.disabled = true;
      } else {
        this.defaultInFocus = false;

        //Show/hide any buttons as required
        setAllResetCSS.innerHTML = ".setAllCourses{ display: none; }";

        //Some fields may need enabling
        this.selector.disabled = false;
        this.addBtn.disabled = false;
        this.showHideMoveBtns();
        this.activeItem.itemName.inputElement.disabled = false;
        this.activeItem.heading.inputElement.disabled = false;

        //Insert new task etc. selectors
        this.activeItem.taskList.selectorSpan.appendChild(this.activeItem.taskList.selector);
      }

      //Populate with values for new selected station and show error/warning messages
      this.activeItem.refreshAllInput();
    }
  }

  class TaskList extends IterableList {
    constructor(parentObj) {
      const obj = {
        selectorSpan: document.getElementById("taskSelect"),
        addBtn: document.getElementById("addTask"),
        deleteBtn: document.getElementById("deleteTask"),
        upBtn: document.getElementById("moveUpTask"),
        downBtn: document.getElementById("moveDownTask"),
        defaultInFocus: false,
        itemInFocus: 0
      };
      super(obj);
      this.parentItem = parentObj;
      this.selector.size = 5;
    }

    newItem(newNode) {
      return new Task({
        parentObj: this,
        optionElement: newNode,
        copyItem: undefined
      });
    }

    setNumItems() {
      //Forcibly adds/removes items from end of list to get required length
      while (this.items.length < this.numItems) {
        this.add();
      }
      while (this.items.length > this.numItems) {
        this.items[this.items.length - 1].deleteThis(false);
      }
    }
  }

  //Types of items
  class IterableItem {
    constructor(obj) {
      this.parentList = obj.parentObj;
      this.optionElement = obj.optionElement;
      if (this.getItemType() === "station") {
        this.station = this;
      } else {
        this.station = this.parentList.parentItem;
      }
    }

    get index() {
      let thisIndex = 0;
      while (this.parentList.items[thisIndex] !== this) { thisIndex++; }
      return thisIndex;
    }

    isActive() {
      return this.parentList.activeItem === this;
    }

    isDefault() {
      return this.parentList.default === this;
    }

    deleteThis(checkFirst = true) {
      //Deletes this item
      if (checkFirst) {
        if (!confirm("Are you sure you want to delete? This cannot be undone.")) {
          return;
        }
      }
      this.optionElement.remove();
      this.parentList.iteminFocus = 0;
      this.parentList.selector.selectedIndex = 0;
      //Remove this station from array
      this.parentList.items.splice(this.index, 1);
      //Don't save deleted values then refresh input fields
      this.parentList.refresh(false);
    }

    move(offset = 0) {
      //Moves this item

      //Check target is in bounds
      const newPos = this.index + offset;
      const arrayLength = this.parentList.items.length
      if (newPos < 0 || newPos >= arrayLength) {
        throw "Moving station to position beyond bounds of array";
      }
      if (!Number.isInteger(newPos)) {
        throw "Station position offset is not an integer";
      }

      //Rearrange selector
      this.optionElement.remove();
      let insertBeforeElement = null;
      if (newPos < arrayLength - 1) {
        insertBeforeElement = this.parentList.items[newPos].optionElement;
      } //Else move to end of list
      //Make sure to get a new element list, as it will have changed
      this.parentList.selector.insertBefore(this.optionElement, insertBeforeElement);

      //Remove this from array
      this.parentList.items.splice(this.index, 1);
      //Insert into new position
      this.parentList.items.splice(newPos, 0, this);

      //Update active station number
      this.parentList.itemInFocus = newPos;

      //Enable/disable move up/down buttons
      this.parentList.showHideMoveBtns();
    }
  }

  class Station extends IterableItem {
    constructor(obj) {
      super(obj);

      //Fields - populate with values given in copyStation, if present
      const coreFieldClasses = [
        StationName,
        ShowStation,
        NumKites,
        Zeroes,
        NumTasks,
        Heading,
        MapShape,
        MapSize,
        MapScale,
        ContourInterval,
      ];
      const customLayoutClasses = [
        IDFontSize,
        CheckWidth,
        CheckHeight,
        CheckFontSize,
        RemoveFontSize,
        PointHeight,
        LetterFontSize,
        PhoneticFontSize
      ];
      const classNames = coreFieldClasses.concat(customLayoutClasses);
      this.coreFields = coreFieldClasses.map((className) => className.getFieldName());
      this.customLayoutFields = customLayoutClasses.map((className) => className.getFieldName());
      this.fieldNames = this.coreFields.concat(this.customLayoutFields);
      if (obj.copyStation === undefined) {
        //Create an object to pass undefined parameters => triggers default values specified in class
        obj.copyStation = {};
        for (const field of this.fieldNames) {
          obj.copyStation[field] = { value: undefined };
        }
      }
      //WARNING: need to make a copy of any sub objects, otherwise will still refer to same memory
      let index = 0;
      for (const field of this.fieldNames) {
        this[field] = new classNames[index](this, obj.copyStation[field].value);
        index++;
      }
      this.taskList = new TaskList();
      this.taskList.setNumItems();
    }

    getItemType() {
      return "station";
    }

    isNonDefaultHidden() {
      //Returns true if the station is hidden and the default station is not in focus
      return !(this.showStation.value || this.isDefault());
    }

    checkValidity(recheckFields = true) {
      if (recheckFields === true) {
        //Recheck validity of all fields
        for (const field of this.fieldNames) {
          this[field].checkValidity();
        }
      }

      this.valid = this.fieldNames.every((field) => this[field].valid);
      if (!this.isDefault()) {
        //Highlight errors in the station selector
        if (this.valid) {
          this.optionElement.classList.remove("error");
        } else {
          this.optionElement.classList.add("error");
        }
      }
    }

    refreshAllInput() {
      for (const field of this.fieldNames) {
        this[field].refreshInput();
      }
    }
  }

  class Task extends IterableItem {
    constructor(obj) {
      super(obj);

      //Fields - populate with values given in copyTask, if present
      const fieldClasses = [
        TaskName,
        CirclePage,
        Circlex,
        Circley,
        CDpage,
        CDx,
        CDy,
        CDwidth,
        CDheight,
        CDscale
      ];
      this.fieldNames = fieldClasses.map((className) => className.getFieldName());
      if (obj.copyTask === undefined) {
        //Create an object to pass undefined parameters => triggers default values specified in class
        obj.copyTask = {};
        for (const field of this.fieldNames) {
          obj.copyTask[field] = { value: undefined };
        }
      }
      //WARNING: need to make a copy of any sub objects, otherwise will still refer to same memory
      let index = 0;
      for (const field of this.fieldNames) {
        this[field] = new fieldClasses[index](this, obj.copyTask[field].value);
        index++;
      }
    }

    getItemType() {
      return "task";
    }
  }

  //Fields
  class Field {
    //Generic field: string or select
    constructor(obj) {
      this.fieldName = this.constructor.getFieldName();
      this.parentItem = obj.parentObj;
      this.parentList = this.parentItem.parentList;
      if (this.parentItem.getItemType() === "station") {
        this.station = this.parentItem;
      } else {
        this.station = this.parentList.parentItem;
      }
      this.stationList = this.station.parentList; //Shortcut
      if (obj.value === undefined) {
        this.value = this.constructor.getOriginalValue();
      } else {
        this.value = obj.value;
      }
      this.inputElement = obj.inputElement;
      this.resetBtn = obj.resetBtn;
      this.setAllBtn = obj.setAllBtn;
      this.errorElement1 = obj.errorElement1;
      this.errorElement2 = obj.errorElement2;
    }

    get inputValue() {
      return this.inputElement.value;
    }

    set inputValue(val) {
      this.inputElement.value = val;
    }

    isDuplicate(ignoreHidden = false) {
      //Determines whether this value is duplicated at another item in this list. Ignores hidden stations on request.

      //Return false if the tested index is ignored
      if (ignoreHidden && this.station.showStation.value === false) { return false; }

      if (ignoreHidden) {
        return this.parentList.items.some((item) => (item[this.fieldName].value === this.value && item !== this.parentItem && this.station.showStation.value === true));
      } else {
        this.parentList.items.some((item) => (item[this.fieldName].value === this.value && item !== this.parentItem));
      }
    }

    matchesAll(ignoreHidden = false) {
      //Returns true if this field matches those on all other items in this list, excluding number fields set to NaN. Ignores NaN if number. Ignores hidden stations on request.

      //Return true if the tested index is ignored
      if (ignoreHidden && this.station.showStation.value === false) { return true; }

      if (ignoreHidden) {
        return this.parentList.items.every((item) => (item[this.fieldName].value === this.value || Number.isNaN(item[this.fieldName].value) || this.station.showStation.value === false));
      } else {
        return this.parentList.items.every((item) => (item[this.fieldName].value === this.value || Number.isNaN(item[this.fieldName].value)));
      }
    }

    refreshInput() {
      //Updates input element value - no user input
      this.inputValue = this.value;
      this.updateMsgs();
    }

    resetValue() {
      if (this.station.isDefault()) {
        //Resets to original value
        this.save(this.constructor.getOriginalValue());
      } else {
        //Resets to current default value
        this.save(this.stationList.default[this.fieldName].value);
      }
      this.refreshInput();
      this.station.checkValidity(false);
    }

    saveInput() {
      //Call when the value of the input element is updated by the user
      this.save(this.inputValue);
      this.updateMsgs();
      this.station.checkValidity(false);
    }

    save(val) {
      //Saves value and determines whether a change has occurred
      if (this.value !== val) {
        //Flag value as changed
        paramsSaved = false;
        //Write new value to memory
        this.value = val;
        //Check whether this new value is valid
        this.checkValidity();
      }
    }

    setAll() {
      //Sets this field to this value in all stations
      for (const station of this.stationList.items) {
        station[this.fieldName].save(this.value);
        station.checkValidity(false);
      }
    }

    //Empty functions, which may be overwritten in inheriting classes if some action is required
    checkValidity() {} //Most likely: field always valid
    updateMsgs() {}
  }

  class BooleanField extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true; //Never updated
    }

    get inputValue() {
      return this.inputElement.checked;
    }

    set inputValue(ticked) {
      this.inputElement.checked = ticked;
    }
  }

  class NumberField extends Field {
    //Same constructor as Field

    get inputValue() {
      const textValue = this.inputElement.value;
      if (textValue === "") {
        return NaN;
      } else {
        return Number(textValue);
      }
    }

    set inputValue(val) {
      //Val needs to be a number
      if (Number.isNaN(val)) {
        this.inputElement.value = "";
      } else {
        this.inputElement.value = val.toString();
      }
    }

    save(val) {
      //New & saved values not equal and at least one of them not NaN (NaN === NaN is false)
      if (this.value !== val && (Number.isNaN(this.value) === false || Number.isNaN(val) === false)) {
        //Flag value as changed
        paramsSaved = false;
        //Write new value to memory
        this.value = val;
        //Check whether this new value is valid
        this.checkValidity();
      }
    }
  }

  class NonNegativeField extends NumberField {
    checkValidity() {
      this.valid = (Number.isFinite(this.value) && this.value >= 0) || this.station.isNonDefaultHidden();
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      if (this.valid) {
        contentFieldClass.remove("error");
        errorMsgStyle.display = "none";
      } else {
        contentFieldClass.add("error");
        errorMsgStyle.display = "";
      }
    }
  }

  class StrictPositiveField extends NonNegativeField {
    checkValidity() {
      //LaTeX crashes if font size is set to zero
      this.valid =  (Number.isFinite(this.value) && this.value > 0) || this.station.isNonDefaultHidden();
    }
  }

  class StationName extends Field {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("stationName"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement1: document.getElementById("nameSyntax"),
        errorElement2: document.getElementById("nameUniqueness")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "itemName";
    }

    static getOriginalValue() {
      return "";
    }

    checkValidity() {
      //Check syntax even if hidden to avoid dodgy strings getting into LaTeX
      const stringFormat = /^[A-Za-z0-9][,.\-+= \w]*$/;
      this.syntaxError = !(this.station.isDefault() || stringFormat.test(this.value));
      this.duplicateError = !(this.station.isDefault()) && this.isDuplicate(false);
      this.valid = !(this.syntaxError || this.duplicateError);

      //Update station list
      if (!this.station.isDefault()) {
        this.station.optionElement.innerHTML = this.value;
      }
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const syntaxMsgStyle = this.errorElement1.style;
      const uniquenessMsgStyle = this.errorElement2.style;
      if (this.valid) {
        contentFieldClass.remove("error");
        syntaxMsgStyle.display = "none";
        uniquenessMsgStyle.display = "none";
      } else {
        contentFieldClass.add("error");
        //Don't show both error messages together
        if (this.syntaxError) {
          syntaxMsgStyle.display = "";
          uniquenessMsgStyle.display = "none";
        } else {
          syntaxMsgStyle.display = "none";
          uniquenessMsgStyle.display = "";
        }
      }
    }
  }

  class ShowStation extends BooleanField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("showStation"),
        resetBtn: document.getElementById("resetShowStation"),
        setAllBtn: document.getElementById("setAllShowStation"),
        errorElement1: undefined,
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "showStation";
    }

    static getOriginalValue() {
      return true;
    }

    checkValidity() {
      //Always valid
      //Most validity checks of other fields are ignored when station is hidden, so rerun all checks
      for (const field of this.station.fieldNames) {
        //Avoid infinite loops
        if (field !== "showStation") {
          this.station[field].checkValidity();
          if (this.station.isActive()) {
            this.station[field].updateMsgs();
          }
        }
      }
    }
  }

  class NumKites extends NumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("numKites"),
        resetBtn: document.getElementById("resetNumKites"),
        setAllBtn: document.getElementById("setAllNumKites"),
        errorElement1: document.getElementById("kitesRule"),
        errorElement2: undefined
      };
      super(inputObj);
      this.valid = true; //Always valid
    }

    static getFieldName() {
      return "numKites";
    }

    static getOriginalValue() {
      return 6;
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const ruleMsgStyle = this.errorElement1.style;
      if (this.value === 6 || this.station.isNonDefaultHidden()) {
        //Field valid, or non-defaults station not displayed
        contentFieldClass.remove("warning");
        ruleMsgStyle.display = "none";
      } else {
        contentFieldClass.add("warning");
        ruleMsgStyle.display = "";
      }
    }
  }

  class Zeroes extends BooleanField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("zeroes"),
        resetBtn: document.getElementById("resetZeroes"),
        setAllBtn: document.getElementById("setAllZeroes"),
        errorElement1: document.getElementById("zeroesWarning"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "zeroes";
    }

    static getOriginalValue() {
      return false;
    }

    updateMsgs() {
      const ruleMsgStyle = this.errorElement1.style;
      const contentFieldClass = this.inputElement.classList;
      //Check all values the same - don't bother if station is hidden or is defaults
      if (this.matchesAll(true) || this.station.isDefault()){
        contentFieldClass.remove("warning");
        ruleMsgStyle.display = "none";
      } else {
        contentFieldClass.add("warning");
        ruleMsgStyle.display = "";
      }
    }
  }

  class NumTasks extends NumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("numTasks"),
        resetBtn: undefined,
        setAllBtn: document.getElementById("setAllNumTasks"),
        errorElement1: document.getElementById("numTasksError"),
        errorElement2: document.getElementById("numTasksRule")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "numTasks";
    }

    static getOriginalValue() {
      return NaN;
    }

    checkValidity() {
      this.valid = (Number.isInteger(this.value) && this.value >= 1) || this.station.isNonDefaultHidden();
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      const ruleMsgStyle = this.errorElement2.style;
      if (this.valid) {
        contentFieldClass.remove("error");
        errorMsgStyle.display = "none";
        //Check all values the same - don't bother if station is hidden or is the defaults
        if (this.matchesAll(true) || this.station.isDefault()) {
          contentFieldClass.remove("warning");
          ruleMsgStyle.display = "none";
        } else {
          contentFieldClass.add("warning");
          ruleMsgStyle.display = "";
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        errorMsgStyle.display = "";
        ruleMsgStyle.display = "";
      }
    }
  }

  class Heading extends NumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("heading"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement1: document.getElementById("headingError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "heading";
    }

    static getOriginalValue() {
      return NaN;
    }

    checkValidity() {
      this.valid = (Number.isFinite(this.value) || this.station.showStation.value === false) || this.station.isDefault();
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      if (this.valid) {
        contentFieldClass.remove("error");
        errorMsgStyle.display = "none";
      } else {
        contentFieldClass.add("error");
        errorMsgStyle.dispaly = "";
      }
    }
  }

  class MapShape extends Field {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("mapShape"),
        resetBtn: document.getElementById("resetMapShape"),
        setAllBtn: document.getElementById("setAllMapShape"),
        errorElement1: document.getElementById("mapShapeRule"),
        errorElement2: undefined
      };
      super(inputObj);
      this.valid = true; //Always valid
    }

    static getFieldName() {
      return "mapShape";
    }

    static getOriginalValue() {
      return "Circle";
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const ruleMsgStyle = this.errorElement1.style;
      const sizeTypeElement = document.getElementById("mapSizeType");

      //Check all values the same - don't bother if station is hidden or is the defaults
      //No warning if station hidden or Defaults or all stations same
      if (this.matchesAll(true) || this.station.isDefault()) {
        contentFieldClass.remove("warning");
        ruleMsgStyle.display = "none";
      } else {
        contentFieldClass.add("warning");
        ruleMsgStyle.display = "";
      }

      //Update labels for map size description
      if (this.value === "Circle") {
        sizeTypeElement.innerHTML = "diameter";
      } else {
        sizeTypeElement.innerHTML = "side length";
      }
    }
  }

  class MapSize extends NumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("mapSize"),
        resetBtn: document.getElementById("resetMapSize"),
        setAllBtn: document.getElementById("setAllMapSize"),
        errorElement1: document.getElementById("mapSizePermitted"),
        errorElement2: document.getElementById("mapSizeRule")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "mapSize";
    }

    static getOriginalValue() {
      return NaN;
    }

    checkValidity() {
      this.valid = (Number.isFinite(this.value) && this.value > 0 && this.value <= 12) || this.station.isNonDefaultHidden();
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      const ruleMsgStyle = this.errorElement2.style;
      if (this.valid) {
        errorMsgStyle.display = "none";
        contentFieldClass.remove("error");
        //Check whether all values are the same and >= 5
        //No warning if non-default hidden station or (size >= 5 and (default or all stations same))
        if (this.station.isNonDefaultHidden() || (this.value >= 5 && (this.station.isDefault() || this.matchesAll(true)))) {
          contentFieldClass.remove("warning");
          ruleMsgStyle.display = "none";
        } else {
          contentFieldClass.add("warning");
          ruleMsgStyle.display = "";
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        errorMsgStyle.display = "";
        ruleMsgStyle.display = "";
      }
    }
  }

  class MapScale extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("mapScale"),
        resetBtn: document.getElementById("resetMapScale"),
        setAllBtn: document.getElementById("setAllMapScale"),
        errorElement1: document.getElementById("mapScalePermitted"),
        errorElement2: document.getElementById("mapScaleRule")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "mapScale";
    }

    static getOriginalValue() {
      return NaN;
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      const ruleMsgStyle = this.errorElement2.style;
      if (this.valid) {
        errorMsgStyle.display = "none";
        contentFieldClass.remove("error");
        //Check whether all values are the same and (equal 4000 or 5000)
        //No warning if non-default hidden station or (4000/5000 and (default or all stations same))
        if (this.station.isNonDefaultHidden() || ((this.value === 4000 || this.value === 5000) && (this.station.isDefault() || this.matchesAll(true)))) {
          contentFieldClass.remove("warning");
          ruleMsgStyle.display = "none";
        } else {
          contentFieldClass.add("warning");
          ruleMsgStyle.display = "";
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        errorMsgStyle.display = "";
        ruleMsgStyle.display = "";
      }
    }
  }

  class ContourInterval extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("contourInterval"),
        resetBtn: document.getElementById("resetContourInterval"),
        setAllBtn: document.getElementById("setAllContourInterval"),
        errorElement1: document.getElementById("contourIntervalPermitted"),
        errorElement2: document.getElementById("contourIntervalRule")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "contourInterval";
    }

    static getOriginalValue() {
      return NaN;
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const errorMsgStyle = this.errorElement1.style;
      const ruleMsgStyle = this.errorElement2.style;

      //Check validity - don't bother if station is non-default hidden
      if (this.valid) {
        errorMsgStyle.display = "none";
        contentFieldClass.remove("error");
        //Check whether all values are the same
        //No warning if non-default hidden station or defaults or all stations same
        if (this.station.isDefault() || this.matchesAll(true)) {
          contentFieldClass.remove("warning");
          ruleMsgStyle.display = "none";
        } else {
          contentFieldClass.add("warning");
          ruleMsgStyle.display = "";
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        errorMsgStyle.display = "";
        ruleMsgStyle.display = "";
      }
    }
  }

  class TaskName extends Field {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("taskName"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement1: document.getElementById("taskNameSyntax"),
        errorElement2: document.getElementById("taskNameUniqueness")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "itemName";
    }

    static getOriginalValue() {
      return "";
    }

    checkValidity() {
      //Check syntax even if hidden to avoid dodgy strings getting into LaTeX
      const stringFormat = /^[A-Za-z0-9][,.\-+= \w]*$/;
      this.syntaxError = !(this.station.isDefault() || stringFormat.test(this.value));
      this.duplicateError = !(this.station.isDefault()) && this.isDuplicate(false) && this.value !== "Kites" && this.value !== "VP";
      this.valid = !(this.syntaxError || this.duplicateError);

      //Update task list
      this.station.optionElement.innerHTML = this.value;
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const syntaxMsgStyle = this.errorElement1.style;
      const uniquenessMsgStyle = this.errorElement2.style;
      if (this.valid) {
        contentFieldClass.remove("error");
        syntaxMsgStyle.display = "none";
        uniquenessMsgStyle.display = "none";
      } else {
        contentFieldClass.add("error");
        //Don't show both error messages together
        if (this.syntaxError) {
          syntaxMsgStyle.display = "";
          uniquenessMsgStyle.display = "none";
        } else {
          syntaxMsgStyle.display = "none";
          uniquenessMsgStyle.display = "";
        }
      }
    }
  }

  class IDFontSize extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("IDFontSize"),
        resetBtn: document.getElementById("resetIDFontSize"),
        setAllBtn: document.getElementById("setAllIDFontSize"),
        errorElement1: document.getElementById("IDFontSizeError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "IDFontSize";
    }

    static getOriginalValue() {
      return 0.7;
    }
  }

  class CheckWidth extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("checkWidth"),
        resetBtn: document.getElementById("resetCheckWidth"),
        setAllBtn: document.getElementById("setAllCheckWidth"),
        errorElement1: document.getElementById("checkWidthError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "checkWidth";
    }

    static getOriginalValue() {
      return 1.5;
    }
  }

  class CheckHeight extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("checkHeight"),
        resetBtn: document.getElementById("resetCheckHeight"),
        setAllBtn: document.getElementById("setAllCheckHeight"),
        errorElement1: document.getElementById("checkHeightError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "checkHeight";
    }

    static getOriginalValue() {
      return 1.5;
    }
  }

  class CheckFontSize extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("checkFontSize"),
        resetBtn: document.getElementById("resetCheckFontSize"),
        setAllBtn: document.getElementById("setAllCheckFontSize"),
        errorElement1: document.getElementById("checkFontSizeError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "checkFontSize";
    }

    static getOriginalValue() {
      return 0.8;
    }
  }

  class RemoveFontSize extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("removeFontSize"),
        resetBtn: document.getElementById("resetRemoveFontSize"),
        setAllBtn: document.getElementById("setAllRemoveFontSize"),
        errorElement1: document.getElementById("removeFontSizeError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "removeFontSize";
    }

    static getOriginalValue() {
      return 0.3;
    }
  }

  class PointHeight extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("pointHeight"),
        resetBtn: document.getElementById("resetPointHeight"),
        setAllBtn: document.getElementById("setAllPointHeight"),
        errorElement1: document.getElementById("pointHeightError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "pointHeight";
    }

    static getOriginalValue() {
      return 2.5;
    }
  }

  class LetterFontSize extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("letterFontSize"),
        resetBtn: document.getElementById("resetLetterFontSize"),
        setAllBtn: document.getElementById("setAllLetterFontSize"),
        errorElement1: document.getElementById("letterFontSizeError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "letterFontSize";
    }

    static getOriginalValue() {
      return 1.8;
    }
  }

  class PhoneticFontSize extends StrictPositiveField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("phoneticFontSize"),
        resetBtn: document.getElementById("resetPhoneticFontSize"),
        setAllBtn: document.getElementById("setAllPhoneticFontSize"),
        errorElement1: document.getElementById("phoneticFontSizeError"),
        errorElement2: undefined
      };
      super(inputObj);
    }

    static getFieldName() {
      return "phoneticFontSize";
    }

    static getOriginalValue() {
      return 0.6;
    }
  }

  //Create root level objects
  const stationList = new StationList();
















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

    freader = new FileReader();
    freader.onload = function () {
      var xmlParser, xmlobj, parsererrorNS, mapFileScale, globalScale, courseNodes, courseNodesId, courseNodesNum, tableRowNode, tableColNode, tableContentNode, layoutRowNode, selectOptionNode, existingRows, existingRowID, existingRow, otherNode, leftcoord, bottomcoord, courseControlNode, controlNode, controlsSkipped, numProblems, stationNameRoot, courseScale;
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
        window.alert("Could not read map scale.");
        return;
      }
      mapFileScale = Number(otherNode.getAttribute("scale"));
      if (!(mapFileScale > 0)) {
        window.alert("Could not read map scale.");
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

          //Create first column - station name + hidden values
          tableColNode = document.createElement("td");
          stationNameRoot = courseNodes[courseNodesId].getElementsByTagName("name")[0].textContent.slice(0,-2);
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
            tableColNode.getElementsByClassName("circlex")[0].innerHTML += (0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("x")) - leftcoord) * mapFileScale / courseScale).toString();
            tableColNode.getElementsByClassName("circley")[0].innerHTML += (0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("y")) - bottomcoord) * mapFileScale / courseScale).toString();

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
          existingRows = document.getElementById("courseTableBody").getElementsByClassName("courseOrder");
          existingRowID = 0;
          for (;;) {
            existingRowID++;
            existingRow = existingRows[existingRowID];
            if (!existingRow) {
              document.getElementById("courseTableBody").appendChild(tableRowNode);
              document.getElementById("layoutTableBody").appendChild(layoutRowNode);
              break;	//No more rows to consider
            }
            if (Number(existingRow.innerHTML) > Number(courseNodes[courseNodesId].getAttribute("order"))) {
              document.getElementById("courseTableBody").insertBefore(tableRowNode, existingRow.parentElement);
              //Need to update existingRow to match layout table. There is one extra header row in the tbody.
              existingRow = document.getElementById("layoutTableBody").getElementsByTagName("tr")[existingRowID + 1];
              document.getElementById("layoutTableBody").insertBefore(layoutRowNode, existingRow.parentElement);
              break;	//Current row needs to be inserted before existingRow
            }
          }
        }
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
    };
    freader.onerror = function () { window.alert("Could not read Purple Pen file. Try reselecting it, then click Reload."); };
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
    fileobj = fileInput.files[0];
    if (fileobj) {
      fname = fileobj.name;
      freader = new FileReader();
      freader.onload = function () {
        //Populates station tables from previous LaTeX parameters file
        var fileString, startPos, endPos, subString, varArray, fields, rowId, numRows, classRoot;

        fileString = freader.result;

        //Indicate that parameters data is currently saved - unedited opened file
        paramsSaved = true;

        //Show station
        startPos = fileString.indexOf("\\def\\ShowStationList{{");
        if (startPos >= 0) {
          endPos = fileString.indexOf("}}", startPos);
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
          endPos = fileString.indexOf("}}", startPos);
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
            if (varArray[rowId] !== "") {
              fields[rowId].value = Number(varArray[rowId]);
            }
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
            if (varArray[rowId] !== "") {
              fields[rowId + 1].selectedIndex = parseInt(varArray[rowId], 10);
            }
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
            if (varArray[rowId] !== "") {
              fields[rowId + 1].value = Number(varArray[rowId]) * 2;
            }
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
            if (varArray[rowId] !== "") {
              fields[rowId + 1].value = Number(varArray[rowId]);
            }
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
            endPos = fileString.indexOf("}}", startPos);
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
                  fields[rowId + 2].value = Number(varArray[rowId].slice(1,-3));
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
          subString = fileString.substr(startPos + 16, 1);
          if (subString == "1") {
            document.getElementById("debugCircle").checked = true;
          } else {
            document.getElementById("debugCircle").checked = false;
          }
        } else {
          document.getElementById("debugCircle").checked = false;
        }
      };
      freader.onerror = function (err) {
        if (err.name == undefined) {
          window.alert("Could not read file due to an unknown error. This occurs on Safari for files containing the % symbol - try deleting all of them.");
        } else {
          window.alert("Could not read file: " + err.toString());
        }
      };
      freader.readAsText(fileobj);   //Reads as UTF-8
    }
  }

  function generateLaTeX() {
    //Generates LaTeX parameters file
    //Returns string with error message or "ok" if no errors

    var rtnstr, tableRows, layoutRows, numTableRows, tableRowID, contentField, numStations, maxProblems, numProblems, showStationList, numProblemsList, stationName, stationNameList, numKites, kitesList, zeroesList, headingList, shapeList, mapSize, sizeList, briefingWidthList, scaleList, contourList, mapFileList, mapPageList, mapxList, mapyList, CDsFileList, CDsPageList, CDsxList, CDsyList, controlsSkipped, CDsxCoord, CDsyCoord, CDsHeightList, CDsWidthList, CDsaFontList, CDsbFontList, fileName, showPointingBoxesList, pointingBoxWidthList, pointingBoxHeightList, pointingLetterFontList, pointingPhoneticFontList, stationIDFontList, checkBoxWidthList, checkBoxHeightList, checkNumberFontList, checkRemoveFontList, fileString, iterNum, CDsxCoordBase, CDsyCoordBase, CDsWidthBase, CDsHeightBase, parametersBlob;

    rtnstr = "ok";

    tableRows = document.getElementById("courseTableBody").getElementsByTagName("tr");
    layoutRows = document.getElementById("layoutTableBody").getElementsByTagName("tr");
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
      if (numStations > 1) {
        //Insert commas in lists
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
        rtnstr = "The number of kites for station " + stationName + " must be an integer between 1 and 6.";
        contentField.focus();
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
        rtnstr = "The heading of station " + stationName + " must be a number.";
        contentField.focus();
      } else {
        headingList += contentField.value;	//Don't add quotes
      }

      //Map shape
      contentField = tableRows[tableRowID].getElementsByClassName("mapShape")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The map shape for station " + stationName + " must be specified.";
        contentField.focus();
      } else {
        shapeList += contentField.selectedIndex;	//Don't add quotes
      }

      //Map size
      contentField = tableRows[tableRowID].getElementsByClassName("mapSize")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The map size for station " + stationName + " must be > 0 and <= 12.";
        contentField.focus();
      }	else if (contentField.value == 0) {
        rtnstr = "The map size for station " + stationName + " must be strictly greater than 0.";
        contentField.focus();
      } else {
        mapSize = Number(contentField.value);
        sizeList += 0.5 * mapSize;	//Don't add quotes
        briefingWidthList += "\"" + (0.7 * contentField.value) + "cm\"";
      }

      //Map scale
      contentField = tableRows[tableRowID].getElementsByClassName("mapScale")[0];
      if (contentField.checkValidity() == false || contentField.value == 0) {
        rtnstr = "The map scale for station " + stationName + " must be strictly greater than 0.";
        contentField.focus();
      } else {
        scaleList += contentField.value;	//Don't add quotes
      }

      //Map contour interval
      contentField = tableRows[tableRowID].getElementsByClassName("contourInterval")[0];
      if (contentField.checkValidity() == false || contentField.value == 0) {
        rtnstr = "The contour interval for station " + stationName + " must be strictly greater than 0.";
        contentField.focus();
      } else {
        contourList += contentField.value;	//Don't add quotes
      }

      //Map and control description files
      fileName = "\"Maps\"";
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

      //Layout parameters. Remember to account for extra header row.

      //Station name font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("IDFontSize")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The name font size for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        stationIDFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box width
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkWidth")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The page order box width for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        checkBoxWidthList += contentField.value;
      }

      //Check box height
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkHeight")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The page order box height for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        checkBoxHeightList += contentField.value;
      }

      //Check box number font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("checkFontSize")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The page number font size for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        checkNumberFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box remove text font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("removeFontSize")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The <em>Remove</em> font size for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        checkRemoveFontList += "\"" + contentField.value + "cm\"";
      }

      //Pointing box height
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("pointHeight")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The pointing box height for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
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
        rtnstr = "The pointing box letter font size for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        pointingLetterFontList += "\"" + contentField.value + "cm\"";
      }

      //Check box remove text font size
      contentField = layoutRows[tableRowID + 1].getElementsByClassName("phoneticFontSize")[0];
      if (contentField.checkValidity() == false) {
        rtnstr = "The pointing box phonetic font size for station " + stationName + " must be between 0 and 29.7.";
        contentField.focus();
      } else {
        pointingPhoneticFontList += "\"" + contentField.value + "cm\"";
      }
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

    //Create file
    parametersBlob = new Blob([fileString], { type: "text/plain" });

    return {str: rtnstr, file:parametersBlob};
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
    downloadFile(rtn.file, "TemplateParameters.tex");

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

  function generatePDF(btn) {
    var statusBox, paramRtn, resourceNames, resourceFileArray, resourceURLs, promiseArray;

    function compileLaTeX(source_code, resourceURLs, resourceNames, btn) {
      var statusBox, texlive, pdftex, texliveEvent;

      texliveEvent = (function(msg) {
        //Called everytime a status message is outputted by TeXLive. Worker thread will crash shortly after an error, so all handling to be done here.
        var logContent, numEvents, statusBox;
        logContent = [];
        numEvents = 0;
        statusBox = document.getElementById("compileStatus");
        return (msg) => {
          var logStr, rowId, logBlob, downloadElement, logURL;
          console.log(msg);
          logContent.push(msg);
          numEvents++;
          if (btn.disabled === true) {
            if (msg.includes("no output PDF file produced!")) {
              //TeXLive encountered an error -> handle it
              if (logContent[numEvents-3] === "!pdfTeX error: /latex (file ./Maps.pdf): PDF inclusion: required page does not ") {
                statusBox.innerHTML = "Failed to compile map cards. The PDF of maps does not contain enough pages. Please follow the instructions in step 8 carefully and try again.";
              } else if (logContent[numEvents-3] === "!pdfTeX error: /latex (file ./CDs.pdf): PDF inclusion: required page does not e") {
                statusBox.innerHTML = "Failed to compile map cards. The PDF of control descriptions does not contain enough pages. Please follow the instructions in step 9 carefully and try again.";
              } else {
                statusBox.innerHTML = "Failed to compile map cards due to an unknown error. Please seek assistance.";
                //Display log
                logStr = "";
                for (rowId = 0; rowId < numEvents; rowId++) {
                  logStr += logContent[rowId] + '\n';
                }
                logBlob = new Blob([logStr], { type: "text/plain" });
                //Revoke old URL
                downloadElement = document.getElementById("viewLog");
                logURL = downloadElement.href;
                if (logURL) {
                  URL.revokeObjectURL(logURL);
                }
                logURL = URL.createObjectURL(logBlob);
                downloadElement.href = logURL;
                downloadElement.hidden = false;
              }
              //Cleanup
              btn.disabled = false;
              texlive.terminate();
            } else {
              statusBox.innerHTML = "Preparing map cards (" + numEvents.toString() + ").";
            }
          }
          return logContent;
        };
      }());

      //Update status
      statusBox = document.getElementById("compileStatus");
      statusBox.innerHTML = "Preparing map cards. This may take a minute.";

      texlive = new TeXLive("texlive.js/");
      pdftex = texlive.pdftex;

      //Use promises of promise.js
      pdftex.set_TOTAL_MEMORY(80*1024*1024).then(function() {
        promise.join(resourceNames.map(function(name, index) {
          return pdftex.FS_createLazyFile('/', name, resourceURLs[index], true, false);
        })).then(function() {
          pdftex.on_stdout = texliveEvent;
          pdftex.on_stderr = texliveEvent;
          return pdftex.compile(source_code);
        }).then(function(pdf_dataurl) {
          var downloadElement;
          if (pdf_dataurl === false) {
            statusBox.innerHTML = "Failed to compile map cards. Please seek assistance.";
          } else {
            //Save PDF
            pdftex.FS_readFile("./input.pdf").then((outfile) => {
              var outBlob, outURL, outLen, outArray, index;
              downloadElement = document.getElementById("savePDF");
              //Revoke old URL
              outURL = downloadElement.href;
              if (outURL) {
                URL.revokeObjectURL(outURL);
              }
              //Make new one, but first need file stream as array buffer
              outLen = outfile.length;
              outArray = new Uint8Array(outLen);
              for (index = 0; index < outLen; index++) {
                //Populate array with unicode value of each character
                outArray[index] = outfile.charCodeAt(index);
              }
              outBlob = new Blob([outArray], { type: "application/pdf" });
              outURL = URL.createObjectURL(outBlob);
              downloadElement.href = outURL;
              downloadElement.hidden = false;
              downloadElement.click();

              //Final clean
              texlive.terminate();
              btn.disabled = false;
              statusBox.innerHTML = "Map cards PDF produced successfully and is now in your downloads folder.";
            });
          }
        });
      });
    }

    //Disable button to avoid double press
    btn.disabled = true;

    //Status line
    statusBox = document.getElementById("compileStatus");
    statusBox.innerHTML = "Loading files. If this message persists, there is a problem. Seek assistance.";
    document.getElementById("savePDF").hidden = true;
    document.getElementById("viewLog").hidden = true;

    //Make LaTeX parameters file
    paramRtn = generateLaTeX();
    if (paramRtn.str !== "ok") {
      statusBox.innerHTML = paramRtn.str;
      btn.disabled = false;
      return;
    }

    //Read maps and CDs PDFs
    resourceNames = ["TemplateParameters.tex", "Maps.pdf", "CDs.pdf"];
    resourceFileArray = [paramRtn.file, document.getElementById("coursePDFSelector").files[0], document.getElementById("CDPDFSelector").files[0]];
    //Load each resource file and get a URL for each
    //Read them using promises native in Javascript
    promiseArray = resourceFileArray.map((fileobj) => {
      return new Promise(function(resolve, reject) {
        var filename = fileobj.name;
        var freader = new FileReader();
        freader.onload = function() {
          resolve(this.result);
        };
        freader.onerror = reject;
        freader.readAsDataURL(fileobj);
      });
    });
    Promise.all(promiseArray).catch((err) => {
      statusBox.innerHTML = "Either the course maps PDF or control descriptions PDF file is missing. Try selecting them again.";
      return Promise.reject("handled");
    }).then((result) => {
      //Download a copy of parameters file to save for later
      if (document.getElementById("autoSave").checked === true && paramsSaved === false) {
        downloadFile(paramRtn.file, "TemplateParameters.tex");

        //Indicate that parameters data is currently saved
        paramsSaved = true;
      }

      resourceURLs = result.slice(0);
      //Load LaTeX code
      return fetch("TCTemplate.tex");
    }).then((response) => {
      if (response.ok === true) {
        return response.text();
      } else {
        throw response.status;
      }
    }).then((sourceCode) => {
      compileLaTeX(sourceCode, resourceURLs, resourceNames, btn);
    }, (err) => {
      if (err !== "handled") {
        statusBox.innerHTML = "Failed to load TCTemplate.tex: " + err;
      }
      //Enable generate PDF button
      btn.disabled = false;
    });
  }

  //Initialisation
  //TODO: From legacy UI
  // document.getElementById("stationProperties").hidden = true;
  document.getElementById("savePDF").hidden = true;
  document.getElementById("viewLog").hidden = true;

  //Initialise variables. Do it this way rather than in HTML to avoid multiple hardcodings of the same initial values.
  stationList.add(stationList.default);
  stationList.refresh(false);

  //Event listeners - always call with arrow functions to ensure this doesn't point to calling DOM element
  stationList.createBtnListeners();
  stationList.defaultRadio.addEventListener("change", () => { stationList.refresh(true); });
  stationList.stationRadio.addEventListener("change", () => { stationList.refresh(true); });
  stationList.activeItem.taskList.createBtnListeners();
  document.getElementById("setAllCore").addEventListener("click", () => { stationList.applyAll("setAll", stationList.default.coreFields); });
  document.getElementById("setAllCustomLayout").addEventListener("click", () => { stationList.applyAll("setAll", stationList.default.customLayoutFields); });
  document.getElementById("resetAllCustomLayout").addEventListener("click", () => { stationList.applyAll("resetValue", stationList.default.customLayoutFields); });
  for (const field of stationList.default.fieldNames) {
    let inputEvent = "input";
    if (stationList.default[field].inputElement.tagName === "SELECT") {
      inputEvent = "change";
    }
    stationList.default[field].inputElement.addEventListener(inputEvent, () => { stationList.activeItem[field].saveInput(); });
    if (stationList.default[field].resetBtn !== undefined) {
      stationList.default[field].resetBtn.addEventListener("click", () => { stationList.activeItem[field].resetValue(); });
      stationList.default[field].setAllBtn.addEventListener("click", () => { stationList.activeItem[field].setAll(); });
    }
  }

  //Make required functions and objects globally visible
  //TODO:This can probably all be removed: now using event listeners
  return {
    stationList: stationList,
    loadppen: loadppen,
    loadTeX: loadTeX,
    saveParameters: saveParameters,
    setAllCourses: setAllCourses,
    resetAllLayout: resetAllLayout,
    setAllLayout: setAllLayout,
    resetField: resetField,
    generatePDF: generatePDF
  };
})();
