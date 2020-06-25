"use strict";

//FIXME: Chrome has bug that defer script loading doesn't work with XHTML
//This function is currently called by DOMContentLoaded event
//Once fixed, load messages.js then this script both with defer, delete domLoad and change tcTemplate() to anonymous IIFE
domLoad.then(async () => {
  //Wait until tcTemplateMsg is defined, then run tcTemplate to make page dynamic
  while (typeof tcTemplateMsg === "undefined") {
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  tcTemplate();
});

//Keep all functions private and put those with events in HTML tags in a namespace
function tcTemplate() {
  // (() => then add opening brace
  let courseList, stationList, pdfjsLib;

  //Keep track of whether an input file has been changed in a table to disable autosave
  var paramsSaved = true;

  //Check browser supports required APIs
  //Items to check:
  //HTMLCanvasElement.toBlob() (Edge >= 79, Firefox >= 19, Safari >= 11)
  //details (Edge >= 79, Firefox >= 49, Safari >= 6)
  //async (Edge >= 15, Firefox >= 52, Safari >= 10.1) - no easy way to check and very few browswers to exclusively trip up
  if (
    typeof document.createElement("canvas").toBlob === "function" &&
    typeof document.createElement("details").open === "boolean"
    //If crashes, API error message is still displayed
  ) {
    document.getElementById("missingAPIs").hidden = true;   //Hide error message that shows by default
  } else {
    document.getElementById("mainView").hidden = true;
    return;
  }

  //Object structure:
  //const Course list->Course->Field
  //const Station list->Station->Field or Task list->Task->Field or Kite list->Kite->Field

  //Classes are not hoisted, so declare them first
  //Class types
  class IterableList {
    constructor(obj) {
      // this.selector = obj.selector;
      // this.btnList = ["addBtn", "deleteBtn", "upBtn", "downBtn"];
      // this.btnList.forEach((btn) => { this[btn] = obj[btn]; });
      if (obj.itemInFocus === undefined) { obj.itemInFocus = 0; };
      this.defaultInFocus = obj.defaultInFocus;
      this.itemInFocus = obj.itemInFocus;
      this.items = [];
      this.counterField = obj.counterField;

      //Event listener functions
      //Always use arrow functions to ensure this points to the IterableList rather than the calling DOM element
      //Save named listener functions so that the event can be removed later
      this.eventListeners = {
        addBtn: () => { this.add(); this.updateCount() },
        deleteBtn: () => { this.activeItem.deleteThis(true); this.updateCount(); },
        upBtn: () => { this.activeItem.move(-1); },
        downBtn: () => { this.activeItem.move(1); }
      };
    }

    get activeItem() {
      if (this.defaultInFocus) {
        return this.default;
      } else {
        return this.items[this.itemInFocus];
      }
    }

    checkValidity(recheckFields = true) {
      if (recheckFields) { this.items.forEach((item) => { item.checkValidity(true); }); }
      this.valid = this.items.every((item) => item.valid);
    }

    addBtnListeners() {
      //Do not call until activeItem is defined
      for (const btn of this.btnList) {
        this[btn].addEventListener("click", this.eventListeners[btn]);
      }
    }

    static addListeners() {
      const listObj = (this === CourseList) ? courseList : stationList;
      this.itemClass.fieldClasses.forEach((field, index) => {
        if (field.inputElement.tagName === undefined) {
          //Radio buttons
          Object.values(field.inputElement).forEach((element) => {
            element.addEventListener("change", () => { listObj.activeItem[this.itemClass.fieldNames[index]].saveInput(); });
          });
        } else {
          const inputEvent = (field.inputElement.tagName === "SELECT" || field.inputElement.type === "checkbox") ? "change" : "input";
          field.inputElement.addEventListener(inputEvent, () => { listObj.activeItem[this.itemClass.fieldNames[index]].saveInput(); });
        }
        if (field.autoElement !== undefined) {
          field.autoElement.addEventListener("change", () => { listObj.activeItem[this.itemClass.fieldNames[index]].auto.saveInput(); });
        }
      });
    }

    add() {
      //New item initially takes default values
      const newItem = this.newItem();

      //Name the new item
      newItem.itemName.value = (this.items.length + 1).toString();
      newItem.optionElement.textContent = newItem.itemName.value;

      //Check validity of data in new item including all fields
      newItem.checkValidity(true);

      //Add to the end
      this.items.push(newItem);
      this.constructor.selector.appendChild(newItem.optionElement);

      //Try to change to new item. The add station button is disabled when on defaults.
      newItem.optionElement.selected = true;
      this.refresh(true);
    }

    applyAll(action, fieldList) {
      //Applies specified action function to all fields with reset buttons
      fieldList.forEach((field) => {
        if (this.activeItem[field].resetBtn !== undefined) {
          this.activeItem[field][action]();
        }
      });
    }

    blockNameError(active = true) {
      //Only permit change of course if the name is valid and unique
      const block = active && !this.activeItem.itemName.valid;
      if (block) {
        //Abort
        alert(this.constructor.nameErrorAlert);
        //Change selectors back to original value
        this.constructor.selector.selectedIndex = this.itemInFocus;
      }
      return block;
    }

    newItem() {
      return new this.constructor.itemClass({
        parentObj: this,
        copy: (this.default === undefined) ? undefined : this.default
      });
    }

    showHideMoveBtns() {
      //Show the delete button only if more than one option
      if (this.constructor.selector.length > 1) {
        this.constructor.deleteBtn.disabled = false;
      } else {
        this.constructor.deleteBtn.disabled = true;
      }

      //Shows or hides the move up and move down buttons
      if (this.constructor.selector.selectedIndex === 0) {
        //First station, so can't move it up
        this.constructor.upBtn.disabled = true;
      } else {
        this.constructor.upBtn.disabled = false;
      }
      if (this.constructor.selector.selectedIndex === this.constructor.selector.length - 1) {
        //First station, so can't move it up
        this.constructor.downBtn.disabled = true;
      } else {
        this.constructor.downBtn.disabled = false;
      }
    }

    updateCount() {
      if (this.counterField !== undefined) {
        this.counterField.save(this.items.length);
        this.counterField.refreshInput();
      }
    }
  }
  IterableList.btnList = ["addBtn", "deleteBtn", "upBtn", "downBtn"];

  class IterableItem {
    constructor(obj) {
      this.parentList = obj.parentObj;
      if (!this.parentList.defaultInFocus) {
        this.optionElement = document.createElement("option");
      }
      //TODO: This seems outdated
      if (this.constructor.itemType === "station") {
        this.station = this;
      } else {
        this.station = this.parentList.parentItem;
      }
    }

    get index() {
      let thisIndex = 0;
      while (this.parentList.items[thisIndex] !== this) { ++thisIndex; }
      return thisIndex;
    }

    static populateFieldNames() {
      this.fieldNames = this.fieldClasses.map((className) => className.fieldName);
    }

    isActive() {
      return this.parentList.activeItem === this;
    }

    isDefault() {
      return this.parentList.default === this;
    }

    // isHidden() {
    //   //TODO: Move to more specialised classes
    //   return false;
    // }

    createFields(copy) {
      if (typeof copy === "undefined") {
        //Create an object to pass undefined parameters => triggers default values specified in class
        copy = {};
        this.constructor.fieldNames.forEach((field) => { copy[field] = { value: undefined }; });
      }
      //WARNING: need to make a copy of any sub objects, otherwise will still refer to same memory
      this.constructor.fieldNames.forEach((field, index) => {
        this[field] = new this.constructor.fieldClasses[index]({
          parentItem: this,
          value: copy[field].value
        });
      });
    }

    deleteThis(checkFirst = true) {
      //Deletes this item
      if (checkFirst) {
        if (!confirm(tcTemplateMsg.confirmDelete)) {
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
      paramsSaved = false;

      //Check target is in bounds
      const newPos = this.index + offset;
      const arrayLength = this.parentList.items.length
      if (newPos < 0 || newPos >= arrayLength) {
        throw new Error("Moving station to position beyond bounds of array");
      }
      if (!Number.isInteger(newPos)) {
        throw new Error("Station position offset is not an integer");
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

    refreshAllInput() {
      this.constructor.fieldNames.forEach((field) => { this[field].refreshInput(true); });
    }
  }

  class Field {
    //Generic field: string or select
    constructor(obj) {
      this.parentItem = obj.parentItem;
      this.parentList = this.parentItem.parentList;
      // if (this.parentItem.getItemType() === "station") {
      //   this.station = this.parentItem;
      // } else {
      //   this.station = this.parentList.parentItem;
      // }
      // this.stationList = this.station.parentList; //Shortcut
      this.inputElement = (obj.inputElement === undefined) ? this.constructor.inputElement : obj.inputElement;
      this.value = (obj.value === undefined) ? this.constructor.originalValue : obj.value;
      if (this.constructor.autoElement !== undefined) {
        this.auto = new AutoField({
          parentItem: obj.parentItem, //TODO: Check this
          parentField: this,
          inputElement: this.constructor.autoElement
        });
      }
    }

    get inputValue() {
      return this.inputElement.value;
    }

    set inputValue(val) {
      this.inputElement.value = val;
    }

    autoEnabled() {
      return (this.auto !== undefined) ? this.auto.value : false;
    }

    isDuplicate(ignoreHidden = false) {
      //Determines whether this value is duplicated at another item in this list. Ignores hidden stations on request.

      //Return false if the tested index is ignored
      if (ignoreHidden && this.station.showStation.value === false) { return false; }

      if (ignoreHidden) {
        return this.parentList.items.some((item) => (item[this.fieldName].value === this.value && item !== this.parentItem && this.station.showStation.value === true));
      } else {
        return this.parentList.items.some((item) => (item[this.fieldName].value === this.value && item !== this.parentItem));
      }
    }

    matchesAll(ignoreHidden = false) {
      //Returns true if this field matches those on all other items in this list, excluding number fields set to NaN. Ignores NaN if number. Ignores hidden stations on request.

      //Return true if the tested index is ignored
      if (ignoreHidden && this.parentItem.isHidden()) { return true; }

      //FIXME: The NaN condition seems dodgy. What if field is not meant to be a number?
      const isMatch = (item) => item[this.constructor.fieldName].value === this.value || Number.isNaN(item[this.constructor.fieldName].value);

      if (ignoreHidden) {
        return this.parentList.items.every((item) => isMatch(item) || item.isHidden());
      } else {
        return this.parentList.items.every(isMatch);
      }
    }

    refreshInput(alsoAuto = false) {
      //Parameters are to avoid circular arguments, hence infinite stacks
      //Updates input element value - no user input
      if (alsoAuto && this.auto !== undefined) { this.auto.refreshInput(false); }
      this.inputValue = this.value;
      this.updateMsgs();
    }

    resetValue() {
      if (this.parentItem.isDefault()) {
        //Resets to original value
        this.save(this.constructor.originalValue);
      } else {
        //Resets to current default value
        this.save(this.parentList.default[this.fieldName].value);
      }
      this.refreshInput(true); //TODO: Check value of alsoAuto
      this.parentItem.checkValidity(false);
    }

    saveInput() {
      //Call when the value of the input element is updated by the user
      this.save(this.inputValue);
      this.updateMsgs();
      this.parentItem.checkValidity(false);
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
      //Sets this field to this value in all items
      this.parentList.items.forEach((item) => {
        item[this.constructor.fieldName].save(this.value);
        item.checkValidity(false);
      });
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      if (this.valid) {
        contentFieldClass.remove("error");
        if (this.constructor.errorElement !== undefined) {
          this.constructor.errorElement.textContent = "";
        }
      } else {
        contentFieldClass.add("error");
        this.constructor.errorElement.textContent = this.errorMsg;
      }
    }

    //Empty functions, which may be overwritten in inheriting classes if some action is required
    checkValidity() {} //Most likely: field always valid
  }
  Field.originalValue = "";

  class BooleanField extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true;
    }

    get inputValue() {
      return this.inputElement.checked;
    }

    set inputValue(ticked) {
      this.inputElement.checked = ticked;
    }
  }
  BooleanField.originalValue = false;

  class AutoField extends BooleanField {
    constructor(obj) {
      super(obj);
      this.parentField = obj.parentField;
      this.calculated = this.parentField.constructor.originalValue;
    }

    updateMsgs() {
      if (this.value) {
        this.parentField.inputElement.readOnly = true;
        //Insert calculated value
        this.parentField.save(this.calculated); //Only checks validity if value changed, so always recheck
        this.parentField.refreshInput(false);
      } else {
        this.parentField.inputElement.readOnly = false;
      }
      this.parentField.checkValidity();
      this.parentField.updateMsgs();
    }
  }
  AutoField.originalValue = true;

  class NumberField extends Field {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.numberFieldError;
    }

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

    checkValidity() {
      this.valid = Number.isFinite(this.value) || this.parentItem.isNonDefaultHidden();
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
  NumberField.originalValue = NaN;

  class NonNegativeField extends NumberField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.nonNegativeFieldError;
    }

    checkValidity() {
      this.valid = (Number.isFinite(this.value) && this.value >= 0) || this.parentItem.isNonDefaultHidden();
    }
  }

  class StrictPositiveField extends NonNegativeField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.strictPositiveFieldError;
    }

    checkValidity() {
      //LaTeX crashes if font size is set to zero
      this.valid = (Number.isFinite(this.value) && this.value > 0) || this.parentItem.isNonDefaultHidden();
    }
  }

  class NaturalNumberField extends NonNegativeField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.naturalNumberFieldError;
    }

    checkValidity() {
      this.valid = (Number.isInteger(this.value) && this.value > 0) || this.parentItem.isNonDefaultHidden();
    }
  }

  class RadioField extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true;
    }
    get inputValue() {
      for (const val in this.inputElement) {
        if (this.inputElement[val].checked) { return val; }
      }
      return undefined;
    }

    set inputValue(val) {
      this.inputElement[val].checked = true;
    }

    updateMsgs() {}
  }

  class NameField extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true;
    }

    checkValidity() {
      if (!this.parentItem.isDefault()) {
        //Check syntax even if hidden to avoid dodgy strings getting into LaTeX
        const stringFormat = /^[A-Za-z0-9][,.\-+= \w]*$/;
        this.syntaxError = !(this.parentItem.isDefault() || stringFormat.test(this.value));
        this.duplicateError = !(this.parentItem.isDefault()) && this.isDuplicate(false);
        this.valid = !(this.syntaxError || this.duplicateError) || this.isDefault();
        if (!this.valid) { this.errorMsg = (this.syntaxError) ? tcTemplateMsg.nameSyntax : tcTemplateMsg.notUnique; }
        //Update station list
        this.parentItem.optionElement.textContent = this.value;
      } //Else remains valid
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      if (this.valid) {
        contentFieldClass.remove("error");
        this.constructor.errorElement.textContent = "";
      } else {
        contentFieldClass.add("error");
        this.constructor.errorElement.textContent = this.errorMsg;
      }
    }
  }
  NameField.fieldName = "itemName";

  //Course - in reverse order of dependency
  class CourseName extends NameField {}
  Object.assign(CourseName, {
    inputElement: document.getElementById("courseName"),
    errorElement: document.getElementById("courseNameError")
  });

  class TasksFile extends RadioField {
    updateMsgs() {
      this.parentItem.refreshAllInput(true);
    }
  }
  Object.assign(TasksFile, {
    fieldName: "tasksFile",
    originalValue: "newFile",
    inputElement: {
      newFile: document.getElementById("newTasksFile"),
      append: document.getElementById("appendTasksFile"),
      hide: document.getElementById("hideTasksFile")
    },
    resetBtn: document.getElementById("resetTasksFile"),
    setAllBtn: document.getElementById("setAllTasksFile")
  });

  class TasksTemplate extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true;
    }
  }
  Object.assign(TasksTemplate, {
    fieldName: "tasksTemplate",
    originalValue: "printA5onA4",
    inputElement: document.getElementById("tasksTemplate")
  });

  class AppendTasksCourse extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true;
    }
  }
  Object.assign(AppendTasksCourse, {
    fieldName: "appendTasksCourse",
    inputElement: document.getElementById("appendTasksCourse")
  });

  class Zeroes extends BooleanField {}
  Object.assign(Zeroes, {
    fieldName: "zeroes",
    inputElement: document.getElementById("zeroes"),
    resetBtn: document.getElementById("resetZeroes"),
    setAllBtn: document.getElementById("setAllZeroes")
  });

  class NumTasks extends NaturalNumberField {
    setAll() {
      //Sets the number of tasks in all stations
      if (this.valid) {
        for (const station of this.stationList.items) {
          if (this.value < station.numTasks.value) {
            if (!confirm(tcTemplateMsg.confirmRemoveTasks(station.itemName.value))) {
              continue;
            }
          }
          //FIXME: Complete this function
          station.numTasks.save(this.value);
          station.taskList.setNumItems();
          station.checkValidity(false);
        }
      }
    }
  }
  Object.assign(NumTasks, {
    fieldName: "numTasks",
    originalValue: 5,
    inputElement: document.getElementById("numTasks"),
    resetBtn: document.getElementById("resetNumTasks"),
    setAllBtn: document.getElementById("setAllNumTasks"),
    errorElement: document.getElementById("numTasksError")
  });

  class MapScale extends StrictPositiveField {
    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      if (this.valid) {
        contentFieldClass.remove("error");
        //Check whether all values are the same and (equal 4000 or 5000)
        //No warning if non-default hidden station or (4000/5000 and (default or all stations same))
        if (
          this.parentItem.isNonDefaultHidden() ||
          ((this.value === 4000 || this.value === 5000) && (this.parentItem.isDefault() || this.matchesAll(true)))
        ) {
          contentFieldClass.remove("warning");
          this.constructor.errorElement.textContent = "";
        } else {
          contentFieldClass.add("warning");
          if (this.value !== 4000 && this.value !== 5000) {
          this.constructor.errorElement.textContent = tcTemplateMsg.mapScaleRule;
        } else {
          this.constructor.errorElement.textContent = tcTemplateMsg.notSameForAllCourses;
        }
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        this.constructor.errorElement.textContent = this.errorMsg;
      }
    }
  }
  Object.assign(MapScale, {
    fieldName: "mapScale",
    originalValue: 4000,
    inputElement: document.getElementById("mapScale"),
    resetBtn: document.getElementById("resetMapScale"),
    setAllBtn: document.getElementById("setAllMapScale"),
    errorElement: document.getElementById("mapScaleError")
  });

  class ContourInterval extends StrictPositiveField {
    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      //Check validity - don't bother if station is non-default hidden
      if (this.valid) {
        contentFieldClass.remove("error");
        //Check whether all values are the same
        //No warning if non-default hidden station or defaults or all stations same
        if (this.parentItem.isDefault() || this.matchesAll(true)) {
          contentFieldClass.remove("warning");
          this.constructor.errorElement.textContent = "";
        } else {
          contentFieldClass.add("warning");
          this.constructor.errorElement.textContent = tcTemplateMsg.notSameForAllCourses;
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        this.constructor.errorElement.textContent = this.errorMsg;
      }
    }
  }
  Object.assign(ContourInterval, {
    fieldName: "contourInterval",
    inputElement: document.getElementById("contourInterval"),
    resetBtn: document.getElementById("resetContourInterval"),
    setAllBtn: document.getElementById("setAllContourInterval"),
    errorElement: document.getElementById("contourIntervalError")
  });

  class MapShape extends Field {
    constructor(obj) {
      super(obj);
      this.valid = true; //Always valid
    }

    updateMsgs() {
      //Check all values the same - don't bother if station is hidden or is the defaults
      //No warning if station hidden or Defaults or all stations same
      const contentFieldClass = this.inputElement.classList;
      if (this.matchesAll(true) || this.parentItem.isDefault()) {
        contentFieldClass.remove("warning");
        this.constructor.errorElement.textContent = "";
      } else {
        contentFieldClass.add("warning");
        this.constructor.errorElement.textContent = tcTemplateMsg.notSameForAllCourses;
      }

      //Update labels for map size description
      const sizeTypeElement = document.getElementById("mapSizeType");
      if (this.value === "circle") {
        sizeTypeElement.textContent = tcTemplateMsg.diameter;
      } else {
        sizeTypeElement.textContent = tcTemplateMsg.sideLength;
      }
    }
  }
  Object.assign(MapShape, {
    fieldName: "mapShape",
    originalValue: "circle",
    inputElement: document.getElementById("mapShape"),
    resetBtn: document.getElementById("resetMapShape"),
    setAllBtn: document.getElementById("setAllMapShape"),
    errorElement: document.getElementById("mapShapeRule")
  });

  class MapSize extends NumberField {
    checkValidity() {
      this.valid = (Number.isFinite(this.value) && this.value > 0 && this.value <= 12) || this.parentItem.isNonDefaultHidden();
      if (this.valid) {
        if (this.parentItem.enforceRules.value && this.value < 5) {
          this.valid = false;
          this.errorMsg = tcTemplateMsg.mapSizeRule;
        }
      } else if (this.auto.value) {
        this.errorMsg = tcTemplateMsg.awaitingData;
      } else {
        this.errorMsg = tcTemplateMsg.mapSizeError;
      }
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      if (this.valid) {
        contentFieldClass.remove("error");
        //Check whether all values are the same
        if (this.parentItem.isNonDefaultHidden() || this.parentItem.isDefault() || this.matchesAll(true)) {
          contentFieldClass.remove("warning");
          this.constructor.errorElement.textContent = "";
        } else {
          contentFieldClass.add("warning");
          this.constructor.errorElement.textContent = tcTemplateMsg.notSameForAllCourses;
        }
      } else {
        contentFieldClass.add("error");
        contentFieldClass.remove("warning");
        this.constructor.errorElement.textContent = this.errorMsg;
      }
    }
  }
  Object.assign(MapSize, {
    fieldName: "mapSize",
    inputElement: document.getElementById("mapSize"),
    autoElement: document.getElementById("mapSizeAuto"),
    resetBtn: document.getElementById("resetMapSize"),
    setAllBtn: document.getElementById("setAllMapSize"),
    errorElement: document.getElementById("mapSizeError")
  });

  class NumKites extends NumberField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.numKitesRule;
    }

    checkValidity() {
      this.valid = !this.parentItem.enforceRules.value || this.parentItem.isHidden() || this.value === 6;
    }
  }
  Object.assign(NumKites, {
    fieldName: "numKites",
    originalValue: 6,
    inputElement: document.getElementById("numKites"),
    resetBtn: document.getElementById("resetNumKites"),
    setAllBtn: document.getElementById("setAllNumKites"),
    errorElement: document.getElementById("numKitesError")
  });

  class EnforceRules extends BooleanField {}
  Object.assign(EnforceRules, {
    fieldName: "enforceRules",
    originalValue: true,
    inputElement: document.getElementById("enforceRules"),
    resetBtn: document.getElementById("resetEnforceRules"),
    setAllBtn: document.getElementById("setAllEnforceRules")
  });

  class IDFontSize extends StrictPositiveField {}
  Object.assign(IDFontSize, {
    fieldName: "IDFontSize",
    originalValue: 0.7,
    inputElement: document.getElementById("IDFontSize"),
    resetBtn: document.getElementById("resetIDFontSize"),
    setAllBtn: document.getElementById("setAllIDFontSize"),
    errorElement: document.getElementById("IDFontSizeError")
  });


  class CheckWidth extends NonNegativeField {}
  Object.assign(CheckWidth, {
    fieldName: "checkWidth",
    originalValue: 1.5,
    inputElement: document.getElementById("checkWidth"),
    resetBtn: document.getElementById("resetCheckWidth"),
    setAllBtn: document.getElementById("setAllCheckWidth"),
    errorElement: document.getElementById("checkWidthError")
  });

  class CheckHeight extends NonNegativeField {}
  Object.assign(CheckHeight, {
    fieldName: "checkHeight",
    originalValue: 1.5,
    inputElement: document.getElementById("checkHeight"),
    resetBtn: document.getElementById("resetCheckHeight"),
    setAllBtn: document.getElementById("setAllCheckHeight"),
    errorElement: document.getElementById("checkHeightError")
  });

  class CheckFontSize extends StrictPositiveField {}
  Object.assign(CheckFontSize, {
    fieldName: "checkFontSize",
    originalValue: "0.8",
    inputElement: document.getElementById("checkFontSize"),
    resetBtn: document.getElementById("resetCheckFontSize"),
    setAllBtn: document.getElementById("setAllCheckFontSize"),
    errorElement: document.getElementById("checkFontSizeError")
  });

  class RemoveFontSize extends StrictPositiveField {}
  Object.assign(RemoveFontSize, {
    fieldName: "removeFontSize",
    originalValue: 0.3,
    inputElement: document.getElementById("removeFontSize"),
    resetBtn: document.getElementById("resetRemoveFontSize"),
    setAllBtn: document.getElementById("setAllRemoveFontSize"),
    errorElement: document.getElementById("removeFontSizeError")
  });

  class PointHeight extends NonNegativeField {}
  Object.assign(PointHeight, {
    fieldName: "pointHeight",
    originalValue: 2.5,
    inputElement: document.getElementById("pointHeight"),
    resetBtn: document.getElementById("resetPointHeight"),
    setAllBtn: document.getElementById("setAllPointHeight"),
    errorElement: document.getElementById("pointHeightError")
  });

  class LetterFontSize extends StrictPositiveField {}
  Object.assign(LetterFontSize, {
    fieldName: "letterFontSize",
    originalValue: 1.8,
    inputElement: document.getElementById("letterFontSize"),
    resetBtn: document.getElementById("resetLetterFontSize"),
    setAllBtn: document.getElementById("setAllLetterFontSize"),
    errorElement: document.getElementById("letterFontSizeError")
  });

  class PhoneticFontSize extends StrictPositiveField {}
  Object.assign(PhoneticFontSize, {
    fieldName: "phoneticFontSize",
    originalValue: 0.6,
    inputElement: document.getElementById("phoneticFontSize"),
    resetBtn: document.getElementById("resetPhoneticFontSize"),
    setAllBtn: document.getElementById("setAllPhoneticFontSize"),
    errorElement: document.getElementById("phoneticFontSizeError")
  });

  class Course extends IterableItem {
    constructor(obj) {
      super(obj);
      //Fields - populate with values given in copyStation, if present
      this.createFields(obj.copy);

      //FIXME: Why keep these?
      // this.coreFields = coreFieldClasses.map((className) => className.getFieldName());
      // this.customLayoutFields = customLayoutClasses.map((className) => className.getFieldName());
    }

    isHidden() {
      return this.tasksFile.value === "hide";
    }

    isNonDefaultHidden() {
      //Returns true if is hidden and default is not in focus
      return this.isHidden() && !this.isDefault();
    }

    checkValidity(recheckFields = true) {
      if (recheckFields) {
        this.constructor.fieldNames.forEach((field) => { this[field].checkValidity(); });
      }
      this.valid = this.constructor.fieldNames.every((field) => this[field].valid);
      //Highlight errors in the station selector
      //Don't flag if uncalculated auto values
      if (!this.isDefault()) {
        if (this.valid || this.constructor.fieldNames.every((field) => this[field].valid || this[field].autoEnabled())) {
          this.optionElement.classList.remove("error");
        } else {
          this.optionElement.classList.add("error");
        }
      }
    }
  }
  Course.itemType = "course"; //FIXME: Is this needed?
  Course.customLayoutClasses = [
    IDFontSize,
    CheckWidth,
    CheckHeight,
    CheckFontSize,
    RemoveFontSize,
    PointHeight,
    LetterFontSize,
    PhoneticFontSize
  ];
  Course.fieldClasses = [
    CourseName,
    TasksFile,
    TasksTemplate,
    AppendTasksCourse,
    NumTasks,
    Zeroes,
    MapScale,
    ContourInterval,
    MapShape,
    MapSize,
    NumKites,
    EnforceRules
  ].concat(Course.customLayoutClasses);
  Course.populateFieldNames();

  class CourseList extends IterableList {
    constructor() {
      const obj = {
        defaultInFocus: false,
        itemInFocus: 0,
        counterField: undefined
      };
      super(obj);
      //Set up current/dynamic defaults
      this.default = this.newItem();
      this.constructor.courseRadio.checked = true;
    }

    static addListeners() {
      super.addListeners();
      [this.defaultRadio, this.courseRadio].forEach((radio) => {
        radio.addEventListener("click", () => { courseList.refresh(true); });
      });
    }

    refresh(blockOnNameError = true) {
      //blockOnNameError must be set to true when keeping the data

      //Other DOM elements
      // const setAllResetCSS = document.getElementById("showSetAllCSS");

      if (this.blockNameError(blockOnNameError)) {
        if (this.defaultInFocus) {
          this.constructor.defaultRadio.checked = true;
        } else {
          this.constructor.courseRadio.checked = true;
        }
        return;
      }

      //TODO: Hide task etc. selectors for station to be hidden
      // if (this.defaultInFocus === false) {
      //   this.activeItem.taskList.selector.style.display = "none";
      //   this.activeItem.taskList.removeBtnListeners();
      // }

      //Show/hide or enable/disable HTML elements according to new selected station
      this.itemInFocus = this.constructor.selector.selectedIndex;
      if (this.constructor.defaultRadio.checked) {
        //Setting defaults for all stations
        this.defaultInFocus = true;

        //Show/hide any buttons as required
        dynamicCSS.hide(".hideIfDefault");
        dynamicCSS.show(".setAllCourses");

        //Some fields need disabling
        this.constructor.selector.disabled = true;
        this.constructor.btnList.forEach((btn) => { this.constructor[btn].disabled = true; });
        CourseName.inputElement.disabled = true;
      } else {
        this.defaultInFocus = false;

        //Show/hide any buttons as required
        dynamicCSS.show(".hideIfDefault");
        dynamicCSS.hide(".setAllCourses");

        //Some fields may need enabling
        this.constructor.selector.disabled = false;
        this.constructor.addBtn.disabled = false;
        this.showHideMoveBtns();
        CourseName.inputElement.disabled = false;

        //TODO: Show new task etc. selectors
        // this.activeItem.taskList.selector.style.display = "";
        // this.activeItem.taskList.addBtnListeners();
      }

      //Populate with values for new selected station and show error/warning messages
      this.activeItem.refreshAllInput();
    }
  }
  Object.assign(CourseList, {
    selector: document.getElementById("courseSelect"),
    addBtn: document.getElementById("addCourse"),
    deleteBtn: document.getElementById("deleteCourse"),
    upBtn: document.getElementById("moveUpCourse"),
    downBtn: document.getElementById("moveDownCourse"),
    itemClass: Course,
    nameErrorAlert: tcTemplateMsg.courseNameAlert,
    //Radio button options
    defaultRadio: document.getElementById("defaultRadio"),
    courseRadio: document.getElementById("courseRadio")
  });

  class StationList extends IterableList {
    constructor() {
      const obj = {
        selector: document.getElementById("stationSelect"),
        addBtn: document.getElementById("addStation"),
        deleteBtn: document.getElementById("deleteStation"),
        upBtn: document.getElementById("moveUpStation"),
        downBtn: document.getElementById("moveDownStation"),
        defaultInFocus: false,
        itemInFocus: 0,
        counterField: undefined
      };
      super(obj);
    }

    refresh(storeValues = false) {
      //storeValues is a boolean stating whether to commit values in form fields to variables in memory.

      //Other DOM elements
      const setAllResetCSS = document.getElementById("showSetAllCSS");

      if (storeValues) {
        //Only permit change of station if the name is valid and unique
        if (this.activeItem.itemName.valid === false) {
          //Abort
          alert(tcTemplateMsg.stationNameAlert);
          //Change radio buttons and selectors back to original values
          if (this.defaultInFocus) {
            this.defaultRadio.checked = true;
          } else {
            this.stationRadio.checked = true;
            this.selector.selectedIndex = this.itemInFocus;
          }
          return;
        }
      }

      //Hide task etc. selectors for station to be hidden
      if (this.defaultInFocus === false) {
        this.activeItem.taskList.selector.style.display = "none";
        this.activeItem.taskList.removeBtnListeners();
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
        this.default.numTasks.inputElement.readOnly = false;
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
        this.activeItem.numTasks.inputElement.readOnly = true;

        //Show new task etc. selectors
        this.activeItem.taskList.selector.style.display = "";
        this.activeItem.taskList.addBtnListeners();
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
        itemInFocus: 0,
        counterField: parentObj.numTasks
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

    refresh(storeValues = false) {
      //storeValues is a boolean stating whether to commit values in form fields to variables in memory.
      if (storeValues) {
        //Only permit change of station if the name is valid and unique
        if (this.activeItem.itemName.valid === false) {
          //Abort
          alert(tcTemplateMsg.taskNameAlert);
          //Change radio buttons and selectors back to original values
          this.selector.selectedIndex = this.itemInFocus;
          return;
        }
      }

      this.itemInFocus = this.selector.selectedIndex;
      this.showHideMoveBtns();
      this.activeItem.refreshAllInput();
    }

    setNumItems() {
      //Forcibly adds/removes items from end of list to get required length
      //If this.numItems is NaN, both inequalities evaluate to false
      while (this.items.length < this.parentItem.numTasks.value) {
        this.add();
      }
      while (this.items.length > this.parentItem.numTasks.value) {
        this.items[this.items.length - 1].deleteThis(false);
      }
    }
  }

  //Types of items


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
      this.taskList = new TaskList(this);
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
        this.taskList.checkValidity(true);
      }

      this.valid = this.fieldNames.every((field) => this[field].valid) && this.taskList.valid;
      if (!this.isDefault()) {
        //Highlight errors in the station selector
        if (this.valid) {
          this.optionElement.classList.remove("error");
        } else {
          this.optionElement.classList.add("error");
        }
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
        CDPage,
        CDx,
        CDy,
        CDWidth,
        CDHeight,
        CDScale
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

    checkValidity(recheckFields = true) {
      if (recheckFields === true) {
        //Recheck validity of all fields
        for (const field of this.fieldNames) {
          this[field].checkValidity();
        }
      }

      this.valid = this.fieldNames.every((field) => this[field].valid);
      //Highlight errors in the task selector
      if (this.valid) {
        this.optionElement.classList.remove("error");
      } else {
        this.optionElement.classList.add("error");
      }
    }
  }

  //Fields


  class StationName extends Field {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("stationName"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("stationNameError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "itemName";
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
      if (this.valid) {
        contentFieldClass.remove("error");
        this.errorElement.innerHTML = "";
      } else {
        contentFieldClass.add("error");
        if (this.syntaxError) {
          this.errorElement.textContent = tcTemplateMsg.nameSyntax;
        } else {
          this.errorElement.textContent = tcTemplateMsg.notUnique;
        }
      }
    }
  }

  class ShowStation extends BooleanField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("showStationTasks"),
        resetBtn: document.getElementById("resetShowStation"),
        setAllBtn: document.getElementById("setAllShowStation"),
        errorElement: undefined
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



  class Heading extends NumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("heading"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("headingError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "heading";
    }

    checkValidity() {
      this.valid = (Number.isFinite(this.value) || this.station.showStation.value === false) || this.station.isDefault();
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
        errorElement: document.getElementById("taskNameError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "itemName";
    }

    checkValidity() {
      //Check syntax even if hidden to avoid dodgy strings getting into LaTeX
      const stringFormat = /^[A-Za-z0-9][,.\-+= \w]*$/;
      this.syntaxError = !(this.station.isDefault() || stringFormat.test(this.value));
      this.duplicateError = (!(this.station.isDefault()) && this.isDuplicate(false)) || this.value === "Kites" || this.value === "VP";
      this.valid = !(this.syntaxError || this.duplicateError);

      //Update task list
      this.parentItem.optionElement.innerHTML = this.value;
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      if (this.valid) {
        contentFieldClass.remove("error");
        this.errorElement.innerHTML = "";
      } else {
        contentFieldClass.add("error");
        //Don't show both error messages together
        if (this.syntaxError) {
          this.errorElement.innerHTML = tcTemplateMsg.nameSyntax;
        } else {
          this.errorElement.innerHTML = tcTemplateMsg.taskNameNotUnique;
        }
      }
    }
  }

  class CirclePage extends NaturalNumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("circlePage"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("circlePageError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "circlePage";
    }
  }

  class Circlex extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("circlex"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("circlexError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "circlex";
    }
  }

  class Circley extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("circley"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("circleyError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "circley";
    }
  }

  class CDPage extends NaturalNumberField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDPage"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDPageError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDPage";
    }
  }

  class CDx extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDx"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDxError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDx";
    }
  }

  class CDy extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDy"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDyError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDy";
    }
  }

  class CDWidth extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDWidth"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDWidthError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDWidth";
    }
  }

  class CDHeight extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDHeight"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDHeightError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDHeight";
    }
  }

  class CDScale extends NonNegativeField {
    constructor(parentObj, value) {
      const inputObj = {
        parentObj: parentObj,
        value: value,
        inputElement: document.getElementById("CDScale"),
        resetBtn: undefined,
        setAllBtn: undefined,
        errorElement: document.getElementById("CDScaleError")
      };
      super(inputObj);
    }

    static getFieldName() {
      return "CDScale";
    }
  }











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
      const ppenStatusBox = document.getElementById("ppenStatus");
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
        courseOrderUsed.sort(function(a, b){return a - b});

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

              //Append comma to lists - always have a trailing comma
              //Circle position in cm with origin in bottom left corner
              tableColNode.getElementsByClassName("circlex")[0].innerHTML += (0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("x")) - leftcoord) * mapFileScale / courseScale).toString() + ",";
              tableColNode.getElementsByClassName("circley")[0].innerHTML += (0.1 * (Number(controlNode.getElementsByTagName("location")[0].getAttribute("y")) - bottomcoord) * mapFileScale / courseScale).toString() + ",";

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
            for (existingRowID = 0;; existingRowID++) {
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
      }	else if (contentField.value == 0) {
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
        //FIXME: Blob.arrayBuffer isn't implemented in Safari - switch to new code when available
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
        if (logLength > 0){
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

  //Initialisation

  const dynamicCSS = (() => {
    //For changing styles of classes on the fly
    const styleElement = document.createElement("style");
    document.head.appendChild(styleElement);
    const stylesheet = styleElement.sheet;
    const selectors = [];

    function getStyle(selector) {
      let index = selectors.indexOf(selector);
      if (index === -1) {
        //Create new rule
        index = stylesheet.insertRule(selector + "{}", selectors.length);
        selectors.pop(selector);
      }
      return stylesheet.cssRules[index].style;
    }

    function show(selector) {
      const style = getStyle(selector);
      style.display = "";
    }

    function hide(selector) {
      const style = getStyle(selector);
      style.display = "none";
    }

    return {
      show: show,
      hide: hide
    }
  })();

  //Create root level objects
  courseList = new CourseList();
  // stationList = new StationList();

  //Populate with essentials
  courseList.add();

  const mapsCompiler = new Compiler({
    statusBox: document.getElementById("compileStatus"),
    outDownload: document.getElementById("savePDF"),
    logDownload: document.getElementById("viewLog")
  });

  //Do not use arrow functions when binding methods as properties
  mapsCompiler.makePNGs = async function(pdfURL) {
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
    for (let stationId = 1; stationId <= numStations; stationId++) {
      const numTasks = Number(tableRows[stationId].getElementsByClassName("numProblems")[0].textContent);
      if (tableRows[stationId].getElementsByClassName("showStation")[0].checked) {
        const stationName = tableRows[stationId].getElementsByClassName("stationName")[0].textContent;

        //Calculate scale required to render PDF at correct resolution
        const circleDiameter = Number(tableRows[stationId].getElementsByClassName("mapSize")[0].value);
        const mapDescriptionSeparation = (9.5 - circleDiameter - 0.77) / 3;
        const pdfHeight = 9.5 - 2 * mapDescriptionSeparation; //cm
        const imDPI = 300 / pdfHeight * 2.54;
        const pdfScale = imDPI / 72; //PDF renders at 72 DPI by default

        for (let taskId = 1; taskId <= numTasks; taskId++) {
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
          const croppedHeight = bottomRow - topRow + 1;
          const croppedWidth = rightCol - leftCol + 1;
          canvasCropped.height = croppedHeight;
          canvasCropped.width = croppedWidth;
          ctxCropped.drawImage(canvasFull, leftCol, topRow, croppedWidth, croppedHeight, 0, 0, croppedWidth, croppedHeight);

          //Save to zip
          const blob = await new Promise((resolve) => { canvasCropped.toBlob(resolve); });
          imPromises.push(zip.file("map-" + stationName + "." + taskId + "z.png", blob));

          pageNum++;
          taskCount++;
          this.statusBox.textContent = "Splitting into images (" + (taskCount / this.compileUnitsItems * 100).toFixed() + "%).";
        }
      } else {
        pageNum += numTasks;
      }
    }
    await Promise.all(imPromises);
    const zipBlob = await zip.generateAsync({ type: "blob" });
    this.downloadOutput(URL.createObjectURL(zipBlob));
  }

  mapsCompiler.setTemplate = function(name) {
    //Compile time units are determined empirically for each template to give an approximate progress of compiler
    switch (name) {
    case "printA5onA4":
      this.texsrc = "maps.tex";
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

  //TODO: From legacy UI
  // document.getElementById("stationProperties").hidden = true;
  document.getElementById("savePDF").hidden = true;
  document.getElementById("viewLog").hidden = true;

  document.getElementById("selectTemplate").addEventListener("change", updateTemplate);
  updateTemplate.call(document.getElementById("selectTemplate"));
  document.getElementById("compileLaTeXBtn").addEventListener("click", Compiler.generateAll, { passive: true });

  //Initialise variables. Do it this way rather than in HTML to avoid multiple hardcodings of the same initial values.
  // stationList.add(stationList.default);
  // stationList.refresh(false);

  //Event listeners - always call with arrow functions to ensure this doesn't point to calling DOM element
  // stationList.addBtnListeners();
  // stationList.defaultRadio.addEventListener("change", () => { stationList.refresh(true); });
  // stationList.stationRadio.addEventListener("change", () => { stationList.refresh(true); });
  // document.getElementById("setAllCore").addEventListener("click", () => { stationList.applyAll("setAll", stationList.default.coreFields); });
  // document.getElementById("setAllCustomLayout").addEventListener("click", () => { stationList.applyAll("setAll", stationList.default.customLayoutFields); });
  // document.getElementById("resetAllCustomLayout").addEventListener("click", () => { stationList.applyAll("resetValue", stationList.default.customLayoutFields); });
  // for (const field of stationList.default.fieldNames) {
  //   let inputEvent = "input";
  //   if (stationList.default[field].inputElement.tagName === "SELECT") {
  //     inputEvent = "change";
  //   }
  //   stationList.default[field].inputElement.addEventListener(inputEvent, () => { stationList.activeItem[field].saveInput(); });
  //   // if (stationList.default[field].resetBtn !== undefined) {
  //   //   stationList.default[field].resetBtn.addEventListener("click", () => { stationList.activeItem[field].resetValue(); });
  //   // }
  //   // if (stationList.default[field].setAllBtn !== undefined) {
  //   //   stationList.default[field].setAllBtn.addEventListener("click", () => { stationList.activeItem[field].setAll(); });
  //   // }
  // }
  // for (const field of stationList.items[0].taskList.items[0].fieldNames) {
  //   let inputEvent = "input";
  //   if (stationList.items[0].taskList.items[0][field].inputElement.tagName === "SELECT") {
  //     inputEvent = "change";
  //   }
  //   stationList.items[0].taskList.items[0][field].inputElement.addEventListener(inputEvent, () => { stationList.activeItem.taskList.activeItem[field].saveInput(); });
  //   if (stationList.items[0].taskList.items[0][field].resetBtn !== undefined) {
  //     stationList.items[0].taskList.items[0][field].resetBtn.addEventListener("click", () => { stationList.activeItem.taskList.activeItem[field].resetValue(); });
  //   }
  //   if (stationList.items[0].taskList.items[0][field].setAllBtn !== undefined) {
  //     stationList.items[0].taskList.items[0][field].setAllBtn.addEventListener("click", () => { stationList.activeItem.taskList.activeItem[field].setAll(); });
  //   }
  // }

  CourseList.addListeners();

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
    resetField: resetField
  };
  // })();
}
